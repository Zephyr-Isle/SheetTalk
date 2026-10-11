"""Agent 主流程:注入表格上下文 → 调用 LLM(流式) → 执行工具调用 → 循环直到给出回答。

run_agent 不是生成器,通过 emit 回调逐个推送事件(dict),由 UI 层决定如何展示。
事件类型:status / plan / delta / note_flush / step / step_result / error / cancelled / done / confirm。

三种工作模式(尽量对齐 Microsoft 365 Copilot in Excel):
  直接干(默认):一次流式输出里先叙述再调用工具。
  计划模式(plan_mode):先单独让模型产出可读的执行计划并展示,再按计划执行。
  可随时中止:cancel 是一个 threading.Event,置位后尽快收尾(不再调工具)。

危险工具二次确认:delete_*/clear_range/close_workbook/protect_* 执行前发 confirm
事件并在 wait_user_confirm 阻塞等待;UI 通过 resolve_confirm(sid, allowed) 回填结果。
用户选「本会话始终允许」得到的豁免按 SESSION_ALLOWED_TTL(30 分钟)自动失效,
两端界面都显示剩余时间(见 session_allowed_items)。
"""
import json
import logging
import threading
import time

from app import llm
from app.tools import (CORE_TOOL_NAMES, DANGEROUS_TOOLS, META_TOOLS, STEP_NAMES, TOOLS,
                       dangerous_tools, retrieve_tools, validate_args,
                       brief_result)

log = logging.getLogger("excelai")

HISTORY_TURNS = 16

# 危险工具二次确认的等待注册表:sid → {event, allowed, tool, args, summary}
_CONFIRM_WAIT = {}
# 用户选过「本会话始终允许」的豁免:{作用域键: 失效时刻(time.monotonic)}。
# 键是「工具名」或「工具名:主参数」(见 _scope_key)。会话豁免不该在长会话里
# 无限放行:超过 SESSION_ALLOWED_TTL 秒自动失效,两端界面显示剩余时间。
_SESSION_ALLOWED_TOOLS = {}
SESSION_ALLOWED_TTL = 30 * 60
# 确认请求监听(web_server SSE 推送用):回调在 agent 线程,须自行线程安全
_CONFIRM_LISTENERS = []


def on_confirm_request(fn):
    """注册确认请求监听(SSE 推送);重复注册忽略。"""
    if fn not in _CONFIRM_LISTENERS:
        _CONFIRM_LISTENERS.append(fn)


def arm_confirm(sid, tool, summary, args=None):
    """登记一个待确认请求并返回其 Event,同时通知监听者(SSE 推送)。

    必须「先登记、后发 confirm 事件」:若 UI 在登记前回传,确认会静默丢失。
    """
    ev = threading.Event()
    log.info("二次确认请求 sid=%s tool=%s", sid, tool)
    _CONFIRM_WAIT[sid] = {"event": ev, "allowed": False, "tool": tool,
                          "summary": summary, "args": args or {}}
    payload = {"type": "confirm", "sid": sid, "tool": tool, "summary": summary}
    for fn in list(_CONFIRM_LISTENERS):
        try:
            fn(payload)
        except Exception:
            log.exception("确认请求监听器异常")
    return ev


def wait_confirm_event(ev, record, cancel, timeout=180):
    """等待确认事件被回传;返回 record 里的 allowed(拒绝 → False,不得执行)。

    注意不能只看事件是否置位:拒绝也是 set 事件,丢了 allowed 判定会把
    「用户拒绝」当成「允许」执行工具。
    """
    try:
        deadline = time.time() + timeout
        while not ev.wait(0.5):
            if cancel is not None and cancel.is_set():
                return False
            if time.time() > deadline:
                return False
        return bool(record.get("allowed", False))
    finally:
        pass


def pending_confirms():
    """列出尚未决断的确认请求(页面崩溃/刷新后重新弹确认条用)。"""
    return [{"sid": sid, "tool": st.get("tool", ""),
             "summary": st.get("summary", "")}
            for sid, st in _CONFIRM_WAIT.items() if not st["event"].is_set()]


def startup_selfcheck():
    """启动期自检:验证确认流/工具注册/校验/检索的关键不变量。

    返回问题列表(空 = 全部通过);问题会在 UI 红色横幅明示,
    把「运行时才炸」的隐患提前到启动期暴露。
    """
    issues = []

    # ① 确认流不变量:登记→允许回填→拒绝语义(wait 必须读 allowed)
    try:
        sid = "__selfcheck_allow__"
        ev = arm_confirm(sid, "__selfcheck_delete__", "自检", {"name": "x"})
        resolve_confirm(sid, True, False)
        if not ev.is_set() or not _CONFIRM_WAIT.get(sid, {}).get("allowed"):
            issues.append("确认流:允许语义失效(resolve 未回填 allowed)")
        _CONFIRM_WAIT.pop(sid, None)
        sid2 = "__selfcheck_deny__"
        ev2 = arm_confirm(sid2, "__selfcheck_delete__", "自检", {"name": "x"})
        _CONFIRM_WAIT[sid2]["allowed"] = False
        ev2.set()
        if wait_confirm_event(ev2, _CONFIRM_WAIT[sid2], None):
            issues.append("确认流:拒绝被当成允许(wait 未读 allowed)")
        _CONFIRM_WAIT.pop(sid2, None)
    except Exception as e:
        issues.append(f"确认流自检异常:{e}")

    # ② 工具注册不变量:危险清单 ⊆ 注册表;每个工具都有执行器;校验与检索可用
    try:
        names = {t["function"]["name"] for t in TOOLS}
        missing = sorted(dangerous_tools() - names)
        if missing:
            issues.append(f"DANGEROUS 含未注册工具:{missing}")
        from app.tools import Executor
        ex = Executor(None)
        no_exec = sorted(n for n in names if not hasattr(ex, "_t_" + n))
        if no_exec:
            issues.append(f"工具缺少执行器:{no_exec}")
        a2, e2 = validate_args({"properties": {"w": {"type": "number"}},
                                "required": ["w"]}, {"w": "3"})
        if e2 or a2.get("w") != 3.0:
            issues.append("参数校验:数字字符串纠正失效")
        if "set_column_width" not in [t["function"]["name"]
                                      for t in retrieve_tools("把列宽调大一点", 5)]:
            issues.append("工具检索未命中 set_column_width")
    except Exception as e:
        issues.append(f"工具自检异常:{e}")

    # ③ 会话豁免的作用域与时效不变量:同对象豁免、异对象隔离、到期自动失效
    try:
        _SESSION_ALLOWED_TOOLS.clear()
        k1 = _scope_key("delete_sheet", {"name": "A"})
        _grant_exemption(k1)
        if not _is_exempt(k1):
            issues.append("会话豁免:同对象豁免丢失")
        if _is_exempt(_scope_key("delete_sheet", {"name": "B"})):
            issues.append("会话豁免:作用域未隔离(不同对象被放行)")
        # 键名带 __selfcheck__ 前缀:到期回收会写日志,一眼能看出是自检造的而不是用户数据
        _grant_exemption("__selfcheck_expired__", ttl=-1)
        if _is_exempt("__selfcheck_expired__") or "__selfcheck_expired__" in _SESSION_ALLOWED_TOOLS:
            issues.append("会话豁免:到期未自动失效")
        _SESSION_ALLOWED_TOOLS.clear()
    except Exception as e:
        issues.append(f"会话豁免自检异常:{e}")
    return issues


def _scope_key(tool, args):
    """会话豁免的作用域键:工具名 + 首个必填参数值(参数前缀)。

    例:delete_sheet + {"name": "Sheet2"} → "delete_sheet:Sheet2",
    只豁免对同一对象的重复操作;换一个对象仍会弹确认。
    """
    schema = next((t["function"].get("parameters") for t in TOOLS
                   if t["function"]["name"] == tool), None) or {}
    req = schema.get("required") or []
    primary = (args or {}).get(req[0]) if req else None
    if primary is None:
        return tool
    return f"{tool}:{primary}"


def _grant_exemption(scope, ttl=None):
    """记一条会话豁免:默认 SESSION_ALLOWED_TTL 秒后自动失效(ttl 仅供自检/测试)。"""
    _SESSION_ALLOWED_TOOLS[str(scope)] = time.monotonic() + (
        SESSION_ALLOWED_TTL if ttl is None else ttl)


def _prune_session_allowed(now=None):
    """回收已到期的豁免:两端读清单前都先走这里,过期的不会再显示出来。"""
    now = time.monotonic() if now is None else now
    for key, expires_at in list(_SESSION_ALLOWED_TOOLS.items()):
        if expires_at <= now:
            _SESSION_ALLOWED_TOOLS.pop(key, None)
            log.info("会话豁免已到期自动失效:%s", key)


def _is_exempt(*scopes):
    """这些作用域键里是否有仍然有效的豁免;顺手回收已到期的条目。"""
    now = time.monotonic()
    for scope in scopes:
        if not scope:
            continue
        expires_at = _SESSION_ALLOWED_TOOLS.get(scope)
        if expires_at is None:
            continue
        if expires_at > now:
            return True
        _SESSION_ALLOWED_TOOLS.pop(scope, None)
        log.info("会话豁免已到期自动失效:%s", scope)
    return False


def session_allowed_tools():
    """当前仍然有效的会话豁免作用域键(过期的先回收)。"""
    _prune_session_allowed()
    return sorted(_SESSION_ALLOWED_TOOLS)


def session_allowed_items():
    """豁免清单的可读条目:{key, tool, label, arg, expires_in}(两端界面共用)。

    label 取 tools.STEP_NAMES 的中文名(未知工具回退工具名),界面不必自备名表;
    expires_in 是剩余秒数,两端各自本地倒计时,不必反复回查。
    arg 是模型给的参数值,前端只允许当文本渲染(不得拼进 HTML)。
    """
    _prune_session_allowed()
    now = time.monotonic()
    items = []
    for key, expires_at in sorted(_SESSION_ALLOWED_TOOLS.items()):
        tool, _, arg = str(key).partition(":")
        items.append({"key": str(key), "tool": tool, "label": STEP_NAMES.get(tool, tool),
                      "arg": arg, "expires_in": max(0, int(round(expires_at - now)))})
    return items


def revoke_session_allowed(scope):
    """撤销单条会话豁免(scope 取 session_allowed_items 里的 key)。

    返回是否真撤到了一条:未知/已过期的键返回 False 而不报错 —— 两端都可能
    拿着稍早拿到的清单去点撤销,键在这中间自然过期是正常情况。
    """
    hit = _SESSION_ALLOWED_TOOLS.pop(str(scope), None) is not None
    log.info("撤销会话豁免:%s(%s)", scope, "已撤销" if hit else "不存在或已过期")
    return hit


def clear_session_allowed():
    """一键撤销全部会话豁免(之后危险操作恢复逐次确认)。"""
    _SESSION_ALLOWED_TOOLS.clear()
    log.info("已清空会话豁免清单")


def resolve_confirm(sid, allowed, remember=False):
    """UI 的确认/取消回调(桌面 ctrl.confirmRun / POST /api/confirm 都走这里)。

    remember=True 且允许时,把该操作的「作用域键」加入会话级允许清单
    (工具名+主参数),之后同范围操作不再重复确认;换对象仍会确认。
    豁免带时效:SESSION_ALLOWED_TTL 秒后自动失效,重新回到逐次确认。
    """
    st = _CONFIRM_WAIT.get(str(sid))
    if not st:
        return False
    log.info("二次确认回传 sid=%s allowed=%s remember=%s",
             str(sid), bool(allowed), bool(remember))
    st["allowed"] = bool(allowed)
    if st.get("tool"):
        scope = _scope_key(st["tool"], st.get("args"))
        if allowed and remember:
            _grant_exemption(scope)
            log.info("已加入会话豁免(%d 分钟后失效):%s", SESSION_ALLOWED_TTL // 60, scope)
        elif not allowed:
            _SESSION_ALLOWED_TOOLS.pop(scope, None)
    st["event"].set()
    return True

SYSTEM_TEMPLATE = """你是 Excel 智能助手,运行在一个独立的桌面应用里,通过 COM 接口直接操作用户当前打开的表格(同时支持 Microsoft Excel 与 WPS 表格)。

【参数契约】所有工具都按这个口径传参(数字/数组写歪会被自动纠正,表名与区域写错则直接失败):
- sheet:要操作的工作表名,省略=当前活动工作表;工作表名必须来自上下文里的 sheets 列表,不要凭空编。
- 区域参数(cell_range / data_range / source_range 等)一律 A1 样式、只写区域本身(如 "A2:D100"、"C5"),不要带工作表名和 $;要动别的表用 sheet 参数。
- 数组参数必须是 JSON 数组:字段名写 ["地区","产品"]、颜色写 ["FFC7CE"],不要写成一个字符串。
- 列序号/行序号是从 1 开始的整数,多数工具是「相对所选区域的序号」(sort_range 的 column_index=2 指该区域的第 2 列,不是工作表的 B 列)。
- 上下限:单次读取约 500 格、单次写入 ≤ 10 万格,超出会被截断或报错;结果里出现 note 说明被截断,要缩小区域或分块再读。
- 数值/布尔参数不要加引号:width 写 20 而不是 "20";call_tool 的 args 是 JSON 对象,不是字符串。

工作规则:
1. 动手修改之前,先用 get_workbook_overview / get_sheet_preview / read_range 弄清数据结构:表头在哪一行、数据从第几行开始、每列是什么含义;后面要用到字段名的地方(透视表/图表/公式)必须以读到的表头为准,不要猜。
2. 写公式一律使用英文函数名和英文逗号分隔,如 =SUMIFS(C2:C100,A2:A100,"华东")。程序会以 Formula 形式写入,Excel/WPS 会自动本地化显示。
3. 整列同构公式优先用 autofill_formula:先在第一个单元格写好公式,再填充到整个区域;不要逐格重复写公式。
4. write_range 中以 = 开头的字符串会被当作公式;普通文本数据不要带等号。
5. 图表用 create_chart(数据区域要含表头);数字/颜色等格式用 set_format;排序用 sort_range。
6. 只操作与任务相关的区域,不要覆盖或清空无关的用户数据;计算类问题优先用 Excel 公式而不是凭空猜数。
7. 全程用简体中文(包括调用工具前的简短叙述),语气简洁专业,可以用 Markdown(标题、列表、表格、代码块)组织内容。完成后用两三句话总结:做了什么、结果在哪里。遇到失败,解释原因并给出手动操作建议。不要在正文里写「[调用 xxx]」式的伪工具日志 —— 工具执行会自动以步骤卡片展示,正文只写给用户看的内容。
8. 如果用户的请求和表格无关,直接回答即可,不需要调用工具。
9. 用户可能只在表格里选中了一块区域就来提问。优先围绕「当前选中区域」作答:
   直接引用它的地址与内容,不要再去全表搜索;除非用户明确说了要看整表。上下文里没有选区块时,不要假装看到了选区。
10. 做「分组统计 / 汇总」优先用 create_pivot(数据透视表),它比手写 SUMIF 更清晰、
    结果可直接下钻;只要一个数字时用 Excel 公式即可。透视表的 rows/columns/values 只能是表头字段名(JSON 数组),字段名写错会报「找不到成员」。
11. 让数据更容易看用条件格式:比较大小用 color_scale(色阶)或 data_bar(数据条),
    超标预警用 cell_value(如大于目标值标红),找重复值用 duplicate。
12. 数据脏了先清洗:重复行用 remove_duplicates(columns 传列序号数组,如 [1,2];省略=整行比对),想固定输入选项用 set_validation
    (下拉列表,选项串总长不超过 255 字符),要规范表格外观用 create_table(可带汇总行)。
13. 注意:同一区域重复设置条件格式/筛选/转表是幂等的,可以直接重跑;但透视表每次
    都会新建一张,别对同一份数据反复建透视表。去重、删除行列、清空区域、覆盖写入都会改到用户数据,动手前先读一遍确认范围。
14. 常用控制手段:批量改文本用 find_replace(不要逐格写);增删行/列用
    insert_rows / delete_rows / insert_columns / delete_columns(行列都传整数序号);合并/取消合并用
    merge_cells;冻结表头用 freeze_panes(冻结首行填 A2);删除图表用 delete_chart;
    取消表格对象用 delete_table;问「这格的公式」用 get_cell_formula;
    跨表搬运数据用 copy_range;加边框用 set_format(border=thin/medium/thick);
    超链接用 set_hyperlink;复制工作表用 copy_sheet;隐藏/显示行列或表用 set_visible;
    取消高亮用 clear_conditional_format;取消筛选用 auto_filter(action=off);
    只粘贴值/转置粘贴用 copy_range(values_only=true / transpose=true);
    自动换行/垂直对齐用 set_format(wrap_text / vertical);只清格式用 clear_formats;
    多列排序在 sort_range 里给 column_index2;问「我选中的是什么」用 get_selection。
15. 删除类操作(删行/删列/删图表/删表/删工作表)破坏性较强:除非用户明确说删,
    先用 get_workbook_overview 确认对象名称与位置,再动手。
16. 工作簿级:打开/另存为/导出 PDF/文档属性/保护工作簿/定义名称(get_names 看
    现有名称)/刷新数据连接/计算模式与重算;工作表级:移动/标签颜色/网格线/
    页面设置/打印区域与打印标题/缩放/拆分窗格/显示公式;区域级:批注/分列/
    格式刷/字体名/下划线/删除线/缩进/锁定与隐藏公式(先 set_format(locked/hide_formula),再 protect_sheet 才生效);
    数据分析:高级筛选(需先在表旁搭条件区域)/行分组/合并计算/单变量求解;
    图表进阶:config_chart 改类型标题图例标签、add_trendline 趋势线、
    refresh_pivot/delete_pivot 维护透视表。print_sheets 会真实打印,用户没明确
    要求打印时不要调用。
17. 上面没列全:还有更多工具可用 search_tools 按需求搜索(返回工具名与参数),
    再用 call_tool(name, args) 调用。不要凭空猜测未列出的工具名;
    search_tools 也用于确认某类操作是否存在;拿到名字后用 call_tool 调用,它的 args 必须是 JSON 对象。
18. 删除工作表/删除行列/删除图表/清空区域/关闭工作簿等破坏性操作会先弹确认条,
    用户可在下拉里选「仅本次允许」「本会话始终允许」或「拒绝执行」:
    用户确认前不要继续别的动作;拒绝后改提替代方案,不要换个名字重试;
    用户选了「始终允许」后,同类操作直接执行即可,不要再反复打扰(该豁免过一段时间会
    自动失效,届时同类操作会重新弹确认,这是正常现象,不要当成用户改了主意)。
    一次只发起一个需要确认的操作,等它出结果再继续。
19. 多条件筛选:auto_filter 一次只支持一个条件(再调一次会替换上一个);要按多个条件过滤就用 advanced_filter(先在表旁写好条件区域,再传 criteria_range)。给区域建表/去重前不需要自己先关筛选,程序内部已经处理。
20. 工具返回 {"error": ...} 时先读懂原因再决定下一步:参数校验失败就按提示改参数(数组要写成 JSON 数组);「找不到工作表/成员/图表」就以读到的真实名字重试一次;同一个原因不要试第三次,直接说明原因和手动做法。
21. 写完重要数据后用 read_range / get_cell_formula 回读一小块核对(尤其批量写入、公式填充、排序之后),确认无误再汇报;不要只凭工具摘要就说「已完成」。
22. 上下文里若显示打开了多个工作簿而用户没有指明,先一句话问清要操作哪一个,不要猜(程序默认跟随用户当前可见的工作簿)。

当前表格上下文(快照,可能滞后):
{context}
{selection}"""


PLAN_PROMPT = """用户想在 Excel 里完成下面这件事。请先用中文给出 3–6 步可执行计划,
每步一行,以「1. 2. 3.」开头,写清:动作 + 目标工作表与区域(如 Sheet1!A2:D100)+ 要用的工具名;最后追加一行以「校验:」开头,说明改完怎么核对结果。不要执行,只列计划。
用户需求:{user_text}"""

MAX_SELECTION_CELLS = 200


def _selection_block(bridge):
    """把用户当前选中的区域渲染成系统提示词的一段(对齐 Copilot 的「选中即上下文」)。

    读取失败或没有有效选中时返回空串,不打扰用户。单元格过多时截断并说明。
    """
    try:
        sel = bridge.selection_context()
    except Exception:
        return ""
    if not sel or not sel.get("address"):
        return ""
    try:
        values = sel.get("values") or []
        flat = sum(len(r) for r in values)
        if flat > MAX_SELECTION_CELLS:
            values = [r[:20] for r in values[:20]]
            note = "(选中区域过大,已截取前 20 行 × 20 列)"
        else:
            note = ""
    except Exception:
        return ""
    body = json.dumps(values, ensure_ascii=False)
    return ("\n当前选中区域(用户框选的内容,默认以它为准):\n"
            f"工作表:{sel.get('sheet')}  地址:{sel.get('address')}  "
            f"尺寸:{sel.get('rows')}行 × {sel.get('cols')}列 {note}\n{body}")


def build_messages(bridge, history, user_text, use_selection=True):
    try:
        ctx = json.dumps(bridge.workbook_context(), ensure_ascii=False)
    except Exception:
        ctx = "{} (未连接或读取失败)"
    # use_selection=False:UI 上的选区附件被用户移除,本轮不把选区内容注入提示词
    sel_block = _selection_block(bridge) if use_selection else ""
    system = SYSTEM_TEMPLATE.replace("{context}", ctx).replace("{selection}", sel_block)
    msgs = [{"role": "system", "content": system}]
    for m in (history or [])[-HISTORY_TURNS:]:
        role = m.get("role")
        content = m.get("content")
        if role in ("user", "assistant") and content:
            msgs.append({"role": role, "content": str(content)[:4000]})
    msgs.append({"role": "user", "content": user_text})
    return msgs


def _make_plan(bridge, settings, user_text, emit, cancel):
    """计划模式:先让模型只列计划(不带 tools),产出后作为 system 提示回灌。"""
    try:
        system = SYSTEM_TEMPLATE.replace("{context}", "").replace("{selection}", "")
    except Exception:
        system = ""
    msgs = [{"role": "system", "content": system},
            {"role": "user", "content": PLAN_PROMPT.replace("{user_text}", user_text)}]
    try:
        plan = llm.chat_completions(settings["base_url"], settings["api_key"],
                                    settings["model"], msgs, tools=None)
    except Exception as e:
        # 计划只是锦上添花,失败就跳过,不影响主流程
        emit({"type": "status", "text": f"生成计划失败,直接执行:{e}"})
        return None
    plan = (plan.get("content") or "").strip()
    if not plan:
        return None
    emit({"type": "plan", "text": plan})
    return plan


def run_agent(bridge, settings, history, user_text, executor, emit, cancel=None, use_selection=True):
    try:
        bridge.ensure()
    except Exception as e:
        emit({"type": "error", "message": str(e)})
        return
    emit({"type": "status", "text": f"已连接 {bridge.host_label},正在分析…"})

    try:
        max_steps = int(settings.get("max_steps") or 30)
    except Exception:
        max_steps = 30
    max_steps = max(5, min(max_steps, 500))

    messages = build_messages(bridge, history, user_text, use_selection=use_selection)

    # 工具 RAG:每轮只给模型「常驻核心 + 按用户请求检索命中 + 元工具」;
    # tool_rag=False 时回退全量注册表。检索没命中不影响:模型可用元工具按需搜索。
    if settings.get("tool_rag", True):
        seen = set()
        active_tools = []
        for t in TOOLS:
            if t["function"]["name"] in CORE_TOOL_NAMES:
                active_tools.append(t)
                seen.add(t["function"]["name"])
        for t in retrieve_tools(user_text, top_k=5):
            if t["function"]["name"] not in seen:
                active_tools.append(t)
                seen.add(t["function"]["name"])
        active_tools += [t for t in META_TOOLS if t["function"]["name"] not in seen]
    else:
        active_tools = TOOLS

    # 计划模式:先出计划再执行,并把计划作为约束回灌给主循环
    if settings.get("plan_mode"):
        emit({"type": "status", "text": "正在生成执行计划…"})
        plan = _make_plan(bridge, settings, user_text, emit, cancel)
        if cancel is not None and cancel.is_set():
            emit({"type": "cancelled"})
            return
        if plan:
            messages[0]["content"] += ("\n\n本次已生成执行计划,请严格按计划推进,不要偏离:\n" + plan)

    for _ in range(max_steps):
        if cancel is not None and cancel.is_set():
            emit({"type": "cancelled"})
            return
        emit({"type": "text_start"})
        try:
            msg = llm.chat_completions(
                settings["base_url"], settings["api_key"], settings["model"], messages,
                tools=active_tools, on_delta=lambda t: emit({"type": "delta", "text": t}))
        except llm.LLMError as e:
            text = str(e)
            if llm.is_tool_unsupported_error(text):
                text += "\n\n(提示:也可能是当前模型不支持函数调用,请在设置中换成 deepseek-chat、glm-4.6 等支持 tools 的模型。)"
            emit({"type": "error", "message": text})
            return

        tool_calls = msg.get("tool_calls") or []
        if not tool_calls:
            if not (msg.get("content") or "").strip():
                emit({"type": "error", "message": "(模型没有返回内容,请重试)"})
                return
            emit({"type": "done"})
            return

        # 这一轮流式输出的是工具调用前的叙述:把它挪成小字备注,正文留给最终回答
        emit({"type": "note_flush"})
        messages.append({"role": "assistant", "content": msg.get("content") or "",
                         "tool_calls": tool_calls})
        for idx, tc in enumerate(tool_calls):
            # 中止要卡在「调用工具之前」:一旦开始改表就不能半途停下,否则工作簿停在中间态
            if cancel is not None and cancel.is_set():
                emit({"type": "cancelled"})
                return
            # 规范化 tool_call id:个别模型/非流式回退返回的 tool_calls 可能没有 id,
            # 而步骤卡、二次确认、tool 消息都依赖它 —— 缺失会导致确认永远等不到回传
            tc_id = str(tc.get("id") or f"call_{idx}")
            fn = tc.get("function") or {}
            name = fn.get("name") or ""
            try:
                args = json.loads(fn.get("arguments") or "{}")
                if not isinstance(args, dict):
                    args = {}
            except Exception:
                args = {}
            # 元工具:call_tool 展示为被调用的内层工具(步骤卡更可读)
            if name == "call_tool":
                step_name = str((args or {}).get("name") or "call_tool")
                step_args = (args or {}).get("args") or {}
            else:
                step_name, step_args = name, args
            emit({"type": "step", "sid": tc_id, "tool": step_name,
                  "summary": executor.describe(step_name, step_args),
                  "ref": executor.describe_ref(step_name, step_args)})
            if name == "search_tools":
                result = executor.search_tools(str((args or {}).get("query", "")))
            else:
                # call_tool 展开内层工具;确认门按「内层工具名」判断:
                # 危险操作无论从哪条路径来都必须过确认/会话允许,不能绕过
                _scope = _scope_key(step_name, step_args)
                # 豁免带时效:过期的在这里被回收,操作重新回到逐次确认
                if step_name in DANGEROUS_TOOLS \
                        and not _is_exempt(step_name, _scope) \
                        and settings.get("confirm_dangerous", True):
                    _summary = executor.describe(step_name, step_args)
                    # 先登记等待、再发事件(防 UI 抢先回传导致确认丢失)
                    evc = arm_confirm(tc_id, step_name, _summary)
                    emit({"type": "confirm", "sid": tc_id, "tool": step_name,
                          "summary": _summary})
                    try:
                        allowed = wait_confirm_event(evc, _CONFIRM_WAIT[tc_id], cancel)
                    finally:
                        # 无论允许 / 拒绝 / 超时 / 中止都要销号:残留的未决条会被
                        # /api/pending_confirm 与 confirm SSE 反复推给侧边栏,
                        # 弹出一个人早已放弃的确认条(且字典只增不减)。
                        # 晚到的回传此刻 resolve_confirm 找不到条目 → 返回 False,正确。
                        _CONFIRM_WAIT.pop(tc_id, None)
                    if not allowed:
                        result = {"error": "用户未确认该操作,已跳过执行"}
                    else:
                        result = executor.execute(step_name, step_args)
                else:
                    result = executor.execute(step_name, step_args)
            emit({"type": "step_result", "sid": tc_id,
                  "ok": "error" not in result, "brief": brief_result(result)})
            messages.append({"role": "tool", "tool_call_id": tc_id,
                             "content": json.dumps(result, ensure_ascii=False)[:4000]})

    emit({"type": "error",
          "message": f"已达本轮工具步骤上限({max_steps})。可在设置中调大「单次任务最大工具步数」,或把需求拆小一点再试。"})
