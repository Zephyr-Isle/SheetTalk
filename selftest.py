"""离线自检:不依赖 Excel 和真实 LLM,验证 Agent 循环、工具执行与事件流是否正确。

用法: python selftest.py
"""
from app import agent
from app.tools import Executor


class FakeBridge:
    host_label = "Microsoft Excel"

    def __init__(self, selection=None):
        self._selection = selection

    def ensure(self):
        pass

    def workbook_context(self):
        return {"workbook": "测试.xlsx", "sheets": [{"name": "Sheet1", "rows": 5, "cols": 3}]}

    def selection_context(self):
        return self._selection


class FakeLLM:
    def __init__(self):
        self.calls = 0

    def __call__(self, *a, **kw):
        self.calls += 1
        if self.calls == 1:
            return {"content": "", "tool_calls": [{
                "id": "call_1", "type": "function",
                "function": {"name": "get_workbook_overview", "arguments": "{}"}}]}
        text = "概况已读取,工作簿共 1 个工作表。"
        on_delta = kw.get("on_delta")
        if on_delta:
            on_delta(text)
        return {"content": text, "tool_calls": None}


def main():
    real = agent.llm.chat_completions
    agent.llm.chat_completions = FakeLLM()
    events = []
    agent.run_agent(FakeBridge(), {"base_url": "http://x", "api_key": "k", "model": "m"},
                    [], "你好", Executor(None), events.append)
    agent.llm.chat_completions = real

    kinds = [e["type"] for e in events]
    for need in ("status", "text_start", "step", "step_result", "delta", "done"):
        assert need in kinds, f"缺少事件 {need},实际:{kinds}"
    assert kinds[-1] == "done", kinds

    # 工具在桥接不可用时也应返回错误 dict 而不是抛异常
    r = Executor(None).execute("read_range", {"cell_range": "A1:B2"})
    assert isinstance(r, dict) and "error" in r, r

    for e in events:
        print(e)

    check_advanced()
    check_hardening()
    check_rag()
    check_safety()
    print("SELFTEST OK")


def check_hardening():
    """加固自检:图标覆盖率 / 数据规整 / TCP 探测 / 「主动退出」标记。"""
    import app.watcher as watcher
    from app import icon_design, runtime
    from app.excel_bridge import _cap_rows, _rgb, _to_rows

    assert _to_rows(None) == [[None]]
    assert _to_rows(5) == [[5]]
    rows, note = _cap_rows([[1, 2], [3, 4], [5, 6]], 4)
    assert rows == [[1, 2], [3, 4]] and note and "截断" in note
    assert _cap_rows([[1]], 500)[1] is None
    assert _rgb("FF8800") == 0x0088FF

    # 图标:中心必须实心、画布角必须透明(回归「内容只画在左上角」的坐标系 bug)
    px, s = icon_design.render(32, ss=2)
    assert px[16 * s + 16][3] > 0.9 and px[0][3] == 0.0

    assert runtime.is_frozen() is False   # 开发态

    # TCP 探测:对确定无人监听的端口必须返回 False 且不抛异常
    old = watcher.PORT
    watcher.PORT = 18765
    try:
        assert watcher.app_running(0.3) is False
    finally:
        watcher.PORT = old

    # 「用户主动退出」标记的写入/清除
    watcher.mark_user_quit()
    assert watcher.user_quit_pending()
    watcher.clear_user_quit()
    assert not watcher.user_quit_pending()

    # 工具不支持判断:只认 tool/function 关键词,普通 HTTP 400 不得误报
    from app import llm
    assert llm.is_tool_unsupported_error("Error: tools is not supported")
    assert llm.is_tool_unsupported_error("This model has no function calling support")
    assert not llm.is_tool_unsupported_error("模型服务返回 HTTP 400: model not found")

    # samples 归一化:字符串/半残对象/垃圾 → 对象数组或 None(三端形状契约)
    from app.web_server import _normalize_samples
    out = _normalize_samples({"samples": [
        "纯文本", {"text": "只有文字"}, 42, None, "", {"text": "T", "prompt": "P"}]})["samples"]
    assert out == [
        {"icon": "\uEA80", "text": "纯文本", "prompt": "纯文本"},
        {"icon": "\uEA80", "text": "只有文字", "prompt": "只有文字"},
        {"icon": "\uEA80", "text": "T", "prompt": "P"},
    ]
    assert _normalize_samples({"samples": "坏形状"})["samples"] is None
    assert _normalize_samples({"samples": []})["samples"] is None
    assert _normalize_samples({}) == {}

    # 工具注册一致性:schema、执行器方法、步骤名三者不得脱节
    from app.tools import TOOLS, STEP_NAMES
    ex2 = Executor(None)
    for t in TOOLS:
        n = t["function"]["name"]
        assert hasattr(ex2, "_t_" + n), f"工具 {n} 缺少 _t_ 执行器"
        assert n in STEP_NAMES, f"工具 {n} 缺少步骤名"
        assert ex2.describe(n, {})
        assert isinstance(ex2.execute(n, {}), dict), f"工具 {n} 未返回 dict"
    print(f"PASS 加固:误报判断 / samples 归一化 / 退出标记 / 工具一致性({len(TOOLS)} 个)")


def check_rag():
    """工具 RAG:检索命中率 + 元工具往返(走 run_agent 全链路)。"""
    import json as _json
    import threading

    from app.tools import META_TOOLS, retrieve_tools

    def names(q, k=5):
        return [t["function"]["name"] for t in retrieve_tools(q, k)]

    cases = {
        "把 C 列宽度调到 20": "set_column_width",
        "删掉那个图表": "delete_chart",
        "做个透视表按产品汇总": "create_pivot",
        "把华东全改成华东区": "find_replace",
        "冻结第一行": "freeze_panes",
        "给这格加个超链接": "set_hyperlink",
        "导出成 pdf": "export_pdf",
        "在第 3 行前插入两行": "insert_rows",
        "只粘贴值到 F1": "copy_range",
        "取消筛选": "auto_filter",
    }
    for q, expect in cases.items():
        got = names(q)
        assert expect in got, f"检索「{q}」未命中 {expect},得到:{got}"
    mnames = [t["function"]["name"] for t in META_TOOLS]
    assert mnames == ["search_tools", "call_tool"], mnames

    # 元工具往返:fake LLM 先 search_tools 再 call_tool,断言事件与结果
    from app import agent as agent_mod

    class MetaLLM:
        def __init__(self):
            self.calls = 0

        def __call__(self, *a, **kw):
            self.calls += 1
            if self.calls == 1:
                return {"content": "", "tool_calls": [{"id": "m1", "type": "function",
                        "function": {"name": "search_tools",
                                     "arguments": _json.dumps({"query": "列宽"})}}]}
            if self.calls == 2:
                return {"content": "", "tool_calls": [{"id": "m2", "type": "function",
                        "function": {"name": "call_tool", "arguments": _json.dumps(
                            {"name": "set_column_width",
                             "args": {"column": "C", "width": 20}})}}]}
            return {"content": "完成", "tool_calls": None}

    real = agent_mod.llm.chat_completions
    agent_mod.llm.chat_completions = MetaLLM()
    ev = []
    try:
        agent_mod.run_agent(FakeBridge(), {"base_url": "http://x", "api_key": "k",
                                           "model": "m"}, [], "把 C 列宽度调到 20",
                            Executor(None), ev.append)
    finally:
        agent_mod.llm.chat_completions = real
    kinds = [e["type"] for e in ev]
    assert "step" in kinds and "step_result" in kinds and "done" in kinds, kinds
    steps = [e for e in ev if e["type"] == "step"]
    assert any(e["tool"] == "set_column_width" for e in steps), steps
    print("PASS 工具 RAG:10 条检索命中 + 元工具往返(检索→调用→完成)")


def check_safety():
    """安全加固:参数校验 / call_tool 白名单 / 危险工具二次确认(拒绝与允许两路径)。"""
    import json as _json
    import threading
    import time as _t

    from app import agent as agent_mod
    from app.tools import DANGEROUS_TOOLS, validate_args

    # 1) 参数校验:数字字符串纠正、必填缺失、枚举越界
    schema = {"properties": {"width": {"type": "number"},
                             "kind": {"type": "string", "enum": ["a", "b"]}},
              "required": ["width"]}
    args, errs = validate_args(schema, {"width": "20", "kind": "a"})
    assert not errs and args["width"] == 20.0, (args, errs)
    _, errs = validate_args(schema, {"kind": "a"})
    assert any("width" in e for e in errs)
    _, errs = validate_args(schema, {"width": 1, "kind": "c"})
    assert any("kind" in e for e in errs)

    # 2) 危险工具清单与白名单关系
    assert "delete_sheet" in DANGEROUS_TOOLS and "close_workbook" in DANGEROUS_TOOLS
    assert "read_range" not in DANGEROUS_TOOLS

    # 3) 二次确认:拒绝路径与允许路径(各跑一次 run_agent 全链路)
    def make_llm(with_id=True, name="delete_sheet", tool_args=None):
        calls = {"n": 0}

        def llm(*a, **kw):
            calls["n"] += 1
            if calls["n"] == 1:
                tc = {"name": name,
                      "arguments": _json.dumps(tool_args or {"name": "表1"})}
                if with_id:
                    tc = dict(_json.loads('{"id":"d1"}'), **tc)
                return {"content": "", "tool_calls": [{"id": "d1" if with_id else None,
                        "type": "function", "function": tc}]}
            return {"content": "已处理", "tool_calls": None}
        return llm

    real = agent_mod.llm.chat_completions
    try:
        for allow in (False, True):
            for with_id in (True, False):
                agent_mod.llm.chat_completions = make_llm(with_id)
                ev = []
                done = threading.Event()

                def run():
                    try:
                        agent_mod.run_agent(FakeBridge(), {"base_url": "http://x",
                            "api_key": "k", "model": "m"}, [], "删掉表1",
                            Executor(None), ev.append)
                    finally:
                        done.set()
                th = threading.Thread(target=run, daemon=True)
                th.start()
                sid = None
                for _ in range(50):          # 最多等 5 秒拿 confirm 事件
                    for e in ev:
                        if e["type"] == "confirm":
                            sid = e["sid"]
                            break
                    if sid:
                        break
                    _t.sleep(0.1)
                assert sid, f"未发出 confirm 事件(with_id={with_id}): {[e.get('type', e.get('kind', e)) for e in ev]}"
                assert agent_mod.resolve_confirm(sid, allow), "resolve 未命中"
                assert done.wait(10), f"确认后未继续(with_id={with_id})"
                kinds = [e["type"] for e in ev]
                assert "confirm" in kinds, kinds
                steps = [e for e in ev if e["type"] == "step_result"]
                assert steps, "缺少 step_result"
                if allow:
                    assert "未确认" not in steps[-1]["brief"], steps[-1]
                else:
                    assert "未确认" in steps[-1]["brief"], steps[-1]
    finally:
        agent_mod.llm.chat_completions = real

    # 3b) remember 路径:「本会话始终允许」后,同类操作不再弹确认
    agent_mod._SESSION_ALLOWED_TOOLS.clear()
    agent_mod.llm.chat_completions = make_llm(False)
    ev = []
    done = threading.Event()

    def run_remember():
        try:
            agent_mod.run_agent(FakeBridge(), {"base_url": "http://x",
                "api_key": "k", "model": "m"}, [], "删掉表1",
                Executor(None), ev.append)
        finally:
            done.set()
    th = threading.Thread(target=run_remember, daemon=True)
    th.start()
    sid = None
    for _ in range(50):
        for e in ev:
            if e["type"] == "confirm":
                sid = e["sid"]
                break
        if sid:
            break
        _t.sleep(0.1)
    assert sid, "未发出 confirm 事件"
    agent_mod.resolve_confirm(sid, True, True)   # remember=True
    assert done.wait(10), "确认后未继续"

    # 第二次:同工具不应再弹确认,直接执行
    agent_mod.llm.chat_completions = make_llm(False)
    ev2 = []
    done2 = threading.Event()

    def run2():
        try:
            agent_mod.run_agent(FakeBridge(), {"base_url": "http://x",
                "api_key": "k", "model": "m"}, [], "删掉表1",
                Executor(None), ev2.append)
        finally:
            done2.set()
    th2 = threading.Thread(target=run2, daemon=True)
    th2.start()
    assert done2.wait(15), "会话允许后未继续"
    kinds2 = [e["type"] for e in ev2]
    assert "confirm" not in kinds2, f"会话允许后不应再弹确认:{kinds2}"
    steps2 = [e for e in ev2 if e["type"] == "step_result"]
    assert steps2 and "未确认" not in steps2[-1]["brief"], steps2
    agent_mod._SESSION_ALLOWED_TOOLS.clear()   # 清理,免得影响其它用例
    agent_mod._SESSION_ALLOWED_TOOLS.discard("clear_range")

    # 3c) remember 只豁免同一工具:换一个危险工具必须仍弹确认
    agent_mod.llm.chat_completions = make_llm(False, "clear_range", {"cell_range": "A1:B2"})
    ev3 = []
    done3 = threading.Event()

    def run3():
        try:
            agent_mod.run_agent(FakeBridge(), {"base_url": "http://x",
                "api_key": "k", "model": "m"}, [], "清空 A1:B2",
                Executor(None), ev3.append)
        finally:
            done3.set()
    th3 = threading.Thread(target=run3, daemon=True)
    th3.start()
    sid3 = None
    for _ in range(50):
        for e in ev3:
            if e["type"] == "confirm" and e.get("tool") == "clear_range":
                sid3 = e["sid"]
                break
        if sid3:
            break
        _t.sleep(0.1)
    assert sid3, "换工具后未弹出确认(remember 不应豁免其它工具)"
    agent_mod.resolve_confirm(sid3, False)
    assert done3.wait(10), "拒绝后未继续"

    # 3d) 启动期自检:不变量全部通过
    assert agent_mod.startup_selfcheck() == [], agent_mod.startup_selfcheck()

    # 3e) SSE 监听:arm_confirm 会通知已注册的监听者(SSE 推送的数据源)
    got = []
    agent_mod.on_confirm_request(lambda p: got.append(dict(p)))
    agent_mod.arm_confirm("sse-selftest", "delete_sheet", "自检", {"name": "x"})
    agent_mod._CONFIRM_WAIT.pop("sse-selftest", None)
    agent_mod._CONFIRM_LISTENERS.remove(lambda p: None) if False else None
    assert got and got[-1]["sid"] == "sse-selftest" and got[-1]["tool"] == "delete_sheet", got
    agent_mod._CONFIRM_LISTENERS.pop()   # 移除测试监听,不影响其它用例

    # 4) web_server 异常恢复:页面崩溃/刷新后,/api/pending_confirm 重新取回
    #    未决断的确认,/api/confirm 回传决定(走真实 HTTP 服务验证)
    import urllib.request as _uq
    from app.web_server import create_server

    class FakeSession:
        def __init__(self):
            self.calls = []

        def confirm(self, sid, allowed, remember=False):
            # 与真实 ChatSession.confirm 同构:回传时必须决断注册表里的等待
            self.calls.append((sid, allowed, remember))
            agent_mod.resolve_confirm(sid, allowed, remember)

    fsession = FakeSession()
    srv = create_server(FakeBridge(), fsession, port=18766)
    assert srv, "测试服务启动失败"
    tbase = "http://127.0.0.1:18766"
    try:
        rsid = "recovery-test-1"
        agent_mod.arm_confirm(rsid, "delete_sheet", "删除工作表:Sheet1")
        with _uq.urlopen(tbase + "/api/pending_confirm", timeout=5) as r:
            d = _json.loads(r.read().decode("utf-8"))
        assert d.get("sid") == rsid and d.get("tool") == "delete_sheet", d
        req = _uq.Request(tbase + "/api/confirm", method="POST",
                          data=_json.dumps({"sid": rsid, "allowed": True}).encode("utf-8"),
                          headers={"Content-Type": "application/json"})
        with _uq.urlopen(req, timeout=5) as r:
            _json.loads(r.read().decode("utf-8"))
        assert fsession.calls and fsession.calls[0][0] == rsid             and fsession.calls[0][1] is True, fsession.calls
        with _uq.urlopen(tbase + "/api/pending_confirm", timeout=5) as r:
            d2 = _json.loads(r.read().decode("utf-8"))
        assert not d2, f"决断后应为空:{d2}"
    finally:
        srv.shutdown()
        agent_mod._CONFIRM_WAIT.pop(rsid, None)
    print("PASS 安全:校验 / 白名单 / 二次确认(四象限+会话记住) / web_server 恢复")


def check_advanced():
    """对齐 Copilot 的交互增强:引用溯源 / 选中区域上下文 / 中止 / 计划模式。"""
    import threading

    from app.chat import ChatSession

    print("\n--- 交互增强自检 ---")

    # 1) 引用溯源:步骤事件要能带上区域
    ex = Executor(None)
    assert ex.describe_ref("read_range", {"sheet": "Sheet1", "cell_range": "A1:C9"}) == "Sheet1!A1:C9"
    assert ex.describe_ref("create_chart", {"data_range": "B2:B10"}) == "活动表!B2:B10"
    assert ex.describe_ref("get_workbook_overview", {}) == ""
    print("PASS 引用溯源 describe_ref")

    # 2) 选中区域上下文:无选中时不干扰,有选中时注入系统提示词
    assert agent._selection_block(FakeBridge()) == ""
    sel = {"sheet": "订单", "address": "B2:C4", "rows": 3, "cols": 2,
           "truncated": False, "values": [["a", 1], ["b", 2], ["c", 3]]}
    blk = agent._selection_block(FakeBridge(sel))
    assert "订单" in blk and "B2:C4" in blk and '"a"' in blk, blk
    sys_msg = agent.build_messages(FakeBridge(sel), [], "总结一下")[0]["content"]
    assert "当前选中区域" in sys_msg
    assert "{context}" not in sys_msg and "{selection}" not in sys_msg, "占位符没被替换"
    print("PASS 选中区域上下文")

    real = agent.llm.chat_completions

    # 3) 中止:预置 cancel 后应只发 status+cancelled,且不再调工具
    calls = {"n": 0}

    def loop_llm(*a, **kw):
        calls["n"] += 1
        return {"content": "", "tool_calls": [{
            "id": "c%d" % calls["n"], "type": "function",
            "function": {"name": "read_range", "arguments": '{"cell_range":"A1:B2"}'}}]}

    agent.llm.chat_completions = loop_llm
    ev = []
    cancel = threading.Event()
    cancel.set()
    agent.run_agent(FakeBridge(), {"base_url": "http://x", "api_key": "k", "model": "m"},
                    [], "别动", Executor(None), ev.append, cancel=cancel)
    kinds = [e["type"] for e in ev]
    assert kinds == ["status", "cancelled"], kinds
    assert calls["n"] == 0, "中止后仍调用了 LLM"
    print("PASS 中止:预置 cancel → %s" % kinds)

    # 4) 计划模式:应额外产出 plan 事件
    def plan_llm(*a, **kw):
        if kw.get("tools") is None:
            return {"content": "1. 读数据\n2. 生成饼图"}
        return {"content": "已完成。", "tool_calls": None}

    agent.llm.chat_completions = plan_llm
    ev2 = []
    agent.run_agent(FakeBridge(), {"base_url": "http://x", "api_key": "k", "model": "m",
                                   "plan_mode": True}, [], "做占比图",
                    Executor(None), ev2.append)
    agent.llm.chat_completions = real
    assert "plan" in [e["type"] for e in ev2]
    assert "生成饼图" in [e for e in ev2 if e["type"] == "plan"][0]["text"]
    print("PASS 计划模式:发出 plan 事件")

    # 5) 事件翻译:cancel 事件要透传,step 要带上 ref
    import app.chat as chatmod
    seen = []

    def fake_run(*a, **kw):
        emit, cancel_kw = a[5], kw.get("cancel")
        assert isinstance(cancel_kw, threading.Event), "cancel 未透传给 agent"
        emit({"type": "plan", "text": "计划"})
        emit({"type": "step", "sid": "s1", "summary": "读取 A1:B2", "ref": "Sheet1!A1:B2"})
        emit({"type": "delta", "text": "半截正文"})
        emit({"type": "cancelled"})

    real_run = chatmod.run_agent
    chatmod.run_agent = fake_run
    try:
        ChatSession(FakeBridge()).send("测试", seen.append)
    finally:
        chatmod.run_agent = real_run
    kk = [e["kind"] for e in seen]
    assert "plan" in kk and "cancelled" in kk, kk
    assert [e for e in seen if e["kind"] == "step"][0]["ref"] == "Sheet1!A1:B2"
    print("PASS 事件翻译:plan / cancelled / ref")


if __name__ == "__main__":
    main()
