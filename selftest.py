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
    check_exemptions()
    check_regressions()
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
    agent_mod._SESSION_ALLOWED_TOOLS.pop("clear_range", None)

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


def check_regressions():
    """回归自检:锁住已修好的真实缺陷(每个都先断言「缺陷行为不存在」)。

    ① 切实例后重拍快照必须生效 ② 二次确认必须销号 ③ guard 不误杀自己
    ④ 数组参数写扁必须归一 ⑤ execute 对非 dict 参数永不抛异常
    """
    import json as _json
    import os
    import sys as _sys
    import threading
    import time as _t

    from app import agent as agent_mod
    from app import runtime
    import app.watcher as watcher
    from app.excel_bridge import ExcelBridge

    # ① 多开 Excel 切换实例后的重拍快照必须真正生效(而不是拿到快照又报「无工作簿」)
    class _FakeWorker:
        def submit(self, fn, timeout=0):
            return fn()

    br = ExcelBridge.__new__(ExcelBridge)
    br.worker = _FakeWorker()
    br.app = object()
    br.host_kind, br.host_label, br._pinned = "excel", "Microsoft Excel", None
    br._alive = lambda: True
    br._switch_to_workbook = lambda: True
    calls = {"n": 0}

    def _snap():
        calls["n"] += 1
        if calls["n"] == 1:
            raise RuntimeError("Excel 忙(非 BridgeError)")   # 走通用 except 分支
        return {"workbook": "测试.xlsx", "path": "C:/测试.xlsx", "sheets": [],
                "active_sheet": "Sheet1", "selection": None,
                "open_workbooks": [], "app_visible": True}

    br._snapshot = _snap
    st = br.status()
    assert calls["n"] == 2, f"未走「切换实例后重拍」路径:{calls}"
    assert st.get("workbook") == "测试.xlsx", f"重拍成功的快照被丢弃:{st}"

    # ② 二次确认无论以何种方式结束都要销号,否则失效的确认条会反复弹回侧边栏
    agent_mod._CONFIRM_WAIT.clear()
    calls2 = {"n": 0}

    def _del_llm(*a, **kw):
        calls2["n"] += 1
        if calls2["n"] == 1:
            return {"content": "", "tool_calls": [{"id": "r1", "type": "function",
                    "function": {"name": "delete_sheet",
                                 "arguments": _json.dumps({"name": "表1"})}}]}
        return {"content": "已处理", "tool_calls": None}

    real_llm = agent_mod.llm.chat_completions
    agent_mod.llm.chat_completions = _del_llm
    cancel = threading.Event()
    ev, done = [], threading.Event()

    def _run():
        try:
            agent_mod.run_agent(FakeBridge(), {"base_url": "http://x", "api_key": "k",
                                               "model": "m"}, [], "删掉表1",
                                Executor(None), ev.append, cancel=cancel)
        finally:
            done.set()

    th = threading.Thread(target=_run, daemon=True)
    th.start()
    sid = None
    for _ in range(50):
        sid = next((e["sid"] for e in ev if e["type"] == "confirm"), None)
        if sid:
            break
        _t.sleep(0.1)
    assert sid, "未发出 confirm 事件"
    cancel.set()                 # 用户点「停止」→ 确认等待被中止
    assert done.wait(10), "中止后 agent 未收尾"
    agent_mod.llm.chat_completions = real_llm
    assert sid not in agent_mod._CONFIRM_WAIT, f"确认等待未销号:{list(agent_mod._CONFIRM_WAIT)}"
    assert agent_mod.pending_confirms() == [], agent_mod.pending_confirms()
    agent_mod._CONFIRM_WAIT.clear()

    # ③ guard 收进程:必须真能枚举出进程名(旧实现调的是 pywin32 里根本不存在的     #     win32process.GetModuleBaseName,于是每个进程都抛异常被 except 吞掉,

    #     「表格全关就收主程序」从未生效),名字要取自「用户启动的 exe」,
    #    且不能误杀 guard 自己 / 自己的 launcher 父进程 / 用户的 python.exe
    killed = []
    guard_pid = os.getpid()
    parent_pid = 777001          # 模拟 onefile 的 launcher 父进程
    want = "sheettalk.exe"      # app_process_name() 取自 launch_exe(),不是 sys.executable
    names = {guard_pid: want, parent_pid: want, 4242: want, 9999: want,
             5555: "python.exe", 6666: "excel.exe"}

    class _Api:
        def OpenProcess(self, access, inherit, pid):
            return pid

        def CloseHandle(self, h):
            pass

        def TerminateProcess(self, h, code):
            killed.append(h)

    class _Proc:
        def EnumProcesses(self):
            return [guard_pid, parent_pid, 4242, 4243, 9999, 5555, 6666]

        def GetModuleFileNameEx(self, h, mod):
            # 实测:模块序号传 1 会报「句柄无效」,必须是 0
            assert mod == 0, "GetModuleFileNameEx 的模块序号必须是 0"
            if h == 4243:
                raise RuntimeError("模拟受保护进程:取不到主模块")
            return "C:\\app\\" + names.get(h, "unknown.exe")

    class _Con:
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        PROCESS_TERMINATE = 0x0001

    saved = {k: _sys.modules.get(k) for k in ("win32api", "win32process", "win32con")}
    real_frozen = runtime.is_frozen
    real_launch = runtime.launch_exe
    real_parent_env = os.environ.get("NUITKA_ONEFILE_PARENT")
    os.environ["NUITKA_ONEFILE_PARENT"] = str(parent_pid)
    runtime.launch_exe = lambda: r"C:\app\SheetTalk.exe"
    _sys.modules.update({"win32api": _Api(), "win32process": _Proc(), "win32con": _Con()})
    try:
        # 开发态(未冻结):与主进程同为 python.exe,绝不能动手
        watcher._stop_app()
        assert killed == [], f"开发态不应收任何进程:{killed}"
        runtime.is_frozen = lambda: True
        watcher._stop_app()
    finally:
        runtime.is_frozen, runtime.launch_exe = real_frozen, real_launch
        if real_parent_env is None:
            os.environ.pop("NUITKA_ONEFILE_PARENT", None)
        else:
            os.environ["NUITKA_ONEFILE_PARENT"] = real_parent_env
        for k, v in saved.items():
            if v is None:
                _sys.modules.pop(k, None)
            else:
                _sys.modules[k] = v
    # 只收「用户启动的 exe」同名的其它进程:自己、launcher 父进程、
    # 以及 python.exe / excel.exe 都不能被碰(旧实现会拿 sys.executable 的
    # python.exe 去匹配,把用户自己的 python 进程也杀掉)
    assert killed == [4242, 9999], f"收进程集合不对(只应含同名的其它进程):{killed}"

    # ③b 真实 API 探一下:pywin32 到底能不能枚举出进程名(旧实现用的
    #     GetModuleBaseName 不存在,模拟桩再完美也跑不出真实效果)
    real_names = dict(watcher._process_names())
    assert real_names, "真实进程枚举一个都没拿到(win32 API 用错了)"
    assert os.path.basename(_sys.executable).lower() in set(real_names.values()), \
        "枚举结果里连本进程都没有:%s" % sorted(set(real_names.values()))[:5]

    # ④ 数组参数写成裸字符串时必须归一(旧行为:逐字符展开,
    #    print_sheets("Sheet1") → ['S','h','e','e','t','1'])
    from app.tools import validate_args as _va

    sch_arr = {"properties": {"sheets": {"type": "array"}}}
    assert _va(sch_arr, {"sheets": "Sheet1"})[0] == {"sheets": ["Sheet1"]}
    assert _va(sch_arr, {"sheets": "Sheet1,Sheet2"})[0] == {"sheets": ["Sheet1", "Sheet2"]}
    assert _va(sch_arr, {"sheets": "A1，B2"})[0] == {"sheets": ["A1", "B2"]}
    assert _va(sch_arr, {"sheets": ["S1"]})[0] == {"sheets": ["S1"]}
    assert _va(sch_arr, {"sheets": ""})[0] == {"sheets": None}   # 空串=未提供,保留默认
    assert _va(sch_arr, {})[0] == {}                              # 未给就未给

    class _FakeBridge:
        def read_range(self, sheet, cell_range, max_cells=500):
            return {"range": cell_range, "values": [[1]]}

        def print_sheets(self, sheets=None, copies=1):
            return {"printed": [str(x) for x in (sheets or [])]}

    _ex = Executor(_FakeBridge())
    assert _ex.execute("print_sheets", {"sheets": "Sheet1"})["printed"] == ["Sheet1"]

    # ④b 系统提示词与注册表不能脱节:提示词里形如工具名的标识符必须真实存在,
    #     占位符必须齐全且能被替换干净 —— 改工具名/加规则时最容易忘的一环
    import re as _re

    from app.agent import PLAN_PROMPT, SYSTEM_TEMPLATE, build_messages
    from app.tools import META_TOOLS as _META
    from app.tools import TOOLS as _TOOLS

    _pname = {t["function"]["name"] for t in _TOOLS} | {t["function"]["name"] for t in _META}
    _argnames = set()      # 参数名 + 枚举值:两者都是提示词里合法会出现的词
    for _t in _TOOLS + _META:
        for _k, _v in (((_t["function"].get("parameters") or {}).get("properties") or {})).items():
            _argnames.add(_k)
            _argnames |= {str(_e) for _e in (_v.get("enum") or [])}
    for _tpl_name, _tpl in (("SYSTEM_TEMPLATE", SYSTEM_TEMPLATE), ("PLAN_PROMPT", PLAN_PROMPT)):
        _flat = _re.sub(r"\{[a-z_]+\}", " ", _tpl)   # 先剔除 {context}/{user_text} 这类占位符
        _toks = set(_re.findall(r"\b[a-z][a-z0-9]*(?:_[a-z0-9]+)+\b", _flat))
        _stray = sorted(_toks - _pname - _argnames)
        assert not _stray, f"{_tpl_name} 引用了不存在的工具名/参数名:{_stray}"
    assert "{context}" in SYSTEM_TEMPLATE and "{selection}" in SYSTEM_TEMPLATE
    assert "{user_text}" in PLAN_PROMPT
    _sysmsg = build_messages(FakeBridge(), [], "看看表")[0]["content"]
    assert "{context}" not in _sysmsg and "{selection}" not in _sysmsg
    assert PLAN_PROMPT.replace("{user_text}", "X").count("{user_text}") == 0
    # 工具调用的关键约束必须在提示词里(靠它保持模型不乱传参)
    for _kw in ("cell_range", "JSON 数组", "read_range", "call_tool", "search_tools"):
        assert _kw in SYSTEM_TEMPLATE, f"系统提示词缺少关键约束:{_kw}"

    # ④c 事件协议三端一致(AGENTS.md:新增事件类型要改四处 —— agent 产出 → chat 翻译 →
    #     Main.qml / addin.html 处理)。QML 的 done 是已知豁免:收尾指示器由 ctrl.busy
    #     驱动,不需要该分支;漏一处就是「某个状态永远不该消失/永远转圈」。
    _root = os.path.dirname(os.path.abspath(__file__))

    def _read(rel):
        with open(os.path.join(_root, *rel.split("/")), encoding="utf-8") as _f:
            return _f.read()

    _agent_src, _chat_src = _read("app/agent.py"), _read("app/chat.py")
    _qml_src, _html_src = _read("qml/Main.qml"), _read("addin/addin.html")
    _a_types = set(_re.findall(r'emit\(\s*\{[^{}]*?"type":\s*"([a-z_]+)"', _agent_src, _re.S))
    _c_types = set(_re.findall(r'\bt\s*==\s*"([a-z_]+)"', _chat_src))
    assert _a_types and _a_types == _c_types, f"agent/chat 事件不匹配:{sorted(_a_types ^ _c_types)}"
    _c_kinds = set(_re.findall(r'"kind":\s*"([a-z_]+)"', _chat_src))
    _q_switch = _re.search(r"function onChatEvent\(ev\)([\s\S]*?)function on", _qml_src)
    assert _q_switch, "Main.qml 里找不到 onChatEvent"
    _q_kinds = set(_re.findall(r'case\s+"([a-z_]+)"', _q_switch.group(1)))
    _h_kinds = set(_re.findall(r'case\s+"([a-z_]+)"', _html_src))
    assert not (_c_kinds - _q_kinds - {"done"}), f"QML 未处理的事件:{sorted(_c_kinds - _q_kinds - {'done'})}"
    assert not (_c_kinds - _h_kinds), f"加载项未处理的事件:{sorted(_c_kinds - _h_kinds)}"

    # ⑤ execute 承诺「永远返回 dict」:参数不是对象也不能抛(旧行为:整轮对话
    #    以「内部错误:dictionary update sequence...」收尾)
    for bad in ("A1:B2", 5, ["A1:B2"], object()):
        r = _ex.execute("read_range", bad)
        assert isinstance(r, dict) and "error" in r, (bad, r)
    # 参数被写成 JSON 字符串(模型常见)时应直接解析成功
    r2 = _ex.execute("read_range", '{"cell_range": "A1:B2"}')
    assert r2.get("range") == "A1:B2", r2

    # ⑥ 开机自启命令:第一个引号片段必须正好是 exe 路径
    #    (整条命令包一对引号 → Windows 把它当成一个不存在的文件名,开机静默失败)
    import importlib.util
    from importlib.util import spec_from_file_location

    sl_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "addin", "sideload.py")
    spec = spec_from_file_location("excelai_sideload_regression", sl_path)
    sl = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(sl)
    parts = [r"C:\Program Files\Py\pythonw.exe", r"D:\proj\main.py", "--watch-guard"]
    cmd = sl._watch_run_command(parts)
    assert cmd.split('"')[1] == parts[0], f"exe 未被单独引起来:{cmd}"
    assert cmd.split('"')[0] == "" and cmd.count('"') == 2, f"引号位置不对:{cmd}"
    # 开发态无空格路径:干脆不需要任何引号
    assert '"' not in sl._watch_run_command([r"C:\py\pythonw.exe", r"D:\p\main.py"])

    # ⑦ 冻结态一律用「用户启动的原始 exe」:onefile 下 sys.executable 指向解包缓存
    #    里并不存在的 python.exe(Nuitka 4.2.2 实测),照它登记开机自启、写桌面
    #    快捷方式、Popen 拉起主程序全都静默失效
    import app.installer as installer

    real_frozen2, real_argv = runtime.is_frozen, _sys.argv
    real_bin = os.environ.pop("NUITKA_ONEFILE_BINARY", None)
    orig = r"D:\Program Files\SheetTalk\ExcelAI.exe"
    try:
        runtime.is_frozen = lambda: True
        _sys.argv = [orig, "--smoke"]
        assert runtime.launch_exe() == os.path.abspath(orig), runtime.launch_exe()
        assert watcher._main_entry() == [os.path.abspath(orig)], watcher._main_entry()
        assert watcher.app_process_name() == "excelai.exe", watcher.app_process_name()
        assert sl._launch_exe() == os.path.abspath(orig), sl._launch_exe()
        assert installer._shortcut_target()[0] == os.path.abspath(orig)
        assert installer._icon_ico_path() == os.path.abspath(orig)
        real_sl_frozen = sl._is_frozen
        sl._is_frozen = lambda: True
        try:
            wcmd_parts = sl._watch_command()
            assert wcmd_parts == [os.path.abspath(orig), "--watch-guard"], wcmd_parts
            # 自启项:exe 单独引号、后面跟 --watch-guard、不含缓存里的 python.exe
            wcmd = sl._watch_run_command()
            assert wcmd.split('"')[1] == os.path.abspath(orig), wcmd
            assert wcmd.endswith(" --watch-guard"), wcmd
            assert "python" not in wcmd.lower(), wcmd
        finally:
            sl._is_frozen = real_sl_frozen
        # 只有拿不到 exe 形态的 argv[0] 时才退回 sys.executable
        _sys.argv = ["some.py"]
        assert runtime.launch_exe() == os.path.abspath(_sys.executable)
    finally:
        runtime.is_frozen, _sys.argv = real_frozen2, real_argv
        if real_bin is not None:
            os.environ["NUITKA_ONEFILE_BINARY"] = real_bin
    if not real_frozen2():
        assert runtime.launch_exe() is None, "开发态不该掺和 exe 路径"

    # ⑧ 工具参数必须「不用猜」:每个参数都要有中文说明(同名同义的在 tools._PARAM_DOCS
    #     里集中维护),且说明里列出的取值清单不能与 bridge 的常量表脱节 ——
    #     chart_type / icon_set 这类参数在 schema 里没有 enum,说明就是唯一依据
    from app.excel_bridge import CHART_TYPES as _CT
    from app.excel_bridge import ExcelBridge as _EB
    from app.tools import META_TOOLS as _MT
    from app.tools import TOOLS as _T

    _undoc = ["%s.%s" % (t["function"]["name"], k)
              for t in _T + _MT
              for k, v in (((t["function"].get("parameters") or {}).get("properties") or {})).items()
              if not (v or {}).get("description")]
    assert not _undoc, "这些参数没有说明,模型只能猜:%s" % _undoc[:8]

    def _param(tool, key):
        for _t in _T:
            if _t["function"]["name"] == tool:
                return _t["function"]["parameters"]["properties"][key]
        raise AssertionError("找不到工具:%s" % tool)

    for _k in _CT:
        assert _k in _param("create_chart", "chart_type")["description"], \
            "chart_type 的说明漏了取值 %s" % _k
    for _k in _EB._CF_ICON_SETS:
        assert _k in _param("set_conditional_format", "icon_set")["description"], \
            "icon_set 的说明漏了取值 %s" % _k

    # ⑧b 说明必须真的随请求下发(而不是只躺在模块里没进 payload)
    _seen = {}
    _real_cc = agent_mod.llm.chat_completions

    def _capture(*a, **kw):
        if kw.get("tools"):
            _seen["tools"] = kw["tools"]
        return {"content": "收到", "tool_calls": None}

    agent_mod.llm.chat_completions = _capture
    try:
        agent_mod.run_agent(FakeBridge(), {"base_url": "http://x", "api_key": "k", "model": "m"},
                            [], "把 C 列宽度调到 20", Executor(None), lambda e: None)
    finally:
        agent_mod.llm.chat_completions = _real_cc
    assert _seen.get("tools"), "请求的 tools 为空"
    # 下发的 schema(真实 payload)里不允许有「没说明的参数」
    _nodoc = ["%s.%s" % (t["function"]["name"], k)
              for t in _seen["tools"]
              for k, v in (((t["function"].get("parameters") or {}).get("properties") or {})).items()
              if not (v or {}).get("description")]
    assert not _nodoc, "下发的 schema 里还有没说明的参数:%s" % _nodoc
    _docstrings = [v["description"] for t in _seen["tools"]
                   for v in (((t["function"].get("parameters") or {}).get("properties") or {})).values()
                   if isinstance(v, dict) and v.get("description")]
    assert len(_docstrings) >= 5, "payload 里的参数说明太少:%s" % _docstrings
    print("PASS 回归:实例切换重拍快照 / 确认销号 / guard 收进程不误杀 /"
          " 数组归一 / 非对象参数不抛异常 / 自启命令引号 / 提示词与三端事件协议一致 /"
          " 冻结态指向原始 exe / 工具参数全有说明且取值清单与 bridge 一致")


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


def check_exemptions():
    """会话豁免的三端契约:agent 存「作用域键+时效」 → 两端渲染可读条目并可撤销。"""
    import os

    from app import agent as agent_mod
    from app.tools import DANGEROUS_TOOLS, STEP_NAMES

    saved = dict(agent_mod._SESSION_ALLOWED_TOOLS)
    try:
        agent_mod.clear_session_allowed()
        assert agent_mod.session_allowed_tools() == [], agent_mod.session_allowed_tools()
        assert agent_mod.session_allowed_items() == [], "空清单应返回空数组"

        # ① 作用域键 → 条目:工具名 / 中文名 / 主参数值 / 剩余秒数(两端直接渲染这四个字段)
        agent_mod._grant_exemption("delete_sheet:Sheet2")
        agent_mod._grant_exemption("clear_range")
        items = agent_mod.session_allowed_items()
        assert [i["key"] for i in items] == ["clear_range", "delete_sheet:Sheet2"], items
        by_key = {i["key"]: i for i in items}
        assert by_key["clear_range"]["arg"] == "" and by_key["clear_range"]["label"] == "清空区域"
        assert by_key["delete_sheet:Sheet2"] == {"key": "delete_sheet:Sheet2",
                                                "tool": "delete_sheet", "label": "删除工作表",
                                                "arg": "Sheet2",
                                                "expires_in": by_key["delete_sheet:Sheet2"]["expires_in"]}
        # 剩余秒数必须在 (0, 总时长] 内:两端据此本地倒计时,发负数或超长都是错的
        assert 0 < by_key["clear_range"]["expires_in"] <= agent_mod.SESSION_ALLOWED_TTL

        # ② 每个危险工具都要有中文名:漏了用户就会在界面上看到英文裸名
        _noname = sorted(t for t in DANGEROUS_TOOLS if STEP_NAMES.get(t, t) == t)
        assert not _noname, "危险工具缺中文名:%s" % _noname

        # ③ 旧会话残留的未知工具键不能崩(回退工具名)
        agent_mod._grant_exemption("no_such_tool:X")
        _labels = {i["key"]: i["label"] for i in agent_mod.session_allowed_items()}
        assert _labels["no_such_tool:X"] == "no_such_tool", _labels

        # ④ 时效:过期的既不再放行、也不再列出来(读清单即回收)
        agent_mod._grant_exemption("delete_sheet:已过期", ttl=-1)
        assert not agent_mod._is_exempt("delete_sheet:已过期"), "过期豁免仍在放行"
        assert "delete_sheet:已过期" not in agent_mod._SESSION_ALLOWED_TOOLS, "过期条目未被回收"
        assert "delete_sheet:已过期" not in agent_mod.session_allowed_tools()
        assert "delete_sheet:已过期" not in [i["key"] for i in agent_mod.session_allowed_items()]
        assert agent_mod._is_exempt("clear_range"), "有效豁免被误判过期"

        # ⑤ 单条撤销:只撤命中那条,其余豁免不受影响;重复/未知键返回 False 而不报错
        agent_mod.clear_session_allowed()
        agent_mod._grant_exemption("delete_sheet:要撤")
        agent_mod._grant_exemption("delete_rows:要留")
        assert agent_mod.revoke_session_allowed("delete_sheet:要撤") is True
        assert agent_mod.revoke_session_allowed("delete_sheet:要撤") is False, "重复撤销应返回未命中"
        assert agent_mod.revoke_session_allowed("查无此键") is False
        assert not agent_mod._is_exempt("delete_sheet:要撤"), "撤过的仍被放行"
        assert [i["key"] for i in agent_mod.session_allowed_items()] == ["delete_rows:要留"], \
            agent_mod.session_allowed_items()

        # ⑥ 一键撤销要真的清空(桌面端与侧边栏读的是同一份状态)
        agent_mod.clear_session_allowed()
        assert agent_mod.session_allowed_tools() == [] \
            and agent_mod.session_allowed_items() == [], "撤销后仍有残留"
    finally:
        agent_mod._SESSION_ALLOWED_TOOLS.clear()
        agent_mod._SESSION_ALLOWED_TOOLS.update(saved)

    # ⑥ 两端接线:接口都用到、条目只用 textContent 写入、剩余时间来自服务端
    #    (arg 是模型给的参数值,拼进 innerHTML 就是注入面)
    root = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(root, "addin", "addin.html"), encoding="utf-8") as f:
        html = f.read()
    for kw in ('id="f-allowed"', 'id="btn-allowed-clear"', 'id="allowed-hint"',
               'id="cb-remember-text"', '"/api/session_allowed"',
               '"/api/session_allowed/clear"', '"/api/session_allowed/revoke"',
               "function renderAllowed", "loadAllowed();", "revokeOne", "cb-x",
               "fmtLeft", "expires_in", "cb-left"):
        assert kw in html, "侧边栏缺少会话豁免接线:%s" % kw
    _body = html[html.index("function renderAllowed"): html.index("function loadAllowed")]
    assert "it.arg" in _body and _body.count("textContent") >= 2, "豁免条目必须用 textContent 写入"
    assert "expires_in" in _body, "豁免条目未渲染剩余时间"
    import re as _re
    assert not _re.search(r"innerHTML\s*\+=\s*[^;]*\b(it|d)\.", _body), \
        "禁止把模型给的条目字段拼进 innerHTML"
    assert not _re.search(r"innerHTML\s*\+=", _body), "豁免条目禁止用 += 拼 HTML"

    with open(os.path.join(root, "qml", "Main.qml"), encoding="utf-8") as f:
        qml = f.read()
    for kw in ("ctrl.sessionAllowed", "ctrl.clearSessionAllowed", "ctrl.revokeSessionAllowed",
               "ctrl.sessionAllowedTtl", "settingsDialog.syncExempt()",
               "onSessionAllowedChanged", "exemptRows"):
        assert kw in qml, "桌面端缺少会话豁免接线:%s" % kw
    with open(os.path.join(root, "app", "qt_controller.py"), encoding="utf-8") as f:
        ctl = f.read()
    for kw in ("def revokeSessionAllowed", "agent.revoke_session_allowed",
               "def clearSessionAllowed", "def _get_session_allowed"):
        assert kw in ctl, "控制器缺少会话豁免接线:%s" % kw
    # 时长数字只在 agent.SESSION_ALLOWED_TTL 定义一份:两端都不许写死分钟数
    assert "30 分钟" not in html, "侧边栏把豁免时长写死了,应使用服务端下发的 ttl"
    assert "30 分钟" not in qml, "桌面端把豁免时长写死了,应使用 ctrl.sessionAllowedTtl"
    print("PASS 会话豁免:作用域键→中文条目+剩余秒数 / 到期自动失效且不再列出 /"
          " 未知工具回退 / 一键撤销清空 / 两端接线且不拼 HTML")


if __name__ == "__main__":
    main()
