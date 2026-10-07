"""框架无关的对话会话:历史管理、流式节流、事件翻译、中止。

桌面端(Qt)与 Excel 内加载项(HTTP)共用同一份会话与表格桥接。
事件 dict 通过 emit 回调推送,kind 取值:
  user_msg / plan / text_start / delta / note_flush / step / step_result /
  status / error / cancelled / done / toast / need_settings
"""
import threading
import time

from app import config
from app.agent import run_agent
from app.tools import Executor


def _needs_key(settings):
    base = settings.get("base_url", "")
    return not settings.get("api_key") and "127.0.0.1" not in base and "localhost" not in base


class ChatSession:
    def __init__(self, bridge):
        self.bridge = bridge
        self.executor = Executor(bridge)
        self.history = []
        self._lock = threading.Lock()
        # 中止信号:UI 点「停止」时置位,agent 在下一个安全点(调工具之前)收尾
        self._cancel = threading.Event()

    def clear(self):
        self.history = []

    def cancel(self):
        """请求中止当前正在跑的这一轮(可从任意线程调用)。"""
        self._cancel.set()

    def confirm(self, sid, allowed, remember=False):
        """回传危险工具二次确认的结果(桌面 ctrl / HTTP /api/confirm 调用)。

        remember=True 且 allowed=True 时,该工具在本会话内不再重复确认。
        """
        from app.agent import resolve_confirm
        resolve_confirm(str(sid), bool(allowed), bool(remember))

    def send(self, text, emit, use_selection=True):
        """执行一轮对话(阻塞)。忙时发 toast 事件并返回 False。

        use_selection=False 时不把当前选区注入系统提示词(UI 上的选区附件被用户移除时用)。
        """
        text = (text or "").strip()
        if not text:
            return False
        if not self._lock.acquire(blocking=False):
            emit({"kind": "toast", "text": "正在处理上一条请求,请稍候…"})
            return False
        self._cancel.clear()
        try:
            self._run(text, emit, use_selection)
        finally:
            self._lock.release()
        return True

    def _run(self, text, emit, use_selection=True):
        settings = config.load()
        if not settings.get("base_url") or _needs_key(settings):
            emit({"kind": "need_settings"})
            return
        emit({"kind": "user_msg", "text": text, "with_selection": bool(use_selection)})
        self.history.append({"role": "user", "content": text})
        acc = []
        state = {"buf": "", "last": 0.0}

        def flush():
            if state["buf"]:
                emit({"kind": "delta", "text": state["buf"]})
                state["buf"] = ""
            state["last"] = time.monotonic()

        def agent_emit(ev):
            t = ev.get("type")
            if t == "delta":
                acc.append(ev.get("text", ""))
                state["buf"] += ev.get("text", "")
                if time.monotonic() - state["last"] > 0.12:
                    flush()
            elif t == "text_start":
                emit({"kind": "text_start"})
            elif t == "plan":
                flush()
                emit({"kind": "plan", "text": ev.get("text", "")})
            elif t == "note_flush":
                flush()
                emit({"kind": "note_flush"})
            elif t == "status":
                emit({"kind": "status", "text": ev.get("text", "")})
            elif t == "step":
                flush()
                emit({"kind": "step", "sid": ev.get("sid"), "summary": ev.get("summary", ""),
                      "ref": ev.get("ref", "")})
            elif t == "confirm":
                # 危险工具二次确认:UI 展示确认条,用户决定后回传 session.confirm
                flush()
                emit({"kind": "confirm", "sid": ev.get("sid", ""),
                      "tool": ev.get("tool", ""), "summary": ev.get("summary", "")})
            elif t == "step_result":
                emit({"kind": "step_result", "sid": ev.get("sid"),
                      "ok": bool(ev.get("ok")), "brief": ev.get("brief", "")})
            elif t == "error":
                flush()
                emit({"kind": "error", "text": ev.get("message", "")})
            elif t == "cancelled":
                # 中止时已流出的半截正文要落库,否则下一轮上下文会莫名其妙断掉
                flush()
                content = "".join(acc).strip()
                if content:
                    self.history.append({"role": "assistant", "content": content})
                acc.clear()  # 注意用 clear() 而不是 acc = []:后者会让 acc 变成闭包内局部变量
                emit({"kind": "cancelled"})
            elif t == "done":
                flush()
                content = "".join(acc).strip()
                if content:
                    self.history.append({"role": "assistant", "content": content})
                emit({"kind": "done"})
        try:
            run_agent(self.bridge, settings, self.history[:-1], text, self.executor,
                      agent_emit, cancel=self._cancel, use_selection=use_selection)
        except Exception as e:
            flush()
            emit({"kind": "error", "text": f"内部错误:{e}"})
