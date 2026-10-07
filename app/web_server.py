"""内置 HTTP 服务:为 Excel 内的 Office 加载项侧边栏提供页面与 API。

仅监听 127.0.0.1;与桌面端共享同一个表格桥接与对话会话。
"""
import json
import logging
import os
import queue
import struct
import sys
import threading
import zlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

from app import config, llm

log = logging.getLogger("excelai")
_SENTINEL = object()

# 确认请求 SSE 推送:订阅队列 + agent 确认监听桥接
_CONFIRM_SSE_CLIENTS = []
_CONFIRM_SSE_LOCK = threading.Lock()
_CONFIRM_SSE_HOOKED = False


def _confirm_sse_listener(payload):
    """agent.arm_confirm 的监听回调:向所有 SSE 订阅者推送确认请求。"""
    with _CONFIRM_SSE_LOCK:
        for q in list(_CONFIRM_SSE_CLIENTS):
            q.put(dict(payload))


def _hook_confirm_listener():
    global _CONFIRM_SSE_HOOKED
    if _CONFIRM_SSE_HOOKED:
        return
    from app.agent import on_confirm_request
    on_confirm_request(_confirm_sse_listener)
    _CONFIRM_SSE_HOOKED = True


def addin_dir():
    """加载项页面目录:PyInstaller _MEIPASS → Nuitka/独立目录(exe 旁边)→ 项目根。"""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass and os.path.isdir(os.path.join(meipass, "addin")):
        return os.path.join(meipass, "addin")
    from app import runtime
    bundled = os.path.join(runtime.exe_dir(), "addin")
    if os.path.isdir(bundled):
        return bundled
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "addin")


def _png_icon(size):
    """统一设计图标:与窗口/托盘/快捷方式/exe 同源(见 app/icon_design.py)。"""
    from app.icon_design import png_bytes, render
    px, _ = render(size, ss=4 if size <= 64 else 2)
    return png_bytes(px, size)


class _Handler(BaseHTTPRequestHandler):
    server_version = "ExcelAI/1.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        pass

    # ---------- 基础工具 ----------

    def _send(self, body, ctype, code=200, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8", code)

    def _file(self, name, ctype):
        path = os.path.join(addin_dir(), name)
        try:
            with open(path, "rb") as f:
                self._send(f.read(), ctype)
        except OSError:
            self._json({"error": "文件不存在"}, 404)

    def _read_json(self):
        """读取 JSON 请求体;过大(>2MB)排干后回 413 并返回 None(调用方据此终止)。

        排干有上限(16MB):直接关连接会让客户端收到 RST 而看不到 413;
        超过排干上限的荒谬体积才放弃应答,让连接直接断掉。
        """
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            return {}
        if n <= 0:
            return {}
        if n > 2 * 1024 * 1024:   # 对话消息远用不到这么大,防御异常/恶意请求
            drained = 0
            while drained < n and drained <= 16 * 1024 * 1024:
                chunk = self.rfile.read(min(65536, n - drained))
                if not chunk:
                    break
                drained += len(chunk)
            self.close_connection = True
            self._json({"error": "请求体过大"}, 413)
            return None
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception:
            return {}

    def _state(self):
        return self.server.state

    # ---------- GET ----------

    def do_GET(self):
        p = urlparse(self.path).path
        st = self._state()
        if p in ("/", "/addin.html"):
            self._file("addin.html", "text/html; charset=utf-8")
        elif p == "/marked.min.js":
            self._file("marked.min.js", "application/javascript; charset=utf-8")
        elif p == "/purify.min.js":
            self._file("purify.min.js", "application/javascript; charset=utf-8")
        elif p == "/icon-32.png":
            self._send(_png_icon(32), "image/png")
        elif p == "/icon-80.png":
            self._send(_png_icon(80), "image/png")
        elif p == "/api/status":
            s = st["bridge"].status()
            cfg = config.load()
            s["has_api_key"] = bool(cfg.get("api_key")) or "127.0.0.1" in cfg.get("base_url", "") \
                or "localhost" in cfg.get("base_url", "")
            s["model"] = cfg.get("model")
            # 加载项侧边栏跟随桌面端主题(auto/light/dark)
            s["theme_mode"] = cfg.get("theme_mode", "auto")
            s["addin"] = True
            from app import runtime
            s["startup_issues"] = list(runtime.STARTUP_ISSUES)
            self._json(s)
        elif p == "/api/confirm_events":
            # SSE 推送:连接即推当前未决确认,新确认请求毫秒级到达
            _hook_confirm_listener()
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream; charset=utf-8")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Connection", "close")
            self.end_headers()
            self.close_connection = True
            q = queue.Queue()
            with _CONFIRM_SSE_LOCK:
                _CONFIRM_SSE_CLIENTS.append(q)
            try:
                from app.agent import pending_confirms
                for pend in pending_confirms():
                    payload = dict(pend, type="confirm")
                    self.wfile.write(("data: " + json.dumps(payload, ensure_ascii=False) + "\n\n").encode("utf-8"))
                self.wfile.flush()
                while True:
                    try:
                        payload = q.get(timeout=15)
                        self.wfile.write(("data: " + json.dumps(payload, ensure_ascii=False) + "\n\n").encode("utf-8"))
                        self.wfile.flush()
                    except queue.Empty:
                        self.wfile.write(b": ping\n\n")   # 心跳,防中间层断开
                        self.wfile.flush()
                        self.wfile.flush()
            except (ConnectionError, BrokenPipeError, OSError):
                pass
            finally:
                with _CONFIRM_SSE_LOCK:
                    try:
                        _CONFIRM_SSE_CLIENTS.remove(q)
                    except ValueError:
                        pass
            return
        elif p == "/api/session_allowed":
            # 会话豁免清单(确认记录抽屉)
            from app.agent import session_allowed_tools
            self._json({"tools": session_allowed_tools()})
        elif p == "/api/selection":
            # 侧边栏选区条点击预览:读当前框选的地址与内容(截断保护在桥接层)
            self._json(st["bridge"].get_selection())
        elif p == "/api/pending_confirm":
            # 异常恢复:侧边栏页面崩溃/刷新后,重新取回未决断的确认请求再展示
            from app.agent import pending_confirms
            pend = pending_confirms()
            self._json(pend[0] if pend else {})
        elif p == "/api/show":
            # 另一个实例启动时发现自己重复,通过这里让已运行的实例把窗口调到前台
            cb = st.get("show_window")
            if cb:
                cb()
            self._json({"ok": bool(cb)})
        elif p == "/api/settings":
            self._json({"settings": config.load(), "presets": config.PRESETS})
        else:
            self._json({"error": "not found"}, 404)

    # ---------- POST ----------

    def do_POST(self):
        p = urlparse(self.path).path
        st = self._state()
        body = self._read_json()
        if body is None:          # 请求体超限,413 已应答
            return
        try:
            if p == "/api/chat":
                self._chat(body)
            elif p == "/api/connect":
                ok = st["bridge"].connect()
                self._json({"connected": ok, "host_label": st["bridge"].host_label})
            elif p == "/api/launch":
                st["bridge"].launch(body.get("host") or "excel")
                self._json({"ok": True, "host_label": st["bridge"].host_label})
            elif p == "/api/settings":
                s = config.save(_normalize_samples(body))
                self._json({"ok": True, "settings": s})
            elif p == "/api/test_llm":
                self._test_llm(body)
            elif p == "/api/clear":
                st["session"].clear()
                self._json({"ok": True})
            elif p == "/api/cancel":
                # 对齐 Copilot 的 Stop:中止后 agent 会在下一个安全点收尾
                st["session"].cancel()
                self._json({"ok": True})
            elif p == "/api/session_allowed/clear":
                # 一键撤销全部会话豁免
                from app.agent import clear_session_allowed
                clear_session_allowed()
                self._json({"ok": True})
            elif p == "/api/confirm":
                # 侧边栏确认条回调:危险工具二次确认结果(remember=本会话始终允许)
                st["session"].confirm(str(body.get("sid") or ""), bool(body.get("allowed")),
                                      bool(body.get("remember")))
                self._json({"ok": True})
            else:
                self._json({"error": "not found"}, 404)
        except Exception as e:
            log.exception("web api 异常")
            self._json({"error": str(e)}, 500)

    def _test_llm(self, body):
        s = config.load()
        base_url = body.get("base_url") or s.get("base_url", "")
        model = body.get("model") or s.get("model", "")
        api_key = body.get("api_key") if body.get("api_key") is not None else s.get("api_key", "")
        try:
            llm.chat_completions(base_url, api_key, model,
                                 [{"role": "user", "content": "回复:ok"}],
                                 max_tokens=8, timeout=30)
            self._json({"ok": True, "message": "连接成功"})
        except llm.LLMError as e:
            self._json({"ok": False, "message": str(e)})

    def _chat(self, body):
        text = (body.get("message") or "").strip()
        if not text:
            self._json({"error": "消息为空"}, 400)
            return
        q = queue.Queue()
        session = self._state()["session"]

        def run():
            try:
                # with_selection=False:用户在侧边栏移除了选区附件,不注入选区上下文
                session.send(text, q.put,
                             use_selection=bool(body.get("with_selection", True)))
            finally:
                q.put(_SENTINEL)
        threading.Thread(target=run, daemon=True).start()

        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "close")
        self.end_headers()
        self.close_connection = True
        try:
            while True:
                ev = q.get(timeout=1800)
                if ev is _SENTINEL:
                    break
                self.wfile.write(("data: " + json.dumps(ev, ensure_ascii=False) + "\n\n")
                                 .encode("utf-8"))
                self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, TimeoutError):
            pass


def _normalize_samples(body):
    """把 samples 统一成 {icon,text,prompt} 对象数组(桌面端按对象渲染)。

    侧边栏/HTTP 接口可能传字符串数组或畸形数据:在这里单点归一化,
    避免坏形状进入 config 后让两端 UI 渲染出错。
    """
    samples = body.get("samples")
    if samples is None:
        return body
    if not isinstance(samples, list):
        body["samples"] = None
        return body
    norm = []
    for x in samples:
        if isinstance(x, str):
            x = x.strip()
            if x:
                norm.append({"icon": "\uEA80", "text": x, "prompt": x})
        elif isinstance(x, dict):
            text = str(x.get("text") or x.get("prompt") or "").strip()
            prompt = str(x.get("prompt") or x.get("text") or "").strip()
            if text and prompt:
                norm.append({"icon": str(x.get("icon") or "\uEA80"),
                             "text": text, "prompt": prompt})
    body["samples"] = norm or None
    return body


class _Server(ThreadingHTTPServer):
    """加大队列:HTTPServer 默认 request_queue_size=5,多个 WebView 面板与
    guard 探测并发连接时会瞬间塞满 backlog,超出即被拒绝(实测 50 并发 8765
    直接 ConnectionRefused)。128 足够本程序的所有合法并发。"""
    daemon_threads = True
    request_queue_size = 128


def create_server(bridge, session, port=8765, on_show=None):
    """启动加载项服务;端口被占用(多半是已在运行)时返回 None。

    on_show:HTTP 线程收到 /api/show 时调用(必须是线程安全的 Qt 信号发射,
    由调用方负责跨线程,qt_app 传 ctrl.requestShowWindow)。
    """
    try:
        srv = _Server(("127.0.0.1", port), _Handler)
    except OSError:
        log.warning("端口 %d 已被占用,加载项服务未启动(可能另一个实例已在运行)", port)
        return None
    srv.daemon_threads = True
    srv.state = {"bridge": bridge, "session": session, "show_window": on_show}
    threading.Thread(target=srv.serve_forever, daemon=True, name="addin-http").start()
    log.info("加载项侧边栏服务已启动: http://localhost:%d/addin.html", port)
    return srv
