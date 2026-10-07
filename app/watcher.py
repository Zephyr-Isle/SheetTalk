"""跟随启动:检测到 Excel / WPS 表格在运行时,自动拉起桌面主程序。

背景:加载项侧边栏依赖本机 8765 端口上的服务,只有主程序跑着才可用。
如果让主程序一直常驻,大多数时候它在空转(轮询 COM、占内存)。这里提供一个
轻量监控模式:自己不创建任何 GUI,只低频枚举一次进程名,看到表格程序才拉起主程序。

两种模式:
  once (默认,--watch):看到表格程序 → 拉起主程序 → 监控进程自身退出。
  guard (--watch-guard):常驻双向跟随,表格在则主程序必须在,表格全关则收掉主程序。

对外接口:
  running_spreadsheets() -> dict[str, str]   命中的「进程名 -> 中文名」
  app_running() -> bool                      主程序是否在运行(TCP 探 8765,毫秒级)
  app_responding() -> bool                   主程序 /api/status 可响应(严格,单实例判定用)
  launch_app() -> bool                       拉起主程序(自动避开重复拉起)
  mark_user_quit() / clear_user_quit()       「用户主动退出」标记(托盘退出 ↔ guard 跟随)
  run_watch(mode="once", interval=2.0)       监控主循环

依赖说明:只用到 requirements.txt 里已有的 pywin32(win32ts)与标准库。
pywin32 没有暴露 ToolHelp32Snapshot 封装,但 WTSEnumerateProcesses 一次调用即可
拿到全部进程名(实测 ~12ms),2 秒轮询约占用 0.6% CPU,无需常驻即可接受。
"""
import json
import logging
import os
import subprocess
import sys
import time

from app import config, runtime

log = logging.getLogger("excelai")

# 表格程序的可执行文件名(小写)。wps.exe 是 WPS 的统一入口,
# 它启动表格模块时一定会同时拉起 et.exe,两者都列上以适配不同版本行为。
HOST_PROCESSES = {
    "excel.exe": "Microsoft Excel",
    "et.exe": "WPS 表格",
    "wps.exe": "WPS Office",
}

PORT = 8765


def _main_entry():
    """返回拉起主程序的命令前缀(兼容 PyInstaller / Nuitka 打包态)。"""
    if runtime.is_frozen():
        return [sys.executable]
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return [sys.executable, os.path.join(root, "main.py")]


def running_process_names():
    """返回当前所有可见进程名的小写集合;枚举失败时返回空集(不抛异常)。"""
    try:
        import win32ts
    except Exception:
        return set()
    try:
        # 参数固定为 (WTS_CURRENT_SERVER_HANDLE, 1, 0);其他组合会返回"参数错误"。
        infos = win32ts.WTSEnumerateProcesses(win32ts.WTS_CURRENT_SERVER_HANDLE, 1, 0)
        return {info[2].lower() for info in infos if info[2]}
    except Exception as e:
        log.debug("枚举进程失败:%s", e)
        return set()


def running_spreadsheets():
    """返回当前正在运行的表格程序,形如 {"excel.exe": "Microsoft Excel"}。"""
    names = running_process_names()
    return {p: label for p, label in HOST_PROCESSES.items() if p in names}


def app_running(timeout=2.0):
    """探测主程序是否在运行:对固定端口 8765 做 TCP 连通性检查。

    故意不用 /api/status 的 HTTP 响应来判断,两个坑都实测踩过:
    1) urllib 的 opener 在进程首次请求时缓存代理配置,guard 常驻进程会被
       冻死在过期代理上,永远误判「不在」→ 每 2 秒拉起重复实例;
    2) /api/status 要过 COM 队列取工作簿信息,Excel 忙时响应超过超时,
       同样误判 → 间歇性拉起重复实例。
    TCP 握手由内核完成(与应用线程是否繁忙无关),三个问题一次解决。
    8765 是本程序专属端口(与 addin/manifest.xml 约定一致),能连上即在运行。
    """
    import socket
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=timeout):
            return True
    except OSError:
        return False


def app_responding(timeout=3.0):
    """主程序的 /api/status 是否可响应(带 addin 标识才算)。

    比 app_running 严格:能区分「本程序」和「碰巧占了 8765 的别的程序」。
    但响应要过 COM 队列,Excel 忙时可能超时 —— 只用于单实例判定这类
    一次性的场景,guard 高频轮询请用 app_running。
    """
    import http.client
    try:
        conn = http.client.HTTPConnection("127.0.0.1", PORT, timeout=timeout)
        try:
            conn.request("GET", "/api/status")
            r = conn.getresponse()
            return bool(json.loads(r.read().decode("utf-8")).get("addin"))
        finally:
            conn.close()
    except Exception:
        return False


def launch_app():
    """拉起主程序。若主程序已在运行则不重复拉起。"""
    if app_running():
        log.info("主程序已在运行,无需重复拉起")
        return False
    clear_user_quit()  # 主动拉起 = 用户想用了,撤销「主动退出」标记,恢复跟随
    try:
        # 关键:必须让 GUI 子进程脱离当前控制台(DETACHED_PROCESS)。
        # 否则它会继承 stdout/stderr 句柄,使得调用 --watch 的终端一直阻塞等待,
        # 用户的命令行看起来像"卡住"了。标准输出丢弃,日志本来就走 logging。
        flags = 0
        if os.name == "nt":
            flags = getattr(subprocess, "DETACHED_PROCESS", 0x00000008) | \
                    getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
        subprocess.Popen(_main_entry(),
                         cwd=os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                         stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL, close_fds=True, creationflags=flags)
        log.info("已拉起主程序")
        return True
    except Exception:
        log.exception("拉起主程序失败")
        return False


def run_watch(mode="once", interval=2.0):
    """监控主循环。

    mode="once":表格程序出现即拉起主程序,然后返回(默认,省资源)。
    mode="guard":常驻,表格在→保证主程序在;表格全关→收掉主进程,随后可继续等待。
    """
    interval = max(0.5, float(interval))
    # guard 由 pythonw 启动,没有控制台,这行只进文件日志(main.py 的 setup_logging),
    # 作用是事后确认"监控进程确实启动过"(app.log 里能看到)。
    log.info("跟随启动监控已启动(模式=%s,间隔=%.1fs),等待表格程序…", mode, interval)
    started = False  # guard 模式下用于避免反复拉起/关闭
    quit_skip_logged = False  # 「用户主动退出」只在第一次跳过时记日志,避免刷屏
    while True:
        hosts = running_spreadsheets()
        if hosts:
            label = "、".join(hosts.values())
            if not app_running():
                if user_quit_pending():
                    # 用户从托盘主动退出过:不再拉起(否则关掉又被弹回来),
                    # 直到主程序被再次启动(手动打开 / launch_app 会清掉标记)。
                    if not quit_skip_logged:
                        log.info("用户已主动退出主程序,跟随启动暂停;重新打开主程序后恢复")
                        quit_skip_logged = True
                else:
                    log.info("检测到 %s 已启动", label)
                    launch_app()
                    started = True
            else:
                started = True
                quit_skip_logged = False
        elif started and mode == "guard":
            _stop_app()
            started = False

        if mode != "guard" and started:
            return

        # 等待:按 interval 分片睡,Ctrl+C 能被及时响应(睡眠用短循环而不是一次长 sleep)
        end = time.time() + interval
        while time.time() < end:
            time.sleep(min(0.2, max(0.0, end - time.time())))


# ---------- 「用户主动退出」标记 ----------
# 主程序关窗后转入托盘后台,不会退出进程;真正的退出只发生在托盘菜单「退出」。
# guard 需要区分「用户主动退出」与「意外崩溃/开机未启动」:前者不再拉起,
# 后者照常拉起。标记由 qt_app 在托盘退出时写入,由主程序下次启动 / launch_app 清除。

def quit_flag_path():
    return os.path.join(config.data_dir(), "user_quit.flag")


def mark_user_quit():
    """托盘「退出」时留痕:guard 见到后暂停跟随,不再把窗口弹回来。"""
    try:
        with open(quit_flag_path(), "w", encoding="utf-8") as f:
            f.write(str(time.time()))
    except OSError:
        pass


def user_quit_pending():
    return os.path.isfile(quit_flag_path())


def clear_user_quit():
    """主程序重新启动(手动/launch_app)即视为用户想用了,恢复跟随。"""
    try:
        os.remove(quit_flag_path())
    except OSError:
        pass


def _stop_app():
    """guard 模式:所有表格程序都退出后,主动关掉主程序。"""
    try:
        import win32api
        import win32con
        import win32process
    except Exception:
        return
    exe = (os.path.basename(sys.executable) or "").lower()
    if not exe:
        return
    try:
        pids = win32process.EnumProcesses()
    except Exception:
        return
    for pid in pids:
        try:
            h = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            try:
                name = win32process.GetModuleBaseName(h, win32process.PROCESS_QUERY_LIMITED_INFORMATION)
            finally:
                win32api.CloseHandle(h)
        except Exception:
            continue
        # 源码运行时主进程是 python.exe,无法据此区分,故只在打包态收进程,
        # 开发态让用户自己关窗口,避免误杀其他 python 程序。
        if name and name.lower() == exe and runtime.is_frozen():
            try:
                win32api.TerminateProcess(win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid), 1)
                log.info("表格程序已全部退出,已关闭主程序")
            except Exception:
                pass