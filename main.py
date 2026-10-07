"""表答 SheetTalk - 入口(原生 Qt 桌面应用,Win11 Fluent 风格)。

用法:
  python main.py                    # 启动桌面应用
  python main.py --watch            # 跟随启动:检测到 Excel/WPS 才拉起主程序(单次)
  python main.py --watch-guard      # 跟随启动(常驻):表格在则主程序在,表格全关则收掉
  python main.py --smoke            # 启动 6 秒后自动退出(冒烟测试)
  python main.py --smoke-dialog     # 同上,但自动打开设置对话框,便于截图核对设置 UI
"""
import argparse
import logging
import os
import sys


def setup_logging():
    handlers = [logging.StreamHandler()]
    # 文件日志对 --watch-guard 尤其重要:它由 pythonw 拉起,没有控制台,
    # StreamHandler 的输出全部丢失 —— 没有文件日志就无法排查"跟随启动为什么没生效"。
    # 打包态与开发态都写;失败静默跳过(只影响日志,不影响功能)。
    try:
        log_dir = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "ExcelAI", "logs")
        os.makedirs(log_dir, exist_ok=True)
        handlers.append(logging.FileHandler(os.path.join(log_dir, "app.log"), encoding="utf-8"))
    except Exception:
        pass
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s",
                        handlers=handlers)


def main():
    parser = argparse.ArgumentParser(description="表答 SheetTalk")
    parser.add_argument("--smoke", action="store_true", help="启动数秒后自动退出(冒烟测试)")
    parser.add_argument("--smoke-dialog", action="store_true",
                        help="冒烟测试时自动打开设置对话框(截图用)")
    parser.add_argument("--watch", action="store_true",
                        help="跟随启动:检测到 Excel/WPS 表格在运行时拉起主程序后退出")
    parser.add_argument("--watch-guard", action="store_true",
                        help="跟随启动(常驻):表格在则确保主程序在,表格全关则关闭主程序")
    parser.add_argument("--watch-interval", type=float, default=2.0,
                        help="跟随启动的轮询间隔秒数(默认 2)")
    args, _unknown = parser.parse_known_args()

    setup_logging()

    # 跟随启动:不需要 Qt,独立在后台空转,只有看到表格程序才去拉主程序
    if args.watch or args.watch_guard:
        from app.watcher import run_watch
        run_watch(mode="guard" if args.watch_guard else "once", interval=args.watch_interval)
        return

    from app.qt_app import run_qt_app
    sys.exit(run_qt_app(smoke_timeout=6000 if (args.smoke or args.smoke_dialog) else 0))


if __name__ == "__main__":
    main()
