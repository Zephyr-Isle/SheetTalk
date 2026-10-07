"""表答 SheetTalk - 一键安装 / 卸载(小白向:双击 一键安装.bat 即可,也可手动运行)。

步骤:检查环境 → 安装依赖 → 注册 Excel/WPS 功能区工具栏按钮 → 登记开机自启
     → 创建桌面快捷方式 → 启动主程序。

用法:
  python install.py                # 一键安装并启动
  python install.py --uninstall    # 一键卸载(按钮/自启/快捷方式全撤)
  python install.py --no-launch    # 安装完不自动启动

依赖只用 requirements.txt 里已有的,不引入任何新第三方库。
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(ROOT)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


def _banner(title):
    print()
    print("=" * 54)
    print("  " + title)
    print("=" * 54)


def install_deps():
    """装 requirements.txt;已满足时 pip 会秒过。失败不致命(可能已装好)。"""
    print("\n[1/3] 检查依赖(PySide6 / requests / pywin32)…")
    try:
        r = subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
    except Exception as e:
        print("  依赖检查失败:%s" % e)
        return False
    if r.returncode != 0:
        print("  ! pip 返回非零。若依赖早已装好可忽略;否则请检查网络后重试。")
        return False
    print("  依赖就绪。")
    return True


def launch_app():
    """拉起主程序(单实例保护:已在运行则不会重复开窗)。"""
    try:
        from app.watcher import launch_app as _launch
        return _launch()
    except Exception as e:
        print("  自动启动失败(可手动双击桌面快捷方式):%s" % e)
        return False


def main(argv):
    uninstall = "--uninstall" in argv
    no_launch = "--no-launch" in argv
    _banner("表答 SheetTalk · " + ("一键卸载" if uninstall else "一键安装"))

    if sys.version_info < (3, 10):
        print("需要 Python 3.10 及以上,当前版本:%s" % sys.version.split()[0])
        return 1

    if not uninstall:
        # 卸载不强制依赖(仅注册表/文件清理),安装才需要完整依赖
        install_deps()

    print("\n[2/3] %s(Excel/WPS 工具栏按钮、开机自启、桌面快捷方式)…"
          % ("正在移除" if uninstall else "正在注册"))
    try:
        from app import installer
    except Exception as e:
        print("  加载安装模块失败:%s" % e)
        return 1
    ok, msg = installer.remove_all() if uninstall else installer.install_all()
    print(("  ✓ " if ok else "  ✗ ") + msg)

    if ok and uninstall:
        print("\n[3/3] 卸载完成。重启 Excel/WPS 后功能区按钮即消失。")
        return 0
    if not ok:
        return 1

    print("\n[3/3] 完成。桌面会有「表答 SheetTalk」快捷方式;开机自启已登记。")
    if not no_launch:
        print("  正在启动…")
        launch_app()
    return 0


if __name__ == "__main__":
    try:
        code = main(sys.argv[1:])
    except KeyboardInterrupt:
        code = 130
    sys.exit(code)
