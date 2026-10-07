"""把「表答 SheetTalk」侧边栏注册到宿主程序。

- Office(Excel):本地受信任目录侧加载 manifest.xml(原有能力)。
- WPS 表格:WPS 自己的 JS 加载项机制 —— 把 ribbon/main 复制到
  %APPDATA%\\kingsoft\\wps\\jsaddons\\ExcelAI_1.0\\,并在同级 publish.xml 里登记。

用法:
  python addin/sideload.py                # 注册 Office + WPS + 跟随启动(默认,一条命令全装好)
  python addin/sideload.py office         # 只注册 Office
  python addin/sideload.py wps            # 只注册 WPS
  python addin/sideload.py remove         # 移除两侧 + 跟随启动
  python addin/sideload.py remove office  # 只移除 Office
  python addin/sideload.py remove wps     # 只移除 WPS
  python addin/sideload.py watch          # 只登记「跟随启动」(开机后台监控,检测到表格才拉起主程序)
  python addin/sideload.py unwatch        # 取消跟随启动

为什么「跟随启动」要单独一个进程(而不是纯加载项):
  侧边栏是 Office Web 加载项,页面由本机 8765 服务提供,而这个服务只有主程序
  跑着才存在 —— Office 不会替我们启动主程序。所以需要 HKCU\\Run 里一个极轻的
  监控进程(仅枚举进程名,~12ms/次),看到 Excel/WPS 才拉起主程序;
  纯 COM 注册表加载项需要原生 DLL,XLSTART 宏需要放开宏信任,都不可取。

注册后需重启对应软件,再从功能区插入侧边栏。
"""
import os
import shutil
import sys
import time
import xml.etree.ElementTree as ET

try:
    import winreg
except ImportError:
    winreg = None

try:
    import ctypes
except ImportError:
    ctypes = None

HERE = os.path.dirname(os.path.abspath(__file__))
APPDATA = os.environ.get("APPDATA", os.path.expanduser("~"))

# ---------- Office(Excel)受信任目录 ----------
CATALOG = os.path.join(APPDATA, "ExcelAI", "addin-catalog")
# 专用 SMB 共享:\\localhost\C$(管理共享)只对提权令牌开放,而读受信任目录的
# Excel 是普通令牌 —— 即使以管理员注册,Excel 也读不到 C$。这里建一个自己的
# 共享(注册时若不可读会弹一次 UAC 创建,永久生效),Url 对 Office 才真正可达。
SHARE_NAME = "ExcelAI$"
# 新版 Office(约 2024 起)用「带 GUID 的子键 + UNC 路径」登记受信任目录;
# 旧版用「TrustedCatalog 下的值名=目录路径」DWORD 1。两种都写,兼容各版本。
WEF_KEY = r"Software\Microsoft\Office\16.0\Wef\TrustedCatalog"
CATALOGS_KEY = r"Software\Microsoft\Office\16.0\Wef\TrustedCatalogs"
# 本加载项的固定 GUID:重复注册时覆盖同一个子键,不会堆积垃圾
CATALOG_GUID = "{7C3D2A61-5B4E-4F8A-9D1C-2E6B0A7F4D53}"
# Wef 的根键:Cache / AutoInstallAddins 都挂在它下面,与 TrustedCatalog 平级。
# 注意别写成 WEF_KEY + "\\Cache":那会指向 TrustedCatalog\Cache(不存在),清缓存会静默无效。
WEF_BASE = r"Software\Microsoft\Office\16.0\Wef"
# 自动安装钩子:在此键下登记本加载项的 Id,Office 启动时会自己把加载项装上,
# 不用再走「插入 → 获取加载项 → 共享文件夹 → 添加」这些手动步骤。
AUTOINSTALL_KEY = WEF_BASE + r"\AutoInstallAddins"


def _to_unc(path):
    """把本地绝对路径转成 \\localhost\\C$\\... 形式(Office 只认 UNC)。

    - 盘符冒号换成 $:C:\\x -> \\\\localhost\\C$\\x
    - UNC 路径原样返回
    - 其他路径(如网络映射盘)退化为 localhost 前缀 + 去掉冒号
    """
    p = os.path.normpath(path)
    if p.startswith("\\\\"):
        return p
    drive, _, rest = p.partition("\\")
    if len(drive) == 2 and drive[1] == ":":
        tail = rest.replace("/", "\\")
        return "\\\\localhost\\" + drive[0] + "$" + ("\\" + tail if tail else "")
    return "\\\\localhost\\" + p.lstrip("\\").replace(":", "").replace("/", "\\")


def _share_unc():
    """专用共享的 UNC 形式。"""
    return "\\\\localhost\\" + SHARE_NAME


def _unc_readable(path):
    """用当前(普通权限)令牌探测 UNC 是否可读 —— 与 Excel 的令牌权限一致。"""
    return os.path.isdir(path)


def _ensure_share():
    """确保 \\localhost\\ExcelAI$ 指向 CATALOG 且普通权限进程可读;成功返回 True。

    会弹一次 UAC 创建共享(删除旧共享后重建,保证指向当前目录),永久生效。
    """
    if ctypes is None:
        return False
    if _unc_readable(_share_unc()):
        return True
    user = os.environ.get("USERNAME", "")
    inner = ('sc start LanmanServer >nul 2>&1 & '          # Server 服务被禁用会导致建共享失败
             'net share %s /delete /y >nul 2>&1 & '        # 已存在则先删,确保指向当前目录
             'net share "%s"="%s" /grant:"%s",READ'
             % (SHARE_NAME, SHARE_NAME, CATALOG, user))
    try:
        # runas 会触发 UAC;用户拒绝时返回值 <= 32
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", "cmd.exe", "/d /c " + inner, None, 0)
    except Exception:
        return False
    if ret <= 32:
        return False
    end = time.time() + 30
    while time.time() < end:
        if _unc_readable(_share_unc()):
            return True
        time.sleep(0.4)
    return False

# ---------- WPS JS 加载项 ----------
WPS_NAME = "ExcelAI"            # publish.xml 里的 name,同时是文件夹名前缀
WPS_VERSION = "1.0"             # 文件夹名 = name + "_" + version
WPS_DIRNAME = WPS_NAME + "_" + WPS_VERSION
JSADDONS = os.path.join(APPDATA, "kingsoft", "wps", "jsaddons")


def _delete_tree(root, subkey):
    """递归删除子键(注册表 API 不允许直接删有子键的键)。"""
    try:
        k = winreg.OpenKey(root, subkey)
    except OSError:
        return
    while True:
        try:
            child = winreg.EnumKey(k, 0)
        except OSError:
            break
        _delete_tree(root, subkey + "\\" + child)
    try:
        winreg.DeleteKey(root, subkey)
    except OSError:
        pass


def _manifest_id(path):
    """从 manifest.xml 读出 <Id>(OfficeApp 在默认命名空间,标签名要带命名空间)。"""
    try:
        root = ET.parse(path).getroot()
    except Exception as e:
        print("解析 manifest 失败:", e)
        return None
    ns = "{http://schemas.microsoft.com/office/appforoffice/1.1}"
    node = root.find(ns + "Id")
    if node is None or not (node.text or "").strip():
        print("manifest 里找不到 <Id>")
        return None
    return node.text.strip()


def install_autoinstall(manifest_path):
    """把加载项登记进 Office 的自动安装列表,省掉用户手动「添加」。

    不同 Office 版本读取的键形态不一致,所以两种都写(幂等,重复注册无害):
      A) 顶层值:AutoInstallAddins 下「值名 = 加载项 Id」,DWORD 1(官方文档写法)
      B) 子键:  AutoInstallAddins\\<Id>\\AutoInstall = 1
    """
    if winreg is None:
        return False
    addin_id = _manifest_id(manifest_path)
    if not addin_id:
        return False
    ok = False
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, AUTOINSTALL_KEY) as k:
            winreg.SetValueEx(k, addin_id, 0, winreg.REG_DWORD, 1)
        ok = True
    except OSError as e:
        print("写入自动安装值失败:", e)
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                              AUTOINSTALL_KEY + "\\" + addin_id) as k:
            winreg.SetValueEx(k, "AutoInstall", 0, winreg.REG_DWORD, 1)
    except OSError as e:
        print("写入自动安装子键失败:", e)
    if ok:
        print("已登记到 Office 自动安装列表(免手动添加):", addin_id)
    return ok


def remove_autoinstall(manifest_path):
    """撤销自动安装登记(两种键形态都清)。"""
    if winreg is None:
        return
    addin_id = _manifest_id(manifest_path) if manifest_path and \
        os.path.exists(manifest_path) else None
    if not addin_id:
        return
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOINSTALL_KEY,
                            0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, addin_id)
    except (FileNotFoundError, OSError):
        pass
    _delete_tree(winreg.HKEY_CURRENT_USER, AUTOINSTALL_KEY + "\\" + addin_id)
    print("已移除 Office 自动安装登记:", addin_id)


def register_office():
    """注册 Excel 受信任加载项目录。"""
    if winreg is None:
        print("此功能只能在 Windows 上运行")
        return False
    src = os.path.join(HERE, "manifest.xml")
    if not os.path.exists(src):
        print("找不到清单文件:", src)
        return False
    os.makedirs(CATALOG, exist_ok=True)
    shutil.copyfile(src, os.path.join(CATALOG, "manifest.xml"))

    # 清理旧版脚本可能建出的错误嵌套结构(C: → Users → ...)
    _delete_tree(winreg.HKEY_CURRENT_USER, WEF_KEY + "\\C:")

    # 新版布局:TrustedCatalogs\{GUID} 子键,含 Id / Url(UNC 路径) / Flags=1
    unc = _to_unc(CATALOG)
    if not _unc_readable(unc) and _ensure_share():
        # C$ 不可读(非管理员是常态)→ 改用专用共享,Office 才真正读得到
        unc = _share_unc()
        print("  C$ 管理共享对 Excel 不可读,已改用专用共享:", unc)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, CATALOGS_KEY + "\\" + CATALOG_GUID) as sub:
        winreg.SetValueEx(sub, "Id", 0, winreg.REG_SZ, CATALOG_GUID)
        winreg.SetValueEx(sub, "Url", 0, winreg.REG_SZ, unc)
        winreg.SetValueEx(sub, "Flags", 0, winreg.REG_DWORD, 1)

    # 旧版布局:TrustedCatalog 下建「值名称 = 目录路径」的 DWORD 1(兼容老版本 Office)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, WEF_KEY) as root:
        winreg.SetValueEx(root, CATALOG, 0, winreg.REG_DWORD, 1)

    # 自动安装:登记后 Office 会自己装,省掉用户手动「插入 → 获取加载项 → 共享文件夹 → 添加」
    install_autoinstall(os.path.join(CATALOG, "manifest.xml"))

    # 关键:Office 会把已读过的清单缓存在 Wef\Cache 下(按 Id+版本)。
    # 改过 manifest.xml(比如新加功能区按钮)却不清缓存,Excel 就会继续用旧清单,
    # 新按钮不会出来。这里直接清掉 Excel 的缓存键。
    cleared = _clear_wef_cache()
    if cleared:
        print("  已清理 Office 清单缓存(否则会继续用旧版 manifest)")

    print("已注册 Excel 受信任加载项目录:", CATALOG)
    print("  UNC 形式(新版 Office 需要):", unc)
    if not _unc_readable(unc):
        print("  [!] 该 UNC 路径当前不可读,Office 将无法发现本加载项。")
        print("      原因通常是:非管理员下 \\localhost\\C$ 不可读,且专用共享未建成。")
        print("      请重新运行注册,并在 UAC 弹窗里点「是」(创建共享 ExcelAI$)。")
    print("  1. 完全退出 Excel(注意任务栏/后台是否还有残留 EXCEL.EXE 进程)")
    print("  2. 重新打开 Excel,点「插入 → 获取加载项」,左下角选「共享文件夹」")
    print("  3. 点「表答 SheetTalk」→「添加」—— 新版 Office(16.0.20430+)已不理会")
    print("     AutoInstallAddins 自动安装,这一步必须手动点一次,之后一直有效")
    print("  4. 功能区「开始」选项卡最右侧应有「表答」组,内有「侧边栏」按钮")
    return True


def _clear_wef_cache():
    """清理 Office 的 Wef 清单缓存,返回是否真的清掉了东西。

    缓存路径是 Wef\\Cache\\<App>,不是 Wef\\TrustedCatalog\\Cache\\<App>。
    """
    if winreg is None:
        return False
    removed = False
    for app in ("Excel", "Word", "PowerPoint"):
        path = WEF_BASE + "\\Cache\\" + app
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
                while True:
                    try:
                        sub = winreg.EnumKey(k, 0)
                    except OSError:
                        break
                    _delete_tree(winreg.HKEY_CURRENT_USER, path + "\\" + sub)
                    removed = True
        except FileNotFoundError:
            continue
        except OSError:
            continue
    return removed


def remove_office():
    """注销 Excel 受信任目录。"""
    if winreg is None:
        print("此功能只能在 Windows 上运行")
        return
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, WEF_KEY, 0, winreg.KEY_SET_VALUE) as root:
            winreg.DeleteValue(root, CATALOG)
        print("已移除 Office 注册(旧版布局):", CATALOG)
    except FileNotFoundError:
        print("未找到 Office 注册项(旧版布局),无需移除")
    except OSError:
        print("未找到 Office 注册项(旧版布局),无需移除")
    _delete_tree(winreg.HKEY_CURRENT_USER, WEF_KEY + "\\C:")

    # 新版布局:删掉本加载项的 GUID 子键
    _delete_tree(winreg.HKEY_CURRENT_USER, CATALOGS_KEY + "\\" + CATALOG_GUID)
    print("已移除 Office 注册(新版 TrustedCatalogs):", CATALOG_GUID)

    # 撤销自动安装登记 + 清清单缓存,否则 Office 还记着这个加载项
    remove_autoinstall(os.path.join(CATALOG, "manifest.xml"))
    if _clear_wef_cache():
        print("已清理 Office 清单缓存")

    # 专用共享一并撤销(会弹一次 UAC;拒绝也无妨,残留共享无害)
    if ctypes is not None and _unc_readable(_share_unc()):
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", "cmd.exe",
                "/d /c net share %s /delete /y" % SHARE_NAME, None, 0)
            print("已请求删除专用共享(若弹出 UAC 请确认):", _share_unc())
        except Exception:
            pass

# ---------- WPS JS 加载项 ----------

def _wps_entry_attrs():
    """publish.xml 里本加载项的条目属性(缺哪个补哪个,不动别人的条目)。"""
    return {"name": WPS_NAME, "type": "et", "url": "file://",
            "version": WPS_VERSION, "enable": "enable_dev",
            "install": "null", "customDomain": "", "debug": ""}


def _read_publish(root_tag="jsplugins"):
    """读取现有 publish.xml,返回 (ElementTree, root);不存在则新建。"""
    path = os.path.join(JSADDONS, "publish.xml")
    if os.path.exists(path):
        try:
            tree = ET.parse(path)
            return tree, tree.getroot(), path
        except ET.ParseError:
            # 文件损坏:备份后重建,避免整个加载项目录失效
            try:
                shutil.copyfile(path, path + ".bad")
                print("警告:publish.xml 解析失败,已备份为 publish.xml.bad 并重建")
            except OSError:
                pass
    tree = ET.ElementTree(ET.Element(root_tag))
    return tree, tree.getroot(), path


def register_wps():
    """把 WPS 加载项文件复制进 jsaddons 并在 publish.xml 登记。"""
    src = os.path.join(HERE, "wps")
    if not os.path.isdir(src):
        print("找不到 WPS 加载项源目录:", src)
        return False
    dst = os.path.join(JSADDONS, WPS_DIRNAME)
    os.makedirs(dst, exist_ok=True)
    # ribbon.xml 与根 main.js 放文件夹根部;接口函数在 js/main.js(WPS 自动建 index.html 引入)
    for rel in ("ribbon.xml", "main.js", os.path.join("js", "main.js")):
        s, d = os.path.join(src, rel), os.path.join(dst, rel)
        os.makedirs(os.path.dirname(d), exist_ok=True)
        shutil.copyfile(s, d)
    print("已复制 WPS 加载项文件到:", dst)

    tree, root, path = _read_publish()
    if root.tag != "jsplugins":  # 保底:根标签不对就换壳,条目保留
        new_root = ET.Element("jsplugins")
        for child in list(root):
            new_root.append(child)
        tree = ET.ElementTree(new_root)
        root = new_root
    entry = None
    for el in root.findall("jsplugin"):
        if el.get("name") == WPS_NAME:
            entry = el
            break
    if entry is None:
        entry = ET.SubElement(root, "jsplugin")
    for k, v in _wps_entry_attrs().items():
        entry.set(k, v)
    os.makedirs(JSADDONS, exist_ok=True)
    tree.write(path, encoding="utf-8", xml_declaration=True)
    print("已登记 publish.xml:", path)

    _check_oem_ini()
    print()
    print("接下来(仅 WPS 表格):")
    print("  1. 完全退出并重新打开 WPS 表格(WPS 启动时读取 jsaddons)")
    print("  2. 功能区出现「表答」标签 → 点「侧边栏」按钮")
    print("  3. 侧边栏打开前,请保持桌面程序「表答 SheetTalk」在本机运行")
    return True


def remove_wps():
    """从 publish.xml 摘除本加载项并删除其文件夹。"""
    tree, root, path = _read_publish()
    removed = False
    for el in list(root.findall("jsplugin")):
        if el.get("name") == WPS_NAME:
            root.remove(el)
            removed = True
    if os.path.exists(path):
        tree.write(path, encoding="utf-8", xml_declaration=True)
        print("已从 publish.xml 摘除条目:", path)
    elif not removed:
        print("未找到 WPS 注册项,无需移除")
    dst = os.path.join(JSADDONS, WPS_DIRNAME)
    if os.path.isdir(dst):
        shutil.rmtree(dst, ignore_errors=True)
        print("已删除加载项目录:", dst)
    # WPS 自动生成的 index.html 若残留在别处,启动时会按 publish.xml 忽略本条目


def _find_oem_ini():
    """定位 WPS 安装目录下的 office6/cfgs/oem.ini,找不到返回 None。"""
    bases = []
    pf86 = os.environ.get("ProgramFiles(x86)")
    for var in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        v = os.environ.get(var)
        if v:
            bases.append(v)
    for base in bases:
        for sub in (os.path.join(base, "Kingsoft", "WPS Office"),
                    os.path.join(base, "kingsoft", "wps")):
            if not os.path.isdir(sub):
                continue
            # 版本号子目录(如 12.1.0.16929)或直接就是 office6 所在目录
            candidates = [os.path.join(sub, "office6", "cfgs", "oem.ini")]
            try:
                for name in os.listdir(sub):
                    candidates.append(os.path.join(sub, name, "office6", "cfgs", "oem.ini"))
            except OSError:
                pass
            for c in candidates:
                if os.path.isfile(c):
                    return c
    del pf86
    return None


def _check_oem_ini():
    """WPS 12.1.0.16910+ 默认关闭 JS 加载项,提示用户在 oem.ini 打开 JsApiPlugin。

    不自动改写 Program Files 下的文件(多半要管理员权限),只做检测与指引。
    """
    path = _find_oem_ini()
    if not path:
        return
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
    except OSError:
        return
    if "JsApiPlugin=true" in text.replace(" ", ""):
        return
    print()
    print("检测到 WPS:", path)
    print("  该版本默认可能关闭了 JS 加载项。若重启 WPS 后看不到「表答」标签,")
    print("  请用记事本打开上面的 oem.ini,在 [support] 段加一行 JsApiPlugin=true")
    print("  (若已有形如 xSik7Ci…=R8YcSmVl… 的行,删掉它),保存后重启 WPS。")


# ---------- 开机自启(跟随启动监控) ----------

# 「跟随表格程序启动」需要一个常驻载体,所以用 HKCU\...\Run 登记一个隐藏启动项。
# 开机后监控进程自己驻留(只枚举进程名,~12ms/次),只有检测到 Excel/WPS 表格才拉起主程序。
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_VALUE = "ExcelAI_FollowWatch"


def _is_frozen():
    """是否在打包后的独立程序里(PyInstaller 或 Nuitka)。

    sideload.py 以数据文件形式存在,不参与编译:PyInstaller 认 sys.frozen;
    Nuitka 没有 sys.frozen,但编译出的 __main__ 模块带 __compiled__ 标记。
    """
    if getattr(sys, "frozen", False):
        return True
    main = sys.modules.get("__main__")
    return main is not None and hasattr(main, "__compiled__")


def _watch_command():
    """返回监控进程的启动命令(list),并尽量用 pythonw 避免弹出控制台黑窗。"""
    if _is_frozen():
        return [sys.executable, "--watch-guard"]
    root = os.path.dirname(HERE)
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    exe = pyw if os.path.isfile(pyw) else sys.executable
    return [exe, os.path.join(root, "main.py"), "--watch-guard"]


def install_watch():
    if winreg is None:
        print("此功能只能在 Windows 上运行")
        return
    cmd = '"%s"' % " ".join(_watch_command())
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, RUN_VALUE, 0, winreg.REG_SZ, cmd)
    except OSError as e:
        print("写入开机启动项失败:", e)
        return
    print("已添加开机自启(跟随表格启动监控):")
    print("  ", cmd)
    print("  登录后监控会在后台静默运行;检测到 Excel/WPS 表格进程才会拉起主程序。")
    print("  取消请执行: python addin/sideload.py unwatch")


def remove_watch():
    if winreg is None:
        print("此功能只能在 Windows 上运行")
        return
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, RUN_VALUE)
        print("已移除开机自启。")
    except FileNotFoundError:
        print("未找到开机启动项,无需移除。")
    except OSError as e:
        print("移除开机启动项失败:", e)


# ---------- 入口 ----------

def _usage():
    print(__doc__.strip())


def main(argv):
    args = [a.lower() for a in argv]
    if not args:
        # 默认一条命令装齐:加载项(按钮/侧边栏)+ 跟随启动(后端服务),
        # 用户不需要知道 watch 的存在也能获得完整体验。
        register_office()
        print()
        register_wps()
        print()
        install_watch()
        return 0
    if args[0] in ("-h", "--help", "help"):
        _usage()
        return 0
    if args[0] == "remove":
        target = args[1] if len(args) > 1 else "all"
        if target in ("all", "office"):
            remove_office()
        if target in ("all", "wps"):
            remove_wps()
        if target == "all":
            # 加载项都移除了,后台监控再拉起主程序已无意义,一并撤销。
            # (只想关掉侧边栏、保留桌面程序跟随启动的,之后可重新 watch。)
            print()
            remove_watch()
        if target not in ("all", "office", "wps"):
            _usage()
            return 2
        return 0
    if args[0] == "office":
        register_office()
        return 0
    if args[0] == "wps":
        register_wps()
        return 0
    if args[0] in ("watch", "follow"):
        install_watch()
        return 0
    if args[0] in ("unwatch", "unfollow"):
        remove_watch()
        return 0
    _usage()
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))


