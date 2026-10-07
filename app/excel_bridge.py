"""通过 COM 同时驱动 Microsoft Excel 与 WPS 表格。

探测顺序:Excel.Application → Ket.Application(WPS 表格) → ET.Application(WPS 旧版)。
所有 COM 调用固定在同一个专用线程上执行(CoInitialize),上层同步调用。
"""
from __future__ import annotations

import concurrent.futures
import datetime
import queue
import threading
import time

import pythoncom
import win32com.client

HOSTS = [
    ("excel", "Excel.Application", "Microsoft Excel"),
    ("wps", "Ket.Application", "WPS 表格"),
    ("wps", "ET.Application", "WPS 表格(旧版)"),
]

CHART_TYPES = {
    "column": 51, "bar": 57, "line": 4, "pie": 5, "area": 1,
    "scatter": -4169, "doughnut": -4120, "radar": -4151,
}

# 单元格错误值:0x800A0000 + XlError 码(2000/2007/2015/2023/2029/2036/2043/2045)
CELL_ERRORS = {
    2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!",
    2029: "#NAME?", 2036: "#NUM!", 2043: "#N/A", 2045: "#GETTING_DATA",
}

ALIGN = {"left": -4131, "center": -4108, "right": -4152}


class BridgeError(Exception):
    """面向用户的、可读的表格操作错误。"""


class ComWorker:
    """把所有 COM 操作串行地固定到一个已 CoInitialize 的线程。

    关键点:必须为这个 STA 线程泵 Windows 消息。Excel 属于单线程套间,像
    Validation.Add 这种会牵涉 UI/重入的调用,遇到调用方不泵消息就会直接返回
    「发生意外」(简单读写不会暴露这个问题,所以很容易漏掉)。
    """

    def __init__(self):
        self._q = queue.Queue()
        threading.Thread(target=self._run, daemon=True, name="com-worker").start()

    def _run(self):
        pythoncom.CoInitialize()
        try:
            while True:
                try:
                    fn, fut = self._q.get(timeout=0.05)
                except queue.Empty:
                    # 空闲时泵消息,让 Excel 的重入调用有机会被处理
                    try:
                        pythoncom.PumpWaitingMessages()
                    except Exception:
                        pass
                    continue
                try:
                    fut.set_result(fn())
                except BaseException as e:
                    fut.set_exception(e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def submit(self, fn, timeout=90):
        fut = concurrent.futures.Future()
        self._q.put((fn, fut))
        try:
            return fut.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            raise BridgeError("表格程序没有响应,可能弹出了对话框,请切到表格窗口处理后再试。")


def _conv(v):
    """把 COM 返回的单个值转成可 JSON 序列化的值。"""
    if v is None or isinstance(v, (str, bool)):
        return v
    if isinstance(v, int):
        code = v + 2146828288  # v - 0x800A0000(带符号)
        if 2000 <= code <= 2050:
            return CELL_ERRORS.get(code, f"#ERROR({code})")
        return v
    if isinstance(v, float):
        return v
    if isinstance(v, datetime.datetime):
        return v.isoformat(sep=" ")
    if isinstance(v, datetime.date):
        return v.isoformat()
    try:
        return float(v)
    except Exception:
        return str(v)


def _to_rows(v):
    """把 Range.Value 归一化为二维列表(单格/单行/单列都处理)。"""
    if v is None:
        return [[None]]
    if isinstance(v, (str, bool, int, float, datetime.date)):
        return [[_conv(v)]]
    rows = [list(r) if isinstance(r, (tuple, list)) else [r] for r in v]
    return [[_conv(c) for c in r] for r in rows]


def _rgb(hex_color):
    """"FF8800" → COM 使用的 BGR 整数。"""
    s = str(hex_color).lstrip("#")
    r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    return r + (g << 8) + (b << 16)


def _cap_rows(rows, max_cells):
    total = sum(len(r) for r in rows)
    if total <= max_cells:
        return rows, None
    kept, count = [], 0
    for r in rows:
        if count + len(r) > max_cells:
            break
        kept.append(r)
        count += len(r)
    return kept, f"结果已截断:原区域共 {len(rows)} 行,只返回前 {len(kept)} 行。如需更多请缩小范围或分批读取。"


def _dup_row_indexes(ws, start_row, nrows, ncols, cols):
    """找出重复行的下标(相对 start_row,0 基),保留首次出现的那些行。

    分批读(每批 500 行)并逐行建元组键,避免一次性把大区域读进内存。
    """
    dups = []
    seen = set()
    BATCH = 500
    for base in range(0, nrows, BATCH):
        cnt = min(BATCH, nrows - base)
        rng = ws.Range(ws.Cells(start_row + base, 1), ws.Cells(start_row + base + cnt - 1, ncols))
        try:
            vals = rng.Value
        except Exception:
            return dups
        if cnt == 1:                      # 单行时 COM 返回的是标量而非二维数组
            vals = [vals]
        for i, row in enumerate(vals):
            try:
                key = tuple(_conv(row[c - 1]) for c in cols)
            except Exception:
                key = tuple(str(row[c - 1]) for c in cols)
            if key in seen:
                dups.append(base + i)
            else:
                seen.add(key)
    return dups


def _sort_val(v):
    """排序键:数字 < 文本,None 最小;文本用区域整理键(中文按拼音,与 Excel 一致)。"""
    import locale
    v = _conv(v)
    if v is None:
        return (0, 0.0, b"")
    if isinstance(v, bool):
        return (2, 0.0, locale.strxfrm(str(v)))
    if isinstance(v, (int, float)):
        return (1, float(v), b"")
    return (2, 0.0, locale.strxfrm(str(v)))


def _py_sort_rows(ws, addr, spec, has_header):
    """Python 侧排序回退:读区域 → 多键稳定排序(先次要键后主要键)→ 写回。

    中文按用户区域的整理顺序(拼音)排,与 Excel 自身的排序习惯一致。
    注意:写回会把排序行里的公式变成值,所以只作为 COM 排序结果不对时的兜底。
    """
    import locale
    rng = ws.Range(addr)
    rows = _to_rows(rng.Value)
    data = rows[1:] if has_header else rows
    if len(data) > 20000:
        raise BridgeError("区域超过 2 万行,无法回退排序;请缩小范围或改在表格内直接排序")
    old = None
    try:
        old = locale.setlocale(locale.LC_COLLATE)
        locale.setlocale(locale.LC_COLLATE, "")   # 用户默认区域
    except Exception:
        old = None
    try:
        for col, desc in reversed(spec):
            data.sort(key=lambda r: _sort_val(r[col - 1]), reverse=desc)
    finally:
        if old is not None:
            try:
                locale.setlocale(locale.LC_COLLATE, old)
            except Exception:
                pass
    if has_header:
        data = [rows[0]] + data
    rng.Value = data


def _a1_to_r1c1(ref):
    """「Sheet1!A1:B4」→「Sheet1!R1C1:R4C3」(Consolidate 的 Sources 只认 R1C1)。"""
    import re
    ref = str(ref).replace("$", "")
    m = re.match(r"^(?:'?([^'!]+)'?!)?([A-Za-z]+[0-9]+)(?::([A-Za-z]+[0-9]+))?$", ref)
    if not m:
        raise BridgeError(f"无法解析区域引用:{ref}")
    sheet, a, b_ = m.group(1), m.group(2), m.group(3) or m.group(2)

    def cell_rc(c):
        mm = re.match(r"^([A-Za-z]+)([0-9]+)$", c)
        col = 0
        for ch in mm.group(1).upper():
            col = col * 26 + (ord(ch) - 64)
        return int(mm.group(2)), col
    r1, c1 = cell_rc(a)
    r2, c2 = cell_rc(b_)
    name = f"'{sheet}'" if sheet else ""
    return f"{name}!R{r1}C{c1}:R{r2}C{c2}"


class ExcelBridge:
    def __init__(self):
        self.worker = ComWorker()
        self.app = None
        self.host_kind = None
        self.host_label = None
        self._pinned = None      # 用户/界面显式指定的工作簿名;为空则跟随活动工作簿

    # ---------- 连接管理 ----------

    def use_workbook(self, name=None):
        """显式指定要操作的工作簿名(传 None 恢复「跟随当前活动工作簿」)。

        多开 Excel 时用得上:GetActiveObject 只能拿到其中一个实例,
        必须让用户明确指定,否则会「看起来连着、其实在改另一个文件」。
        """
        self._pinned = name or None
        return self._pinned

    def connect(self):
        """探测正在运行的 Excel / WPS,成功返回 True。

        多开场景(几十个无工作簿的后台实例)下,GetActiveObject 拿到的往往是
        「第一个注册进 ROT」的实例 —— 它可能一个工作簿都没开,界面就会一直
        显示「无工作簿」。所以拿到后再去 ROT 里找「带工作簿」的实例,找到就换过去。
        """
        for kind, progid, label in HOSTS:
            try:
                app = self.worker.submit(
                    lambda p=progid: win32com.client.GetActiveObject(p), timeout=10)
            except Exception:
                continue
            try:
                better = self.worker.submit(lambda a=app: self._rot_pick(a), timeout=10)
                if better is not None:
                    app = better
            except Exception:
                pass
            self.app, self.host_kind, self.host_label = app, kind, label
            return True
        # 快速路径全失败(如 class moniker 失效):兜底扫 ROT,带工作簿的实例优先
        try:
            app = self.worker.submit(
                lambda: self._rot_pick(None, require_workbook=False), timeout=10)
        except Exception:
            app = None
        if app is not None:
            self.app = app
            self.host_kind, self.host_label = self._kind_of(app)
            return True
        self.app = self.host_kind = self.host_label = None
        return False

    def _rot_pick(self, current=None, require_workbook=True):
        """在 ROT 里挑一个表格实例:优先「带打开工作簿」的。

        为什么必须扫 ROT:GetActiveObject 只会返回 class moniker 绑定的那一个
        实例;每个 Excel/WPS 实例还会用「文档 moniker」注册进 ROT,只有遍历 ROT
        并 QueryInterface(IID_IDispatch) 才能拿到它们(直接把 PyIUnknown 丢给
        win32com.client.Dispatch 会报 GetTypeInfo 错)。没有 Workbooks 属性的对象
        (Word、Shell 等)读取时会抛异常,一律跳过。
        """
        if current is not None:
            try:
                if int(current.Workbooks.Count) >= 1:
                    return current
            except Exception:
                pass
        try:
            rot = pythoncom.GetRunningObjectTable()
            enum = rot.EnumRunning()
            any_app = None
            for _ in range(64):     # ROT 条目通常个位数,64 足够且有界
                mons = enum.Next(1)
                if not mons:
                    break
                mon = mons[0] if isinstance(mons, (list, tuple)) else mons
                try:
                    disp = rot.GetObject(mon).QueryInterface(pythoncom.IID_IDispatch)
                    cand = win32com.client.Dispatch(disp).Application
                    wc = int(cand.Workbooks.Count)
                except Exception:
                    continue
                if wc >= 1:
                    return cand
                if any_app is None:
                    any_app = cand
            if require_workbook:
                return None
            return any_app
        except Exception:
            return None

    @staticmethod
    def _kind_of(app):
        """由 Application.Name 判断宿主(ROT 兜底路径拿不到 progid 时用)。"""
        try:
            name = str(app.Name or "")
        except Exception:
            name = ""
        if "WPS" in name.upper():
            return "wps", "WPS 表格"
        return "excel", "Microsoft Excel"

    def _switch_to_workbook(self):
        """当前实例没开着工作簿 → 切到 ROT 里「带工作簿」的实例;切不动返回 False。"""
        try:
            cand = self.worker.submit(lambda: self._rot_pick(self.app), timeout=10)
        except Exception:
            return False
        if cand is None or cand is self.app:
            return False
        self.app = cand
        self.host_kind, self.host_label = self._kind_of(cand)
        return True

    def ensure(self):
        if self.app is not None and self._alive():
            return
        if not self.connect():
            raise BridgeError("没有检测到正在运行的 Excel / WPS 表格。请先打开一个表格窗口,然后点“重新连接”。")

    def _alive(self):
        try:
            self.worker.submit(lambda: self.app.Version, timeout=5)
            return True
        except Exception:
            self.app = None
            return False

    def new_workbook(self):
        """在已连接的实例里新建空白工作簿(并把隐藏的实例显出来)。"""
        def _do():
            try:
                self.app.Visible = True
            except Exception:
                pass
            wb = self.app.Workbooks.Add()
            try:
                wb.Activate()
            except Exception:
                pass
            return {"workbook": wb.Name}
        return self._call(_do)

    def launch(self, kind):
        """已连接时直接新建工作簿;否则启动一个新的 Excel / WPS 实例。"""
        if self.app is not None and self._alive():
            return self.new_workbook()
        progid = "Excel.Application" if kind == "excel" else "Ket.Application"
        label = "Microsoft Excel" if kind == "excel" else "WPS 表格"

        def _do():
            app = win32com.client.DispatchEx(progid)
            app.Visible = True
            app.Workbooks.Add()
            return win32com.client.GetActiveObject(progid)

        try:
            self.app = self.worker.submit(_do, timeout=60)
            self.host_kind = "excel" if kind == "excel" else "wps"
            self.host_label = label
        except Exception as e:
            raise BridgeError(f"启动 {label} 失败:可能没有安装。{e}")

    def status(self):
        if (self.app is None or not self._alive()) and not self.connect():
            return {"connected": False, "host": None, "host_label": None, "workbook": None,
                    "sheets": [], "active_sheet": None, "selection": None}
        try:
            info = self.worker.submit(self._snapshot, timeout=20)
        except BridgeError as e:
            # 用户指定的工作簿已关闭等原因 → 原样透出;若只是「当前实例没开着
            # 工作簿」,先试着切到带工作簿的实例再快照一次(多开时 GetActiveObject
            # 经常连到无工作簿的后台实例 —— 这是「检测不到工作簿」的主修复路径)
            if "没有打开的工作簿" in str(e) and self._switch_to_workbook():
                try:
                    info = self.worker.submit(self._snapshot, timeout=20)
                except Exception as e2:
                    return self._no_workbook(str(e2))
            else:
                return self._no_workbook(str(e))
        except Exception:
            if self._switch_to_workbook():
                try:
                    info = self.worker.submit(self._snapshot, timeout=20)
                except Exception:
                    return self._no_workbook("没有打开的工作簿")
            return self._no_workbook("没有打开的工作簿")
        info.update({"connected": True, "host": self.host_kind, "host_label": self.host_label,
                    "pinned": self._pinned})
        return info
    def _no_workbook(self, note):
        """连上了表格程序但拿不到工作簿时的统一返回(note 供界面提示原因)。"""
        return {"connected": True, "host": self.host_kind, "host_label": self.host_label,
                "workbook": None, "sheets": [], "active_sheet": None, "selection": None,
                "open_workbooks": [], "app_visible": None, "note": note}



    # ---------- 内部工具 ----------

    def _call(self, fn, timeout=90):
        try:
            return self.worker.submit(fn, timeout=timeout)
        except BridgeError:
            raise
        except Exception as e:
            # 用 from e 保留原始异常链:COM 报错信息往往很含糊("找不到成员"),
            # 保留 cause 才能在排查时看到真正的调用点。
            raise BridgeError(f"表格操作失败:{e}") from e

    def _wb(self):
        """取得要操作的工作簿。

        依次尝试,保证「操作的表 = 用户看到的表」:
          1. 用户显式指定的名字(use_workbook)
          2. 当前活动窗口所属的工作簿(比 ActiveWorkbook 更贴合用户视线)
          3. ActiveWorkbook
          4. 最近打开的一个
        另外:如果连到的实例整个不可见(app.Visible=False)却带着可见窗口,
        说明这是个后台残留实例,直接报错要求用户明确指定,避免静默改错文件。
        """
        app = self.app
        wb = None
        if self._pinned:
            try:
                wb = app.Workbooks(self._pinned)
            except Exception:
                wb = None
            if wb is None:
                raise BridgeError(
                    f"指定的工作簿「{self._pinned}」已不存在(可能已被关闭或重命名)。"
                    "请在状态栏重新选择当前要操作的表格。")

        if wb is None:
            # ActiveWindow.Parent 才是「用户眼前这个窗口」对应的工作簿;
            # ActiveWorkbook 在多窗口/窗口未聚焦时经常指向别的文件。
            try:
                win = app.ActiveWindow
                if win is not None:
                    wb = win.Parent
            except Exception:
                wb = None
        if wb is None:
            try:
                wb = app.ActiveWorkbook
            except Exception:
                wb = None
        if wb is None:
            try:
                if app.Workbooks.Count >= 1:
                    wb = app.Workbooks(app.Workbooks.Count)
            except Exception:
                wb = None
        if wb is None:
            raise BridgeError("没有打开的工作簿")
        return wb

    def _ws(self, sheet):
        wb = self._wb()
        if sheet:
            try:
                return wb.Worksheets(sheet)
            except Exception:
                raise BridgeError(f"找不到工作表“{sheet}”")
        ws = self.app.ActiveSheet
        if ws is None:
            ws = wb.Worksheets(1)
        return ws

    def _snapshot(self):
        try:
            wb = self._wb()
        except BridgeError:
            # 用户显式指定的工作簿不存在,这个错误必须透出(否则界面会误显示成
            # 「没有打开的工作簿」,让人以为 Excel 没连上)
            raise
        except Exception:
            return {"workbook": None, "path": None, "sheets": [], "active_sheet": None,
                    "selection": None, "open_workbooks": [], "app_visible": None}
        sheets = []
        for i in range(1, min(wb.Worksheets.Count, 30) + 1):
            ws = wb.Worksheets(i)
            ur = ws.UsedRange
            sheets.append({"name": ws.Name, "used_range": ur.Address,
                           "rows": ur.Rows.Count, "cols": ur.Columns.Count})
        try:
            sel = self.app.Selection.Address
        except Exception:
            sel = None
        try:
            path = wb.FullName
        except Exception:
            path = ""
        # 附带「本实例里所有工作簿」与实例可见性:界面据此告诉用户现在改的是哪个文件,
        # 避免多开 Excel 时静默改错表。
        others = []
        try:
            for i in range(1, self.app.Workbooks.Count + 1):
                w = self.app.Workbooks(i)
                others.append({"name": w.Name, "current": (w.Name == wb.Name)})
        except Exception:
            pass
        try:
            visible = bool(self.app.Visible)
        except Exception:
            visible = None
        try:
            active_sheet = wb.ActiveSheet.Name
        except Exception:
            active_sheet = None
        return {"workbook": wb.Name, "path": path, "sheets": sheets,
                "active_sheet": active_sheet,
                "selection": sel, "open_workbooks": others, "app_visible": visible}

    def workbook_context(self):
        """给 LLM 的轻量上下文:概况 + 每个表的表头预览。"""
        return self._call(self._workbook_context, timeout=30)

    def _workbook_context(self):
        ctx = self._snapshot()
        previews = []
        for s in ctx["sheets"][:10]:
            try:
                ws = self._wb().Worksheets(s["name"])
                ur = ws.UsedRange
                nr = min(2, ur.Rows.Count)
                nc = min(15, ur.Columns.Count)
                rng = ws.Range(ws.Cells(ur.Row, ur.Column),
                               ws.Cells(ur.Row + nr - 1, ur.Column + nc - 1))
                previews.append({"sheet": s["name"], "first_rows": _to_rows(rng.Value)})
            except Exception:
                continue
        ctx["sheet_previews"] = previews
        return ctx

    def selection_context(self, max_rows=50, max_cols=15, timeout=20):
        """当前选中区域的内容,作为对话上下文。

        对应 Copilot 的「选中即上下文」体验:用户在 Excel 里框一片区域再提问,
        AI 就知道在说谁。不选或读取失败时返回 None,由调用方决定怎么提示。
        注意必须截断行列:整列/整行选择时 Rows.Count 会上百万,直接读会卡死表格。
        """
        def _do():
            sel = self.app.Selection
            try:
                address = sel.Address
            except Exception:
                # 选中的是图表/形状等非单元格对象:视为没有有效选区
                return None
            ws = sel.Worksheet
            total_rows, total_cols = sel.Rows.Count, sel.Columns.Count
            # 单一单元格只取它自己;多单元格才按截断后的区域取值
            if total_rows == 1 and total_cols == 1:
                rng = sel
                nr = nc = 1
            else:
                nr = min(total_rows, max(1, int(max_rows)))
                nc = min(total_cols, max(1, int(max_cols)))
                rng = ws.Range(sel.Cells(1, 1), sel.Cells(nr, nc))
            return {"sheet": ws.Name, "address": address,
                    "rows": total_rows, "cols": total_cols,
                    "truncated": nr < total_rows or nc < total_cols,
                    "values": _to_rows(rng.Value)}

        try:
            return self._call(_do, timeout=timeout)
        except BridgeError:
            raise
        except Exception:
            return None

    # ---------- 读取 ----------

    def read_range(self, sheet, cell_range, max_cells=500):
        def _do():
            rng = self._ws(sheet).Range(cell_range)
            rows = _to_rows(rng.Value)
            vals, note = _cap_rows(rows, max_cells)
            out = {"range": rng.Address, "rows": len(rows),
                   "cols": len(rows[0]) if rows else 0, "values": vals}
            if note:
                out["note"] = note
            return out
        return self._call(_do)

    def preview_sheet(self, sheet, max_rows=10, max_cols=15):
        def _do():
            ws = self._ws(sheet)
            ur = ws.UsedRange
            nr = min(ur.Rows.Count, max_rows)
            nc = min(ur.Columns.Count, max_cols)
            rng = ws.Range(ws.Cells(ur.Row, ur.Column), ws.Cells(ur.Row + nr - 1, ur.Column + nc - 1))
            return {"sheet": ws.Name, "used_range": ur.Address,
                    "total_rows": ur.Rows.Count, "total_cols": ur.Columns.Count,
                    "values": _to_rows(rng.Value),
                    "note": "预览从使用区域左上角开始,只含前几行"}
        return self._call(_do)

    # ---------- 写入 ----------

    def write_range(self, sheet, start_cell, values):
        if not isinstance(values, list) or not values:
            raise BridgeError("values 必须是非空的二维数组")
        rows = values if isinstance(values[0], list) else [values]
        ncols = max(len(r) for r in rows)
        rows = [list(r) + [None] * (ncols - len(r)) for r in rows]
        if len(rows) * ncols > 100000:
            raise BridgeError("单次写入不能超过 100000 个单元格,请分批写入")
        has_formula = any(isinstance(c, str) and c.startswith("=") for r in rows for c in r)

        def _do():
            ws = self._ws(sheet)
            # 注意:动态分发下 Resize(rows, cols) 会被错误解析成 Item,必须用两参 Range(Cells, Cells)
            anchor = ws.Range(start_cell)
            r0, c0 = anchor.Row, anchor.Column
            rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(rows) - 1, c0 + ncols - 1))
            if has_formula:
                rng.Formula = rows
            else:
                rng.Value = rows
            return {"written": len(rows) * ncols, "range": rng.Address}
        return self._call(_do)

    def autofill_formula(self, sheet, source_cell, target_range):
        def _do():
            ws = self._ws(sheet)
            src = ws.Range(source_cell)
            dest = ws.Range(target_range)
            src.AutoFill(dest, 0)  # xlFillDefault,相对引用自动递增
            return {"filled": dest.Address}
        return self._call(_do)

    # ---------- 图表 / 格式 / 表结构 ----------

    def create_chart(self, sheet, data_range, chart_type="column", title="",
                     anchor="H2", width=480, height=300):
        ct = CHART_TYPES.get(chart_type, 51)

        def _do():
            ws = self._ws(sheet)
            a = ws.Range(anchor)
            left, top = float(a.Left) + 12, float(a.Top) + 12
            try:
                shape = ws.Shapes.AddChart2(-1, ct, left, top, width, height)
            except Exception:
                shape = ws.Shapes.AddChart(ct, left, top, width, height)
            chart = shape.Chart
            chart.SetSourceData(ws.Range(data_range))
            if title:
                chart.HasTitle = True
                chart.ChartTitle.Text = str(title)
            return {"chart": shape.Name, "anchor": anchor, "data_range": data_range}
        return self._call(_do)

    def set_format(self, sheet, cell_range, bold=None, italic=None, font_size=None,
                   font_color=None, fill_color=None, number_format=None, align=None,
                   border=None, wrap_text=None, vertical=None, font_name=None,
                   underline=None, strikethrough=None, indent=None, locked=None,
                   hide_formula=None):
        _WEIGHTS = {"thin": -4139, "medium": -4138, "thick": 4}   # XlBorderWeight
        _VALIGN = {"top": -4160, "middle": -4108, "bottom": -4107}

        def _do():
            rng = self._ws(sheet).Range(cell_range)
            if font_name:
                rng.Font.Name = str(font_name)
            if bold is not None:
                rng.Font.Bold = bool(bold)
            if italic is not None:
                rng.Font.Italic = bool(italic)
            if underline is not None:
                rng.Font.Underline = 2 if underline else -4142   # 单下划线 / 无
            if strikethrough is not None:
                rng.Font.Strikethrough = bool(strikethrough)
            if font_size:
                rng.Font.Size = float(font_size)
            if font_color:
                rng.Font.Color = _rgb(font_color)
            if fill_color:
                rng.Interior.Color = _rgb(fill_color)
            if number_format:
                rng.NumberFormat = str(number_format)
            if align in ALIGN:
                rng.HorizontalAlignment = ALIGN[align]
            if vertical in _VALIGN:
                rng.VerticalAlignment = _VALIGN[vertical]
            if wrap_text is not None:
                rng.WrapText = bool(wrap_text)
            if indent is not None:
                rng.IndentLevel = max(0, min(15, int(indent)))
            if border:
                # 四边框:LineStyle=1(xlContinuous),粗细用 XlBorderWeight 数值(WPS 兼容)
                rng.Borders.LineStyle = 1
                rng.Borders.Weight = _WEIGHTS.get(str(border), -4139)
            if locked is not None:
                rng.Locked = bool(locked)          # 工作表保护后生效
            if hide_formula is not None:
                rng.FormulaHidden = bool(hide_formula)   # 工作表保护后生效
            return {"formatted": rng.Address}
        return self._call(_do)

    def add_sheet(self, name):
        def _do():
            wb = self._wb()
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            target, i = str(name), 2
            while True:
                try:
                    ws.Name = target
                    break
                except Exception:
                    target = f"{name}{i}"
                    i += 1
                    if i > 50:
                        raise
            return {"sheet": ws.Name}
        return self._call(_do)

    def rename_sheet(self, old, new):
        def _do():
            self._wb().Worksheets(old).Name = str(new)
            return {"renamed": f"{old} → {new}"}
        return self._call(_do)

    def delete_sheet(self, name):
        def _do():
            wb = self._wb()
            # DisplayAlerts 设在 Application 层;个别宿主/派发方式下 wb.DisplayAlerts
            # 不可写会直接抛错,所以取别名时一并容错,不要因此中断删除。
            try:
                app = self.app
                old = app.DisplayAlerts
                app.DisplayAlerts = False
            except Exception:
                app, old = None, None
            try:
                wb.Worksheets(name).Delete()
            finally:
                if app is not None:
                    try:
                        app.DisplayAlerts = old
                    except Exception:
                        pass
            return {"deleted": name}
        return self._call(_do)

    def clear_range(self, sheet, cell_range, contents_only=True):
        def _do():
            rng = self._ws(sheet).Range(cell_range)
            addr = rng.Address
            if contents_only:
                rng.ClearContents()
            else:
                rng.Clear()
            return {"cleared": addr}
        return self._call(_do)

    def sort_range(self, sheet, cell_range, column_index=1, descending=False,
                   has_header=True, column_index2=None, descending2=False):
        def _do():
            ws = self._ws(sheet)
            rng = ws.Range(cell_range)
            if column_index < 1 or column_index > rng.Columns.Count:
                raise BridgeError("排序列序号超出区域列数")
            if column_index2 and (int(column_index2) < 1 or int(column_index2) > rng.Columns.Count):
                raise BridgeError("第二排序列序号超出区域列数")
            addr = rng.Address
            spec = [(int(column_index), bool(descending))] + \
                   ([(int(column_index2), bool(descending2))] if column_index2 else [])
            original = _to_rows(rng.Value)      # 整块原始数据(含表头),回退时先还原
            header_row = original[0] if has_header else None
            # Excel 原生排序:Key 必须用单格。但实测部分派发环境会忽略 Header/Key2/Order2
            # 具名参数(表头被排进数据、次键失效)——排完校验表头,不对就回退 Python 排序。
            kw = {"Key1": rng.Cells(1, int(column_index)),
                  "Order1": 2 if descending else 1,
                  "Header": 1 if has_header else 2, "Orientation": 1}
            if column_index2:
                kw["Key2"] = rng.Cells(1, int(column_index2))
                kw["Order2"] = 2 if descending2 else 1
            rng.Sort(**kw)
            used_fallback = False
            if has_header:
                now = _to_rows(ws.Range(addr).Rows(1).Value)
                if now != header_row:
                    # COM 排序把表头排进了数据:先把原始块还原,再走 Python 排序
                    ws.Range(addr).Value = original
                    _py_sort_rows(ws, addr, spec, True)
                    used_fallback = True
            by = "第%d列 %s" % (int(column_index), "降序" if descending else "升序")
            if column_index2:
                by += ",再按第%d列 %s" % (int(column_index2), "降序" if descending2 else "升序")
            out = {"sorted": addr, "by": by}
            if used_fallback:
                out["note"] = "已用回退排序(排序行内的公式会变为值)"
            return out
        return self._call(_do)

    def clear_formats(self, sheet, cell_range):
        """只清格式(边框/颜色/数字格式等),保留内容。"""
        def _do():
            rng = self._ws(sheet).Range(cell_range)
            rng.ClearFormats()
            return {"cleared_formats": rng.Address}
        return self._call(_do)

    def close_workbook(self, name=None, save=False):
        """关闭工作簿(默认不保存);name 省略 = 当前操作的工作簿。"""
        def _do():
            wb = self._wb() if not name else self.app.Workbooks(str(name))
            title = wb.Name
            if save and not wb.Path:
                raise BridgeError("该工作簿从未保存过,无法按保存关闭;请先手动另存。")
            try:
                self.app.DisplayAlerts = False
            except Exception:
                pass
            try:
                wb.Close(SaveChanges=bool(save))
            finally:
                try:
                    self.app.DisplayAlerts = True
                except Exception:
                    pass
            if self._pinned == title:
                self._pinned = None      # 被固定的工作簿已关,恢复跟随活动工作簿
            return {"closed": title, "saved": bool(save)}
        return self._call(_do)

    def get_selection(self, timeout=5):
        """当前选中区域的地址与内容(无选中/读取失败返回提示 dict)。"""
        try:
            sel = self.selection_context(timeout=timeout)
        except Exception:
            sel = None
        if not sel:
            return {"note": "当前没有选中区域"}
        return sel

    # ---------- 条件格式 / 数据验证(Copilot 的高亮与校验能力) ----------

    # Excel 枚举常量写死为数字:是为了同时兼容 Excel 与 WPS,WPS 对部分枚举名支持不佳
    _CF_VALUE_TYPES = {"lowest": 1, "highest": 2, "percentile": 5}
    _CF_CELL_VALUE = 1        # xlCellValue
    _CF_OPERATORS = {"between": 1, "not_between": 2, "equal": 3, "not_equal": 4,
                     "greater": 5, "less": 6, "greater_equal": 7, "less_equal": 8}
    _CF_ICON_SETS = {"3_arrows": 1, "3_arrows_gray": 2, "3_flags": 3,
                     "3_traffic_lights": 4, "3_signs": 5, "4_arrows": 6,
                     "5_quarters": 7, "5_rating": 8}

    def _clear_cf(self, rng):
        """清掉区域上已有的条件格式。不带参数的 Delete 在各版本行为不一致,
        这里先整体删一次,失败再逐条兜底删。"""
        try:
            rng.FormatConditions.Delete()
        except Exception:
            for i in range(rng.FormatConditions.Count, 0, -1):
                try:
                    rng.FormatConditions(i).Delete()
                except Exception:
                    pass

    def set_conditional_format(self, sheet, cell_range, rule, colors=None,
                               operator="greater", value=None, icon_set="3_traffic_lights"):
        """给区域加条件格式。rule:color_scale / data_bar / icon_set /
        cell_value / duplicate / top_bottom。"""
        def _do():
            rng = self._ws(sheet).Range(cell_range)
            addr = rng.Address
            self._clear_cf(rng)
            fc = rng.FormatConditions

            if rule == "color_scale":
                pal = colors or ["F8696B", "FFEB84", "63BE7B"]
                n = len(pal)
                cs = fc.AddColorScale(n)
                for i in range(1, n + 1):
                    c = cs.ColorScaleCriteria(i)
                    # 首尾固定最低/最高,中间按百分位均分
                    if i == 1:
                        c.Type = self._CF_VALUE_TYPES["lowest"]
                    elif i == n:
                        c.Type = self._CF_VALUE_TYPES["highest"]
                    else:
                        c.Type = self._CF_VALUE_TYPES["percentile"]
                        c.Value = int(100.0 * i / (n - 1))
                    c.FormatColor.Color = _rgb(pal[i - 1])
                kind = "%d色色阶" % n
            elif rule == "data_bar":
                db = fc.AddDatabar()
                db.BarColor.Color = _rgb((colors or ["638EC6"])[0])
                kind = "数据条"
            elif rule == "icon_set":
                ic = fc.AddIconSetCondition()
                ic.IconSet = self.app.ActiveWorkbook.IconSets(
                    self._CF_ICON_SETS.get(icon_set, 4))
                kind = "图标集(%s)" % icon_set
            elif rule == "cell_value":
                op = self._CF_OPERATORS.get(operator)
                if op is None:
                    raise BridgeError("不支持的条件运算符:%s" % operator)
                if value is None:
                    raise BridgeError("按值高亮需要提供 value")
                cond = fc.Add(self._CF_CELL_VALUE, op, str(value))
                pal = colors or ["FFC7CE", "9C0006"]
                cond.Interior.Color = _rgb(pal[0])
                if len(pal) > 1:
                    cond.Font.Color = _rgb(pal[1])
                kind = "值 %s %s" % (operator, value)
            elif rule == "duplicate":
                cond = fc.AddUniqueValues()
                cond.DupeUnique = 1          # xlDuplicate
                cond.Interior.Color = _rgb((colors or ["FFC7CE"])[0])
                kind = "重复值高亮"
            elif rule == "top_bottom":
                cond = fc.AddTop10()
                cond.TopBottom = 0 if operator == "bottom" else 1
                cond.Rank = int(value if value is not None else 10)
                cond.Percent = False
                cond.Interior.Color = _rgb((colors or ["C6EFCE"])[0])
                kind = "%s %d 名" % ("后" if operator == "bottom" else "前", cond.Rank)
            else:
                raise BridgeError("不支持的条件格式类型:%s" % rule)

            return {"formatted": addr, "rule": rule, "kind": kind}

        return self._call(_do)

    _VALIDATION_TYPES = {"list": 3, "whole": 1, "decimal": 2,
                         "date": 4, "text_length": 6, "custom": 7}

    def set_validation(self, sheet, cell_range, kind="list", values=None,
                       minimum=None, maximum=None, formula=None, allow_blank=True):
        """给区域加数据验证(下拉列表 / 数值区间 / 文本长度)。"""
        def _do():
            rng = self._ws(sheet).Range(cell_range)
            addr = rng.Address
            t = self._VALIDATION_TYPES.get(kind)
            if t is None:
                raise BridgeError("不支持的数据验证类型:%s" % kind)
            v = rng.Validation
            try:
                v.Delete()          # 区域里已有验证时 Add 会失败,必须先删
            except Exception:
                pass
            # 注意顺序:Delete 之后立刻设置 IgnoreBlank / InCellDropdown 会抛「发生意外」,
            # 必须先 Add 出规则,再回头调属性。
            # 另外不要把 Validation 对象缓存下来跨多次 COM 调用使用 ——
            # 它是隐式对象,在 ComWorker 线程里持有后再调用 Add 会失败(主线程却正常)。
            if kind == "list":
                seq = ",".join(str(x) for x in (values or []))
                if not seq:
                    raise BridgeError("列表验证需要提供 values")
                if len(seq) > 255:
                    raise BridgeError("列表内容超过 255 字符,Excel 不允许直接序列")
                rng.Validation.Add(t, 1, 1, seq)
            elif kind in ("whole", "decimal"):
                rng.Validation.Add(t, 1, 1, str(minimum if minimum is not None else 0),
                                   str(maximum if maximum is not None else 1000000))
            elif kind == "text_length":
                rng.Validation.Add(t, 1, 1, str(minimum if minimum is not None else 0),
                                   str(maximum if maximum is not None else 255))
            elif kind == "date":
                rng.Validation.Add(t, 1, 1, str(minimum or "2000-01-01"),
                                   str(maximum or "2100-01-01"))
            else:  # custom
                rng.Validation.Add(t, 1, 1, str(formula or "TRUE"))

            for attr, val in (("IgnoreBlank", bool(allow_blank)),
                              ("InCellDropdown", True), ("ShowError", True)):
                try:
                    setattr(rng.Validation, attr, val)
                except Exception:
                    pass          # 个别宿主不支持的属性直接忽略,不影响规则本身
            return {"validated": addr, "kind": kind}

        return self._call(_do)

    def auto_filter(self, sheet, cell_range, field=1, criteria=None, operator=None,
                    action="on"):
        """给区域加自动筛选。criteria 为空只打开筛选器;给了条件就按列筛选;
        action=off 取消该工作表的筛选。"""
        def _do():
            ws = self._ws(sheet)
            if str(action) == "off":
                try:
                    ws.AutoFilterMode = False
                except Exception:
                    pass
                return {"filter_off": True}
            rng = ws.Range(cell_range)
            addr = rng.Address
            # 同一区域重复加筛选会报错,先关掉既有的
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
            if field < 1 or field > rng.Columns.Count:
                raise BridgeError("筛选列序号超出区域列数")
            if criteria:
                rng.AutoFilter(Field=int(field), Criteria1=str(criteria))
                kind = "第%d列 条件 %s" % (field, criteria)
            else:
                rng.AutoFilter()
                kind = "筛选下拉"
            return {"filtered": addr, "kind": kind}

        return self._call(_do)

    def set_hyperlink(self, sheet, cell, url, text=None):
        def _do():
            ws = self._ws(sheet)
            anchor = ws.Range(str(cell))
            disp = str(text) if text else str(anchor.Value or "")
            ws.Hyperlinks.Add(Anchor=anchor, Address=str(url), TextToDisplay=disp)
            return {"hyperlinked": anchor.Address}
        return self._call(_do)

    def copy_sheet(self, sheet, name):
        """复制工作表(值+格式+列宽)。

        不用 Worksheet.Copy(After=):实测该具名参数在 win32com 动态派发下不生效,
        副本会落到新工作簿;改用「新建表 + 整表 Cells.Copy」,确定性复制。
        """
        def _do():
            wb = self._wb()
            src = wb.Worksheets(str(sheet))
            new = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            src.Cells.Copy(new.Cells(1, 1))
            target, i = str(name), 2
            while True:
                try:
                    new.Name = target
                    break
                except Exception:
                    target = "%s%d" % (name, i)
                    i += 1
                    if i > 50:
                        raise
            return {"copied_sheet": new.Name, "from": src.Name}
        return self._call(_do, timeout=120)

    def set_visible(self, sheet, kind, target, visible=True):
        """kind=row/column/sheet;target:行号/列字母或序号/工作表名。"""
        def _do():
            wb = self._wb()
            k = str(kind).lower()      # 别复用 kind:闭包内自赋值会 UnboundLocalError
            vis = bool(visible)
            if k == "sheet":
                ws = wb.Worksheets(str(target))
                ws.Visible = -1 if vis else 0      # xlSheetVisible / xlSheetHidden
                return {"visible": vis, "target": "工作表 %s" % target}
            ws = self._ws(sheet)
            if k == "row":
                ws.Rows(int(target)).Hidden = not vis
                label = "第%d行" % int(target)
            elif k == "column":
                col = str(target).strip()
                (ws.Columns(int(col)) if col.isdigit() else ws.Columns(col)).Hidden = not vis
                label = "列%s" % col
            else:
                raise BridgeError("kind 只能是 row/column/sheet")
            return {"visible": vis, "target": label}
        return self._call(_do)

    def clear_conditional_format(self, sheet, cell_range):
        def _do():
            rng = self._ws(sheet).Range(cell_range)
            self._clear_cf(rng)
            return {"cleared_cf": rng.Address}
        return self._call(_do)

    def remove_duplicates(self, sheet, cell_range, columns=None, has_header=True):
        """删除区域中的重复行。columns 为参与比对的列序号列表,空=整行。"""
        def _do():
            ws = self._ws(sheet)
            rng = ws.Range(cell_range)
            addr = rng.Address
            before = rng.Rows.Count
            ncol = rng.Columns.Count
            # 关键:区域上有自动筛选时,被筛掉的行处于隐藏状态,
            # RemoveDuplicates 会把它们一起算进去,结果常常一行都删不掉。
            # 去重前先关掉筛选,保证按数据的真实内容去重。
            try:
                if ws.AutoFilterMode:
                    ws.AutoFilterMode = False
            except Exception:
                pass
            cols = tuple(int(c) for c in columns) if columns else tuple(range(1, ncol + 1))
            # 先用 Excel 原生 RemoveDuplicates。Columns 必须是列序号数组而非逗号字符串。
            native = False
            try:
                rng.RemoveDuplicates(Columns=cols, Header=1 if has_header else 2)
                native = rng.Rows.Count < before
            except Exception:
                native = False
            if not native:
                # 回退:实测部分环境下原生 RemoveDuplicates 会「返回成功但一行没删」,
                # 所以自己读出数据、比对键、从后往前删掉重复行(结果确定可靠)。
                start = rng.Row + (1 if has_header else 0)
                n = min(before - (1 if has_header else 0), 2000)
                dup = _dup_row_indexes(ws, start, n, ncol, cols)
                for r in sorted(dup, reverse=True):
                    try:
                        ws.Rows(start + r).Delete()
                    except Exception:
                        continue
            removed = before - rng.Rows.Count
            return {"deduped": addr, "removed_rows": removed, "remaining": rng.Rows.Count}

        return self._call(_do, timeout=120)

    def create_table(self, sheet, cell_range, name=None, style="TableStyleMedium2",
                     show_totals=False, totals_column=None):
        """把区域转成 Excel 表格对象(ListObject),可选加汇总行。"""
        def _do():
            ws = self._ws(sheet)
            rng = ws.Range(cell_range)
            addr = rng.Address
            # 同一区域重复转表会报「表不能互相重叠」,先把与目标区域相交的表拆掉。
            # 不能只比地址(开了汇总行后表的区域会多出一行),也不依赖 Range.Intersects
            # (win32com 下不可靠),直接用行列坐标判断重叠。
            def _overlaps(a, b):
                ar1, ac1 = a.Row, a.Column
                ar2, ac2 = ar1 + a.Rows.Count - 1, ac1 + a.Columns.Count - 1
                br1, bc1 = b.Row, b.Column
                br2, bc2 = br1 + b.Rows.Count - 1, bc1 + b.Columns.Count - 1
                return not (ar2 < br1 or br2 < ar1 or ac2 < bc1 or bc2 < ac1)

            for i in range(ws.ListObjects.Count, 0, -1):
                try:
                    if not _overlaps(rng, ws.ListObjects(i).Range):
                        continue
                except Exception:
                    continue
                try:
                    ws.ListObjects(i).Unlist()
                except Exception:
                    pass
            # 1 = xlSrcRange, 1 = xlYes(区域第一行是表头)
            lo = ws.ListObjects.Add(1, rng, None, 1)
            if name:
                lo.Name = name
            try:
                lo.TableStyle = style
            except Exception:
                pass          # WPS 对样式名支持有限,失败不影响表格本身
            if show_totals:
                lo.ShowTotals = True
                if totals_column:
                    idx = lo.ListColumns.Count + 1 if totals_column == "last" else int(totals_column)
                    if 1 <= idx <= lo.ListColumns.Count:
                        # 1 = xlTotalsCalculationSum;Copilot 默认给合计行
                        lo.ListColumns(idx).TotalsCalculation = 1
            return {"table": lo.Name, "range": addr,
                    "totals": bool(lo.ShowTotals), "style": lo.TableStyle.Name
                    if hasattr(lo.TableStyle, "Name") else str(lo.TableStyle)}

        return self._call(_do)

    # ---------- 数据透视表(Copilot 的核心统计能力) ----------

    # 透视表枚举:行/列/页/值 字段,以及各种汇总函数
    _PT_ORIENTATION = {"row": 1, "column": 2, "page": 3, "value": 4}
    _PT_FUNCTIONS = {"sum": -4157, "average": -4106, "count": -4112,
                     "count_nums": -4133, "max": -4136, "min": -4135,
                     "product": -4140, "stddev": -4134, "var": -4163}

    def create_pivot(self, sheet, data_range, rows=None, columns=None, values=None,
                     agg="sum", anchor=None, table_name=None):
        """在数据源区域上建数据透视表。

        rows/columns/values 均为字段名(表头)列表;values 为空时默认用第一个非文本列。
        agg:sum / average / count / count_nums / max / min / product / stddev / var。
        """
        def _do():
            wb = self._wb()
            ws = self._ws(sheet)
            rng = ws.Range(data_range)
            # 数据源必须是「工作表名!绝对地址」。注意 Range.Address 在 win32com 下是
            # 属性而非方法,不能写 Address(True, True),只能自己拼工作表名。
            src = "%s!%s" % (ws.Name, rng.Address)
            # 源区域上如果开着自动筛选,建透视表会失败(且会弹出模态对话框,
            # 进而阻塞后续所有 COM 调用)。先关掉筛选再继续。
            try:
                if ws.AutoFilterMode:
                    ws.AutoFilterMode = False
            except Exception:
                pass
            # 表格对象(ListObject)区域不能直接作透视表数据源,会报「找不到成员」,
            # 所以先把压在源区域上的表拆掉(拆表只影响样式,数据不变)。
            for i in range(ws.ListObjects.Count, 0, -1):
                try:
                    lo = ws.ListObjects(i).Range
                    if (lo.Row <= rng.Row + rng.Rows.Count - 1
                            and lo.Row + lo.Rows.Count - 1 >= rng.Row
                            and lo.Column <= rng.Column + rng.Columns.Count - 1
                            and lo.Column + lo.Columns.Count - 1 >= rng.Column):
                        # Unlist 是 ListObject 的方法;Range.Parent 是工作表,
                        # 写成 lo.Parent.Unlist() 会静默失败,表根本没拆掉。
                        ws.ListObjects(i).Unlist()
                except Exception:
                    pass
            # 找落点:默认放在数据源右侧空列,避开源区域与既有透视表
            base = anchor
            if not base:
                # 落点要留足空间:透视表至少要占 3 列,只找到「第一个空列」会让 Excel
                # 报目标区域无效(并弹出模态对话框卡住后续所有 COM 调用)。
                # 这里直接越过「源区域」与「已用区域」的右边界,再留 1 列空隙。
                col = rng.Column + rng.Columns.Count
                try:
                    used = ws.UsedRange
                    edge = used.Column + used.Columns.Count
                    if edge >= col:
                        col = edge + 1
                except Exception:
                    pass
                base = ws.Cells(rng.Row, col)
            else:
                base = ws.Range(base)

            pc = wb.PivotCaches().Create(SourceType=1, SourceData=src)
            # 注意:不要用 wb.PivotTables().Count 来编号 —— 在 win32com 动态派发下
            # PivotTables 这个 collection 的成员访问会抛「找不到成员」,把整个建表带崩。
            # 改用时间戳命名,真撞名了再顺延。
            name = table_name or ("PT%d" % (int(time.time() * 1000) % 100000000))
            pt = None
            for attempt in range(20):
                try:
                    pt = pc.CreatePivotTable(TableDestination=base, TableName=name)
                    break
                except Exception as e:
                    if attempt == 19:
                        raise BridgeError("创建数据透视表失败:%s" % e)
                    name = "%s_%d" % (name, attempt + 2)

            # 不要试图「先按序号把所有字段设成隐藏」:PivotFields("1") 这种数字索引
            # 在 Excel 里是非法调用(报「类 PivotTable 的 PivotFields 方法无效」);
            # 而且新建的透视表本来就没有任何字段进入布局,无需隐藏。
            for nm in (rows or []):
                pt.PivotFields(str(nm)).Orientation = 1
            for nm in (columns or []):
                pt.PivotFields(str(nm)).Orientation = 2

            # 值字段:必须用 AddDataField 一次性完成「设为值字段 + 指定汇总方式」,
            # 分两步给 Orientation / Function 赋值在新版 Excel 上会报「不能设置 Function 属性」
            func = self._PT_FUNCTIONS.get(agg)
            if func is None:
                raise BridgeError("不支持的汇总方式:%s" % agg)
            vnames = values or []
            if not vnames:
                # 没指定就挑第一个非文本列(透视表不能对文本做求和)
                for i in range(1, rng.Columns.Count + 1):
                    try:
                        rng.Cells(2, i).Value
                        vnames = [str(rng.Cells(1, i).Value)]
                        break
                    except Exception:
                        continue
            if not vnames:
                raise BridgeError("没有可用于汇总的数值列,请显式指定 values")
            label = {"sum": "求和", "average": "平均", "count": "计数",
                     "count_nums": "计数(数值)", "max": "最大值", "min": "最小值",
                     "product": "乘积", "stddev": "标准偏差", "var": "方差"}[agg]
            for nm in vnames:
                try:
                    pt.AddDataField(pt.PivotFields(str(nm)), "%s:%s" % (label, nm), func)
                except Exception:
                    # 个别 Excel 会话里 AddDataField 会 E_INVALIDARG(环境相关,同一台
                    # 机器昨天可用今天失败)。尝试两步赋值;仍失败则清掉半成品透视表,
                    # 给出可读错误,避免留下无法使用的空透视表。
                    try:
                        fld = pt.PivotFields(str(nm))
                        fld.Orientation = 4          # xlDataField
                        try:
                            fld.Function = func
                        except Exception:
                            pass
                    except Exception:
                        pass
                    try:
                        pt.TableRange2.Clear()
                    except Exception:
                        pass
                    raise BridgeError(
                        "透视表创建后无法添加值字段(此 Excel 会话的透视表接口异常,"
                        "可能与多实例/版本状态有关)。已清理半成品;可改用 SUMIFS 等公式"
                        "做分组统计,或重启 Excel 后重试。")

            return {"pivot": pt.Name, "source": src, "anchor": base.Address,
                    "rows": rows or [], "columns": columns or [],
                    "values": vnames, "agg": agg,
                    "result_range": pt.TableRange2.Address}

        return self._call(_do, timeout=120)   # 透视表在大数据源上很慢,给足超时

    def auto_fit_columns(self, sheet, cell_range=None):
        def _do():
            ws = self._ws(sheet)
            rng = ws.Range(cell_range) if cell_range else ws.UsedRange
            rng.Columns.AutoFit()
            return {"autofit": rng.Address}
        return self._call(_do)

    # ---------- 结构与对象控制(删图表/插删行列/查找替换/冻结/保护等) ----------

    def delete_chart(self, sheet, name=None):
        """删除图表;name 为空删除该表所有嵌入图表。"""
        def _do():
            ws = self._ws(sheet)
            co = ws.ChartObjects()
            try:
                cnt = co.Count
            except Exception:
                cnt = 0
            n = 0
            if name:
                try:
                    co(str(name)).Delete()
                    n = 1
                except Exception:
                    raise BridgeError(f"找不到图表「{name}」(可先不带 name 调用以删除全部)")
            else:
                for i in range(cnt, 0, -1):
                    try:
                        co(i).Delete()
                        n += 1
                    except Exception:
                        pass
            return {"deleted_charts": n}
        return self._call(_do)

    def find_replace(self, sheet, find, replace="", cell_range=None,
                     whole_cell=False, case_sensitive=False):
        """查找替换。先数出命中数再执行替换(计数上限 1 万,防整表大区域卡死)。"""
        if not str(find or "").strip():
            raise BridgeError("find 不能为空")
        def _do():
            ws = self._ws(sheet)
            rng = ws.Range(cell_range) if cell_range else ws.Cells
            lookat = 1 if whole_cell else 2          # xlWhole / xlPart
            mc = bool(case_sensitive)
            first = rng.Find(What=str(find), LookAt=lookat, MatchCase=mc)
            if first is None:
                return {"replaced": 0, "range": rng.Address}
            addr0 = first.Address
            count = 1
            cur = first
            while count < 10000:
                cur = rng.FindNext(After=cur)
                if cur is None or cur.Address == addr0:
                    break
                count += 1
            rng.Replace(What=str(find), Replacement=str(replace), LookAt=lookat, MatchCase=mc)
            return {"replaced": count, "range": rng.Address}
        return self._call(_do, timeout=120)

    def insert_rows(self, sheet, row, count=1):
        def _do():
            ws = self._ws(sheet)
            r0 = int(row)
            n = max(1, int(count))
            ws.Rows("%d:%d" % (r0, r0 + n - 1)).Insert()
            return {"inserted_rows": n, "at": r0}
        return self._call(_do)

    def delete_rows(self, sheet, row, count=1):
        def _do():
            ws = self._ws(sheet)
            r0 = int(row)
            n = max(1, int(count))
            ws.Rows("%d:%d" % (r0, r0 + n - 1)).Delete()
            return {"deleted_rows": n, "at": r0}
        return self._call(_do)

    def insert_columns(self, sheet, column, count=1):
        def _do():
            ws = self._ws(sheet)
            c0 = int(column)
            n = max(1, int(count))
            ws.Range(ws.Cells(1, c0), ws.Cells(1, c0 + n - 1)).EntireColumn.Insert()
            return {"inserted_cols": n, "at": c0}
        return self._call(_do)

    def delete_columns(self, sheet, column, count=1):
        def _do():
            ws = self._ws(sheet)
            c0 = int(column)
            n = max(1, int(count))
            ws.Range(ws.Cells(1, c0), ws.Cells(1, c0 + n - 1)).EntireColumn.Delete()
            return {"deleted_cols": n, "at": c0}
        return self._call(_do)

    def merge_cells(self, sheet, cell_range, action="merge", center=True):
        def _do():
            rng = self._ws(sheet).Range(cell_range)
            if str(action) == "unmerge":
                rng.UnMerge()
                return {"merged": False, "range": rng.Address}
            rng.Merge()
            if center:
                rng.HorizontalAlignment = -4108       # xlCenter
            return {"merged": True, "range": rng.Address}
        return self._call(_do)

    def set_column_width(self, sheet, column, width):
        def _do():
            ws = self._ws(sheet)
            col = str(column).strip()
            target = ws.Columns(int(col)) if col.isdigit() else ws.Columns(col)
            target.ColumnWidth = float(width)
            return {"width_set": "列%s=%.1f" % (col, float(width))}
        return self._call(_do)

    def set_row_height(self, sheet, row, height):
        def _do():
            ws = self._ws(sheet)
            ws.Rows(int(row)).RowHeight = float(height)
            return {"height_set": "行%d=%.1f" % (int(row), float(height))}
        return self._call(_do)

    def freeze_panes(self, sheet, cell=None):
        """冻结窗格:cell=冻结点右下方第一格(如 B2 冻结首行首列);空=取消冻结。"""
        def _do():
            ws = self._ws(sheet)
            ws.Activate()
            win = self.app.ActiveWindow
            try:
                win.FreezePanes = False
            except Exception:
                pass
            if cell:
                ws.Range(str(cell)).Select()
                win.FreezePanes = True
            return {"frozen": str(cell) if cell else "已取消冻结"}
        return self._call(_do)

    def protect_sheet(self, sheet, password=None, protect=True):
        def _do():
            ws = self._ws(sheet)
            pw = str(password or "")
            if protect:
                ws.Protect(Password=pw)
            else:
                ws.Unprotect(Password=pw)
            return {"protected": bool(protect), "sheet": ws.Name}
        return self._call(_do)

    def copy_range(self, source_range, target_start, sheet=None, target_sheet=None,
                   values_only=False, transpose=False):
        """复制区域。values_only=只粘贴值;transpose=转置粘贴(粘贴为值)。

        值/转置模式不走剪贴板(读值后直接写入):Range.PasteSpecial 在动态派发下
        传参会被拒(实测「方法无效」),读写的确定性反而更高。
        """
        def _do():
            src_ws = self._ws(sheet)
            src = src_ws.Range(source_range)
            dst_ws = self._ws(target_sheet or sheet)
            dst = dst_ws.Range(target_start)
            if values_only or transpose:
                rows = _to_rows(src.Value)      # 公式会被解析成当前值,即「只粘贴值」
                if transpose:
                    rows = [list(r) for r in zip(*rows)]
                ncols = max(len(r) for r in rows)
                rows = [list(r) + [None] * (ncols - len(r)) for r in rows]
                if len(rows) * ncols > 100000:
                    raise BridgeError("单次写入不能超过 100000 个单元格,请分批")
                r0, c0 = dst.Row, dst.Column
                out = dst_ws.Range(dst_ws.Cells(r0, c0),
                                   dst_ws.Cells(r0 + len(rows) - 1, c0 + ncols - 1))
                out.Value = rows
                return {"copied": "%s → %s" % (src.Address, out.Address),
                        "mode": "转置" if transpose else "只粘贴值"}
            src.Copy(Destination=dst)
            return {"copied": "%s → %s" % (src.Address, dst.Address)}
        return self._call(_do)

    def get_cell_formula(self, sheet, cell_range):
        def _do():
            rng = self._ws(sheet).Range(cell_range)
            rows = _to_rows(rng.Formula)
            vals, note = _cap_rows(rows, 500)
            out = {"range": rng.Address, "formulas": vals}
            if note:
                out["note"] = note
            return out
        return self._call(_do)

    def delete_table(self, sheet, name=None):
        """取消表格对象(ListObject→普通区域,数据不动);name 为空取消该表全部。"""
        def _do():
            ws = self._ws(sheet)
            n = 0
            if name:
                try:
                    ws.ListObjects(str(name)).Unlist()
                    n = 1
                except Exception:
                    raise BridgeError(f"找不到表格「{name}」")
            else:
                for i in range(ws.ListObjects.Count, 0, -1):
                    try:
                        ws.ListObjects(i).Unlist()
                        n += 1
                    except Exception:
                        pass
            return {"unlisted": n}
        return self._call(_do)

    # ---------- 工作簿/文件级(打开/另存/导出/打印/属性/保护/名称/计算) ----------

    def open_workbook(self, path):
        def _do():
            import os
            p = os.path.abspath(str(path))
            if not os.path.isfile(p):
                raise BridgeError(f"找不到文件:{p}")
            wb = self.app.Workbooks.Open(p)
            return {"opened": wb.Name}
        return self._call(_do, timeout=120)

    def save_workbook_as(self, path, file_format=None):
        """另存为;file_format 省略按扩展名推断(xlsx=51/xlsm=52/xls=56/csv=6)。"""
        def _do():
            import os
            wb = self._wb()
            p = os.path.abspath(str(path))
            fmt = int(file_format) if file_format else None
            if fmt is None:
                ext = os.path.splitext(p)[1].lower()
                fmt = {".xlsx": 51, ".xlsm": 52, ".xls": 56, ".csv": 6, ".txt": 42}.get(ext)
                if fmt is None:
                    raise BridgeError("无法识别的扩展名,请用 .xlsx/.xlsm/.xls/.csv 或显式传 file_format")
            old = None
            try:
                old = self.app.DisplayAlerts
                self.app.DisplayAlerts = False
            except Exception:
                pass
            try:
                wb.SaveAs(p, FileFormat=fmt)
            finally:
                if old is not None:
                    try:
                        self.app.DisplayAlerts = old
                    except Exception:
                        pass
            return {"saved_as": wb.Name, "path": p}
        return self._call(_do, timeout=120)

    def export_pdf(self, path, sheet=None):
        def _do():
            import os
            wb = self._wb()
            p = os.path.abspath(str(path))
            if not p.lower().endswith(".pdf"):
                p += ".pdf"
            src = self._ws(sheet) if sheet else wb.ActiveSheet
            src.ExportAsFixedFormat(0, p)      # 0 = xlTypePDF
            return {"exported": p}
        return self._call(_do, timeout=180)

    def print_sheets(self, sheets=None, copies=1):
        """把工作表发送到默认打印机(直接打印,注意纸张)。"""
        def _do():
            wb = self._wb()
            names = [str(s) for s in sheets] if sheets else [wb.ActiveSheet.Name]
            cp = max(1, int(copies or 1))
            for n in names:
                wb.Worksheets(n).PrintOut(Copies=cp)
            return {"printed": names, "copies": cp}
        return self._call(_do, timeout=180)

    def set_doc_properties(self, title=None, author=None, subject=None,
                           keywords=None, comments=None):
        def _do():
            wb = self._wb()
            props = wb.BuiltinDocumentProperties
            out = []
            for key, val in (("Title", title), ("Author", author), ("Subject", subject),
                             ("Keywords", keywords), ("Comments", comments)):
                if val is None:
                    continue
                try:
                    props(key).Value = str(val)
                    out.append(key)
                except Exception:
                    pass
            return {"properties_set": out}
        return self._call(_do)

    def get_doc_properties(self):
        def _do():
            wb = self._wb()
            props = wb.BuiltinDocumentProperties
            out = {}
            for key in ("Title", "Author", "Subject", "Keywords", "Comments",
                        "Last Author", "Creation Date", "Last Save Time"):
                try:
                    v = props(key).Value
                    out[key] = str(v) if v is not None else ""
                except Exception:
                    out[key] = ""
            return out
        return self._call(_do)

    def protect_workbook(self, password=None, protect=True):
        def _do():
            wb = self._wb()
            pw = str(password or "")
            if protect:
                wb.Protect(Password=pw, Structure=True, Windows=False)
            else:
                wb.Unprotect(Password=pw)
            return {"protected": bool(protect), "workbook": wb.Name}
        return self._call(_do)

    def define_name(self, name, refers_to):
        def _do():
            wb = self._wb()
            wb.Names.Add(Name=str(name), RefersTo=str(refers_to))
            return {"defined": str(name), "refers_to": str(refers_to)}
        return self._call(_do)

    def get_names(self):
        def _do():
            wb = self._wb()
            out = []
            for i in range(1, wb.Names.Count + 1):
                nm = wb.Names(i)
                out.append({"name": nm.Name, "refers_to": nm.RefersTo})
            return {"names": out}
        return self._call(_do)

    def refresh_all(self):
        def _do():
            self._wb().RefreshAll()
            return {"refreshed": True}
        return self._call(_do, timeout=180)

    def set_calc_mode(self, mode="auto"):
        def _do():
            modes = {"auto": -4105, "manual": -4135, "semiauto": 2}
            m = modes.get(str(mode).lower())
            if m is None:
                raise BridgeError("mode 只能是 auto/manual/semiauto")
            self.app.Calculation = m
            return {"calc_mode": str(mode).lower()}
        return self._call(_do)

    def calculate(self):
        def _do():
            self.app.Calculate()
            return {"calculated": True}
        return self._call(_do, timeout=180)

    # ---------- 工作表级(移动/标签色/网格线/页面/打印区域/拆分/缩放) ----------

    def move_sheet(self, sheet, after=None, before=None, position=None):
        def _do():
            wb = self._wb()
            ws = wb.Worksheets(str(sheet))
            if after:
                ws.Move(After=wb.Worksheets(str(after)))
            elif before:
                ws.Move(Before=wb.Worksheets(str(before)))
            elif position:
                ws.Move(Before=wb.Worksheets(max(1, min(int(position), wb.Worksheets.Count))))
            idx = None
            for i in range(1, wb.Worksheets.Count + 1):
                if wb.Worksheets(i).Name == str(sheet):
                    idx = i
                    break
            return {"moved": str(sheet), "position": idx,
                    "note": "以实际落位为准(Move 参数在个别派发环境可能被忽略)"}
        return self._call(_do)

    def set_tab_color(self, sheet, color=None):
        def _do():
            ws = self._ws(sheet)
            if color:
                ws.Tab.Color = _rgb(color)
            else:
                ws.Tab.ColorIndex = -4142
            return {"tab_color": color or "已清除", "sheet": ws.Name}
        return self._call(_do)

    def set_gridlines(self, sheet, visible=True):
        def _do():
            ws = self._ws(sheet)
            ws.Activate()
            self.app.ActiveWindow.DisplayGridlines = bool(visible)
            return {"gridlines": bool(visible), "sheet": ws.Name}
        return self._call(_do)

    def set_page_setup(self, sheet, orientation=None, fit_to_width=None, paper_size=None):
        def _do():
            ws = self._ws(sheet)
            ps = ws.PageSetup
            out = []
            if orientation in ("portrait", "landscape"):
                ps.Orientation = 1 if orientation == "portrait" else 2
                out.append(orientation)
            if fit_to_width is not None:
                ps.Zoom = False
                ps.FitToPagesWide = int(fit_to_width)
                ps.FitToPagesTall = False
                out.append("宽度缩放到%d页" % int(fit_to_width))
            if paper_size:
                sizes = {"A4": 9, "A3": 8, "B5": 13, "letter": 1}
                up = str(paper_size).upper()
                if up in sizes:
                    ps.PaperSize = sizes[up]
                    out.append(up)
            return {"page_setup": out or "未变更", "sheet": ws.Name}
        return self._call(_do)

    def set_print_area(self, sheet, cell_range=None):
        def _do():
            ws = self._ws(sheet)
            ws.PageSetup.PrintArea = str(cell_range) if cell_range else ""
            return {"print_area": str(cell_range) if cell_range else "已清除", "sheet": ws.Name}
        return self._call(_do)

    def set_print_titles(self, sheet, rows=None, columns=None):
        def _do():
            ws = self._ws(sheet)
            if rows:
                ws.PageSetup.PrintTitleRows = str(rows)      # 如 "1:1"
            if columns:
                ws.PageSetup.PrintTitleColumns = str(columns)
            return {"print_titles": {"rows": rows, "columns": columns}, "sheet": ws.Name}
        return self._call(_do)

    def set_zoom(self, sheet, percent=100):
        def _do():
            ws = self._ws(sheet)
            ws.Activate()
            self.app.ActiveWindow.Zoom = max(10, min(400, int(percent)))
            return {"zoom": self.app.ActiveWindow.Zoom, "sheet": ws.Name}
        return self._call(_do)

    def split_panes(self, sheet, cell=None):
        """拆分窗格:cell 为拆分点右下方第一格;省略 = 取消拆分。"""
        def _do():
            ws = self._ws(sheet)
            ws.Activate()
            win = self.app.ActiveWindow
            win.Split = False
            try:
                win.FreezePanes = False
            except Exception:
                pass
            if cell:
                ws.Range(str(cell)).Select()
                win.Split = True
            return {"split": str(cell) if cell else "已取消拆分", "sheet": ws.Name}
        return self._call(_do)

    # ---------- 区域级扩展(批注/分列/格式刷/显示公式) ----------

    def set_comment(self, sheet, cell, text=""):
        """加/改批注;text 为空字符串时清除批注。"""
        def _do():
            rng = self._ws(sheet).Range(str(cell))
            if str(text) == "":
                rng.ClearComments()
                return {"comment": "已清除", "cell": str(cell)}
            if rng.Comment is None:
                rng.AddComment(str(text))
            else:
                rng.Comment.Text(str(text))
            return {"comment": str(text), "cell": str(cell)}
        return self._call(_do)

    def text_to_columns(self, sheet, cell_range, delimiter=","):
        def _do():
            ws = self._ws(sheet)
            rng = ws.Range(cell_range)
            d = str(delimiter)
            known = d in (",", ";", " ", "\t")
            rng.TextToColumns(DataType=1,               # xlDelimited
                              ConsecutiveDelimiter=True,
                              Comma=(d == ","), Semicolon=(d == ";"),
                              Space=(d == " "), Tab=(d == "\t"),
                              Other=not known, OtherChar=(d if not known else ""))
            return {"split": rng.Address, "delimiter": d}
        return self._call(_do)

    def copy_format(self, source_range, target_range, sheet=None, target_sheet=None):
        """格式刷:读源区域首个单元格的核心格式,套到目标区域(确定性实现)。"""
        def _do():
            src = self._ws(sheet).Range(source_range).Cells(1, 1)
            dst = self._ws(target_sheet or sheet).Range(target_range)
            f = src.Font
            dst.Font.Name = f.Name
            dst.Font.Size = float(f.Size)
            dst.Font.Bold = bool(f.Bold)
            dst.Font.Italic = bool(f.Italic)
            dst.Font.Color = int(f.Color)
            dst.Interior.Color = int(src.Interior.Color)
            dst.NumberFormat = src.NumberFormat
            dst.Borders.LineStyle = int(src.Borders.LineStyle)
            return {"copied_format": "%s → %s" % (src.Address, dst.Address)}
        return self._call(_do)

    def set_display_formulas(self, sheet, show=True):
        def _do():
            ws = self._ws(sheet)
            ws.Activate()
            self.app.ActiveWindow.DisplayFormulas = bool(show)
            return {"display_formulas": bool(show), "sheet": ws.Name}
        return self._call(_do)

    # ---------- 数据分析(高级筛选/分组/合并计算/单变量求解) ----------

    def advanced_filter(self, sheet, cell_range, criteria_range, copy_to=None,
                        unique_only=False):
        def _do():
            ws = self._ws(sheet)
            rng = ws.Range(cell_range)
            crit = ws.Range(criteria_range)
            if copy_to:
                dst = ws.Range(copy_to)
                rng.AdvancedFilter(Action=2, CriteriaRange=crit, CopyToRange=dst,
                                   Unique=bool(unique_only))
                return {"filtered_to": dst.Address}
            rng.AdvancedFilter(Action=1, CriteriaRange=crit, Unique=bool(unique_only))
            return {"filtered_in_place": rng.Address}
        return self._call(_do)

    def group_rows(self, sheet, start_row, end_row=None, ungroup=False):
        def _do():
            ws = self._ws(sheet)
            r0, r1 = int(start_row), int(end_row or start_row)
            rng = ws.Rows("%d:%d" % (r0, r1))
            if ungroup:
                rng.Ungroup()
            else:
                rng.Group()
            return {"grouped": not ungroup, "rows": "%d:%d" % (r0, r1)}
        return self._call(_do, timeout=60)

    def consolidate(self, sources, target_start, function="sum", target_sheet=None,
                    top_row=True, left_column=True):
        """合并计算:sources 为「工作表!区域」列表,如 ["Sheet1!A1:B10"]。"""
        funcs = {"sum": -4157, "average": -4106, "count": -4112, "max": -4136, "min": -4135}
        def _do():
            # Sources 只认 R1C1 样式引用:A1 样式会报「引用无效」,这里统一转换
            src_list = [_a1_to_r1c1(s) for s in (sources or [])]
            if not src_list:
                raise BridgeError("sources 必须是「工作表!区域」的列表,如 [\"Sheet1!A1:B10\"]")
            fn = funcs.get(str(function).lower())
            if fn is None:
                raise BridgeError("不支持的汇总函数:%s(function/sum/average/max/min)" % function)
            dst = self._ws(target_sheet).Range(target_start)
            dst.Consolidate(Sources=src_list, Function=fn,
                            TopRow=bool(top_row), LeftColumn=bool(left_column))
            return {"consolidated": dst.Address, "function": str(function).lower()}
        return self._call(_do, timeout=120)

    def goal_seek(self, sheet, target_cell, target_value, changing_cell):
        def _do():
            ws = self._ws(sheet)
            ok = ws.Range(str(target_cell)).GoalSeek(Goal=float(target_value),
                                                     ChangingCell=ws.Range(str(changing_cell)))
            return {"goal_seek": bool(ok), "target": str(target_cell),
                    "value": ws.Range(str(target_cell)).Value,
                    "changed": str(changing_cell)}
        return self._call(_do)

    # ---------- 图表/透视表进阶(改型/标题/图例/标签/趋势线/刷新/删除) ----------

    def config_chart(self, sheet, name, title=None, chart_type=None, legend=None,
                     data_labels=None):
        def _do():
            ws = self._ws(sheet)
            ch = ws.ChartObjects(str(name)).Chart
            out = []
            if chart_type:
                ch.ChartType = CHART_TYPES.get(str(chart_type), 51)
                out.append("类型=" + str(chart_type))
            if title is not None:
                if str(title) == "":
                    ch.HasTitle = False
                else:
                    ch.HasTitle = True
                    ch.ChartTitle.Text = str(title)
                out.append("标题")
            if legend is not None:
                ch.HasLegend = bool(legend)
                out.append("图例=" + ("开" if legend else "关"))
            if data_labels is not None:
                ch.SeriesCollection(1).HasDataLabels = bool(data_labels)
                out.append("数据标签=" + ("开" if data_labels else "关"))
            return {"configured": str(name), "changes": out}
        return self._call(_do)

    def add_trendline(self, sheet, name, kind="linear"):
        # 枚举值经隔离实例实测:linear=-4132、logarithmic=-4133、power=4(中文界面验证)
        kinds = {"linear": -4132, "logarithmic": -4133, "power": 4}
        def _do():
            ws = self._ws(sheet)
            ch = ws.ChartObjects(str(name)).Chart
            # Trendlines 是方法,须调用后访问集合;Add 的 Type 用位置参数
            # (具名参数会被动态派发丢弃,报「参数无效」)
            t = ch.SeriesCollection(1).Trendlines().Add(kinds.get(str(kind), -4132))
            return {"trendline": t.Name, "kind": str(kind)}
        return self._call(_do)

    def refresh_pivot(self, sheet, name):
        def _do():
            ws = self._ws(sheet)
            ws.PivotTables(str(name)).RefreshTable()
            return {"refreshed": str(name)}
        return self._call(_do, timeout=120)

    def delete_pivot(self, sheet, name):
        def _do():
            ws = self._ws(sheet)
            ws.PivotTables(str(name)).TableRange2.Clear()
            return {"deleted_pivot": str(name)}
        return self._call(_do)

    # ---------- 公式/视图(显示公式) ----------

    def set_display_formulas(self, sheet, show=True):
        def _do():
            ws = self._ws(sheet)
            ws.Activate()
            self.app.ActiveWindow.DisplayFormulas = bool(show)
            return {"display_formulas": bool(show), "sheet": ws.Name}
        return self._call(_do)

    def save_workbook(self):
        def _do():
            wb = self._wb()
            if not wb.Path:
                return {"saved": False, "note": "该工作簿从未保存过,请先手动另存,避免弹出对话框。"}
            wb.Save()
            return {"saved": True, "path": wb.Path}
        return self._call(_do)
