"""端到端真机测试:在新建的空白工作簿里跑通全部核心操作。

数据安全:默认【不关闭】任何已存在的工作簿 —— 旧版本的清理逻辑会无差别
关闭所有未保存的工作簿,可能丢用户数据。要顺带关掉上次测试残留的空白
演示工作簿,显式加 --clean(仍跳过已保存的文件与 PERSONAL.XLSB)。
"""
import sys
import time

from app.excel_bridge import ExcelBridge

clean = "--clean" in sys.argv
b = ExcelBridge()
assert b.connect(), "未检测到 Excel/WPS"
print("host:", b.host_label)

# --clean 时才关闭残留的未保存工作簿(不保存),默认一律不碰
if clean:
    def _cleanup():
        closed = []
        for i in range(b.app.Workbooks.Count, 0, -1):
            wb = b.app.Workbooks(i)
            if wb.Path == "" and wb.Name != "PERSONAL.XLSB":
                closed.append(wb.Name)
                wb.Close(SaveChanges=False)
        return closed
    print("cleanup:", b.worker.submit(_cleanup))

# 新建工作簿(测试全程只碰这个新工作簿)
print("new workbook:", b.new_workbook())
time.sleep(0.6)
print("workbook:", b.status()["workbook"])

# 1. 写入数据
print("write:", b.write_range(None, "A1", [
    ["月份", "销售额", "产品"],
    ["2025-01", 128000, "A型"], ["2025-02", 145000, "A型"],
    ["2025-03", 132000, "B型"], ["2025-04", 167000, "B型"],
    ["2025-05", 158000, "A型"], ["2025-06", 189000, "C型"],
    ["2025-07", 176000, "B型"], ["2025-08", 201000, "C型"],
    ["2025-09", 194000, "A型"], ["2025-10", 218000, "C型"],
    ["2025-11", 235000, "B型"], ["2025-12", 242000, "A型"],
]))

# 2. 公式 + 自动填充
print("cell:", b.write_range(None, "D1", [["环比增长"]]))
print("cell:", b.write_range(None, "D3", ["=(B3-B2)/B2"]))
print("fill:", b.autofill_formula(None, "D3", "D3:D13"))

# 3. 格式
print("fmt:", b.set_format(None, "A1:D1", bold=True, fill_color="4472C4", font_color="FFFFFF", align="center"))
print("numfmt:", b.set_format(None, "B2:B13", number_format="#,##0"))
print("pct:", b.set_format(None, "D3:D13", number_format="0.0%"))
print("autofit:", b.auto_fit_columns(None))

# 4. 排序(按销售额降序)
print("sort:", b.sort_range(None, "A1:D13", column_index=2, descending=True, has_header=True))

# 5. 图表
print("chart:", b.create_chart(None, "A1:B13", "column", "月度销售额", anchor="F2"))

# 6. 回读验证
r = b.read_range(None, "A1:D4", max_cells=32)
print("read back:")
for row in r["values"]:
    print("   ", row)

print("save:", b.save_workbook())
print("LIVE TEST OK")
