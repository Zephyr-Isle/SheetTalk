"""Agent 可用的工具:OpenAI function-calling schema + 对 bridge 的执行封装。

工具 RAG:TOOLS 是全量注册表(74 个);每轮请求只下发「常驻核心 + 按用户请求
检索命中的 top-k + 元工具(search_tools/call_tool)」,模型按需获取其余工具。

安全模型:
- DANGEROUS_TOOLS(删表/删行/删图表/清空/关簿/保护)不进 call_tool 白名单,
  只能作为直接工具调用,且执行前会向用户发起二次确认(agent 层负责)。
- 所有工具的参数先经 validate_args 按 schema 校验/纠正,再交给 bridge。
"""
import json
import logging
import os

from app.excel_bridge import BridgeError

log = logging.getLogger("excelai")

TOOLS = [
    {"type": "function", "function": {
        "name": "get_workbook_overview",
        "description": "获取工作簿概况:各工作表及大小、活动表、当前选中区域、表头预览。何时用:任何任务的第一个动作,先摸清结构再动手。",
        "parameters": {"type": "object", "properties": {}, "required": []}}},
    {"type": "function", "function": {
        "name": "get_sheet_preview",
        "description": "预览某工作表的表头与前几行样例。何时用:想快速了解一张陌生表的结构,不必读全量数据。默认 10 行 15 列。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string", "description": "工作表名,缺省为活动工作表"},
            "max_rows": {"type": "integer"}, "max_cols": {"type": "integer"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "read_range",
        "description": "读取指定区域的实际值。何时用:需要看具体单元格内容/数值时。单次不超过约 500 个单元格,大区域请分批或只读关键列。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string", "description": "工作表名,缺省为活动工作表"},
            "cell_range": {"type": "string", "description": "如 A1:F20 或 A:C"}}, "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "write_range",
        "description": "从 start_cell 开始写入二维数据(2D 数组)。字符串以 = 开头会被当作公式(用英文函数名和英文逗号,如 =SUM(A1:A10));纯数据不要带等号。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"},
            "start_cell": {"type": "string", "description": "起始单元格,如 G1"},
            "values": {"type": "array", "items": {"type": "array"},
                       "description": "二维数组,每行长度一致"}}, "required": ["start_cell", "values"]}}},
    {"type": "function", "function": {
        "name": "write_cell",
        "description": "写入单个单元格(值或公式)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell": {"type": "string"},
            "content": {"type": "string", "description": "内容;以 = 开头即为公式"}}, "required": ["cell", "content"]}}},
    {"type": "function", "function": {
        "name": "autofill_formula",
        "description": "把 source_cell 中的公式按相对引用自动填充到 target_range。整列同构公式务必优先用它,不要逐格写公式。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"},
            "source_cell": {"type": "string", "description": "已含公式的单元格,如 G2"},
            "target_range": {"type": "string", "description": "填充目标区域,如 G2:G200"}}, "required": ["source_cell", "target_range"]}}},
    {"type": "function", "function": {
        "name": "create_chart",
        "description": "根据数据区域插入图表。chart_type 取值:column/bar/line/pie/area/scatter/doughnut/radar。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "data_range": {"type": "string", "description": "含表头的数据区域"},
            "chart_type": {"type": "string"}, "title": {"type": "string"},
            "anchor": {"type": "string", "description": "图表左上角锚定单元格,默认 H2"}}, "required": ["data_range"]}}},
    {"type": "function", "function": {
        "name": "set_format",
        "description": "设置区域格式:加粗/斜体、字号、字体色与填充色(十六进制 RGB,如 FF0000)、数字格式(如 0.00%、#,##0、yyyy-mm-dd)、水平对齐。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "bold": {"type": "boolean"}, "italic": {"type": "boolean"}, "font_size": {"type": "number"},
            "font_color": {"type": "string"}, "fill_color": {"type": "string"},
            "number_format": {"type": "string"},
            "align": {"type": "string", "enum": ["left", "center", "right"]},
            "vertical": {"type": "string", "enum": ["top", "middle", "bottom"]},
            "wrap_text": {"type": "boolean"},
            "font_name": {"type": "string", "description": "字体名,如 微软雅黑 / Consolas"},
            "underline": {"type": "boolean", "description": "下划线"},
            "strikethrough": {"type": "boolean", "description": "删除线"},
            "indent": {"type": "integer", "description": "缩进级别 0–15"},
            "locked": {"type": "boolean", "description": "锁定单元格(工作表保护后生效)"},
            "hide_formula": {"type": "boolean",
                             "description": "隐藏公式(工作表保护后生效)"},
            "border": {"type": "string", "enum": ["thin", "medium", "thick"],
                       "description": "四边框粗细,省略=不动边框"}}, "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "add_sheet",
        "description": "在当前工作簿末尾新增工作表。",
        "parameters": {"type": "object", "properties": {"name": {"type": "string"}}, "required": ["name"]}}},
    {"type": "function", "function": {
        "name": "rename_sheet",
        "description": "重命名工作表。",
        "parameters": {"type": "object", "properties": {
            "old": {"type": "string"}, "new": {"type": "string"}}, "required": ["old", "new"]}}},
    {"type": "function", "function": {
        "name": "delete_sheet",
        "description": "删除工作表(不可恢复,谨慎使用)。",
        "parameters": {"type": "object", "properties": {"name": {"type": "string"}}, "required": ["name"]}}},
    {"type": "function", "function": {
        "name": "clear_range",
        "description": "清空区域。默认只清内容不清格式;contents_only=false 时连格式一起清。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "contents_only": {"type": "boolean"}}, "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "sort_range",
        "description": "对区域排序。column_index 是相对该区域的第几列(从 1 开始);"
                       "多列排序再给 column_index2/descending2(先按第一列,再按第二列)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "column_index": {"type": "integer"}, "descending": {"type": "boolean"},
            "has_header": {"type": "boolean"},
            "column_index2": {"type": "integer"}, "descending2": {"type": "boolean"}},
            "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "auto_fit_columns",
        "description": "自动调整列宽(缺省为整个使用区域)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_conditional_format",
        "description": "给区域加条件格式(高亮/预警)。rule 取值:color_scale 三色色阶、data_bar 数据条、"
                       "icon_set 图标集、cell_value 按单元格值(如大于某值标红)、duplicate 重复值高亮、"
                       "top_bottom 前/后 N 名。colors 用十六进制色值如 [\"FFEB84\"]。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "rule": {"type": "string", "enum": ["color_scale", "data_bar", "icon_set",
                                                "cell_value", "duplicate", "top_bottom"]},
            "colors": {"type": "array", "items": {"type": "string"}},
            "operator": {"type": "string",
                         "enum": ["greater", "less", "greater_equal", "less_equal",
                                  "equal", "not_equal", "between", "bottom"]},
            "value": {"type": "number"},
            "icon_set": {"type": "string"}},
            "required": ["cell_range", "rule"]}}},
    {"type": "function", "function": {
        "name": "set_validation",
        "description": "给区域加数据验证(下拉/输入约束)。kind 取值:list 下拉列表(用 values)、"
                       "whole 整数区间、decimal 小数区间、text_length 文本长度、date 日期区间。"
                       "区间类用 minimum/maximum 指定上下限。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "kind": {"type": "string", "enum": ["list", "whole", "decimal",
                                                "text_length", "date"]},
            "values": {"type": "array", "items": {"type": "string"}},
            "minimum": {"type": "number"}, "maximum": {"type": "number"}},
            "required": ["cell_range", "kind"]}}},
    {"type": "function", "function": {
        "name": "auto_filter",
        "description": "给区域加自动筛选。criteria 省略则只打开筛选下拉;"
                       "给了条件(如 \">100\"、\"华东\")则按第 field 列筛选;"
                       "action=off 取消该表的筛选。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "field": {"type": "integer"}, "criteria": {"type": "string"},
            "action": {"type": "string", "enum": ["on", "off"]}},
            "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "remove_duplicates",
        "description": "删除区域中的重复行。columns 可指定参与比对的列序号(如 [1,2]),"
                       "省略则按整行去重。会返回删掉了多少行。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "columns": {"type": "array", "items": {"type": "integer"}},
            "has_header": {"type": "boolean"}}, "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "create_table",
        "description": "把区域转成 Excel 表格对象(带筛选按钮与隔行底色的规范表),"
                       "可同时加汇总行(show_totals)。totals_column 传 \"last\" 表示对最后一列求和。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "name": {"type": "string"}, "style": {"type": "string"},
            "show_totals": {"type": "boolean"}, "totals_column": {"type": "string"}},
            "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "create_pivot",
        "description": "创建数据透视表做分组统计(这是做汇总分析的首选,比手写 SUMIF 更清晰)。"
                       "rows/columns/values 都填表头字段名;values 为空时自动取第一个数值列。"
                       "agg 取值:sum 求和、average 平均、count 计数、count_nums 数值计数、"
                       "max 最大值、min 最小值、product 乘积、stddev 标准差、var 方差。"
                       "不要把透视表和 create_chart 混用:要数字结论用透视表,要图用 create_chart。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "data_range": {"type": "string"},
            "rows": {"type": "array", "items": {"type": "string"}},
            "columns": {"type": "array", "items": {"type": "string"}},
            "values": {"type": "array", "items": {"type": "string"}},
            "agg": {"type": "string", "enum": ["sum", "average", "count", "count_nums",
                                               "max", "min", "product", "stddev", "var"]},
            "anchor": {"type": "string"}},
            "required": ["data_range"]}}},
    {"type": "function", "function": {
        "name": "delete_chart",
        "description": "删除图表。省略 name 时删除该工作表上的全部嵌入图表。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"},
            "name": {"type": "string", "description": "图表对象名,省略=全部"}},
            "required": []}}},
    {"type": "function", "function": {
        "name": "find_replace",
        "description": "查找替换文本。cell_range 省略则全表;whole_cell=true 整格精确匹配;"
                       "批量把某个值改成另一个值时优先用它,不要逐格写入。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"},
            "find": {"type": "string"}, "replace": {"type": "string"},
            "cell_range": {"type": "string"},
            "whole_cell": {"type": "boolean"}, "case_sensitive": {"type": "boolean"}},
            "required": ["find"]}}},
    {"type": "function", "function": {
        "name": "insert_rows",
        "description": "从第 row 行起插入 count 行(原数据下移)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "row": {"type": "integer"},
            "count": {"type": "integer"}}, "required": ["row"]}}},
    {"type": "function", "function": {
        "name": "delete_rows",
        "description": "删除从第 row 行起的 count 行(下方数据上移)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "row": {"type": "integer"},
            "count": {"type": "integer"}}, "required": ["row"]}}},
    {"type": "function", "function": {
        "name": "insert_columns",
        "description": "在第 column 列(1 基序号)起插入 count 列。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "column": {"type": "integer"},
            "count": {"type": "integer"}}, "required": ["column"]}}},
    {"type": "function", "function": {
        "name": "delete_columns",
        "description": "删除从第 column 列(1 基序号)起的 count 列。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "column": {"type": "integer"},
            "count": {"type": "integer"}}, "required": ["column"]}}},
    {"type": "function", "function": {
        "name": "merge_cells",
        "description": "合并单元格(action=merge)或取消合并(unmerge)。合并默认居中。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "action": {"type": "string", "enum": ["merge", "unmerge"]},
            "center": {"type": "boolean"}}, "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "set_column_width",
        "description": "设置列宽。column 可以是列字母(如 C)或 1 基序号。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"},
            "column": {"type": "string", "description": "列字母或序号"},
            "width": {"type": "number"}}, "required": ["column", "width"]}}},
    {"type": "function", "function": {
        "name": "set_row_height",
        "description": "设置行高(磅)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "row": {"type": "integer"},
            "height": {"type": "number"}}, "required": ["row", "height"]}}},
    {"type": "function", "function": {
        "name": "freeze_panes",
        "description": "冻结窗格。cell 是冻结点右下方第一格(冻结首行填 A2,首行首列填 B2);"
                       "省略或留空 = 取消冻结。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "protect_sheet",
        "description": "保护(protect)或取消保护(protect=false)工作表,可带密码。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "password": {"type": "string"},
            "protect": {"type": "boolean"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "copy_range",
        "description": "把源区域复制到目标起始格,可跨工作表。values_only=true 只粘贴值"
                       "(去掉公式与格式);transpose=true 转置粘贴;两者可叠加。",
        "parameters": {"type": "object", "properties": {
            "source_range": {"type": "string"}, "target_start": {"type": "string"},
            "sheet": {"type": "string"}, "target_sheet": {"type": "string"},
            "values_only": {"type": "boolean"}, "transpose": {"type": "boolean"}},
            "required": ["source_range", "target_start"]}}},
    {"type": "function", "function": {
        "name": "get_cell_formula",
        "description": "读取区域的公式文本(与 read_range 读值相对;用户问「这格公式是什么」时用)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"}},
            "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "delete_table",
        "description": "取消表格对象(ListObject 转回普通区域,数据不丢)。何时用:用户说「取消表格/转回普通区域」时;不要与删除工作表混淆。name 省略=该表全部。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "name": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_hyperlink",
        "description": "给单元格加超链接。text 省略则保留单元格原显示文本。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell": {"type": "string"},
            "url": {"type": "string"}, "text": {"type": "string"}},
            "required": ["cell", "url"]}}},
    {"type": "function", "function": {
        "name": "copy_sheet",
        "description": "复制一个工作表(含数据与格式)到工作簿末尾并命名。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string", "description": "源工作表名"},
            "name": {"type": "string", "description": "新工作表名"}},
            "required": ["sheet", "name"]}}},
    {"type": "function", "function": {
        "name": "set_visible",
        "description": "显示/隐藏行、列或工作表。kind=row(数字行号)/column(列字母或序号)/"
                       "sheet(工作表名);visible=false 隐藏、true 显示。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"},
            "kind": {"type": "string", "enum": ["row", "column", "sheet"]},
            "target": {"type": "string"}, "visible": {"type": "boolean"}},
            "required": ["kind", "target"]}}},
    {"type": "function", "function": {
        "name": "clear_conditional_format",
        "description": "清除区域上的全部条件格式(高亮/色阶/数据条等)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"}},
            "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "clear_formats",
        "description": "只清除区域的格式(边框/颜色/数字格式),保留内容。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"}},
            "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "close_workbook",
        "description": "关闭工作簿(默认不保存;save=true 时保存后关闭,从未保存过的会报错)。",
        "parameters": {"type": "object", "properties": {
            "name": {"type": "string", "description": "工作簿名,省略=当前操作的工作簿"},
            "save": {"type": "boolean"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "get_selection",
        "description": "读取用户当前选中区域的地址、尺寸与内容(用户问「我选中的是什么」时用)。",
        "parameters": {"type": "object", "properties": {}, "required": []}}},
    {"type": "function", "function": {
        "name": "open_workbook",
        "description": "打开本机已有的工作簿文件(路径为绝对路径)。",
        "parameters": {"type": "object", "properties": {
            "path": {"type": "string"}}, "required": ["path"]}}},
    {"type": "function", "function": {
        "name": "save_workbook_as",
        "description": "另存为。扩展名决定格式(.xlsx/.xlsm/.xls/.csv),也可显式传 file_format。",
        "parameters": {"type": "object", "properties": {
            "path": {"type": "string"}, "file_format": {"type": "integer"}},
            "required": ["path"]}}},
    {"type": "function", "function": {
        "name": "export_pdf",
        "description": "把工作表导出为 PDF 文件(绝对路径;缺 .pdf 后缀会自动补)。",
        "parameters": {"type": "object", "properties": {
            "path": {"type": "string"}, "sheet": {"type": "string"}},
            "required": ["path"]}}},
    {"type": "function", "function": {
        "name": "print_sheets",
        "description": "把工作表发送到默认打印机直接打印(注意纸张)。sheets 为工作表名列表,省略=当前表。",
        "parameters": {"type": "object", "properties": {
            "sheets": {"type": "array", "items": {"type": "string"}},
            "copies": {"type": "integer"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_doc_properties",
        "description": "设置工作簿文档属性(标题/作者/主题/关键词/备注)。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"}, "author": {"type": "string"},
            "subject": {"type": "string"}, "keywords": {"type": "string"},
            "comments": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "get_doc_properties",
        "description": "读取工作簿文档属性(标题/作者/主题/创建时间/最后保存时间等)。",
        "parameters": {"type": "object", "properties": {}, "required": []}}},
    {"type": "function", "function": {
        "name": "protect_workbook",
        "description": "保护/取消保护工作簿结构(防止增删移动工作表)。",
        "parameters": {"type": "object", "properties": {
            "password": {"type": "string"}, "protect": {"type": "boolean"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "define_name",
        "description": "定义名称。refers_to 用绝对引用,如 \"=Sheet1!$A$1:$A$10\" 或 \"=Sheet1!$B$2*2\"。",
        "parameters": {"type": "object", "properties": {
            "name": {"type": "string"}, "refers_to": {"type": "string"}},
            "required": ["name", "refers_to"]}}},
    {"type": "function", "function": {
        "name": "get_names",
        "description": "列出工作簿中全部定义名称及其引用。",
        "parameters": {"type": "object", "properties": {}, "required": []}}},
    {"type": "function", "function": {
        "name": "refresh_all",
        "description": "刷新工作簿的全部外部数据连接与透视表。",
        "parameters": {"type": "object", "properties": {}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_calc_mode",
        "description": "设置计算模式:auto 自动(默认)/ manual 手动 / semiauto 半自动。",
        "parameters": {"type": "object", "properties": {
            "mode": {"type": "string", "enum": ["auto", "manual", "semiauto"]}}, "required": []}}},
    {"type": "function", "function": {
        "name": "calculate",
        "description": "强制重算全部公式(手动计算模式下用)。",
        "parameters": {"type": "object", "properties": {}, "required": []}}},
    {"type": "function", "function": {
        "name": "move_sheet",
        "description": "移动工作表位置:给 after/before(工作表名)或 position(1 基序号)。"
                       "以返回的实际落位为准。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "after": {"type": "string"},
            "before": {"type": "string"}, "position": {"type": "integer"}},
            "required": ["sheet"]}}},
    {"type": "function", "function": {
        "name": "set_tab_color",
        "description": "设置工作表标签颜色(十六进制,如 FF8800);color 省略=清除。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "color": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_gridlines",
        "description": "显示/隐藏当前工作表的网格线(视图级)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "visible": {"type": "boolean"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_page_setup",
        "description": "页面设置:orientation=portrait/landscape、fit_to_width(宽度缩放页数)、"
                       "paper_size=A4/A3/B5/letter。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "orientation": {"type": "string"},
            "fit_to_width": {"type": "integer"}, "paper_size": {"type": "string"}},
            "required": []}}},
    {"type": "function", "function": {
        "name": "set_print_area",
        "description": "设置打印区域(如 A1:H30);cell_range 省略=清除。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_print_titles",
        "description": "设置打印标题:每页重复的行(如 \"1:1\")或列(如 \"A:A\")。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "rows": {"type": "string"},
            "columns": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_zoom",
        "description": "设置工作表视图缩放(10–400)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "percent": {"type": "integer"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "split_panes",
        "description": "拆分窗格:cell 为拆分点右下方第一格(如 B2);省略 = 取消拆分。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "set_comment",
        "description": "添加/修改单元格批注;text 为空字符串时清除批注。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell": {"type": "string"},
            "text": {"type": "string"}}, "required": ["cell", "text"]}}},
    {"type": "function", "function": {
        "name": "text_to_columns",
        "description": "把一列按分隔符拆成多列(分列)。delimiter:逗号/分号/空格/制表符/其他字符。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "delimiter": {"type": "string"}}, "required": ["cell_range"]}}},
    {"type": "function", "function": {
        "name": "copy_format",
        "description": "格式刷:把源区域首个单元格的核心格式(字体/填充/数字格式/边框)套到目标区域。",
        "parameters": {"type": "object", "properties": {
            "source_range": {"type": "string"}, "target_range": {"type": "string"},
            "sheet": {"type": "string"}, "target_sheet": {"type": "string"}},
            "required": ["source_range", "target_range"]}}},
    {"type": "function", "function": {
        "name": "set_display_formulas",
        "description": "在工作表视图里显示公式(而非计算结果)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "show": {"type": "boolean"}}, "required": []}}},
    {"type": "function", "function": {
        "name": "advanced_filter",
        "description": "高级筛选:按条件区域做多条件筛选。何时用:auto_filter 的单条件不够用时。"
                       "criteria_range 为含表头的条件区域;copy_to 给出则结果复制到该处,"
                       "否则原地筛选;unique_only=true 只留不重复记录。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "cell_range": {"type": "string"},
            "criteria_range": {"type": "string"}, "copy_to": {"type": "string"},
            "unique_only": {"type": "boolean"}},
            "required": ["cell_range", "criteria_range"]}}},
    {"type": "function", "function": {
        "name": "group_rows",
        "description": "把行分组(大纲/折叠);ungroup=true 取消分组。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "start_row": {"type": "integer"},
            "end_row": {"type": "integer"}, "ungroup": {"type": "boolean"}},
            "required": ["start_row"]}}},
    {"type": "function", "function": {
        "name": "consolidate",
        "description": "合并计算:把多个结构相同的区域按标签汇总成一个结果。何时用:多张表/多个区域要汇总到一张总表时(单表分组统计用 create_pivot)。sources 为「工作表!区域」列表,如 [\"Sheet1!A1:B10\"]。"
                       "sources 如 [\"Sheet1!A1:B10\", \"Sheet2!A1:B10\"]。",
        "parameters": {"type": "object", "properties": {
            "sources": {"type": "array", "items": {"type": "string"}},
            "target_start": {"type": "string"}, "function": {"type": "string",
            "enum": ["sum", "average", "count", "max", "min"]},
            "target_sheet": {"type": "string"},
            "top_row": {"type": "boolean"}, "left_column": {"type": "boolean"}},
            "required": ["sources", "target_start"]}}},
    {"type": "function", "function": {
        "name": "goal_seek",
        "description": "单变量求解:调整 changing_cell 的值,使 target_cell 的公式结果等于 target_value。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "target_cell": {"type": "string"},
            "target_value": {"type": "number"}, "changing_cell": {"type": "string"}},
            "required": ["target_cell", "target_value", "changing_cell"]}}},
    {"type": "function", "function": {
        "name": "config_chart",
        "description": "修改已有图表:chart_type(类型)、title(标题,空串=隐藏)、"
                       "legend(图例开关)、data_labels(数据标签开关)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "name": {"type": "string", "description": "图表对象名"},
            "title": {"type": "string"}, "chart_type": {"type": "string"},
            "legend": {"type": "boolean"}, "data_labels": {"type": "boolean"}},
            "required": ["name"]}}},
    {"type": "function", "function": {
        "name": "add_trendline",
        "description": "给图表第一个系列加趋势线。kind:linear 线性(默认)/logarithmic 对数/power 幂。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "name": {"type": "string"},
            "kind": {"type": "string", "enum": ["linear", "logarithmic", "power"]}},
            "required": ["name"]}}},
    {"type": "function", "function": {
        "name": "refresh_pivot",
        "description": "刷新指定名称的数据透视表(全部刷新用 refresh_all)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "name": {"type": "string"}}, "required": ["name"]}}},
    {"type": "function", "function": {
        "name": "delete_pivot",
        "description": "删除指定名称的数据透视表(按名删除,清空其区域)。",
        "parameters": {"type": "object", "properties": {
            "sheet": {"type": "string"}, "name": {"type": "string"}}, "required": ["name"]}}},
    {"type": "function", "function": {
        "name": "save_workbook",
        "description": "保存当前工作簿。",
        "parameters": {"type": "object", "properties": {}, "required": []}}},
]

# ---------------- 工具 RAG:按用户请求检索最相关的工具 ----------------

# 常驻核心:任何表格任务大概率都会用到,不依赖检索命中
CORE_TOOL_NAMES = ("get_workbook_overview", "read_range", "write_range", "save_workbook")

# 口语/同义词 → 补充检索词(工具描述里没出现的说法,靠这里桥接)
_SYNONYMS = {
    "透视": "数据透视表 汇总", "汇总": "透视 求和 合并计算", "做图": "图表",
    "饼图": "图表 饼图", "柱状": "图表", "去重": "删除重复行", "查重": "删除重复行",
    "冻结": "冻结窗格", "批注": "批注 备注", "分列": "分列 拆分",
    "高亮": "条件格式", "标红": "条件格式 颜色", "打印": "打印 页面设置",
    "pdf": "导出 pdf", "导出": "导出 另存为", "超链接": "超链接",
    "列宽": "列宽 调整列宽", "行高": "行高", "合并": "合并", "保护": "保护 锁定",
    "隐藏": "隐藏 显示", "筛选": "筛选", "排序": "排序", "保存": "保存",
    "另存": "另存为 导出", "宽度": "列宽", "删掉图": "删除图表", "换行": "自动换行",
    "边框": "边框", "大写": "文本 公式", "粘贴": "复制 粘贴",
}


def _bigrams(s):
    return {s[i:i + 2] for i in range(len(s) - 1)}


def retrieve_tools(query, top_k=5):
    """按用户请求检索最相关的工具 schema(降序)。

    纯 stdlib:字符二元组重叠 + 工具名命中加分 + 同义词桥接;无需分词库。
    query 为空或没有命中时返回 [](调用方仍保有常驻核心与元工具)。
    """
    q = str(query or "").strip()
    if not q:
        return []
    extra = " ".join(v for k, v in _SYNONYMS.items() if k in q)
    qb = _bigrams(q + " " + extra)
    qn = q.replace(" ", "")
    scored = []
    for t in TOOLS:
        fn = t["function"]
        text = fn["description"] + " " + fn["name"].replace("_", " ")
        score = len(qb & _bigrams(text))
        if qn and qn in fn["name"].replace("_", ""):
            score += 5
        for k, v in _SYNONYMS.items():
            if k in q and any(w in fn["description"] for w in v.split()):
                score += 2
        if score > 0:
            scored.append((score, fn["name"], t))
    scored.sort(key=lambda x: (-x[0], x[1]))
    hits = [t for _, _, t in scored[:top_k]]
    log.info("工具检索 query=%r → %s", q[:60],
             [t["function"]["name"] for t in hits] or "无命中")
    return hits


# 危险工具:不进 call_tool 白名单;直接调用时需用户二次确认(agent 层)
DANGEROUS_TOOLS = frozenset({
    "delete_sheet", "delete_rows", "delete_columns", "delete_chart",
    "delete_table", "delete_pivot", "clear_range", "close_workbook",
    "protect_sheet", "protect_workbook",
})

# 用户自定义危险工具规则(dangerous_tools.json,热加载):
# {"dangerous": ["工具名", ...]} —— 与内置清单合并生效,改文件即生效,无需重打包
_DANGEROUS_RULE_KEY = None        # (st_mtime_ns, st_size) 缓存键
_DANGEROUS_RULE_EXTRA = frozenset()


def _load_dangerous_rules():
    """热加载用户自定义危险工具规则(文件 mtime_ns/size 变化才重读)。"""
    global _DANGEROUS_RULE_KEY, _DANGEROUS_RULE_EXTRA
    try:
        from app.config import data_dir
        path = os.path.join(data_dir(), "dangerous_tools.json")
        st = os.stat(path)
        key = (st.st_mtime_ns, st.st_size)
    except Exception:
        if _DANGEROUS_RULE_KEY is not None:   # 文件被删:清掉旧规则
            _DANGEROUS_RULE_KEY = None
            _DANGEROUS_RULE_EXTRA = frozenset()
        return frozenset()
    if key == _DANGEROUS_RULE_KEY:
        return _DANGEROUS_RULE_EXTRA
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        extra = frozenset(str(x).strip() for x in (data.get("dangerous") or [])
                          if str(x).strip())
        log.info("已加载自定义危险工具规则:%s", sorted(extra) or "(空)")
    except Exception as e:
        log.warning("dangerous_tools.json 解析失败,忽略自定义规则:%s", e)
        extra = frozenset()
    _DANGEROUS_RULE_KEY = key
    _DANGEROUS_RULE_EXTRA = extra
    return extra


def dangerous_tools():
    """生效的危险工具全集 = 内置清单 ∪ 用户 JSON 规则(热加载)。"""
    return DANGEROUS_TOOLS | _load_dangerous_rules()


def _coerce(value, typ):
    """按 schema 类型纠正模型参数(模型常把数字传成字符串等)。"""
    try:
        if typ == "number":
            if isinstance(value, bool):
                return value, True
            if isinstance(value, (int, float)):
                return value, True
            return float(str(value).strip()), True
        if typ == "integer":
            if isinstance(value, bool):
                return value, True
            if isinstance(value, int):
                return value, True
            return int(float(str(value).strip())), True
        if typ == "boolean":
            if isinstance(value, bool):
                return value, True
            s = str(value).strip().lower()
            if s in ("true", "1", "yes", "是"):
                return True, True
            if s in ("false", "0", "no", "否"):
                return False, True
            return value, False
        if typ == "string":
            return str(value), True
    except (TypeError, ValueError):
        pass
    return value, False


def validate_args(params, args):
    """按工具的 parameters(JSON Schema 子集)校验并纠正 args。

    支持子集:type / required / properties / enum / items;未声明的键原样保留。
    返回 (纠正后的 args, 错误列表);错误非空时不应执行。
    """
    args = dict(args or {})
    errors = []
    props = (params or {}).get("properties") or {}
    for key, rule in props.items():
        if key not in args:
            continue
        typ = rule.get("type")
        if typ and typ not in ("object", "array"):
            coerced, ok = _coerce(args[key], typ)
            if ok:
                args[key] = coerced
            else:
                errors.append(f"{key} 应为 {typ}")
                continue
        if rule.get("enum") and args[key] not in rule["enum"]:
            errors.append(f"{key} 只能是 {rule['enum']} 之一")
    for key in (params or {}).get("required") or []:
        if key not in args or args[key] in (None, ""):
            errors.append(f"缺少必填参数 {key}")
    return args, errors


# 元工具:让模型按需搜索并调用注册表里的任意工具
META_TOOLS = [
    {"type": "function", "function": {
        "name": "search_tools",
        "description": "搜索可用的表格工具:返回与需求相关的工具名、说明和参数定义。"
                       "当你要做的操作没有现成工具可直接调用时,先用它查找。",
        "parameters": {"type": "object", "properties": {
            "query": {"type": "string", "description": "想做什么,如「设置列宽」「导入 csv」"}},
            "required": ["query"]}}},
    {"type": "function", "function": {
        "name": "call_tool",
        "description": "调用任意表格工具。name 填 search_tools 返回的工具名,"
                       "args 按该工具 parameters 传(没有就先 search_tools)。",
        "parameters": {"type": "object", "properties": {
            "name": {"type": "string"}, "args": {"type": "object"}},
            "required": ["name"]}}},
]

STEP_NAMES = {
    "get_workbook_overview": "获取工作簿概况",
    "get_sheet_preview": "预览工作表",
    "read_range": "读取区域",
    "write_range": "写入数据",
    "write_cell": "写入单元格",
    "autofill_formula": "填充公式",
    "create_chart": "插入图表",
    "set_format": "设置格式",
    "add_sheet": "新增工作表",
    "rename_sheet": "重命名工作表",
    "delete_sheet": "删除工作表",
    "clear_range": "清空区域",
    "sort_range": "排序",
    "auto_fit_columns": "调整列宽",
    "set_conditional_format": "设置条件格式",
    "set_validation": "设置数据验证",
    "auto_filter": "设置筛选",
    "remove_duplicates": "删除重复行",
    "create_table": "创建表格",
    "create_pivot": "创建数据透视表",
    "delete_chart": "删除图表",
    "find_replace": "查找替换",
    "insert_rows": "插入行",
    "delete_rows": "删除行",
    "insert_columns": "插入列",
    "delete_columns": "删除列",
    "merge_cells": "合并单元格",
    "set_column_width": "设置列宽",
    "set_row_height": "设置行高",
    "freeze_panes": "冻结窗格",
    "protect_sheet": "保护工作表",
    "copy_range": "复制区域",
    "get_cell_formula": "读取公式",
    "delete_table": "取消表格对象",
    "set_hyperlink": "设置超链接",
    "copy_sheet": "复制工作表",
    "set_visible": "显示/隐藏",
    "clear_conditional_format": "清除条件格式",
    "clear_formats": "清除格式",
    "close_workbook": "关闭工作簿",
    "get_selection": "读取当前选区",
    "open_workbook": "打开工作簿",
    "save_workbook_as": "另存为",
    "export_pdf": "导出 PDF",
    "print_sheets": "打印",
    "set_doc_properties": "设置文档属性",
    "get_doc_properties": "读取文档属性",
    "protect_workbook": "保护工作簿",
    "define_name": "定义名称",
    "get_names": "获取名称",
    "refresh_all": "全部刷新",
    "set_calc_mode": "设置计算模式",
    "calculate": "重算",
    "move_sheet": "移动工作表",
    "set_tab_color": "标签颜色",
    "set_gridlines": "网格线",
    "set_page_setup": "页面设置",
    "set_print_area": "打印区域",
    "set_print_titles": "打印标题",
    "set_zoom": "缩放",
    "split_panes": "拆分窗格",
    "set_comment": "批注",
    "text_to_columns": "分列",
    "copy_format": "格式刷",
    "set_display_formulas": "显示公式",
    "advanced_filter": "高级筛选",
    "group_rows": "行分组",
    "consolidate": "合并计算",
    "goal_seek": "单变量求解",
    "config_chart": "配置图表",
    "add_trendline": "加趋势线",
    "refresh_pivot": "刷新透视表",
    "delete_pivot": "删除透视表",
    "save_workbook": "保存工作簿",
}


class Executor:
    def __init__(self, bridge):
        self.bridge = bridge

    def describe(self, name, args):
        args = args or {}
        base = STEP_NAMES.get(name, name)
        sheet = args.get("sheet") or "活动表"
        if name == "read_range":
            return f"读取 {sheet}!{args.get('cell_range', '')}"
        if name == "get_sheet_preview":
            return f"预览 {sheet}"
        if name == "write_range":
            return f"向 {sheet}!{args.get('start_cell', '')} 写入数据"
        if name == "write_cell":
            return f"写入 {sheet}!{args.get('cell', '')}"
        if name == "autofill_formula":
            return f"填充公式 {args.get('source_cell', '')} → {args.get('target_range', '')}"
        if name == "create_chart":
            return f"插入图表({args.get('chart_type', 'column')}) 源 {args.get('data_range', '')}"
        if name == "set_format":
            return f"设置格式 {sheet}!{args.get('cell_range', '')}"
        if name in ("add_sheet", "rename_sheet", "delete_sheet"):
            return f"{base}:{args.get('name') or (args.get('old', '') + '→' + args.get('new', ''))}"
        if name == "clear_range":
            return f"清空 {sheet}!{args.get('cell_range', '')}"
        if name == "sort_range":
            return f"排序 {sheet}!{args.get('cell_range', '')}"
        if name == "set_conditional_format":
            kinds = {"color_scale": "色阶", "data_bar": "数据条", "icon_set": "图标集",
                     "cell_value": "按值高亮", "duplicate": "重复值", "top_bottom": "前/后N名"}
            r = args.get("rule", "")
            extra = ""
            if r == "cell_value":
                extra = f" {args.get('operator', '')} {args.get('value', '')}"
            elif r == "top_bottom":
                extra = f" {args.get('value', 10)} 名"
            return f"{kinds.get(r, r)}{extra} {sheet}!{args.get('cell_range', '')}"
        if name == "set_validation":
            kinds = {"list": "下拉列表", "whole": "整数区间", "decimal": "数值区间",
                     "text_length": "文本长度", "date": "日期区间"}
            k = args.get("kind", "")
            extra = f"({len(args.get('values') or [])} 项)" if k == "list" else \
                f"({args.get('minimum', '')}~{args.get('maximum', '')})"
            return f"{kinds.get(k, k)}{extra} {sheet}!{args.get('cell_range', '')}"
        if name == "auto_filter":
            c = args.get("criteria")
            return f"筛选 {sheet}!{args.get('cell_range', '')}" + (f" 第{args.get('field', 1)}列 {c}" if c else "")
        if name == "remove_duplicates":
            return f"去重 {sheet}!{args.get('cell_range', '')}"
        if name == "create_table":
            return f"转表格 {sheet}!{args.get('cell_range', '')}" + \
                   ("(带汇总行)" if args.get("show_totals") else "")
        if name == "create_pivot":
            dims = []
            if args.get("rows"):
                dims.append("行=" + "/".join(args["rows"]))
            if args.get("columns"):
                dims.append("列=" + "/".join(args["columns"]))
            if args.get("values"):
                dims.append("值=" + "/".join(args["values"]))
            return f"透视表 {args.get('agg', 'sum')} 源 {sheet}!{args.get('data_range', '')}" + \
                   (" [" + " ".join(dims) + "]" if dims else "")
        if name == "delete_chart":
            return f"删除图表 {args.get('name') or '(全部)'} @ {sheet}"
        if name == "find_replace":
            return f"查找替换「{args.get('find', '')}」→「{args.get('replace', '')}」" + \
                   (f" {sheet}!{args['cell_range']}" if args.get("cell_range") else " 全表")
        if name in ("insert_rows", "delete_rows"):
            act = "插入" if name.startswith("insert") else "删除"
            return f"{act} {args.get('count', 1)} 行 @ 第{args.get('row')}行 {sheet}"
        if name in ("insert_columns", "delete_columns"):
            act = "插入" if name.startswith("insert") else "删除"
            return f"{act} {args.get('count', 1)} 列 @ 第{args.get('column')}列 {sheet}"
        if name == "merge_cells":
            return ("取消合并 " if args.get("action") == "unmerge" else "合并 ") + \
                   f"{sheet}!{args.get('cell_range', '')}"
        if name == "set_column_width":
            return f"列宽 {args.get('column')}={args.get('width')} {sheet}"
        if name == "set_row_height":
            return f"行高 {args.get('row')}={args.get('height')} {sheet}"
        if name == "freeze_panes":
            return f"冻结窗格 {args.get('cell') or '(取消)'} @ {sheet}"
        if name == "protect_sheet":
            return ("保护 " if args.get("protect", True) else "取消保护 ") + sheet
        if name == "copy_range":
            return f"复制 {sheet}!{args.get('source_range', '')} → " \
                   f"{args.get('target_sheet') or sheet}!{args.get('target_start', '')}"
        if name == "get_cell_formula":
            return f"读取公式 {sheet}!{args.get('cell_range', '')}"
        if name == "delete_table":
            return f"取消表格对象 {args.get('name') or '(全部)'} @ {sheet}"
        if name == "set_hyperlink":
            return f"超链接 {sheet}!{args.get('cell', '')} → {args.get('url', '')}"
        if name == "copy_sheet":
            return f"复制工作表 {args.get('sheet', '')} → {args.get('name', '')}"
        if name == "set_visible":
            return ("显示 " if args.get("visible", True) else "隐藏 ") + \
                   f"{args.get('kind', '')} {args.get('target', '')}"
        if name == "clear_conditional_format":
            return f"清除条件格式 {sheet}!{args.get('cell_range', '')}"
        if name == "clear_formats":
            return f"清除格式 {sheet}!{args.get('cell_range', '')}"
        if name == "close_workbook":
            return f"关闭工作簿 {args.get('name') or '(当前)'}"
        if name == "get_selection":
            return "读取当前选区"
        return base

    def describe_ref(self, name, args):
        """返回这一步的数据来源区域(引用溯源),无来源时返回空串。

        对应 Copilot 的「上标编号 → 悬停看源」体验:这里把来源做成步骤卡片上的
        灰色标签,用户能立刻确认 AI 到底是读/改了哪块区域。
        """
        args = args or {}
        sheet = args.get("sheet")
        rng = None
        if name in ("read_range", "sort_range", "clear_range", "set_format", "autofill_formula",
               "set_conditional_format", "set_validation", "auto_filter",
               "remove_duplicates", "create_table", "merge_cells", "find_replace",
               "get_cell_formula", "clear_formats"):
            rng = args.get("cell_range") or args.get("target_range")
        elif name in ("write_cell",):
            rng = args.get("cell")
        elif name == "copy_range":
            rng = args.get("source_range")
        elif name == "write_range":
            rng = args.get("start_cell")
        elif name in ("create_chart", "create_pivot"):
            rng = args.get("data_range")
        if not rng:
            return ""
        return f"{sheet or '活动表'}!{rng}"

    def search_tools(self, query):
        """元工具:按需求搜索注册表,返回工具名/说明/参数定义(供 LLM 组装 call_tool)。"""
        hits = retrieve_tools(query, top_k=6)
        return {"tools": [
            {"name": t["function"]["name"],
             "description": t["function"]["description"],
             "parameters": t["function"]["parameters"]} for t in hits]}

    def execute(self, name, args):
        fn = getattr(self, "_t_" + name, None)
        if fn is None:
            return {"error": f"未知工具 {name}"}
        # schema 校验/纠正:模型常把数字传成字符串、漏掉枚举值,先纠正再执行
        schema = next((t["function"].get("parameters") for t in TOOLS
                       if t["function"]["name"] == name), None)
        args, errs = validate_args(schema, args)
        if errs:
            return {"error": "参数校验失败:" + ";".join(errs)}
        log.info("工具调用 %s args=%s", name, json.dumps(args, ensure_ascii=False)[:200])
        try:
            result = fn(**args)
        except BridgeError as e:
            result = {"error": str(e)}
        except TypeError as e:
            result = {"error": f"参数错误:{e}"}
        except Exception as e:
            result = {"error": f"执行失败:{e}"}
        log.info("工具结果 %s ok=%s", name, "error" not in result)
        return result

    # ---- 工具实现(转调 bridge) ----

    def _t_get_workbook_overview(self):
        return self.bridge.workbook_context()

    def _t_get_sheet_preview(self, sheet=None, max_rows=10, max_cols=15):
        return self.bridge.preview_sheet(sheet, int(max_rows or 10), int(max_cols or 15))

    def _t_read_range(self, sheet=None, cell_range=None, max_cells=500):
        if not cell_range:
            return {"error": "缺少 cell_range"}
        return self.bridge.read_range(sheet, cell_range, int(max_cells or 500))

    def _t_write_range(self, start_cell, values, sheet=None):
        return self.bridge.write_range(sheet, start_cell, values)

    def _t_write_cell(self, cell, content, sheet=None):
        return self.bridge.write_range(sheet, cell, [[content]])

    def _t_autofill_formula(self, source_cell, target_range, sheet=None):
        return self.bridge.autofill_formula(sheet, source_cell, target_range)

    def _t_create_chart(self, data_range, chart_type="column", title="", anchor="H2", sheet=None):
        return self.bridge.create_chart(sheet, data_range, chart_type, title, anchor)

    def _t_set_format(self, cell_range, sheet=None, **kw):
        return self.bridge.set_format(sheet, cell_range, **kw)

    def _t_add_sheet(self, name):
        return self.bridge.add_sheet(name)

    def _t_rename_sheet(self, old, new):
        return self.bridge.rename_sheet(old, new)

    def _t_delete_sheet(self, name):
        return self.bridge.delete_sheet(name)

    def _t_clear_range(self, cell_range, sheet=None, contents_only=True):
        return self.bridge.clear_range(sheet, cell_range, bool(contents_only))

    def _t_sort_range(self, cell_range, sheet=None, column_index=1, descending=False,
                      has_header=True, column_index2=None, descending2=False):
        return self.bridge.sort_range(sheet, cell_range, int(column_index or 1),
                                      bool(descending), bool(has_header),
                                      column_index2, bool(descending2 or False))

    def _t_auto_fit_columns(self, sheet=None, cell_range=None):
        return self.bridge.auto_fit_columns(sheet, cell_range)

    def _t_set_conditional_format(self, cell_range, rule, sheet=None, colors=None,
                                  operator="greater", value=None, icon_set="3_traffic_lights"):
        return self.bridge.set_conditional_format(sheet, cell_range, rule, colors,
                                                   operator, value, icon_set)

    def _t_set_validation(self, cell_range, kind, sheet=None, values=None,
                         minimum=None, maximum=None):
        return self.bridge.set_validation(sheet, cell_range, kind, values, minimum, maximum)

    def _t_auto_filter(self, cell_range, sheet=None, field=1, criteria=None, action="on"):
        return self.bridge.auto_filter(sheet, cell_range, int(field or 1), criteria,
                                       action=str(action))

    def _t_remove_duplicates(self, cell_range, sheet=None, columns=None, has_header=True):
        return self.bridge.remove_duplicates(sheet, cell_range, columns, bool(has_header))

    def _t_create_table(self, cell_range, sheet=None, name=None, style="TableStyleMedium2",
                        show_totals=False, totals_column=None):
        return self.bridge.create_table(sheet, cell_range, name, style,
                                        bool(show_totals), totals_column)

    def _t_create_pivot(self, data_range, sheet=None, rows=None, columns=None,
                        values=None, agg="sum", anchor=None):
        return self.bridge.create_pivot(sheet, data_range, rows, columns, values, agg, anchor)

    def _t_delete_chart(self, sheet=None, name=None):
        return self.bridge.delete_chart(sheet, name)

    def _t_find_replace(self, find, replace="", sheet=None, cell_range=None,
                        whole_cell=False, case_sensitive=False):
        return self.bridge.find_replace(sheet, find, replace, cell_range,
                                        bool(whole_cell), bool(case_sensitive))

    def _t_insert_rows(self, row, count=1, sheet=None):
        return self.bridge.insert_rows(sheet, int(row), int(count))

    def _t_delete_rows(self, row, count=1, sheet=None):
        return self.bridge.delete_rows(sheet, int(row), int(count))

    def _t_insert_columns(self, column, count=1, sheet=None):
        return self.bridge.insert_columns(sheet, int(column), int(count))

    def _t_delete_columns(self, column, count=1, sheet=None):
        return self.bridge.delete_columns(sheet, int(column), int(count))

    def _t_merge_cells(self, cell_range, sheet=None, action="merge", center=True):
        return self.bridge.merge_cells(sheet, cell_range, str(action), bool(center))

    def _t_set_column_width(self, column, width, sheet=None):
        return self.bridge.set_column_width(sheet, column, width)

    def _t_set_row_height(self, row, height, sheet=None):
        return self.bridge.set_row_height(sheet, row, height)

    def _t_freeze_panes(self, cell=None, sheet=None):
        return self.bridge.freeze_panes(sheet, cell)

    def _t_protect_sheet(self, sheet=None, password=None, protect=True):
        return self.bridge.protect_sheet(sheet, password, bool(protect))

    def _t_copy_range(self, source_range, target_start, sheet=None, target_sheet=None,
                      values_only=False, transpose=False):
        return self.bridge.copy_range(source_range, target_start, sheet, target_sheet,
                                      bool(values_only), bool(transpose))

    def _t_get_cell_formula(self, cell_range, sheet=None):
        return self.bridge.get_cell_formula(sheet, cell_range)

    def _t_delete_table(self, sheet=None, name=None):
        return self.bridge.delete_table(sheet, name)

    def _t_clear_formats(self, cell_range, sheet=None):
        return self.bridge.clear_formats(sheet, cell_range)

    def _t_close_workbook(self, name=None, save=False):
        return self.bridge.close_workbook(name, bool(save))

    def _t_get_selection(self):
        return self.bridge.get_selection()

    def _t_set_hyperlink(self, cell, url, sheet=None, text=None):
        return self.bridge.set_hyperlink(sheet, cell, url, text)

    def _t_copy_sheet(self, sheet, name):
        return self.bridge.copy_sheet(sheet, name)

    def _t_set_visible(self, kind, target, visible=True, sheet=None):
        return self.bridge.set_visible(sheet, kind, target, bool(visible))

    def _t_clear_conditional_format(self, cell_range, sheet=None):
        return self.bridge.clear_conditional_format(sheet, cell_range)

    def _t_open_workbook(self, path):
        return self.bridge.open_workbook(path)

    def _t_save_workbook_as(self, path, file_format=None):
        return self.bridge.save_workbook_as(path, file_format)

    def _t_export_pdf(self, path, sheet=None):
        return self.bridge.export_pdf(path, sheet)

    def _t_print_sheets(self, sheets=None, copies=1):
        return self.bridge.print_sheets(sheets, copies)

    def _t_set_doc_properties(self, title=None, author=None, subject=None,
                              keywords=None, comments=None):
        return self.bridge.set_doc_properties(title, author, subject, keywords, comments)

    def _t_get_doc_properties(self):
        return self.bridge.get_doc_properties()

    def _t_protect_workbook(self, password=None, protect=True):
        return self.bridge.protect_workbook(password, bool(protect))

    def _t_define_name(self, name, refers_to):
        return self.bridge.define_name(name, refers_to)

    def _t_get_names(self):
        return self.bridge.get_names()

    def _t_refresh_all(self):
        return self.bridge.refresh_all()

    def _t_set_calc_mode(self, mode="auto"):
        return self.bridge.set_calc_mode(mode)

    def _t_calculate(self):
        return self.bridge.calculate()

    def _t_move_sheet(self, sheet, after=None, before=None, position=None):
        return self.bridge.move_sheet(sheet, after, before, position)

    def _t_set_tab_color(self, sheet=None, color=None):
        return self.bridge.set_tab_color(sheet, color)

    def _t_set_gridlines(self, visible=True, sheet=None):
        return self.bridge.set_gridlines(sheet, bool(visible))

    def _t_set_page_setup(self, sheet=None, orientation=None, fit_to_width=None,
                          paper_size=None):
        return self.bridge.set_page_setup(sheet, orientation, fit_to_width, paper_size)

    def _t_set_print_area(self, cell_range=None, sheet=None):
        return self.bridge.set_print_area(sheet, cell_range)

    def _t_set_print_titles(self, sheet=None, rows=None, columns=None):
        return self.bridge.set_print_titles(sheet, rows, columns)

    def _t_set_zoom(self, percent=100, sheet=None):
        return self.bridge.set_zoom(sheet, percent)

    def _t_split_panes(self, sheet=None, cell=None):
        return self.bridge.split_panes(sheet, cell)

    def _t_set_comment(self, cell, text="", sheet=None):
        return self.bridge.set_comment(sheet, cell, text)

    def _t_text_to_columns(self, cell_range, sheet=None, delimiter=","):
        return self.bridge.text_to_columns(sheet, cell_range, delimiter)

    def _t_copy_format(self, source_range, target_range, sheet=None, target_sheet=None):
        return self.bridge.copy_format(source_range, target_range, sheet, target_sheet)

    def _t_set_display_formulas(self, show=True, sheet=None):
        return self.bridge.set_display_formulas(sheet, bool(show))

    def _t_advanced_filter(self, cell_range, criteria_range, sheet=None, copy_to=None,
                           unique_only=False):
        return self.bridge.advanced_filter(sheet, cell_range, criteria_range,
                                           copy_to, bool(unique_only))

    def _t_group_rows(self, start_row, end_row=None, ungroup=False, sheet=None):
        return self.bridge.group_rows(sheet, start_row, end_row, bool(ungroup))

    def _t_consolidate(self, sources, target_start, function="sum", target_sheet=None,
                       top_row=True, left_column=True):
        return self.bridge.consolidate(sources, target_start, function,
                                       target_sheet, bool(top_row), bool(left_column))

    def _t_goal_seek(self, target_cell, target_value, changing_cell, sheet=None):
        return self.bridge.goal_seek(sheet, target_cell, target_value, changing_cell)

    def _t_config_chart(self, name, sheet=None, title=None, chart_type=None,
                        legend=None, data_labels=None):
        return self.bridge.config_chart(sheet, name, title, chart_type,
                                        legend, data_labels)

    def _t_add_trendline(self, name, sheet=None, kind="linear"):
        return self.bridge.add_trendline(sheet, name, kind)

    def _t_refresh_pivot(self, name, sheet=None):
        return self.bridge.refresh_pivot(sheet, name)

    def _t_delete_pivot(self, name, sheet=None):
        return self.bridge.delete_pivot(sheet, name)

    def _t_save_workbook(self):
        return self.bridge.save_workbook()


def brief_result(result):
    """给 UI 的单步结果摘要。

    不要在这里加 ✓/✗ 前缀 —— 步骤卡片左侧已经有状态图标了,再带符号会重复。
    """
    if not isinstance(result, dict):
        return "完成"
    if "error" in result:
        return str(result["error"])[:120]
    if "written" in result:
        return f"已写入 {result['written']} 格 → {result.get('range', '')}"
    if "filled" in result:
        return f"已填充 {result['filled']}"
    # 注意顺序:create_pivot / create_table 的返回值里也带 values、rows 等键,
    # 必须排在下面的 "values" 分支之前,否则会被误判成"读取返回 N 行数据"。
    if "pivot" in result:
        agg = {"sum": "求和", "average": "平均", "count": "计数",
               "count_nums": "数值计数", "max": "最大值", "min": "最小值"}.get(
                   result.get("agg", ""), result.get("agg", ""))
        return f"透视表已建,结果在 {result.get('result_range', '')}({agg})"
    if "table" in result:
        return f"表格「{result['table']}」{result.get('range', '')}" + \
               ("(含汇总行)" if result.get("totals") else "")
    if "deduped" in result:
        return f"删除重复 {result.get('removed_rows', 0)} 行,剩 {result.get('remaining', '?')} 行"
    if "filtered" in result:
        return f"{result.get('kind', '已筛选')} → {result['filtered']}"
    if "validated" in result:
        return f"已加{result.get('kind', '')}验证 → {result['validated']}"
    if "formatted" in result:
        return f"已应用{result.get('kind', '格式')} → {result['formatted']}"
    if "deleted_charts" in result:
        return f"已删除 {result['deleted_charts']} 个图表"
    if "replaced" in result:
        return f"替换 {result['replaced']} 处 → {result.get('range', '')}"
    if "inserted_rows" in result:
        return f"已插入 {result['inserted_rows']} 行(第 {result.get('at')} 行起)"
    if "deleted_rows" in result:
        return f"已删除 {result['deleted_rows']} 行(第 {result.get('at')} 行起)"
    if "inserted_cols" in result:
        return f"已插入 {result['inserted_cols']} 列(第 {result.get('at')} 列起)"
    if "deleted_cols" in result:
        return f"已删除 {result['deleted_cols']} 列(第 {result.get('at')} 列起)"
    if "merged" in result:
        return ("已合并 " if result["merged"] else "已取消合并 ") + str(result.get("range", ""))
    if "width_set" in result:
        return "已设" + str(result["width_set"])
    if "height_set" in result:
        return "已设" + str(result["height_set"])
    if "frozen" in result:
        return "冻结于 " + str(result["frozen"])
    if "protected" in result:
        return ("已保护 " if result["protected"] else "已取消保护 ") + str(result.get("sheet", ""))
    if "copied" in result:
        return "已复制 " + str(result["copied"])
    if "formulas" in result:
        return f"返回 {len(result['formulas'])} 行公式"
    if "unlisted" in result:
        return f"已取消 {result['unlisted']} 个表格对象"
    if "hyperlinked" in result:
        return "已加超链接 " + str(result["hyperlinked"])
    if "copied_sheet" in result:
        return f"已复制工作表 {result.get('from', '')} → {result['copied_sheet']}"
    if "visible" in result:
        return ("已显示 " if result["visible"] else "已隐藏 ") + str(result.get("target", ""))
    if "cleared_cf" in result:
        return "已清除条件格式 " + str(result["cleared_cf"])
    if "cleared_formats" in result:
        return "已清除格式 " + str(result["cleared_formats"])
    if "opened" in result:
        return "已打开 " + str(result["opened"])
    if "saved_as" in result:
        return "已另存为 " + str(result.get("path", ""))
    if "exported" in result:
        return "已导出 PDF " + str(result["exported"])
    if "printed" in result:
        return "已发送打印:" + ",".join(str(x) for x in result["printed"])
    if "properties_set" in result:
        return "已设置文档属性:" + ",".join(str(x) for x in result["properties_set"])
    if "names" in result:
        return f"共 {len(result['names'])} 个定义名称"
    if "refreshed" in result:
        return "已刷新"
    if "calc_mode" in result:
        return "计算模式=" + str(result["calc_mode"])
    if "calculated" in result:
        return "已重算"
    if "moved" in result:
        return f"已移动 {result['moved']} → 第 {result.get('position')} 位"
    if "tab_color" in result:
        return "标签色=" + str(result["tab_color"])
    if "gridlines" in result:
        return "网格线" + ("开" if result["gridlines"] else "关")
    if "page_setup" in result:
        ps = result["page_setup"]
        return "页面设置:" + (",".join(str(x) for x in ps) if isinstance(ps, list) else str(ps))
    if "print_area" in result:
        return "打印区域=" + str(result["print_area"])
    if "print_titles" in result:
        return "打印标题已设置"
    if "zoom" in result:
        return "缩放=" + str(result["zoom"]) + "%"
    if "split" in result:
        return "拆分=" + str(result["split"])
    if "comment" in result:
        return "批注:" + str(result["comment"])[:40]
    if "copied_format" in result:
        return "格式刷 " + str(result["copied_format"])
    if "display_formulas" in result:
        return "显示公式" + ("开" if result["display_formulas"] else "关")
    if "filtered_to" in result:
        return "筛选结果 → " + str(result["filtered_to"])
    if "filtered_in_place" in result:
        return "已原地筛选"
    if "grouped" in result:
        return ("已分组 " if result["grouped"] else "已取消分组 ") + str(result.get("rows", ""))
    if "consolidated" in result:
        return "合并计算(" + str(result.get("function", "")) + ")→ " + str(result["consolidated"])
    if "goal_seek" in result:
        return "单变量求解" + ("成功" if result["goal_seek"] else "失败") + \
               ",目标值=" + str(result.get("value", ""))
    if "configured" in result:
        ch = result.get("changes") or []
        return "图表已更新:" + (",".join(str(x) for x in ch) or str(result["configured"]))
    if "trendline" in result:
        return "已加趋势线(" + str(result.get("kind", "")) + ")"
    if "deleted_pivot" in result:
        return "已删除透视表 " + str(result["deleted_pivot"])
    if "closed" in result:
        return ("已保存并关闭 " if result.get("saved") else "已关闭 ") + str(result["closed"])
    if "note" in result and "selection" in result:
        return "已读取当前选区"
    if "filter_off" in result:
        return "已取消筛选"
    if "values" in result:
        return f"返回 {len(result['values'])} 行数据"
    if "chart" in result:
        return f"图表已插入({result.get('data_range', '')})"
    if "deleted" in result:
        return f"已删除 {result['deleted']}"
    if "saved" in result:
        return "已保存" if result["saved"] else "未保存:" + str(result.get("note", ""))
    return "完成"
