/* 表答 SheetTalk - WPS 加载项接口函数(由 WPS 自动生成的 index.html 引入)。
 *
 * ribbon.xml 里按钮的 onAction 指向本文件的 toggleExcelAiPane():
 * 打开/关闭停靠在表格右侧的任务窗格,页面复用桌面端同一份 addin.html
 * (由内置服务 127.0.0.1:8765 提供),因此对话历史与 COM 桥接完全共享。
 */

/* 任务窗格地址:host=wps 让 addin.html 跳过加载 Office JS(WPS 里没有该运行时) */
var EXCEL_AI_PANE_URL = "http://127.0.0.1:8765/addin.html?host=wps";
/* PluginStorage 里记录当前任务窗格 ID 的键名 */
var EXCEL_AI_PANE_KEY = "excelAiPaneId";

function toggleExcelAiPane() {
  try {
    // 先看上次的窗格是否还在:在就关闭(开关式)
    var lastId = wps.PluginStorage.getItem(EXCEL_AI_PANE_KEY);
    if (lastId) {
      var existing = null;
      try { existing = wps.GetTaskPane(lastId); } catch (e0) { existing = null; }
      if (existing) {
        existing.Delete();
        wps.PluginStorage.setItem(EXCEL_AI_PANE_KEY, "");
        return true;
      }
      wps.PluginStorage.setItem(EXCEL_AI_PANE_KEY, "");
    }

    // 新建任务窗格(默认停靠右侧,只调 Width,避免依赖不同版本的枚举值)
    var pane = wps.CreateTaskpane(EXCEL_AI_PANE_URL, "表答 SheetTalk");
    wps.PluginStorage.setItem(EXCEL_AI_PANE_KEY, pane.ID);
    try { pane.Width = 430; } catch (e1) { /* 个别版本只读宽度,忽略 */ }
    pane.Visible = true;
  } catch (e) {
    // 面向用户的可读提示:多半是桌面程序没开
    try {
      wps.Alert("无法打开侧边栏:" + e + "\n\n请确认桌面程序「表答 SheetTalk」正在本机运行"
        + "(它提供 127.0.0.1:8765 页面服务)。");
    } catch (e2) { /* Alert 也不可用就只能放弃 */ }
  }
  return true;
}
