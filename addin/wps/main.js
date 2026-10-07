/* 表答 SheetTalk - WPS 加载项入口垫片。
 *
 * WPS 启动时会在这个文件夹里自动生成 index.html 并引入 main.js:
 *   - 若它引入的是根目录 main.js(官方文档说法)→ 由本文件转手加载 js/main.js;
 *   - 若它引入的是 js/main.js(wpsjs 模板/社区惯例)→ 直接生效,本文件不会被用到。
 * 两条路径最终都只执行一次 js/main.js,请把所有接口函数写在 js/main.js 里。
 * 注意:不要在本目录创建 index.html(WPS 会自动生成)。
 */
document.write("<script src='js/main.js'><\/script>");
