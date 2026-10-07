import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Window

ApplicationWindow {
    id: win
    width: 470
    height: 920
    minimumWidth: 380
    minimumHeight: 600
    visible: true
    title: qsTr("表答 SheetTalk")
    // 窗口底色用 Fluent 的分层表面基色
    color: win.baseColor

    // 全局 palette 兜底:未显式设色的文字/控件继承这里,深浅色切换时全部跟着 win.* 令牌走,
    // 否则深色模式下会出现「黑字黑底」的隐形文字(空状态标题/步骤行就踩过)。
    palette.windowText: win.textColor
    palette.text: win.textColor
    palette.window: win.baseColor
    palette.base: win.layerColor
    palette.alternateBase: win.layerAlt
    palette.button: win.layerColor
    palette.buttonText: win.textColor
    palette.placeholderText: win.subText
    palette.highlight: win.accent
    palette.highlightedText: "#ffffff"
    palette.toolTipBase: win.layerColor
    palette.toolTipText: win.textColor

    // ---------------- Win11 Fluent 设计令牌 ----------------
    // 不依赖 Material 附加属性,控件外观全部由本文件的 background/contentItem 显式定义,
    // 好处是深浅色、圆角、描边、悬浮态完全可控,且不新增任何第三方依赖。
    // 主题跟随系统:themeMode=auto(默认)时读 Python 侧轮询到的系统亮暗(Windows 注册表,3s 一次)
    readonly property bool dark: ctrl.themeMode === "dark"
        || (ctrl.themeMode === "auto" && ctrl.systemDark)

    // 强调色:Excel 绿(Microsoft 365 Excel 的 #107C41),与表格程序的功能区观感一致
    readonly property color accent: "#107c41"
    readonly property color accentHover: Qt.lighter(accent, 1.12)
    readonly property color accentPressed: Qt.darker(accent, 1.15)
    readonly property color accentSoft: dark ? Qt.rgba(0.25, 0.62, 0.40, 0.30) : Qt.rgba(0.06, 0.49, 0.25, 0.12)
    // 更淡的强调色底(空态发送按钮等「可点但未激活」的控件)
    readonly property color accentFaint: dark ? Qt.rgba(0.25, 0.62, 0.40, 0.16) : Qt.rgba(0.06, 0.49, 0.25, 0.10)

    // 分层表面:base(窗口)→ layer(卡片)→ alt(次级控件)
    readonly property color baseColor: dark ? "#202020" : "#f3f3f3"
    readonly property color layerColor: dark ? "#2c2c2c" : "#ffffff"
    readonly property color layerAlt: dark ? "#333333" : "#fafafa"
    readonly property color stroke: dark ? "#484848" : "#e5e5e5"
    readonly property color strokeStrong: dark ? "#6a6a6a" : "#d0d0d0"
    readonly property color textColor: dark ? "#ffffff" : "#1b1b1b"
    readonly property color subText: dark ? "#bdbdbd" : "#5d5d5d"

    // 兼容旧属性名(下方大量引用 win.surface / win.outline / win.okColor …)
    readonly property color surface: layerColor
    readonly property color surfaceDim: layerAlt
    readonly property color outline: stroke
    readonly property color okColor: "#0f7b0f"
    readonly property color badColor: "#c42b1c"
    readonly property color warnColor: "#9d5d00"
    readonly property color grad1: accent
    readonly property color grad2: Qt.darker(accent, 1.18)

    // 圆角与高度按 Fluent 规格:控件 4px / 卡片 8px / 悬浮层 8px
    readonly property int rControl: 4
    readonly property int rCard: 8
    readonly property int hControl: 32
    readonly property int hButton: 32

    // 全局字体在 Python 侧通过 QFont.setFamilies() 设置(见 app/qt_app.py 的 _pick_font_family),
    // 因为 QML 的 font.family 只接受单个字体名,写成 "A, B, C" 会被当成一个不存在的字体名,
    // 从而回退到默认衬线体(中文界面尤其明显)。此处只保留首选项供个别控件覆盖。
    readonly property string fontFamily: "Segoe UI Variable Text"

    property int curAi: -1
    property var sidIndex: ({})

    // ---------------- Fluent 控件样式(集中定义,供全文件复用) ----------------

    // ---------------- 图标(Windows 自带 Segoe Fluent Icons 字体,无新增依赖) ----------------
    // QML 的 Text 无法直接显示 PUA 码位,统一用这些转义常量;页面里凡是要放图标的地方
    // 都必须把 font.family 设为 "Segoe Fluent Icons",否则会显示成豆腐块。
    readonly property string iconFont: "Segoe Fluent Icons"
    readonly property string icoClear: "\uE74D"     // 删除/清空
    readonly property string icoSettings: "\uE713"  // 齿轮
    readonly property string icoChart: "\uE9D2"     // 柱状图/分析
    readonly property string icoPie: "\uE8B9"       // 饼图/图表
    readonly property string icoCalc: "\uE8EF"      // 计算器/公式
    readonly property string icoBulb: "\uEA80"      // 灯泡/建议
    readonly property string icoSend: "\uE72A"      // 发送
    readonly property string icoSheet: "\uE7B3"     // 表格
    readonly property string icoCheck: "\uE73E"     // 对勾
    readonly property string icoError: "\uEA39"     // 错误
    readonly property string icoClock: "\uE823"     // 进行中/等待
    readonly property string icoWarn: "\uE7BA"      // 警告
    readonly property string icoClean: "\uE71C"     // 筛选/清洗
    readonly property string icoList: "\ue8fd"
    readonly property string icoStop: "\ue71a"
    readonly property string icoClose: "\uE711"    // 关闭(对话框右上角 ×)
    readonly property string icoAttach: "\uE724"   // 附件别针

    // 图标文本:统一字号与字体,避免各处重复写 font
    component FluIcon: Text {
        font.family: win.iconFont
        font.pixelSize: 14
        color: win.textColor
        horizontalAlignment: Text.AlignHCenter
        verticalAlignment: Text.AlignVCenter
        renderType: Text.NativeRendering
    }

    // 图标按钮:图标专用(与 FluIconButton 的区别是内容是图标而非文本)
    component FluIconAction: AbstractButton {
        id: ctl
        property string glyph: ""
        property int glyphSize: 14
        implicitWidth: 30
        implicitHeight: 30
        hoverEnabled: true
        // Fluent 触觉确认:按下轻微压缩;键盘焦点给描边(禁止忽略焦点状态)
        scale: ctl.down ? 0.92 : 1.0
        Behavior on scale { NumberAnimation { duration: 150; easing.type: Easing.OutCubic } }
        contentItem: FluIcon {
            text: ctl.glyph
            font.pixelSize: ctl.glyphSize
            color: win.textColor
        }
        background: Rectangle {
            radius: win.rControl
            color: ctl.down ? Qt.rgba(0, 0, 0, 0.08)
                 : (ctl.hovered || ctl.activeFocus) ? (win.dark ? Qt.rgba(1, 1, 1, 0.07) : Qt.rgba(0, 0, 0, 0.05))
                 : "transparent"
            border.width: ctl.activeFocus ? 1 : 0
            border.color: win.accent
            Behavior on color { ColorAnimation { duration: 120 } }
        }
        // 手型光标:图标动作都是可点击的
        MouseArea { anchors.fill: parent; acceptedButtons: Qt.NoButton; cursorShape: Qt.PointingHandCursor }
    }

    // 次要按钮:透明底 + 描边,悬浮时轻微着色
    component FluSecondaryButton: Button {
        id: ctl
        implicitHeight: win.hButton
        implicitWidth: Math.max(88, contentItem.implicitWidth + 28)
        font.pixelSize: 13
        hoverEnabled: true
        // Fluent 触觉确认:按下压缩到 0.97
        scale: ctl.down ? 0.97 : 1.0
        Behavior on scale { NumberAnimation { duration: 150; easing.type: Easing.OutCubic } }
        contentItem: Label {
            text: ctl.text
            font.pixelSize: 13
            color: ctl.enabled ? win.textColor : win.subText
            horizontalAlignment: Text.AlignHCenter
            verticalAlignment: Text.AlignVCenter
            elide: Text.ElideRight
        }
        background: Rectangle {
            radius: win.rControl
            color: ctl.down ? Qt.rgba(0.06, 0.49, 0.25, 0.20)
                 : ctl.hovered ? (win.dark ? Qt.rgba(1, 1, 1, 0.06) : Qt.rgba(0, 0, 0, 0.04))
                 : "transparent"
            border.width: !ctl.enabled ? 0 : (ctl.activeFocus ? 2 : 1)
            border.color: !ctl.enabled ? "transparent"
                         : ctl.activeFocus ? win.accent
                         : (ctl.hovered ? win.strokeStrong : win.stroke)
            Behavior on color { ColorAnimation { duration: 120 } }
        }
        // 手型光标:次要按钮也是动作入口
        MouseArea { anchors.fill: parent; acceptedButtons: Qt.NoButton; cursorShape: Qt.PointingHandCursor }
    }

    // 强调按钮:实心强调色 + 悬浮/按下加深
    component FluFilledButton: Button {
        id: ctl
        property color bgColor: win.accent
        property color bgHovered: win.accentHover
        property color bgDown: win.accentPressed
        implicitHeight: win.hButton
        implicitWidth: Math.max(88, contentItem.implicitWidth + 28)
        font.pixelSize: 13
        hoverEnabled: true
        // Fluent 触觉确认:同次要按钮(0.97 按压)
        scale: ctl.down ? 0.97 : 1.0
        Behavior on scale { NumberAnimation { duration: 150; easing.type: Easing.OutCubic } }
        contentItem: Label {
            text: ctl.text
            font.pixelSize: 13
            color: "#ffffff"
            horizontalAlignment: Text.AlignHCenter
            verticalAlignment: Text.AlignVCenter
            elide: Text.ElideRight
        }
        background: Rectangle {
            radius: win.rControl
            color: !ctl.enabled ? Qt.rgba(0.06, 0.49, 0.25, 0.30)
                 : ctl.down ? ctl.bgDown
                 : ctl.hovered ? ctl.bgHovered
                 : ctl.bgColor
            // 键盘焦点:WinUI 风格的 2px 内描边(实心底上用白环更醒目)
            border.width: ctl.activeFocus ? 2 : 0
            border.color: "#ffffff"
            Behavior on color { ColorAnimation { duration: 120 } }
        }
        // 手型光标:强调按钮同样可点
        MouseArea { anchors.fill: parent; acceptedButtons: Qt.NoButton; cursorShape: Qt.PointingHandCursor }
    }

    // 数字微调框:Fluent 的纤细 ± 按钮(默认 Basic 风格的深灰实心块太重)
    component FluSpinBox: SpinBox {
        id: ctl
        implicitHeight: win.hControl
        hoverEnabled: true
        editable: true
        // Material/Basic 风格的 SpinBox 输入后不会立即提交,失焦时才生效,
        // 不显式 commit 会造成「调了但没生效」的错觉。回车提交(SpinBox 无 onEditingFinished)。
        Keys.onReturnPressed: ctl.commit()
        onValueModified: ctl.commit()
        contentItem: TextInput {
            text: ctl.textFromValue(ctl.value, ctl.locale)
            font.pixelSize: 13
            color: win.textColor
            selectionColor: win.accent
            selectedTextColor: "#ffffff"
            horizontalAlignment: Text.AlignHCenter
            verticalAlignment: Text.AlignVCenter
            readOnly: !ctl.editable
            validator: ctl.validator
            inputMethodHints: Qt.ImhDigitsOnly
            onTextEdited: {
                var v = parseInt(text, 10)
                if (!isNaN(v)) ctl.value = v
            }
        }
        background: Item {
            Rectangle {
                anchors.fill: parent
                radius: win.rControl
                color: win.dark ? Qt.rgba(1, 1, 1, 0.04) : Qt.rgba(0, 0, 0, 0.02)
            }
            // 底部指示线,与输入框同一套语言
            Rectangle {
                anchors { left: parent.left; right: parent.right; bottom: parent.bottom }
                height: 1
                color: ctl.activeFocus ? win.accent : (ctl.hovered ? win.strokeStrong : win.stroke)
                Behavior on color { ColorAnimation { duration: 120 } }
            }
        }
        // 上/下箭头:细线三角,悬浮才显强调色
        up.indicator: Rectangle {
            x: ctl.width - width - 6
            y: ctl.height / 2 - height / 2 - 6
            implicitWidth: 20
            implicitHeight: 20
            radius: 3
            color: ctl.up.pressed ? Qt.rgba(0, 0, 0, 0.12) : (ctl.up.hovered ? win.accentSoft : "transparent")
            Text {
                anchors.centerIn: parent
                text: "▲"
                font.pixelSize: 7
                color: ctl.up.hovered ? win.accent : win.subText
            }
        }
        down.indicator: Rectangle {
            x: ctl.width - width - 6
            y: ctl.height / 2 - height / 2 + 6
            implicitWidth: 20
            implicitHeight: 20
            radius: 3
            color: ctl.down.pressed ? Qt.rgba(0, 0, 0, 0.12) : (ctl.down.hovered ? win.accentSoft : "transparent")
            Text {
                anchors.centerIn: parent
                text: "▼"
                font.pixelSize: 7
                color: ctl.down.hovered ? win.accent : win.subText
            }
        }
    }

    // 滚动条:Fluent 的细条设计,默认隐藏、悬浮才展开
    component FluScrollBar: ScrollBar {
        id: ctl
        policy: ScrollBar.AsNeeded
        implicitWidth: 10
        contentItem: Rectangle {
            implicitWidth: 6
            radius: 3
            color: ctl.pressed ? (win.dark ? Qt.rgba(1, 1, 1, 0.50) : Qt.rgba(0, 0, 0, 0.45))
                 : ctl.hovered ? (win.dark ? Qt.rgba(1, 1, 1, 0.36) : Qt.rgba(0, 0, 0, 0.32))
                 : (win.dark ? Qt.rgba(1, 1, 1, 0.22) : Qt.rgba(0, 0, 0, 0.16))
            Behavior on color { ColorAnimation { duration: 130 } }
            opacity: ctl.policy === ScrollBar.AlwaysOn || ctl.active || ctl.hovered ? 1.0 : 0.0
            Behavior on opacity { NumberAnimation { duration: 160 } }
        }
        background: Rectangle { color: "transparent" }
    }

    // 图标按钮:Fluent 的透明底 + 悬浮浅色底
    component FluIconButton: AbstractButton {
        id: ctl
        implicitWidth: 30
        implicitHeight: 30
        hoverEnabled: true
        // Fluent 触觉确认 + 键盘焦点描边(与 FluIconAction 同款)
        scale: ctl.down ? 0.92 : 1.0
        Behavior on scale { NumberAnimation { duration: 150; easing.type: Easing.OutCubic } }
        contentItem: Label {
            text: ctl.text
            font.pixelSize: 14
            color: ctl.enabled ? win.textColor : win.subText
            horizontalAlignment: Text.AlignHCenter
            verticalAlignment: Text.AlignVCenter
        }
        background: Rectangle {
            radius: win.rControl
            color: ctl.down ? Qt.rgba(0, 0, 0, 0.08)
                 : (ctl.hovered || ctl.activeFocus) ? (win.dark ? Qt.rgba(1, 1, 1, 0.07) : Qt.rgba(0, 0, 0, 0.05))
                 : "transparent"
            border.width: ctl.activeFocus ? 1 : 0
            border.color: win.accent
            Behavior on color { ColorAnimation { duration: 120 } }
        }
        // 手型光标:同 FluIconAction,所有透明底图标按钮都可点击
        MouseArea { anchors.fill: parent; acceptedButtons: Qt.NoButton; cursorShape: Qt.PointingHandCursor }
    }

    // 开关:Win11 的「拨杆」样式,滑块用强调色
    component FluSwitch: Switch {
        id: ctl
        implicitWidth: 40
        implicitHeight: 20
        hoverEnabled: true
        indicator: Rectangle {
            implicitWidth: 40
            implicitHeight: 20
            radius: 10
            color: ctl.checked ? win.accent
                 : (ctl.hovered ? win.strokeStrong : win.dark ? "#484848" : "#c8c8c8")
            border.width: 0
            Behavior on color { ColorAnimation { duration: 130 } }
            Rectangle {
                x: ctl.checked ? parent.width - width - 3 : 3
                y: 3
                width: 14
                height: 14
                radius: 7
                color: "#ffffff"
                // Win11 滑块悬浮时略微放大
                scale: ctl.hovered ? 1.12 : 1.0
                Behavior on x { NumberAnimation { duration: 130; easing.type: Easing.OutCubic } }
                Behavior on scale { NumberAnimation { duration: 130 } }
            }
        }
        contentItem: Item {}
    }

    // 输入框:Fluent 风格 = 底部 1px 描边 + 悬浮加深 + 聚焦转强调色
    component FluTextField: TextField {
        id: ctl
        implicitHeight: win.hControl
        font.pixelSize: 13
        color: win.textColor
        placeholderTextColor: win.subText
        selectionColor: win.accent
        selectedTextColor: "#ffffff"
        leftPadding: 4
        rightPadding: 4
        background: Item {
            implicitHeight: win.hControl
            Rectangle {
                anchors.fill: parent
                radius: win.rControl
                color: win.dark ? Qt.rgba(1, 1, 1, 0.04) : Qt.rgba(0, 0, 0, 0.02)
                border.width: 0
                Behavior on color { ColorAnimation { duration: 120 } }
            }
            // Fluent 的标志性「底边线」指示器(收敛在控件宽度内,不外扩)
            Rectangle {
                anchors { left: parent.left; right: parent.right; bottom: parent.bottom }
                height: 1
                color: ctl.activeFocus ? win.accent : (ctl.hovered ? win.strokeStrong : win.stroke)
                Behavior on color { ColorAnimation { duration: 120 } }
            }
        }
    }

    // 全局字体:一次设定,所有 Label/Button/TextField 继承
    Component.onCompleted: {
        // 默认落点:当前屏幕的右侧居中。必须加上屏幕原点(virtualX/virtualY):
        // 多显示器时副屏原点不是 (0,0),直接用 desktopAvailableWidth 会落错屏
        x = Screen.virtualX + Screen.desktopAvailableWidth - width - 20
        y = Screen.virtualY + Math.max(0, (Screen.desktopAvailableHeight - height) / 2 - 24)
        var s = ctrl.loadSettings()
        // 示例按钮:config.samples 为 null 表示从未自定义 → 用内置默认;[] 是用户主动清空
        if (s.samples !== null && s.samples !== undefined && Array.isArray(s.samples))
            // 归一化:侧边栏旧版本存过纯字符串数组,统一转成 {icon,text,prompt},
            // 否则设置区的 modelData.icon / modelData.text 会渲染成空行
            win.samples = s.samples.map(function (x) {
                if (typeof x === "string")
                    return { icon: win.icoBulb, text: x, prompt: x }
                return { icon: x.icon || win.icoBulb, text: x.text || x.prompt || "", prompt: x.prompt || x.text || "" }
            })
        // 启动时 Excel 里若已有框选,直接附上
        if (ctrl.selectionRef) {
            win.lastSelRef = ctrl.selectionRef
            win.selAttached = true
        }
        if (!s.api_key) settingsDialog.openNormal()
        // 冒烟测试专用:带 --smoke-dialog 启动时自动打开设置对话框,便于截图核对 UI
        else if (Qt.application.arguments.indexOf("--smoke-dialog") >= 0) settingsDialog.openNormal()
    }

    // 窗口落点保护:显示后把窗口钳回当前屏幕的可见区域。
    // 背景:预显示阶段设置的坐标在 DPI 缩放(125%/150%)下会被二次换算,
    // 窗口会整体落到屏幕外(标题栏不可达,只能靠 Alt+Tab 救)。显示完成后再
    // 校正一次,落点一定可见;若用户自己把窗口拖到副屏也只在下次启动时校正。
    onVisibleChanged: if (visible) Qt.callLater(win.clampIntoScreen)

    function clampIntoScreen() {
        var sx = Screen.virtualX, sy = Screen.virtualY
        var aw = Screen.desktopAvailableWidth, ah = Screen.desktopAvailableHeight
        if (!isFinite(sx) || !isFinite(sy) || !isFinite(aw) || !isFinite(ah) || aw <= 0) return
        var minX = sx + 4, maxX = Math.max(minX, sx + aw - width - 4)
        var minY = sy + 4, maxY = Math.max(minY, sy + ah - 56)  // 56:给标题栏留出可拖拽高度
        x = Math.round(Math.max(minX, Math.min(x, maxX)))
        y = Math.round(Math.max(minY, Math.min(y, maxY)))
    }

    // 点关闭按钮 = 转入托盘后台运行:进程不退出,Excel/WPS 侧边栏的本地服务保持可用
    // (guard 监控因此也不会再「关掉又弹回来」)。真正退出走托盘菜单「退出」。
    // 冒烟测试实例(--smoke)是临时窗口,关闭就是真退出,不留后台进程。
    onClosing: (close) => {
        if (Qt.application.arguments.indexOf("--smoke") >= 0) { close.accepted = true; return }
        close.accepted = false
        win.hide()
        ctrl.windowHiddenToTray()
    }

    // 返回 true 表示消息已真正发出(调用方据此决定是否清空输入框)。
    // 发不出去的三种情况都给明确反馈,避免「点了没反应/文字被吞」:
    //   空文本 → false;忙碌 → toast;缺配置 → 先弹设置(保留原文,配好即发)。
    function send(text) {
        text = ("" + text).trim()
        if (!text) return false
        if (ctrl.busy) { toast(qsTr("正在处理上一条请求,请稍候…")); return false }
        // 发前先查配置:缺 base_url 或缺 Key(本地 Ollama 除外)时直接弹设置,
        // 不清空输入,用户配好回来文字还在(与 chat.py 的 _needs_key 口径一致)。
        var s = ctrl.loadSettings()
        var base = (s.base_url || "").trim()
        var local = base.indexOf("127.0.0.1") >= 0 || base.indexOf("localhost") >= 0
        if (!base || (!s.api_key && !local)) {
            toast(qsTr("请先配置模型服务与 API Key"))
            settingsDialog.openNormal()
            return false
        }
        // 选区附件:点过 × 的(同一选区)不带;默认自动带上当前框选内容
        var withSel = win.selAttached && ctrl.selectionRef !== ""
        ctrl.sendMessage(text, withSel)
        // 发送即回到底部:上一轮若停在上方浏览,新消息与回复仍会实时跟随
        chatList.follow = true
        return true
    }
    function applyThemeMode() {
        ctrl.setThemeMode(autoThemeSwitch.checked ? "auto" : (themeSwitch.checked ? "dark" : "light"))
    }

    function toast(msg) {
        toastLabel.text = msg
        toast.show()
    }

    // ---------------- 事件接线 ----------------
    Connections {
        target: ctrl
        function onChatEvent(ev) {
            switch (ev.kind) {
            case "user_msg":
                chatModel.append({ kind: "user", text: ev.text,
                                   sel: ev.with_selection ? ctrl.selectionRef : "",
                                   md: "", note: "", summary: "", brief: "", state: "" })
                break
            case "text_start":
                win.curAi = chatModel.count
                chatModel.append({ kind: "ai", text: "", md: "", note: "", summary: "", brief: "", state: "" })
                break
            case "delta":
                if (win.curAi >= 0 && win.curAi < chatModel.count)
                    chatModel.setProperty(win.curAi, "md", chatModel.get(win.curAi).md + ev.text)
                break
            case "note_flush":
                if (win.curAi >= 0 && win.curAi < chatModel.count) {
                    var m = chatModel.get(win.curAi)
                    chatModel.setProperty(win.curAi, "note", m.md)
                    chatModel.setProperty(win.curAi, "md", "")
                }
                break
            case "plan":
                chatModel.append({ kind: "plan", text: ev.text, md: "", note: "", summary: "", brief: "", ref: "", state: "" })
                break
            case "step":
                win.sidIndex[ev.sid] = chatModel.count
                chatModel.append({ kind: "step", summary: ev.summary, brief: "", state: "running", text: "", md: "", note: "", ref: ev.ref || "" })
                break
            case "step_result":
                var i = win.sidIndex[ev.sid]
                if (i !== undefined && i < chatModel.count) {
                    chatModel.setProperty(i, "state", ev.ok ? "done" : "fail")
                    if (!ev.ok) chatModel.setProperty(i, "brief", ev.brief)
                }
                break
            case "status":
                activityLabel.text = ev.text
                break
            case "error":
                chatModel.append({ kind: "error", text: ev.text, md: "", note: "", summary: "", brief: "", ref: "", state: "" })
                break
            case "cancelled":
                activityLabel.text = ""
                chatModel.append({ kind: "stopped", text: qsTr("已停止"), md: "", note: "", summary: "", brief: "", ref: "", state: "" })
                break
            case "clear":
                chatModel.clear()
                win.sidIndex = ({})
                win.curAi = -1
                break
            case "toast":
                win.toast(ev.text)
                break
            case "confirm":
                // 危险工具二次确认:底部弹出确认条,等用户决定
                pendingConfirm = { sid: ev.sid, tool: ev.tool, summary: ev.summary }
                break
            case "need_settings":
                win.toast("请先配置模型服务与 API Key")
                settingsDialog.openNormal()
                break
            }
        }
        function onTestFinished(ok, msg) {
            testResult.text = msg
            testResult.color = ok ? win.okColor : win.badColor
        }
        function onStatusChanged() {
            if (settingsDialog.visible && !ctrl.installBusy) {
                settingsDialog._watchGuard = true
                try { watchSwitch.checked = ctrl.watchEnabled } catch (e) {}
                settingsDialog._watchGuard = false
            }
            // 选区附件:用户在 Excel 里换了框选区域 → 自动重新附上;清掉选区 → 附件收起
            var ref = ctrl.selectionRef
            if (ref && ref !== win.lastSelRef) {
                win.lastSelRef = ref
                win.selAttached = true
            } else if (!ref) {
                win.lastSelRef = ""
                win.selAttached = false
            }
        }
    }

    // ---------------- 顶栏 ----------------
    header: ToolBar {
        // 排版修复:contentItem 不能用 anchors.fill + margins —— 控件本身会把 contentItem
        // 布置到 padding 区域,双重锚定在窗口尺寸变化时会互相打架,导致顶栏偶发错位/裁切。
        // 间距一律交给 padding,由布局系统单点决定。
        leftPadding: 14
        rightPadding: 8
        topPadding: 6
        bottomPadding: 6
        background: Rectangle {
            color: win.layerColor
            // Fluent:层级之间用一条细分隔线而非阴影
            Rectangle {
                anchors { left: parent.left; right: parent.right; bottom: parent.bottom }
                height: 1
                color: win.stroke
            }
        }
        contentItem: RowLayout {
            spacing: 8

            // 应用图标:与 exe / 托盘 / 快捷方式同一份资源,顶栏更有产品感
            Image {
                source: "../assets/icon.png"
                Layout.preferredWidth: 18
                Layout.preferredHeight: 18
                fillMode: Image.PreserveAspectFit
                smooth: true
                mipmap: true
                asynchronous: true
            }
            Label {
                text: qsTr("表答 SheetTalk")
                font.pixelSize: 14
                font.weight: Font.DemiBold
                color: win.textColor
                elide: Text.ElideRight
                Layout.fillWidth: true
                Layout.minimumWidth: 0
            }

            AbstractButton {
                id: statusPill
                implicitHeight: 28
                Layout.maximumWidth: 220
                hoverEnabled: true
                // Fluent 触觉:按下压缩反馈(此胶囊可点击重连)
                scale: statusPill.down ? 0.97 : 1.0
                Behavior on scale { NumberAnimation { duration: 150; easing.type: Easing.OutCubic } }
                onClicked: ctrl.reconnect()
                contentItem: RowLayout {
                    spacing: 6
                    Rectangle {
                        id: statusDot
                        Layout.preferredWidth: 7
                        Layout.preferredHeight: 7
                        radius: 3.5
                        color: ctrl.connected ? win.okColor : (ctrl.warnState ? win.warnColor : win.badColor)
                        // 未连接时呼吸闪烁,提示「这里可以点一下重连」;恢复后回满不透明
                        SequentialAnimation {
                            running: !ctrl.connected
                            loops: Animation.Infinite
                            NumberAnimation { target: statusDot; property: "opacity"; to: 0.35; duration: 700; easing.type: Easing.InOutSine }
                            NumberAnimation { target: statusDot; property: "opacity"; to: 1.0; duration: 700; easing.type: Easing.InOutSine }
                            onStopped: statusDot.opacity = 1.0
                        }
                    }
                    Label {
                        text: ctrl.statusLine
                        font.pixelSize: 12
                        color: win.subText
                        elide: Text.ElideRight
                        Layout.maximumWidth: 180
                    }
                }
                background: Rectangle {
                    radius: win.rControl
                    color: statusPill.hovered ? win.accentSoft : win.dark ? Qt.rgba(1, 1, 1, 0.04) : Qt.rgba(0, 0, 0, 0.03)
                    border.width: 1
                    border.color: win.stroke
                    Behavior on color { ColorAnimation { duration: 120 } }
                }
                MouseArea { anchors.fill: parent; acceptedButtons: Qt.NoButton; cursorShape: Qt.PointingHandCursor }
            }
            FluIconAction {
                glyph: win.icoClear
                ToolTip.visible: hovered
                ToolTip.text: qsTr("清空对话")
                onClicked: ctrl.clearChat()
            }
            FluIconAction {
                id: settingsBtn
                glyph: win.icoSettings
                ToolTip.visible: hovered
                ToolTip.text: qsTr("设置")
                onClicked: settingsDialog.openNormal()
            }
        }
    }

    // 示例问题:图标与文案分开存,便于统一套用图标字体(不用 Emoji)
    readonly property var builtinSamples: [
        { icon: "\uE9D2", text: qsTr("分析这个表"),   prompt: qsTr("分析当前活动工作表的数据,总结关键发现和值得注意的点") },
        { icon: "\uE8B9", text: qsTr("生成图表"),     prompt: qsTr("根据当前表格的数据选一个合适的图表类型并生成图表") },
        { icon: "\uE8EF", text: qsTr("写同比公式"),   prompt: qsTr("查看表结构后,帮我写同比/环比增长公式并填充(如果表里没有合适的列,请说明需要什么)") },
        { icon: "\uE71C", text: qsTr("清洗建议"),     prompt: qsTr("检查当前工作表的数据质量问题(空值、重复、格式不一致等)并给出清洗建议") }
    ]
    // 实际展示的示例按钮:设置里可增删并持久化(config.samples);空状态引导与输入框上方共用这一份
    property var samples: builtinSamples

    // 选区附件:Excel 里框选后自动作为本轮上下文;点 × 移除后同一选区不再自动附回(换选区会重新附上)
    property bool selAttached: false
    property string lastSelRef: ""
    property var selPreviewData: ({})   // 点选区芯片后读到的内容预览
    property var pendingConfirm: ({})   // 危险工具待确认:{sid, tool, summary}

    function persistSamples() {
        ctrl.saveSamples(win.samples)
    }
    function addSample() {
        var t = newSampleText.text.trim()
        var p = newSamplePrompt.text.trim()
        if (!t || !p) { win.toast(qsTr("请填写按钮文字与提示词")); return }
        var arr = win.samples.slice()
        arr.push({ icon: win.icoBulb, text: t, prompt: p })
        win.samples = arr
        win.persistSamples()
        newSampleText.text = ""
        newSamplePrompt.text = ""
        win.toast(qsTr("已添加,立即生效"))
    }
    function removeSample(idx) {
        var arr = win.samples.slice()
        if (idx < 0 || idx >= arr.length) return
        arr.splice(idx, 1)
        win.samples = arr
        win.persistSamples()
    }

    // ---------------- 聊天列表 ----------------
    ListModel { id: chatModel }

    ListView {
        id: chatList
        anchors.fill: parent
        model: chatModel
        spacing: 10
        clip: true
        header: Item { width: chatList.width; height: 10 }
        footer: Item { width: chatList.width; height: 12 }
        ScrollBar.vertical: FluScrollBar { }

        // 新消息淡入:轻量不干扰流式期间跟随底部;populate 覆盖历史恢复时的批量创建
        add: Transition {
            NumberAnimation { property: "opacity"; from: 0; to: 1; duration: 140 }
        }
        populate: Transition {
            NumberAnimation { property: "opacity"; from: 0; to: 1; duration: 140 }
        }

        property bool follow: true
        onAtYEndChanged: follow = atYEnd
        onCountChanged: if (follow) Qt.callLater(positionViewAtEnd)
        onContentHeightChanged: if (follow) Qt.callLater(positionViewAtEnd)

        delegate: DelegateChooser {
            role: "kind"

            // ---- 用户消息 ----
            DelegateChoice {
                roleValue: "user"
                Item {
                    width: chatList.width
                    height: userBubble.height + 2
                    Rectangle {
                        id: userBubble
                        anchors.right: parent.right
                        anchors.rightMargin: 14
                        width: Math.min(parent.width * 0.82, userCol.implicitWidth + 30)
                        height: userCol.implicitHeight + 22
                        // Fluent:用户气泡用强调色纵向渐变 + 8px 圆角(与侧边栏同款)
                        radius: win.rCard
                        gradient: Gradient {
                            GradientStop { position: 0; color: win.accent }
                            GradientStop { position: 1; color: win.grad2 }
                        }

                        ColumnLayout {
                            id: userCol
                            x: 14; y: 11
                            width: parent.width - 28
                            spacing: 5

                            Text {
                                id: userText
                                Layout.fillWidth: true
                                text: model.text
                                color: "#ffffff"
                                font.pixelSize: 14
                                wrapMode: Text.Wrap
                                textFormat: Text.PlainText
                            }
                            // 附件行:这条消息随带的选区(Excel 里框选的内容)
                            // fillWidth 必须给:否则行宽按内容自然宽计算,长选区文本会伸出气泡
                            RowLayout {
                                id: selRow
                                Layout.fillWidth: true
                                visible: (model.sel || "") !== ""
                                spacing: 5
                                FluIcon {
                                    text: win.icoSheet
                                    font.pixelSize: 11
                                    color: Qt.rgba(1, 1, 1, 0.85)
                                }
                                Label {
                                    text: model.sel
                                    font.pixelSize: 11
                                    color: Qt.rgba(1, 1, 1, 0.85)
                                    elide: Text.ElideRight
                                    Layout.fillWidth: true
                                    Layout.maximumWidth: 240
                                    horizontalAlignment: Text.AlignLeft
                                }
                            }
                        }
                    }
                }
            }

            // ---- AI 回答(Markdown) ----
            DelegateChoice {
                roleValue: "ai"
                Item {
                    width: chatList.width
                    height: aiBubble.height + 2
                    Rectangle {
                        id: aiBubble
                        anchors.left: parent.left
                        anchors.leftMargin: 14
                        width: Math.min(parent.width - 40, Math.max(120, aiBody.implicitWidth + 30))
                        height: aiCol.implicitHeight + 22
                        radius: win.rCard
                        color: win.layerColor
                        border.width: 1
                        border.color: win.stroke

                        ColumnLayout {
                            id: aiCol
                            x: 14; y: 11
                            width: parent.width - 28
                            spacing: 6

                            Label {
                                visible: model.note !== ""
                                text: model.note
                                font.pixelSize: 12
                                color: win.subText
                                wrapMode: Text.Wrap
                                Layout.fillWidth: true
                                maximumLineCount: 4
                                elide: Text.ElideRight
                            }
                            // 加载动画:卡片还没等到任何内容且任务进行中时显示
                            // 三点打字动画(比 BusyIndicator 在浅色气泡上更醒目)
                            Row {
                                visible: model.md === "" && model.note === "" && ctrl.busy
                                spacing: 4
                                Layout.leftMargin: 2

                                Repeater {
                                    model: 3
                                    Rectangle {
                                        width: 6
                                        height: 6
                                        radius: 3
                                        color: win.subText
                                        opacity: 0.25

                                        SequentialAnimation on opacity {
                                            loops: Animation.Infinite
                                            PauseAnimation { duration: index * 220 }
                                            NumberAnimation { to: 0.9; duration: 180 }
                                            NumberAnimation { to: 0.25; duration: 380 }
                                        }
                                    }
                                }
                            }
                            Text {
                                id: aiBody
                                Layout.fillWidth: true
                                text: model.md
                                textFormat: Text.MarkdownText
                                wrapMode: Text.Wrap
                                font.pixelSize: 14
                                color: win.textColor
                                onLinkActivated: function(link) { Qt.openUrlExternally(link) }
                            }
                        }
                    }
                }
            }

            // ---- 执行计划(Plan 模式:动手前先给用户看计划) ----
            DelegateChoice {
                roleValue: "plan"
                Item {
                    width: chatList.width
                    // 高度必须由内容撑开:ColumnLayout 用 anchors.fill 会把可用高度
                    // 减掉上下 margins,若 Item 只按 implicitHeight + 20 设高,最后一行会被裁掉。
                    height: planCol.implicitHeight + 36
                    Rectangle {
                        anchors.fill: parent
                        anchors.margins: 6
                        radius: win.rCard
                        color: win.accentSoft
                        border.width: 1
                        border.color: win.accent
                    }
                    ColumnLayout {
                        id: planCol
                        x: 16
                        y: 16
                        width: parent.width - 32
                        spacing: 6
                        RowLayout {
                            spacing: 6
                            FluIcon {
                                text: win.icoList
                                font.pixelSize: 13
                                color: win.accent
                            }
                            Label {
                                text: qsTr("执行计划")
                                font.pixelSize: 12
                                font.weight: Font.DemiBold
                                color: win.accent
                            }
                        }
                        Label {
                            Layout.fillWidth: true
                            text: model.text
                            font.pixelSize: 12
                            color: win.textColor
                            wrapMode: Text.Wrap
                        }
                    }
                }
            }

            // ---- 工具执行步骤 ----
            DelegateChoice {
                roleValue: "step"
                Item {
                    width: chatList.width
                    height: stepCard.height + 2
                    // 步骤包一层浅色底卡:工具调用在视觉上成组;失败时整卡染红,比单行文字更醒目
                    Rectangle {
                        id: stepCard
                        x: 14
                        width: parent.width - 28
                        height: stepRow.height + 12
                        radius: 6
                        color: model.state === "fail"
                               ? (win.dark ? Qt.rgba(0.77, 0.17, 0.11, 0.16) : Qt.rgba(0.77, 0.17, 0.11, 0.08))
                               : (win.dark ? Qt.rgba(1, 1, 1, 0.045) : Qt.rgba(0, 0, 0, 0.03))
                        Behavior on color { ColorAnimation { duration: 150 } }
                        RowLayout {
                            id: stepRow
                            x: 8
                            y: 6
                            width: parent.width - 16
                            spacing: 8
                            FluIcon {
                                text: model.state === "done" ? win.icoCheck
                                      : (model.state === "fail" ? win.icoError : win.icoClock)
                                color: model.state === "done" ? win.okColor : (model.state === "fail" ? win.badColor : win.warnColor)
                                font.pixelSize: 12
                            }
                            Label {
                                Layout.fillWidth: true
                                text: model.state === "fail" && model.brief !== "" ? model.summary + "  —  " + model.brief : model.summary
                                font.pixelSize: 12
                                color: model.state === "fail" ? win.badColor : win.subText
                                wrapMode: Text.Wrap
                                elide: Text.ElideRight
                                maximumLineCount: 2
                            }
                            // 引用溯源:灰底小标签,标明这一步动的是哪块区域(对齐 Copilot 的「看源」)
                            Rectangle {
                                visible: model.ref !== ""
                                Layout.alignment: Qt.AlignTop
                                implicitWidth: refLabel.implicitWidth + 12
                                implicitHeight: refLabel.implicitHeight + 4
                                radius: 4
                                color: win.dark ? Qt.rgba(1, 1, 1, 0.06) : Qt.rgba(0, 0, 0, 0.05)
                                Label {
                                    id: refLabel
                                    anchors.centerIn: parent
                                    text: model.ref
                                    font.pixelSize: 10
                                    color: win.subText
                                }
                                ToolTip.visible: refHover.hovered
                                ToolTip.text: qsTr("这一步涉及的区域")
                                HoverHandler { id: refHover }
                            }
                        }
                    }
                }
            }

            // ---- 已停止 ----
            DelegateChoice {
                roleValue: "stopped"
                RowLayout {
                    x: 14
                    width: parent.width - 28
                    spacing: 6
                    FluIcon {
                        text: win.icoWarn
                        font.pixelSize: 12
                        color: win.subText
                    }
                    Label {
                        text: model.text
                        font.pixelSize: 11
                        color: win.subText
                        // 不给 fillWidth 的话,Label 会占满剩余宽度并在内部居中,
                        // 看起来就跟步骤列表对不齐了
                        Layout.fillWidth: true
                        horizontalAlignment: Text.AlignLeft
                    }
                }
            }

            // ---- 错误 ----
            DelegateChoice {
                roleValue: "error"
                Item {
                    width: chatList.width
                    height: errBox.height + 2
                    Rectangle {
                        id: errBox
                        anchors.left: parent.left
                        anchors.leftMargin: 14
                        width: Math.min(parent.width - 40, errText.implicitWidth + 28)
                        height: errText.implicitHeight + 20
                        radius: 12
                        color: win.dark ? "#2b1d1e" : "#fdeceb"
                        border.color: win.badColor
                        Text {
                            id: errText
                            x: 14; y: 10
                            width: Math.min(parent.width - 28, implicitWidth)
                            text: model.text
                            color: win.badColor
                            font.pixelSize: 13
                            wrapMode: Text.Wrap
                            textFormat: Text.PlainText
                        }
                    }
                }
            }
        }
    }

    // ---------------- 空状态引导 ----------------
    ColumnLayout {
        anchors.fill: parent
        visible: chatModel.count === 0
        spacing: 12

        Item { Layout.fillHeight: true; Layout.fillWidth: true }

        ColumnLayout {
            Layout.alignment: Qt.AlignHCenter
            Layout.maximumWidth: 360
            spacing: 10

            // 品牌渐变方块:与侧边栏空状态同款,比裸图标更有产品感
            Rectangle {
                Layout.alignment: Qt.AlignHCenter
                width: 46; height: 46; radius: 12
                gradient: Gradient {
                    GradientStop { position: 0; color: win.grad1 }
                    GradientStop { position: 1; color: win.grad2 }
                }
                FluIcon {
                    anchors.centerIn: parent
                    text: win.icoBulb
                    font.pixelSize: 22
                    color: "#ffffff"
                }
            }
            Label {
                Layout.alignment: Qt.AlignHCenter
                text: qsTr("让 AI 帮你操作表格")
                font.pixelSize: 17
                font.weight: Font.DemiBold
                color: win.textColor
            }
            Label {
                Layout.alignment: Qt.AlignHCenter
                Layout.fillWidth: true
                horizontalAlignment: Text.AlignHCenter
                wrapMode: Text.Wrap
                text: qsTr("支持 Microsoft Excel 和 WPS 表格\n用中文直接说出你想做的事")
                font.pixelSize: 13
                color: win.subText
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.topMargin: 8
                implicitHeight: guideCol.implicitHeight + 24
                radius: 16
                color: win.surface
                border.color: win.outline

                ColumnLayout {
                    id: guideCol
                    anchors.fill: parent
                    anchors.margins: 12
                    spacing: 8

                    RowLayout {
                        spacing: 8
                        Rectangle { width: 18; height: 18; radius: 9
                            gradient: Gradient { orientation: Gradient.Horizontal
                                GradientStop { position: 0; color: win.grad1 }
                                GradientStop { position: 1; color: win.grad2 } }
                            Label { anchors.centerIn: parent; text: "1"; color: "white"; font.pixelSize: 10 }
                        }
                        Label { text: qsTr("打开一个 Excel / WPS 表格(先开表格,再说话)"); font.pixelSize: 13; color: win.textColor; wrapMode: Text.Wrap; Layout.fillWidth: true }
                    }
                    RowLayout {
                        spacing: 8
                        Rectangle { width: 18; height: 18; radius: 9
                            gradient: Gradient { orientation: Gradient.Horizontal
                                GradientStop { position: 0; color: win.grad1 }
                                GradientStop { position: 1; color: win.grad2 } }
                            Label { anchors.centerIn: parent; text: "2"; color: "white"; font.pixelSize: 10 }
                        }
                        Label { text: qsTr("点右上角齿轮填好模型 API Key(只需一次)"); font.pixelSize: 13; color: win.textColor; wrapMode: Text.Wrap; Layout.fillWidth: true }
                    }
                    RowLayout {
                        spacing: 8
                        Rectangle { width: 18; height: 18; radius: 9
                            gradient: Gradient { orientation: Gradient.Horizontal
                                GradientStop { position: 0; color: win.grad1 }
                                GradientStop { position: 1; color: win.grad2 } }
                            Label { anchors.centerIn: parent; text: "3"; color: "white"; font.pixelSize: 10 }
                        }
                        Label { text: qsTr("下方直接说中文需求,AI 自动改表"); font.pixelSize: 13; color: win.textColor; wrapMode: Text.Wrap; Layout.fillWidth: true }
                    }

                    RowLayout {
                        Layout.topMargin: 4
                        Layout.fillWidth: true
                        spacing: 8
                        FluFilledButton {
                            text: qsTr("重新连接")
                            onClicked: ctrl.reconnect()
                        }
                        FluSecondaryButton {
                            text: qsTr("启动 Excel")
                            onClicked: ctrl.launch("excel")
                        }
                        FluSecondaryButton {
                            text: qsTr("启动 WPS")
                            onClicked: ctrl.launch("wps")
                        }
                    }
                    RowLayout {
                        Layout.topMargin: 2
                        Layout.fillWidth: true
                        spacing: 8
                        FluSecondaryButton {
                            text: qsTr("一键安装到 Excel / WPS(含自启动)")
                            Layout.fillWidth: true
                            onClicked: settingsDialog.openInstall()
                        }
                    }
                }
            }

            Label {
                Layout.topMargin: 8
                Layout.alignment: Qt.AlignHCenter
                text: qsTr("小白先点下面一句试试")
                font.pixelSize: 12
                color: win.subText
            }
            Repeater {
                // 与输入框上方按钮共用 win.samples:设置里增删,这里自动跟着变。
                // 从纯文字行升级为胶囊按钮,与 Fluent 的链接按钮观感一致
                model: win.samples
                Rectangle {
                    id: emptyChip
                    required property var modelData
                    Layout.alignment: Qt.AlignHCenter
                    implicitWidth: chipInner.implicitWidth + 22
                    implicitHeight: 30
                    radius: 15
                    color: chipHover.hovered ? win.accentSoft : win.layerColor
                    border.width: 1
                    border.color: chipHover.hovered ? win.accent : win.stroke
                    Behavior on color { ColorAnimation { duration: 120 } }
                    // Fluent 触觉确认:按下压缩到 0.97
                    scale: chipPress.pressed ? 0.97 : 1.0
                    Behavior on scale { NumberAnimation { duration: 150; easing.type: Easing.OutCubic } }
                    Row {
                        id: chipInner
                        anchors.centerIn: parent
                        spacing: 6
                        FluIcon {
                            text: emptyChip.modelData.icon
                            font.pixelSize: 12
                            color: win.accent
                        }
                        Label {
                            text: emptyChip.modelData.text
                            font.pixelSize: 12
                            color: win.textColor
                        }
                    }
                    MouseArea {
                        id: chipPress
                        anchors.fill: parent
                        cursorShape: Qt.PointingHandCursor
                        onClicked: win.send(emptyChip.modelData.prompt)
                    }
                    HoverHandler { id: chipHover }
                }
            }
            Label {
                visible: win.samples.length === 0
                Layout.alignment: Qt.AlignHCenter
                text: qsTr("(没有示例按钮,可在设置里添加)")
                font.pixelSize: 12
                color: win.subText
            }
        }

        Item { Layout.fillHeight: true; Layout.fillWidth: true }
    }

    // ---------------- 底部输入区 ----------------
    footer: Pane {
        background: Rectangle {
            color: win.layerColor
            Rectangle {
                anchors { left: parent.left; right: parent.right; top: parent.top }
                height: 1
                color: win.stroke
            }
        }
        contentItem: ColumnLayout {
            spacing: 8

            // ---- 危险操作确认条:Agent 要执行删除/清空/关簿等操作时弹出 ----
            Rectangle {
                Layout.fillWidth: true
                visible: !!pendingConfirm.sid
                color: win.dark ? "#3a2f1d" : "#fff7e6"
                border.color: win.warnColor
                border.width: 1
                radius: win.rCard
                height: confirmCol.implicitHeight + 20

                ColumnLayout {
                    id: confirmCol
                    anchors { left: parent.left; right: parent.right; margins: 10 }
                    anchors.verticalCenter: parent.verticalCenter
                    spacing: 6

                    RowLayout {
                        spacing: 5
                        FluIcon { text: win.icoWarn; font.pixelSize: 13; color: win.warnColor }
                        Label {
                            text: qsTr("需要你的确认")
                            font.pixelSize: 12
                            font.weight: Font.DemiBold
                            color: win.warnColor
                        }
                    }
                    Label {
                        text: pendingConfirm.summary || pendingConfirm.tool || ""
                        font.pixelSize: 12
                        color: win.textColor
                        wrapMode: Text.Wrap
                        Layout.fillWidth: true
                    }
                    RowLayout {
                        Layout.fillWidth: true
                        spacing: 8
                        // Codex 风格:处理方式用下拉选择,再点确认
                        ComboBox {
                            id: confirmChoice
                            Layout.preferredWidth: 220
                            implicitHeight: win.hControl
                            font.pixelSize: 13
                            model: ["仅本次允许执行", "本会话始终允许该工具", "拒绝执行"]
                            contentItem: Label {
                                text: confirmChoice.displayText
                                font.pixelSize: 13
                                color: win.textColor
                                verticalAlignment: Text.AlignVCenter
                                leftPadding: 10
                            }
                            background: Rectangle {
                                radius: win.rControl
                                color: win.dark ? Qt.rgba(1, 1, 1, 0.04) : Qt.rgba(0, 0, 0, 0.02)
                                border.width: 1
                                border.color: confirmChoice.activeFocus ? win.accent : win.stroke
                            }
                        }
                        // 红色只给拒绝语义;允许保持品牌绿
                        FluFilledButton {
                            text: qsTr("确认")
                            bgColor: confirmChoice.currentIndex === 2 ? win.badColor : win.accent
                            bgHovered: confirmChoice.currentIndex === 2
                                       ? Qt.lighter(win.badColor, 1.15) : win.accentHover
                            bgDown: confirmChoice.currentIndex === 2
                                    ? Qt.darker(win.badColor, 1.15) : win.accentPressed
                            onClicked: {
                                var c = confirmChoice.currentIndex
                                // 0=仅本次 1=本会话始终允许(需勾选豁免) 2=拒绝
                                ctrl.confirmRun(pendingConfirm.sid, c !== 2, c === 1 && cbRemember.checked)
                                pendingConfirm = ({})
                            }
                        }
                    }
                    // 二次勾选:避免手滑永久放行
                    RowLayout {
                        visible: confirmChoice.currentIndex === 1
                        spacing: 6
                        Layout.fillWidth: true
                        CheckBox {
                            id: cbRemember
                            font.pixelSize: 11
                        }
                        Label {
                            text: cbRemember.checked
                                  ? qsTr("本会话内该工具不再弹确认")
                                  : qsTr("勾选后本会话内该工具不再弹确认;不勾选按仅本次处理")
                            font.pixelSize: 11
                            color: win.subText
                            elide: Text.ElideRight
                            Layout.fillWidth: true
                        }
                    }
                }
            }

            // ---- 选区附件:Excel 里框选后自动带上,点 × 移除(同一选区不再自动附回) ----
            RowLayout {
                Layout.fillWidth: true
                spacing: 8
                visible: win.selAttached && ctrl.selectionRef !== ""

                Rectangle {
                    implicitHeight: 26
                    implicitWidth: selChipInner.implicitWidth + 18
                    radius: 13
                    color: win.accentSoft
                    border.width: 1
                    border.color: win.accent

                    RowLayout {
                        id: selChipInner
                        anchors.verticalCenter: parent.verticalCenter
                        anchors.left: parent.left
                        anchors.leftMargin: 9
                        spacing: 5

                        FluIcon {
                            text: win.icoSheet
                            font.pixelSize: 12
                            color: win.accent
                        }
                        Label {
                            text: ctrl.selectionRef
                            font.pixelSize: 12
                            color: win.textColor
                            elide: Text.ElideRight
                            Layout.maximumWidth: 240

                            // 点选区文字:弹出内容预览(当前框选的值,前 6 行)
                            MouseArea {
                                anchors.fill: parent
                                cursorShape: Qt.PointingHandCursor
                                onClicked: {
                                    selPreviewData = ctrl.selectionPreview()
                                    selPreview.open()
                                }
                            }
                        }
                        Label {
                            visible: ctrl.selectionDims !== ""
                            text: ctrl.selectionDims
                            font.pixelSize: 11
                            color: win.subText
                        }
                        FluIconAction {
                            glyph: win.icoClose
                            glyphSize: 10
                            implicitWidth: 18
                            implicitHeight: 18
                            ToolTip.visible: hovered
                            ToolTip.text: qsTr("不把该选区作为附件")
                            onClicked: win.selAttached = false
                        }
                    }
                }
                Label {
                    Layout.fillWidth: true
                    text: qsTr("已框选,将随消息发给 AI")
                    font.pixelSize: 11
                    color: win.subText
                    elide: Text.ElideRight
                    horizontalAlignment: Text.AlignRight
                }
            }

            // ---- 示例按钮:Flow 自动换行,增减/改窗口宽度都会自动重排 ----
            Flow {
                Layout.fillWidth: true
                Layout.preferredHeight: implicitHeight
                spacing: 8
                Repeater {
                    model: win.samples
                    FluSecondaryButton {
                        id: sampleBtn
                        property var sample: modelData
                        enabled: !ctrl.busy
                        onClicked: win.send(sampleBtn.sample.prompt)
                        contentItem: RowLayout {
                            spacing: 6
                            FluIcon {
                                text: sampleBtn.sample.icon
                                font.pixelSize: 12
                                // 悬浮时图标转强调色(Fluent 图标微交互)
                                color: sampleBtn.hovered ? win.accent : win.subText
                                Behavior on color { ColorAnimation { duration: 120 } }
                            }
                            Label {
                                text: sampleBtn.sample.text
                                font.pixelSize: 12
                                color: sampleBtn.enabled ? win.textColor : win.subText
                            }
                        }
                    }
                }
            }

            RowLayout {
                spacing: 10

                Rectangle {
                    Layout.fillWidth: true
                    Layout.preferredHeight: Math.min(input.implicitHeight + 20, 130)
                    // Fluent:输入区是「卡片」而非气泡,8px 圆角 + 描边
                    radius: win.rCard
                    color: win.layerColor
                    border.width: 1
                    border.color: input.activeFocus ? win.accent : win.stroke
                    Behavior on border.color { ColorAnimation { duration: 120 } }

                    TextArea {
                        id: input
                        anchors.fill: parent
                        anchors.margins: 9
                        placeholderText: qsTr("输入你想对表格做的事…")
                        placeholderTextColor: win.subText
                        color: win.textColor
                        font.pixelSize: 14
                        wrapMode: TextArea.Wrap
                        background: null
                        selectionColor: win.accent
                        selectedTextColor: "#ffffff"

                        Keys.onPressed: function(event) {
                            if ((event.key === Qt.Key_Return || event.key === Qt.Key_Enter)
                                && !(event.modifiers & Qt.ShiftModifier) && !input.inputMethodComposing) {
                                // 只有真正发出去才清空:忙/缺配置时文字保留,用户回来还能发
                                if (win.send(input.text)) input.clear()
                                event.accepted = true
                            }
                        }
                    }
                }

                RoundButton {
                    id: sendBtn
                    Layout.alignment: Qt.AlignBottom
                    Layout.bottomMargin: 2
                    implicitWidth: 40
                    implicitHeight: 40
                    // 忙碌时变成「停止」(对齐 Copilot 输入框右下角的 Stop)。
                    // 始终可点:空文本点击 = 聚焦输入框(不再出现「灰着点不动」的死按钮),
                    // 忙碌点击 = 停止,有文字点击 = 发送 —— 任何状态下点击都有明确反馈。
                    readonly property bool hasText: input.text.trim() !== ""
                    hoverEnabled: true
                    MouseArea { anchors.fill: parent; acceptedButtons: Qt.NoButton; cursorShape: Qt.PointingHandCursor }
                    contentItem: FluIcon {
                        text: ctrl.busy ? win.icoStop : win.icoSend
                        color: (ctrl.busy || sendBtn.hasText) ? "#ffffff" : win.accent
                        font.pixelSize: ctrl.busy ? 12 : 15
                        horizontalAlignment: Text.AlignHCenter
                        verticalAlignment: Text.AlignVCenter
                    }
                    background: Rectangle {
                        radius: width / 2
                        color: ctrl.busy ? (sendBtn.hovered ? Qt.rgba(0, 0, 0, 0.62) : Qt.rgba(0, 0, 0, 0.5))
                             : !sendBtn.hasText ? (sendBtn.hovered ? win.accentSoft : win.accentFaint)
                             : sendBtn.down ? win.accentPressed
                             : sendBtn.hovered ? win.accentHover
                             : win.accent
                        Behavior on color { ColorAnimation { duration: 120 } }
                        scale: sendBtn.down ? 0.94 : (sendBtn.hovered ? 1.06 : 1.0)
                        Behavior on scale { NumberAnimation { duration: 120 } }
                    }
                    ToolTip.visible: hovered
                    ToolTip.text: ctrl.busy ? qsTr("停止生成")
                                 : sendBtn.hasText ? qsTr("发送")
                                 : qsTr("先输入要做什么")
                    onClicked: {
                        if (ctrl.busy) {
                            ctrl.cancelRun()
                        } else if (sendBtn.hasText) {
                            if (win.send(input.text)) input.clear()
                        } else {
                            // 空文本:把光标请回输入框,避免「点了没反应」的困惑
                            input.forceActiveFocus()
                        }
                    }
                }
            }

            RowLayout {
                visible: ctrl.busy
                spacing: 8
                BusyIndicator {
                    running: ctrl.busy
                    implicitWidth: 18
                    implicitHeight: 18
                }
                Label {
                    id: activityLabel
                    text: qsTr("正在处理…")
                    font.pixelSize: 12
                    color: win.subText
                    Layout.fillWidth: true
                    elide: Text.ElideRight
                }
            }
            Label {
                visible: !ctrl.busy
                text: qsTr("Enter 发送 · Shift+Enter 换行 · 表格里的修改可用 Ctrl+Z 撤销")
                font.pixelSize: 11
                color: win.subText
                Layout.alignment: Qt.AlignHCenter
            }
        }
    }

    // 选区内容预览:点击底部选区芯片的文字弹出,展示当前框选的前几行值
    Popup {
        id: selPreview
        x: 12
        y: win.height - (win.footer ? win.footer.height : 0) - 190
        width: Math.min(win.width - 24, 400)
        padding: 10
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        background: Rectangle {
            radius: win.rCard
            color: win.layerColor
            border.width: 1
            border.color: win.stroke
        }
        contentItem: ColumnLayout {
            spacing: 4
            Label {
                text: ctrl.selectionRef
                font.pixelSize: 12
                font.weight: Font.DemiBold
                color: win.accent
                elide: Text.ElideRight
                Layout.fillWidth: true
            }
            Label {
                visible: !selPreviewData.values
                text: selPreviewData.note ? qsTr("当前没有选中区域")
                                          : qsTr("读取中…")
                font.pixelSize: 11
                color: win.subText
                Layout.fillWidth: true
            }
            Repeater {
                model: selPreviewData.values ? selPreviewData.values.slice(0, 6) : []
                Label {
                    required property var modelData
                    text: modelData.map(function (c) { return c === null ? "" : c }).join("  │  ")
                    font.pixelSize: 11
                    color: win.textColor
                    elide: Text.ElideRight
                    Layout.fillWidth: true
                }
            }
            Label {
                visible: !!selPreviewData.values && selPreviewData.values.length > 6
                text: qsTr("(仅预览前 6 行,发送后 AI 看到完整选区)")
                font.pixelSize: 10
                color: win.subText
            }
        }
    }

    // 启动自检异常横幅:红色,可关闭(不变量被破坏时明示,不等到运行时才炸)
    Rectangle {
        id: selfCheckBanner
        visible: startupIssues && startupIssues.length > 0
        anchors { top: parent.top; left: parent.left; right: parent.right }
        color: win.dark ? "#3a1d1e" : "#fdeceb"
        height: scCol.implicitHeight + 24
        z: 60

        ColumnLayout {
            id: scCol
            anchors { left: parent.left; right: parent.right; margins: 10 }
            anchors.verticalCenter: parent.verticalCenter
            spacing: 4

            RowLayout {
                Layout.fillWidth: true
                spacing: 6
                FluIcon { text: win.icoError; font.pixelSize: 13; color: win.badColor }
                Label {
                    text: qsTr("启动自检发现异常(功能可能受影响)")
                    font.pixelSize: 12
                    font.weight: Font.DemiBold
                    color: win.badColor
                    Layout.fillWidth: true
                }
                FluIconAction {
                    glyph: win.icoClose
                    glyphSize: 11
                    implicitWidth: 22
                    implicitHeight: 22
                    onClicked: selfCheckBanner.visible = false
                }
            }
            Label {
                text: Array.isArray(startupIssues) ? startupIssues.join("\n") : ""
                font.pixelSize: 11
                color: win.textColor
                wrapMode: Text.Wrap
                Layout.fillWidth: true
            }
        }
    }

    // 模态遮罩:对话框打开时压暗背景,形成 Fluent 的层级感。    // 只压暗 header/footer 之外的聊天区,避免连顶栏一起变灰(顶栏是常驻导航)。
    Rectangle {
        id: dimmer
        x: chatList.x
        y: win.header.height
        width: chatList.width
        height: win.height - win.header.height - (win.footer ? win.footer.height : 0)
        color: "#000000"
        opacity: settingsDialog.opened ? 0.32 : 0
        visible: opacity > 0
        z: 90
        Behavior on opacity { NumberAnimation { duration: 160 } }
    }

    // ---------------- 设置对话框 ----------------
    Dialog {
        id: settingsDialog
        z: 100
        modal: true
        parent: Overlay.overlay
        anchors.centerIn: parent
        width: Math.min(win.width - 36, 420)
        padding: 18
        closePolicy: Popup.CloseOnEscape
        // 内容变多(示例按钮管理)时限制高度,超出部分在 ScrollView 里滚动
        height: Math.min(settingsCol.implicitHeight + padding * 2, win.height - 80)
        background: Rectangle {
            // Fluent 对话框:8px 圆角 + 分隔描边 + 柔和阴影
            radius: win.rCard
            color: win.layerColor
            border.width: 1
            border.color: win.stroke
        }

        property var presets: ({})
        property var presetKeys: []

        property bool _watchGuard: false
        function openInstall() {
            openNormal()
            // 用纵向滚动条定位到安装区:position 是 0~1 归一化位置,要扣除滑块自身长度。
            // 之前用 settingsScroll.contentItem.contentY 在 Qt6 下静默失败(Qt6 的
            // ScrollView.contentItem 不是 Flickable,没有 contentY 可设)。
            var tries = 0
            function scrollOnce() {
                try {
                    var target = Math.max(0, installHeader.y - 12)
                    var maxY = Math.max(1, settingsCol.implicitHeight - settingsScroll.availableHeight)
                    var frac = Math.min(1, Math.max(0, target / maxY))
                    settingsBar.position = frac * (1 - settingsBar.size)
                } catch (e) {}
                if (++tries < 3) Qt.callLater(scrollOnce)
            }
            Qt.callLater(scrollOnce)
        }

        function openNormal() {
            presets = ctrl.presets()
            presetKeys = Object.keys(presets)
            providerBox.model = presetKeys
            var s = ctrl.loadSettings()
            var pi = presetKeys.indexOf(s.provider)
            providerBox.currentIndex = pi >= 0 ? pi : 0
            baseField.text = s.base_url
            modelField.text = s.model
            keyField.text = s.api_key
            // 回填时先钳制到 SpinBox 允许的范围,否则旧配置里的越界值会让 SpinBox 显示异常
            stepsSpin.value = Math.max(5, Math.min(Math.round(Number(s.max_steps) || 30), 500))
            planSwitch.checked = s.plan_mode === true
            // 主题三态:theme_mode=auto 跟随系统(默认)/ light / dark;旧配置只有 theme 键时也按 auto 处理
            var tm = s.theme_mode || "auto"
            autoThemeSwitch.checked = (tm === "auto")
            themeSwitch.checked = (tm === "dark")
            // 安装区:回显真实注册表状态(读注册表,幂等,不写);守卫避免回填触发写注册表
            settingsDialog._watchGuard = true
            try { watchSwitch.checked = ctrl.watchEnabled } catch (e) {}
            settingsDialog._watchGuard = false
            testResult.text = ""
            open()
        }

        contentItem: ScrollView {
            id: settingsScroll
            clip: true
            contentWidth: availableWidth
            ScrollBar.vertical: FluScrollBar { id: settingsBar }

            ColumnLayout {
                id: settingsCol
                width: settingsScroll.availableWidth
                spacing: 14

            RowLayout {
                spacing: 8
                Layout.fillWidth: true
                FluIcon {
                    text: win.icoSettings
                    font.pixelSize: 16
                    color: win.accent
                }
                Label {
                    text: qsTr("模型设置")
                    font.pixelSize: 15
                    font.weight: Font.DemiBold
                    color: win.textColor
                }
                Item { Layout.fillWidth: true }
                FluIconAction {
                    glyph: win.icoClose
                    glyphSize: 13
                    ToolTip.visible: hovered
                    ToolTip.text: qsTr("关闭")
                    onClicked: settingsDialog.close()
                }
            }

            Label { text: qsTr("服务商预设"); font.pixelSize: 12; color: win.subText }
            ComboBox {
                id: providerBox
                Layout.fillWidth: true
                implicitHeight: win.hControl
                font.pixelSize: 13
                textRole: "modelData"
                displayText: settingsDialog.presets[currentValue] ? settingsDialog.presets[currentValue].label : currentValue
                onActivated: {
                    baseField.text = settingsDialog.presets[settingsDialog.presetKeys[currentIndex]].base_url
                    modelField.text = settingsDialog.presets[settingsDialog.presetKeys[currentIndex]].model
                }
                contentItem: Label {
                    text: providerBox.displayText
                    font.pixelSize: 13
                    color: win.textColor
                    verticalAlignment: Text.AlignVCenter
                    leftPadding: 10
                    rightPadding: 28
                    elide: Text.ElideRight
                }
                background: Rectangle {
                    radius: win.rControl
                    color: win.dark ? Qt.rgba(1, 1, 1, 0.04) : Qt.rgba(0, 0, 0, 0.02)
                    border.width: 1
                    border.color: providerBox.activeFocus ? win.accent : win.stroke
                }
                delegate: ItemDelegate {
                    width: providerBox.width
                    height: 34
                    contentItem: Label {
                        text: settingsDialog.presets[modelData] ? settingsDialog.presets[modelData].label : modelData
                        font.pixelSize: 13
                        color: win.textColor
                        verticalAlignment: Text.AlignVCenter
                    }
                    highlighted: providerBox.highlightedIndex === index
                    background: Rectangle {
                        radius: win.rControl
                        color: highlighted ? win.accentSoft : "transparent"
                    }
                }
                // Fluent 下拉箭头:细线倒三角,替换 Basic 的粗箭头
                indicator: Text {
                    x: providerBox.width - width - 10
                    y: providerBox.topPadding + (providerBox.availableHeight - height) / 2
                    text: "⌄"
                    font.pixelSize: 14
                    color: providerBox.pressed ? win.subText : win.textColor
                }
            }

            Label { text: qsTr("接口地址(OpenAI 兼容 Base URL)"); font.pixelSize: 12; color: win.subText }
            FluTextField {
                id: baseField
                Layout.fillWidth: true
                placeholderText: "https://api.deepseek.com"
                selectByMouse: true
            }
            Label { text: qsTr("模型名称"); font.pixelSize: 12; color: win.subText }
            FluTextField {
                id: modelField
                Layout.fillWidth: true
                placeholderText: "deepseek-chat"
                selectByMouse: true
            }
            Label { text: qsTr("API Key(本地 Ollama 可留空)"); font.pixelSize: 12; color: win.subText }
            FluTextField {
                id: keyField
                Layout.fillWidth: true
                echoMode: TextInput.Password
                placeholderText: "sk-…"
                selectByMouse: true
            }

            Label { text: qsTr("单次任务最大工具步数(5–500,调大可完成更复杂的任务)"); font.pixelSize: 12; color: win.subText; Layout.fillWidth: true; wrapMode: Text.WordWrap }
            RowLayout {
                spacing: 10
                Layout.fillWidth: true

                FluSpinBox {
                    id: stepsSpin
                    from: 5; to: 500; stepSize: 1
                    editable: true
                    Layout.preferredWidth: 132
                }
                Label {
                    // 实时回显最终生效值(与 agent.py 的 max(5, min(v, 500)) 钳制口径一致)
                    text: qsTr("当前生效:%1 步").arg(Math.max(5, Math.min(Math.round(stepsSpin.value), 500)))
                    font.pixelSize: 12
                    color: win.subText
                    Layout.fillWidth: true
                }
            }
            Label {
                text: qsTr("简单任务 20–30 足够;批量改写、跨表分析等复杂任务建议 60–150。")
                font.pixelSize: 11
                color: win.subText
                wrapMode: Text.WordWrap
                Layout.fillWidth: true
            }

            // Fluent 对话框惯例:主题切换前先来一条分隔线,分组更清晰
            Rectangle {
                Layout.fillWidth: true
                Layout.topMargin: 2
                height: 1
                color: win.stroke
            }

            RowLayout {
                spacing: 8
                Label {
                    text: qsTr("跟随系统外观(亮 / 暗自动切换)")
                    font.pixelSize: 13
                    color: win.textColor
                    Layout.fillWidth: true
                }
                FluSwitch {
                    id: autoThemeSwitch
                    onCheckedChanged: if (settingsDialog.visible) win.applyThemeMode()
                }
            }

            RowLayout {
                spacing: 8
                Label {
                    text: qsTr("深色主题(不跟随系统时生效)")
                    font.pixelSize: 13
                    color: autoThemeSwitch.checked ? win.subText : win.textColor
                    Layout.fillWidth: true
                }
                FluSwitch {
                    id: themeSwitch
                    enabled: !autoThemeSwitch.checked
                    onCheckedChanged: if (settingsDialog.visible) win.applyThemeMode()
                }
            }

            RowLayout {
                spacing: 8
                Label {
                    text: qsTr("计划模式(先出计划再动手)")
                    font.pixelSize: 13
                    color: win.textColor
                    Layout.fillWidth: true
                }
                FluSwitch {
                    id: planSwitch
                    // 切换即生效,不必等点保存
                    onCheckedChanged: ctrl.setPlanMode(checked)
                }
            }

            Label {
                text: qsTr("开启后每次任务会先展示一份执行计划再开始改表,适合复杂操作;简单问答不受影响。")
                font.pixelSize: 11
                color: win.subText
                wrapMode: Text.Wrap
                Layout.fillWidth: true
            }

            // ---------- 一键安装 / 自启动(小白友好:全程界面操作,免命令行) ----------
            Rectangle {
                Layout.fillWidth: true
                Layout.topMargin: 2
                height: 1
                color: win.stroke
            }
            Label {
                id: installHeader
                text: qsTr("一键安装(Excel / WPS 工具栏按钮)")
                font.pixelSize: 13
                font.weight: Font.DemiBold
                color: win.textColor
                Layout.fillWidth: true
            }
            Label {
                text: qsTr("把「表答」装进 Excel 和 WPS 的功能区,并创建桌面快捷方式、登记开机自启;重复点击无副作用。")
                font.pixelSize: 11
                color: win.subText
                wrapMode: Text.Wrap
                Layout.fillWidth: true
            }
            RowLayout {
                spacing: 8
                Layout.fillWidth: true
                FluFilledButton {
                    text: ctrl.installBusy ? qsTr("正在安装…") : qsTr("一键安装 / 修复")
                    Layout.fillWidth: true
                    enabled: !ctrl.installBusy
                    onClicked: ctrl.installAddins()
                }
                FluSecondaryButton {
                    text: ctrl.installBusy ? qsTr("请稍候…") : qsTr("一键卸载")
                    Layout.fillWidth: true
                    enabled: !ctrl.installBusy
                    onClicked: ctrl.removeAddins()
                }
            }

            RowLayout {
                spacing: 8
                Label {
                    text: qsTr("开机自启动(打开表格自动就绪)")
                    font.pixelSize: 13
                    color: win.textColor
                    Layout.fillWidth: true
                }
                FluSwitch {
                    id: watchSwitch
                    // 与主题开关同款守卫:openNormal 回填时不触发写注册表
                    onCheckedChanged: if (settingsDialog.visible && !settingsDialog._watchGuard) ctrl.setWatchEnabled(checked)
                }
            }
            Label {
                text: qsTr("开启后不用每次手动启动:打开 Excel/WPS 本程序会自动就绪;关闭表格则自动退出。")
                font.pixelSize: 11
                color: win.subText
                wrapMode: Text.Wrap
                Layout.fillWidth: true
            }

            // ---------- 示例按钮(输入框上方)增删 ----------
            Rectangle {
                Layout.fillWidth: true
                Layout.topMargin: 2
                height: 1
                color: win.stroke
            }
            Label {
                text: qsTr("示例按钮(输入框上方,增删即保存)")
                font.pixelSize: 13
                font.weight: Font.DemiBold
                color: win.textColor
                Layout.fillWidth: true
            }
            Label {
                text: qsTr("点击按钮会把对应提示词发给 AI;空状态的「小白先点下面一句试试」也用同一份列表。")
                font.pixelSize: 11
                color: win.subText
                wrapMode: Text.Wrap
                Layout.fillWidth: true
            }
            ColumnLayout {
                spacing: 6
                Layout.fillWidth: true

                Repeater {
                    model: win.samples
                    RowLayout {
                        spacing: 8
                        Layout.fillWidth: true

                        FluIcon {
                            text: modelData.icon
                            font.pixelSize: 13
                            color: win.accent
                        }
                        ColumnLayout {
                            spacing: 1
                            Layout.fillWidth: true
                            Label {
                                text: modelData.text
                                font.pixelSize: 12
                                color: win.textColor
                                elide: Text.ElideRight
                                Layout.fillWidth: true
                                horizontalAlignment: Text.AlignLeft
                            }
                            Label {
                                text: modelData.prompt
                                font.pixelSize: 10
                                color: win.subText
                                elide: Text.ElideRight
                                Layout.fillWidth: true
                                horizontalAlignment: Text.AlignLeft
                            }
                        }
                        FluIconAction {
                            glyph: win.icoClear
                            glyphSize: 12
                            implicitWidth: 26
                            implicitHeight: 26
                            ToolTip.visible: hovered
                            ToolTip.text: qsTr("删除这个按钮")
                            onClicked: win.removeSample(index)
                        }
                    }
                }
                Label {
                    visible: win.samples.length === 0
                    text: qsTr("(已全部删除,输入框上方不再显示示例按钮)")
                    font.pixelSize: 11
                    color: win.subText
                    wrapMode: Text.Wrap
                    Layout.fillWidth: true
                }
                FluTextField {
                    id: newSampleText
                    Layout.fillWidth: true
                    placeholderText: qsTr("按钮文字,例如「按月份汇总」")
                    selectByMouse: true
                }
                FluTextField {
                    id: newSamplePrompt
                    Layout.fillWidth: true
                    placeholderText: qsTr("点击后发送给 AI 的提示词,例如「按月份汇总销售额并生成柱状图」")
                    selectByMouse: true
                }
                RowLayout {
                    Layout.fillWidth: true
                    spacing: 8
                    FluSecondaryButton {
                        text: qsTr("添加按钮")
                        onClicked: win.addSample()
                    }
                    Item { Layout.fillWidth: true }
                }
            }

            Label {
                id: testResult
                font.pixelSize: 12
                wrapMode: Text.WrapAnywhere
                Layout.fillWidth: true
                visible: text !== ""
            }

            RowLayout {
                spacing: 10
                Layout.alignment: Qt.AlignRight
                Layout.topMargin: 4

                FluSecondaryButton {
                    text: qsTr("测试连接")
                    onClicked: ctrl.testLlm(baseField.text, modelField.text, keyField.text)
                }
                FluFilledButton {
                    text: qsTr("保存")
                    onClicked: {
                        ctrl.saveSettings(presetKeys[providerBox.currentIndex] || "custom",
                                          baseField.text, modelField.text, keyField.text,
                                          stepsSpin.value, planSwitch.checked)
                        settingsDialog.close()
                        win.toast(qsTr("设置已保存"))
                    }
                }
            }
            } // settingsCol
        }     // ScrollView
    }

    // ---------------- Toast ----------------
    Rectangle {
        id: toast
        property bool shown: false
        anchors.bottom: parent.bottom
        anchors.bottomMargin: 110
        anchors.horizontalCenter: parent.horizontalCenter
        width: Math.min(parent.width - 60, toastLabel.implicitWidth + 32)
        height: toastLabel.implicitHeight + 18
        radius: 10
        color: win.dark ? "#2c313d" : "#1f2430"
        opacity: shown ? 0.96 : 0
        visible: opacity > 0.01
        Behavior on opacity { NumberAnimation { duration: 200 } }

        function show() {
            shown = true
            toastTimer.restart()
        }
        Timer {
            id: toastTimer
            interval: 2400
            onTriggered: toast.shown = false
        }
        Label {
            id: toastLabel
            anchors.centerIn: parent
            color: "white"
            font.pixelSize: 13
            wrapMode: Text.Wrap
            width: Math.min(implicitWidth, 320)
        }
    }
}
