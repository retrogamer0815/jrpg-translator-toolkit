#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn
#NoTrayIcon
; =; === Taskbar grouping: shared AppUserModelID ===
DllCall("shell32\SetCurrentProcessExplicitAppUserModelID", "wstr", "JRPGTranslator", "int")

; Big Box and other front ends can pass --background to initialize everything
; without showing or activating the control panel.
global CP_BACKGROUND_START := false
for __cpArg in A_Args {
    if (StrLower(__cpArg) = "--background") {
        CP_BACKGROUND_START := true
        break
    }
}

; Conditional single-instance handling: a background duplicate exits silently,
; while a normal launch brings the existing control panel forward.
global __CP_MUTEX := DllCall("kernel32\CreateMutexW", "ptr", 0, "int", 0
    , "wstr", "Local\JRPGTranslatorControlPanel", "ptr")
global __CP_ALREADY_RUNNING := (A_LastError = 183) ; ERROR_ALREADY_EXISTS
global CPPreviousForegroundHwnd := 0
global CPOverlayAdjustState := Map("active", false)
global CPOverlayAdjustHotkeysBound := false
global CP_OVERLAY_ADJUST_FLAG := A_Temp "\JRPG_Overlay\controller_adjust.active"

CloseControlPanelMutex(*) {
    global __CP_MUTEX
    if (__CP_MUTEX) {
        DllCall("kernel32\CloseHandle", "ptr", __CP_MUTEX)
        __CP_MUTEX := 0
    }
}

ShowWindowNoActivate(hwnd) {
    if !hwnd
        return false
    static SW_SHOWNOACTIVATE := 4
    static SWP_NOSIZE := 0x0001, SWP_NOMOVE := 0x0002
    static SWP_NOZORDER := 0x0004, SWP_NOACTIVATE := 0x0010
    static SWP_SHOWWINDOW := 0x0040
    DllCall("user32\ShowWindow", "ptr", hwnd, "int", SW_SHOWNOACTIVATE)
    DllCall("user32\SetWindowPos", "ptr", hwnd, "ptr", 0
        , "int", 0, "int", 0, "int", 0, "int", 0
        , "uint", SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW)
    return true
}

OnExit(CloseControlPanelMutex)

if (__CP_ALREADY_RUNNING) {
    if (!CP_BACKGROUND_START) {
        __cpOldDhw := A_DetectHiddenWindows
        __cpOldTitleMode := A_TitleMatchMode
        try {
            DetectHiddenWindows true
            SetTitleMatchMode 3
            __cpExistingHwnd := WinExist("JRPG Translator")
            if (__cpExistingHwnd) {
                DllCall("user32\ShowWindow", "ptr", __cpExistingHwnd, "int", 5) ; SW_SHOW
                try WinActivate("ahk_id " __cpExistingHwnd)
            }
        } finally {
            SetTitleMatchMode __cpOldTitleMode
            DetectHiddenWindows __cpOldDhw
        }
    }
    ExitApp
}

; Only the primary process owns the temporary controller-adjustment marker.
CPOverlayAdjustFlag(false)
OnExit(CPOverlayAdjustOnExit)

SafeCall(fn) {
    global CPOverlayAdjustState
    if (CPOverlayAdjustState.Has("active") && CPOverlayAdjustState["active"])
        return
    Try
        fn()
    Catch as ex
    {
        MsgBox("Control Panel error:`n" ex.Message "`n`n" ex.Extra)
    }
}

; --- Small helper: find 1-based index in an Array ---
ArrIndexOf(arr, needle) {
    for i, v in arr
        if (v = needle)
            return i
    return 0
}

; --- Explanations archive folder (global) ---
explainsDir := A_ScriptDir "\Settings\Explanations"
DirCreate(explainsDir)

; --- Feature flag for the new "Explanation Window" tab ---
global CP_ENABLE_EXPLAINER_DESIGN := true

; ===== Debug helpers =====
global __DBG_ENABLED_CP := (Trim(IniRead(A_ScriptDir "\Settings\control.ini", "cfg", "debugMode", 0)) != "0")
global __DBG_LOG := A_Temp "\JRPG_Control\debug.log"

SetDebugMode(enabled) {
    global __DBG_ENABLED_CP
    __DBG_ENABLED_CP := enabled ? true : false
    EnvSet("JRPG_DEBUG", __DBG_ENABLED_CP ? "1" : "0")
}

CPOnDebugModeToggle(*) {
    global cbDebug, debugMode
    debugMode := cbDebug.Value ? 1 : 0
    SetDebugMode(debugMode)
    MarkDirty()
}

SetDebugMode(__DBG_ENABLED_CP)

DbgCP(msg) {
    global __DBG_ENABLED_CP, __DBG_LOG
    if !__DBG_ENABLED_CP
        return
    Try {
        ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        dbgDir := A_Temp "\JRPG_Control"
        if !DirExist(dbgDir)
            DirCreate(dbgDir)
        FileAppend("[" ts "] CONTROL  " msg "`n", __DBG_LOG, "UTF-8")
    }
    Catch as exx
    {
        ; swallow logging issues
    }
}
GetWindowDPI(hwnd) {
    dpi := 96
    try dpi := DllCall("user32\GetDpiForWindow", "ptr", hwnd, "uint")
    return (dpi > 0) ? dpi : 96
}

; -------- Control Panel light/dark theme --------
global CPThemeBrushWindow := 0
global CPThemeBrushSurface := 0
global CPThemeBrushFocus := 0
global CPThemeMutedHwnds := Map()
global CPThemeColorSwatchHwnds := Map()
global CPColorFocusFrame := []
global CPControllerColorGradientSliders := Map()
global CPControllerColorGradientMessageRegistered := false
global CPThemeMessagesRegistered := false
global CPComboArrowOverlays := []

; -------- Scrollable control-panel canvas --------
; The existing layout remains a fixed minimum design surface. When the window is
; smaller, native scrollbars expose that surface instead of compressing controls.
global CP_CANVAS_MIN_W := 890
global CP_CANVAS_MIN_H := 680
global CP_VIEWPORT_MIN_W := 640
global CP_VIEWPORT_MIN_H := 480
global CPCanvasScrollX := 0
global CPCanvasScrollY := 0
global CPCanvasScrollMaxX := 0
global CPCanvasScrollMaxY := 0
global CPCanvasViewportW := CP_CANVAS_MIN_W
global CPCanvasViewportH := CP_CANVAS_MIN_H
global CPCanvasMessagesRegistered := false

CPRegisterCanvasMessages() {
    global ui, CPCanvasMessagesRegistered
    if CPCanvasMessagesRegistered
        return
    OnMessage(0x0114, CPOnCanvasScroll) ; WM_HSCROLL
    OnMessage(0x0115, CPOnCanvasScroll) ; WM_VSCROLL
    OnMessage(0x020A, CPOnCanvasMouseWheel) ; WM_MOUSEWHEEL
    ; Start with the non-client scrollbars hidden. SetScrollInfo will reveal
    ; either one only when the viewport is smaller than the design surface.
    try DllCall("user32\ShowScrollBar", "ptr", ui.Hwnd, "int", 3, "int", 0)
    CPCanvasMessagesRegistered := true
}

CPCanvasDirectControls() {
    global ui
    controls := []
    for childHwnd in CPGetControlHwnds() {
        if (DllCall("user32\GetParent", "ptr", childHwnd, "ptr") != ui.Hwnd)
            continue
        try {
            ctrl := GuiCtrlFromHwnd(childHwnd)
            if IsObject(ctrl)
                controls.Push(ctrl)
        }
    }
    return controls
}

CPCanvasMoveChildren(dx, dy, redraw := true, manageRedraw := true) {
    global ui
    if (!dx && !dy)
        return

    hwnd := ui.Hwnd
    if (hwnd && manageRedraw)
        DllCall("user32\SendMessage", "ptr", hwnd, "uint", 0x000B, "ptr", 0, "ptr", 0) ; WM_SETREDRAW
    try {
        for ctrl in CPCanvasDirectControls() {
            try {
                ctrl.GetPos(&x, &y)
                ctrl.Move(x + dx, y + dy)
            }
        }
    } finally {
        if (hwnd && manageRedraw) {
            DllCall("user32\SendMessage", "ptr", hwnd, "uint", 0x000B, "ptr", 1, "ptr", 0)
            if redraw
                DllCall("user32\RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100)
        }
    }
}

CPCanvasSetScrollInfo(bar, canvasSize, pageSize, pos) {
    global ui
    si := Buffer(28, 0)
    NumPut("uint", 28, si, 0)
    NumPut("uint", 0x0001 | 0x0002 | 0x0004, si, 4) ; SIF_RANGE|SIF_PAGE|SIF_POS
    NumPut("int", 0, si, 8)
    NumPut("int", Max(0, canvasSize - 1), si, 12)
    NumPut("uint", Max(0, pageSize), si, 16)
    NumPut("int", Max(0, pos), si, 20)
    return DllCall("user32\SetScrollInfo", "ptr", ui.Hwnd, "int", bar, "ptr", si.Ptr, "int", 1, "int")
}

CPCanvasGetScrollInfo(bar) {
    global ui
    si := Buffer(28, 0)
    NumPut("uint", 28, si, 0)
    NumPut("uint", 0x0017, si, 4) ; SIF_ALL
    if !DllCall("user32\GetScrollInfo", "ptr", ui.Hwnd, "int", bar, "ptr", si.Ptr, "int")
        return Map("page", 0, "pos", 0, "track", 0)
    return Map(
        "page", NumGet(si, 16, "uint"),
        "pos", NumGet(si, 20, "int"),
        "track", NumGet(si, 24, "int")
    )
}

CPCanvasScrollTo(newX, newY, redraw := true, manageRedraw := true) {
    global CPCanvasScrollX, CPCanvasScrollY, CPCanvasScrollMaxX, CPCanvasScrollMaxY
    global CPCanvasViewportW, CPCanvasViewportH, CP_CANVAS_MIN_W, CP_CANVAS_MIN_H

    newX := Max(0, Min(CPCanvasScrollMaxX, Round(newX)))
    newY := Max(0, Min(CPCanvasScrollMaxY, Round(newY)))
    dx := CPCanvasScrollX - newX
    dy := CPCanvasScrollY - newY
    CPCanvasSetScrollInfo(0, CP_CANVAS_MIN_W, CPCanvasViewportW, newX)
    CPCanvasSetScrollInfo(1, CP_CANVAS_MIN_H, CPCanvasViewportH, newY)
    CPCanvasMoveChildren(dx, dy, redraw, manageRedraw)
    CPCanvasScrollX := newX
    CPCanvasScrollY := newY
    return (dx || dy)
}

CPCanvasResetForLayout(manageRedraw := true) {
    global CPCanvasScrollX, CPCanvasScrollY
    oldX := CPCanvasScrollX
    oldY := CPCanvasScrollY
    if (oldX || oldY)
        CPCanvasMoveChildren(oldX, oldY, false, manageRedraw)
    CPCanvasScrollX := 0
    CPCanvasScrollY := 0
    return Map("x", oldX, "y", oldY)
}

CPCanvasFinishLayout(viewW, viewH, restorePos, manageRedraw := true) {
    global CPCanvasScrollMaxX, CPCanvasScrollMaxY, CPCanvasViewportW, CPCanvasViewportH
    global CP_CANVAS_MIN_W, CP_CANVAS_MIN_H

    CPCanvasViewportW := Max(1, viewW)
    CPCanvasViewportH := Max(1, viewH)
    CPCanvasScrollMaxX := Max(0, CP_CANVAS_MIN_W - CPCanvasViewportW)
    CPCanvasScrollMaxY := Max(0, CP_CANVAS_MIN_H - CPCanvasViewportH)
    CPCanvasSetScrollInfo(0, CP_CANVAS_MIN_W, CPCanvasViewportW, 0)
    CPCanvasSetScrollInfo(1, CP_CANVAS_MIN_H, CPCanvasViewportH, 0)
    CPCanvasScrollTo(restorePos["x"], restorePos["y"], false, manageRedraw)
}

CPOnCanvasScroll(wParam, lParam, msg, hwnd) {
    global ui, CPCanvasScrollX, CPCanvasScrollY, CPCanvasScrollMaxX, CPCanvasScrollMaxY
    global CPCanvasViewportW, CPCanvasViewportH
    if !(IsSet(ui) && ui && hwnd = ui.Hwnd) || lParam
        return

    bar := (msg = 0x0114) ? 0 : 1
    code := wParam & 0xFFFF
    info := CPCanvasGetScrollInfo(bar)
    current := (bar = 0) ? CPCanvasScrollX : CPCanvasScrollY
    maximum := (bar = 0) ? CPCanvasScrollMaxX : CPCanvasScrollMaxY
    page := (bar = 0) ? CPCanvasViewportW : CPCanvasViewportH
    lineStep := 40
    pageStep := Max(lineStep, page - lineStep)

    switch code {
        case 0: nextPos := current - lineStep       ; SB_LINEUP / SB_LINELEFT
        case 1: nextPos := current + lineStep       ; SB_LINEDOWN / SB_LINERIGHT
        case 2: nextPos := current - pageStep       ; SB_PAGEUP / SB_PAGELEFT
        case 3: nextPos := current + pageStep       ; SB_PAGEDOWN / SB_PAGERIGHT
        case 4, 5: nextPos := info["track"]         ; SB_THUMBPOSITION / SB_THUMBTRACK
        case 6: nextPos := 0                        ; SB_TOP / SB_LEFT
        case 7: nextPos := maximum                  ; SB_BOTTOM / SB_RIGHT
        default: return 0
    }

    if (bar = 0)
        CPCanvasScrollTo(nextPos, CPCanvasScrollY)
    else
        CPCanvasScrollTo(CPCanvasScrollX, nextPos)
    return 0
}

CPOnCanvasMouseWheel(wParam, lParam, msg, hwnd) {
    global ui, CPCanvasScrollX, CPCanvasScrollY, CPCanvasScrollMaxY
    if !(IsSet(ui) && ui && ui.Hwnd && CPCanvasScrollMaxY > 0)
        return
    if (hwnd != ui.Hwnd && !DllCall("user32\IsChild", "ptr", ui.Hwnd, "ptr", hwnd, "int"))
        return
    if !WinActive("ahk_id " ui.Hwnd)
        return

    focusedHwnd := CPFocusRingTargetHwnd(CPFocusedHwnd())
    if (focusedHwnd && CPComboDropped(focusedHwnd))
        return
    focusedClass := ""
    try focusedClass := WinGetClass("ahk_id " focusedHwnd)
    if (focusedClass = "msctls_trackbar32")
        return

    delta := (wParam >> 16) & 0xFFFF
    if (delta & 0x8000)
        delta -= 0x10000
    if !delta
        return
    CPCanvasScrollTo(CPCanvasScrollX, CPCanvasScrollY - (delta / 120) * 48)
    return 0
}

CPEnsureFocusedControlVisible(*) {
    global ui, CPCanvasScrollX, CPCanvasScrollY, CPCanvasScrollMaxX, CPCanvasScrollMaxY
    if !(IsSet(ui) && ui && ui.Hwnd && (CPCanvasScrollMaxX > 0 || CPCanvasScrollMaxY > 0))
        return
    focusHwnd := CPFocusRingTargetHwnd(CPFocusedHwnd())
    if !focusHwnd
        return

    ; The native tab is a full-page host used only as the custom tab bar's
    ; keyboard focus proxy. Revealing its oversized rectangle can nudge both
    ; scroll axes as pages change, so tab navigation always uses the canvas
    ; origin instead.
    if CPHwndIsTab(focusHwnd) {
        CPCanvasScrollTo(0, 0)
        return
    }

    ctrlRect := Buffer(16, 0)
    clientRect := Buffer(16, 0)
    clientOrigin := Buffer(8, 0)
    if !DllCall("user32\GetWindowRect", "ptr", focusHwnd, "ptr", ctrlRect.Ptr, "int")
        return
    if !DllCall("user32\GetClientRect", "ptr", ui.Hwnd, "ptr", clientRect.Ptr, "int")
        return
    DllCall("user32\ClientToScreen", "ptr", ui.Hwnd, "ptr", clientOrigin.Ptr, "int")

    margin := 12
    viewL := NumGet(clientOrigin, 0, "int") + margin
    viewT := NumGet(clientOrigin, 4, "int") + margin
    viewR := NumGet(clientOrigin, 0, "int") + NumGet(clientRect, 8, "int") - margin
    viewB := NumGet(clientOrigin, 4, "int") + NumGet(clientRect, 12, "int") - margin
    ctrlL := NumGet(ctrlRect, 0, "int")
    ctrlT := NumGet(ctrlRect, 4, "int")
    ctrlR := NumGet(ctrlRect, 8, "int")
    ctrlB := NumGet(ctrlRect, 12, "int")
    nextX := CPCanvasScrollX
    nextY := CPCanvasScrollY

    if (ctrlR - ctrlL > viewR - viewL)
        nextX += ctrlL - viewL
    else if (ctrlL < viewL)
        nextX -= viewL - ctrlL
    else if (ctrlR > viewR)
        nextX += ctrlR - viewR

    if (ctrlB - ctrlT > viewB - viewT)
        nextY += ctrlT - viewT
    else if (ctrlT < viewT)
        nextY -= viewT - ctrlT
    else if (ctrlB > viewB)
        nextY += ctrlB - viewB

    CPCanvasScrollTo(nextX, nextY)
}

CPPalette(darkMode := -1) {
    global controlDarkMode
    static light := Map(
        "window", "F0F0F0",
        "surface", "FFFFFF",
        "surfaceAlt", "F3F3F3",
        "focus", "F3F8FF",
        "text", "202020",
        "muted", "808080",
        "border", "A0A0A0",
        "accent", "0078D4",
        "accentFocus", "005A9E",
        "accentText", "FFFFFF"
    )
    static dark := Map(
        "window", "202124",
        "surface", "2B2D30",
        "surfaceAlt", "303236",
        "focus", "3A3D42",
        "text", "ECEDEF",
        "muted", "A8ABB2",
        "border", "4A4D52",
        "accent", "0A72C7",
        "accentFocus", "0F5F9E",
        "accentText", "FFFFFF"
    )
    if (darkMode = -1) {
        darkMode := 0
        try darkMode := controlDarkMode ? 1 : 0
    }
    return darkMode ? dark : light
}

CPColorRef(hexColor) {
    cpRgb := Integer("0x" Trim(hexColor, "#"))
    return ((cpRgb & 0xFF) << 16) | (cpRgb & 0x00FF00) | ((cpRgb >> 16) & 0xFF)
}

CPDestroyThemeBrushes(*) {
    global CPThemeBrushWindow, CPThemeBrushSurface, CPThemeBrushFocus
    if CPThemeBrushWindow
        try DllCall("gdi32\DeleteObject", "ptr", CPThemeBrushWindow)
    if CPThemeBrushSurface
        try DllCall("gdi32\DeleteObject", "ptr", CPThemeBrushSurface)
    if CPThemeBrushFocus
        try DllCall("gdi32\DeleteObject", "ptr", CPThemeBrushFocus)
    CPThemeBrushWindow := 0
    CPThemeBrushSurface := 0
    CPThemeBrushFocus := 0
}

CPRefreshThemeBrushes() {
    global CPThemeBrushWindow, CPThemeBrushSurface, CPThemeBrushFocus
    CPDestroyThemeBrushes()
    cpColors := CPPalette()
    CPThemeBrushWindow := DllCall("gdi32\CreateSolidBrush", "uint", CPColorRef(cpColors["window"]), "ptr")
    CPThemeBrushSurface := DllCall("gdi32\CreateSolidBrush", "uint", CPColorRef(cpColors["surface"]), "ptr")
    CPThemeBrushFocus := DllCall("gdi32\CreateSolidBrush", "uint", CPColorRef(cpColors["focus"]), "ptr")
}

CPRegisterMutedControl(cpMutedCtrl) {
    global CPThemeMutedHwnds
    if IsObject(cpMutedCtrl) && cpMutedCtrl.Hwnd
        CPThemeMutedHwnds[cpMutedCtrl.Hwnd] := true
    return cpMutedCtrl
}

CPIsMutedControl(cpMutedHwnd) {
    global CPThemeMutedHwnds
    return IsSet(CPThemeMutedHwnds) && IsObject(CPThemeMutedHwnds) && CPThemeMutedHwnds.Has(cpMutedHwnd)
}

CPRegisterColorSwatch(cpSwatchCtrl, cpSwatchTarget := "", cpFocusable := true) {
    global CPThemeColorSwatchHwnds
    if IsObject(cpSwatchCtrl) && cpSwatchCtrl.Hwnd {
        ; SS_NOTIFY + WS_TABSTOP lets a color preview participate in controller
        ; navigation while retaining its normal mouse-click behavior.
        if cpFocusable
            try cpSwatchCtrl.Opt("+0x100 +0x10000")
        CPThemeColorSwatchHwnds[cpSwatchCtrl.Hwnd] := cpSwatchTarget
    }
    return cpSwatchCtrl
}

CPUnregisterColorSwatch(cpSwatchCtrlOrHwnd) {
    global CPThemeColorSwatchHwnds
    cpSwatchHwnd := IsObject(cpSwatchCtrlOrHwnd) ? cpSwatchCtrlOrHwnd.Hwnd : cpSwatchCtrlOrHwnd
    if (cpSwatchHwnd && CPThemeColorSwatchHwnds.Has(cpSwatchHwnd))
        CPThemeColorSwatchHwnds.Delete(cpSwatchHwnd)
}

CPIsColorSwatchControl(cpSwatchHwnd) {
    global CPThemeColorSwatchHwnds
    return IsSet(CPThemeColorSwatchHwnds) && IsObject(CPThemeColorSwatchHwnds)
        && CPThemeColorSwatchHwnds.Has(cpSwatchHwnd)
}

CPColorSwatchTarget(cpSwatchHwnd) {
    global CPThemeColorSwatchHwnds
    if CPIsColorSwatchControl(cpSwatchHwnd)
        return CPThemeColorSwatchHwnds[cpSwatchHwnd]
    return ""
}

CPIsCustomTabControl(cpThemeHwnd) {
    global CPTabBarFill, CPTabButtons
    if (IsSet(CPTabBarFill) && CPTabBarFill && cpThemeHwnd = CPTabBarFill.Hwnd)
        return true
    if IsSet(CPTabButtons) && IsObject(CPTabButtons) {
        for cpThemeTabCtrl in CPTabButtons {
            if (cpThemeTabCtrl && cpThemeHwnd = cpThemeTabCtrl.Hwnd)
                return true
        }
    }
    return false
}

CPIsComboArrowControl(cpArrowHwnd) {
    global CPComboArrowOverlays
    if !IsSet(CPComboArrowOverlays) || !IsObject(CPComboArrowOverlays)
        return false
    for cpArrowEntry in CPComboArrowOverlays {
        if (cpArrowEntry["arrow"].Hwnd = cpArrowHwnd)
            return true
    }
    return false
}

CPComboArrowClick(cpComboHwnd, *) {
    if !cpComboHwnd || !DllCall("user32\IsWindowEnabled", "ptr", cpComboHwnd, "int")
        return
    try ControlFocus("ahk_id " cpComboHwnd)
    CPShowCombo(cpComboHwnd, !CPComboDropped(cpComboHwnd))
}

CPGetControlHwnds() {
    global ui
    cpControlsOldDetectHidden := A_DetectHiddenWindows
    try {
        DetectHiddenWindows true
        return WinGetControlsHwnd("ahk_id " ui.Hwnd)
    } finally {
        DetectHiddenWindows cpControlsOldDetectHidden
    }
}

CPCreateComboArrowOverlays() {
    global ui, CPComboArrowOverlays
    CPComboArrowOverlays := []
    for cpComboHwnd in CPGetControlHwnds() {
        cpComboClass := ""
        try cpComboClass := WinGetClass("ahk_id " cpComboHwnd)
        if (cpComboClass != "ComboBox")
            continue
        cpComboStyle := DllCall("user32\GetWindowLongPtr", "ptr", cpComboHwnd, "int", -16, "ptr")
        if ((cpComboStyle & 0x3) != 0x3) ; CBS_DROPDOWNLIST only
            continue
        cpComboCtrl := 0
        try cpComboCtrl := GuiCtrlFromHwnd(cpComboHwnd)
        if !IsObject(cpComboCtrl)
            continue
        cpArrowCtrl := ui.Add("Text", "x0 y0 w1 h1 Hidden Center +0x100 +0x200 +0x04000000", Chr(9662))
        cpArrowCtrl.Cursor := "Hand"
        cpArrowCtrl.OnEvent("Click", CPComboArrowClick.Bind(cpComboHwnd))
        CPComboArrowOverlays.Push(Map("combo", cpComboCtrl, "arrow", cpArrowCtrl))
    }
    CPUpdateComboArrowOverlays()
}

CPUpdateComboArrowOverlays(*) {
    global ui, controlDarkMode, CPComboArrowOverlays
    if !(IsSet(ui) && ui && ui.Hwnd && IsSet(CPComboArrowOverlays) && IsObject(CPComboArrowOverlays))
        return
    cpArrowColors := CPPalette(controlDarkMode)
    static SWP_KEEP_GEOMETRY := 0x0001 | 0x0002 | 0x0010

    for cpArrowEntry in CPComboArrowOverlays {
        cpArrowCombo := cpArrowEntry["combo"]
        cpArrowCtrl := cpArrowEntry["arrow"]
        cpArrowShow := controlDarkMode && DllCall("user32\IsWindowVisible", "ptr", cpArrowCombo.Hwnd, "int")
        if !cpArrowShow {
            cpArrowCtrl.Visible := false
            continue
        }

        cpArrowCombo.GetPos(&cpArrowX, &cpArrowY, &cpArrowW, &cpArrowH)
        cpArrowWidth := Min(26, Max(20, Floor(cpArrowH * 0.85)))
        cpArrowCtrl.Move(cpArrowX + cpArrowW - cpArrowWidth, cpArrowY + 1, cpArrowWidth - 1, Max(1, cpArrowH - 2))
        cpArrowEnabled := DllCall("user32\IsWindowEnabled", "ptr", cpArrowCombo.Hwnd, "int")
        cpArrowCtrl.Opt("+Background" cpArrowColors["surface"])
        cpArrowCtrl.SetFont("s9 c" (cpArrowEnabled ? cpArrowColors["text"] : cpArrowColors["muted"]))
        cpArrowCtrl.Visible := true
        try DllCall("user32\SetWindowPos", "ptr", cpArrowCtrl.Hwnd, "ptr", 0
            , "int", 0, "int", 0, "int", 0, "int", 0, "uint", SWP_KEEP_GEOMETRY)
        try cpArrowCtrl.Redraw()
    }
}

CPThemeComboParts(cpComboHwnd, darkMode) {
    cpComboInfoSize := (A_PtrSize = 8) ? 64 : 52
    cpComboInfo := Buffer(cpComboInfoSize, 0)
    NumPut("uint", cpComboInfoSize, cpComboInfo, 0)
    if !DllCall("user32\GetComboBoxInfo", "ptr", cpComboHwnd, "ptr", cpComboInfo.Ptr, "int")
        return

    cpItemOffset := (A_PtrSize = 8) ? 48 : 44
    cpListOffset := (A_PtrSize = 8) ? 56 : 48
    for cpPartHwnd in [NumGet(cpComboInfo, cpItemOffset, "ptr"), NumGet(cpComboInfo, cpListOffset, "ptr")] {
        if !cpPartHwnd
            continue
        if darkMode {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpPartHwnd, "wstr", "DarkMode_Explorer", "ptr", 0)
        } else {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpPartHwnd, "ptr", 0, "ptr", 0)
        }
        try DllCall("user32\InvalidateRect", "ptr", cpPartHwnd, "ptr", 0, "int", 1)
    }
}

CPApplyThemeToControl(cpThemeHwnd, darkMode := -1) {
    global controlDarkMode
    if !cpThemeHwnd || !DllCall("user32\IsWindow", "ptr", cpThemeHwnd, "int")
        return
    if (darkMode = -1) {
        darkMode := controlDarkMode ? 1 : 0
    }

    cpThemeClass := ""
    try cpThemeClass := WinGetClass("ahk_id " cpThemeHwnd)
    if (cpThemeClass = "")
        return

    cpThemeColors := CPPalette(darkMode)
    cpThemeCtrl := 0
    try cpThemeCtrl := GuiCtrlFromHwnd(cpThemeHwnd)

    if (cpThemeClass = "Static") {
        if CPIsCustomTabControl(cpThemeHwnd) || CPIsComboArrowControl(cpThemeHwnd)
            return
        cpThemeText := ""
        try cpThemeText := cpThemeCtrl.Text
        if (cpThemeText = "")
            return
        cpThemeTextColor := CPIsMutedControl(cpThemeHwnd) ? cpThemeColors["muted"] : cpThemeColors["text"]
        try cpThemeCtrl.SetFont("c" cpThemeTextColor)
        try cpThemeCtrl.Opt("+Background" cpThemeColors["window"])
    } else if (cpThemeClass = "Button") {
        try cpThemeCtrl.SetFont("c" cpThemeColors["text"])
        if darkMode {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "wstr", "DarkMode_Explorer", "ptr", 0)
        } else {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "ptr", 0, "ptr", 0)
        }
    } else if (cpThemeClass = "Edit" || cpThemeClass = "ComboBox" || cpThemeClass = "ListBox") {
        try cpThemeCtrl.SetFont("c" cpThemeColors["text"])
        try cpThemeCtrl.Opt("+Background" cpThemeColors["surface"])
        if darkMode {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "wstr", "DarkMode_CFD", "ptr", 0)
        } else {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "ptr", 0, "ptr", 0)
        }
        if (cpThemeClass = "ComboBox")
            CPThemeComboParts(cpThemeHwnd, darkMode)
    } else if (cpThemeClass = "msctls_trackbar32" || cpThemeClass = "msctls_updown32" || cpThemeClass = "SysTabControl32") {
        if darkMode {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "wstr", "DarkMode_Explorer", "ptr", 0)
        } else {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "ptr", 0, "ptr", 0)
        }
    }
    try DllCall("user32\InvalidateRect", "ptr", cpThemeHwnd, "ptr", 0, "int", 1)
}

CPApplyDarkTitleBar(cpThemeGuiHwnd, darkMode) {
    cpDarkValue := Buffer(4, 0)
    NumPut("int", darkMode ? 1 : 0, cpDarkValue, 0)
    try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", cpThemeGuiHwnd, "uint", 20, "ptr", cpDarkValue.Ptr, "uint", 4)
    try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", cpThemeGuiHwnd, "uint", 19, "ptr", cpDarkValue.Ptr, "uint", 4)
}

CPApplyWindowScrollbarTheme(cpThemeGuiHwnd, darkMode) {
    ; The canvas scrollbars belong to the top-level window rather than a child
    ; control, so they need their own Explorer theme assignment.
    if darkMode {
        try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeGuiHwnd, "wstr", "DarkMode_Explorer", "ptr", 0)
    } else {
        try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeGuiHwnd, "ptr", 0, "ptr", 0)
    }
}

CPSetPreferredAppDarkMode(darkMode, cpThemeGuiHwnd := 0) {
    ; Windows 10/11 expose per-app dark control rendering through uxtheme.
    ; Calls are guarded so older Windows versions simply fall back to our brushes.
    try DllCall("uxtheme\#135", "int", darkMode ? 2 : 3, "int") ; AllowDark / ForceLight
    if cpThemeGuiHwnd
        try DllCall("uxtheme\#133", "ptr", cpThemeGuiHwnd, "int", darkMode ? 1 : 0, "int")
    try DllCall("uxtheme\#136") ; FlushMenuThemes
}

CPThemeCtlColor(wParam, lParam, msg, parentHwnd) {
    global ui, controlDarkMode, CPThemeBrushWindow, CPThemeBrushSurface
    if !controlDarkMode || !(IsSet(ui) && ui && ui.Hwnd)
        return
    cpThemeOwner := DllCall("user32\GetWindow", "ptr", parentHwnd, "uint", 4, "ptr") ; GW_OWNER
    if (parentHwnd != ui.Hwnd
        && cpThemeOwner != ui.Hwnd
        && !DllCall("user32\IsChild", "ptr", ui.Hwnd, "ptr", lParam, "int"))
        return
    ; Color previews retain their configured color instead of inheriting the window theme.
    if CPIsColorSwatchControl(lParam)
        return
    if (msg = 0x0138 && (CPIsCustomTabControl(lParam) || CPIsComboArrowControl(lParam)))
        return

    cpCtlColors := CPPalette(true)
    cpCtlEnabled := DllCall("user32\IsWindowEnabled", "ptr", lParam, "int")
    cpCtlTextColor := cpCtlEnabled ? cpCtlColors["text"] : cpCtlColors["muted"]
    DllCall("gdi32\SetTextColor", "ptr", wParam, "uint", CPColorRef(cpCtlTextColor))

    if (msg = 0x0133 || msg = 0x0134) { ; Edit / ListBox
        DllCall("gdi32\SetBkColor", "ptr", wParam, "uint", CPColorRef(cpCtlColors["surface"]))
        return CPThemeBrushSurface
    }

    DllCall("gdi32\SetBkMode", "ptr", wParam, "int", 1) ; TRANSPARENT
    return CPThemeBrushWindow
}

CPThemeMeasureItem(wParam, lParam, msg, parentHwnd) {
    global ui
    if !lParam || NumGet(lParam, 0, "uint") != 3 ; ODT_COMBOBOX
        return
    cpMeasureDpi := (IsSet(ui) && ui && ui.Hwnd) ? GetWindowDPI(ui.Hwnd) : 96
    NumPut("uint", Max(20, Round(24 * cpMeasureDpi / 96)), lParam, 16)
    return true
}

CPThemeDrawItem(wParam, lParam, msg, parentHwnd) {
    global ui, controlDarkMode, CPThemeBrushSurface, CPThemeBrushFocus
    if !lParam || NumGet(lParam, 0, "uint") != 3 ; ODT_COMBOBOX
        return
    if !(IsSet(ui) && ui && ui.Hwnd)
        return

    cpDrawItemId := NumGet(lParam, 8, "uint")
    cpDrawState := NumGet(lParam, 16, "uint")
    cpDrawHwndOffset := (A_PtrSize = 8) ? 24 : 20
    cpDrawHdcOffset := (A_PtrSize = 8) ? 32 : 24
    cpDrawRectOffset := (A_PtrSize = 8) ? 40 : 28
    cpDrawComboHwnd := NumGet(lParam, cpDrawHwndOffset, "ptr")
    cpDrawHdc := NumGet(lParam, cpDrawHdcOffset, "ptr")
    if !cpDrawComboHwnd || !cpDrawHdc
        return

    if (cpDrawItemId = 0xFFFFFFFF)
        cpDrawItemId := SendMessage(0x0147, 0, 0, cpDrawComboHwnd) ; CB_GETCURSEL

    cpDrawText := ""
    if (cpDrawItemId >= 0) {
        cpDrawTextLen := SendMessage(0x0149, cpDrawItemId, 0, cpDrawComboHwnd) ; CB_GETLBTEXTLEN
        if (cpDrawTextLen >= 0) {
            cpDrawTextBuf := Buffer((cpDrawTextLen + 1) * 2, 0)
            SendMessage(0x0148, cpDrawItemId, cpDrawTextBuf.Ptr, cpDrawComboHwnd) ; CB_GETLBTEXT
            cpDrawText := StrGet(cpDrawTextBuf, "UTF-16")
        }
    }

    cpDrawSelected := (cpDrawState & 0x0001) != 0 ; ODS_SELECTED
    cpDrawDisabled := (cpDrawState & 0x0002) != 0 || (cpDrawState & 0x0004) != 0
    cpDrawBrush := cpDrawSelected ? CPThemeBrushFocus : CPThemeBrushSurface
    DllCall("user32\FillRect", "ptr", cpDrawHdc, "ptr", lParam + cpDrawRectOffset, "ptr", cpDrawBrush)

    cpDrawColors := CPPalette(controlDarkMode)
    if cpDrawDisabled
        cpDrawTextHex := cpDrawColors["muted"]
    else if cpDrawSelected && !controlDarkMode
        cpDrawTextHex := cpDrawColors["accentFocus"]
    else
        cpDrawTextHex := cpDrawColors["text"]
    DllCall("gdi32\SetTextColor", "ptr", cpDrawHdc, "uint", CPColorRef(cpDrawTextHex))
    DllCall("gdi32\SetBkMode", "ptr", cpDrawHdc, "int", 1)

    cpDrawFont := SendMessage(0x0031, 0, 0, cpDrawComboHwnd) ; WM_GETFONT
    cpDrawOldFont := cpDrawFont ? DllCall("gdi32\SelectObject", "ptr", cpDrawHdc, "ptr", cpDrawFont, "ptr") : 0
    cpDrawRect := Buffer(16, 0)
    cpDrawLeft := NumGet(lParam, cpDrawRectOffset, "int") + 8
    cpDrawTop := NumGet(lParam, cpDrawRectOffset + 4, "int")
    cpDrawRight := NumGet(lParam, cpDrawRectOffset + 8, "int") - 4
    cpDrawBottom := NumGet(lParam, cpDrawRectOffset + 12, "int")
    NumPut("int", cpDrawLeft, "int", cpDrawTop, "int", cpDrawRight, "int", cpDrawBottom, cpDrawRect, 0)
    DllCall("user32\DrawTextW", "ptr", cpDrawHdc, "wstr", cpDrawText, "int", -1, "ptr", cpDrawRect.Ptr
        , "uint", 0x0020 | 0x0004 | 0x0800 | 0x8000) ; SINGLELINE | VCENTER | NOPREFIX | END_ELLIPSIS
    if cpDrawOldFont
        DllCall("gdi32\SelectObject", "ptr", cpDrawHdc, "ptr", cpDrawOldFont, "ptr")
    if (cpDrawState & 0x0010)
        DllCall("user32\DrawFocusRect", "ptr", cpDrawHdc, "ptr", lParam + cpDrawRectOffset)
    return true
}

CPRegisterThemeMessages() {
    global CPThemeMessagesRegistered
    if CPThemeMessagesRegistered
        return
    for cpThemeMsg in [0x0133, 0x0134, 0x0135, 0x0138]
        OnMessage(cpThemeMsg, CPThemeCtlColor)
    OnMessage(0x002B, CPThemeDrawItem)
    OnMessage(0x002C, CPThemeMeasureItem)
    OnExit(CPDestroyThemeBrushes)
    CPThemeMessagesRegistered := true
}

CPApplyControlPanelTheme(forceRedraw := true) {
    global ui, controlDarkMode, CPTabBarFill, CPTabRenderedState, hkConflictText
    if !(IsSet(ui) && ui && ui.Hwnd)
        return

    cpApplyDark := controlDarkMode ? 1 : 0
    cpApplyColors := CPPalette(cpApplyDark)
    CPSetPreferredAppDarkMode(cpApplyDark, ui.Hwnd)
    CPApplyWindowScrollbarTheme(ui.Hwnd, cpApplyDark)
    CPRefreshThemeBrushes()
    ui.BackColor := cpApplyColors["window"]
    CPApplyDarkTitleBar(ui.Hwnd, cpApplyDark)

    try {
        for cpApplyHwnd in CPGetControlHwnds()
            CPApplyThemeToControl(cpApplyHwnd, cpApplyDark)
    }

    if (IsSet(CPTabBarFill) && CPTabBarFill)
        try CPTabBarFill.Opt("+Background" cpApplyColors["window"])
    if (IsSet(hkConflictText) && hkConflictText)
        try hkConflictText.SetFont("c" (cpApplyDark ? "FF7B72" : "FF0000"))

    CPTabRenderedState := ""
    CPRenderCustomTabBar(true)
    CPUpdateComboArrowOverlays()
    if forceRedraw
        try DllCall("user32\RedrawWindow", "ptr", ui.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400)
}

CPOnDarkModeToggle(*) {
    global chkDarkMode, controlDarkMode, iniPath
    controlDarkMode := chkDarkMode.Value ? 1 : 0
    IniWrite(controlDarkMode, iniPath, "cfg_control", "darkMode")
    CPApplyControlPanelTheme()
}

; =========================
; JRPG Translator Control Panel
; =========================

; --- Window & tray icon helper (stable) ---
SetGuiAndTrayIcon(guiObj, icoPath) {
    if !FileExist(icoPath)
        return false
    if (guiObj.Hwnd = 0)
        guiObj.Show("NA Hide")
    hwnd := guiObj.Hwnd
    if (hwnd = 0)
        return false
    try TraySetIcon(icoPath)
    IMAGE_ICON := 1, LR_LOADFROMFILE := 0x10, LR_DEFAULTSIZE := 0x40
    hBig := DllCall("LoadImage", "ptr", 0, "str", icoPath, "uint", IMAGE_ICON
                  , "int", 0, "int", 0, "uint", LR_LOADFROMFILE|LR_DEFAULTSIZE, "ptr")
    hSmall := DllCall("LoadImage", "ptr", 0, "str", icoPath, "uint", IMAGE_ICON
                    , "int", 16, "int", 16, "uint", LR_LOADFROMFILE, "ptr")
    If (hBig)
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x0080, "ptr", 1, "ptr", hBig)   ; WM_SETICON, ICON_BIG
    If (hSmall)
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x0080, "ptr", 0, "ptr", hSmall) ; WM_SETICON, ICON_SMALL
    OnExit((*) => (
        hBig   && DllCall("DestroyIcon","ptr",hBig),
        hSmall && DllCall("DestroyIcon","ptr",hSmall)
    ))
    return true
}

; Check if a window is topmost (WS_EX_TOPMOST 0x00000008)
IsWindowTopmost(winTitle) {
    return (WinGetExStyle(winTitle) & 0x00000008) != 0
}

; -------- layout constants (globals) --------
global pad := 12
global gap := 8

; -------- locations / INI (portable) --------
; Keep all settings next to the scripts in .\Settings
portableRoot := A_ScriptDir "\Settings"
if !DirExist(portableRoot)
    DirCreate(portableRoot)

; One-time migration from the old %APPDATA% location (if present)
oldRoot := A_AppData "\JRPG_Overlay"
try {
    if FileExist(oldRoot "\control.ini") && !FileExist(portableRoot "\control.ini")
        FileCopy(oldRoot "\control.ini", portableRoot "\control.ini", true)
    if DirExist(oldRoot "\profiles") && !DirExist(portableRoot "\profiles")
        DirCopy(oldRoot "\profiles", portableRoot "\profiles", true)
    if DirExist(oldRoot "\prompts") && !DirExist(portableRoot "\prompts")
        DirCopy(oldRoot "\prompts", portableRoot "\prompts", true)
}

; Now use the portable folder for everything
appDir      := portableRoot
iniPath     := appDir "\control.ini"
envPath     := appDir "\.env"
overlayDir  := A_Temp "\JRPG_Overlay"      ; runtime-only stuff stays in temp
profilesDir := appDir "\profiles"
; --- Explanation profiles subfolder ---
profilesDir_EW := profilesDir "\explainer"
try DirCreate(profilesDir_EW)

; Keeps the current registration so we can unbind/rebind on changes
global __HK_LAUNCH_EXPL_REQ := ""
Rebind_LaunchExplainerRequest() {
    global __HK_LAUNCH_EXPL_REQ, iniPath

    newHK := Trim(IniRead(iniPath, "hotkeys", "launch_explainer_request", ""))

    ; Unbind previous, if any
    if (__HK_LAUNCH_EXPL_REQ != "") {
        try Hotkey(__HK_LAUNCH_EXPL_REQ, "Off")
        __HK_LAUNCH_EXPL_REQ := ""
    }

    ; Bind fresh if configured
    if (newHK != "") {
        try {
            Hotkey(newHK, (*) => SafeCall(CP_LaunchExplainerRequest))
            __HK_LAUNCH_EXPL_REQ := newHK
        } catch as ex {
            ; If user typed an invalid AHK key string, donâ€™t crash the panel
            DbgCP("Failed to bind launch_explainer_request hotkey '" newHK "': " ex.Message)
        }
    }
}

; Keeps the current registration for explain-last so we can unbind/rebind on changes
global __HK_EXPLAIN_LAST := ""
global __HK_STARTSTOP_AUDIO := ""
global __HK_HIDE_SHOW_CP := ""
Rebind_ExplainLastTranslation() {
    global __HK_EXPLAIN_LAST, iniPath

    newHK := Trim(IniRead(iniPath, "hotkeys", "explain_last_translation", ""))

    ; Unbind previous, if any
    if (__HK_EXPLAIN_LAST != "") {
        try Hotkey(__HK_EXPLAIN_LAST, "Off")
        __HK_EXPLAIN_LAST := ""
    }

    ; Bind fresh if configured
    if (newHK != "") {
        try {
            ; Call the same function as the "Explain last jp. Text" buttonâ€”no window toggling.
            Hotkey(newHK, (*) => SafeCall(ExplainNow))
            __HK_EXPLAIN_LAST := newHK
        } catch as ex {
            ; If user typed an invalid AHK key string, donâ€™t crash the panel
            DbgCP("Failed to bind explain_last_translation hotkey '" newHK "': " ex.Message)
        }
    }
}

; NEW: bind/unbind the Start/Stop Audio toggle hotkey
Rebind_StartStopAudio() {
    global __HK_STARTSTOP_AUDIO, iniPath
    newHK := Trim(IniRead(iniPath, "hotkeys", "start_stop_audio", ""))

    ; Unbind previous, if any (guard for first run)
    if (IsSet(__HK_STARTSTOP_AUDIO) && __HK_STARTSTOP_AUDIO != "") {
        try Hotkey(__HK_STARTSTOP_AUDIO, "Off")
        __HK_STARTSTOP_AUDIO := ""
    }

    ; Bind fresh if configured
    if (newHK != "") {
        try {
            Hotkey(newHK, (*) => SafeCall(StartStopAudio))
            __HK_STARTSTOP_AUDIO := newHK
        } catch as ex {
            DbgCP("Failed to bind start_stop_audio hotkey '" newHK "': " ex.Message)
        }
    }
}

; Bind/unbind the Show/Hide Control Panel global hotkey
Rebind_HideShowControlPanel() {
    global __HK_HIDE_SHOW_CP, iniPath
    newHK := Trim(IniRead(iniPath, "hotkeys", "hide_show_control_panel", ""))

    ; Unbind previous, if any (guard for first run)
    if (IsSet(__HK_HIDE_SHOW_CP) && __HK_HIDE_SHOW_CP != "") {
        try Hotkey(__HK_HIDE_SHOW_CP, "Off")
        __HK_HIDE_SHOW_CP := ""
    }

    ; Bind fresh if configured
    if (newHK != "") {
        try {
            Hotkey(newHK, (*) => SafeCall(ToggleControlPanel))
            __HK_HIDE_SHOW_CP := newHK
        } catch as ex {
            DbgCP("Failed to bind hide_show_control_panel hotkey '" newHK "': " ex.Message)
        }
    }
}

ToggleControlPanel(*) {
    global ui, ddlProv

    try {
        if !ui.Hwnd
            return
        if DllCall("IsWindowVisible", "ptr", ui.Hwnd, "int") {
            SavePanelBounds()
            ui.Hide()
            SetTimer(RestoreControlPanelReturnWindow, -1)
        } else {
            CaptureControlPanelReturnWindow()
            ui.Show()
            try WinActivate("ahk_id " ui.Hwnd)
            try (IsSet(ddlProv) && ddlProv) ? ddlProv.Focus() : 0
        }
    } catch as ex {
        DbgCP("ToggleControlPanel failed: " ex.Message)
    }
}

HideControlPanel(*) {
    global ui
    try {
        if IsSet(ui) && ui && ui.Hwnd {
            SavePanelBounds()
            ui.Hide()
            SetTimer(RestoreControlPanelReturnWindow, -1)
        }
    }
}

CaptureControlPanelReturnWindow() {
    global ui, CPPreviousForegroundHwnd
    cpForegroundHwnd := DllCall("user32\GetForegroundWindow", "ptr")
    if (cpForegroundHwnd && (!IsSet(ui) || !ui || cpForegroundHwnd != ui.Hwnd))
        CPPreviousForegroundHwnd := cpForegroundHwnd
}

RestoreControlPanelReturnWindow(*) {
    global ui, CPPreviousForegroundHwnd
    cpReturnHwnd := CPPreviousForegroundHwnd
    CPPreviousForegroundHwnd := 0
    if (!cpReturnHwnd
     || (IsSet(ui) && ui && cpReturnHwnd = ui.Hwnd)
     || !DllCall("user32\IsWindow", "ptr", cpReturnHwnd, "int"))
        return
    if !DllCall("user32\SetForegroundWindow", "ptr", cpReturnHwnd, "int")
        try WinActivate("ahk_id " cpReturnHwnd)
}

CPFocusedControl() {
    global ui
    try {
        c := ui.FocusedCtrl
        return IsObject(c) ? c : 0
    }
    return 0
}

CPFocusedHwnd() {
    global ui
    hwnd := DllCall("user32\GetFocus", "ptr")
    if (hwnd && IsSet(ui) && ui && ui.Hwnd && DllCall("user32\IsChild", "ptr", ui.Hwnd, "ptr", hwnd, "int"))
        return hwnd
    ctrl := CPFocusedControl()
    try return ctrl && ctrl.Hwnd ? ctrl.Hwnd : 0
    return 0
}

CPFocusRingTargetHwnd(hwnd) {
    if !hwnd
        return 0
    try {
        parent := DllCall("user32\GetParent", "ptr", hwnd, "ptr")
        if (parent && InStr(WinGetClass("ahk_id " parent), "ComboBox"))
            return parent
    }
    return hwnd
}

CPControlFromHwnd(hwnd) {
    try return GuiCtrlFromHwnd(hwnd)
    return 0
}

CPEnsureColorFocusFrame() {
    global ui, CPColorFocusFrame, controlDarkMode
    if (IsSet(CPColorFocusFrame) && IsObject(CPColorFocusFrame) && CPColorFocusFrame.Length = 4)
        return

    CPColorFocusFrame := []
    cpFrameColor := controlDarkMode ? "FFFFFF" : CPPalette(0)["accentFocus"]
    Loop 4 {
        ; Register frame segments as non-focusable color surfaces so dark-mode
        ; WM_CTLCOLORSTATIC handling does not repaint them with the window brush.
        cpFramePart := CPRegisterColorSwatch(
            ui.Add("Text", "x0 y0 w1 h1 Hidden Disabled Background" cpFrameColor " +0x04000000")
          , "", false)
        CPColorFocusFrame.Push(cpFramePart)
    }
}

CPHideColorFocusFrame() {
    global CPColorFocusFrame
    if !IsSet(CPColorFocusFrame) || !IsObject(CPColorFocusFrame)
        return
    for cpFramePart in CPColorFocusFrame
        try cpFramePart.Visible := false
}

CPShowColorFocusFrame(swatchHwnd) {
    global CPColorFocusFrame, controlDarkMode
    if !CPIsColorSwatchControl(swatchHwnd)
        return CPHideColorFocusFrame()

    swatchCtrl := CPControlFromHwnd(swatchHwnd)
    if !IsObject(swatchCtrl)
        return CPHideColorFocusFrame()
    CPEnsureColorFocusFrame()

    swatchCtrl.GetPos(&swatchX, &swatchY, &swatchW, &swatchH)
    cpFrameColor := controlDarkMode ? "FFFFFF" : CPPalette(0)["accentFocus"]
    cpPad := 3
    cpThickness := 2
    cpOuterX := swatchX - cpPad
    cpOuterY := swatchY - cpPad
    cpOuterW := swatchW + cpPad * 2
    cpOuterH := swatchH + cpPad * 2
    cpFrameRects := [
        [cpOuterX, cpOuterY, cpOuterW, cpThickness],
        [cpOuterX, cpOuterY + cpOuterH - cpThickness, cpOuterW, cpThickness],
        [cpOuterX, cpOuterY, cpThickness, cpOuterH],
        [cpOuterX + cpOuterW - cpThickness, cpOuterY, cpThickness, cpOuterH]
    ]
    static SWP_KEEP_GEOMETRY := 0x0001 | 0x0002 | 0x0010
    for cpFrameIndex, cpFramePart in CPColorFocusFrame {
        cpFrameRect := cpFrameRects[cpFrameIndex]
        cpFramePart.Opt("+Background" cpFrameColor)
        cpFramePart.Move(cpFrameRect[1], cpFrameRect[2], cpFrameRect[3], cpFrameRect[4])
        cpFramePart.Visible := true
        try DllCall("user32\SetWindowPos", "ptr", cpFramePart.Hwnd, "ptr", 0
            , "int", 0, "int", 0, "int", 0, "int", 0, "uint", SWP_KEEP_GEOMETRY)
        try cpFramePart.Redraw()
    }
}

CPRestoreFocusVisual() {
    global CPFocusVisualCtrl, CPFocusVisualHwnd
    CPHideColorFocusFrame()
    if (IsSet(CPFocusVisualCtrl) && CPFocusVisualCtrl) {
        if CPIsColorSwatchControl(CPFocusVisualHwnd) {
            try CPFocusVisualCtrl.Redraw()
        } else {
            try CPFocusVisualCtrl.SetFont("Norm")
            try CPApplyThemeToControl(CPFocusVisualCtrl.Hwnd)
        }
    }
    CPFocusVisualCtrl := 0
    CPFocusVisualHwnd := 0
}

CPSetTabFocusIndicator(show := false) {
    global CPTabBarHasNavFocus
    cpTabNextState := show ? true : false
    if (!IsSet(CPTabBarHasNavFocus) || CPTabBarHasNavFocus != cpTabNextState) {
        CPTabBarHasNavFocus := cpTabNextState
        CPRenderCustomTabBar()
    }
}

CPRenderCustomTabBar(force := false) {
    global tab, tabNames, CPTabVisiblePages, CPTabButtons, CPTabBarFill, CPTabBarHasNavFocus, CPTabRenderedState, controlDarkMode
    if !(IsSet(tab) && tab && tab.Hwnd && IsSet(CPTabButtons) && IsObject(CPTabButtons))
        return

    cpTabActive := 1
    try cpTabActive := tab.Value
    if (cpTabActive < 1 || cpTabActive > tabNames.Length)
        cpTabActive := 1

    cpTabNavFocused := IsSet(CPTabBarHasNavFocus) && CPTabBarHasNavFocus
    cpTabColors := CPPalette(controlDarkMode)
    cpTabRenderState := cpTabActive "|" (cpTabNavFocused ? 1 : 0) "|" controlDarkMode
    if (!force && IsSet(CPTabRenderedState) && CPTabRenderedState = cpTabRenderState)
        return

    for cpTabButtonIndex, cpTabCtrl in CPTabButtons {
        cpTabPage := CPTabVisiblePages[cpTabButtonIndex]
        if (cpTabPage = cpTabActive) {
            if cpTabNavFocused {
                cpTabCtrl.Opt("+Background" cpTabColors["accentFocus"])
                cpTabCtrl.SetFont("Bold c" cpTabColors["accentText"])
            } else {
                cpTabCtrl.Opt("+Background" cpTabColors["accent"])
                cpTabCtrl.SetFont("Norm c" cpTabColors["accentText"])
            }
        } else {
            cpTabCtrl.Opt("+Background" cpTabColors["surfaceAlt"])
            cpTabCtrl.SetFont("Norm c" cpTabColors["text"])
        }
        try cpTabCtrl.Redraw()
    }
    try CPTabBarFill.Opt("+Background" cpTabColors["window"])
    try CPTabBarFill.Redraw()
    CPTabRenderedState := cpTabRenderState
    CPMaintainCustomTabZOrder()
}

CPMaintainCustomTabZOrder() {
    global tab, CPTabButtons, CPTabBarFill
    if !(IsSet(tab) && tab && tab.Hwnd && IsSet(CPTabButtons) && IsObject(CPTabButtons))
        return

    static SWP_KEEP_GEOMETRY := 0x0001 | 0x0002 | 0x0010 ; NOSIZE | NOMOVE | NOACTIVATE
    ; The native tab stays enabled and visible so AutoHotkey keeps every page
    ; interactive. Only its sibling-window order changes.
    try DllCall("user32\SetWindowPos", "ptr", tab.Hwnd, "ptr", 1
        , "int", 0, "int", 0, "int", 0, "int", 0, "uint", SWP_KEEP_GEOMETRY) ; HWND_BOTTOM
    try DllCall("user32\SetWindowPos", "ptr", CPTabBarFill.Hwnd, "ptr", 0
        , "int", 0, "int", 0, "int", 0, "int", 0, "uint", SWP_KEEP_GEOMETRY) ; HWND_TOP
    for cpTabCtrl in CPTabButtons
        try DllCall("user32\SetWindowPos", "ptr", cpTabCtrl.Hwnd, "ptr", 0
            , "int", 0, "int", 0, "int", 0, "int", 0, "uint", SWP_KEEP_GEOMETRY)
}

CPLayoutCustomTabBar(cpViewportW := 0) {
    global tab, CPTabButtons, CPTabNaturalWidths, CPTabBarFill
    if !(IsSet(tab) && tab && tab.Hwnd && IsSet(CPTabButtons) && IsObject(CPTabButtons))
        return

    tab.GetPos(&cpTabX, &cpTabY, &cpTabW)
    cpTabBarH := 30
    cpTabAvailableW := Max(1, cpTabW)
    if (cpViewportW > 0)
        cpTabAvailableW := Min(cpTabAvailableW, Max(1, cpViewportW - cpTabX * 2))

    cpTabDesiredTotal := 0
    for cpTabNaturalWidth in CPTabNaturalWidths
        cpTabDesiredTotal += cpTabNaturalWidth

    cpTabCount := CPTabButtons.Length
    cpTabMinW := 42
    cpTabCompact := cpTabDesiredTotal > cpTabAvailableW
    cpTabUseMinimums := cpTabAvailableW >= cpTabCount * cpTabMinW
    cpTabDesiredExtraTotal := 0
    if (cpTabCompact && cpTabUseMinimums) {
        for cpTabNaturalWidth in CPTabNaturalWidths
            cpTabDesiredExtraTotal += Max(0, cpTabNaturalWidth - cpTabMinW)
    }

    ; Keep the opaque cover across the full native tab header. Only the custom
    ; buttons shrink to the viewport; otherwise native tab labels and arrows
    ; become visible in the uncovered area as duplicate "ghost" controls.
    CPTabBarFill.Move(cpTabX, cpTabY, cpTabW, cpTabBarH)
    CPTabBarFill.Visible := true

    cpTabDrawX := cpTabX
    cpTabRemainingW := cpTabAvailableW
    for cpTabIndex, cpTabCtrl in CPTabButtons {
        ; Preserve the beginning of compact labels. SS_ENDELLIPSIS (0x4000)
        ; truncates only their right edge; full-width labels remain centered.
        try cpTabCtrl.Opt(cpTabCompact ? "-Center +0x4000" : "+Center -0x4000")

        if cpTabCompact {
            if (cpTabIndex = CPTabButtons.Length)
                cpTabItemW := cpTabRemainingW
            else if cpTabUseMinimums
                cpTabItemW := cpTabMinW + Floor(
                    Max(0, CPTabNaturalWidths[cpTabIndex] - cpTabMinW)
                    * (cpTabAvailableW - cpTabCount * cpTabMinW)
                    / Max(1, cpTabDesiredExtraTotal)
                )
            else
                cpTabItemW := Max(1, Floor(cpTabAvailableW / cpTabCount))
        } else {
            cpTabItemW := CPTabNaturalWidths[cpTabIndex]
        }

        cpTabItemW := Min(cpTabItemW, cpTabRemainingW)
        cpTabCtrl.Move(cpTabDrawX, cpTabY, cpTabItemW, cpTabBarH)
        cpTabCtrl.Visible := true
        cpTabDrawX += cpTabItemW
        cpTabRemainingW := Max(0, cpTabX + cpTabAvailableW - cpTabDrawX)
    }
    CPRenderCustomTabBar(true)
}

CPSelectCustomTab(cpTabIndex, *) {
    global tab, CPFocusVisualNavHwnd
    if !(IsSet(tab) && tab && tab.Hwnd)
        return

    CPFocusVisualNavHwnd := 0
    try tab.Value := cpTabIndex
    CPSetTabFocusIndicator(false)
    CPRenderCustomTabBar(true)
}

CPMouseTabClick(*) {
    global ui, CPTabVisiblePages, CPTabButtons
    if !(IsSet(ui) && ui && ui.Hwnd && IsSet(CPTabButtons) && IsObject(CPTabButtons))
        return

    ; Use the actual child window beneath the pointer. Coordinate comparisons are
    ; unsafe here because MouseGetPos and GetWindowRect can use different origins.
    MouseGetPos &cpMouseX, &cpMouseY, &cpMouseWindowHwnd, &cpMouseControlHwnd, 2
    if (cpMouseWindowHwnd != ui.Hwnd || !cpMouseControlHwnd)
        return

    for cpTabButtonIndex, cpTabCtrl in CPTabButtons {
        if (cpMouseControlHwnd = cpTabCtrl.Hwnd
            && DllCall("user32\IsWindowVisible", "ptr", cpTabCtrl.Hwnd, "int")) {
            CPSelectCustomTab(CPTabVisiblePages[cpTabButtonIndex])
            return
        }
    }
}

CPCreateCustomTabBar() {
    global ui, tabNames, CPTabVisiblePages, CPTabButtons, CPTabNaturalWidths, CPTabBarFill
    global CPTabBarHasNavFocus, CPTabRenderedState

    CPTabButtons := []
    CPTabNaturalWidths := []
    CPTabBarHasNavFocus := false
    CPTabRenderedState := ""

    ; This opaque strip covers the native tab header. The real tab control remains
    ; underneath only to manage page visibility and provide a keyboard focus target.
    try tab.Opt("+0x04000000") ; WS_CLIPSIBLINGS keeps native painting below the custom row.
    CPTabBarFill := ui.Add("Text", "x0 y0 w1 h1 Hidden BackgroundF0F0F0 +0x100 +0x04000000")
    for cpTabPage in CPTabVisiblePages {
        cpTabLabel := tabNames[cpTabPage]
        cpTabCtrl := ui.Add("Text", "x0 y0 h30 Hidden Center Border +0x100 +0x200 +0x04000000 BackgroundF3F3F3 c202020", cpTabLabel)
        cpTabCtrl.GetPos(,, &cpTabNaturalW)
        CPTabNaturalWidths.Push(Max(46, cpTabNaturalW + 18))
        cpTabCtrl.Cursor := "Hand"
        cpTabCtrl.OnEvent("Click", CPSelectCustomTab.Bind(cpTabPage))
        CPTabButtons.Push(cpTabCtrl)
    }
    CPLayoutCustomTabBar()
}

CPMarkFocusedTabFromFallback(*) {
    global CPFocusVisualNavHwnd
    hwnd := CPFocusRingTargetHwnd(CPFocusedHwnd())
    if (hwnd && CPHwndIsTab(hwnd)) {
        CPFocusVisualNavHwnd := hwnd
        CPSetTabFocusIndicator(true)
    }
    CPEnsureFocusedControlVisible()
}

UpdateCPFocusRing(*) {
    global ui, CPFocusVisualCtrl, CPFocusVisualHwnd, CPFocusVisualNavHwnd, controlDarkMode
    if !(IsSet(ui) && ui && ui.Hwnd && DllCall("user32\IsWindowVisible", "ptr", ui.Hwnd, "int")) {
        CPRestoreFocusVisual()
        CPSetTabFocusIndicator(false)
        return
    }

    hwnd := CPFocusRingTargetHwnd(CPFocusedHwnd())
    if !hwnd {
        CPRestoreFocusVisual()
        CPSetTabFocusIndicator(false)
        return
    }
    if (IsSet(CPFocusVisualHwnd) && CPFocusVisualHwnd = hwnd) {
        if CPIsColorSwatchControl(hwnd)
            CPShowColorFocusFrame(hwnd)
        return
    }

    CPRestoreFocusVisual()
    if CPHwndIsTab(hwnd) {
        CPSetTabFocusIndicator(IsSet(CPFocusVisualNavHwnd) && CPFocusVisualNavHwnd = hwnd)
        return
    }
    CPSetTabFocusIndicator(false)

    if (!IsSet(CPFocusVisualNavHwnd) || CPFocusVisualNavHwnd != hwnd) {
        CPFocusVisualNavHwnd := 0
        return
    }

    ctrl := CPControlFromHwnd(hwnd)
    if !IsObject(ctrl)
        return

    CPFocusVisualCtrl := ctrl
    CPFocusVisualHwnd := hwnd
    if CPIsColorSwatchControl(hwnd) {
        ; Keep the configured color visible and draw a high-contrast frame around it.
        CPShowColorFocusFrame(hwnd)
        return
    }
    cpFocusColors := CPPalette()
    try ctrl.Opt("+Background" cpFocusColors["focus"])
    try ctrl.SetFont("Bold c" (controlDarkMode ? cpFocusColors["accentText"] : cpFocusColors["accentFocus"]))
}

UpdateCPActiveTabHighlight(*) {
    global tab, CPTabVisiblePages
    static cpLastActiveTab := 0

    cpActiveTab := 0
    try cpActiveTab := tab.Value
    cpActiveTabVisible := false
    for cpVisiblePage in CPTabVisiblePages {
        if (cpVisiblePage = cpActiveTab) {
            cpActiveTabVisible := true
            break
        }
    }
    if (cpActiveTab && !cpActiveTabVisible && CPTabVisiblePages.Length) {
        cpActiveTab := CPTabVisiblePages[1]
        try tab.Value := cpActiveTab
    }
    if (cpActiveTab && cpLastActiveTab && cpActiveTab != cpLastActiveTab)
        CPCanvasScrollTo(0, 0)
    if cpActiveTab
        cpLastActiveTab := cpActiveTab

    CPRenderCustomTabBar()
    CPUpdateComboArrowOverlays()
}

CPHwndIsCombo(hwnd) {
    try return hwnd && InStr(WinGetClass("ahk_id " hwnd), "ComboBox") != 0
    return false
}

CPHwndIsTab(hwnd) {
    try return hwnd && InStr(WinGetClass("ahk_id " hwnd), "SysTabControl32") != 0
    return false
}

CPHwndIsButtonToggle(hwnd) {
    if !hwnd
        return false
    try {
        if (WinGetClass("ahk_id " hwnd) != "Button")
            return false
        buttonStyle := DllCall("user32\GetWindowLongPtr", "ptr", hwnd, "int", -16, "ptr")
        buttonType := buttonStyle & 0xF
        return (buttonType = 2 || buttonType = 3 || buttonType = 4 || buttonType = 5 || buttonType = 6 || buttonType = 9)
    }
    return false
}

CPComboDropped(hwnd) {
    static CB_GETDROPPEDSTATE := 0x0157
    if !CPHwndIsCombo(hwnd)
        return false
    try return SendMessage(CB_GETDROPPEDSTATE, 0, 0, hwnd) != 0
    return false
}

CPShowCombo(hwnd, show := true) {
    static CB_SHOWDROPDOWN := 0x014F
    if !CPHwndIsCombo(hwnd)
        return false
    try {
        SendMessage(CB_SHOWDROPDOWN, show ? 1 : 0, 0, hwnd)
        return true
    }
    return false
}

CPHwndIsFocusable(hwnd) {
    if !hwnd
        return false
    if !DllCall("user32\IsWindowVisible", "ptr", hwnd, "int")
        return false
    if !DllCall("user32\IsWindowEnabled", "ptr", hwnd, "int")
        return false

    className := ""
    try className := WinGetClass("ahk_id " hwnd)
    if (className = "" || (className = "Static" && !CPIsColorSwatchControl(hwnd)))
        return false

    style := DllCall("user32\GetWindowLongPtr", "ptr", hwnd, "int", -16, "ptr")
    static WS_TABSTOP := 0x00010000
    if !(style & WS_TABSTOP)
        return false

    ; Skip group boxes: they share the Button class but are not useful controller targets.
    if (className = "Button" && ((style & 0xF) = 0x7))
        return false

    rect := CPGetHwndRect(hwnd)
    return rect["w"] > 4 && rect["h"] > 4
}

CPEnumFocusableProc(hwnd, lParam) {
    global __CP_NAV_ITEMS
    if CPHwndIsFocusable(hwnd)
        __CP_NAV_ITEMS.Push(hwnd)
    return true
}

CPFocusableHwnds() {
    global ui, __CP_NAV_ITEMS
    __CP_NAV_ITEMS := []
    if !(IsSet(ui) && ui && ui.Hwnd)
        return __CP_NAV_ITEMS
    cb := CallbackCreate(CPEnumFocusableProc, "F")
    try DllCall("user32\EnumChildWindows", "ptr", ui.Hwnd, "ptr", cb, "ptr", 0)
    CallbackFree(cb)
    return __CP_NAV_ITEMS
}

CPGetHwndRect(hwnd) {
    r := Buffer(16, 0)
    DllCall("user32\GetWindowRect", "ptr", hwnd, "ptr", r)
    x1 := NumGet(r, 0, "int"), y1 := NumGet(r, 4, "int")
    x2 := NumGet(r, 8, "int"), y2 := NumGet(r, 12, "int")
    return Map("x", x1, "y", y1, "w", x2 - x1, "h", y2 - y1
        , "l", x1, "t", y1, "r", x2, "b", y2
        , "cx", x1 + (x2 - x1) / 2, "cy", y1 + (y2 - y1) / 2)
}

CPGetCurrentTabItemRect(guiClient := false) {
    global tab, ui
    if !(IsSet(tab) && tab && tab.Hwnd)
        return 0

    static TCM_GETCURSEL := 0x130B
    static TCM_GETITEMRECT := 0x130A
    tabIdx := DllCall("user32\SendMessageW", "ptr", tab.Hwnd, "uint", TCM_GETCURSEL, "ptr", 0, "ptr", 0, "ptr")
    if (tabIdx < 0)
        tabIdx := 0

    r := Buffer(16, 0)
    if !DllCall("user32\SendMessageW", "ptr", tab.Hwnd, "uint", TCM_GETITEMRECT, "ptr", tabIdx, "ptr", r.Ptr, "ptr")
        return CPGetHwndRect(tab.Hwnd)

    tabWin := Buffer(16, 0)
    DllCall("user32\GetWindowRect", "ptr", tab.Hwnd, "ptr", tabWin.Ptr)
    tabScreenX := NumGet(tabWin, 0, "int"), tabScreenY := NumGet(tabWin, 4, "int")

    x1 := tabScreenX + NumGet(r, 0, "int"), y1 := tabScreenY + NumGet(r, 4, "int")
    x2 := tabScreenX + NumGet(r, 8, "int"), y2 := tabScreenY + NumGet(r, 12, "int")
    if (guiClient && IsSet(ui) && ui && ui.Hwnd) {
        pts := Buffer(16, 0)
        NumPut("int", x1, "int", y1, "int", x2, "int", y2, pts, 0)
        DllCall("user32\MapWindowPoints", "ptr", 0, "ptr", ui.Hwnd, "ptr", pts.Ptr, "uint", 2)
        x1 := NumGet(pts, 0, "int"), y1 := NumGet(pts, 4, "int")
        x2 := NumGet(pts, 8, "int"), y2 := NumGet(pts, 12, "int")
    }
    return Map("x", x1, "y", y1, "w", x2 - x1, "h", y2 - y1
        , "l", x1, "t", y1, "r", x2, "b", y2
        , "cx", x1 + (x2 - x1) / 2, "cy", y1 + (y2 - y1) / 2)
}

CPFocusFirstControlInCurrentTab() {
    tabRect := CPGetCurrentTabItemRect()
    if !IsObject(tabRect)
        return false

    items := CPFocusableHwnds()
    bestHwnd := 0
    bestScore := 0

    for hwnd in items {
        if CPHwndIsTab(hwnd)
            continue
        rect := CPGetHwndRect(hwnd)
        if (rect["t"] <= tabRect["b"] + 8)
            continue

        score := (rect["t"] * 10000) + rect["l"]
        if (!bestHwnd || score < bestScore) {
            bestHwnd := hwnd
            bestScore := score
        }
    }

    if bestHwnd {
        CPSetFocusHwnd(bestHwnd)
        return true
    }
    return false
}

CPHwndIsTrackbar(hwnd) {
    if !hwnd
        return false
    try return WinGetClass("ahk_id " hwnd) = "msctls_trackbar32"
    return false
}

CPFocusTabBar() {
    global tab, CPFocusVisualNavHwnd
    if !(IsSet(tab) && tab && tab.Hwnd)
        return false

    CPFocusVisualNavHwnd := tab.Hwnd
    try tab.Focus()
    CPSetTabFocusIndicator(true)
    SetTimer(CPEnsureFocusedControlVisible, -1)
    return true
}

CPSetFocusHwnd(hwnd) {
    global ui, CPFocusVisualNavHwnd
    if !hwnd
        return
    CPFocusVisualNavHwnd := CPFocusRingTargetHwnd(hwnd)
    ; Make keyboard focus indicators visible even when focus is moved programmatically.
    try SendMessage(0x0127, 0x00030002, 0, ui.Hwnd) ; WM_CHANGEUISTATE, UIS_CLEAR, HIDEFOCUS|HIDEACCEL
    try {
        ctrl := GuiCtrlFromHwnd(hwnd)
        if IsObject(ctrl) {
            ctrl.Focus()
            SetTimer(CPEnsureFocusedControlVisible, -1)
            return
        }
    }
    try ControlFocus("ahk_id " hwnd, "ahk_id " ui.Hwnd)
    try DllCall("user32\SetFocus", "ptr", hwnd, "ptr")
    SetTimer(CPEnsureFocusedControlVisible, -1)
}

CPActionRowHwnds() {
    global btnOv, btnOvClose, btnAudio, btnExplainerLaunch, btnExplainerClose
    row := []
    for ctrl in [btnOv, btnOvClose, btnAudio, btnExplainerLaunch, btnExplainerClose] {
        try {
            if (ctrl && ctrl.Hwnd && CPHwndIsFocusable(ctrl.Hwnd))
                row.Push(ctrl.Hwnd)
        }
    }
    return row
}

CPActionRowIndex(hwnd, row := 0) {
    if !IsObject(row)
        row := CPActionRowHwnds()
    for idx, itemHwnd in row {
        if (itemHwnd = hwnd)
            return idx
    }
    return 0
}

CPNavActionRowHorizontal(hwnd, dir) {
    row := CPActionRowHwnds()
    idx := CPActionRowIndex(hwnd, row)
    if !idx
        return false

    if (dir = "Left")
        nextIdx := idx > 1 ? idx - 1 : row.Length
    else
        nextIdx := idx < row.Length ? idx + 1 : 1
    CPSetFocusHwnd(row[nextIdx])
    return true
}

CPFocusNearestAboveActionRow(curHwnd) {
    curRect := CPGetHwndRect(curHwnd)
    items := CPFocusableHwnds()
    bestHwnd := 0
    bestScore := 0

    for hwnd in items {
        if (hwnd = curHwnd || CPHwndIsTab(hwnd) || CPActionRowIndex(hwnd))
            continue
        rect := CPGetHwndRect(hwnd)
        if (rect["b"] > curRect["t"] + 4)
            continue

        dx := Abs(rect["cx"] - curRect["cx"])
        overlap := (rect["r"] > curRect["l"] && rect["l"] < curRect["r"])
        primary := curRect["t"] - rect["b"]
        score := (primary * 10000) + (overlap ? 0 : 2000) + dx
        if (!bestHwnd || score < bestScore) {
            bestHwnd := hwnd
            bestScore := score
        }
    }

    if bestHwnd {
        CPSetFocusHwnd(bestHwnd)
        return true
    }
    return false
}

CPForwardIfComboOpen(keyName) {
    hwnd := CPFocusedHwnd()
    if !(hwnd && CPComboDropped(hwnd))
        return false
    SendEvent("{" keyName "}")
    return true
}

CPNavMove(dir, *) {
    if CPForwardIfComboOpen(dir)
        return

    curHwnd := CPFocusedHwnd()
    if (curHwnd && CPHwndIsTrackbar(curHwnd) && (dir = "Left" || dir = "Right")) {
        ; The $-prefixed navigation hotkeys do not retrigger on this synthetic
        ; key, so the native slider receives it and fires its Change event.
        SendEvent("{" dir "}")
        return
    }
    if (curHwnd && CPActionRowIndex(curHwnd)) {
        if (dir = "Left" || dir = "Right") {
            CPNavActionRowHorizontal(curHwnd, dir)
            return
        }
        if (dir = "Up" && CPFocusNearestAboveActionRow(curHwnd))
            return
    }

    if (curHwnd && CPHwndIsTab(curHwnd) && (dir = "Left" || dir = "Right")) {
        CPNavSwitchTab(dir = "Left" ? -1 : 1)
        return
    }
    if (curHwnd && CPHwndIsTab(curHwnd) && dir = "Down") {
        if CPFocusFirstControlInCurrentTab()
            return
    }

    items := CPFocusableHwnds()
    if (items.Length = 0)
        return

    if !curHwnd {
        CPSetFocusHwnd(items[1])
        return
    }

    curRect := CPGetHwndRect(curHwnd)
    bestHwnd := 0
    bestScore := 0

    for hwnd in items {
        if (hwnd = curHwnd)
            continue
        if (curHwnd && !CPHwndIsTab(curHwnd) && CPHwndIsTab(hwnd))
            continue
        rect := CPGetHwndRect(hwnd)
        dx := rect["cx"] - curRect["cx"]
        dy := rect["cy"] - curRect["cy"]
        overlap := false

        if (dir = "Up") {
            if (rect["b"] > curRect["t"] + 4)
                continue
            primary := curRect["t"] - rect["b"]
            secondary := Abs(dx)
            overlap := (rect["r"] > curRect["l"] && rect["l"] < curRect["r"])
        } else if (dir = "Down") {
            if (rect["t"] < curRect["b"] - 4)
                continue
            primary := rect["t"] - curRect["b"]
            secondary := Abs(dx)
            overlap := (rect["r"] > curRect["l"] && rect["l"] < curRect["r"])
        } else if (dir = "Left") {
            if (rect["r"] > curRect["l"] + 4)
                continue
            primary := curRect["l"] - rect["r"]
            secondary := Abs(dy)
            overlap := (rect["b"] > curRect["t"] && rect["t"] < curRect["b"])
        } else {
            if (rect["l"] < curRect["r"] - 4)
                continue
            primary := rect["l"] - curRect["r"]
            secondary := Abs(dy)
            overlap := (rect["b"] > curRect["t"] && rect["t"] < curRect["b"])
        }

        ; Human-feeling spatial navigation: same visual row/column wins first,
        ; then nearest edge distance, then center alignment.
        score := (overlap ? 0 : 1000000000) + (primary * 10000) + secondary
        if (!bestHwnd || score < bestScore) {
            bestHwnd := hwnd
            bestScore := score
        }
    }

    if bestHwnd {
        CPSetFocusHwnd(bestHwnd)
        return
    }

    if (dir = "Up" && CPFocusTabBar())
        return

    ; If spatial navigation has no candidate in that direction, fall back to
    ; normal keyboard focus movement so the user is never stuck.
    if (dir = "Up" || dir = "Left")
        SendEvent("+{Tab}")
    else
        SendEvent("{Tab}")
    SetTimer(CPMarkFocusedTabFromFallback, -1)
}

CPNavUp(*) {
    CPNavMove("Up")
}

CPNavDown(*) {
    CPNavMove("Down")
}

CPNavLeft(*) {
    CPNavMove("Left")
}

CPNavRight(*) {
    CPNavMove("Right")
}

CPNavActivate(keyName := "Enter", *) {
    hwnd := CPFocusedHwnd()
    if (keyName = "Enter" && CPIsColorSwatchControl(hwnd)) {
        CPAdjustColorSwatchWithController(hwnd)
        return
    }
    if (hwnd && CPHwndIsCombo(hwnd) && !CPComboDropped(hwnd)) {
        CPShowCombo(hwnd, true)
        return
    }
    if (keyName = "Enter" && CPHwndIsButtonToggle(hwnd)) {
        SendMessage(0x00F5, 0, 0, hwnd) ; BM_CLICK
        return
    }
    SendEvent("{" keyName "}")
}

CPNavEnter(*) {
    CPNavActivate("Enter")
}

CPNavSpace(*) {
    CPNavActivate("Space")
}

CPNavSwitchTab(dir) {
    global tab, CPTabVisiblePages, CPFocusVisualNavHwnd
    if !(IsSet(tab) && tab && tab.Hwnd)
        return

    count := CPTabVisiblePages.Length
    if (count <= 0)
        return

    currentPage := 1
    try currentPage := tab.Value
    visibleIndex := 0
    for cpTabIndex, cpTabPage in CPTabVisiblePages {
        if (cpTabPage = currentPage) {
            visibleIndex := cpTabIndex
            break
        }
    }
    if !visibleIndex
        visibleIndex := 1

    nextVisibleIndex := Mod(visibleIndex - 1 + dir + count, count) + 1
    try tab.Value := CPTabVisiblePages[nextVisibleIndex]
    try tab.Focus()
    CPFocusVisualNavHwnd := tab.Hwnd
    CPSetTabFocusIndicator(true)
    CPRenderCustomTabBar(true)
    SetTimer(CPEnsureFocusedControlVisible, -1)
}


CPNavPrevTab(*) {
    CPNavSwitchTab(-1)
}

CPNavNextTab(*) {
    CPNavSwitchTab(1)
}

RegisterControlPanelArrowNavigation() {
    global ui, __CP_ARROW_NAV_BOUND
    if !(IsSet(ui) && ui && ui.Hwnd)
        return
    if (IsSet(__CP_ARROW_NAV_BOUND) && __CP_ARROW_NAV_BOUND)
        return

    HotIfWinActive("ahk_id " ui.Hwnd)
    try Hotkey("$Down", CPNavDown, "On")
    try Hotkey("$Up", CPNavUp, "On")
    try Hotkey("$Right", CPNavRight, "On")
    try Hotkey("$Left", CPNavLeft, "On")
    try Hotkey("$Enter", CPNavEnter, "On")
    try Hotkey("$NumpadEnter", CPNavEnter, "On")
    try Hotkey("$Space", CPNavSpace, "On")
    try Hotkey("$PgUp", CPNavPrevTab, "On")
    try Hotkey("$PgDn", CPNavNextTab, "On")
    try Hotkey("~LButton", CPMouseTabClick, "On")
    try Hotkey("$Esc", HideControlPanel, "On")
    HotIfWinActive()

    __CP_ARROW_NAV_BOUND := true
}

; --- Hotkeys registry (actions, labels, defaults) ---
; We keep hotkeys in [hotkeys] section of control.ini for now (global scope).
; Later we can add per-profile overrides if desired.
global hotkeyActions := [
    "screenshot_translate",
	"explain_last_translation",
	"hide_show_translator",
	"hide_show_explainer",
	"hide_show_control_panel",
	"take_screenshot",
	"screenshot_translation", 
	"launch_explainer_request",
	"recapture_region",
	"start_stop_audio"
]

global hotkeyLabels := Map(
    "screenshot_translate",       "Screenshot + Translate",
    "explain_last_translation",   "Explain last translation",
	"hide_show_translator",       "Show/Hide Translator",
    "hide_show_explainer",        "Show/Hide Explainer",
    "hide_show_control_panel",    "Show/Hide Control Panel",
    "take_screenshot",            "Take Screenshot",
    "screenshot_translation",     "Translate Screenshots",
	"launch_explainer_request",   "Launch Explainer + Req.",
	"recapture_region",           "Recapture Region",
	"start_stop_audio",           "Audio Translation On/Off"
)

global hotkeyDefaults := Map(
    "screenshot_translate",       "^+t",    ; Ctrl+Shift+T
    "explain_last_translation",   "^+e",    ; Ctrl+Shift+E
    "hide_show_translator",       "^+h",    ; Ctrl+Shift+H
    "hide_show_explainer",        "^+x",    ; Ctrl+Shift+X
    "hide_show_control_panel",    "^+c",    ; Ctrl+Shift+C
    "take_screenshot",            "^+s",    ; Ctrl+Shift+S
    "screenshot_translation",     "^+d",    ; Ctrl+Shift+D
    "launch_explainer_request",   "^+a",    ; Ctrl+Shift+A
    "recapture_region",           "^+r",    ; Ctrl+Shift+R
	"start_stop_audio",           "^+l"     ; Ctrl+Shift+L
)

; UI control maps for later wiring (Change/Disable/Default)
global hkEdits  := Map()  ; action -> Edit control (shows current binding)
global hkBtnChg := Map()  ; action -> "Changeâ€¦" button
global hkBtnDis := Map()  ; action -> "Disable" button
global hkBtnDef := Map()  ; action -> "Default" button
global hkDirty  := false


if !DirExist(profilesDir)
    DirCreate(profilesDir)
promptsDir  := appDir "\prompts"
if !DirExist(promptsDir)
    DirCreate(promptsDir)
	
; separate folder for EXPLANATION prompt profiles
explainPromptsDir := appDir "\prompts_explain"
if !DirExist(explainPromptsDir)
    DirCreate(explainPromptsDir)	
	
	; base folder for JPâ†’EN / ENâ†’EN glossaries (profileed)
glossariesDir := appDir "\glossaries"
if !DirExist(glossariesDir)
    DirCreate(glossariesDir)
; (Do NOT auto-create per-profile folders/files hereâ€”only when the user clicks New)

; -------- defaults --------
defPython       := ".\python\python.exe"
defAudioPy      := ".\scripts\live_audio_translator.py"
defOverlay      := A_IsCompiled ? ".\bin\overlay.exe" : ".\bin\overlay.ahk"
defImgPy        := ".\scripts\screenshot_translator.py"
defExplainPy    := ".\scripts\explainer.py"
defCaptureDir   := ".\Settings\Screenshots"
defOverlayTrans := 255

; overlay color defaults
defBoxBg  := "102040"
defBdrOut := defBoxBg
defBdrIn  := defBoxBg
defTxtCol := "FFFFFF"
defNameCol := "FFD166"  ; default speaker name color (soft amber)

; border width defaults
defBdrOutW := 3
defBdrInW  := 1

; overlay font defaults
defFontName := "Segoe UI"
defFontSize := 21

; -------------------- Hardcoded default schemes (SEED-ONLY) --------------------
; Edit these two maps to define the baseline look for a first run.
; They are ONLY applied to control.ini if the respective keys are missing.
defT := Map( ; Translator overlay -> section [cfg]
    "overlayTrans", 253
  , "boxBg",       0x000044
  , "bdrOut",      0x000044
  , "bdrIn",       0x000044
  , "txtColor",    0xFFFFFF
  , "bdrOutW",     0
  , "bdrInW",      0
  , "fontName",    defFontName
  , "fontSize",    defFontSize
  , "fontBold",    0
)

defE := Map( ; Explainer overlay -> section [cfg_explainer]
    "overlayTrans", 253
  , "boxBg",       0x000000
  , "bdrOut",      0x000000
  , "bdrIn",       0x000000
  , "txtColor",    0xFFFFFF
  , "bdrOutW",     0
  , "bdrInW",      0
  , "fontName",    defFontName
  , "fontSize",    defFontSize
  , "fontBold",    0
)

; Helper: write only if the key is missing (never overwrites user-changed values)
EnsureIniDefault(iniPath, section, key, default) {
    sentinel := "__MISSING__"
    val := IniRead(iniPath, section, key, sentinel)
    if (val = sentinel)
        IniWrite(default, iniPath, section, key)
}

; Seed control.ini with HARDcoded overlay defaults if keys are absent
SeedHardcodedOverlayDefaults() {
    global iniPath, defT, defE
    for k, v in defT
        EnsureIniDefault(iniPath, "cfg", k, v)
    for k, v in defE
        EnsureIniDefault(iniPath, "cfg_explainer", k, v)
}

; Run once at startup so first-run has visible/complete defaults in control.ini
SeedHardcodedOverlayDefaults()
; -----------------------------------------------------------------------------

; live audio translation models
defTrans := "gpt-realtime-translate"

; providers + model defaults for dropdown lists
defAudioProvider    := "openai"
defGeminiAudioModel := "gemini-3.5-live-translate-preview"
defImgProvider      := "openai"
defImgModel         := "gpt-4o"
defGeminiImgModel   := "gemini-2.5-flash"
; --- Explanation tab defaults (provider + text models)
defExplainProvider    := "openai"
defExplainOpenAIModel := "gpt-4o-mini"
defExplainGeminiModel := "gemini-2.5-flash"
; --- Advanced UI + debug defaults ---
defShowPathsTab := 0
defDebugMode := 0
; --- Glossary profile defaults ---
defJP2ENGlossaryProfile := "default"
defEN2ENGlossaryProfile := "default"

; NEW: default prompt profile name
defPromptProfile := "default_en"
; EXPLAIN: default prompt profile
defExplainPromptProfile := "default_en"
; (from previous build) screenshot post-processing
defImgPostproc := "tt"  ; "tt" | "translation" | "none"

; AUDIO live translation target language default
defAudioTargetLang := "English (en)"
audioTargetLangs := [
    "English (en)"
  , "German (de)"
  , "French (fr)"
  , "Spanish (es)"
  , "Italian (it)"
  , "Portuguese (pt)"
  , "Dutch (nl)"
  , "Polish (pl)"
  , "Russian (ru)"
  , "Ukrainian (uk)"
  , "Korean (ko)"
  , "Chinese Simplified (zh-CN)"
  , "Chinese Traditional (zh-TW)"
  , "Japanese (ja)"
]

; -------- state --------
global gPidAudio := 0
global gJustStoppedUntil := 0
global gLastAction := ""
global CPFocusVisualNavHwnd := 0

; -------- INI helpers --------
; Trim everything we read from the INI to avoid invisible whitespace / BOM residue issues.
Load(k, d, s := "cfg") => Trim(IniRead(iniPath, s, k, d))

SyncUnifiedWindowAppearance() {
    global boxBgHex, bdrOutHex, bdrInHex, bdrOutW, bdrInW
    global boxBgHex_EW, bdrOutHex_EW, bdrInHex_EW, bdrOutW_EW, bdrInW_EW

    bdrOutHex := boxBgHex
    bdrInHex := boxBgHex
    bdrOutW := 0
    bdrInW := 0

    bdrOutHex_EW := boxBgHex_EW
    bdrInHex_EW := boxBgHex_EW
    bdrOutW_EW := 0
    bdrInW_EW := 0
}

pythonExe       := Load("pythonExe",        defPython)
audioScript     := Load("audioScript",      defAudioPy)
if RegExMatch(StrLower(audioScript), "(^|\\|/|^\.\x5c)scripts[\\/]+audio_translator\.py$")
    audioScript := defAudioPy
overlayAhk      := Load("overlayAhk",       defOverlay)
if (A_IsCompiled) {
    ; Migrate release configurations from the former executable names while
    ; leaving genuinely custom overlay paths untouched.
    __overlayPathNormalized := StrLower(StrReplace(Trim(overlayAhk), "/", "\"))
    __legacyOverlayPaths := [
        ".\bin\jrpg_overlay.exe"
      , StrLower(A_ScriptDir "\bin\jrpg_overlay.exe")
    ]
    for __legacyOverlayPath in __legacyOverlayPaths {
        if (__overlayPathNormalized = __legacyOverlayPath) {
            overlayAhk := defOverlay
            IniWrite(overlayAhk, iniPath, "cfg", "overlayAhk")
            break
        }
    }
}
imgScript       := Load("imgScript",        defImgPy)
overlayTrans    := Load("overlayTrans",     defOverlayTrans)
explainScript   := Load("explainScript",   defExplainPy)
captureDir      := IniRead(iniPath, "paths", "captureDir", defCaptureDir)

; --- NEW: Native capture settings (safe defaults) ---
capMaxKB   := Integer(IniRead(iniPath, "capture", "maxKB", 1400))     ; cap file size in KB
capMode    := IniRead(iniPath, "capture", "mode", "region")           ; "region" or "window"
capWinInfo := IniRead(iniPath, "capture", "winTitle", "")             ; window title (fallback)
capRect    := IniRead(iniPath, "capture", "rect", "")                 ; "x,y,w,h" once selected

showPathsTab := Integer(Load("showPathsTab", defShowPathsTab, "cfg")) ? 1 : 0
debugMode := Integer(Load("debugMode", defDebugMode, "cfg"))
SetDebugMode(debugMode)
controlDarkMode := Integer(Load("darkMode", 0, "cfg_control")) ? 1 : 0
CPSetPreferredAppDarkMode(controlDarkMode)

; overlay colors
boxBgHex   := StrUpper(Load("boxBg",    defBoxBg))
bdrOutHex  := boxBgHex
bdrInHex   := boxBgHex
txtHex     := StrUpper(Load("txtColor", defTxtCol))
nameHex    := StrUpper(Load("nameColor", defNameCol))

; overlay border widths
bdrOutW := 0
bdrInW  := 0

; overlay font
fontName := Load("fontName", defFontName)
fontSize := Integer(Load("fontSize", defFontSize))
fontBold := Integer(Load("fontBold", 0)) ? 1 : 0

; === EXPLAINER overlay (separate state, section: cfg_explainer) ===
overlayTrans_EW := Load("overlayTrans",     defOverlayTrans, "cfg_explainer")

boxBgHex_EW  := StrUpper(Load("boxBg",      defBoxBg,       "cfg_explainer"))
bdrOutHex_EW := boxBgHex_EW
bdrInHex_EW  := boxBgHex_EW
txtHex_EW    := StrUpper(Load("txtColor",   defTxtCol,      "cfg_explainer"))

bdrOutW_EW := 0
bdrInW_EW  := 0

fontName_EW := Load("fontName",             defFontName,    "cfg_explainer")
fontSize_EW := Integer(Load("fontSize",     defFontSize,    "cfg_explainer"))
fontBold_EW := Integer(Load("fontBold",     0,              "cfg_explainer")) ? 1 : 0
SyncUnifiedWindowAppearance()

; --- EXPLAINER provider + model selections (own section)
explainProvider    := Load("explainProvider",    defExplainProvider,    "cfg_explainer")
explainOpenAIModel := Load("explainOpenAIModel", defExplainOpenAIModel, "cfg_explainer")
explainGeminiModel := Load("explainGeminiModel", defExplainGeminiModel, "cfg_explainer")

; safety: if INI was empty on first run
if (!explainProvider)    explainProvider    := defExplainProvider
if (!explainOpenAIModel) explainOpenAIModel := defExplainOpenAIModel
if (!explainGeminiModel) explainGeminiModel := defExplainGeminiModel

; --- Explainer bounds (persisted separately) ---
ewTmp := Load("x", "", "explainer_bounds"), ewX := (ewTmp = "" ? "" : Integer(ewTmp))
ewTmp := Load("y", "", "explainer_bounds"), ewY := (ewTmp = "" ? "" : Integer(ewTmp))
ewTmp := Load("w", "", "explainer_bounds"), ewW := (ewTmp = "" ? "" : Integer(ewTmp))
ewTmp := Load("h", "", "explainer_bounds"), ewH := (ewTmp = "" ? "" : Integer(ewTmp))

; track last-saved to avoid spam writes
ew_lastX := ewX, ew_lastY := ewY, ew_lastW := ewW, ew_lastH := ewH
ew_bounds_watch_running := false

trModel          := Load("trModel",          defTrans)
audioProvider    := Load("audioProvider",    defAudioProvider)
geminiAudioModel := Load("geminiAudioModel", defGeminiAudioModel)
audioProvider := (StrLower(audioProvider) = "gemini") ? "Gemini" : "OpenAI"
if !InStr(StrLower(trModel), "realtime")
    trModel := defTrans
if !InStr(StrLower(geminiAudioModel), "live-translate")
    geminiAudioModel := defGeminiAudioModel
imgProvider      := Load("imgProvider",      defImgProvider)
imgModel         := Load("imgModel",         defImgModel)
geminiImgModel   := Load("geminiImgModel",   defGeminiImgModel)
speakerName      := Load("speakerName", "")
; NEW: current prompt profile
promptProfile    := Load("promptProfile",    defPromptProfile)
; EXPLAIN: current prompt profile
explainPromptProfile := Load("explainPromptProfile", defExplainPromptProfile)

; (from previous build) post-processing mode
imgPostproc      := Load("imgPostproc",      defImgPostproc)
audioTargetLang := Load("audioTargetLanguage", defAudioTargetLang)
if !IndexOf(audioTargetLangs, audioTargetLang)
    audioTargetLang := defAudioTargetLang

; Load main window bounds
tmpW := Load("w", "", "gui_bounds")
tmpH := Load("h", "", "gui_bounds")
tmpX := Load("x", "", "gui_bounds")
tmpY := Load("y", "", "gui_bounds")
bounds_mode := Load("bounds_mode", "", "gui_bounds")  ; "client" once we've converted

guiW_saved := (tmpW != "" && IsNumber(tmpW)) ? Integer(tmpW) : ""
guiH_saved := (tmpH != "" && IsNumber(tmpH)) ? Integer(tmpH) : ""
guiX_saved := (tmpX != "" && IsNumber(tmpX)) ? Integer(tmpX) : ""
guiY_saved := (tmpY != "" && IsNumber(tmpY)) ? Integer(tmpY) : ""

; Glossary profile selections
jp2enGlossaryProfile := Load("jp2enGlossaryProfile", defJP2ENGlossaryProfile)
en2enGlossaryProfile := Load("en2enGlossaryProfile", defEN2ENGlossaryProfile)

; ---------- Model list persistence (new) ----------
; We store lists under [models] with comma-separated values.
StrJoin(arr, sep := ",") {
    out := ""
    for v in arr
        out .= (out = "" ? "" : sep) . v
    return out
}
IndexOf(arr, val) {
    for i, v in arr
        if (v = val)
            return i
    return 0
}
ModelListNaturalCompare(left, right) {
    ; Windows Explorer-style comparison keeps numeric model versions together,
    ; e.g. 5.4 before 5.5, while sorting suffixes alphabetically.
    return DllCall("Shlwapi\StrCmpLogicalW", "str", StrLower(left)
        , "str", StrLower(right), "int")
}
ModelListSort(arr) {
    if (!IsObject(arr) || arr.Length < 2)
        return arr

    ; Stable in-place insertion sort; model lists are small and this preserves
    ; the original spelling when two entries compare equally.
    Loop arr.Length - 1 {
        sourceIndex := A_Index + 1
        modelName := arr[sourceIndex]
        insertIndex := sourceIndex - 1
        while (insertIndex >= 1) {
            if (ModelListNaturalCompare(arr[insertIndex], modelName) <= 0)
                break
            arr[insertIndex + 1] := arr[insertIndex]
            insertIndex -= 1
        }
        arr[insertIndex + 1] := modelName
    }
    return arr
}
ModelListRead(key, defaultsArr) {
    raw := ""
    try raw := IniRead(iniPath, "models", key, "")
    if (Trim(raw) = "") {
        return ModelListSort(defaultsArr.Clone())
    }
    out := []
    for it in StrSplit(raw, ",") {
        s := Trim(it)
        if (s != "")
            out.Push(s)
    }
    return ModelListSort(out.Length ? out : defaultsArr.Clone())
}
ModelListWrite(key, arr) {
    ModelListSort(arr)
    IniWrite(StrJoin(arr, ","), iniPath, "models", key)
}
ModelListEnsure(arr, val, prepend := true) {
    if !val
        return
    if IndexOf(arr, val)
        return
    if prepend
        arr.InsertAt(1, val)
    else
        arr.Push(val)
}
ModelListMergeUnique(primary, secondary) {
    out := primary.Clone()
    for val in secondary
        ModelListEnsure(out, val, false)
    return ModelListSort(out)
}
AudioTargetCode(label) {
    if RegExMatch(label, "\(([A-Za-z-]+)\)\s*$", &m)
        return m[1]
    return "en"
}
AudioTargetName(label) {
    return Trim(RegExReplace(label, "\s*\([^)]+\)\s*$"))
}

; default lists
def_openai_img     := ["gpt-5.5","gpt-5.4-nano","gpt-5.4-pro","gpt-4o","gpt-4o-mini"]
def_gemini_img     := ["gemini-3.1-flash-lite","gemini-3.5-flash","gemini-3.1-pro-preview","gemini-2.5-flash","gemini-2.5-flash-lite","gemini-2.5-pro"]
def_openai_explain := def_openai_img.Clone()
def_gemini_explain := def_gemini_img.Clone()
def_openai_audio   := ["gpt-realtime-translate"]
def_gemini_audio   := ["gemini-3.5-live-translate-preview"]

; --- Explanation tab defaults (provider + text models)
defExplainProvider    := "openai"
defExplainOpenAIModel := "gpt-4o-mini"            ; uses text/chat models
defExplainGeminiModel := "gemini-2.5-flash"       ; Gemini text

; load lists from INI (or defaults)
model_openai_img     := ModelListRead("openai_img",     def_openai_img)
model_gemini_img     := ModelListRead("gemini_img",     def_gemini_img)
legacyOpenAIExplain  := ModelListRead("openai_tr",      def_openai_explain)
model_openai_explain := ModelListRead("openai_explain", legacyOpenAIExplain)
model_gemini_explain := ModelListRead("gemini_explain", model_gemini_img)
model_openai_audio   := ModelListRead("openai_audio",   def_openai_audio)
model_gemini_audio   := ModelListRead("gemini_audio",   def_gemini_audio)
ModelListEnsure(model_openai_explain, explainOpenAIModel)
ModelListEnsure(model_gemini_explain, explainGeminiModel)
ModelListSort(model_openai_explain)
ModelListSort(model_gemini_explain)
; Persist the one-time split immediately so later screenshot-list edits cannot
; become new explanation defaults on a subsequent launch.
ModelListWrite("openai_explain", model_openai_explain)
ModelListWrite("gemini_explain", model_gemini_explain)
ModelListEnsure(model_openai_audio, "gpt-realtime-translate")
ModelListEnsure(model_gemini_audio, "gemini-3.5-live-translate-preview")
ModelListSort(model_openai_audio)
ModelListSort(model_gemini_audio)

SaveAll(){
    global pythonExe,audioScript,overlayAhk,imgScript,overlayTrans,captureDir
    global trModel,audioProvider,geminiAudioModel,audioTargetLang
    global imgProvider,imgModel,geminiImgModel
    global iniPath, debugMode, showPathsTab
    global capMaxKB,capMode,capRect
    global boxBgHex,bdrOutHex,bdrInHex,txtHex
    global fontName,fontSize,fontBold
    global fontBold_EW
    global bdrOutW,bdrInW
    global model_openai_img, model_gemini_img, model_openai_explain, model_gemini_explain
    global model_openai_audio, model_gemini_audio
    global promptProfile, imgPostproc
 	global promptProfile, imgPostproc, chkDel, chkTop, chkDarkMode, controlDarkMode

    SyncUnifiedWindowAppearance()
    IniWrite(pythonExe,       iniPath, "cfg", "pythonExe")
	IniWrite(captureDir,      iniPath, "paths", "captureDir")

    ; --- Capture: refresh values from INI (picker writes them) before saving ---
    capModeLive := IniRead(iniPath, "capture", "mode", capMode)
    capRectLive := IniRead(iniPath, "capture", "rect", capRect)

    IniWrite(capMaxKB,        iniPath, "capture", "maxKB")
    IniWrite(capModeLive,     iniPath, "capture", "mode")
    IniWrite(capRectLive,     iniPath, "capture", "rect")

    ; keep UI variables in sync too
    capMode := capModeLive
    capRect := capRectLive

    IniWrite(audioScript,     iniPath, "cfg", "audioScript")
    IniWrite(overlayAhk,      iniPath, "cfg", "overlayAhk")
    IniWrite(imgScript,       iniPath, "cfg", "imgScript")
    IniWrite(overlayTrans,    iniPath, "cfg", "overlayTrans")
    IniWrite(explainScript,   iniPath, "cfg", "explainScript")
	IniWrite(chkTop.Value ? 1 : 0, iniPath, "cfg_control", "winTop")
    IniWrite(controlDarkMode ? 1 : 0, iniPath, "cfg_control", "darkMode")
    IniWrite(trModel,         iniPath, "cfg", "trModel")
    IniWrite(audioProvider,   iniPath, "cfg", "audioProvider")
    IniWrite(geminiAudioModel,iniPath, "cfg", "geminiAudioModel")
    IniWrite(audioTargetLang, iniPath, "cfg", "audioTargetLanguage")
    IniWrite(imgProvider,     iniPath, "cfg", "imgProvider")
    IniWrite(imgModel,        iniPath, "cfg", "imgModel")
    IniWrite(geminiImgModel,  iniPath, "cfg", "geminiImgModel")
    ; NEW: persist chosen prompt profile and postproc
    IniWrite(promptProfile,   iniPath, "cfg", "promptProfile")
    IniWrite(imgPostproc,     iniPath, "cfg", "imgPostproc")
    ; Back-compat: overlay reads "post"
    IniWrite(imgPostproc,     iniPath, "cfg", "post")
	IniWrite(showPathsTab, iniPath, "cfg", "showPathsTab")
	IniWrite(debugMode, iniPath, "cfg", "debugMode")
	; Also persist the delete-after-use toggle to [paths]
    IniWrite(chkDel.Value ? 1 : 0, iniPath, "paths", "deleteAfterUse")

    ; colors
    IniWrite(boxBgHex,        iniPath, "cfg", "boxBg")
    IniWrite(boxBgHex,        iniPath, "cfg", "bdrOut")
    IniWrite(boxBgHex,        iniPath, "cfg", "bdrIn")
    IniWrite(txtHex,          iniPath, "cfg", "txtColor")
	IniWrite(nameHex,         iniPath, "cfg", "nameColor")

    ; border widths
    IniWrite(0,               iniPath, "cfg", "bdrOutW")
    IniWrite(0,               iniPath, "cfg", "bdrInW")

    ; font
    IniWrite(fontName,        iniPath, "cfg", "fontName")
    IniWrite(fontSize,        iniPath, "cfg", "fontSize")
    IniWrite(fontBold,        iniPath, "cfg", "fontBold")
	
	    ; === EXPLAINER (separate section) ===
    IniWrite(overlayTrans_EW, iniPath, "cfg_explainer", "overlayTrans")
	IniWrite(explainProvider,    iniPath, "cfg_explainer", "explainProvider")
    IniWrite(explainOpenAIModel, iniPath, "cfg_explainer", "explainOpenAIModel")
    IniWrite(explainGeminiModel, iniPath, "cfg_explainer", "explainGeminiModel")


    ; colors
    IniWrite(boxBgHex_EW,     iniPath, "cfg_explainer", "boxBg")
    IniWrite(boxBgHex_EW,     iniPath, "cfg_explainer", "bdrOut")
    IniWrite(boxBgHex_EW,     iniPath, "cfg_explainer", "bdrIn")
    IniWrite(txtHex_EW,       iniPath, "cfg_explainer", "txtColor")

    ; border widths
    IniWrite(0,                iniPath, "cfg_explainer", "bdrOutW")
    IniWrite(0,                iniPath, "cfg_explainer", "bdrInW")

    ; font
    IniWrite(fontName_EW,     iniPath, "cfg_explainer", "fontName")
    IniWrite(fontSize_EW,     iniPath, "cfg_explainer", "fontSize")
    IniWrite(fontBold_EW,     iniPath, "cfg_explainer", "fontBold")
	
	    ; Explainer bounds
    if (ewX != "")
        IniWrite(ewX, iniPath, "explainer_bounds", "x")
    if (ewY != "")
        IniWrite(ewY, iniPath, "explainer_bounds", "y")
    if (ewW != "")
        IniWrite(ewW, iniPath, "explainer_bounds", "w")
    if (ewH != "")
        IniWrite(ewH, iniPath, "explainer_bounds", "h")

    ; persist lists
    ModelListWrite("openai_img",   model_openai_img)
    ModelListWrite("gemini_img",   model_gemini_img)
    ModelListWrite("openai_explain", model_openai_explain)
    ModelListWrite("gemini_explain", model_gemini_explain)
    ModelListWrite("openai_audio", model_openai_audio)
    ModelListWrite("gemini_audio", model_gemini_audio)
    DbgCP("SaveAll() persisted current config.")
}

SetCapMaxKB(v) {
    global capMaxKB
    try {
        capMaxKB := Integer(v)
    } catch as ex {   ; <-- renamed from "e" to avoid clashing with a global
        capMaxKB := 1400
    }
    ; optional bounds
    if (capMaxKB < 100)
        capMaxKB := 100
    else if (capMaxKB > 10000)
        capMaxKB := 10000
    SaveAll()
}

EnsureOverlayDir(){
    global overlayDir
    if !DirExist(overlayDir)
        DirCreate(overlayDir)
}

SignalExplainerBusy() {
    global overlayDir
    oldMode := A_TitleMatchMode
    SetTitleMatchMode 3
    hasExplainer := WinExist("Explainer")
    SetTitleMatchMode oldMode
    if !hasExplainer
        return

    try {
        EnsureOverlayDir()
        path := overlayDir "\cmd.explain_start"
        if FileExist(path)
            FileDelete(path)
        FileAppend("", path, "UTF-8")
    } catch as ex {
        DbgCP("SignalExplainerBusy failed: " ex.Message)
    }
}

ExpandEnv(str) {
    if !str
        return ""
    cap := 32767
    buf := Buffer(cap * 2, 0)
    DllCall("Kernel32\ExpandEnvironmentStringsW", "str", str, "ptr", buf, "int", cap, "int")
    return StrGet(buf, "UTF-16")
}

ResolvePath(p) {
    if !p
        return ""
    expanded := ExpandEnv(p)
    if RegExMatch(expanded, 'i)^(?:[A-Z]:\\|\\\\)')
        return expanded
    if RegExMatch(expanded, 'i)^(?:\./|\.\\|\.\./|\.\.\\)') {
        return A_ScriptDir "\" expanded
    }
    return A_ScriptDir "\" expanded
}

; =========================
; Audio state helper
; =========================
AudioIsRunning() {
    global gPidAudio
    if (gPidAudio && ProcessExist(gPidAudio))
        return true

    pids := AudioPidsByScript()
    if (pids.Length) {
        gPidAudio := pids[1]
        return true
    }

    gPidAudio := 0
    return false
}

UpdateStatus(){
    global btnAudio
    running := AudioIsRunning()
    if (IsSet(btnAudio) && IsObject(btnAudio))
        btnAudio.Text := running ? "Audio Translation On" : "Audio Translation Off"
}

_UpdateStatus(){
    UpdateStatus()
}

; === Dirty-flag & autosave helpers ===
MarkDirty(){
    global isDirty, bSave
    isDirty := true
    if IsSet(bSave) && IsObject(bSave){
        try bSave.Enabled := true
        try bSave.Text := "Save *"
    }
}
ClearDirty(){
    global isDirty, bSave
    isDirty := false
    if IsSet(bSave) && IsObject(bSave){
        try bSave.Enabled := false
        try bSave.Text := "Save"
    }
}
; Use this for changes that should immediately persist and NOT wake the Save button
AutoPersist(){
    UpdateVars()
    SaveAll()
    ClearDirty()
}

; -------- Browse handlers --------
BrowsePythonExe(*) {
    global pythonExe
    sel := FileSelect(3,, "Select python.exe", "Programs (*.exe)")
    if (sel != "")
        pythonExe := sel, SaveAll(), Repaint(), DbgCP("BrowsePythonExe -> " sel)
}
BrowseAudioScript(*) {
    global audioScript
    sel := FileSelect(3,, "Select audio subtitle .py", "Python (*.py)")
    if (sel != "")
        audioScript := sel, SaveAll(), Repaint(), DbgCP("BrowseAudioScript -> " sel)
}
BrowseOverlayAhk(*) {
    global overlayAhk
    sel := FileSelect(3,, "Select overlay .exe", "AutoHotkey (*.exe)")
    if (sel != "")
        overlayAhk := sel, SaveAll(), Repaint(), DbgCP("BrowseOverlayAhk -> " sel)
}
BrowseImageScript(*) {
    global imgScript
    sel := FileSelect(3,, "Select image translator .py", "Python (*.py)")
    if (sel != "")
        imgScript := sel, SaveAll(), Repaint(), DbgCP("BrowseImageScript -> " sel)
}

BrowseCaptureDir(*) {
    global captureDir
    sel := DirSelect(, , "Select screenshot folder")
    if (sel != "")
        captureDir := sel, SaveAll(), Repaint(), DbgCP("BrowseCaptureDir -> " sel)
}

BrowseExplainScript(*) {
    global explainScript
    sel := FileSelect(3,, "Select explainer .py", "Python (*.py)")
    if (sel != "")
        explainScript := sel, SaveAll(), Repaint(), DbgCP("BrowseExplainScript -> " sel)
}

; --- Screenshots: trigger ShareX â€œdefine capture regionâ€ (Ctrl+Alt+F2) ---
DefineCaptureRegion(*) {
    global iniPath
    hk := ""
    try hk := IniRead(iniPath, "cfg", "sharexDefineHotkey", "^!F2")
    if (Trim(hk) = "")
        hk := "^!F2"
    SendInput("{Ctrl up}{Alt up}{Shift up}")
    Sleep(20)
    if (hk = "^!F2")
        SendInput("^!{F2}")
    else
        Send(hk)
    ToolTip("ShareX: define capture region")
    SetTimer(() => ToolTip(""), -900)
    DbgCP("DefineCaptureRegion hotkey sent: " hk)
}

; --- Temporarily hide the Control Panel during capture, then auto-show when done ---
StartTempHideWatcher(kind := "region") {
    global ui, iniPath
    ; snapshot current values to detect a change (no ErrorLevel use in v2)
    oldMode := IniRead(iniPath, "capture", "mode", "")
    oldRect := IniRead(iniPath, "capture", "rect", "")
    oldTit  := IniRead(iniPath, "capture", "winTitle", "")

    ; store snapshots in globals for the poller
    global __HideWatchKind := kind
    global __OldMode := oldMode
    global __OldRect := oldRect
    global __OldTit  := oldTit
	; also snapshot INI modified time so "same rect again" still counts as done
    oldMTime := ""
    try oldMTime := FileGetTime(iniPath, "M")
    global __OldIniMTime := oldMTime
    global __CapWatchActive := true


    ; hide the Control Panel now
    try ui.Hide()

    ; start polling the INI every 150ms for a completed selection
    SetTimer(WatchCapDone, 150)

    ; fail-safe: force show again after 15s
    SetTimer(CapWatchFallback, -15000)
}

FinishCapWatch() {
    global ui, __CapWatchActive
    __CapWatchActive := false
    SetTimer(WatchCapDone, 0)
    SetTimer(CapWatchFallback, 0)
    try ui.Show()
}

CapWatchFallback(*) {
    global ui, __CapWatchActive
    if !__CapWatchActive
        return
    __CapWatchActive := false
    SetTimer(WatchCapDone, 0)
    try ui.Show()
}

WatchCapDone(*) {
    global ui, iniPath, __HideWatchKind, __OldMode, __OldRect, __OldTit, __OldIniMTime

    curMode := IniRead(iniPath, "capture", "mode", "")
	    curMTime := ""
    try curMTime := FileGetTime(iniPath, "M")
    if (__HideWatchKind = "region") {
        curRect := IniRead(iniPath, "capture", "rect", "")
        ; re-show when a new rect is written under mode=region
            if (curMode = "region" && curRect != "" && (curRect != __OldRect || curMTime != __OldIniMTime)) {
            FinishCapWatch()
        }
    } else if (__HideWatchKind = "window") {
        curTit := IniRead(iniPath, "capture", "winTitle", "")
        ; re-show when a new window title is written under mode=window
            if (curMode = "window" && curTit != "" && (curTit != __OldTit || curMTime != __OldIniMTime)) {
            FinishCapWatch()
        }
    }
}

; --- NEW: tiny modal picker to choose Region vs Window (native) ---
; --- Native capture mode picker (clamped to virtual desktop, near mouse) ---
OpenCapturePicker(*) {
    global capMaxKB, capMode

    CoordMode "Mouse", "Screen"
    MouseGetPos &mx, &my

    ; virtual desktop across all monitors
    vsx := SysGet(76)  ; SM_XVIRTUALSCREEN
    vsy := SysGet(77)  ; SM_YVIRTUALSCREEN
    vsw := SysGet(78)  ; SM_CXVIRTUALSCREEN
    vsh := SysGet(79)  ; SM_CYVIRTUALSCREEN

    g := Gui("+ToolWindow -Caption +AlwaysOnTop -DPIScale")
    g.BackColor := "101825"
    g.MarginX := 12, g.MarginY := 12
    g.SetFont("s10", "Segoe UI")

        g.Add("Text", "cWhite", "Select capture mode")
    g.Add("Text", "w0 h0") ; spacer

    b1 := g.Add("Button", "xm y+8 w120", "Region")
    b2 := g.Add("Button", "x+m w140",      "Window")   ; widen to prevent wrap
    bCancel := g.Add("Button", "xm y+8 w272", "Cancel") ; 120 + 12 (margin) + 140

    b1.OnEvent("Click", (*) => (g.Hide(), _SendCapPick("region")))
    b2.OnEvent("Click", (*) => (g.Hide(), _SendCapPick("window")))
    bCancel.OnEvent("Click", (*) => g.Destroy())

    g.Show("AutoSize NoActivate x" vsx+vsw " y" vsy+vsh)
    g.GetPos(, &dlgW, &dlgH)
    x := mx - Floor(dlgW/2)
    y := my + 20
    x := Max(vsx, Min(x, vsx + vsw - dlgW))
    y := Max(vsy, Min(y, vsy + vsh - dlgH))
    g.Move(x, y)
    g.Show()


    ; clean up if left open
    SetTimer(() => g.Destroy(), -120000)
}

TranslatorWindowExists() {
    oldMode := A_TitleMatchMode
    oldDetectHidden := A_DetectHiddenWindows
    SetTitleMatchMode(3)
    DetectHiddenWindows(false)
    hwnd := WinExist("Translator")
    DetectHiddenWindows(oldDetectHidden)
    SetTitleMatchMode(oldMode)
    return hwnd
}

EnsureTranslatorForCapture() {
    if TranslatorWindowExists()
        return true

    Toast("Opening Translator...")
    LaunchOverlay()
    if TranslatorWindowExists()
        return true

    deadline := A_TickCount + 1500
    while (A_TickCount < deadline) {
        Sleep(50)
        if TranslatorWindowExists()
            return true
    }
    return false
}

_SendCapPick(kind) {
    global capMaxKB
    ; compose a simple, future-proof key=value payload
    payload := "capcmd=pick"
            .  "|kind=" kind
            .  "|maxkb=" capMaxKB

    if !EnsureTranslatorForCapture() {
        Toast("Could not open Translator")
        return
    }

    StartTempHideWatcher(kind)
    ok := SendOverlayCmd(payload)
    if !ok {
        FinishCapWatch()
        Toast("Translator did not become ready")
        return
    }
    Toast("Pick " (kind = "region" ? "region" : "window"))
}

; Push current Screenshot Translation selections to environment so the next run uses them
; Push current Screenshot Translation selections to environment so the next run uses them
ApplyShotSettings(*) {
    global ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt, ddlPost
    global postCodes  ; mapping we added earlier
    global chkGuess   ; <â€” new: UI toggle for highlighting

    ; Read current selections
    provider := ddlProv.Text
    ; pick model depending on provider
    if (provider = "gemini") {
        modelToSet := ddlIMG_GM.Text
        if (SubStr(modelToSet, 1, 7) != "models/")
            modelToSet := "models/" . modelToSet
        EnvSet("GEMINI_MODEL_NAME", modelToSet)
    } else {
        EnvSet("MODEL_NAME", ddlIMG.Text)
    }

    ; Provider + prompt + postproc
    EnvSet("PROVIDER", provider)
    EnvSet("PROMPT_PROFILE", ddlPrompt.Text)
    EnvSet("PROMPT_FILE", "")
    EnvSet("POSTPROC_MODE", postCodes[ddlPost.Value])

    ; --- Highlight guessed subjects ---
    EnvSet("SHOT_ITALICIZE_GUESSED", chkGuess.Value ? "1" : "0")
    EnvSet("SHOT_GUESS_DELIM", Chr(0x60))  ; literal backtick

    ; --- Speaker name color toggle (JP+EN; Python strips ã€Œâ€¦ã€ when ON) ---
    global chkName
    EnvSet("SHOT_COLOR_SPEAKER", chkName.Value ? "1" : "0")
    }

ExplainNow(*) {
    global pythonExe, explainScript
    global explainProvider, explainOpenAIModel, explainGeminiModel
    global debugMode, explainsDir
    px := ResolvePath(pythonExe)
    ex := ResolvePath(explainScript)
    if !(FileExist(px) && FileExist(ex)) {
        MsgBox("Set valid paths for python.exe and explainer script first.`n`npythonExe:`n" px "`n`nexplainer:`n" ex, "Missing", 48)
        return
    }
	
	; Use the selected EXPLAIN profile
    p := ExplainProfilePath(Trim(ddlEPr.Text))
    if FileExist(p)
        EnvSet "EXPLAIN_PROMPT_FILE", p
    else
        EnvSet "EXPLAIN_PROMPT_FILE", ""  ; Python falls back to BASE_PROMPT


        ; NEW: propagate debug toggle to the Python explainer process
        EnvSet "JRPG_DEBUG", (debugMode ? "1" : "0")

    ; Provider/model for explainer: use Explanation tab
    prov := StrLower(Trim(explainProvider))
    EnvSet("EXPLAIN_PROVIDER", prov)
    if (prov = "gemini") {
        modelToSet := explainGeminiModel
        if (SubStr(modelToSet, 1, 7) != "models/")
            modelToSet := "models/" . modelToSet
        EnvSet("GEMINI_EXPLAIN_MODEL", modelToSet)
        EnvSet("EXPLAIN_MODEL","")
    } else {
        EnvSet("EXPLAIN_MODEL", explainOpenAIModel)
        EnvSet("GEMINI_EXPLAIN_MODEL","")
    }
    EnvSet("PYTHONIOENCODING","utf-8")
	
	; --- Save-to-textfiles (archive) wiring for explainer.py ---
    ; Read the user's toggle from the INI (written by the checkbox in the Explanation tab)
    saveExpl := Integer(IniRead(iniPath, "cfg", "saveExplains", 0))

    ; Pass environment variables to explainer.py
    ; SAVE_EXPLAINS: "1" to archive each explanation; "0" to skip (default)
    ; EXPLAIN_SAVE_DIR: directory where time-stamped files are written
    ; SETTINGS_DIR: optional hint for Python's fallback resolution
    EnvSet "SAVE_EXPLAINS", (saveExpl ? "1" : "0")
    EnvSet "EXPLAIN_SAVE_DIR", explainsDir
    EnvSet "SETTINGS_DIR", A_ScriptDir "\Settings"
    
    outFile := A_Temp "\learn_out.txt"
    errFile := A_Temp "\learn_err.txt"
    try FileDelete(outFile)
    try FileDelete(errFile)

    cmd := Format('cmd /c chcp 65001>nul & "{1}" "{2}" 1>"{3}" 2>"{4}"', px, ex, outFile, errFile)
    DbgCP("ExplainNow -> " cmd)
    SignalExplainerBusy()
    exitCode := RunWait(cmd, , "Hide")
    out := (FileExist(outFile) ? Trim(FileRead(outFile, "UTF-8")) : "")
    err := (FileExist(errFile)  ? FileRead(errFile, "UTF-8")      : "")

    if (exitCode = 0) {
        Toast("Explanation updated")
        DbgCP("ExplainNow OK: " out)
    } else {
        msg := "(Explain exit " exitCode ")`n" (Trim(err)!="" ? err : out)
        MsgBox(msg, "Explain failed", 16)
        DbgCP("ExplainNow ERR: " msg)
    }
}

; Force the color swatches to repaint immediately (no warnings, no flicker)
RefreshColorSwatches() {
    global ui, rectBg, rectTxt, rectName
    global boxBgHex, txtHex, nameHex

    rectBg.Opt("Background" . boxBgHex)
    rectTxt.Opt("Background" . txtHex)
    if IsSet(rectName)
        rectName.Opt("Background" . nameHex)

    for swatch in [rectBg, rectTxt, rectName] {
        if IsSet(swatch) {
            try {
                swatch.Redraw()
            } catch as __swErr {      ; <-- use a unique local name to avoid #Warn
                ; no-op: control may not exist yet during early draws
            }
        }
    }

    DllCall("user32\RedrawWindow", "ptr", ui.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x0181)
}

; -- Explainer version (same idea, different controls/vars)
RefreshColorSwatches_EW() {
    global ui, rectBg_EW, rectTxt_EW
    global boxBgHex_EW, txtHex_EW
    rectBg_EW.Opt("Background" . boxBgHex_EW)
    rectTxt_EW.Opt("Background" . txtHex_EW)
    for swatch in [rectBg_EW, rectTxt_EW] {
        try swatch.Redraw()
    }
    DllCall("user32\RedrawWindow", "ptr", ui.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x0181)
}

; --- Clear selection highlight in editable ComboBox (removes white-on-blue) ---
ComboUnselectText(ctrl) {
    try {
        len := StrLen(ctrl.Text)
        ; 1) Ask the ComboBox to clear selection
        ; CB_SETEDITSEL = 0x0143  (start & end packed into lParam)
        SendMessage(0x0143, 0, (len << 16) | len, ctrl.Hwnd)

        ; 2) Also clear on the CHILD Edit (some themes still show highlight otherwise)
        hEdit := DllCall("FindWindowEx", "ptr", ctrl.Hwnd, "ptr", 0, "str", "Edit", "ptr", 0, "ptr")
        if (hEdit)
            ; EM_SETSEL = 0x00B1  (wParam=start, lParam=end)
            SendMessage(0x00B1, len, len, hEdit)
    }
}
; --- NEW: remove ES_NOHIDESEL from the ComboBox's child Edit so selection isn't painted ---
FixEditableCombo(ctrl) {
    try {
        hEdit := DllCall("FindWindowEx", "ptr", ctrl.Hwnd, "ptr", 0, "str", "Edit", "ptr", 0, "ptr")
        if !hEdit
            return
        GWL_STYLE := -16, ES_NOHIDESEL := 0x100
        get := (A_PtrSize=8 ? "GetWindowLongPtr" : "GetWindowLong")
        set := (A_PtrSize=8 ? "SetWindowLongPtr" : "SetWindowLong")
        style := DllCall(get, "ptr", hEdit, "int", GWL_STYLE, "ptr")
        if (style & ES_NOHIDESEL) {
            DllCall(set, "ptr", hEdit, "int", GWL_STYLE, "ptr", style & ~ES_NOHIDESEL)
            ; force a non-client refresh so the new style is applied immediately
            DllCall("SetWindowPos", "ptr", hEdit, "ptr", 0
                , "int", 0, "int", 0, "int", 0, "int", 0
                , "uint", 0x0027) ; NOSIZE|NOMOVE|NOZORDER|FRAMECHANGED
        }
        ; also ensure nothing is selected
        SendMessage(0x00B1, 0, 0, hEdit) ; EM_SETSEL
    }
}

; Convenience: fix all editable combos we use
FixAllEditableCombos() {
    global ddlAProv, ddlA_GM, ddlTR, ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt, ddlPost
    for c in [ddlAProv, ddlA_GM, ddlTR, ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt, ddlPost]
        FixEditableCombo(c)
}

; ============ One-key Explainer show/hide + (optional) request ============
; Behavior (unchanged intent, but "hide" now delegates to your existing toggle hotkey):
; - If Explainer not running: launch, set topmost, Immediately ExplainNow()
; - If Explainer running and HIDDEN: show, set topmost, ExplainNow()
; - If Explainer running and VISIBLE + NOT topmost: set topmost, ExplainNow()
; - If Explainer running and VISIBLE + TOPMOST: trigger your "hide_show_explainer" hotkey
;   (safer than WinHide to avoid overlay crashes). If not configured, just drop topmost.
CP_LaunchExplainerRequest(*) {
    oldMode := A_TitleMatchMode
    SetTitleMatchMode 3  ; exact "Explainer"

    oldDHW := A_DetectHiddenWindows
    DetectHiddenWindows true

    if !WinExist("Explainer") {
        ; Not present -> launch
        LaunchExplainerOverlay()
        WinWait("Explainer",, 3)
        if WinExist("Explainer") {
            hwnd := WinExist("Explainer")
            try ShowWindowNoActivate(hwnd)
            try WinSetAlwaysOnTop(1, "Explainer")
            ExplainNow()
        }
        DetectHiddenWindows oldDHW
        SetTitleMatchMode oldMode
        return
    }

    hwnd := WinExist("Explainer")
    isHidden := !DllCall("user32\IsWindowVisible", "ptr", hwnd, "int")

    if isHidden {
        ; Hidden -> show + topmost + request
        try ShowWindowNoActivate(hwnd)
        try WinSetAlwaysOnTop(1, "ahk_id " hwnd)
        ExplainNow()
        DetectHiddenWindows oldDHW
        SetTitleMatchMode oldMode
        return
    }

    if IsWindowTopmost("ahk_id " hwnd) {
        ; Visible + Topmost -> delegate "hide" to your existing overlay toggle hotkey
                ; Read the configured toggle hotkey from INI
        global iniPath, overlayDir
        toggleHK := Trim(IniRead(iniPath, "hotkeys", "hide_show_explainer", ""))

                ; Primary: signal the overlay directly (bullet-proof vs synthetic keys)
        try {
            if !DirExist(overlayDir)
                DirCreate(overlayDir)
            FileAppend("", overlayDir "\cmd.toggle_explainer", "UTF-8")
        } catch as ex {
            ; Fallback: if the file signal fails for any reason, use the userâ€™s toggle hotkey
            if (toggleHK != "") {
                try SendEvent toggleHK
            } else {
                ; Last-resort soft hide: just drop topmost
                try WinSetAlwaysOnTop(0, "ahk_id " hwnd)
            }
        }
    } else {
        ; Visible + Not topmost -> show + request
        try WinSetAlwaysOnTop(1, "ahk_id " hwnd)
        ExplainNow()
    }

    DetectHiddenWindows oldDHW
    SetTitleMatchMode oldMode
}

; --- Helper: send the configured hotkey for a given action name ---
; Falls back to the default mapping if the user hasn't customized it yet.
FireOverlayCommandAction(action) {
    global overlayDir
    cmdName := ""
    if (action = "screenshot_translate")
        cmdName := "cmd.oneshot_translate"
    else if (action = "take_screenshot")
        cmdName := "cmd.take_screenshot"
    else if (action = "screenshot_translation")
        cmdName := "cmd.screenshot_translation"
    else
        return ""

    try {
        EnsureOverlayDir()
        path := overlayDir "\" cmdName
        if FileExist(path)
            FileDelete(path)
        FileAppend("", path, "UTF-8")
        return path
    } catch as ex {
        DbgCP("FireOverlayCommandAction failed for " action ": " ex.Message)
    }
    return ""
}

HotkeyToSendSpec(hk) {
    hk := Trim(hk)
    if (hk = "")
        return ""

    ; Hotkey notation accepts +F9, but Send needs +{F9}.
    ; Keep simple character keys like ^+t unchanged.
    while (hk != "" && InStr("*~$", SubStr(hk, 1, 1)))
        hk := SubStr(hk, 2)

    mods := ""
    while (hk != "" && InStr("^+!#", SubStr(hk, 1, 1))) {
        mods .= SubStr(hk, 1, 1)
        hk := SubStr(hk, 2)
    }

    if (hk = "")
        return mods
    if RegExMatch(hk, "^\{.+\}$")
        return mods hk
    if RegExMatch(hk, "i)^(F[1-9]|F1[0-9]|F2[0-4]|Home|End|PgUp|PgDn|PageUp|PageDown|Insert|Ins|Delete|Del|Up|Down|Left|Right|Space|Tab|Enter|Escape|Esc|Backspace|BS)$")
        return mods "{" hk "}"
    return mods hk
}

FireHotkeyAction(action) {
    global iniPath, hotkeyDefaults

    ; Push latest Screenshot-Translation settings (incl. â€œHighlight guessed subjectsâ€)
    ; immediately before any screenshot-related trigger.
    if (action = "screenshot_translate"
     || action = "screenshot_translation"
     || action = "take_screenshot"
     || action = "recapture_region") {
        try ApplyShotSettings()
    }

    ; Prefer a direct command file for overlay screenshot actions. This avoids
    ; focus-sensitive synthetic hotkeys from the Control Panel. If an old
    ; running overlay doesn't consume it, fall back to the configured hotkey.
    cmdPath := FireOverlayCommandAction(action)
    if (cmdPath != "") {
        Sleep(750)
        if !FileExist(cmdPath)
            return
        try FileDelete(cmdPath)
    }

    try {
        hk := Trim(IniRead(iniPath, "hotkeys", action, ""))
        if (hk = "" && hotkeyDefaults.Has(action))
            hk := hotkeyDefaults[action]
        if (hk != "") {
            SendEvent HotkeyToSendSpec(hk)
            return
        }
    } catch as ex {
        ; ignore and fall through to beep
    }
    SoundBeep 1500
    ToolTip("No hotkey set for '" action "'", , , 3)
    SetTimer(() => ToolTip("",,,3), -1200)
}

; ===================== Hotkey UI helpers =====================
; Opens a modal dialog with a "Hotkey" capture field.
; Returns: AHK v2 hotkey string (e.g. "^!t"), "" if clearing, or "__CANCEL__" on cancel.
CaptureHotkey(init := "") {
    global ui
    result := ""
    closed := false
    dlg := Gui("+Owner" ui.Hwnd " +AlwaysOnTop", "Set Hotkey")
    dlg.MarginX := 14, dlg.MarginY := 12

    dlg.Add("Text", "xm", "Press the new shortcut (or clear the field to disable).")
    hk := dlg.Add("Hotkey", "xm y+6 w260 vHK")
    if (init != "")
        hk.Value := init

    dlg.Add("Text", "xm y+10 w0 h0")  ; spacer
    btnOK := dlg.Add("Button", "xm y+8 w90 Default", "OK")
    btnCancel := dlg.Add("Button", "x+6 w90", "Cancel")

    btnOK.OnEvent("Click", (*) => (result := hk.Value, closed := true, dlg.Destroy()))
    btnCancel.OnEvent("Click", (*) => (result := "__CANCEL__", closed := true, dlg.Destroy()))
    dlg.OnEvent("Escape", (*) => (result := "__CANCEL__", closed := true, dlg.Destroy()))
    dlg.OnEvent("Close",  (*) => (result := "__CANCEL__", closed := true))  ; window X

    dlg.Show("AutoSize Center")
    while !closed
        Sleep(30)
    return result
}

; (Optional pretty-printer; we can use raw AHK strings for now)
HotkeyPretty(hk) {
    repl := Map("^","Ctrl","+","Shift","!","Alt","#","Win")
    out := ""
    For i, hkCh in StrSplit(hk, "")
        out .= repl.Has(hkCh) ? repl[hkCh] " + " : Format("{:U}", hkCh)
    out := RegExReplace(out, "\s*\+\s*$", "")
    return out
}

NormalizeHotkey(hk) {
    mods := Map("^","^","+","+","!","!","#","#")
    out := ""
    i := 1, len := StrLen(hk)
    While (i <= len) {
        hkCh := SubStr(hk, i, 1)
        If (mods.Has(hkCh)) {
            out .= mods[hkCh]
        } else {
            out .= hkCh
        }
        i++
    }
    return out
}

; ===================== Hotkey storage & conflict helpers =====================
; Build a Map(action -> string) from current UI rows
Hotkeys_GetMap() {
    global hotkeyActions, hkEdits
    m := Map()
    for act in hotkeyActions {
        m[act] := hkEdits[act].Value
    }
    return m
}

; Check duplicates (same non-empty hotkey assigned to multiple actions)
; Returns Map(hotkeyString -> [actions...]) for those with conflicts
Hotkeys_FindConflicts() {
    m := Hotkeys_GetMap()
    rev := Map()
    for actionName, hk in m {
        if (hk = "")
            continue
        if !rev.Has(hk)
            rev[hk] := []
        rev[hk].Push(actionName)
}
    conflicts := Map()
    for hk, arr in rev {
        if (arr.Length >= 2)
            conflicts[hk] := arr
    }
    return conflicts
}

; Visualize conflicts by tinting Edit boxes and updating banner text
Hotkeys_ShowConflicts() {
    global hkEdits, hkConflictText
    ; Reset all tints
    for actionName, ctrl in hkEdits {
        try ctrl.Opt("Background")
    }
    c := Hotkeys_FindConflicts()
    if (c.Count = 0) {
        if IsSet(hkConflictText)
            hkConflictText.Text := ""
        return 0
    }
    ; Tint conflicted rows and compose banner
    banner := []
    for hk, arr in c {
        for actionName in arr {
            ctrl := hkEdits[actionName]
            try ctrl.Opt("BackgroundFFCCCC")
        }
        banner.Push(Format("{}  <-  {}", hk, JoinWith(arr, ", ")))
    }
    if IsSet(hkConflictText)
        hkConflictText.Text := "Conflicts: " . JoinWith(banner, "    ")
    return c.Count
}

; Save to INI [hotkeys] section
Hotkeys_SaveToIni() {
    global iniPath, hotkeyActions, hkEdits
    for act in hotkeyActions {
    IniWrite(hkEdits[act].Value, iniPath, "hotkeys", act)
    }
}

; Reload UI rows from INI (discard in-memory edits)
Hotkeys_ReloadFromIni() {
    global iniPath, hotkeyActions, hotkeyDefaults, hkEdits
    for act in hotkeyActions {
    curVal := IniRead(iniPath, "hotkeys", act, hotkeyDefaults[act])
    hkEdits[act].Value := curVal
    }
    Hotkeys_ShowConflicts()
}
; ===========================================================================

; small utility used above (renamed to avoid conflicts)
JoinWith(arr, sep) {
    out := ""
    for i, v in arr
        out .= (i=1 ? "" : sep) v
    return out
}
; =============================================================

; ---- Add/Delete model helpers (TOP-LEVEL, outside any braces) ----
; ===================== Hotkeys row handlers =====================
Hotkey_Row_Change(action, *) {
    global hkEdits, hkDirty
    curVal := hkEdits[action].Value
    new   := CaptureHotkey(curVal)
    if (new = "__CANCEL__")
        return
    hkEdits[action].Value := NormalizeHotkey(new)
    hkDirty := true
    Hotkeys_ShowConflicts()
    ; Auto-apply & persist immediately after user confirms OK
    Hotkeys_OnApply()
}

Hotkey_Row_Disable(action, *) {
    global hkEdits, hkDirty
    hkEdits[action].Value := ""
    hkDirty := true
    Hotkeys_ShowConflicts()
    ; Auto-apply & persist immediately when disabling
    Hotkeys_OnApply()
}

Hotkey_Row_Default(action, *) {
    global hkEdits, hotkeyDefaults, hkDirty
    if (hotkeyDefaults.Has(action))
        hkEdits[action].Value := NormalizeHotkey(hotkeyDefaults[action])  ; keep notation consistent (^ before +)
    else
        hkEdits[action].Value := ""
    hkDirty := true
    Hotkeys_ShowConflicts()
    ; Apply immediately: writes INI, drops hotkeys.reload, and rebinds in the overlay
    Hotkeys_OnApply()
}

; ===================== Hotkeys Apply/Revert =====================
Hotkeys_OnApply() {
    global hkDirty
    ; Refuse apply if there are conflicts (confirm override if you want)
    hasConflicts := Hotkeys_ShowConflicts()
    if (hasConflicts) {
        res := MsgBox("There are duplicate hotkeys. Apply anyway?", "Conflicts detected", 0x21)
        if (res != "OK")
            return
    }
    Hotkeys_SaveToIni()

    ; --- signal the overlay to live-reload hotkeys ---
    global overlayDir
    if !DirExist(overlayDir)
        DirCreate(overlayDir)
    try FileAppend("", overlayDir "\hotkeys.reload")
    ; -----------------------------------------------

    Rebind_LaunchExplainerRequest()
	Rebind_ExplainLastTranslation()
	Rebind_StartStopAudio()
	Rebind_HideShowControlPanel()

    hkDirty := false
    ToolTip("Hotkeys saved", A_ScreenWidth-220, 20)
    SetTimer(() => ToolTip(), -900)
}

Hotkeys_OnRevert() {
    global hkDirty
    Hotkeys_ReloadFromIni()
    hkDirty := false
    ToolTip("Changes reverted", A_ScreenWidth-220, 20)
    SetTimer(() => ToolTip(), -900)
}
; ===============================================================

; ===============================================================

SetComboItems(combo, arr) {
    SendMessage(0x14B, 0, 0, combo.Hwnd) ; CB_RESETCONTENT
    if (arr.Length)
        combo.Add(arr)
}

SetComboToExistingItem(combo, arr, desired := "") {
    desired := Trim(desired)
    idx := desired != "" ? ArrIndexOf(arr, desired) : 0
    if (idx) {
        combo.Choose(idx)
        return arr[idx]
    }
    if (arr.Length) {
        combo.Choose(1)
        return arr[1]
    }
    try combo.Text := ""
    return ""
}

RefreshModelCombos(key, activeCombo := 0, activePreferredText := "") {
    global model_openai_img, model_gemini_img, model_openai_explain, model_gemini_explain
    global model_openai_audio, model_gemini_audio
    global ddlIMG, ddlIMG_GM, ddlEOpenAI, ddlEGem, ddlTR, ddlA_GM

    arr := 0
    combos := []
    switch key {
        case "openai_img":
            arr := model_openai_img
            if IsSet(ddlIMG)
                combos.Push(ddlIMG)
        case "gemini_img":
            arr := model_gemini_img
            if IsSet(ddlIMG_GM)
                combos.Push(ddlIMG_GM)
        case "openai_explain":
            arr := model_openai_explain
            if IsSet(ddlEOpenAI)
                combos.Push(ddlEOpenAI)
        case "gemini_explain":
            arr := model_gemini_explain
            if IsSet(ddlEGem)
                combos.Push(ddlEGem)
        case "openai_audio":
            arr := model_openai_audio
            if IsSet(ddlTR)
                combos.Push(ddlTR)
        case "gemini_audio":
            arr := model_gemini_audio
            if IsSet(ddlA_GM)
                combos.Push(ddlA_GM)
        default:
            return
    }
    if !IsObject(arr)
        return
    ModelListSort(arr)

    for combo in combos {
        modelSelectionText := Trim(combo.Text)
        if (IsObject(activeCombo) && combo.Hwnd = activeCombo.Hwnd && activePreferredText != "")
            modelSelectionText := activePreferredText
        SetComboItems(combo, arr)
        SetComboToExistingItem(combo, arr, modelSelectionText)
    }
}

AddModelValue(arr, key, combo, newModel) {
    newModel := Trim(newModel)
    if (newModel = "")
        return false
    for v in arr
        if (StrLower(v) = StrLower(newModel)) {
            MsgBox("Already in the list: " newModel, "Add model")
            return false
        }
    arr.Push(newModel)            ; modifies the original array
    ModelListWrite(key, arr)
    RefreshModelCombos(key, combo, newModel)
    DbgCP("Model added under [" key "]: " newModel)
    return true
}

AddModel(arr, key, combo) {
    newModel := Trim(InputBox("Add model:", "Add").Value)
    return AddModelValue(arr, key, combo, newModel)
}

CPApplyOwnedDialogTheme(dlg) {
    global controlDarkMode
    if !(IsObject(dlg) && dlg.Hwnd)
        return

    dialogColors := CPPalette(controlDarkMode)
    dlg.BackColor := dialogColors["window"]
    CPApplyDarkTitleBar(dlg.Hwnd, controlDarkMode)
    try CPSetPreferredAppDarkMode(controlDarkMode, dlg.Hwnd)
    try {
        for controlHwnd in WinGetControlsHwnd("ahk_id " dlg.Hwnd)
            CPApplyThemeToControl(controlHwnd, controlDarkMode)
    }
    try DllCall("user32\RedrawWindow", "ptr", dlg.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x185)
}

CPShowTextEditorDialog(dlg, editorCtrl) {
    dlg.Show()
    CPApplyOwnedDialogTheme(dlg)

    ; Keep the editor ready for typing without selecting all existing text.
    try editorCtrl.Focus()
    try SendMessage(0x00B1, 0, 0, editorCtrl.Hwnd) ; EM_SETSEL
    try SendMessage(0x00B7, 0, 0, editorCtrl.Hwnd) ; EM_SCROLLCARET
}

AddModelDialogSelect(which, rbOnline, rbManual, focusChoice := true, *) {
    chooseOnline := (which = "online")
    rbOnline.Value := chooseOnline ? 1 : 0
    rbManual.Value := chooseOnline ? 0 : 1
    if focusChoice
        (chooseOnline ? rbOnline : rbManual).Focus()
}

AddModelDialogNavigate(direction, rbOnline, rbManual, btnContinue, btnCancel, *) {
    focusHwnd := DllCall("user32\GetFocus", "ptr")
    if (focusHwnd = rbOnline.Hwnd) {
        if (direction = "Down" || direction = "Right")
            rbManual.Focus()
        return
    }
    if (focusHwnd = rbManual.Hwnd) {
        if (direction = "Up" || direction = "Left")
            rbOnline.Focus()
        else if (direction = "Down")
            btnContinue.Focus()
        return
    }
    if (focusHwnd = btnContinue.Hwnd) {
        if (direction = "Up")
            rbManual.Focus()
        else if (direction = "Right")
            btnCancel.Focus()
        return
    }
    if (focusHwnd = btnCancel.Hwnd) {
        if (direction = "Up")
            rbManual.Focus()
        else if (direction = "Left")
            btnContinue.Focus()
        return
    }
    (rbOnline.Value ? rbOnline : rbManual).Focus()
}

AddModelDialogActivate(rbOnline, rbManual, btnContinue, btnCancel, *) {
    focusHwnd := DllCall("user32\GetFocus", "ptr")
    if (focusHwnd = rbOnline.Hwnd) {
        AddModelDialogSelect("online", rbOnline, rbManual)
        return
    }
    if (focusHwnd = rbManual.Hwnd) {
        AddModelDialogSelect("manual", rbOnline, rbManual)
        return
    }
    if (focusHwnd = btnCancel.Hwnd)
        SendMessage(0x00F5, 0, 0, btnCancel.Hwnd) ; BM_CLICK
    else
        SendMessage(0x00F5, 0, 0, btnContinue.Hwnd)
}

CPApplyDialogRadioTheme(radioCtrl) {
    global controlDarkMode
    colors := CPPalette(controlDarkMode)
    try radioCtrl.SetFont("c" colors["text"])
    if controlDarkMode {
        ; Unthemed radio text honors WM_CTLCOLORBTN, unlike the dark Explorer radio.
        try DllCall("uxtheme\SetWindowTheme", "ptr", radioCtrl.Hwnd, "wstr", "", "wstr", "")
    }
    try DllCall("user32\InvalidateRect", "ptr", radioCtrl.Hwnd, "ptr", 0, "int", 1)
}

AddModelSourceDialog(provider, purpose) {
    global ui
    providerLabel := StrLower(provider) = "gemini" ? "Gemini" : "OpenAI"
    purposeName := StrLower(purpose)
    if (purposeName = "screenshot")
        purposeLabel := "Screenshot Translation"
    else if (purposeName = "explanation")
        purposeLabel := "Explanation"
    else if (purposeName = "audio")
        purposeLabel := "Audio Translation"
    else
        purposeLabel := "JRPG Translator"

    result := "cancel"
    closed := false
    dlg := Gui("+Owner" ui.Hwnd " +AlwaysOnTop", "Add model")
    dlg.MarginX := 18, dlg.MarginY := 16
    dlg.SetFont("s10", "Segoe UI")
    dlg.Add("Text", "xm w390", "Add a " providerLabel " model for " purposeLabel ".")
    dlg.Add("Text", "xm y+5 w390", "Choose where the model ID should come from.")
    rbOnline := dlg.Add("Radio", "xm y+14 w390 h28 Checked Group", "Browse models available to this API key online")
    rbManual := dlg.Add("Radio", "xm y+6 w390 h28", "Enter a model ID manually")
    btnContinue := dlg.Add("Button", "xm y+16 w120 Default", "Continue")
    btnCancel := dlg.Add("Button", "x+8 w100", "Cancel")

    finish := (selection) => (result := selection, closed := true, dlg.Destroy())
    rbOnline.OnEvent("Click", AddModelDialogSelect.Bind("online", rbOnline, rbManual, true))
    rbManual.OnEvent("Click", AddModelDialogSelect.Bind("manual", rbOnline, rbManual, true))
    btnContinue.OnEvent("Click", (*) => finish(rbOnline.Value ? "online" : "manual"))
    btnCancel.OnEvent("Click", (*) => finish("cancel"))
    dlg.OnEvent("Escape", (*) => finish("cancel"))
    dlg.OnEvent("Close", (*) => finish("cancel"))

    dialogHotIf := "ahk_id " dlg.Hwnd
    dialogArrowHotkeys := Map("$Up", "Up", "$Down", "Down", "$Left", "Left", "$Right", "Right")
    HotIfWinActive(dialogHotIf)
    for keyName, direction in dialogArrowHotkeys
        try Hotkey(keyName, AddModelDialogNavigate.Bind(direction, rbOnline, rbManual, btnContinue, btnCancel), "On")
    try Hotkey("$Enter", AddModelDialogActivate.Bind(rbOnline, rbManual, btnContinue, btnCancel), "On")
    try Hotkey("$NumpadEnter", AddModelDialogActivate.Bind(rbOnline, rbManual, btnContinue, btnCancel), "On")
    HotIfWinActive()

    try {
        dlg.Show("AutoSize Center")
        CPApplyOwnedDialogTheme(dlg)
        CPApplyDialogRadioTheme(rbOnline)
        CPApplyDialogRadioTheme(rbManual)
        AddModelDialogSelect("online", rbOnline, rbManual)
        while !closed
            Sleep(30)
    } finally {
        HotIfWinActive(dialogHotIf)
        for keyName, direction in dialogArrowHotkeys
            try Hotkey(keyName, "Off")
        try Hotkey("$Enter", "Off")
        try Hotkey("$NumpadEnter", "Off")
        HotIfWinActive()
    }
    return result
}

DeleteModel(arr, key, combo) {
    selText := Trim(combo.Text)
    If (selText = "")
        Return
    for i, v in arr
        if (StrLower(v) = StrLower(selText)) {
            arr.RemoveAt(i)      ; modifies the original array
            ModelListWrite(key, arr)
            RefreshModelCombos(key, combo, "")
            DbgCP("Model removed under [" key "]: " selText)
            return
        }
    MsgBox("Not found in list: " selText, "Delete model")
}

; =========================
; AUDIO start/stop
; =========================
StartAudio(*) {
    global pythonExe, audioScript, trModel, gPidAudio
    global audioProvider, geminiAudioModel, audioTargetLang, gJustStoppedUntil, gLastAction
    ; also read current UI controls (so Start works without pressing Apply)
    global ddlAProv, ddlTR, ddlA_GM, ddlAudioTarget
    ; used by logging / env for prompt/speaker
    global ddlSpeaker, debugMode
    gLastAction := "start"
    px := ResolvePath(pythonExe)
    ap := ResolvePath(audioScript)
    if (gPidAudio && ProcessExist(gPidAudio)) {
        ToolTip("Audio already running"), SetTimer(() => ToolTip(""), -800)
        return
    }
    if !(FileExist(px) && FileExist(ap)) {
        MsgBox("Set valid paths for python.exe and audio script first.`n`npythonExe:`n" px "`n`naudioScript:`n" ap, "Missing", 48)
        return
    }
    ; Snapshot CURRENT UI (so Start works even if user didn't press Apply)
    tProv   := (IsSet(ddlAProv)     ? Trim(ddlAProv.Text)    : audioProvider)
    tTR     := (IsSet(ddlTR)        ? Trim(ddlTR.Text)       : trModel)
    tGModel := (IsSet(ddlA_GM)      ? Trim(ddlA_GM.Text)     : geminiAudioModel)
    tLang   := (IsSet(ddlAudioTarget) ? Trim(ddlAudioTarget.Text) : audioTargetLang)

    audioProvider    := tProv, trModel  := tTR
    geminiAudioModel := tGModel
    audioTargetLang  := tLang

    EnvSet("AUDIO_PROVIDER", (StrLower(audioProvider) = "gemini") ? "gemini" : "openai")
    EnvSet("TEXT_PROVIDER",  (StrLower(audioProvider) = "gemini") ? "gemini" : "openai")
    EnvSet("TRANSLATE_MODEL", trModel)
    EnvSet("GEMINI_AUDIO_MODEL", geminiAudioModel)
    EnvSet("TARGET_LANGUAGE_CODE", AudioTargetCode(audioTargetLang))
    EnvSet("TARGET_LANGUAGE_NAME", AudioTargetName(audioTargetLang))
    EnvSet("SETTINGS_DIR", A_ScriptDir "\Settings")
    EnvSet "JRPG_DEBUG", (debugMode ? "1" : "0")
    EnvSet("PYTHONIOENCODING","utf-8")
	; Select loopback device: empty => default output
    spick := Trim(ddlSpeaker.Text)
    if (spick = "" || spick = "[Windows Default]")
        EnvSet("SPEAKER_NAME", "")
    else
        EnvSet("SPEAKER_NAME", spick)

    DbgCP("StartAudio live provider=" audioProvider " openaiModel=" trModel " geminiModel=" geminiAudioModel " target=" AudioTargetCode(audioTargetLang) " speaker=" spick)
        try {
        gPidAudio := Run('"' px '" "' ap '"', , "Hide")
    } Catch as exrr {
        MsgBox("Failed to start audio script:`n" exrr.Message, "Error", 16)
        UpdateStatus()
        DbgCP("StartAudio failed: " exrr.Message)
        return
    }

    /*
      NEW: give Windows up to ~2 seconds to surface the process
      (prevents the â€œhave to click twiceâ€ / false-fail fallback)
    */
    started := false
    Loop 20 {                         ; 20Ã—100ms = ~2 seconds
        Sleep(100)
        if (gPidAudio && ProcessExist(gPidAudio)) {
            started := true
            break
        }
        ; trust WMI if it already sees our script path
        pids := AudioPidsByScript()
        if (pids.Length) {
            if (!gPidAudio)
                gPidAudio := pids[1]
            started := true
            break
        }
    }

        ; Extra guard: if it died immediately after spawn, treat as failure
    if (started) {
        Sleep(200)
        if !(gPidAudio && ProcessExist(gPidAudio)) {
            started := false
        }
    }
    if (started) {
        DbgCP("StartAudio: process confirmed (pid=" gPidAudio ")")
        UpdateStatus()
        Toast("Audio Translation On")
        return true
    }

        ; --- original error-capture fallback (only skip if we just stopped) ---
    if (gLastAction = "stop" && A_TickCount < gJustStoppedUntil) {
        UpdateStatus()
        return
    }
    outFile := A_Temp "\jrpg_audio_out.txt"
    errFile := A_Temp "\jrpg_audio_err.txt"
    try FileDelete(outFile)
    try FileDelete(errFile)
    cmd := Format('cmd /c chcp 65001>nul & "{1}" "{2}" 1>"{3}" 2>"{4}"', px, ap, outFile, errFile)
    exitCode := RunWait(cmd, , "Hide")

    out := (FileExist(outFile) ? Trim(FileRead(outFile, "UTF-8")) : "")
    err := (FileExist(errFile)  ? FileRead(errFile, "UTF-8")        : "")
    ; keep the old silero cache line filter
    err := RegExReplace(err, "(?im)^\s*Using cache found in .+snakers4_silero-vad_master\s*$", "")
    ; drop Gemini/gRPC noise
    err := FilterPythonStderr(err)

    if (exitCode = 0) {
        UpdateStatus()
        return
    }

    msg := "(Python exit code " exitCode ")`n"
    if (Trim(err) != "")
        msg .= "(stderr)`n" Trim(err)
    else if (Trim(out) != "")
        msg .= Trim(out)
    else
        msg .= "No output captured."
    MsgBox(msg, "Audio failed to start", 16)
    DbgCP("StartAudio stderr/out: " msg)
    UpdateStatus()
}

AudioPidsByScript() {
    global audioScript
    ap := ResolvePath(audioScript)
    apL := StrLower(StrReplace(ap, "/", "\"))   ; normalize & lower
    out := []
    try {
        wm := ComObjGet("winmgmts:")
        for p in wm.ExecQuery("Select ProcessId,CommandLine from Win32_Process Where Name='python.exe'") {
            cmd := p.CommandLine ? p.CommandLine : ""
            cmdL := StrLower(StrReplace(cmd, "/", "\"))  ; normalize & lower
            if (InStr(cmdL, apL) && !InStr(cmdL, "--list-speakers"))
                out.Push(p.ProcessId)
        }
    }
    return out
}

OverlayPidsByScript() {
    global overlayAhk
    ov := ResolvePath(overlayAhk)
    ovL := StrLower(StrReplace(ov, "/", "\"))  ; normalize & lower
    out := []
    try {
        wm := ComObjGet("winmgmts:")
        for p in wm.ExecQuery("Select ProcessId,CommandLine,Name from Win32_Process Where Name='AutoHotkey64.exe' OR Name='AutoHotkey.exe'") {
            cmd := p.CommandLine ? p.CommandLine : ""
            cmdL := StrLower(StrReplace(cmd, "/", "\"))  ; normalize & lower
            if InStr(cmdL, ovL)
                out.Push(p.ProcessId)
        }
    }
    return out
}

DumpWindowsForPids(pids) {
    if (!pids.Length) {
        DbgCP("DumpWindowsForPids: no PIDs given")
        return
    }
    DbgCP("DumpWindowsForPids: scanning " pids.Length " PID(s)")
    for hwnd in WinGetList() {
        pid := WinGetPID("ahk_id " hwnd)
        for p in pids {
            if (pid = p) {
                ttl := WinGetTitle("ahk_id " hwnd)
                cls := WinGetClass("ahk_id " hwnd)
                vis := DllCall("IsWindowVisible", "ptr", hwnd, "int")
                DbgCP(Format("  pid={} hwnd=0x{:X} class={} visible={} title='{}'", pid, hwnd, cls, vis, ttl))
                break
            }
        }
    }
}

; --- Filter noisy-but-benign stderr from Gemini/gRPC/absl ---
FilterPythonStderr(s) {
    ; absl pre-init notice
    s := RegExReplace(s, "(?im)^WARNING:\s+All log messages before absl::InitializeLog\(\).*\R?", "")
    ; gRPC ALTS creds ignored (not on GCP)
    s := RegExReplace(s, "(?im)^\s*E\d+\s+\S+\s+\d+\s+alts_credentials\.cc:\d+\]\s+ALTS creds ignored\.[^\r\n]*\R?", "")
    return Trim(s)
}

StopAudio(*) {
    global gPidAudio, gJustStoppedUntil, gLastAction
    gLastAction := "stop"
    gJustStoppedUntil := A_TickCount + 5000
    if (gPidAudio && ProcessExist(gPidAudio)) {
        try ProcessClose(gPidAudio)
    }
    for pid in AudioPidsByScript() {
        try ProcessClose(pid)
    }

    stopped := false
    Loop 20 {
        primaryAlive := gPidAudio && ProcessExist(gPidAudio)
        remaining := AudioPidsByScript()
        if (!primaryAlive && !remaining.Length) {
            stopped := true
            break
        }
        for pid in remaining {
            try ProcessClose(pid)
        }
        Sleep(50)
    }

    if (stopped)
        gPidAudio := 0
    DbgCP("StopAudio() requested; confirmed=" stopped)
    _UpdateStatus()
    Toast(stopped ? "Audio Translation Off" : "Audio Translation still running")
    return stopped
}

; NEW: unified toggle used by the single button
ToggleAudioFromButton(*) {
    if AudioIsRunning() {
        StopAudio()
    } else {
        StartAudio()
    }
}

; Toggle audio translation from the configured hotkey.
StartStopAudio(*) {
    ToggleAudioFromButton()
}

; Run a command hidden and capture its stdout via a temp file (works with python.exe)
ExecCaptureHidden(px, ap, args:="") {
    tmp := A_Temp "\spk_" A_TickCount ".txt"
    ; Capture ONLY stdout (no stderr), hide the window
    cmd := Format('"{1}" /c ""{2}" "{3}" {4} 1> "{5}""'
        , A_ComSpec      ; cmd.exe
        , px             ; python exe (python.exe or pythonw.exe)
        , ap             ; script path
        , args
        , tmp)
    RunWait cmd, , "Hide"
    out := ""
    try out := FileRead(tmp, "UTF-8")
    try FileDelete(tmp)
    return out
}

; Generic companion to ExecCaptureHidden that preserves stderr and the exit code.
; The model-catalog UI uses this so provider/network failures can be explained
; without exposing API keys or opening a console window.
ExecCaptureHiddenResult(px, ap, args := "", tempPrefix := "jrpg_cmd") {
    unique := A_TickCount "_" Random(1000, 9999)
    stdoutPath := A_Temp "\" tempPrefix "_" unique "_out.txt"
    stderrPath := A_Temp "\" tempPrefix "_" unique "_err.txt"
    command := Format('"{1}" /d /s /c ""{2}" "{3}" {4} 1> "{5}" 2> "{6}""'
        , A_ComSpec, px, ap, args, stdoutPath, stderrPath)

    exitCode := -1
    launchError := ""
    try exitCode := RunWait(command, A_ScriptDir, "Hide")
    catch as ex
        launchError := ex.Message

    stdout := ""
    stderr := ""
    try stdout := FileRead(stdoutPath, "UTF-8")
    try stderr := FileRead(stderrPath, "UTF-8")
    try FileDelete(stdoutPath)
    try FileDelete(stderrPath)
    if (launchError != "")
        stderr := launchError (stderr != "" ? "`n" stderr : "")

    return Map("exitCode", exitCode, "stdout", stdout, "stderr", stderr)
}

ModelCatalogQuery(provider, purpose, forceRefresh := false) {
    global pythonExe
    result := Map(
        "ok", false,
        "provider", StrLower(provider),
        "purpose", StrLower(purpose),
        "source", "none",
        "fetchedAt", "",
        "models", [],
        "displayNames", Map(),
        "warnings", [],
        "error", ""
    )

    px := ResolvePath(pythonExe)
    helper := ResolvePath(".\scripts\model_catalog.py")
    if (px = "" || !FileExist(px)) {
        result["error"] := "The configured Python executable was not found: " px
        return result
    }
    if (helper = "" || !FileExist(helper)) {
        result["error"] := "The online model helper was not found: " helper
        return result
    }

    args := '--provider "' StrLower(provider) '" --purpose "' StrLower(purpose) '" --format ahk'
    if forceRefresh
        args .= " --refresh"
    captured := ExecCaptureHiddenResult(px, helper, args, "jrpg_models")
    output := StrReplace(captured["stdout"], "`r", "")
    lines := StrSplit(output, "`n")
    if (!lines.Length || Trim(lines[1]) != "JRPG_MODEL_CATALOG_V1") {
        errorText := Trim(captured["stderr"])
        if (errorText = "")
            errorText := "The model helper returned an unreadable response."
        result["error"] := errorText
        return result
    }

    for lineNumber, line in lines {
        if (lineNumber = 1 || line = "")
            continue
        fields := StrSplit(line, "`t",, 3)
        fieldName := fields.Length ? fields[1] : ""
        fieldValue := fields.Length >= 2 ? fields[2] : ""
        switch fieldName {
            case "STATUS":
                result["ok"] := fieldValue = "OK"
            case "PROVIDER":
                result["provider"] := fieldValue
            case "PURPOSE":
                result["purpose"] := fieldValue
            case "SOURCE":
                result["source"] := fieldValue
            case "FETCHED_AT":
                result["fetchedAt"] := fieldValue
            case "MODEL":
                if (fieldValue != "") {
                    result["models"].Push(fieldValue)
                    result["displayNames"][fieldValue] := fields.Length >= 3 ? fields[3] : fieldValue
                }
            case "WARNING":
                if (fieldValue != "")
                    result["warnings"].Push(fieldValue)
            case "ERROR":
                result["error"] := fieldValue
        }
    }

    if (!result["ok"] && result["error"] = "")
        result["error"] := Trim(captured["stderr"]) != "" ? Trim(captured["stderr"]) : "The model list could not be loaded."
    return result
}

ModelAlreadyAdded(modelArray, modelId) {
    modelNeedle := StrLower(Trim(modelId))
    for modelEntry in modelArray
        if (StrLower(Trim(modelEntry)) = modelNeedle)
            return true
    return false
}

ModelCatalogQueryWithFeedback(provider, purpose, forceRefresh := false) {
    providerName := StrLower(provider) = "gemini" ? "Gemini" : "OpenAI"
    loadingText := forceRefresh ? "Refreshing " : "Loading "
    ToolTip(loadingText providerName " models...")
    try return ModelCatalogQuery(provider, purpose, forceRefresh)
    finally ToolTip()
}

ModelPickerPopulate(modelListBox, modelStatus, modelAddButton, catalog, existingModels) {
    SendMessage(0x0184, 0, 0, modelListBox.Hwnd) ; LB_RESETCONTENT
    availableModels := []
    for modelId in catalog["models"]
        if !ModelAlreadyAdded(existingModels, modelId)
            availableModels.Push(modelId)

    if availableModels.Length {
        modelListBox.Add(availableModels)
        modelListBox.Choose(1)
    }
    modelListBox.Enabled := availableModels.Length > 0
    modelAddButton.Enabled := availableModels.Length > 0

    catalogSource := catalog["source"]
    if (availableModels.Length = 0)
        statusText := "All compatible models in this catalog are already in the list."
    else if (catalogSource = "online")
        statusText := availableModels.Length " available models loaded online."
    else if (catalogSource = "stale_cache")
        statusText := availableModels.Length " available models loaded from an older cache because the online refresh failed."
    else
        statusText := availableModels.Length " available models loaded from the local cache."

    if catalog["warnings"].Length
        statusText .= " " catalog["warnings"][1]
    modelStatus.Text := statusText
    return availableModels.Length
}

ModelPickerAccept(modelListBox, finishPicker, *) {
    chosenModel := Trim(modelListBox.Text)
    if (chosenModel = "") {
        SoundBeep(1100, 80)
        return
    }
    finishPicker.Call(chosenModel)
}

ModelPickerNavigate(direction, modelListBox, modelAddButton, refreshButton, cancelButton, *) {
    focusHwnd := DllCall("user32\GetFocus", "ptr")
    if (focusHwnd = modelListBox.Hwnd) {
        itemCount := SendMessage(0x018B, 0, 0, modelListBox.Hwnd) ; LB_GETCOUNT
        selectedIndex := SendMessage(0x0188, 0, 0, modelListBox.Hwnd) ; LB_GETCURSEL
        if (direction = "Down") {
            if (itemCount > 0 && selectedIndex >= itemCount - 1)
                modelAddButton.Focus()
            else if (selectedIndex < itemCount - 1)
                SendMessage(0x0186, selectedIndex + 1, 0, modelListBox.Hwnd) ; LB_SETCURSEL
        } else if (direction = "Up" && selectedIndex > 0) {
            SendMessage(0x0186, selectedIndex - 1, 0, modelListBox.Hwnd)
        }
        return
    }

    if (focusHwnd = modelAddButton.Hwnd) {
        if (direction = "Up" && modelListBox.Enabled)
            modelListBox.Focus()
        else if (direction = "Left")
            cancelButton.Focus()
        else if (direction = "Right")
            refreshButton.Focus()
        return
    }
    if (focusHwnd = refreshButton.Hwnd) {
        if (direction = "Up" && modelListBox.Enabled)
            modelListBox.Focus()
        else if (direction = "Left")
            modelAddButton.Focus()
        else if (direction = "Right")
            cancelButton.Focus()
        return
    }
    if (focusHwnd = cancelButton.Hwnd) {
        if (direction = "Up" && modelListBox.Enabled)
            modelListBox.Focus()
        else if (direction = "Left")
            refreshButton.Focus()
        else if (direction = "Right")
            modelAddButton.Focus()
        return
    }

    if modelListBox.Enabled
        modelListBox.Focus()
    else
        refreshButton.Focus()
}

ModelPickerRefresh(provider, purpose, existingModels, pickerDialog, modelListBox, modelStatus, modelAddButton, refreshButton, *) {
    previousStatus := modelStatus.Text
    modelListBox.Enabled := false
    modelAddButton.Enabled := false
    refreshButton.Enabled := false
    modelStatus.Text := "Refreshing the online model catalog..."

    refreshedCatalog := ModelCatalogQueryWithFeedback(provider, purpose, true)
    refreshButton.Enabled := true
    if !refreshedCatalog["ok"] {
        modelStatus.Text := previousStatus
        modelListBox.Enabled := true
        modelAddButton.Enabled := Trim(modelListBox.Text) != ""
        MsgBox("The online model list could not be refreshed.`n`n" refreshedCatalog["error"], "Browse models", 48)
        return
    }

    ModelPickerPopulate(modelListBox, modelStatus, modelAddButton, refreshedCatalog, existingModels)
    CPApplyOwnedDialogTheme(pickerDialog)
    if modelListBox.Enabled
        modelListBox.Focus()
    else
        refreshButton.Focus()
}

OnlineModelPicker(provider, purpose, existingModels, catalog) {
    global ui
    providerName := StrLower(provider) = "gemini" ? "Gemini" : "OpenAI"
    purposeName := StrLower(purpose)
    if (purposeName = "screenshot")
        purposeLabel := "Screenshot Translation"
    else if (purposeName = "explanation")
        purposeLabel := "Explanation"
    else if (purposeName = "audio")
        purposeLabel := "Audio Translation"
    else
        purposeLabel := "JRPG Translator"
    pickerResult := ""
    pickerClosed := false
    pickerDialog := Gui("+Owner" ui.Hwnd " +AlwaysOnTop", "Browse " providerName " models")
    pickerDialog.MarginX := 18, pickerDialog.MarginY := 16
    pickerDialog.SetFont("s10", "Segoe UI")
    pickerDialog.Add("Text", "xm w560", "Select one model to add to " purposeLabel ".")
    pickerDialog.Add("Text", "xm y+5 w560", "The list contains compatible models available to the configured API key.")
    modelListBox := pickerDialog.Add("ListBox", "xm y+12 w560 r16")
    modelStatus := pickerDialog.Add("Text", "xm y+8 w560 h42")
    modelAddButton := pickerDialog.Add("Button", "xm y+12 w120 Default", "Add model")
    refreshButton := pickerDialog.Add("Button", "x+8 w100", "Refresh")
    cancelButton := pickerDialog.Add("Button", "x+8 w100", "Cancel")

    finishPicker := (selection) => (pickerResult := selection, pickerClosed := true, pickerDialog.Destroy())
    modelAddButton.OnEvent("Click", ModelPickerAccept.Bind(modelListBox, finishPicker))
    modelListBox.OnEvent("DoubleClick", ModelPickerAccept.Bind(modelListBox, finishPicker))
    refreshButton.OnEvent("Click", ModelPickerRefresh.Bind(provider, purpose, existingModels, pickerDialog, modelListBox, modelStatus, modelAddButton, refreshButton))
    cancelButton.OnEvent("Click", (*) => finishPicker.Call(""))
    pickerDialog.OnEvent("Escape", (*) => finishPicker.Call(""))
    pickerDialog.OnEvent("Close", (*) => finishPicker.Call(""))

    ModelPickerPopulate(modelListBox, modelStatus, modelAddButton, catalog, existingModels)
    pickerHotIf := "ahk_id " pickerDialog.Hwnd
    pickerArrowHotkeys := Map("$Up", "Up", "$Down", "Down", "$Left", "Left", "$Right", "Right")
    HotIfWinActive(pickerHotIf)
    for keyName, direction in pickerArrowHotkeys
        try Hotkey(keyName, ModelPickerNavigate.Bind(direction, modelListBox, modelAddButton, refreshButton, cancelButton), "On")
    HotIfWinActive()

    try {
        pickerDialog.Show("AutoSize Center")
        CPApplyOwnedDialogTheme(pickerDialog)
        if modelListBox.Enabled
            modelListBox.Focus()
        else
            refreshButton.Focus()
        while !pickerClosed
            Sleep(30)
    } finally {
        HotIfWinActive(pickerHotIf)
        for keyName, direction in pickerArrowHotkeys
            try Hotkey(keyName, "Off")
        HotIfWinActive()
    }
    return pickerResult
}

AddModelInteractive(modelArray, modelKey, modelCombo, provider, purpose) {
    modelSource := AddModelSourceDialog(provider, purpose)
    if (modelSource = "cancel")
        return false
    if (modelSource = "manual")
        return AddModel(modelArray, modelKey, modelCombo)

    modelCatalog := ModelCatalogQueryWithFeedback(provider, purpose)
    if !modelCatalog["ok"] {
        MsgBox("The online model list could not be loaded.`n`n" modelCatalog["error"], "Browse models", 48)
        return false
    }
    selectedModel := OnlineModelPicker(provider, purpose, modelArray, modelCatalog)
    if (selectedModel = "")
        return false
    return AddModelValue(modelArray, modelKey, modelCombo, selectedModel)
}

LaunchOverlay(*) {
    global overlayAhk, imgProvider, imgModel, geminiImgModel, overlayTrans
    global promptsDir, promptProfile, imgPostproc
    global debugMode  ; use the real checkbox state
    ov := ResolvePath(overlayAhk)
    if (ov = "" || !FileExist(ov)) {
        MsgBox("Set a valid Overlay .ahk path.`n`n" ov, "Missing", 48)
        return
    }
    EnvSet("PROVIDER", imgProvider)

    if (imgProvider = "gemini") {
        modelToSet := geminiImgModel
        if (SubStr(modelToSet, 1, 7) != "models/") {
            modelToSet := "models/" . modelToSet
        }
        EnvSet("GEMINI_MODEL_NAME", modelToSet)
    } else {
        EnvSet("MODEL_NAME", imgModel)
    }

    ; Pass prompt profile & post-processing.
    EnvSet("PROMPT_PROFILE", promptProfile)
    EnvSet("PROMPT_FILE", "")
    EnvSet("POSTPROC_MODE", imgPostproc)
	EnvSet "JRPG_DEBUG", (debugMode ? "1" : "0")
    EnvSet("PYTHONIOENCODING","utf-8")

    ; --- FIX: Explicitly clear EXPLAIN_MODE ---
    EnvSet("EXPLAIN_MODE","")

    DbgCP("LaunchOverlay run='" ov "' provider=" imgProvider " model=" (imgProvider="gemini"?geminiImgModel:imgModel) " prompt=" promptProfile " postproc=" imgPostproc " trans=" overlayTrans)

    ; Always tell the overlay where the app root is
    EnvSet("APP_ROOT", A_ScriptDir)

    SplitPath(ov, , &ovDir, &ext)
    runDir := A_ScriptDir  ; force stable working dir at app root

    if (StrLower(ext) = "exe") {
        cmd := Format('"{}" --root "{}"', ov, A_ScriptDir)
    } else {
        exe := A_AhkPath
        cmd := Format('"{}" "{}" --root "{}"', exe, ov, A_ScriptDir)
    }

    pid := 0
    try {
        pid := Run(cmd, runDir)
        DbgCP("LaunchOverlay Run OK, pid=" pid " cmd=" cmd " wd=" runDir)
    } Catch as ex {

        DbgCP("LaunchOverlay Run EXCEPTION: " e.Message "  cmd=" cmd " wd=" ovDir)
    }

    Sleep(150)
    ok := (pid && ProcessExist(pid))
    pids := OverlayPidsByScript()
    DbgCP("OverlayPidsByScript() -> count=" pids.Length (pids.Length ? " first=" pids[1] : ""))

    if (!ok && pids.Length) {
        pid := pids[1]
        ok := true
        DbgCP("LaunchOverlay adopting PID from WMI: " pid)
    }
    
    oldMode := A_TitleMatchMode
    SetTitleMatchMode 3
    WinWait("Translator",, 3)
    if WinExist("Translator") {
        DbgCP("LaunchOverlay found window: Translator")
        try WinSetTransparent(overlayTrans, "Translator")
        Catch as ex {
            DbgCP("WinSetTransparent failed on 'Translator': " e.Message)
        }
        SetTitleMatchMode oldMode
        SendOverlayTheme("Translator")
        return
    }
    SetTitleMatchMode oldMode

    DumpWindowsForPids(pids)
    DbgCP("LaunchOverlay: window not found, running diagnostic with /ErrorStdOut â€¦")
    diag := A_Temp "\jrpg_overlay_diag.txt"
    try FileDelete(diag)
    if (StrLower(ext) = "exe") {
        diagCmd := Format('cmd /c chcp 65001>nul & cd /d "{}" & "{}" --root "{}" 1>"{}" 2>&1', runDir, ov, A_ScriptDir, diag)
    } else {
        diagCmd := Format('cmd /c chcp 65001>nul & cd /d "{}" & "{}" /ErrorStdOut "{}" --root "{}" 1>"{}" 2>&1', runDir, A_AhkPath, ov, A_ScriptDir, diag)
    }
    DbgCP("LaunchOverlay diag cmd=" diagCmd)
    RunWait(diagCmd, , "Hide")
    diagOut := FileExist(diag) ? Trim(FileRead(diag, "UTF-8")) : "(no diag output)"
    DbgCP("LaunchOverlay diag output: " (StrLen(diagOut) ? SubStr(diagOut, 1, 2000) : "(empty)"))
}

CloseAllOverlays(*) {
    pids := OverlayPidsByScript()
    if !pids.Length {
        ToolTip("No overlay processes found.")
        SetTimer(() => ToolTip(""), -1200)
        Return
    }

    count := 0
    for pid in pids {
        try ProcessClose(pid)
        count++
    }
    Toast("Closed " count " overlay process(es).")
    DbgCP("CloseAllOverlays closed " count " processes.")
}

CloseTranslatorOverlay(*) {
    oldMode := A_TitleMatchMode
    SetTitleMatchMode 3
    if WinExist("Translator") {
        WinClose("Translator")
        Toast("Closed Translator overlay.")
        DbgCP("CloseTranslatorOverlay: Window found and closed.")
    } else {
        ToolTip("Translator overlay is not running.")
        SetTimer(() => ToolTip(""), -1200)
        DbgCP("CloseTranslatorOverlay: Window not found.")
    }
    SetTitleMatchMode oldMode
}

CloseExplainerOverlay(*) {
    old := A_TitleMatchMode
    SetTitleMatchMode 3
    if WinExist("Explainer")
        WinClose
    SetTitleMatchMode old
    StopExplainerBoundsWatcher()
}

LaunchExplainerOverlay(*) {
    global overlayAhk, overlayTrans
    global imgProvider, imgModel, geminiImgModel
	global explainProvider, explainOpenAIModel, explainGeminiModel
	global debugMode

    ov := ResolvePath(overlayAhk)
    if (ov = "" || !FileExist(ov)) {
        MsgBox("Set a valid Overlay .ahk path.`n`n" ov, "Missing", 48)
        return
    }

    ; --- Set explain-mode variables ---
    EnvSet("EXPLAIN_MODE", "1")
    EnvSet("PROMPT_PROFILE", "") ; Clear this to avoid confusion
	
	; NEW: propagate debug toggle to the overlay process
    EnvSet "JRPG_DEBUG", (debugMode ? "1" : "0")
	
    ; use selected EXPLAIN profile
    p := ExplainProfilePath(Trim(ddlEPr.Text))
    if FileExist(p)
        EnvSet "EXPLAIN_PROMPT_FILE", p
    else
        EnvSet "EXPLAIN_PROMPT_FILE", ""  ; Python falls back to BASE_PROMPT


    ; Use explainer-specific provider + models
    EnvSet("EXPLAIN_PROVIDER", explainProvider)
    if (explainProvider = "gemini") {
        modelToSet := explainGeminiModel
        if (SubStr(modelToSet, 1, 7) != "models/")
            modelToSet := "models/" . modelToSet
        EnvSet("GEMINI_EXPLAIN_MODEL", modelToSet)
        EnvSet("EXPLAIN_MODEL","")
    } else {
        EnvSet("EXPLAIN_MODEL", explainOpenAIModel)
        EnvSet("GEMINI_EXPLAIN_MODEL","")
    }

    EnvSet("PYTHONIOENCODING","utf-8")
    DbgCP("LaunchExplainerOverlay run='" ov "' provider=" imgProvider " model=" (imgProvider="gemini"?geminiImgModel:imgModel))

   ; Tell overlay where the app root is
    EnvSet("APP_ROOT", A_ScriptDir)

    SplitPath(ov, , , &ext)
    runDir := A_ScriptDir

    if (StrLower(ext) = "exe") {
        cmd := Format('"{}" --root "{}"', ov, A_ScriptDir)
    } else {
        exe := A_AhkPath
        cmd := Format('"{}" "{}" --root "{}"', exe, ov, A_ScriptDir)
    }

    pid := 0
    try pid := Run(cmd, runDir)
    DbgCP("LaunchExplainerOverlay pid=" pid " cmd=" cmd " wd=" runDir)


    oldMode := A_TitleMatchMode
    SetTitleMatchMode 3
    WinWait("Explainer",, 3)
    if WinExist("Explainer") {
        ; apply saved size/pos first
        ApplyExplainerBounds()
        ; then transparency
        try WinSetTransparent(overlayTrans_EW, "Explainer")
        Catch as ex {
            DbgCP("WinSetTransparent failed: " e.Message)
        }
        ; start periodic watcher to persist movement/resizes
        StartExplainerBoundsWatcher()
    }
    SetTitleMatchMode oldMode

    SendOverlayTheme("Explainer")
}

Toast(msg){
    static g := 0
    try g.Destroy()
    g := Gui("+ToolWindow -Caption +AlwaysOnTop +E0x20")
    g.BackColor := "101825"
    g.MarginX := 12, g.MarginY := 8
    g.SetFont("s11", "Segoe UI")
    g.Add("Text", "cWhite", msg)
    g.Show("AutoSize NoActivate x20 y20")
    SetTimer(() => g.Destroy(), -1100)
}

; =========================
; Controller overlay positioning
; =========================
CPFindExactWindow(title, includeHidden := true) {
    oldMode := A_TitleMatchMode
    oldHidden := A_DetectHiddenWindows
    hwnd := 0
    try {
        SetTitleMatchMode 3
        DetectHiddenWindows includeHidden
        hwnd := WinExist(title)
    } finally {
        SetTitleMatchMode oldMode
        DetectHiddenWindows oldHidden
    }
    return hwnd
}

CPEnsureOverlayForAdjustment(title) {
    hwnd := CPFindExactWindow(title)
    if !hwnd {
        if (title = "Translator")
            LaunchOverlay()
        else
            LaunchExplainerOverlay()
        hwnd := CPFindExactWindow(title)
    }
    if hwnd
        ShowWindowNoActivate(hwnd)
    return hwnd
}

CPOverlayAdjustFlag(enable) {
    global CP_OVERLAY_ADJUST_FLAG
    flagDir := A_Temp "\JRPG_Overlay"
    try {
        if enable {
            if !DirExist(flagDir)
                DirCreate(flagDir)
            if !FileExist(CP_OVERLAY_ADJUST_FLAG)
                FileAppend(ProcessExist(), CP_OVERLAY_ADJUST_FLAG, "UTF-8")
        } else if FileExist(CP_OVERLAY_ADJUST_FLAG) {
            FileDelete(CP_OVERLAY_ADJUST_FLAG)
        }
    }
}

CPOverlayAdjustHotIf(*) {
    global CPOverlayAdjustState
    return CPOverlayAdjustState.Has("active") && CPOverlayAdjustState["active"]
}

CPOverlayAdjustConsume(*) {
}

CPRegisterOverlayAdjustHotkeys() {
    global CPOverlayAdjustHotkeysBound
    if CPOverlayAdjustHotkeysBound
        return

    HotIf(CPOverlayAdjustHotIf)
    try Hotkey("$Enter", CPOverlayAdjustConfirm, "On")
    try Hotkey("$NumpadEnter", CPOverlayAdjustConfirm, "On")
    try Hotkey("$Escape", CPOverlayAdjustCancel, "On")
    try Hotkey("*$Left", CPOverlayAdjustArrow.Bind(-18, 0), "On")
    try Hotkey("*$Right", CPOverlayAdjustArrow.Bind(18, 0), "On")
    try Hotkey("*$Up", CPOverlayAdjustArrow.Bind(0, -18), "On")
    try Hotkey("*$Down", CPOverlayAdjustArrow.Bind(0, 18), "On")
    try Hotkey("WheelUp", CPOverlayAdjustConsume, "On")
    try Hotkey("WheelDown", CPOverlayAdjustConsume, "On")
    HotIf()
    CPOverlayAdjustHotkeysBound := true
}

CPJoystickAxis(controllerId, axisName) {
    value := ""
    try value := GetKeyState(controllerId "Joy" axisName)
    if (value = "")
        return ""
    return value + 0.0
}

CPXInputLibrary() {
    static initialized := false, library := ""
    if initialized
        return library

    initialized := true
    for dllName in ["XInput1_4.dll", "XInput1_3.dll", "XInput9_1_0.dll"] {
        module := 0
        try module := DllCall("kernel32\LoadLibraryW", "str", dllName, "ptr")
        if module {
            library := dllName
            break
        }
    }
    return library
}

CPNormalizeXInputAxis(value) {
    return (value >= 0) ? value / 32767.0 : value / 32768.0
}

CPXInputGetState(userIndex) {
    library := CPXInputLibrary()
    if (library = "")
        return false

    stateBuffer := Buffer(16, 0)
    result := 1
    try result := DllCall(library "\XInputGetState"
        , "uint", userIndex, "ptr", stateBuffer.Ptr, "uint")
    if (result != 0)
        return false

    return Map(
        "leftX", CPNormalizeXInputAxis(NumGet(stateBuffer, 8, "short")),
        "leftY", CPNormalizeXInputAxis(NumGet(stateBuffer, 10, "short")),
        "rightX", CPNormalizeXInputAxis(NumGet(stateBuffer, 12, "short")),
        "rightY", CPNormalizeXInputAxis(NumGet(stateBuffer, 14, "short"))
    )
}

CPControllerRightAxes(controllerName) {
    lowerName := StrLower(controllerName)
    if (InStr(lowerName, "dualsense")
     || InStr(lowerName, "dualshock")
     || InStr(lowerName, "wireless controller"))
        return ["Z", "R"]
    return ["R", "U"]
}

CPScanOverlayAdjustControllers() {
    controllers := []

    Loop 4 {
        userIndex := A_Index - 1
        xinputState := CPXInputGetState(userIndex)
        if IsObject(xinputState) {
            controllers.Push(Map(
                "type", "xinput",
                "id", userIndex,
                "name", "XInput controller " A_Index,
                "baseline", xinputState
            ))
        }
    }

    Loop 16 {
        controllerId := A_Index
        controllerName := ""
        try controllerName := Trim(GetKeyState(controllerId "JoyName"))
        if (controllerName = "")
            continue

        rightAxes := CPControllerRightAxes(controllerName)
        baseline := Map()
        for axisName in ["X", "Y", "Z", "R", "U", "V"] {
            value := CPJoystickAxis(controllerId, axisName)
            baseline[axisName] := (value = "") ? 50.0 : value
        }
        controllers.Push(Map(
            "type", "legacy",
            "id", controllerId,
            "name", controllerName,
            "rightX", rightAxes[1],
            "rightY", rightAxes[2],
            "baseline", baseline
        ))
    }
    return controllers
}

CPReadOverlayAdjustAxes(controller) {
    if (controller["type"] = "xinput") {
        current := CPXInputGetState(controller["id"])
        if !IsObject(current)
            return false
        baseline := controller["baseline"]
        return Map(
            "moveX", Max(-1.0, Min(1.0, current["leftX"] - baseline["leftX"])),
            "moveY", Max(-1.0, Min(1.0, baseline["leftY"] - current["leftY"])),
            "sizeX", Max(-1.0, Min(1.0, current["rightX"] - baseline["rightX"])),
            "sizeY", Max(-1.0, Min(1.0, baseline["rightY"] - current["rightY"]))
        )
    }

    return Map(
        "moveX", CPOverlayAdjustAxis(controller, "X"),
        "moveY", CPOverlayAdjustAxis(controller, "Y"),
        "sizeX", CPOverlayAdjustAxis(controller, controller["rightX"]),
        "sizeY", CPOverlayAdjustAxis(controller, controller["rightY"])
    )
}

CPDetectOverlayAdjustController() {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !state["controllers"].Length
        return false

    bestController := 0
    bestScore := 0.0
    for controller in state["controllers"] {
        axes := CPReadOverlayAdjustAxes(controller)
        if !IsObject(axes)
            continue
        score := Max(Abs(axes["moveX"]), Abs(axes["moveY"])
            , Abs(axes["sizeX"]), Abs(axes["sizeY"]))
        if (score > bestScore) {
            bestScore := score
            bestController := controller
        }
    }

    if (!IsObject(bestController) || bestScore < 0.15)
        return false

    state["controller"] := bestController
    CPUpdateOverlayAdjustHud(true)
    return true
}

CPOverlayAdjustAxis(controller, axisName) {
    value := CPJoystickAxis(controller["id"], axisName)
    if (value = "")
        return 0.0
    centered := (value - controller["baseline"][axisName]) / 50.0
    return Max(-1.0, Min(1.0, centered))
}

CPOverlayAdjustVelocity(axisValue, maximumSpeed := 620.0) {
    magnitude := Abs(axisValue)
    deadZone := 0.16
    if (magnitude <= deadZone)
        return 0.0
    normalized := (magnitude - deadZone) / (1.0 - deadZone)
    speed := 12.0 + (maximumSpeed - 12.0) * normalized * normalized

    precisionPoint := 0.15
    halfTiltPoint := (0.5 - deadZone) / (1.0 - deadZone)
    if (normalized <= precisionPoint) {
        boostFactor := 1.0
    } else if (normalized <= halfTiltPoint) {
        boostFactor := 1.0 + 1.25
            * (normalized - precisionPoint) / (halfTiltPoint - precisionPoint)
    } else {
        boostFactor := 2.25 + 0.75
            * (normalized - halfTiltPoint) / (1.0 - halfTiltPoint)
    }
    speed *= boostFactor
    return (axisValue < 0) ? -speed : speed
}

CPOverlayAdjustWholePixels(stateKey, amount) {
    global CPOverlayAdjustState
    total := CPOverlayAdjustState[stateKey] + amount
    whole := (total >= 0) ? Floor(total) : Ceil(total)
    CPOverlayAdjustState[stateKey] := total - whole
    return whole
}

CPClampOverlayAdjustRect(&x, &y, &w, &h) {
    virtualX := DllCall("user32\GetSystemMetrics", "int", 76, "int")
    virtualY := DllCall("user32\GetSystemMetrics", "int", 77, "int")
    virtualW := DllCall("user32\GetSystemMetrics", "int", 78, "int")
    virtualH := DllCall("user32\GetSystemMetrics", "int", 79, "int")
    if (virtualW <= 0 || virtualH <= 0)
        return

    w := Max(200, Min(w, virtualW))
    h := Max(120, Min(h, virtualH))
    x := Max(virtualX, Min(x, virtualX + virtualW - w))
    y := Max(virtualY, Min(y, virtualY + virtualH - h))
}

CPApplyOverlayAdjustRect(x, y, w, h) {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state["active"] && DllCall("user32\IsWindow", "ptr", state["hwnd"], "int"))
        return false

    CPClampOverlayAdjustRect(&x, &y, &w, &h)
    state["x"] := x, state["y"] := y, state["w"] := w, state["h"] := h
    try WinMove(x, y, w, h, "ahk_id " state["hwnd"])
    CPUpdateOverlayAdjustHud()
    return true
}

CPOverlayAdjustNudge(deltaX, deltaY, deltaW, deltaH, *) {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("active") && state["active"])
        return
    CPApplyOverlayAdjustRect(
        state["x"] + deltaX,
        state["y"] + deltaY,
        state["w"] + deltaW,
        state["h"] + deltaH
    )
}

CPOverlayAdjustModifierKey(hotkeyText) {
    keyName := RegExReplace(Trim(hotkeyText), "i)\s+up$")
    if InStr(keyName, " & ") {
        keyParts := StrSplit(keyName, " & ")
        keyName := Trim(keyParts[keyParts.Length])
    }
    return RegExReplace(keyName, "^[~*$<>!^+#]+")
}

CPOverlayAdjustResizeModifierHeld() {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("resizeModifierKey") && state["resizeModifierKey"] != "")
        return false
    isPressed := false
    try isPressed := GetKeyState(state["resizeModifierKey"], "P")
    return !!isPressed
}

CPOverlayAdjustArrow(deltaX, deltaY, *) {
    if CPOverlayAdjustResizeModifierHeld()
        CPOverlayAdjustNudge(0, 0, deltaX, deltaY)
    else
        CPOverlayAdjustNudge(deltaX, deltaY, 0, 0)
}

CPCreateOverlayAdjustHud() {
    global CPOverlayAdjustState, controlDarkMode
    state := CPOverlayAdjustState
    hud := Gui("+ToolWindow -Caption +AlwaysOnTop +E0x20")
    hud.BackColor := controlDarkMode ? "202124" : "F3F3F3"
    hud.MarginX := 12, hud.MarginY := 8
    hud.SetFont("s10 " (controlDarkMode ? "cFFFFFF" : "c202124"), "Segoe UI")
    hudText := hud.Add("Text", "w720 h62 Center", "")
    state["hud"] := hud
    state["hudText"] := hudText
    CPUpdateOverlayAdjustHud(true)
    hud.Show("NA AutoSize")
    CPPositionOverlayAdjustHud()
}

CPPositionOverlayAdjustHud() {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("hud") && state["hud"])
        return

    centerX := state["x"] + state["w"] / 2
    centerY := state["y"] + state["h"] / 2
    workLeft := 0, workTop := 0, workRight := A_ScreenWidth, workBottom := A_ScreenHeight
    try monitorCount := MonitorGetCount()
    catch {
        monitorCount := 0
    }
    Loop monitorCount {
        try MonitorGetWorkArea(A_Index, &left, &top, &right, &bottom)
        catch
            continue
        if (centerX >= left && centerX < right && centerY >= top && centerY < bottom) {
            workLeft := left, workTop := top, workRight := right, workBottom := bottom
            break
        }
    }

    try state["hud"].GetPos(,, &hudW, &hudH)
    catch
        return
    hudX := Round(workLeft + (workRight - workLeft - hudW) / 2)
    hudY := workTop + 16
    try state["hud"].Move(hudX, hudY)
}

CPUpdateOverlayAdjustHud(force := false) {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("hudText") && state["hudText"])
        return
    now := A_TickCount
    if (!force && state.Has("lastHudTick") && now - state["lastHudTick"] < 100)
        return
    state["lastHudTick"] := now

    controllerLine := "Move either analog stick to select a controller"
    if (state.Has("controller") && IsObject(state["controller"]))
        controllerLine := state["controller"]["name"]
    state["hudText"].Text := "Adjusting " state["title"] " | " controllerLine
        . "`nLeft stick or arrows: move | Right stick or hold Screenshot + Translate + arrows: resize"
        . "`nEnter saves | Esc cancels | " state["w"] " x " state["h"]
    CPPositionOverlayAdjustHud()
}

CPOverlayAdjustTick(*) {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("active") && state["active"])
        return
    if !DllCall("user32\IsWindow", "ptr", state["hwnd"], "int") {
        CPFinishOverlayAdjustment(false)
        return
    }

    now := A_TickCount
    deltaSeconds := Min(0.05, Max(0.001, (now - state["lastTick"]) / 1000.0))
    state["lastTick"] := now

    if !(state.Has("controller") && IsObject(state["controller"])) {
        if (now - state["lastScanTick"] >= 1000) {
            state["controllers"] := CPScanOverlayAdjustControllers()
            state["lastScanTick"] := now
        }
        CPDetectOverlayAdjustController()
        CPUpdateOverlayAdjustHud()
        return
    }

    controller := state["controller"]
    axes := CPReadOverlayAdjustAxes(controller)
    if !IsObject(axes) {
        state.Delete("controller")
        state["controllers"] := CPScanOverlayAdjustControllers()
        state["lastScanTick"] := now
        CPUpdateOverlayAdjustHud(true)
        return
    }
    moveX := CPOverlayAdjustVelocity(axes["moveX"], 620.0)
    moveY := CPOverlayAdjustVelocity(axes["moveY"], 620.0)
    sizeX := CPOverlayAdjustVelocity(axes["sizeX"], 520.0)
    sizeY := CPOverlayAdjustVelocity(axes["sizeY"], 520.0)

    deltaX := CPOverlayAdjustWholePixels("fractionX", moveX * deltaSeconds)
    deltaY := CPOverlayAdjustWholePixels("fractionY", moveY * deltaSeconds)
    deltaW := CPOverlayAdjustWholePixels("fractionW", sizeX * deltaSeconds)
    deltaH := CPOverlayAdjustWholePixels("fractionH", sizeY * deltaSeconds)
    if (deltaX || deltaY || deltaW || deltaH)
        CPApplyOverlayAdjustRect(state["x"] + deltaX, state["y"] + deltaY
            , state["w"] + deltaW, state["h"] + deltaH)
}

StartOverlayAdjustment(title, *) {
    global ui, CPOverlayAdjustState, CPPreviousForegroundHwnd, iniPath
    state := CPOverlayAdjustState
    if (state.Has("active") && state["active"])
        return

    hwnd := CPEnsureOverlayForAdjustment(title)
    if !hwnd {
        MsgBox("The " title " overlay could not be opened.", "Move / Resize", 48)
        return
    }

    x := 0, y := 0, w := 0, h := 0
    try WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
    if (w <= 0 || h <= 0) {
        MsgBox("The " title " overlay bounds could not be read.", "Move / Resize", 48)
        return
    }

    state.Clear()
    state["active"] := true
    state["title"] := title
    state["hwnd"] := hwnd
    state["originalX"] := x, state["originalY"] := y
    state["originalW"] := w, state["originalH"] := h
    state["x"] := x, state["y"] := y, state["w"] := w, state["h"] := h
    state["wasTopmost"] := !!(WinGetExStyle("ahk_id " hwnd) & 0x00000008)
    state["returnHwnd"] := CPPreviousForegroundHwnd
    screenshotHotkey := Trim(IniRead(iniPath, "hotkeys", "screenshot_translate", "^+t"))
    state["resizeModifierKey"] := CPOverlayAdjustModifierKey(screenshotHotkey)
    state["controllers"] := CPScanOverlayAdjustControllers()
    state["lastScanTick"] := A_TickCount
    state["lastTick"] := A_TickCount
    state["acceptAfter"] := A_TickCount + 400
    state["fractionX"] := 0.0, state["fractionY"] := 0.0
    state["fractionW"] := 0.0, state["fractionH"] := 0.0
    state["lastHudTick"] := 0

    CPOverlayAdjustFlag(true)
    CPRegisterOverlayAdjustHotkeys()
    DllCall("user32\SetWindowPos", "ptr", hwnd, "ptr", -1
        , "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x0013)
    CPCreateOverlayAdjustHud()

    SavePanelBounds()
    ui.Hide()
    SetTimer(RestoreControlPanelReturnWindow, -1)
    SetTimer(CPOverlayAdjustTick, 20)
}

CPOverlayAdjustConfirm(*) {
    global CPOverlayAdjustState
    if (A_TickCount < CPOverlayAdjustState["acceptAfter"])
        return
    CPFinishOverlayAdjustment(true)
}

CPOverlayAdjustCancel(*) {
    CPFinishOverlayAdjustment(false)
}

CPRestoreControlPanelAfterAdjustment(title, returnHwnd := 0) {
    global ui, CPPreviousForegroundHwnd, btnMoveResize, btnMoveResize_EW
    if !(IsSet(ui) && ui && ui.Hwnd)
        return

    CPPreviousForegroundHwnd := returnHwnd
    ui.Show()
    try WinActivate("ahk_id " ui.Hwnd)
    try {
        if (title = "Translator")
            btnMoveResize.Focus()
        else
            btnMoveResize_EW.Focus()
    }
}

CPFinishOverlayAdjustment(saveChanges, quiet := false) {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("active") && state["active"])
        return

    SetTimer(CPOverlayAdjustTick, 0)
    title := state["title"]
    hwnd := state["hwnd"]
    if (!saveChanges && DllCall("user32\IsWindow", "ptr", hwnd, "int"))
        try WinMove(state["originalX"], state["originalY"], state["originalW"], state["originalH"], "ahk_id " hwnd)

    if DllCall("user32\IsWindow", "ptr", hwnd, "int") {
        if !state["wasTopmost"]
            DllCall("user32\SetWindowPos", "ptr", hwnd, "ptr", -2
                , "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x0013)
        SendOverlayCmdTo(title, "action=save_bounds")
        if (title = "Explainer")
            SetTimer(SaveExplainerBoundsIfChanged, -1)
    }

    try state["hud"].Destroy()
    state["active"] := false
    CPOverlayAdjustFlag(false)
    if !quiet {
        returnHwnd := state.Has("returnHwnd") ? state["returnHwnd"] : 0
        SetTimer(CPRestoreControlPanelAfterAdjustment.Bind(title, returnHwnd), -100)
        Toast(title (saveChanges ? " position saved" : " adjustment canceled"))
    }
}

CPOverlayAdjustOnExit(*) {
    global CPOverlayAdjustState
    if (CPOverlayAdjustState.Has("active") && CPOverlayAdjustState["active"])
        CPFinishOverlayAdjustment(false, true)
    else
        CPOverlayAdjustFlag(false)
}

; =========================
; GUI
; =========================
ui := Gui("+Resize +MinSize" CP_VIEWPORT_MIN_W "x" CP_VIEWPORT_MIN_H " +0x300000", "JRPG Translator")
CPRegisterThemeMessages()
CPRegisterCanvasMessages()

; --- Control Panel default bounds (used only if no valid [gui_bounds] exist) ---
defGuiX := 140
defGuiY := 140
defGuiW := CP_CANVAS_MIN_W
defGuiH := CP_CANVAS_MIN_H

IsValidBounds(x, y, w, h) {
    if !((x is number) && (y is number) && (w is number) && (h is number))
        return false
    if (w < 640 || h < 480)
        return false
    if (x <= -30000 || y <= -30000)
        return false

    try {
        monitorCount := MonitorGetCount()
        Loop monitorCount {
            MonitorGetWorkArea(A_Index, &left, &top, &right, &bottom)
            if (x < right && x + w > left && y < bottom && y + h > top)
                return true
        }
    }
    return false
}

ui.MarginX := pad, ui.MarginY := pad
ui.SetFont("s10", "Segoe UI")
ui.BackColor := CPPalette(controlDarkMode)["window"]

; The native tab remains as a page host and focus proxy. A custom tab bar is
; created after all page controls so its styling and geometry are predictable.
tabNames := ["Screenshot Translation","Audio Translation","Translation Window","Explanation","Explanation Window","Terminology Overrides","Hotkeys","API Keys","Paths"]
CPTabVisiblePages := [1, 2, 3, 4, 5, 6, 7, 8]
if showPathsTab
    CPTabVisiblePages.Push(9)
tab := ui.Add("Tab", "xm ym w760 h420 Buttons", tabNames)

; --- Tab 1: SCREENSHOT TRANSLATION
tab.UseTab(1)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer
ui.Add("Text", "xm y+6 w90", "AI Provider:")
ddlProv := ui.Add("DropDownList", "x+m w220 0x210", ["Gemini","OpenAI"])
provSelIdx := (StrLower(imgProvider) = "gemini") ? 1 : 2
ddlProv.Choose(provSelIdx)
ddlProv.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))

ui.Add("Text", "xm y+12 w90", "Gemini model:")
ddlIMG_GM := ui.Add("DropDownList", "x+m w260 0x210", model_gemini_img)
imgGMInitIdx := ArrIndexOf(model_gemini_img, geminiImgModel)
ddlIMG_GM.Choose(imgGMInitIdx ? imgGMInitIdx : 1)
ddlIMG_GM.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))
btnIMG_GM_Add := ui.Add("Button", "x+82 w70", "Add")
btnIMG_GM_Del := ui.Add("Button", "x+6 w70", "Delete")

ui.Add("Text", "xm y+12 w90", "OpenAI model:")
ddlIMG := ui.Add("DropDownList", "x+m w260 0x210", model_openai_img)
imgInitIdx := ArrIndexOf(model_openai_img, imgModel)
ddlIMG.Choose(imgInitIdx ? imgInitIdx : 1)
ddlIMG.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))
btnIMG_Add := ui.Add("Button", "x+82 w70", "Add")
btnIMG_Del := ui.Add("Button", "x+6 w70", "Delete")

; Prompt profile (FIRST)
ui.Add("Text", "xm y+12 w90", "Prompt:")
ddlPrompt := ui.Add("DropDownList", "x+m w260 0x210", ListPromptProfiles())
btnPrEdit  := ui.Add("Button", "x+6 w70", "Edit")
btnPrNew   := ui.Add("Button", "x+6 w70", "Add")
btnPrDel   := ui.Add("Button", "x+6 w70", "Delete")

ui.Add("Text", "xm y+12", "Translation post-processing:")
postLabels := ["Translation with transcript","Translation only","Direct model output"]
postCodes  := ["tt","translation","none"]

; AltSubmit => .Value returns 1..N (index into postCodes)
ddlPost := ui.Add("DropDownList", "x+m w260 AltSubmit 0x210", postLabels)

; Initialize selection from saved code (imgPostproc)
postInitIdx := ArrIndexOf(postCodes, imgPostproc)
if (!postInitIdx)
    postInitIdx := 1
ddlPost.Value := postInitIdx

; Toggle: delete screenshots after translation
delAfterUse := Integer(IniRead(iniPath, "paths", "deleteAfterUse", 0))
chkDel := ui.Add("Checkbox", "xm y+10", "Delete screenshots after translation")
chkDel.Value := delAfterUse ? 1 : 0
; Persist to control.ini immediately when toggled
chkDel.OnEvent("Click", (*) => IniWrite(chkDel.Value ? 1 : 0, iniPath, "paths", "deleteAfterUse"))

; Toggle: highlight guessed subjects (shifted right to avoid size box overlap)
hlGuess := Integer(IniRead(iniPath, "cfg", "highlightGuessed", 1))
chkGuess := ui.Add("Checkbox", "x+240 yp", "Highlight guessed subjects")
chkGuess.Value := hlGuess ? 1 : 0
chkGuess.OnEvent("Click", (*) => (IniWrite(chkGuess.Value ? 1 : 0, iniPath, "cfg", "highlightGuessed"), ApplyShotSettings()))

; Help text under â€œHighlight guessed subjectsâ€ (start under the word, not under the checkbox box)
chkGuess.GetPos(&gx, &gy, &gWidth, &gHeight)
cbIndent := 22  ; ~checkbox box width + label gap
txtGuessHelp := ui.Add(
    "Text"
  , Format("x{} y+2 w420 cGray", gx + cbIndent)  ; initial width; will be resized on window Size
  , "When enabled the subjects or pronouns the model adds for natural English phrasing are shown in italics for clarity."
)
CPRegisterMutedControl(txtGuessHelp)

; Toggle: use color for speaker names (one switch for JP+EN) â€” place lower to leave space for the help text
hlName := Integer(IniRead(iniPath, "cfg", "colorSpeaker", 1))
txtGuessHelp.GetPos(, , , &gHelpH)
chkName := ui.Add("Checkbox", Format("x{} y{}", gx, gy + gHeight + 8 + gHelpH + 6), "Use speaker name color")
chkName.Value := hlName ? 1 : 0
chkName.OnEvent("Click", (*) => (IniWrite(chkName.Value ? 1 : 0, iniPath, "cfg", "colorSpeaker"), ApplyShotSettings()))

; Help text under â€œUse speaker name colorâ€ (start under the word, not under the checkbox box)
chkName.GetPos(&nx, &ny, &nWidth, &nHeight)
txtNameHelp := ui.Add(
    "Text"
  , Format("x{} y+2 w420 cGray", nx + cbIndent)  ; initial width; will be resized on window Size
  , "When enabled, detected speaker names are shown in color picked in Translation Window tab. Turn off for plain output."
)
CPRegisterMutedControl(txtNameHelp)

; Make changes effective immediately + persist to INI
ddlProv.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))
ddlIMG.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))
ddlIMG_GM.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))
ddlPrompt.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))
ddlPost.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))

; --- Max size + Capture picker (native path, non-breaking) ---
; Pin this row to the left-column baseline under the "Delete screenshots after translation" checkbox
chkDel.GetPos(&delX, &delY, , &delH)
leftBaseY := delY + delH + 16  ; spacing under the delete checkbox

ui.Add("Text", Format("xm y{} w160", leftBaseY), "Max PNG size (KB):")
eCapMax := ui.Add("Edit", "x+m yp w80 Number", capMaxKB)
eCapMax.OnEvent("Change", (*) => SetCapMaxKB(eCapMax.Value))

btnCapPick := ui.Add("Button", "xm y+10 w160", "Capture...")
btnCapPick.OnEvent("Click", OpenCapturePicker)

; --- Quick actions: trigger the same functions as the corresponding hotkeys ---
; They simply send the currently configured hotkey from [hotkeys] in control.ini
sp1 := ui.Add("Text", "xm y+8", "")  ; small spacer below Capture

; One-click workflow (standalone)
btnST := ui.Add("Button", "xm y+6 w200 h28", "Screenshot + Translate")
btnST.OnEvent("Click", (*) => FireHotkeyAction("screenshot_translate"))

; Visual separator: etched vertical line (SS_ETCHEDVERT = 0x11)
sepQuick := ui.Add("Text", "x+14 w2 h28 0x11", "")  ; vertical rule + a bit more spacing

; Two-step workflow (used together)
btnTS := ui.Add("Button", "x+14 w170 h28", "Take Screenshot")
btnTS.OnEvent("Click", (*) => FireHotkeyAction("take_screenshot"))

btnSTO := ui.Add("Button", "x+8 w220 h28", "Screenshot -> Translation")
btnSTO.OnEvent("Click", (*) => FireHotkeyAction("screenshot_translation"))

; Subtle hint to clarify intent (auto-wraps with window width)
txtHint := ui.Add(
    "Text"
  , "xm y+6 w620 cGray"  ; give it an initial width so wrapping can happen
  , 'Tip: "Screenshot + Translate" is a one-click action. The other two are a 2-step workflow, allowing multiple screenshots to be translated at once, useful if a longer Japanese sentence did not fit into a single textbox, if ordered in the prompt the AI model can stitch those together.'
)
CPRegisterMutedControl(txtHint)

; Keep the hint and the two help texts wrapping nicely when the window is resized
; v2 Size event passes (gui, minMax, w, h)
ui.OnEvent("Size", (gui, minMax, w, h) => (
    ; keep ~20px margins on both sides for the big tip
    txtHint.Move(, , Max(260, w - 40))
  , (IsSet(txtGuessHelp) ? (
        txtGuessHelp.GetPos(&tgx,, ,)
      , txtGuessHelp.Move(, , Max(240, w - tgx - 40))  ; width = window width minus left x minus right margin
    ) : 0)
  , (IsSet(txtNameHelp) ? (
        txtNameHelp.GetPos(&tnx,, ,)
      , txtNameHelp.Move(, , Max(240, w - tnx - 40))
    ) : 0)
))

; Toggle: open Translator overlay when Control Panel opens
autoOpenTW := Integer(IniRead(iniPath, "cfg", "openTranslatorOnLaunch", 0))
chkOpenTW := ui.Add("CheckBox", "xm y+8", "Open translation window with JRPG Translator")
chkOpenTW.Value := autoOpenTW ? 1 : 0
chkOpenTW.OnEvent("Click", (*) => IniWrite(chkOpenTW.Value ? 1 : 0, iniPath, "cfg", "openTranslatorOnLaunch"))

; Translator window "Always on top" toggle (persists to [cfg])
chkTop_TW := ui.Add("CheckBox", "xm y+12", "Open translation window always on top")
chkTop_TW.Value := Integer(IniRead(iniPath, "cfg", "winTop", 1)) ? 1 : 0
chkTop_TW.OnEvent("Click", (*) => IniWrite(chkTop_TW.Value ? 1 : 0, iniPath, "cfg", "winTop"))


; --- Tab 2: AUDIO TRANSLATION
tab.UseTab(2)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer

; ===== Live input =====
lblLiveInput := ui.Add("Text", "xm y+8 w150", "Live input")
lblLiveInput.SetFont("Bold")

; Listen device (WASAPI loopback target)
ui.Add("Text", "xm y+10", "Listen device")
ddlSpeaker := ui.Add("DropDownList", "x+m w360 0x210", [])
btnSpRef   := ui.Add("Button", "x+6 w80", "Refresh")

; Live translation provider
lblLiveTranslation := ui.Add("Text", "xm y+20 w180", "Live translation")
lblLiveTranslation.SetFont("Bold")
ui.Add("Text", "xm y+10", "AI Provider:")
ddlAProv := ui.Add("DropDownList", "x+m w220 0x210", ["Gemini","OpenAI"])
ddlAProv.Text := audioProvider
; Keep model dropdowns in sync with provider choice
ddlAProv.OnEvent("Change", (*) => (ToggleAudioControls(), AutoPersist()))

; Gemini audio model â€” its own row directly under provider
ui.Add("Text", "xm y+12", "Gemini live model:")
ddlA_GM := ui.Add("DropDownList", "x+m w260 0x210", model_gemini_audio)
SetComboToExistingItem(ddlA_GM, model_gemini_audio, geminiAudioModel)
btnA_GM_Add := ui.Add("Button", "x+6 w60", "Add...")
btnA_GM_Del := ui.Add("Button", "x+6 w60", "Delete")

ui.Add("Text", "xm y+12", "OpenAI live model:")
ddlTR := ui.Add("DropDownList", "x+m w420 0x210", model_openai_audio) ; initial width; ResizeUI will adjust
SetComboToExistingItem(ddlTR, model_openai_audio, trModel)
btnTR_Add := ui.Add("Button", "x+6 w60", "Add...")
btnTR_Del := ui.Add("Button", "x+6 w60", "Delete")

; Ensure correct initial enabled/disabled state based on provider
ToggleAudioControls()

; Output language for live audio translation
ui.Add("Text", "xm y+12 w150", "Output language:")
ddlAudioTarget := ui.Add("DropDownList", "x+m w260 0x210", audioTargetLangs)
ddlAudioTarget.Text := audioTargetLang
ddlAudioTarget.OnEvent("Change", (*) => AutoPersist())

; fill and wire the device dropdown
PopulateSpeakersList(speakerName)
btnSpRef.OnEvent("Click", RefreshSpeakerList)
ddlSpeaker.OnEvent("Change", SpeakerChanged)

; --- Tab 3: TRANSLATION WINDOW
tab.UseTab(3)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer
twLabelX := pad + 16
twLabelW := 140
twControlW := 280
twSwatchW := 84
twSwatchX := twLabelX + twLabelW + pad + twControlW - twSwatchW

ui.Add("Text", "x" twLabelX " y+10 w" twLabelW, "Overlay Transparency:")
slTrans := ui.Add("Slider", "x+m w" twControlW " Range0-255 ToolTip")
lblTransPct := ui.Add("Text", "x+m", "100%")

ui.Add("Text", "x" twLabelX " y+28 w" twLabelW, "Window color:")
rectBg := CPRegisterColorSwatch(ui.Add("Text", "x" twSwatchX " yp w" twSwatchW " h34 Border"), "translator:bg")
ui.Add("Text", "x" twLabelX " y+18 w" twLabelW, "Text color:")
rectTxt := CPRegisterColorSwatch(ui.Add("Text", "x" twSwatchX " yp w" twSwatchW " h34 Border"), "translator:txt")
ui.Add("Text", "x" twLabelX " y+18 w" twLabelW, "Speaker name:")
rectName := CPRegisterColorSwatch(ui.Add("Text", "x" twSwatchX " yp w" twSwatchW " h34 Border"), "translator:name")

RefreshColorSwatches()

ui.Add("Text", "x" twLabelX " y+32 w" twLabelW, "Font:")
ddlFont := ui.Add("DropDownList", "x+m w" twControlW " 0x210", [])
ui.Add("Text", "x+14 yp", "Size:")
edFSize := ui.Add("Edit", "x+m w60 Number", fontSize)
udFSize := ui.Add("UpDown", "Range6-128", fontSize)
chkFontBold := ui.Add("CheckBox", "x+14 yp+2", "Bold")
chkFontBold.Value := fontBold

ui.Add("Text", "x" twLabelX " y+26 w" twLabelW, "Profile:")
ddlProf  := ui.Add("DropDownList", "x+m w" twControlW " 0x210", [])
btnPNew  := ui.Add("Button", "x+8 w80",  "Add")
btnPSave := ui.Add("Button", "x+6 w70",  "Save")
btnPLoad := ui.Add("Button", "x+6 w70",  "Load")
btnPDel  := ui.Add("Button", "x+6 w70",  "Delete")

twControlX := twLabelX + twLabelW + pad
btnMoveResize := ui.Add("Button", "x" twControlX " y+20 w180", "Move / Resize")
txtMoveResize := ui.Add("Text", "x" twControlX " y+5 w590 cGray"
    , "Left stick or arrows move; right stick or Screenshot + Translate + arrows resize. Enter saves, Esc cancels.")
CPRegisterMutedControl(txtMoveResize)
btnMoveResize.OnEvent("Click", StartOverlayAdjustment.Bind("Translator"))

RefreshProfilesList( IniRead(iniPath, "profiles", "translator_last", "") )

btnPSave.OnEvent("Click", (*) => (
    name := Trim(ddlProf.Text),
    name := (name = "" ? "default" : name),
    SaveProfile(name),
    RefreshProfilesList(name)
))
btnPLoad.OnEvent("Click", (*) => (
    name := Trim(ddlProf.Text),
    name != "" ? LoadProfile(name) : 0
))
btnPDel.OnEvent("Click", (*) => (
    name := Trim(ddlProf.Text),
    (name != "" && MsgBox("Delete profile '" name "'?",, 0x21) = "OK")
        ? (DeleteProfile(name), RefreshProfilesList()) : 0
))
btnPNew.OnEvent("Click", (*) => CreateProfile())

for sw in [rectBg,rectTxt,rectName]
    sw.Cursor := "Hand"

rectBg.OnEvent("Click", (*) => PickAndApply("bg"))
rectTxt.OnEvent("Click", (*) => PickAndApply("txt"))
rectName.OnEvent("Click", (*) => PickAndApply("name"))

ddlFont.OnEvent("Change", FontChanged)
edFSize.OnEvent("LoseFocus", FontSizeCommit)
udFSize.OnEvent("Change", (*) => FontSizeCommit(edFSize))
chkFontBold.OnEvent("Click", FontBoldChanged)

; Prompt profile events + initial list
btnPrEdit.OnEvent("Click", OpenPromptEditor)
btnPrNew.OnEvent("Click",  NewPromptProfile)
btnPrDel.OnEvent("Click",  DeletePromptProfile)
RefreshPromptProfilesList(promptProfile)


; --- Explanation: Provider + Models (independent from Screenshot Translation)
tab.UseTab(4)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer
ui.Add("Text", "xm y+6 w90", "AI Provider:")
ddlEProv := ui.Add("DropDownList", "x+m w300 0x210", ["Gemini","OpenAI"])

eProvIdx := (StrLower(explainProvider) = "gemini") ? 1 : 2
ddlEProv.Choose(eProvIdx)
ddlEProv.OnEvent("Change", (*) => (UpdateVars(), SaveAll()))

; --- Gemini row (unchanged)
ui.Add("Text", "xm y+12 w90", "Gemini model:")
ddlEGem := ui.Add("DropDownList", "x+m w300 0x210", model_gemini_explain)
eGemIdx := ArrIndexOf(model_gemini_explain, explainGeminiModel)
ddlEGem.Choose(eGemIdx ? eGemIdx : 1)
ddlEGem.OnEvent("Change", (*) => (UpdateVars(), SaveAll()))
btnEGem_Add := ui.Add("Button", "x+82 w70", "Add")
btnEGem_Del := ui.Add("Button", "x+6 w70", "Delete")

; --- OpenAI row (moved here; use a fresh y step so it sits below Gemini)
ui.Add("Text", "xm y+12 w90", "OpenAI model:")
ddlEOpenAI := ui.Add("DropDownList", "x+m w300 0x210", model_openai_explain)
eOpenAIIdx := ArrIndexOf(model_openai_explain, explainOpenAIModel)
ddlEOpenAI.Choose(eOpenAIIdx ? eOpenAIIdx : 1)
ddlEOpenAI.OnEvent("Change", (*) => (UpdateVars(), SaveAll()))
btnEOpenAI_Add := ui.Add("Button", "x+82 w70", "Add")
btnEOpenAI_Del := ui.Add("Button", "x+6 w70", "Delete")


; Initialize enabled/disabled state for Explanation models
ToggleExplanationControls()
; Force a post-build sync from INI so later generic repainting canâ€™t overwrite these
SyncExplanationFromIni()

; EXPLANATION prompt profile (independent from Screenshot/Audio prompts)
ui.Add("Text", "xm y+10 w90", "Prompt:")
ddlEPr     := ui.Add("DropDownList", "x+m w300 0x210", [])
btnEPrEdit := ui.Add("Button", "x+6 w70", "Edit")
btnEPrNew  := ui.Add("Button", "x+6 w70", "Add")
btnEPrDel  := ui.Add("Button", "x+6 w70", "Delete")

; anchor to the current Sectionâ€™s left edge, keep same row spacing
btnExplainNow := ui.Add("Button", "xs y+20 w220", "Explain last jp. Text")

; Move: Save explanations checkbox â€” placed under the button row, left-aligned to the Section
saveExplChk := ui.Add("CheckBox", "xs y+16", "Save explanations to textfiles")
; Short info note about where the files are stored
txtExplainSaveInfo := ui.Add("Text"
    , "xm y+4 w720 cGray"
    , "Saved explanations are stored in the 'Settings\\Explanations' folder inside your JRPG Translator directory."
)
; Load previous setting (defaults to 0 if missing)
CPRegisterMutedControl(txtExplainSaveInfo)
saveExplVal := IniRead(iniPath, "cfg", "saveExplains", 0)
saveExplChk.Value := saveExplVal ? 1 : 0
; Persist on click
saveExplChk.OnEvent("Click", (*) => (
    IniWrite(saveExplChk.Value ? 1 : 0, iniPath, "cfg", "saveExplains")
))

; Checkbox on the NEXT line, left-aligned under the first button
autoOpenEW := Integer(IniRead(iniPath, "cfg", "openExplainerOnLaunch", 0))
chkOpenEW  := ui.Add("CheckBox", "xs y+10", "Open explanation window with JRPG Translator")
chkOpenEW.Value := autoOpenEW ? 1 : 0
chkOpenEW.OnEvent("Click", (*) => IniWrite(chkOpenEW.Value ? 1 : 0, iniPath, "cfg", "openExplainerOnLaunch"))

; Explanation window "Always on top" toggle (persists to [cfg_explainer])
chkTop_EW := ui.Add("CheckBox", "xm y+12", "Open explanation window always on top")
chkTop_EW.Value := Integer(IniRead(iniPath, "cfg_explainer", "winTop", 0)) ? 1 : 0
chkTop_EW.OnEvent("Click", (*) => IniWrite(chkTop_EW.Value ? 1 : 0, iniPath, "cfg_explainer", "winTop"))

btnEPrEdit.OnEvent("Click", OpenExplainPromptEditor_Multi)
btnEPrNew.OnEvent("Click",  NewExplainPromptProfile)
btnEPrDel.OnEvent("Click",  DeleteExplainPromptProfile)
ddlEPr.OnEvent("Change", ExplainPromptChanged)
RefreshExplainPromptProfilesList(explainPromptProfile)

btnExplainNow .OnEvent("Click", ExplainNow)

btnEOpenAI_Add.OnEvent("Click", (*) => AddModelInteractive(model_openai_explain, "openai_explain", ddlEOpenAI, "openai", "explanation"))
btnEOpenAI_Del.OnEvent("Click", (*) => DeleteModel(model_openai_explain, "openai_explain", ddlEOpenAI))

btnEGem_Add.OnEvent("Click", (*) => AddModelInteractive(model_gemini_explain, "gemini_explain", ddlEGem, "gemini", "explanation"))
btnEGem_Del.OnEvent("Click", (*) => DeleteModel(model_gemini_explain, "gemini_explain", ddlEGem))


; --- Tab 5: EXPLANATION WINDOW  (UI only, not wired yet)
tab.UseTab(5)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer
; Layout parity with "Translation Window" (Tab 3), distinct control names (EW_*)
ewLabelX := twLabelX
ewLabelW := twLabelW
ewControlW := twControlW
ewSwatchW := twSwatchW
ewSwatchX := twSwatchX

ui.Add("Text", "x" ewLabelX " y+10 w" ewLabelW, "Overlay Transparency:")
slTrans_EW := ui.Add("Slider", "x+m w" ewControlW " Range0-255 ToolTip")
lblTransPct_EW := ui.Add("Text", "x+m", "100%")

ui.Add("Text", "x" ewLabelX " y+28 w" ewLabelW, "Window color:")
rectBg_EW := CPRegisterColorSwatch(ui.Add("Text", "x" ewSwatchX " yp w" ewSwatchW " h34 Border"), "explainer:bg")
ui.Add("Text", "x" ewLabelX " y+18 w" ewLabelW, "Text color:")
rectTxt_EW := CPRegisterColorSwatch(ui.Add("Text", "x" ewSwatchX " yp w" ewSwatchW " h34 Border"), "explainer:txt")

RefreshColorSwatches_EW()

ui.Add("Text", "x" ewLabelX " y+32 w" ewLabelW, "Font:")
ddlFont_EW := ui.Add("DropDownList", "x+m w" ewControlW " 0x210", [])
ui.Add("Text", "x+14 yp", "Size:")
edFSize_EW := ui.Add("Edit", "x+m w60 Number")
udFSize_EW := ui.Add("UpDown", "Range8-96")
chkFontBold_EW := ui.Add("CheckBox", "x+14 yp+2", "Bold")
chkFontBold_EW.Value := fontBold_EW

; --- Explanation Profiles row ---
ui.Add("Text", "x" ewLabelX " y+26 w" ewLabelW, "Profile:")
ddlProf_EW := ui.Add("DropDownList", "x+m w" ewControlW " 0x210", [])
btnProfCreate_EW := ui.Add("Button", "x+8 w80", "Add")
btnProfSave_EW := ui.Add("Button", "x+6 w70", "Save")
btnProfLoad_EW := ui.Add("Button", "x+6 w70", "Load")
btnProfDel_EW  := ui.Add("Button", "x+6 w70", "Delete")

ewControlX := ewLabelX + ewLabelW + pad
btnMoveResize_EW := ui.Add("Button", "x" ewControlX " y+20 w180", "Move / Resize")
txtMoveResize_EW := ui.Add("Text", "x" ewControlX " y+5 w590 cGray"
    , "Left stick or arrows move; right stick or Screenshot + Translate + arrows resize. Enter saves, Esc cancels.")
CPRegisterMutedControl(txtMoveResize_EW)
btnMoveResize_EW.OnEvent("Click", StartOverlayAdjustment.Bind("Explainer"))

; populate dropdown
Refresh_ddlProf_EW(*) {
    global iniPath, ddlProf_EW
    list := EW_ListProfiles()
    ddlProf_EW.Delete()
    ; AHK v2: Add() expects an Array. Add all at once if we have items.
    if (list.Length)
        ddlProf_EW.Add(list)
    sel := IniRead(iniPath, "profiles", "explainer_last", "")
    if (sel != "") {
        idx := ArrayIndexOf(list, sel)
        ddlProf_EW.Choose(idx ? idx : 1)
    } else if (list.Length) {
        ddlProf_EW.Choose(1)
    }
}
Refresh_ddlProf_EW()

; wire buttons
CreateProfile_EW(*) {
    global ddlProf_EW
    ; Always prompt for a name (same behavior as Translation Window)
    ip := InputBox("Enter new profile name:", "Create Explainer Profile",, "")
    if (ip.Result != "OK")
        return
    name := Trim(ip.Value)
    if (name = "")
        return

    path := EW_ProfilePath(name)
    if FileExist(path) {
        r := MsgBox(Format('Profile "{}" already exists. Overwrite?', name), "Create Explainer Profile", "YesNo Icon!")
        if (r != "Yes")
            return
    }

    EW_SaveProfile(name)
    Refresh_ddlProf_EW()
    ddlProf_EW.Text := name
}


btnProfCreate_EW.OnEvent("Click", CreateProfile_EW)

btnProfSave_EW.OnEvent("Click", (*) => (
    (name := Trim(ddlProf_EW.Text)) = "" 
        ? MsgBox("Enter a profile name in the box.") 
        : (EW_SaveProfile(name), Refresh_ddlProf_EW(), ddlProf_EW.Text := name)
))
btnProfLoad_EW.OnEvent("Click", (*) => (
    (name := Trim(ddlProf_EW.Text)) = "" ? MsgBox("Pick a profile to load.") : EW_LoadProfile(name)
))
btnProfDel_EW.OnEvent("Click", (*) => (
    (name := Trim(ddlProf_EW.Text)) = "" ? MsgBox("Pick a profile to delete.") : (EW_DeleteProfile(name), Refresh_ddlProf_EW())
))

; --- wire EW events ---
for sw in [rectBg_EW,rectTxt_EW]
    sw.Cursor := "Hand"

slTrans_EW.OnEvent("Change", (c, e) => (HandleTransparencyChange_EW(c), SaveAll(), SendOverlayTheme()))

rectBg_EW.OnEvent("Click", (*) => PickAndApply_EW("bg"))
rectTxt_EW.OnEvent("Click", (*) => PickAndApply_EW("txt"))

ddlFont_EW.OnEvent("Change", FontChanged_EW)
edFSize_EW.OnEvent("LoseFocus", FontSizeCommit_EW)
udFSize_EW.OnEvent("Change",   (*) => FontSizeCommit_EW(edFSize_EW))
chkFontBold_EW.OnEvent("Click", FontBoldChanged_EW)

; --- Tab 6: TERMINOLOGY OVERRIDES
tab.UseTab(6)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer

; --- Help text (what these two glossaries do & how to use them)
txtGlossaryHelp1 := ui.Add("Text", "xm y+8 cGray w760"
  , 'Japanese terms often have multiple translations (e.g., 宰相 can mean "Chancellor" or "Prime Minister"), and names may vary in spelling. Set fixed rules here to ensure a consistent translation for your chosen Target Language (TL) throughout your playthrough.')
txtGlossaryHelp2 := ui.Add("Text", "xm y+20 cGray w760"
  , "JP -> TL glossary: Maps specific Japanese terms to fixed terms in the target language during translation, stabilizing names and terminology. Example: エステル -> Estelle (to avoid inconsistent interpretations like Esuteru).")
txtGlossaryHelp3 := ui.Add("Text", "xm y+2 cGray w760"
  , "TL -> TL glossary: Rewrites translated output afterward for aliases, preferred wording, spelling, or capitalization, useful if you struggle with Japanese text input. Example: Esuteru -> Estelle.")
txtGlossaryHelp4 := ui.Add("Text", "xm y+20 cGray w760"
  , 'How to use: Choose a profile from the menu. Click "Edit" to add one "source -> target" mapping per line. "Add" makes a new profile, for example for one game; "Delete" removes it. Changes apply to the next screenshot translation.')

; --- Row 1: Japanese -> target-language glossary
for cpMutedGlossaryCtrl in [txtGlossaryHelp1, txtGlossaryHelp2, txtGlossaryHelp3, txtGlossaryHelp4]
    CPRegisterMutedControl(cpMutedGlossaryCtrl)
ui.Add("Text", "xm y+20", "JP -> TL glossary:")
ddlJPG := ui.Add("DropDownList", "x+m w260 0x210", [])   ; filled by RefreshGlossaryProfilesList
btnJPG_Edit := ui.Add("Button", "x+6 w70", "Edit")
btnJPG_New  := ui.Add("Button", "x+6 w70", "Add")
btnJPG_Del  := ui.Add("Button", "x+6 w70", "Delete")

; --- Row 2: target-language -> target-language glossary
ui.Add("Text", "xm y+12", "TL -> TL glossary:")
ddlENG := ui.Add("DropDownList", "x+m w260 0x210", [])
btnENG_Edit := ui.Add("Button", "x+6 w70", "Edit")
btnENG_New  := ui.Add("Button", "x+6 w70", "Add")
btnENG_Del  := ui.Add("Button", "x+6 w70", "Delete")

; wire up events
btnJPG_Edit.OnEvent("Click", (*) => OpenGlossaryEditor("jp"))
btnJPG_New .OnEvent("Click", (*) => NewGlossaryProfile())
btnJPG_Del .OnEvent("Click", (*) => DeleteGlossaryProfile())

btnENG_Edit.OnEvent("Click", (*) => OpenGlossaryEditor("en"))
btnENG_New .OnEvent("Click", (*) => NewGlossaryProfile())
btnENG_Del .OnEvent("Click", (*) => DeleteGlossaryProfile())

ddlJPG.OnEvent("Change", (*) => GlossaryChanged("jp"))
ddlENG.OnEvent("Change", (*) => GlossaryChanged("en"))

; initial fill + selection
RefreshGlossaryProfilesList(jp2enGlossaryProfile, en2enGlossaryProfile)

; --- Tab 9: PATHS
tab.UseTab(9)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer
tPython := ui.Add("Text",  "xm y+6", "Python (.exe)")
ePython := ui.Add("Edit",  "x+m w560", pythonExe)
bPy     := ui.Add("Button","x+m w80", "Browse")
bPy.OnEvent("Click", BrowsePythonExe)

tOv     := ui.Add("Text",  "xm y+10", "Overlay script (.exe)")
eOverlay:= ui.Add("Edit",  "x+m w560", overlayAhk)
bOvSel  := ui.Add("Button","x+m w80", "Browse")
bOvSel.OnEvent("Click", BrowseOverlayAhk)

tImg    := ui.Add("Text",  "xm y+10", "Screenshot translator (.py)")
eImg    := ui.Add("Edit",  "x+m w560", imgScript)
bImgSel := ui.Add("Button","x+m w80", "Browse")
bImgSel.OnEvent("Click", BrowseImageScript)

tAud    := ui.Add("Text",  "xm y+10", "Audio translator (.py)")
eAudio  := ui.Add("Edit",  "x+m w560", audioScript)
bAud    := ui.Add("Button","x+m w80", "Browse")
bAud.OnEvent("Click", BrowseAudioScript)

tExplain := ui.Add("Text",  "xm y+10", "Explainer (.py)")
eExplain := ui.Add("Edit",  "x+m w560", explainScript)
bExplainSel := ui.Add("Button","x+m w80", "Browse")
bExplainSel.OnEvent("Click", BrowseExplainScript)

; --- Debug toggle (bottom of Paths tab) ---
opts := "xm y+18 w140"
if (debugMode)
    opts .= " Checked"
cbDebug := ui.Add("CheckBox", opts, "Debug mode")
TooltipBind(cbDebug, "Write diagnostic logs for the Control Panel, overlays, and live audio translator")
cbDebug.OnEvent("Click", CPOnDebugModeToggle)

; --- Tab 7: HOTKEYS
tab.UseTab(7)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer/header row

; Column headers
ui.Add("Text", "xm y+6 w260", "Action")
ui.Add("Text", "x+10 w240",   "Current hotkey")
ui.Add("Text", "x+10 w320",   "")  ; space for buttons

; Render one row per action (label | current binding | Change / Disable / Default)
for action in hotkeyActions {
    label := hotkeyLabels[action]
    cur   := IniRead(iniPath, "hotkeys", action, hotkeyDefaults[action])

     ui.Add("Text", "xm y+10 w260", label)
    e := ui.Add("Edit", "x+10 w240 ReadOnly", cur)
    bChg := ui.Add("Button", "x+10 w90",  "Change...")
    bDis := ui.Add("Button", "x+6  w80",  "Disable")
    bDef := ui.Add("Button", "x+6  w100", "Default")

    ; store controls for later wiring
    hkEdits[action]  := e
    hkBtnChg[action] := bChg
    hkBtnDis[action] := bDis
    hkBtnDef[action] := bDef
	; --- Wire per-row buttons (bind action key so each row is independent) ---
    actKey := action  ; freeze the current key for this row

    bChg.OnEvent("Click", Hotkey_Row_Change.Bind(actKey))
    bDis.OnEvent("Click", Hotkey_Row_Disable.Bind(actKey))
    bDef.OnEvent("Click", Hotkey_Row_Default.Bind(actKey))

}

; --- Conflicts banner + bottom buttons ---
global hkConflictText
; reduce top gap and fix a small text height so it doesnâ€™t reserve extra space
hkConflictText := ui.Add("Text", "xm y+4 w800 h1 cRed", "")

; initial conflict pass
Hotkeys_ShowConflicts()

tab.UseTab()
FixAllEditableCombos()

; --- visual separator above global (non-tab) controls ---
; SS_ETCHEDHORZ = 0x10 -> draws a 1â€“2 px horizontal etched line
sepAction := ui.Add("Text", "xm y+6 w1000 h2 0x10")
btnOv      := ui.Add("Button", "xm y+18 w130",  "Open Translator")
btnOvClose := ui.Add("Button", "x+6 w140",  "Close Translator")
btnAudio  := ui.Add("Button", "x+6 w180", "Audio Translation Off")
btnExplainerLaunch := ui.Add("Button", "x+6 w140", "Open Explainer")
btnExplainerClose  := ui.Add("Button", "x+6 w140", "Close Explainer")

bSave     := ui.Add("Button", "xm y+14 w120", "Save")
try bSave.Enabled := false
bClose    := ui.Add("Button", "x+8 w120", "Close all")
chkTop    := ui.Add("CheckBox", "x+12 yp+6", "Always on top")
chkDarkMode := ui.Add("CheckBox", "x+18 yp", "Dark mode")
chkDarkMode.Value := controlDarkMode
chkDarkMode.OnEvent("Click", CPOnDarkModeToggle)

; --- Dirty wiring (manual-save-worthy settings) ---
; Topmost toggle
if IsSet(chkTop)
    chkTop.OnEvent("Click", (*) => MarkDirty())

; Transparency sliders (Translator / Explainer)
if IsSet(slTrans)
    slTrans.OnEvent("Change", (c, e) => (HandleTransparencyChange(c),    MarkDirty(), SendOverlayTheme()))
if IsSet(slTrans_EW)
    slTrans_EW.OnEvent("Change", (c, e) => (HandleTransparencyChange_EW(c), MarkDirty(), SendOverlayTheme()))

; Theme / font / sizing / borders
if IsSet(ddlTheme)
    ddlTheme.OnEvent("Change", (*) => MarkDirty())
if IsSet(ddlFont)
    ddlFont.OnEvent("Change", (*) => MarkDirty())
if IsSet(spnFontSz)
    spnFontSz.OnEvent("Change", (*) => MarkDirty())
if IsSet(spnBorder)
    spnBorder.OnEvent("Change", (*) => MarkDirty())

; Color pickers
if IsSet(clrBg)
    clrBg.OnEvent("Change", (*) => MarkDirty())
if IsSet(clrText)
    clrText.OnEvent("Change", (*) => MarkDirty())

; --- Paths tab edits: typing should mark the config as dirty ---
; (These are created in the Paths tab as: ePython, eOverlay, eImg, eAudio, eExplain)
if IsSet(ePython)
    ePython.OnEvent("Change", (*) => MarkDirty())
if IsSet(eOverlay)
    eOverlay.OnEvent("Change", (*) => MarkDirty())
if IsSet(eImg)
    eImg.OnEvent("Change", (*) => MarkDirty())
if IsSet(eAudio)
    eAudio.OnEvent("Change", (*) => MarkDirty())
if IsSet(eExplain)
    eExplain.OnEvent("Change", (*) => MarkDirty())

btnAudio.OnEvent("Click", ToggleAudioFromButton)
btnOv.OnEvent("Click",    LaunchOverlay)
btnOvClose.OnEvent("Click", CloseTranslatorOverlay)
btnExplainerLaunch.OnEvent("Click", LaunchExplainerOverlay)
btnExplainerClose.OnEvent("Click",  CloseExplainerOverlay)

bSave.OnEvent("Click", (*) => (UpdateVars(), SaveAll(), ClearDirty(), Toast("Saved"), DbgCP("Manual Save clicked")))
bClose.OnEvent("Click", ClosePanel)

ClosePanel(*) {
    static closing := false
    if closing
        return
    closing := true

    ; 1) Close overlays first (same as pressing the dedicated buttons)
    try CloseTranslatorOverlay()
    try CloseExplainerOverlay()

    ; 2) Give them a brief moment to exit cleanly
    old := A_TitleMatchMode
    SetTitleMatchMode 3
    WinWaitClose("Translator", , 0.5)
    WinWaitClose("Explainer",  , 0.5)
    SetTitleMatchMode old

    ; 3) Close our GUI and exit the app, so no AHK process remains
    try PostMessage(0x0010, 0, 0, , "ahk_id " ui.Hwnd)  ; WM_CLOSE to Control Panel
    Sleep 50
    ExitApp
}

; Load persisted Control Panel topmost state from INI and apply it
chkTop.Value := Integer(IniRead(iniPath, "cfg_control", "winTop", 0)) ? 1 : 0
ui.Opt(chkTop.Value ? "+AlwaysOnTop" : "-AlwaysOnTop")
chkTop.OnEvent("Click", (*) => ui.Opt(chkTop.Value ? "+AlwaysOnTop" : "-AlwaysOnTop"))

slTrans.OnEvent("Change", (c, e) => (HandleTransparencyChange(c), UpdateVars(), SaveAll(), SendOverlayTheme()))

; Build a safe list of existing dropdown controls before wiring handlers
ctls := []

if IsSet(ddlAProv)     ctls.Push(ddlAProv)
if IsSet(ddlA_GM)      ctls.Push(ddlA_GM)
if IsSet(ddlTR)        ctls.Push(ddlTR)
if IsSet(ddlAudioTarget) ctls.Push(ddlAudioTarget)
if IsSet(ddlProv)      ctls.Push(ddlProv)
if IsSet(ddlIMG)       ctls.Push(ddlIMG)
if IsSet(ddlIMG_GM)    ctls.Push(ddlIMG_GM)
if IsSet(ddlEProv)     ctls.Push(ddlEProv)
if IsSet(ddlEOpenAI)   ctls.Push(ddlEOpenAI)

for ctl in ctls
    ctl.OnEvent("Change", (*) => (AutoPersist(), ToggleAudioControls(), ToggleModelControls(), ToggleExplanationControls()))

ddlPrompt.OnEvent("Change", (*) => (UpdateVars(), SaveAll()))
ddlPost.OnEvent("Change",   (*) => (UpdateVars(), SaveAll()))

; --- Tab 8: API KEYS
tab.UseTab(8)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer

; Master toggle: default OFF (use Windows env). If .env exists, reflect that.
cbApiInApp := ui.Add("CheckBox", "xm y+6", "Enter API Keys in JRPG Translator (.env)")

; Info panel (storage & instructions)
ui.SetFont("w600")
ui.Add("Text", "xm y+10", "About storing your API keys")
ui.SetFont("w400")
txtApiHelp1 := ui.Add("Text"
    , "xm y+4 w740 cGray"
    , "If this box is ticked, your keys are stored in a .env file in the JRPG Translator Settings folder. This file is plain text (convenient, but not secure)."
)
txtApiHelp2 := ui.Add("Text"
    , "xm y+6 w740 cGray"
    , "Recommended: Keep keys in Windows environment variables and leave the box unticked. The app will read GEMINI_API_KEY and OPENAI_API_KEY from Windows."
)
txtApiHelp3 := ui.Add("Text"
    , "xm y+8 w740 cGray"
    , "How to store the keys in Windows: Click Start -> Search for and select 'Edit the system environment variables' -> In the 'Advanced' tab click 'Environment Variables...' -> Under 'User variables' click 'New...' -> Name: GEMINI_API_KEY (or OPENAI_API_KEY) -> Value: your key -> OK. Restart apps."
)

for cpMutedApiCtrl in [txtApiHelp1, txtApiHelp2, txtApiHelp3]
    CPRegisterMutedControl(cpMutedApiCtrl)

ui.Add("Text", "xm y+16 w200", "Gemini API key:")
eGemini := ui.Add("Edit", "x+m w420 Password")

ui.Add("Text", "xm y+10 w200", "OpenAI API key:")
eOpenAI := ui.Add("Edit", "x+m w420 Password")

; Buttons row
btnSaveEnv  := ui.Add("Button", "xm y+14 w120", "Save Keys")
btnDelEnv   := ui.Add("Button", "x+8 w120", "Delete .env")


; Helper to parse simple KEY=VALUE lines
ParseEnvLine(str, key) {
    ; returns value (string) or ""
    patt := "m)^\s*" key "\s*=\s*(.*)$"
    if RegExMatch(str, patt, &m)
        return Trim(m[1], "`r`n`t ")
    return ""
}

; Load existing .env (if present) and prefill; also set the toggle
prefOpenAI := ""
prefGemini := ""
if FileExist(envPath) {
    try {
        envBody := FileRead(envPath, "UTF-8")
        prefOpenAI := ParseEnvLine(envBody, "OPENAI_API_KEY")
        ; Accept either GOOGLE_API_KEY or GEMINI_API_KEY for Gemini; we will write both
        prefGemini := ParseEnvLine(envBody, "GEMINI_API_KEY")
        if (prefGemini = "")
            prefGemini := ParseEnvLine(envBody, "GOOGLE_API_KEY")
        cbApiInApp.Value := 1  ; .env exists â†’ assume enabled
    }
}

; Prefill edits
if (prefOpenAI != "")
    eOpenAI.Value := prefOpenAI
if (prefGemini != "")
    eGemini.Value := prefGemini

; Track last-saved values for .env to drive the dirty flag
envSavedOpenAI := eOpenAI.Value
envSavedGemini := eGemini.Value

UpdateEnvDirty(*) {
    global eOpenAI, eGemini, btnSaveEnv, cbApiInApp, envSavedOpenAI, envSavedGemini
    dirty := (cbApiInApp.Value = 1)
        && (Trim(eOpenAI.Value) != Trim(envSavedOpenAI) || Trim(eGemini.Value) != Trim(envSavedGemini))
    btnSaveEnv.Enabled := dirty
}

; Enable/disable edits + buttons based on toggle (Save button handled by UpdateEnvDirty)
ToggleApiKeyControls := (*) => (
    eOpenAI.Enabled := cbApiInApp.Value = 1,
    eGemini.Enabled := cbApiInApp.Value = 1,
    btnDelEnv.Enabled  := cbApiInApp.Value = 1,
    UpdateEnvDirty()
)

; Initial state: off if there is no .env, on if file exists
ToggleApiKeyControls()

; Watch for changes to enable Save .env only when something actually changed
eOpenAI.OnEvent("Change", UpdateEnvDirty)
eGemini.OnEvent("Change", UpdateEnvDirty)
UpdateEnvDirty()

; When user flips the toggle:
; - If turned OFF: we don't delete .env automatically; we just disable the fields.
;   (Windows env remains the default source. Python workers will ignore .env if you delete it.)
cbApiInApp.OnEvent("Click", (*) => (
    IniWrite(cbApiInApp.Value, iniPath, "api", "keys_in_env"),
    ToggleApiKeyControls()
))

; Save .env atomically (writes OPENAI_API_KEY, GOOGLE_API_KEY and GEMINI_API_KEY)
SaveApiEnv(*) {
    global eOpenAI, eGemini, envPath, cbApiInApp, ToggleApiKeyControls
    global envSavedOpenAI, envSavedGemini
    openai := Trim(eOpenAI.Value)
    gemini := Trim(eGemini.Value)

    body := "OPENAI_API_KEY=" openai "`r`n"
          . "GOOGLE_API_KEY=" gemini "`r`n"
          . "GEMINI_API_KEY=" gemini "`r`n"

    ; Ensure the folder for envPath exists
    SplitPath envPath, , &envDir
    if !DirExist(envDir)
        DirCreate(envDir)

    tmp := envPath ".tmp"
    try {
        if FileExist(tmp)
            FileDelete(tmp)
        FileAppend(body, tmp, "UTF-8")

        if FileExist(envPath)
            FileDelete(envPath)

        FileMove(tmp, envPath, true)

        cbApiInApp.Value := 1
        envSavedOpenAI := openai
        envSavedGemini := gemini
        ToggleApiKeyControls()
        UpdateEnvDirty()
        Toast("Saved .env to " envPath)
    } catch as ex {
        try if FileExist(tmp) FileDelete(tmp)
        MsgBox("Saving .env failed:`n" ex.Message)
    }
}

btnSaveEnv.OnEvent("Click", SaveApiEnv)

; Delete .env (and keep toggle OFF)
DeleteEnvFile(*) {
    global envPath, cbApiInApp, eOpenAI, eGemini, iniPath
    global ToggleApiKeyControls, UpdateEnvDirty
    global envSavedOpenAI, envSavedGemini
    if FileExist(envPath)
        FileDelete(envPath)
    cbApiInApp.Value := 0
    eOpenAI.Value := ""
    eGemini.Value := ""
    envSavedOpenAI := ""
    envSavedGemini := ""
    IniWrite(0, iniPath, "api", "keys_in_env")
    ToggleApiKeyControls()
    UpdateEnvDirty()
    Toast("Deleted .env")
}
btnDelEnv.OnEvent("Click", DeleteEnvFile)

OpenEnvFolder(*) {
    global appDir
    try
        Run('explorer.exe "' appDir '"')
    catch as ex
        MsgBox("Couldn't open folder:`n" appDir "`n`n" ex.Message)
}

; Create the visible tab bar last so it stays above the native page host.
tab.UseTab()
CPCreateCustomTabBar()
CPRefreshThemeBrushes()

ui.OnEvent("Close",  (*) => (SetTimer(_UpdateStatus, 0), SetTimer(UpdateCPFocusRing, 0), SetTimer(UpdateCPActiveTabHighlight, 0), SavePanelBounds(), ExitApp()))
ui.OnEvent("Escape", (*) => (SetTimer(_UpdateStatus, 0), SetTimer(UpdateCPFocusRing, 0), SetTimer(UpdateCPActiveTabHighlight, 0), SavePanelBounds(), ExitApp()))
ui.OnEvent("Size",   ResizeUI)

; wire buttons
btnA_GM_Add   .OnEvent("Click", (*) => AddModelInteractive(model_gemini_audio, "gemini_audio", ddlA_GM, "gemini", "audio"))
btnA_GM_Del   .OnEvent("Click", (*) => DeleteModel(model_gemini_audio, "gemini_audio", ddlA_GM))

btnTR_Add     .OnEvent("Click", (*) => AddModelInteractive(model_openai_audio, "openai_audio", ddlTR, "openai", "audio"))
btnTR_Del     .OnEvent("Click", (*) => DeleteModel(model_openai_audio, "openai_audio", ddlTR))

btnIMG_Add    .OnEvent("Click", (*) => AddModelInteractive(model_openai_img, "openai_img", ddlIMG, "openai", "screenshot"))
btnIMG_Del    .OnEvent("Click", (*) => DeleteModel(model_openai_img,"openai_img",   ddlIMG))

btnIMG_GM_Add .OnEvent("Click", (*) => AddModelInteractive(model_gemini_img, "gemini_img", ddlIMG_GM, "gemini", "screenshot"))
btnIMG_GM_Del .OnEvent("Click", (*) => DeleteModel(model_gemini_img,"gemini_img",   ddlIMG_GM))

; initial paint + status, then start timer
LoadFontsIntoCombo()
LoadFontsIntoCombo_EW()
Repaint()
_UpdateStatus()
SetTimer(_UpdateStatus, 1000)

; During background initialization, keep the temporary native window out of the
; taskbar and prevent Windows from activating it while controls are measured.
if (CP_BACKGROUND_START)
    ui.Opt("+ToolWindow +E0x08000000") ; WS_EX_NOACTIVATE

; Show the window, restore saved bounds if valid; otherwise use hardcoded defaults and seed control.ini
if (IsValidBounds(guiX_saved, guiY_saved, guiW_saved, guiH_saved)) {
    ; If the INI was written before this fix, it likely holds OUTER sizes.
    ; Measure non-client deltas and convert once to CLIENT size for Show().
    if (bounds_mode != "client") {
        ui.Show("Hide")                                  ; create handle and metrics
        ui.GetPos(,, &ow, &oh)                           ; outer size
        ui.GetClientPos(,, &cw, &ch)                     ; client size
        ncW := ow - cw, ncH := oh - ch                   ; non-client deltas
        ui.Hide()
        guiW_saved := Max(0, guiW_saved - ncW)           ; convert outer->client
        guiH_saved := Max(0, guiH_saved - ncH)
        IniWrite("client", iniPath, "gui_bounds", "bounds_mode")
        DbgCP("Converted saved bounds from OUTER to CLIENT using ncW=" ncW " ncH=" ncH)
    }
    DbgCP("Restore saved panel bounds (client): x=" guiX_saved " y=" guiY_saved " w=" guiW_saved " h=" guiH_saved)
    cpShowOptions := "w" guiW_saved " h" guiH_saved " x" guiX_saved " y" guiY_saved
    ui.Show((CP_BACKGROUND_START ? "NA Hide " : "") cpShowOptions)
    if (CP_BACKGROUND_START)
        ui.Hide()
} else {
    DbgCP("Use default panel bounds: x=" defGuiX " y=" defGuiY " w=" defGuiW " h=" defGuiH)
    cpShowOptions := "w" defGuiW " h" defGuiH " x" defGuiX " y" defGuiY
    ui.Show((CP_BACKGROUND_START ? "NA Hide " : "") cpShowOptions)
    if (CP_BACKGROUND_START)
        ui.Hide()
    ; Seed control.ini [gui_bounds] immediately so subsequent launches restore these
    SavePanelBounds()
}
; Ensure first paint draws all children cleanly (fixes clipped checkbox text/box)
DllCall("RedrawWindow"
    , "ptr", ui.Hwnd
    , "ptr", 0
    , "ptr", 0
    , "uint", 0x0001 | 0x0080 | 0x0100) ; RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW
CPCreateComboArrowOverlays()
CPApplyControlPanelTheme()

Rebind_LaunchExplainerRequest()
Rebind_ExplainLastTranslation()
Rebind_StartStopAudio()
Rebind_HideShowControlPanel()
RegisterControlPanelArrowNavigation()
SetTimer(UpdateCPFocusRing, 60)
UpdateCPFocusRing()
SetTimer(UpdateCPActiveTabHighlight, 80)
UpdateCPActiveTabHighlight()

; Auto-open overlays if toggled in cfg
if (Integer(IniRead(iniPath, "cfg", "openTranslatorOnLaunch", 0))) {
    SetTimer(LaunchOverlay, -100)
}
if (Integer(IniRead(iniPath, "cfg", "openExplainerOnLaunch", 0))) {
    SetTimer(LaunchExplainerOverlay, -200)
}

; Some custom controls repaint themselves during initialization. Enforce the
; hidden state once more, then restore normal window styles for the first
; intentional hotkey/double-launch reveal.
if (CP_BACKGROUND_START) {
    ui.Hide()
    ui.Opt("-ToolWindow -E0x08000000")
    DbgCP("Background start complete: control panel remains hidden.")
}

; Force immediate paint so text is visible without hover
ForcePaint(ctrls*) {
    for c in ctrls {
        try {
            DllCall("user32\RedrawWindow", "ptr", c.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x0401) ; RDW_INVALIDATE|RDW_UPDATENOW
        }
    }
}
ForcePaint(ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt, ddlPost)
; Do one more pass after creation so no dropdowns look â€œselectedâ€ on first open
ClearAllComboSelections(*) {
    global ddlAProv, ddlA_GM, ddlTR, ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt, ddlPost
    for cmb in [ddlAProv, ddlA_GM, ddlTR, ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt, ddlPost]
        ComboUnselectText(cmb)
}

SetGuiAndTrayIcon(ui, A_ScriptDir "\icon.ico")

; =========================
; Helpers (GUI)
; =========================
SavePanelBounds() {
    global ui, iniPath
    try {
        if !IsObject(ui) || !ui.Hwnd
            return
        ui.GetPos(&x, &y)                 ; outer position
        ui.GetClientPos(,, &cliW, &cliH)  ; client size
        if !IsValidBounds(x, y, cliW, cliH) {
            DbgCP("Skipped saving invalid panel bounds: x=" x " y=" y " w=" cliW " h=" cliH)
            return
        }
        IniWrite(x,     iniPath, "gui_bounds", "x")
        IniWrite(y,     iniPath, "gui_bounds", "y")
        IniWrite(cliW,  iniPath, "gui_bounds", "w")
        IniWrite(cliH,  iniPath, "gui_bounds", "h")
        IniWrite("client", iniPath, "gui_bounds", "bounds_mode")
        DbgCP("Saved panel bounds (client): x=" x " y=" y " w=" cliW " h=" cliH)
    }
}

HandleTransparencyChange(sliderCtrl) {
    global lblTransPct, overlayTrans
    val := sliderCtrl.Value
    overlayTrans := val
    pct := Round(val / 255 * 100)
    lblTransPct.Value := pct . "%"
    oldMode := A_TitleMatchMode
    SetTitleMatchMode 3
    try WinSetTransparent(val, "Translator")
    catch ; ignore
    SetTitleMatchMode oldMode

    DbgCP("Transparency slider -> " val " (" pct "%)")
}

Repaint(){
    global ePython,eAudio,eOverlay,eImg,ddlTR,ddlAProv,ddlA_GM,ddlAudioTarget,ddlProv,ddlIMG,ddlIMG_GM
    global pythonExe,audioScript,overlayAhk,imgScript,trModel,audioProvider,geminiAudioModel,audioTargetLang
    global imgProvider,imgModel,geminiImgModel,overlayTrans
    global slTrans, lblTransPct
    global rectBg,rectTxt, boxBgHex,txtHex
    global ddlFont, edFSize, chkFontBold, fontName, fontSize, fontBold
    global chkFontBold_EW, fontBold_EW
    global ddlPrompt, promptProfile
    global ddlPost, imgPostproc, postCodes
	global ddlEProv, ddlEOpenAI, ddlEGem
    global explainProvider, explainOpenAIModel, explainGeminiModel, iniPath
    global model_openai_img, model_gemini_img, model_openai_explain, model_gemini_explain
    global model_openai_audio, model_gemini_audio

    ePython.Value := pythonExe
    eAudio.Value  := audioScript
    eOverlay.Value:= overlayAhk
    eImg.Value    := imgScript
    eExplain.Value := explainScript

    ddlAProv.Text := audioProvider
    SetComboToExistingItem(ddlA_GM, model_gemini_audio, geminiAudioModel)
    SetComboToExistingItem(ddlTR, model_openai_audio, trModel)
    ddlAudioTarget.Text := audioTargetLang

    ; AFTER (use names unique to Repaint)
    provIdx_r := (StrLower(imgProvider) = "gemini") ? 1 : 2
    ddlProv.Choose(provIdx_r)

    imgIdx_r := ArrIndexOf(model_openai_img, imgModel)
    ddlIMG.Choose(imgIdx_r ? imgIdx_r : 1)

    imgGMIdx_r := ArrIndexOf(model_gemini_img, geminiImgModel)
    ddlIMG_GM.Choose(imgGMIdx_r ? imgGMIdx_r : 1)




    slTrans.Value := overlayTrans
    lblTransPct.Value := Round(overlayTrans / 255 * 100) . "%"

    rectBg.Opt("Background" . boxBgHex)
    rectTxt.Opt("Background" . txtHex)

    try ddlFont.Text := fontName
    edFSize.Value := fontSize
    chkFontBold.Value := fontBold

    ; (Do not set ddlPrompt.Text here â€“ list may be empty on first run)
    postSelIdx := ArrIndexOf(postCodes, imgPostproc)
    if (!postSelIdx)
    postSelIdx := 1
    ddlPost.Value := postSelIdx

    RefreshPromptProfilesList(promptProfile)

    ; ===== Explanation tab: reflect persisted provider/model =====
    if IsSet(ddlEProv) {
        idx := (StrLower(explainProvider) = "gemini") ? 1 : 2
        ddlEProv.Choose(idx)
    }
    if IsSet(ddlEGem)
        SetComboToExistingItem(ddlEGem, model_gemini_explain, explainGeminiModel)
    if IsSet(ddlEOpenAI)
        SetComboToExistingItem(ddlEOpenAI, model_openai_explain, explainOpenAIModel)
    ToggleExplanationControls()
    ; Defensive: pull from INI again to win against any other state that might have run before/after
    SyncExplanationFromIni()

      ; ===== Populate Explanation Window controls (own state) =====
    if (CP_ENABLE_EXPLAINER_DESIGN) {
        ; Transparency + label
        slTrans_EW.Value := overlayTrans_EW
        try lblTransPct_EW.Value := Round(overlayTrans_EW / 255 * 100) . "%"

        ; Color preview rectangles
        try rectBg_EW.Opt("Background" . boxBgHex_EW)
        try rectTxt_EW.Opt("Background" . txtHex_EW)

        ; Font and size
        try ddlFont_EW.Text := fontName_EW
        try edFSize_EW.Value := fontSize_EW
        try udFSize_EW.Value := fontSize_EW
        try chkFontBold_EW.Value := fontBold_EW

                ; Profile row: use last Explainer profile from control.ini
        try ddlProf_EW.Text := IniRead(iniPath, "profiles", "explainer_last", "")
    }

    ToggleModelControls()
}

ToggleModelControls(){
    global ddlProv, ddlIMG, ddlIMG_GM
    global btnIMG_Add, btnIMG_Del, btnIMG_GM_Add, btnIMG_GM_Del
    prov := StrLower(Trim(ddlProv.Text))
    openAIEnabled := (prov = "openai")
    geminiEnabled := (prov = "gemini")
    ddlIMG.Enabled := openAIEnabled
    btnIMG_Add.Enabled := openAIEnabled
    btnIMG_Del.Enabled := openAIEnabled
    ddlIMG_GM.Enabled := geminiEnabled
    btnIMG_GM_Add.Enabled := geminiEnabled
    btnIMG_GM_Del.Enabled := geminiEnabled
}
ToggleAudioControls(){
    global ddlAProv, ddlTR, ddlA_GM
    global btnTR_Add, btnTR_Del, btnA_GM_Add, btnA_GM_Del
    ap := StrLower(ddlAProv.Text)
    isOpenAI := (ap = "openai")
    ddlTR.Enabled := isOpenAI
    btnTR_Add.Enabled := isOpenAI
    btnTR_Del.Enabled := isOpenAI
    ddlA_GM.Enabled := !isOpenAI
    btnA_GM_Add.Enabled := !isOpenAI
    btnA_GM_Del.Enabled := !isOpenAI
}
; NEW: Explanation tab toggles
ToggleExplanationControls(){
    global ddlEProv, ddlEOpenAI, ddlEGem
    global btnEOpenAI_Add, btnEOpenAI_Del, btnEGem_Add, btnEGem_Del
    ep := StrLower(Trim(ddlEProv.Text))
    openAIEnabled := (ep = "openai")
    geminiEnabled := (ep = "gemini")
    ddlEOpenAI.Enabled := openAIEnabled
    btnEOpenAI_Add.Enabled := openAIEnabled
    btnEOpenAI_Del.Enabled := openAIEnabled
    ddlEGem.Enabled := geminiEnabled
    btnEGem_Add.Enabled := geminiEnabled
    btnEGem_Del.Enabled := geminiEnabled
}

; NEW: force-sync Explanation dropdowns from INI (defensive against any later repaint)
SyncExplanationFromIni(){
    global iniPath
    global ddlEProv, ddlEOpenAI, ddlEGem
    global model_openai_explain, model_gemini_explain

    prov := StrLower(Trim(IniRead(iniPath, "cfg_explainer", "explainProvider", "")))
    gm   := Trim(IniRead(iniPath, "cfg_explainer", "explainGeminiModel", ""))
    om   := Trim(IniRead(iniPath, "cfg_explainer", "explainOpenAIModel", ""))

    if (prov != "")
        ddlEProv.Choose(prov = "gemini" ? 1 : 2)
    if (gm != "")
        SetComboToExistingItem(ddlEGem, model_gemini_explain, gm)
    if (om != "")
        SetComboToExistingItem(ddlEOpenAI, model_openai_explain, om)
    ToggleExplanationControls()
}

PopulateSpeakersList(select := "") {
    global ddlSpeaker, pythonExe, audioScript, speakerName
    px := ResolvePath(pythonExe)
    ap := ResolvePath(audioScript)
    if !(FileExist(px) && FileExist(ap)) {
        ; silently skip if paths arenâ€™t set yet
        return
    }
    txt := ExecCaptureHidden(px, ap, "--list-speakers")
    txt := Trim(txt, "`r`n `t")
        arr := (txt = "" ? [] : StrSplit(txt, "`r`n"))

    ; Sanitize: drop filesystem paths or stray Python chatter
    clean := []
    for _, n in arr {
        n := Trim(n)
        if (n = "")
            continue
        ; Skip obvious paths / noise
        if RegExMatch(n, "i)^(?:[A-Z]:\\|\\\\|/).+")          ; drive or UNC path
            continue
        if InStr(n, "import pkg_resources")                   ; common noisy line
            continue
        if RegExMatch(n, "i)\.py($| )")                       ; python file mentions
            continue
        clean.Push(n)
    }
    arr := clean

    ; First entry: Windows default
    ddlSpeaker.Delete()
    ddlSpeaker.Add(["[Windows Default]"])
    for n in arr
        if (Trim(n) != "")
            ddlSpeaker.Add([n])

    ; choose selection: explicit 'select', else saved speakerName, else default
    pick := select != "" ? select : speakerName
    if (pick = "" || pick = "[Windows Default]") {
        ddlSpeaker.Choose(1)  ; first item is default
    } else {
        ; arr holds the device names we just added after the default
        idx := 0
        for i, name in arr
            if (name = pick) {
                idx := i + 1  ; +1 because item 1 is [Windows Default]
                break
            }
        if (idx)
            ddlSpeaker.Choose(idx)
        else
            ddlSpeaker.Choose(1)
    }

}

RefreshSpeakerList(*) {
    PopulateSpeakersList(Trim(ddlSpeaker.Text))
}

SpeakerChanged(*) {
    global speakerName, ddlSpeaker, iniPath
    speakerName := Trim(ddlSpeaker.Text)
    IniWrite(speakerName, iniPath, "cfg", "speakerName")
}

ResolvePythonNoConsole(px) {
    try {
        if InStr(px, "\python.exe") {
            alt := StrReplace(px, "\python.exe", "\pythonw.exe")
            if FileExist(alt)
                return alt
        }
    }
    return px
}

UpdateVars(){
    global pythonExe,audioScript,overlayAhk,imgScript,overlayTrans
    global trModel,audioProvider,geminiAudioModel,audioTargetLang
    global imgProvider,imgModel,geminiImgModel
    global eExplain, explainScript
	global explainProvider, explainOpenAIModel, explainGeminiModel
    global ddlEProv, ddlEOpenAI, ddlEGem
    global ePython,eAudio,eOverlay,eImg,ddlTR,ddlAProv,ddlA_GM,ddlAudioTarget,ddlProv,ddlIMG,ddlIMG_GM,slTrans
	global ddlFont, edFSize, chkFontBold, fontName, fontSize, fontBold
    global ddlPrompt, promptProfile
    global ddlPost, imgPostproc
	global debugMode, cbDebug
    pythonExe        := ePython.Value
    audioScript      := eAudio.Value
    overlayAhk       := eOverlay.Value
    imgScript        := eImg.Value
    overlayTrans     := slTrans.Value
    explainScript    := eExplain.Value
    trModel          := ddlTR.Text
    audioProvider    := ddlAProv.Text
    geminiAudioModel := ddlA_GM.Text
    audioTargetLang  := ddlAudioTarget.Text
    imgProvider      := ddlProv.Text
    imgModel         := ddlIMG.Text
    geminiImgModel   := ddlIMG_GM.Text
    fontName         := ddlFont.Text
    fontSize         := Integer(edFSize.Value)
    fontBold         := chkFontBold.Value ? 1 : 0
    SyncUnifiedWindowAppearance()
    promptProfile    := ddlPrompt.Text
    imgPostproc      := postCodes[ddlPost.Value]
	explainProvider    := ddlEProv.Text
    explainOpenAIModel := ddlEOpenAI.Text
    explainGeminiModel := ddlEGem.Text
    debugMode := cbDebug.Value ? 1 : 0
    SetDebugMode(debugMode)
}

MoveRowWithButtons(combo, btnAdd, btnDel, rightEdge, btnH, maxW := 0) {
    btnW := 60, g := 6
    combo.GetPos(&cx,&cy,,)
    newW := Max(160, rightEdge - cx - (btnW*2 + g*2))
    if (maxW && newW > maxW)
        newW := maxW
    combo.Move(, , newW)
    btnAdd.Move(cx + newW + g, cy, btnW, btnH)
    btnDel.Move(cx + newW + g + btnW + g, cy, btnW, btnH)
}
; --- bounded version used when a row has two combos (ASR + Translate) ---
MoveRowWithButtonsBound(combo, btnAdd, btnDel, rightLimitX, btnH) {
    btnW := 60, g := 6
    combo.GetPos(&cx,&cy,,)
    newW := Max(160, rightLimitX - cx - (btnW*2 + g*2))
    combo.Move(, , newW)
    btnAdd.Move(cx + newW + g, cy, btnW, btnH)
    btnDel.Move(cx + newW + g + btnW + g, cy, btnW, btnH)
}

ResizeUI(gui, minMax, w, h){
    ; --- Prevent flicker & â€œinvisible until hoverâ€ by suspending redraw during bulk moves ---
    hwnd := gui.Hwnd
    ; WM_SETREDRAW (0x000B) â†’ 0 = suspend redraw (call WinAPI directly with the HWND)
    if (hwnd)
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x000B, "ptr", 0, "ptr", 0)

    global CP_CANVAS_MIN_W, CP_CANVAS_MIN_H
    viewportW := w
    viewportH := h
    canvasRestore := CPCanvasResetForLayout(false)
    w := Max(CP_CANVAS_MIN_W, viewportW)
    h := Max(CP_CANVAS_MIN_H, viewportH)

    global pad, gap
    global tab, sepAction
    global tPython,ePython,bPy,tAud,eAudio,bAud,tOv,eOverlay,bOvSel,tImg,eImg,bImgSel,tExplain,eExplain,bExplainSel
    global ddlSpeaker,ddlAProv,ddlA_GM,ddlTR,ddlAudioTarget,ddlProv,ddlIMG,ddlIMG_GM
    global btnStart,btnStop,btnAudio,btnOv,btnOvClose,btnExplainerLaunch,btnExplainerClose,btnExplainNow,bSave,bClose,chkTop,chkDarkMode,chkGuess
    global btnA_GM_Add, btnA_GM_Del, btnTR_Add, btnTR_Del
    global btnIMG_Add, btnIMG_Del, btnIMG_GM_Add, btnIMG_GM_Del
    ; NEW prompt widgets
        ; NEW prompt widgets
    global ddlPrompt, btnPrEdit, btnPrNew, btnPrDel, ddlPost
    ; AUDIO target language widget
    global ddlAudioTarget

    browseW := 80
    btnH    := 32

    gap1 := 10
    gap2 := 12
    bottomBlockH := btnH*2 + gap1 + gap2 + pad

    tabH := Max(260, h - pad*2 - bottomBlockH)
    tab.Move(pad, pad, w - pad*2, tabH)
    CPLayoutCustomTabBar(viewportW)

    tab.GetPos(&tx,&ty,&tw,&th)
    rightEdge := tx + tw - (pad + 28)

    innerW := (w - pad*2)
    ; reserve enough space so the Edit column starts at a fixed x, regardless of label text width
    labelW := 180                      ; widened to accommodate the longest label
    editX  := tx + pad + labelW        ; fixed left edge for all path Edit controls
    editW  := Max(260, rightEdge - editX - (gap + browseW))  ; width that keeps the Browse button inside the tab

    ; Keep all four rows perfectly aligned (same x and widths)
        for pair in [[ePython,bPy],[eAudio,bAud],[eOverlay,bOvSel],[eImg,bImgSel],[eExplain,bExplainSel]] {
        ctrl := pair[1], btn := pair[2]
        ctrl.GetPos(, &ey,,)
        ctrl.Move(editX, ey, editW)
        ; Clamp the Browse button so it never spills past the tab's right edge
        btnX := Min(editX + editW + gap, rightEdge - browseW)
        btn.Move(btnX, ey, browseW, btnH)
    }

    ; Live translation uses one shared combo rectangle. Its right edge follows
    ; Listen device, while its left edge leaves room for both model labels.
    ddlSpeaker.GetPos(&audioDeviceX, , &audioDeviceW)
    ddlA_GM.GetPos(&audioGeminiX, &audioGeminiY)
    ddlTR.GetPos(&audioOpenAIX, &audioOpenAIY)
    audioComboLeft := Max(audioGeminiX, audioOpenAIX)
    audioComboRight := audioDeviceX + audioDeviceW
    audioComboW := Max(160, audioComboRight - audioComboLeft)

    for audioCombo in [ddlAProv, ddlA_GM, ddlTR, ddlAudioTarget]
        audioCombo.Move(audioComboLeft, , audioComboW)

    btnA_GM_Add.Move(audioComboRight + gap, audioGeminiY, 60, btnH)
    btnA_GM_Del.Move(audioComboRight + gap + 60 + gap, audioGeminiY, 60, btnH)
    btnTR_Add.Move(audioComboRight + gap, audioOpenAIY, 60, btnH)
    btnTR_Del.Move(audioComboRight + gap + 60 + gap, audioOpenAIY, 60, btnH)
    ; Screenshot Translation uses one shared combo boundary: the left edge of
    ; "Highlight guessed subjects". This keeps selected values readable when
    ; a narrow viewport scrolls focused controls into view.
    chkGuess.GetPos(&shotComboRight)
    ddlProv.GetPos(&shotProvX)
    ddlProv.Move(, , Max(160, shotComboRight - shotProvX))

    btnW := 70, g := 6
    for shotRow in [[ddlIMG_GM, btnIMG_GM_Add, btnIMG_GM_Del], [ddlIMG, btnIMG_Add, btnIMG_Del]] {
        shotCombo := shotRow[1], shotAdd := shotRow[2], shotDel := shotRow[3]
        shotCombo.GetPos(&shotX, &shotY)
        shotW := Max(160, shotComboRight - shotX)
        shotCombo.Move(, , shotW)
        ; Match the prompt row's Add/Delete columns, leaving its Edit column empty.
        shotAdd.Move(shotX + shotW + g + btnW + g, shotY, btnW, btnH)
        shotDel.Move(shotX + shotW + g + (btnW + g)*2, shotY, btnW, btnH)
    }

    ; NEW: prompt row (combo + 3 buttons)
    ddlPrompt.GetPos(&pcx,&pcy,,)
    pW := Max(160, shotComboRight - pcx)
    ddlPrompt.Move(, , pW)
    btnPrEdit.Move(pcx + pW + g, pcy, btnW, btnH)
    btnPrNew .Move(pcx + pW + g + btnW + g, pcy, btnW, btnH)
    btnPrDel .Move(pcx + pW + g + (btnW+g)*2, pcy, btnW, btnH)

    ; post-processing row (single combo)
    ddlPost.GetPos(&ppx,&ppy,,)
    ppW := Max(160, shotComboRight - ppx)
    ddlPost.Move(, , ppW)

            ; place a thin separator directly under the tab
    sepY := pad + tabH + 4
    try sepAction.Move(pad, sepY, w - pad*2, 2)

    ; action row always sits just below the separator
    yAction := sepY + 8
    ; Order (left -> right): Open, Close, Audio Translation, Open Explainer, Close Explainer
    btnOv.Move(pad, yAction, 130, btnH)                                ; Open Translator
    btnOvClose.Move(pad + 130 + gap, yAction, 140, btnH)               ; Close Translator
    btnAudio.Move(pad + 130 + 140 + gap*2, yAction, 180, btnH)          ; Audio Translation On/Off
    btnExplainerLaunch.Move(pad + 130 + 140 + 180 + gap*3, yAction, 140, btnH)  ; Open Explainer
    btnExplainerClose.Move(pad + 130 + 140 + 180 + 140 + gap*4, yAction, 140, btnH) ; Close Explainer

    ; Use the virtual canvas height so this row remains at the design bottom
    ; when the visible window is smaller and scrollbars are present.
    ySave := h - btnH - pad

    bSave.Move(pad, ySave, 120, btnH)
    bClose.Move(pad + 120 + 8, ySave, 120, btnH)
    bClose.GetPos(&cx,&cy,&btnW,)
    chkTop.Move(cx + btnW + 12, ySave + 6)
    chkTop.GetPos(&cpTopX, &cpTopY, &cpTopW)
    chkDarkMode.Move(cpTopX + cpTopW + 18, ySave + 6)
    CPUpdateComboArrowOverlays()
    CPCanvasFinishLayout(viewportW, viewportH, canvasRestore, false)

    ; --- Re-enable redraw and force repaint of the whole window and all children ---
    if (hwnd) {
        ; WM_SETREDRAW â†’ 1 = resume redraw (WinAPI with HWND)
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x000B, "ptr", 1, "ptr", 0)
        ; RedrawWindow(hwnd, NULL, NULL, RDW_INVALIDATE|RDW_ERASE|RDW_ALLCHILDREN|RDW_UPDATENOW)
        DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", 0x0001|0x0004|0x0080|0x0100)
    }
}

; =========================
; Overlay color picking + messaging
; =========================
CPColorHexToHSV(hexColor) {
    rgb := Integer("0x" Trim(hexColor, "#"))
    red := ((rgb >> 16) & 0xFF) / 255
    green := ((rgb >> 8) & 0xFF) / 255
    blue := (rgb & 0xFF) / 255
    high := Max(red, green, blue)
    low := Min(red, green, blue)
    delta := high - low

    hue := 0
    if (delta != 0) {
        if (high = red)
            hue := 60 * Mod((green - blue) / delta, 6)
        else if (high = green)
            hue := 60 * (((blue - red) / delta) + 2)
        else
            hue := 60 * (((red - green) / delta) + 4)
    }
    if (hue < 0)
        hue += 360

    saturation := (high = 0) ? 0 : (delta / high)
    return Map("h", Round(hue), "s", Round(saturation * 100), "v", Round(high * 100))
}

CPColorHSVToHex(hue, saturation, brightness) {
    hue := Mod(Max(0, Min(359, hue + 0)), 360)
    saturation := Max(0, Min(100, saturation + 0)) / 100
    brightness := Max(0, Min(100, brightness + 0)) / 100

    chroma := brightness * saturation
    section := hue / 60
    intermediate := chroma * (1 - Abs(Mod(section, 2) - 1))
    redPart := 0, greenPart := 0, bluePart := 0
    if (section < 1)
        redPart := chroma, greenPart := intermediate
    else if (section < 2)
        redPart := intermediate, greenPart := chroma
    else if (section < 3)
        greenPart := chroma, bluePart := intermediate
    else if (section < 4)
        greenPart := intermediate, bluePart := chroma
    else if (section < 5)
        redPart := intermediate, bluePart := chroma
    else
        redPart := chroma, bluePart := intermediate

    match := brightness - chroma
    red := Round((redPart + match) * 255)
    green := Round((greenPart + match) * 255)
    blue := Round((bluePart + match) * 255)
    return Format("{:02X}{:02X}{:02X}", red, green, blue)
}

CPColorGradientWriteVertex(vertices, offset, x, y, colorHex) {
    rgb := Integer("0x" Trim(colorHex, "#"))
    NumPut("Int", x, vertices, offset)
    NumPut("Int", y, vertices, offset + 4)
    NumPut("UShort", ((rgb >> 16) & 0xFF) << 8, vertices, offset + 8)
    NumPut("UShort", ((rgb >> 8) & 0xFF) << 8, vertices, offset + 10)
    NumPut("UShort", (rgb & 0xFF) << 8, vertices, offset + 12)
    NumPut("UShort", 0, vertices, offset + 14)
}

CPColorGradientFillRect(hdc, left, top, right, bottom, startHex, endHex) {
    if (right <= left || bottom <= top)
        return
    vertices := Buffer(32, 0)
    CPColorGradientWriteVertex(vertices, 0, left, top, startHex)
    CPColorGradientWriteVertex(vertices, 16, right, bottom, endHex)
    gradientRect := Buffer(8, 0)
    NumPut("UInt", 0, gradientRect, 0)
    NumPut("UInt", 1, gradientRect, 4)
    DllCall("msimg32\GradientFill", "ptr", hdc, "ptr", vertices.Ptr, "uint", 2
        , "ptr", gradientRect.Ptr, "uint", 1, "uint", 0) ; GRADIENT_FILL_RECT_H
}

CPDrawControllerColorGradient(hdc, sourceHwnd, channelRect, sliderData) {
    global controlDarkMode
    clientRect := Buffer(16, 0)
    DllCall("user32\GetClientRect", "ptr", sourceHwnd, "ptr", clientRect.Ptr)
    clientH := NumGet(clientRect, 12, "int")
    left := NumGet(channelRect, 0, "int")
    right := NumGet(channelRect, 8, "int")
    barH := Max(8, Round(10 * GetWindowDPI(sourceHwnd) / 96))
    top := Max(1, Floor((clientH - barH) / 2))
    bottom := Min(clientH - 1, top + barH)

    focused := DllCall("user32\GetFocus", "ptr") = sourceHwnd
    borderHex := focused ? (controlDarkMode ? "FFFFFF" : CPPalette(0)["accentFocus"])
        : (controlDarkMode ? "777777" : "8A8A8A")
    frameRect := Buffer(16, 0)
    NumPut("Int", left, frameRect, 0)
    NumPut("Int", top, frameRect, 4)
    NumPut("Int", right, frameRect, 8)
    NumPut("Int", bottom, frameRect, 12)
    frameBrush := DllCall("gdi32\CreateSolidBrush", "uint", CPColorRef(borderHex), "ptr")
    try DllCall("user32\FrameRect", "ptr", hdc, "ptr", frameRect.Ptr, "ptr", frameBrush)
    finally DllCall("gdi32\DeleteObject", "ptr", frameBrush)
    left += 1, top += 1, right -= 1, bottom -= 1

    hue := sliderData["hue"].Value
    saturation := sliderData["saturation"].Value
    brightness := sliderData["brightness"].Value
    switch sliderData["kind"] {
        case "hue":
            hueStops := ["FF0000", "FFFF00", "00FF00", "00FFFF", "0000FF", "FF00FF", "FF0000"]
            gradientW := Max(1, right - left)
            Loop 6 {
                segmentLeft := left + Floor(gradientW * (A_Index - 1) / 6)
                segmentRight := left + Floor(gradientW * A_Index / 6)
                CPColorGradientFillRect(hdc, segmentLeft, top, segmentRight, bottom
                    , hueStops[A_Index], hueStops[A_Index + 1])
            }
        case "saturation":
            CPColorGradientFillRect(hdc, left, top, right, bottom, "FFFFFF"
                , CPColorHSVToHex(hue, 100, 100))
        case "brightness":
            CPColorGradientFillRect(hdc, left, top, right, bottom, "000000"
                , CPColorHSVToHex(hue, saturation, 100))
    }
}

CPControllerColorGradientCustomDraw(wParam, lParam, msg, parentHwnd) {
    global CPControllerColorGradientSliders
    if !lParam
        return
    sourceHwnd := NumGet(lParam, 0, "ptr")
    if !CPControllerColorGradientSliders.Has(sourceHwnd)
        return
    notifyCode := NumGet(lParam, 2 * A_PtrSize, "int")
    if (notifyCode != -12) ; NM_CUSTOMDRAW
        return

    stageOffset := (A_PtrSize = 8) ? 24 : 12
    drawStage := NumGet(lParam, stageOffset, "uint")
    if (drawStage = 0x00000001) ; CDDS_PREPAINT
        return 0x00000020 ; CDRF_NOTIFYITEMDRAW
    if (drawStage != 0x00010001) ; CDDS_ITEMPREPAINT
        return

    itemOffset := (A_PtrSize = 8) ? 56 : 36
    if (NumGet(lParam, itemOffset, "uptr") != 3) ; TBCD_CHANNEL
        return
    hdcOffset := (A_PtrSize = 8) ? 32 : 16
    rectOffset := (A_PtrSize = 8) ? 40 : 20
    CPDrawControllerColorGradient(NumGet(lParam, hdcOffset, "ptr"), sourceHwnd
        , lParam + rectOffset, CPControllerColorGradientSliders[sourceHwnd])
    return 0x00000004 ; CDRF_SKIPDEFAULT
}

CPRegisterControllerColorGradients(hueSlider, saturationSlider, brightnessSlider) {
    global CPControllerColorGradientSliders, CPControllerColorGradientMessageRegistered
    sliderSet := Map("hue", hueSlider, "saturation", saturationSlider, "brightness", brightnessSlider)
    for kind, slider in sliderSet {
        sliderData := Map("kind", kind, "hue", hueSlider
            , "saturation", saturationSlider, "brightness", brightnessSlider)
        CPControllerColorGradientSliders[slider.Hwnd] := sliderData
    }
    if !CPControllerColorGradientMessageRegistered {
        OnMessage(0x004E, CPControllerColorGradientCustomDraw) ; WM_NOTIFY
        CPControllerColorGradientMessageRegistered := true
    }
}

CPUnregisterControllerColorGradients(sliders*) {
    global CPControllerColorGradientSliders
    for slider in sliders {
        sliderHwnd := 0
        if IsObject(slider) {
            try sliderHwnd := slider.Hwnd
        } else {
            sliderHwnd := slider
        }
        if (sliderHwnd && CPControllerColorGradientSliders.Has(sliderHwnd))
            CPControllerColorGradientSliders.Delete(sliderHwnd)
    }
}

CPControllerColorDeferredGradientRedraw(hueHwnd, saturationHwnd, brightnessHwnd) {
    for sliderHwnd in [hueHwnd, saturationHwnd, brightnessHwnd] {
        if (!sliderHwnd || !DllCall("user32\IsWindow", "ptr", sliderHwnd, "int"))
            continue
        ; Trackbars cache an unfocused custom-drawn channel even after ordinary
        ; invalidation. WM_THEMECHANGED makes the control request it again.
        try DllCall("user32\SendMessage", "ptr", sliderHwnd, "uint", 0x031A
            , "ptr", 0, "ptr", 0) ; WM_THEMECHANGED
        try DllCall("user32\RedrawWindow", "ptr", sliderHwnd, "ptr", 0, "ptr", 0
            , "uint", 0x0001 | 0x0004 | 0x0100) ; INVALIDATE | ERASE | UPDATENOW
    }
}

CPControllerColorUpdatePreview(preview, hueSlider, saturationSlider, brightnessSlider
    , hueValue, saturationValue, brightnessValue, *) {
    colorHex := CPColorHSVToHex(hueSlider.Value, saturationSlider.Value, brightnessSlider.Value)
    preview.Opt("+Background" colorHex)
    hueValue.Text := Round(hueSlider.Value)
    saturationValue.Text := Round(saturationSlider.Value) "%"
    brightnessValue.Text := Round(brightnessSlider.Value) "%"
    try preview.Redraw()
    ; Run after the active trackbar's Change notification returns. Windows can
    ; otherwise defer repainting the two tracks that do not currently have focus.
    SetTimer(CPControllerColorDeferredGradientRedraw.Bind(
        hueSlider.Hwnd, saturationSlider.Hwnd, brightnessSlider.Hwnd), -1)
    return colorHex
}

CPControllerColorNavigate(direction, hueSlider, saturationSlider, brightnessSlider
    , applyButton, cancelButton, *) {
    focusHwnd := DllCall("user32\GetFocus", "ptr")
    sliders := [hueSlider, saturationSlider, brightnessSlider]
    for index, slider in sliders {
        if (focusHwnd != slider.Hwnd)
            continue
        if (direction = "Left" || direction = "Right")
            SendEvent("{" direction "}")
        else if (direction = "Up")
            (index = 1 ? hueSlider : sliders[index - 1]).Focus()
        else if (direction = "Down")
            (index = sliders.Length ? applyButton : sliders[index + 1]).Focus()
        return
    }

    if (focusHwnd = applyButton.Hwnd) {
        if (direction = "Up")
            brightnessSlider.Focus()
        else if (direction = "Right")
            cancelButton.Focus()
        return
    }
    if (focusHwnd = cancelButton.Hwnd) {
        if (direction = "Up")
            brightnessSlider.Focus()
        else if (direction = "Left")
            applyButton.Focus()
        return
    }
    hueSlider.Focus()
}

CPControllerColorActivate(hueSlider, saturationSlider, brightnessSlider, applyButton, cancelButton, *) {
    focusHwnd := DllCall("user32\GetFocus", "ptr")
    if (focusHwnd = applyButton.Hwnd)
        SendMessage(0x00F5, 0, 0, applyButton.Hwnd) ; BM_CLICK
    else if (focusHwnd = cancelButton.Hwnd)
        SendMessage(0x00F5, 0, 0, cancelButton.Hwnd)
}

CPControllerColorDialog(initHex, dialogTitle := "Adjust color") {
    global ui
    hsv := CPColorHexToHSV(initHex)
    result := ""
    closed := false
    dlg := Gui("+Owner" ui.Hwnd " +AlwaysOnTop", dialogTitle)
    dlg.MarginX := 18, dlg.MarginY := 16
    dlg.SetFont("s10", "Segoe UI")

    preview := CPRegisterColorSwatch(
        dlg.Add("Text", "xm w390 h54 Border Background" initHex), "", false)
    previewHwnd := preview.Hwnd
    dlg.Add("Text", "xm y+16 w82", "Hue:")
    hueSlider := dlg.Add("Slider", "x+8 yp-4 w250 Range0-359 ToolTip")
    hueSlider.Value := hsv["h"]
    hueValue := dlg.Add("Text", "x+8 yp+4 w42 Right", hsv["h"])

    dlg.Add("Text", "xm y+14 w82", "Saturation:")
    saturationSlider := dlg.Add("Slider", "x+8 yp-4 w250 Range0-100 ToolTip")
    saturationSlider.Value := hsv["s"]
    saturationValue := dlg.Add("Text", "x+8 yp+4 w42 Right", hsv["s"] "%")

    dlg.Add("Text", "xm y+14 w82", "Brightness:")
    brightnessSlider := dlg.Add("Slider", "x+8 yp-4 w250 Range0-100 ToolTip")
    brightnessSlider.Value := hsv["v"]
    brightnessValue := dlg.Add("Text", "x+8 yp+4 w42 Right", hsv["v"] "%")

    applyButton := dlg.Add("Button", "xm y+20 w120 Default", "Apply")
    cancelButton := dlg.Add("Button", "x+8 w100", "Cancel")
    updatePreview := CPControllerColorUpdatePreview.Bind(preview, hueSlider, saturationSlider
        , brightnessSlider, hueValue, saturationValue, brightnessValue)
    hueSlider.OnEvent("Change", updatePreview)
    saturationSlider.OnEvent("Change", updatePreview)
    brightnessSlider.OnEvent("Change", updatePreview)
    CPRegisterControllerColorGradients(hueSlider, saturationSlider, brightnessSlider)
    gradientSliderHwnds := [hueSlider.Hwnd, saturationSlider.Hwnd, brightnessSlider.Hwnd]

    finish := (colorValue) => (result := colorValue, closed := true, dlg.Destroy())
    applyButton.OnEvent("Click", (*) => finish.Call(CPColorHSVToHex(
        hueSlider.Value, saturationSlider.Value, brightnessSlider.Value)))
    cancelButton.OnEvent("Click", (*) => finish.Call(""))
    dlg.OnEvent("Escape", (*) => finish.Call(""))
    dlg.OnEvent("Close", (*) => finish.Call(""))

    dialogHotIf := "ahk_id " dlg.Hwnd
    dialogArrowHotkeys := Map("$Up", "Up", "$Down", "Down", "$Left", "Left", "$Right", "Right")
    HotIfWinActive(dialogHotIf)
    for keyName, direction in dialogArrowHotkeys
        try Hotkey(keyName, CPControllerColorNavigate.Bind(direction, hueSlider, saturationSlider
            , brightnessSlider, applyButton, cancelButton), "On")
    try Hotkey("$Enter", CPControllerColorActivate.Bind(hueSlider, saturationSlider
        , brightnessSlider, applyButton, cancelButton), "On")
    try Hotkey("$NumpadEnter", CPControllerColorActivate.Bind(hueSlider, saturationSlider
        , brightnessSlider, applyButton, cancelButton), "On")
    HotIfWinActive()

    try {
        dlg.Show("AutoSize Center")
        CPApplyOwnedDialogTheme(dlg)
        updatePreview.Call()
        hueSlider.Focus()
        while !closed
            Sleep(30)
    } finally {
        CPUnregisterColorSwatch(previewHwnd)
        CPUnregisterControllerColorGradients(gradientSliderHwnds*)
        HotIfWinActive(dialogHotIf)
        for keyName, direction in dialogArrowHotkeys
            try Hotkey(keyName, "Off")
        try Hotkey("$Enter", "Off")
        try Hotkey("$NumpadEnter", "Off")
        HotIfWinActive()
    }
    return result
}

CPAdjustColorSwatchWithController(swatchHwnd) {
    global boxBgHex, txtHex, nameHex, boxBgHex_EW, txtHex_EW
    target := CPColorSwatchTarget(swatchHwnd)
    switch target {
        case "translator:bg":
            initialColor := boxBgHex, dialogTitle := "Adjust Translator window color"
        case "translator:txt":
            initialColor := txtHex, dialogTitle := "Adjust Translator text color"
        case "translator:name":
            initialColor := nameHex, dialogTitle := "Adjust speaker-name color"
        case "explainer:bg":
            initialColor := boxBgHex_EW, dialogTitle := "Adjust Explainer window color"
        case "explainer:txt":
            initialColor := txtHex_EW, dialogTitle := "Adjust Explainer text color"
        default:
            return
    }

    selectedColor := CPControllerColorDialog(initialColor, dialogTitle)
    if (selectedColor = "")
        return
    if InStr(target, "translator:") = 1
        ApplyColorValue(SubStr(target, StrLen("translator:") + 1), selectedColor)
    else
        ApplyColorValue_EW(SubStr(target, StrLen("explainer:") + 1), selectedColor)
}

PickColorDialogDarkHook(dialogHwnd, msg, wParam, lParam) {
    global controlDarkMode, CPThemeBrushWindow, CPThemeBrushSurface
    if !controlDarkMode
        return 0

    if (msg = 0x0110) { ; WM_INITDIALOG
        CPApplyDarkTitleBar(dialogHwnd, true)
        CPSetPreferredAppDarkMode(true, dialogHwnd)
        try DllCall("uxtheme\SetWindowTheme", "ptr", dialogHwnd, "wstr", "DarkMode_Explorer", "ptr", 0)

        oldDetectHidden := A_DetectHiddenWindows
        try {
            DetectHiddenWindows true
            for controlHwnd in WinGetControlsHwnd("ahk_id " dialogHwnd) {
                controlClass := ""
                try controlClass := WinGetClass("ahk_id " controlHwnd)
                if (controlClass = "Edit") {
                    try DllCall("uxtheme\SetWindowTheme", "ptr", controlHwnd, "wstr", "DarkMode_CFD", "ptr", 0)
                } else if (controlClass = "Button" || controlClass = "ScrollBar") {
                    try DllCall("uxtheme\SetWindowTheme", "ptr", controlHwnd, "wstr", "DarkMode_Explorer", "ptr", 0)
                }
            }
        } finally {
            DetectHiddenWindows oldDetectHidden
        }
        try DllCall("user32\RedrawWindow", "ptr", dialogHwnd, "ptr", 0, "ptr", 0, "uint", 0x185)
        return 0
    }

    if (msg = 0x0136) ; WM_CTLCOLORDLG
        return CPThemeBrushWindow

    if (msg = 0x0133 || msg = 0x0134 || msg = 0x0135 || msg = 0x0138) {
        colors := CPPalette(true)
        DllCall("gdi32\SetTextColor", "ptr", wParam, "uint", CPColorRef(colors["text"]))
        if (msg = 0x0133 || msg = 0x0134) { ; Edit / ListBox
            DllCall("gdi32\SetBkColor", "ptr", wParam, "uint", CPColorRef(colors["surface"]))
            return CPThemeBrushSurface
        }
        DllCall("gdi32\SetBkMode", "ptr", wParam, "int", 1) ; TRANSPARENT
        return CPThemeBrushWindow
    }
    return 0
}

PickColorDialog(initHex := "FFFFFF") {
    global ui, controlDarkMode
    rgb := Integer("0x" initHex)
    bgr := ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
    ccSize := (A_PtrSize = 8 ? 72 : 36)
    cc := Buffer(ccSize, 0)
    custom := Buffer(16*4, 0)
    NumPut("UInt", ccSize, cc, 0)
    NumPut("Ptr", ui.Hwnd, cc, A_PtrSize)
    NumPut("Ptr", 0, cc, 2*A_PtrSize)
    NumPut("UInt", bgr, cc, 3*A_PtrSize)
    NumPut("Ptr", custom.Ptr, cc, 4*A_PtrSize)
    flags := 0x00000001 | 0x00000002
    colorHook := 0
    if controlDarkMode {
        CPRefreshThemeBrushes()
        colorHook := CallbackCreate(PickColorDialogDarkHook)
        flags |= 0x00000010 ; CC_ENABLEHOOK
        NumPut("Ptr", colorHook, cc, (A_PtrSize = 8 ? 56 : 28))
    }
    NumPut("UInt", flags, cc, (A_PtrSize=8 ? 40 : 20))
    try ret := DllCall("Comdlg32\ChooseColorW", "Ptr", cc.Ptr, "Int")
    finally {
        if colorHook
            CallbackFree(colorHook)
    }
    if (ret = 0)
        return ""
    gotBGR := NumGet(cc, 3*A_PtrSize, "UInt")
    gotRGB := ((gotBGR & 0xFF) << 16) | (gotBGR & 0xFF00) | ((gotBGR >> 16) & 0xFF)
    return Format("{:06X}", gotRGB)
}

PickAndApply(which) {
    global boxBgHex,txtHex,nameHex

    colorCur := (which="bg")    ? boxBgHex
             : (which="name")  ? nameHex
             :                   txtHex

    got := PickColorDialog(colorCur)
    If (got = "")
        Return

    ApplyColorValue(which, got)
}

ApplyColorValue(which, got) {
    global boxBgHex,txtHex,nameHex
    global rectBg,rectTxt,rectName

    if (which="bg") {
        boxBgHex := got
        rectBg.Opt("Background" . got)
    } else if (which="name") {
        nameHex := got
        if IsSet(rectName)
            rectName.Opt("Background" . got)
    } else {
        txtHex := got
        rectTxt.Opt("Background" . got)
    }

    SyncUnifiedWindowAppearance()
    SaveAll()
    RefreshColorSwatches()
    DbgCP("Color change '" which "' -> " got)
    SendOverlayTheme()
}

FontChanged(ctrl, *) {
    global fontName
    fontName := ctrl.Text
    SaveAll()
    DbgCP("FontChanged -> " fontName)
    SendOverlayTheme()
}

FontSizeCommit(ctrl, *) {
    global fontSize
    txt := Trim(ctrl.Value)
    if (txt = "") {
        ctrl.Value := fontSize
        return
    }
    val := Integer(txt)
    if (val < 6)
        val := 6
    else if (val > 128)
        val := 128
    if (val = fontSize) {
        ctrl.Value := val
        return
    }
    fontSize := val
    ctrl.Value := fontSize
    SaveAll()
    DbgCP("FontSizeCommit -> " fontSize)
    SendOverlayTheme()
}

FontBoldChanged(ctrl, *) {
    global fontBold
    fontBold := ctrl.Value ? 1 : 0
    SaveAll()
    DbgCP("FontBoldChanged -> " fontBold)
    SendOverlayTheme()
}

; =========================
; EXPLAINER handlers (separate state)
; =========================
HandleTransparencyChange_EW(sliderCtrl) {
    global lblTransPct_EW, overlayTrans_EW
    val := sliderCtrl.Value
    overlayTrans_EW := val
    pct := Round(val / 255 * 100)
    lblTransPct_EW.Value := pct . "%"
        SaveAll(), ClearDirty()
    try WinSetTransparent(overlayTrans_EW, "Explainer")
    SendOverlayTheme()
}

PickAndApply_EW(which) {
    global boxBgHex_EW,txtHex_EW

    colorCur := (which="bg") ? boxBgHex_EW : txtHex_EW
    got := PickColorDialog(colorCur)
    if (got = "")
        return

    ApplyColorValue_EW(which, got)
}

ApplyColorValue_EW(which, got) {
    global boxBgHex_EW,txtHex_EW
    global rectBg_EW,rectTxt_EW

    if (which="bg") {
        boxBgHex_EW := got
        rectBg_EW.Opt("Background" . got)
    } else {
        txtHex_EW := got
        rectTxt_EW.Opt("Background" . got)
    }
    SyncUnifiedWindowAppearance()
    SaveAll()
    RefreshColorSwatches_EW()
    DbgCP("EW Color change '" which "' -> " got)
    SendOverlayTheme()
}

FontChanged_EW(ctrl, *) {
    global fontName_EW
    fontName_EW := ctrl.Text
    SaveAll()
    DbgCP("EW FontChanged -> " fontName_EW)
    SendOverlayTheme()
}

FontSizeCommit_EW(ctrl, *) {
    global fontSize_EW
    txt := Trim(ctrl.Value)
    if (txt = "") {
        ctrl.Value := fontSize_EW
        return
    }
    val := Integer(txt)
    if (val < 6)
        val := 6
    if (val > 200)
        val := 200
    fontSize_EW := val
    ctrl.Value := val
    SaveAll()
    DbgCP("EW FontSizeCommit -> " fontSize_EW)
    SendOverlayTheme()
}

FontBoldChanged_EW(ctrl, *) {
    global fontBold_EW
    fontBold_EW := ctrl.Value ? 1 : 0
    SaveAll()
    DbgCP("EW FontBoldChanged -> " fontBold_EW)
    SendOverlayTheme()
}

; =========================
; Explainer bounds helpers
; =========================
ApplyExplainerBounds() {
    global ewX, ewY, ewW, ewH
    old := A_TitleMatchMode
    SetTitleMatchMode 3
    hwnd := WinExist("Explainer")
    SetTitleMatchMode old
    if !hwnd
        return
    ; only apply if we have stored values
    if (ewX != "" && ewY != "" && ewW != "" && ewH != "") {
        try WinMove ewX, ewY, ewW, ewH, "ahk_id " hwnd
    }
}

SaveExplainerBoundsIfChanged() {
    global ewX, ewY, ewW, ewH
    global ew_lastX, ew_lastY, ew_lastW, ew_lastH
    global iniPath

    old := A_TitleMatchMode
    SetTitleMatchMode 3
    hwnd := WinExist("Explainer")
    SetTitleMatchMode old
    if !hwnd {
        return
    }

    x := 0, y := 0, w := 0, h := 0
    try WinGetPos &x, &y, &w, &h, "ahk_id " hwnd
    if (x = "" || y = "" || w = "" || h = "")
        return

    ; first time? seed the public vars so UI shows correct data if needed
    if (ewX = "") ewX := x
    if (ewY = "") ewY := y
    if (ewW = "") ewW := w
    if (ewH = "") ewH := h

    changed := (x != ew_lastX) || (y != ew_lastY) || (w != ew_lastW) || (h != ew_lastH)
    if !changed
        return

    ew_lastX := x, ew_lastY := y, ew_lastW := w, ew_lastH := h
    ewX := x, ewY := y, ewW := w, ewH := h

    try {
        IniWrite(x, iniPath, "explainer_bounds", "x")
        IniWrite(y, iniPath, "explainer_bounds", "y")
        IniWrite(w, iniPath, "explainer_bounds", "w")
        IniWrite(h, iniPath, "explainer_bounds", "h")
    }
}

StartExplainerBoundsWatcher() {
    global ew_bounds_watch_running
    if (ew_bounds_watch_running)
        return
    SetTimer SaveExplainerBoundsIfChanged, 700
    ew_bounds_watch_running := true
}

StopExplainerBoundsWatcher() {
    global ew_bounds_watch_running
    if (!ew_bounds_watch_running)
        return
    SetTimer SaveExplainerBoundsIfChanged, 0
    ew_bounds_watch_running := false
}

; =========================
; Explainer profiles (save/load/delete)
; =========================
EW_ProfilePath(name) {
    global profilesDir_EW
    return profilesDir_EW "\" RegExReplace(Trim(name), "[^\w\-\. ]", "_") ".ini"
}

EW_SaveProfile(name) {
    global boxBgHex_EW,bdrOutHex_EW,bdrInHex_EW,txtHex_EW
    global fontName_EW,fontSize_EW,fontBold_EW,bdrOutW_EW,bdrInW_EW,overlayTrans_EW
    global ewX,ewY,ewW,ewH

    p := EW_ProfilePath(name)
    SyncUnifiedWindowAppearance()
    try {
        ; theme
        IniWrite(overlayTrans_EW, p, "cfg_explainer", "overlayTrans")
        IniWrite(boxBgHex_EW,     p, "cfg_explainer", "boxBg")
        IniWrite(boxBgHex_EW,     p, "cfg_explainer", "bdrOut")
        IniWrite(boxBgHex_EW,     p, "cfg_explainer", "bdrIn")
        IniWrite(txtHex_EW,       p, "cfg_explainer", "txtColor")
        IniWrite(fontName_EW,     p, "cfg_explainer", "fontName")
        IniWrite(fontSize_EW,     p, "cfg_explainer", "fontSize")
        IniWrite(fontBold_EW,     p, "cfg_explainer", "fontBold")
        IniWrite(0,                p, "cfg_explainer", "bdrOutW")
        IniWrite(0,                p, "cfg_explainer", "bdrInW")

        ; bounds (optional; only if known)
        if (ewX != "")
            IniWrite(ewX, p, "explainer_bounds", "x")
        if (ewY != "")
            IniWrite(ewY, p, "explainer_bounds", "y")
        if (ewW != "")
            IniWrite(ewW, p, "explainer_bounds", "w")
        if (ewH != "")
            IniWrite(ewH, p, "explainer_bounds", "h")
    } Catch as ex {
        MsgBox "Failed to save profile:`n" e.Message
    }
    DbgCP("EW_SaveProfile -> " p)
}

EW_LoadProfile(name) {
    global boxBgHex_EW,bdrOutHex_EW,bdrInHex_EW,txtHex_EW
    global fontName_EW,fontSize_EW,fontBold_EW,bdrOutW_EW,bdrInW_EW,overlayTrans_EW
    global ewX,ewY,ewW,ewH
    global ui, slTrans_EW, lblTransPct_EW
    global rectBg_EW,rectTxt_EW
    global ddlFont_EW, edFSize_EW, udFSize_EW, chkFontBold_EW
    global iniPath, ddlProf_EW

    p := EW_ProfilePath(name)
    if !FileExist(p) {
        MsgBox "Profile not found: " p
        return
    }
    ; remember last used Explainer profile in control.ini
    try IniWrite(name, iniPath, "profiles", "explainer_last")
    try ddlProf_EW.Text := name

    overlayTrans_EW := Integer(IniRead(p, "cfg_explainer", "overlayTrans", overlayTrans_EW))
    boxBgHex_EW     := StrUpper(IniRead(p, "cfg_explainer", "boxBg",    boxBgHex_EW))
    txtHex_EW       := StrUpper(IniRead(p, "cfg_explainer", "txtColor", txtHex_EW))
    fontName_EW     := IniRead(p, "cfg_explainer", "fontName", fontName_EW)
    fontSize_EW     := Integer(IniRead(p, "cfg_explainer", "fontSize",  fontSize_EW))
    fontBold_EW     := Integer(IniRead(p, "cfg_explainer", "fontBold",  fontBold_EW)) ? 1 : 0
    SyncUnifiedWindowAppearance()

    ; bounds (may not exist in profile; keep current if empty)
    _x := IniRead(p, "explainer_bounds", "x", "")
    _y := IniRead(p, "explainer_bounds", "y", "")
    _w := IniRead(p, "explainer_bounds", "w", "")
    _h := IniRead(p, "explainer_bounds", "h", "")
    if (_x != "" && _y != "" && _w != "" && _h != "") {
        ewX := Integer(_x), ewY := Integer(_y), ewW := Integer(_w), ewH := Integer(_h)
    }

    ; reflect in UI
    slTrans_EW.Value := overlayTrans_EW
    try lblTransPct_EW.Value := Round(overlayTrans_EW / 255 * 100) . "%"
    try rectBg_EW.Opt("Background" . boxBgHex_EW)
    try rectTxt_EW.Opt("Background" . txtHex_EW)
    try ddlFont_EW.Text := fontName_EW
    try edFSize_EW.Value := fontSize_EW, udFSize_EW.Value := fontSize_EW
    try chkFontBold_EW.Value := fontBold_EW
    RefreshColorSwatches_EW()

    SaveAll()
    SendOverlayTheme()
    ; and if window is open, also apply bounds immediately
    ApplyExplainerBounds()
    DbgCP("EW_LoadProfile <- " p)
}

EW_DeleteProfile(name) {
    p := EW_ProfilePath(name)
    if FileExist(p) {
        try FileDelete(p)
        DbgCP("EW_DeleteProfile x " p)
    }
}

EW_ListProfiles() {
    global profilesDir_EW
    list := []
    Loop Files profilesDir_EW "\*.ini", "F" {
        list.Push( StrReplace(A_LoopFileName, ".ini") )
    }
    return list
}

; ---- fonts ------------------------------------------------------
; --- Private font loader + TTF/OTF name reader (fonts subfolder) ----------------
EnsurePrivateFontsLoaded(){
    static loaded := false
    global __PRIVATE_FONT_NAMES
    if (loaded)
        return
    __PRIVATE_FONT_NAMES := []

    dir := A_ScriptDir "\fonts"
    if !DirExist(dir) {
        loaded := true
        return
    }

        exts := ["ttf","otf","ttc"]
    for ext in exts {
        Loop Files, dir "\*." ext, "F" {
            f := A_LoopFileFullPath
            try DllCall("AddFontResourceEx", "str", f, "uint", 0x10, "ptr", 0)  ; FR_PRIVATE

            added := false
            for name in TTF_GetFamilyNames(f) {
                if (name != "") {
                    __PRIVATE_FONT_NAMES.Push(name)
                    added := true
                }
            }
            ; Fallback: if name table parsing failed, add the fileâ€™s base name
            if (!added) {
                base := RegExReplace(A_LoopFileName, "\.(ttf|otf|ttc)$", "",, 1)
                base := RegExReplace(base, "i)-(Regular|Bold|Italic|Oblique)$")
                if (base = "PressStart2P")
                    base := "Press Start 2P"
                if (base != "")
                    __PRIVATE_FONT_NAMES.Push(base)
            }
        }
    }
    loaded := true
}

; Return an Array of family names from a TTF/OTF/TTC (minimal 'name' table parser)
TTF_GetFamilyNames(path){
    out := []
    try {
        f := FileOpen(path, "r")
        if (!f)
            return out
        size := f.Length
        buf := Buffer(size, 0)
        f.RawRead(buf, size)
        f.Close()

        ; OpenType numeric fields are big-endian, unlike native Windows integers.
        Num16(off) => (NumGet(buf, off, "UChar") << 8)
                    | NumGet(buf, off + 1, "UChar")
        Num32(off) => (NumGet(buf, off, "UChar") << 24)
                    | (NumGet(buf, off + 1, "UChar") << 16)
                    | (NumGet(buf, off + 2, "UChar") << 8)
                    | NumGet(buf, off + 3, "UChar")

        ; sfnt header
        numTables := Num16(4)
        ; table records start at 12, 16 bytes each
        nameOff := -1, nameLen := 0
        base := 12
        Loop numTables {
            off := base + (A_Index-1)*16
            tag := StrGet(buf.Ptr + off, 4, "CP0")
            ; offset and length fields
            toff := Num32(off+8)
            tlen := Num32(off+12)
            if (tag = "name") {
                nameOff := toff, nameLen := tlen
                break
            }
        }
        if (nameOff < 0)
            return out

        nameFormat := Num16(nameOff)
        nameCount  := Num16(nameOff + 2)
        strOff := Num16(nameOff + 4) + nameOff

        ; gather all NameID 1 (Font Family)
        seen := Map()
        Loop nameCount {
            rec := nameOff + 6 + (A_Index-1)*12
            platformID := Num16(rec)
            encodingID := Num16(rec+2)
            languageID := Num16(rec+4)
            nameID     := Num16(rec+6)
            length     := Num16(rec+8)
            roff       := Num16(rec+10)

            if (nameID != 1) ; family name
                continue
            p := strOff + roff
            if (p+length > buf.Size)
                continue

            s := ""
            ; Windows/Unicode (UTF-16BE) â†’ CP1201
            if (platformID = 3) {
                try s := StrGet(buf.Ptr + p, length//2, 1201) ; UTF-16BE
            } else if (platformID = 0) { ; Unicode
                try s := StrGet(buf.Ptr + p, length//2, 1201)
            } else {
                ; Mac/others â€“ treat as ANSI
                try s := StrGet(buf.Ptr + p, length, "CP1252")
            }
            s := Trim(s)
            if (s != "" && !seen.Has(s)) {
                seen[s] := true
                out.Push(s)
            }
        }
    } catch {
        ; ignore parse errors -> return whatever we got
    }
    return out
}

GetInstalledFonts(){
    global __PRIVATE_FONT_NAMES
    ; Merge registry fonts + private fonts loaded from .\fonts
    EnsurePrivateFontsLoaded()
    seen := Map()

    ; 1) System-installed (registry)
    for root in ["HKLM","HKCU"] {
        key := root "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        Loop Reg, key, "V" {
            n := A_LoopRegName
            n := RegExReplace(n, "\s*\(.*\)$", "")
            n := RegExReplace(n, "\b(Regular|Bold|Italic|Oblique|Condensed|Extended|Extra|ExtraBold|ExtraLight|Heavy|Demi)\b", "")
            n := RegExReplace(n, "\s{2,}", " ")
            n := Trim(n)
            if (n != "")
                seen[n] := true
        }
    }

    ; 2) Private fonts
    global __PRIVATE_FONT_NAMES
    if (IsSet(__PRIVATE_FONT_NAMES)) {
        for _, n in __PRIVATE_FONT_NAMES {
            if (n != "")
                seen[n] := true
        }
    }

    tmp := ""
    for k, _ in seen
        tmp .= k "`n"
    tmp := RTrim(tmp, "`n")
    if (tmp != "")
        tmp := Sort(tmp)
    out := []
    for name in StrSplit(tmp, "`n")
        out.Push(name)
    return out.Length ? out : ["Segoe UI","Arial","Consolas"]
}

LoadFontsIntoCombo(){
    global ddlFont, fontName
    EnsurePrivateFontsLoaded()
    SendMessage(0x14B, 0, 0, ddlFont.Hwnd)  ; CB_RESETCONTENT
    fonts := GetInstalledFonts()
    ddlFont.Add(fonts)
    desiredFont := ArrIndexOf(fonts, fontName) ? fontName : "Segoe UI"
    selectedFont := SetComboToExistingItem(ddlFont, fonts, desiredFont)
    if (selectedFont != "")
        fontName := selectedFont
}

LoadFontsIntoCombo_EW(){
    global ddlFont_EW, fontName_EW
    if !IsSet(ddlFont_EW) || !ddlFont_EW
        return
    EnsurePrivateFontsLoaded()
    SendMessage(0x14B, 0, 0, ddlFont_EW.Hwnd)  ; CB_RESETCONTENT
    fonts := GetInstalledFonts()
    ddlFont_EW.Add(fonts)
    desiredFont := ArrIndexOf(fonts, fontName_EW) ? fontName_EW : "Segoe UI"
    selectedFont := SetComboToExistingItem(ddlFont_EW, fonts, desiredFont)
    if (selectedFont != "")
        fontName_EW := selectedFont
}

; ---- prompt profile helpers -----------------------------------
PromptFilePath(name) {
    global promptsDir
    return promptsDir "\" name ".txt"
}

; --- UI helper: bind a tooltip to a control (clean + reusable)
; --- UI helper: set a native tooltip on a control (hover to see it)
TooltipBind(ctrl, text) {
    try ctrl.SetTip(text)  ; show this tip on hover
    ; To remove later: ctrl.SetTip("")  (optional, not used here)
}

ExplainPromptFilePath() {
    global promptsDir
    return promptsDir "\explain_prompt.txt"
}

OpenExplainPromptEditor(*) {
    path := ExplainPromptFilePath()
    txt := ""
    try txt := FileExist(path) ? FileRead(path, "UTF-8") : ""

    g := Gui("+Resize", "Edit Explanation Prompt")
    edt := g.Add("Edit", "xm ym w680 h420 WantTab WantReturn Wrap", txt)
    ; ^ WantReturn ensures Enter inserts a line break, like your other editor
    btnSave  := g.Add("Button", "xm y+8 w100", "Save")
    btnSave.OnEvent("Click", (*) => (
    (FileExist(path) ? (FileCopy(path, path ".bak", true)) : 0),
    f := FileOpen(path, "w", "UTF-8"),
    f.Write(edt.Value),
    f.Close(),
    Toast("Saved explanation prompt"),
    DbgCP("Explanation prompt saved to: " path)
))
    btnClose := g.Add("Button", "x+8 yp w100", "Close")

    btnClose.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Size", (gui, mm, w, h) => (
        edt.Move(, , Max(300, w-40), Max(180, h-90)),
        y := h - 52,
        btnSave.Move(20, y, 100, 32),
        btnClose.Move(20+100+8, y, 100, 32)
    ))
    CPShowTextEditorDialog(g, edt)
}
ListPromptProfiles() {
    global promptsDir
    out := []
    Loop Files promptsDir "\*.txt" {
        n := A_LoopFileName
        n := RegExReplace(n, "\.txt$", "")
        out.Push(n)
    }
    if out.Length {
        txt := ""
        for n in out
            txt .= n "`n"
        txt := RTrim(txt, "`n")
        txt := Sort(txt)
        out := []
        for n in StrSplit(txt, "`n")
            out.Push(n)
    }
    return out
}

; ---- EXPLANATION prompt profile helpers (separate folder) ----
ExplainProfilePath(name) {
    global explainPromptsDir
    return explainPromptsDir "\" name ".txt"
}

ListExplainPromptProfiles() {
    global explainPromptsDir
    out := []
    Loop Files explainPromptsDir "\*.txt" {
        n := A_LoopFileName
        n := RegExReplace(n, "\.txt$", "")
        out.Push(n)
    }
    if out.Length {
        txt := ""
        for n in out
            txt .= n "`n"
        txt := RTrim(txt, "`n")
        txt := Sort(txt)
        out := []
        for n in StrSplit(txt, "`n")
            out.Push(n)
    }
    return out
}

RefreshExplainPromptProfilesList(select := "") {
    global ddlEPr, explainPromptProfile
    list := ListExplainPromptProfiles()
    ddlEPr.Delete()
    if (list.Length) {
        ddlEPr.Add(list)
        selIdx := (select!="") ? ArrayIndexOf(list, select) : ArrayIndexOf(list, explainPromptProfile)
        if (selIdx = 0)
            selIdx := 1
        ddlEPr.Choose(selIdx)
    } else {
        ; list empty â€“ donâ€™t assign a non-existent item
        try ddlEPr.Text := ""   ; clear display safely
        ; (We still remember explainPromptProfile in INI; once a file exists, it will be selected.)
    }
}

ExplainPromptChanged(*) {
    global explainPromptProfile, ddlEPr, iniPath
    explainPromptProfile := Trim(ddlEPr.Text)
    IniWrite(explainPromptProfile, iniPath, "cfg", "explainPromptProfile")
}

OpenExplainPromptEditor_Multi(*) {
    global ddlEPr
    name := Trim(ddlEPr.Text)
    if (name = "")
        name := "default"
    path := ExplainProfilePath(name)
    txt := ""
    try txt := FileExist(path) ? FileRead(path, "UTF-8") : ""

    g := Gui("+Resize", "Edit Explanation Prompt - " name)
    edt := g.Add("Edit", "xm ym w680 h420 WantTab WantReturn Wrap", txt)
    btnSave  := g.Add("Button", "xm y+8 w100", "Save")
    btnSave.OnEvent("Click", (*) => (
        (FileExist(path) ? (FileCopy(path, path ".bak", true)) : 0),
        f := FileOpen(path, "w", "UTF-8"),
        f.Write(edt.Value),
        f.Close(),
        Toast("Saved EXPLAIN prompt: " name)
    ))
    btnClose := g.Add("Button", "x+8 yp w100", "Close")
    btnClose.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Size", (gui, mm, w, h) => (
        edt.Move(, , Max(300, w-40), Max(180, h-90)),
        y := h - 52,
        btnSave.Move(20, y, 100, 32),
        btnClose.Move(20+100+8, y, 100, 32)
    ))
    CPShowTextEditorDialog(g, edt)
}

NewExplainPromptProfile(*) {
    global ddlEPr
    ib := InputBox("Enter a name for the new EXPLANATION prompt:", "New EXPLAIN prompt", "w320 h140")
    if (ib.Result = "Cancel")
        return
    name := Trim(ib.Value)
    if (name = "") {
        MsgBox("Please enter a non-empty name.", "New EXPLAIN prompt", "OK Icon!")
        return
    }
    name := RegExReplace(name, '[\\/:*?"<>|]+', "_")
    if RegExMatch(name, 'i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$')
        name := "_" name
    path := ExplainProfilePath(name)
    if FileExist(path) {
        MsgBox("A prompt with that name already exists.",, "OK Icon!")
        return
    }
    FileAppend("You are a friendly tutor for learners of Japanese." . "`r`n`r`nJapanese:" . "`r`n{jp}", path, "UTF-8")
    RefreshExplainPromptProfilesList(name)
}

DeleteExplainPromptProfile(*) {
    global ddlEPr
    name := Trim(ddlEPr.Text)
    if (name = "") {
        MsgBox("No EXPLAIN prompt selected.",, "OK Icon!")
        return
    }
    if (MsgBox("Delete EXPLAIN prompt '" name "'?",, "YesNo Icon!")!="Yes")
        return
    path := ExplainProfilePath(name)
    try FileDelete(path)
    RefreshExplainPromptProfilesList()
}

RefreshPromptProfilesList(select := "") {
    global ddlPrompt, promptProfile
    list := ListPromptProfiles()
    ddlPrompt.Delete()
    if (list.Length) {
        ddlPrompt.Add(list)
        selIdx := (select!="") ? ArrayIndexOf(list, select) : ArrayIndexOf(list, promptProfile)
        if (selIdx = 0)
            selIdx := 1
        ddlPrompt.Choose(selIdx)
    } else {
        ; list empty â€“ donâ€™t assign a non-existent item
        try ddlPrompt.Text := ""   ; safe clear (equivalently: ddlPrompt.Choose(0))
    }
}
OpenPromptEditor(*) {
    global ddlPrompt
    name := Trim(ddlPrompt.Text)
    if (name = "")
        name := "default"
    path := PromptFilePath(name)
    txt := ""
    try txt := FileExist(path) ? FileRead(path, "UTF-8") : ""

    g := Gui("+Resize", "Edit Prompt - " name)
    edt := g.Add("Edit", "xm ym w680 h420 WantTab WantReturn Wrap", txt)
    ; ^^^^^^^^^^^^^^^ lets Enter insert a line break
    btnSave  := g.Add("Button", "xm y+8 w100", "Save")
    btnSave.OnEvent("Click", (*) => (
    (FileExist(path) ? (FileCopy(path, path ".bak", true)) : 0),
    f := FileOpen(path, "w", "UTF-8"),
    f.Write(edt.Value),
    f.Close(),
    Toast("Saved prompt: " name),
    DbgCP("Prompt saved: " name)
))

    btnClose := g.Add("Button", "x+8 yp w100", "Close")   ; yp = same Y as Save
    btnClose.OnEvent("Click", (*) => g.Destroy())

    g.OnEvent("Size", (gui, mm, w, h) => (
        edt.Move(, , Max(300, w-40), Max(180, h-90)),
        y := h - 52,                     ; bottom padding
        btnSave.Move(20, y, 100, 32),    ; left-aligned
        btnClose.Move(20+100+8, y, 100, 32)
    ))
    CPShowTextEditorDialog(g, edt)
}
NewPromptProfile(*) {
    global ddlPrompt
    ib := InputBox("Enter a name for the new prompt profile:", "New prompt", "w320 h140")
    if (ib.Result = "Cancel")
        return
    name := Trim(ib.Value)
    if (name = "") {
        MsgBox("Please enter a non-empty name.", "New prompt", "OK Icon!")
        return
    }
    name := RegExReplace(name, '[\\/:*?"<>|]+', "_")
    if RegExMatch(name, 'i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$')
        name := "_" name
    path := PromptFilePath(name)
    if FileExist(path) {
        if (MsgBox("Prompt '" name "' already exists.`nOpen editor?", "New prompt", "YesNo Icon!")="Yes") {
            RefreshPromptProfilesList(name)   ; select it safely
            OpenPromptEditor()
        }
        return
    }
    FileAppend("", path, "UTF-8")
    RefreshPromptProfilesList(name)   ; rebuild list and select the new name
    OpenPromptEditor()
}
DeletePromptProfile(*) {
    global ddlPrompt
    name := Trim(ddlPrompt.Text)
    if (name = "") {
        MsgBox("No profile selected.",, "OK Icon!")
        return
    }
    if (MsgBox("Delete prompt '" name "'?",, "YesNo Icon!")!="Yes")
        return
    path := PromptFilePath(name)
    try FileDelete(path)
    RefreshPromptProfilesList()
    DbgCP("Prompt deleted: " name)
}

; ---- GLOSSARY profile helpers ---------------------------------
GlossaryProfileDir(name) {
    global glossariesDir
    return glossariesDir "\" name
}
GlossaryJP2ENPath(name) {
    return GlossaryProfileDir(name) "\jp2en.txt"
}
GlossaryEN2ENPath(name) {
    return GlossaryProfileDir(name) "\en2en.txt"
}

ListGlossaryProfiles() {
    global glossariesDir
    out := []

    ; collect folder names that contain either file
    if DirExist(glossariesDir) {
        Loop Files glossariesDir "\*", "D" {
            prof := A_LoopFileName
            if FileExist(glossariesDir "\" prof "\jp2en.txt")
             || FileExist(glossariesDir "\" prof "\en2en.txt")
                out.Push(prof)
        }
    }

    ; make sure "default" is available in the list (but do NOT create files now)
    if !ArrHas(out, "default")
        out.Push("default")

    if out.Length {
        txt := ""
        for n in out
            txt .= n "`n"
        txt := RTrim(txt, "`n")
        txt := Sort(txt)
        out := []
        for n in StrSplit(txt, "`n")
            out.Push(n)
    }
    return out
}

RefreshGlossaryProfilesList(selJP := "", selEN := "") {
    global ddlJPG, ddlENG
    SendMessage(0x14B, 0, 0, ddlJPG.Hwnd) ; CB_RESETCONTENT
    SendMessage(0x14B, 0, 0, ddlENG.Hwnd)

    lst := ListGlossaryProfiles()
    ddlJPG.Add(lst)
    ddlENG.Add(lst)

    if (selJP != "")
        ddlJPG.Text := selJP
    if (selEN != "")
        ddlENG.Text := selEN

    ; if selection was invalid/missing, default to "default" without creating files
    if (Trim(ddlJPG.Text) = "")
        ddlJPG.Text := "default"
    if (Trim(ddlENG.Text) = "")
        ddlENG.Text := "default"
}

OpenGlossaryEditor(kind := "jp") {
    global ddlJPG, ddlENG
    prof := (kind = "jp") ? Trim(ddlJPG.Text) : Trim(ddlENG.Text)
    if (prof = "")
        prof := "default"

    title := (kind = "jp") ? "Edit JP -> TL Glossary - " prof : "Edit TL -> TL Glossary - " prof
    path  := (kind = "jp") ? GlossaryJP2ENPath(prof) : GlossaryEN2ENPath(prof)

    ; ensure the profile folder exists (only when editing/saving), but do NOT auto-create other profiles
    if !DirExist(GlossaryProfileDir(prof))
        DirCreate(GlossaryProfileDir(prof))

    txt := ""
    try txt := FileExist(path) ? FileRead(path, "UTF-8") : "# One mapping per line: JP -> TL (or TL -> TL)`r`n"

   g := Gui("+Resize", title)
    edGloss := g.Add("Edit", "xm ym w700 h420 WantTab WantReturn Wrap", txt)
    btnSave  := g.Add("Button", "xm y+8 w100", "Save")
    btnClose := g.Add("Button", "x+8 yp w100", "Close")

    btnSave.OnEvent("Click", (*) => (
    SaveTextAtomic(path, edGloss.Value),
    Toast("Saved " ((kind="jp")?"JP -> TL":"TL -> TL") " glossary for profile '" prof "'")
))

    btnClose.OnEvent("Click", (*) => g.Destroy())

    g.OnEvent("Size", (gui, mm, w, h) => (
        edGloss.Move(, , Max(320, w-40), Max(160, h-90)),
        y := h - 52,
        btnSave.Move(20, y, 100, 32),
        btnClose.Move(130, y, 100, 32)
    ))
    CPShowTextEditorDialog(g, edGloss)
}

NewGlossaryProfile(*) {
    global ddlJPG, ddlENG, iniPath, jp2enGlossaryProfile, en2enGlossaryProfile
    ib := InputBox("Enter a name for the new glossary profile:", "New glossary", "w320 h140")
    if (ib.Result = "Cancel")
        return
    name := Trim(ib.Value)
    if (name = "") {
        MsgBox("Please enter a non-empty name.", "New glossary", "OK Icon!")
        return
    }
    name := RegExReplace(name, '[\\/:*?"<>|]+', "_")
    if RegExMatch(name, 'i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$')
        name := "_" name

    dir := GlossaryProfileDir(name)
    if DirExist(dir) {
        if (MsgBox("Profile '" name "' already exists.`nOpen editors?", "New glossary", "YesNo Icon!")="Yes") {
            RefreshGlossaryProfilesList(name, name)
            OpenGlossaryEditor("jp")
            OpenGlossaryEditor("en")
        }
        return
    }

    ; create both partner files so both rows immediately have it
    DirCreate(dir)
    FileAppend("# One mapping per line: JP -> TL`r`n", GlossaryJP2ENPath(name), "UTF-8")
    FileAppend("# One mapping per line: TL -> TL`r`n", GlossaryEN2ENPath(name), "UTF-8")

    ; select it in BOTH rows and persist
    RefreshGlossaryProfilesList(name, name)
    jp2enGlossaryProfile := name
    en2enGlossaryProfile := name
    IniWrite(jp2enGlossaryProfile, iniPath, "cfg", "jp2enGlossaryProfile")
    IniWrite(en2enGlossaryProfile, iniPath, "cfg", "en2enGlossaryProfile")

    OpenGlossaryEditor("jp")
    OpenGlossaryEditor("en")
}

DeleteGlossaryProfile(*) {
    global ddlJPG, ddlENG, iniPath, jp2enGlossaryProfile, en2enGlossaryProfile
    name := Trim(ddlJPG.Text)
    if (name = "")
        name := Trim(ddlENG.Text)
    if (name = "") {
        MsgBox("No profile selected.",, "OK Icon!")
        return
    }
    if (name = "default") {
        MsgBox("The 'default' profile cannot be deleted.",, "OK Icon!")
        return
    }
    if (MsgBox("Delete glossary profile '" name "' (both files)?",, "YesNo Icon!")!="Yes")
        return

    ; delete the entire folder for this profile
    dir := GlossaryProfileDir(name)
    try DirDelete(dir, true)

    ; if you just deleted the selected one, fall back to 'default' and persist
    if (jp2enGlossaryProfile = name)
        jp2enGlossaryProfile := "default"
    if (en2enGlossaryProfile = name)
        en2enGlossaryProfile := "default"
    IniWrite(jp2enGlossaryProfile, iniPath, "cfg", "jp2enGlossaryProfile")
    IniWrite(en2enGlossaryProfile, iniPath, "cfg", "en2enGlossaryProfile")

    RefreshGlossaryProfilesList(jp2enGlossaryProfile, en2enGlossaryProfile)
}

GlossaryChanged(kind) {
    global ddlJPG, ddlENG, iniPath, jp2enGlossaryProfile, en2enGlossaryProfile
    if (kind = "jp") {
        jp2enGlossaryProfile := Trim(ddlJPG.Text)
        if (jp2enGlossaryProfile = "")
            jp2enGlossaryProfile := "default"
        IniWrite(jp2enGlossaryProfile, iniPath, "cfg", "jp2enGlossaryProfile")
    } else {
        en2enGlossaryProfile := Trim(ddlENG.Text)
        if (en2enGlossaryProfile = "")
            en2enGlossaryProfile := "default"
        IniWrite(en2enGlossaryProfile, iniPath, "cfg", "en2enGlossaryProfile")
    }
}

; small helper you already use patterns like this across the file:
ArrHas(arr, val) {
    for v in arr
        if (v = val)
            return true
    return false
}

; ========= Atomic/Retry Save Helpers =========

; Text files (prompts, glossaries, profiles-as-text, etc.)
SaveTextAtomic(path, text, doBackup := true) {
    tmp := path ".tmp"

    ; ensure dir exists
    SplitPath(path, , &dir)
    if !DirExist(dir)
        DirCreate(dir)

    ; optional backup
    if doBackup && FileExist(path) {
        try FileCopy(path, path ".bak", true)
    }

    ; up to 5 retries for sharing violations (cloud sync / AV)
    loop 5 {
        try {
            if FileExist(tmp)
                FileDelete(tmp)
            f := FileOpen(tmp, "w", "UTF-8")
            f.Write(text)
            f.Close()
            FileMove(tmp, path, true) ; atomic replace
            return
        } catch as ex {
            if (A_LastError = 32) {  ; ERROR_SHARING_VIOLATION
                Sleep(150)
                continue
            }
            throw ex
        }
    }
    throw Error("Could not save file (sharing violation persisted): " path)
}

; INI writes with small retry (keeps IniWrite semantics)
IniWriteRetry(value, path, section, key) {
    SplitPath(path, , &dir)
    if !DirExist(dir)
        DirCreate(dir)

    loop 5 {
        try {
            IniWrite(value, path, section, key)
            return
        } catch as ex {
            if (A_LastError = 32) {
                Sleep(150)
                continue
            }
            throw ex
        }
    }
    throw Error("IniWrite failed (sharing violation persisted): " path " [" section "/" key "]")
}

; ---- profiles (overlay theme) ---------------------------------
ListProfiles() {
    global profilesDir
    out := []
    Loop Files profilesDir "\*.ini" {
        n := A_LoopFileName
        n := RegExReplace(n, "\.ini$", "")
        out.Push(n)
    }
    if out.Length {
        txt := ""
        for n in out
            txt .= n "`n"
        txt := RTrim(txt, "`n")
        txt := Sort(txt)
        out := []
        for n in StrSplit(txt, "`n")
            out.Push(n)
    }
    return out
}

ArrayIndexOf(arr, val) {
    for i, v in arr
        if (v = val)
            return i
    return 0
}

RefreshProfilesList(select := "") {
    global ddlProf
    profList := ListProfiles()
    ddlProf.Delete()
    if (profList.Length) {
        ddlProf.Add(profList)
        selIdx := select != "" ? ArrayIndexOf(profList, select) : 1
        if (selIdx = 0)
            selIdx := 1
        ddlProf.Choose(selIdx)
    } else {
        ddlProf.Text := ""
    }
}

GetOverlayStateForProfile() {
    global overlayTrans, boxBgHex, bdrOutHex, bdrInHex, txtHex
    global fontName, fontSize, fontBold
    global bdrOutW, bdrInW
    SyncUnifiedWindowAppearance()
    return Map(
        "overlayTrans", overlayTrans,
        "boxBg",  boxBgHex,
        "bdrOut", boxBgHex,
        "bdrIn",  boxBgHex,
        "txt",    txtHex,
        "font",   fontName,
        "size",   fontSize,
        "bold",   fontBold,
        "outw",   0,
        "inw",    0
    )
}

ApplyOverlayState(st) {
    global overlayTrans, boxBgHex, bdrOutHex, bdrInHex, txtHex
    global fontName, fontSize, fontBold, bdrOutW, bdrInW
    global slTrans, lblTransPct
    global rectBg, rectTxt
    global ddlFont, edFSize, chkFontBold

    if st.Has("overlayTrans") {
        overlayTrans := Integer(st["overlayTrans"])
        slTrans.Value := overlayTrans
        lblTransPct.Value := Round(overlayTrans / 255 * 100) . "%"
        oldMode := A_TitleMatchMode
        SetTitleMatchMode 3
        WinSetTransparent(overlayTrans, "Translator")
        SetTitleMatchMode oldMode
    }
    if st.Has("boxBg") {
        boxBgHex := st["boxBg"], rectBg.Opt("Background" . boxBgHex)
    }
    if st.Has("txt") {
        txtHex := st["txt"], rectTxt.Opt("Background" . txtHex)
    }
    if st.Has("font") {
        fontName := st["font"]
        try ddlFont.Text := fontName
        catch {
            availableFont := Trim(ddlFont.Text)
            if (availableFont != "")
                fontName := availableFont
        }
    }
    if st.Has("size") {
        fontSize := Integer(st["size"]), edFSize.Value := fontSize
    }
    if st.Has("bold") {
        fontBold := Integer(st["bold"]) ? 1 : 0
        chkFontBold.Value := fontBold
    }
    SyncUnifiedWindowAppearance()
    RefreshColorSwatches()
    SaveAll()
    SendOverlayTheme()
}

SaveProfile(name) {
    global profilesDir
    st := GetOverlayStateForProfile()
    path := profilesDir "\" name ".ini"
    for k, v in st
        IniWriteRetry(v, path, "overlay", k)

    oldMode := A_TitleMatchMode
    SetTitleMatchMode 3
    WinGetPos &x, &y, &w, &h, "Translator"
    if (x != "")
    {
        IniWriteRetry(x,   path, "overlay", "ovX")
        IniWriteRetry(y,   path, "overlay", "ovY")
        IniWriteRetry(w,   path, "overlay", "ovW")
        IniWriteRetry(h,   path, "overlay", "ovH")
        try {
            dpi := GetWindowDPI(WinExist("Translator"))
            IniWriteRetry(dpi, path, "overlay", "ovDPI")
            DbgCP("SaveProfile '" name "' pos=(" x "," y "," w "," h ") dpi=" dpi)
        } catch {
            DbgCP("SaveProfile '" name "' pos=(" x "," y "," w "," h ") dpi=?")
        }
    }
    SetTitleMatchMode oldMode
}


LoadProfile(name) {
    global profilesDir, iniPath, ddlProf
    path := profilesDir "\" name ".ini"
    if !FileExist(path) {
        MsgBox("Profile not found:`n" path, "Missing", 48)
        return
    }
    ; remember last used Translation Window profile in control.ini
    try IniWrite(name, iniPath, "profiles", "translator_last")
    try ddlProf.Text := name
    st := Map()
    for k in ["overlayTrans","boxBg","txt","font","size","bold","ovX","ovY","ovW","ovH","ovDPI"] {
        v := IniRead(path, "overlay", k, "")
        if (v != "")
            st[k] := v
    }
    ApplyOverlayState(st)

    oldMode := A_TitleMatchMode
    SetTitleMatchMode 3
    if (st.Has("ovX") && st.Has("ovY") && st.Has("ovW") && st.Has("ovH")) {
        x := Integer(st["ovX"])
        y := Integer(st["ovY"])
        w := Integer(st["ovW"])
        h := Integer(st["ovH"])
        DbgCP("LoadProfile '" name "' -> WinMove x=" x " y=" y " w=" w " h=" h " (savedDPI=" (st.Has("ovDPI")?st["ovDPI"]:"") ")")
        WinMove x, y, w, h, "Translator"
		try {
            s := "action=save_bounds"
            DbgCP("Sending save_bounds command to overlay.")
            buf := Buffer(StrLen(s)*2 + 2, 0)
            StrPut(s, buf, "UTF-16")
            cds := Buffer(A_PtrSize*3, 0)
            NumPut("UPtr", 0,        cds, 0)
            NumPut("UPtr", buf.Size, cds, A_PtrSize)
            NumPut("Ptr",  buf.Ptr,  cds, 2*A_PtrSize)
            target := WinExist("Translator")
            if (target)
                DllCall("User32\SendMessageW", "Ptr", target, "UInt", 0x004A, "Ptr", 0, "Ptr", cds.Ptr)
        }
    }
    SetTitleMatchMode oldMode
}

DeleteProfile(name) {
    global profilesDir
    path := profilesDir "\" name ".ini"
    try FileDelete(path)
    DbgCP("DeleteProfile '" name "'")
}

CreateProfile() {
    global ddlProf, profilesDir
    ib := InputBox("Enter a name for the new profile:", "Create profile", "w320 h140")
    if (ib.Result = "Cancel")
        return
    name := Trim(ib.Value)
    if (name = "") {
        MsgBox("Please enter a non-empty name.", "Create profile", "OK Icon!")
        ddlProf.Focus()
        return
    }
    name := RegExReplace(name, '[\\/:*?"<>|]+', "_")
    if RegExMatch(name, 'i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$')
        name := "_" name

    path := profilesDir "\" name ".ini"
    if FileExist(path) {
        if (MsgBox("Profile '" name "' already exists.`nOverwrite it?", "Create profile", "YesNo Icon!") != "Yes") {
            ddlProf.Focus()
            return
        }
    }
    SaveProfile(name)
    list := ListProfiles()
    ddlProf.Delete()
    if (list.Length)
        ddlProf.Add(list)
    ddlProf.Text := name
    ddlProf.Focus()
    try Toast("Saved profile: " name), DbgCP("CreateProfile done: " name)
}

; ---------------------------------------------------------------
; Send current theme to overlay via WM_COPYDATA
SendOverlayTheme(targetTitle := "") {
    ; ===== Vars for Translator =====
    global overlayTrans, boxBgHex, bdrOutHex, bdrInHex, txtHex, nameHex
    global fontName, fontSize, fontBold, bdrOutW, bdrInW
    ; ===== Vars for Explainer =====
    global overlayTrans_EW, boxBgHex_EW, bdrOutHex_EW, bdrInHex_EW, txtHex_EW, nameHex_EW
    global fontName_EW, fontSize_EW, fontBold_EW, bdrOutW_EW, bdrInW_EW

    SyncUnifiedWindowAppearance()

    ; Launch-time initialization can target one overlay without reformatting the
    ; other window's existing RichEdit content. Normal settings changes update both.
    targetTitles := (targetTitle = "") ? ["Translator", "Explainer"] : [targetTitle]
    for title in targetTitles {
        oldMode := A_TitleMatchMode
        SetTitleMatchMode 3
        target := WinExist(title)
        SetTitleMatchMode oldMode
        if !target
            continue

        if (title = "Explainer") {
            s := "trans=" overlayTrans_EW
               . "|bg="    boxBgHex_EW
               . "|b_out=" boxBgHex_EW
               . "|b_in="  boxBgHex_EW
               . "|txt="   txtHex_EW
               . "|font="  fontName_EW
               . "|size="  fontSize_EW
               . "|bold="  fontBold_EW
               . "|outw=0"
               . "|inw=0"
        } else {
            s := "trans=" overlayTrans
               . "|bg="    boxBgHex
               . "|b_out=" boxBgHex
               . "|b_in="  boxBgHex
               . "|txt="   txtHex
               . "|name="  nameHex
               . "|font="  fontName
               . "|size="  fontSize
               . "|bold="  fontBold
               . "|outw=0"
               . "|inw=0"
        }

        DbgCP("SendTheme(" title ") " s)

        ; --- WM_COPYDATA send (UTF-16 string) ---
        buf := Buffer(StrLen(s)*2 + 2, 0)
        StrPut(s, buf, "UTF-16")
        cds := Buffer(A_PtrSize*3, 0)
        NumPut("UPtr", 0,        cds, 0)
        NumPut("UPtr", buf.Size, cds, A_PtrSize)
        NumPut("Ptr",  buf.Ptr,  cds, 2*A_PtrSize)
        DllCall("User32\SendMessageW", "Ptr", target, "UInt", 0x004A, "Ptr", 0, "Ptr", cds.Ptr)
    }
}

; ---------------------------------------------------------------
; Send a generic command string to either overlay via WM_COPYDATA.
SendOverlayCmdTo(title, s) {
    target := CPFindExactWindow(title)
    if !target
        return false
    ; --- Send UTF-16 payload ---
    buf := Buffer(StrLen(s)*2 + 2, 0)
    StrPut(s, buf, "UTF-16")
    cds := Buffer(A_PtrSize*3, 0)
    NumPut("UPtr", 0,        cds, 0)
    NumPut("UPtr", buf.Size, cds, A_PtrSize)
    NumPut("Ptr",  buf.Ptr,  cds, 2*A_PtrSize)
    DllCall("User32\SendMessageW", "Ptr", target, "UInt", 0x004A, "Ptr", 0, "Ptr", cds.Ptr)
    return true
}

SendOverlayCmd(s) {
    return SendOverlayCmdTo("Translator", s)
}
