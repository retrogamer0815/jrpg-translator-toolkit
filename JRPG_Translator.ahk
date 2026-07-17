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

SafeCall(fn) {
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
global __DBG_ENABLED_CP := true
global __DBG_LOG := A_ScriptDir "\Settings\debug.log"
DbgCP(msg) {
    Try {
	    ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        dbgDir := A_Temp "\JRPG_Control"
        if !DirExist(dbgDir)
            DirCreate(dbgDir)
        FileAppend("[" ts "] CONTROL  " msg "`n", dbgDir "\debug.log", "UTF-8")
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
global CPThemeMessagesRegistered := false
global CPComboArrowOverlays := []

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

CPRegisterColorSwatch(cpSwatchCtrl) {
    global CPThemeColorSwatchHwnds
    if IsObject(cpSwatchCtrl) && cpSwatchCtrl.Hwnd
        CPThemeColorSwatchHwnds[cpSwatchCtrl.Hwnd] := true
    return cpSwatchCtrl
}

CPIsColorSwatchControl(cpSwatchHwnd) {
    global CPThemeColorSwatchHwnds
    return IsSet(CPThemeColorSwatchHwnds) && IsObject(CPThemeColorSwatchHwnds)
        && CPThemeColorSwatchHwnds.Has(cpSwatchHwnd)
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
    } else if (cpThemeClass = "Edit" || cpThemeClass = "ComboBox") {
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
    if (parentHwnd != ui.Hwnd && !DllCall("user32\IsChild", "ptr", ui.Hwnd, "ptr", lParam, "int"))
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
        try DllCall("user32\RedrawWindow", "ptr", ui.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100)
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

CPRestoreFocusVisual() {
    global CPFocusVisualCtrl, CPFocusVisualHwnd
    if (IsSet(CPFocusVisualCtrl) && CPFocusVisualCtrl) {
        try CPFocusVisualCtrl.SetFont("Norm")
        try CPApplyThemeToControl(CPFocusVisualCtrl.Hwnd)
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
    global tab, tabNames, CPTabButtons, CPTabBarFill, CPTabBarHasNavFocus, CPTabRenderedState, controlDarkMode
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

    for cpTabIndex, cpTabCtrl in CPTabButtons {
        if (cpTabIndex = cpTabActive) {
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

CPLayoutCustomTabBar(*) {
    global tab, CPTabButtons, CPTabNaturalWidths, CPTabBarFill
    if !(IsSet(tab) && tab && tab.Hwnd && IsSet(CPTabButtons) && IsObject(CPTabButtons))
        return

    tab.GetPos(&cpTabX, &cpTabY, &cpTabW)
    cpTabBarH := 30
    cpTabAvailableW := Max(1, cpTabW)
    cpTabDesiredTotal := 0
    for cpTabNaturalWidth in CPTabNaturalWidths
        cpTabDesiredTotal += cpTabNaturalWidth

    CPTabBarFill.Move(cpTabX, cpTabY, cpTabAvailableW, cpTabBarH)
    CPTabBarFill.Visible := true

    cpTabDrawX := cpTabX
    cpTabRemainingW := cpTabAvailableW
    for cpTabIndex, cpTabCtrl in CPTabButtons {
        if (cpTabDesiredTotal > cpTabAvailableW) {
            if (cpTabIndex = CPTabButtons.Length)
                cpTabItemW := cpTabRemainingW
            else
                cpTabItemW := Max(38, Floor(CPTabNaturalWidths[cpTabIndex] * cpTabAvailableW / cpTabDesiredTotal))
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
    global ui, CPTabButtons
    if !(IsSet(ui) && ui && ui.Hwnd && IsSet(CPTabButtons) && IsObject(CPTabButtons))
        return

    ; Use the actual child window beneath the pointer. Coordinate comparisons are
    ; unsafe here because MouseGetPos and GetWindowRect can use different origins.
    MouseGetPos &cpMouseX, &cpMouseY, &cpMouseWindowHwnd, &cpMouseControlHwnd, 2
    if (cpMouseWindowHwnd != ui.Hwnd || !cpMouseControlHwnd)
        return

    for cpTabIndex, cpTabCtrl in CPTabButtons {
        if (cpMouseControlHwnd = cpTabCtrl.Hwnd
            && DllCall("user32\IsWindowVisible", "ptr", cpTabCtrl.Hwnd, "int")) {
            CPSelectCustomTab(cpTabIndex)
            return
        }
    }
}

CPCreateCustomTabBar() {
    global ui, tabNames, CPTabButtons, CPTabNaturalWidths, CPTabBarFill
    global CPTabBarHasNavFocus, CPTabRenderedState

    CPTabButtons := []
    CPTabNaturalWidths := []
    CPTabBarHasNavFocus := false
    CPTabRenderedState := ""

    ; This opaque strip covers the native tab header. The real tab control remains
    ; underneath only to manage page visibility and provide a keyboard focus target.
    try tab.Opt("+0x04000000") ; WS_CLIPSIBLINGS keeps native painting below the custom row.
    CPTabBarFill := ui.Add("Text", "x0 y0 w1 h1 Hidden BackgroundF0F0F0 +0x100 +0x04000000")
    for cpTabIndex, cpTabLabel in tabNames {
        cpTabCtrl := ui.Add("Text", "x0 y0 h30 Hidden Center Border +0x100 +0x200 +0x04000000 BackgroundF3F3F3 c202020", cpTabLabel)
        cpTabCtrl.GetPos(,, &cpTabNaturalW)
        CPTabNaturalWidths.Push(Max(46, cpTabNaturalW + 18))
        cpTabCtrl.Cursor := "Hand"
        cpTabCtrl.OnEvent("Click", CPSelectCustomTab.Bind(cpTabIndex))
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
    if (IsSet(CPFocusVisualHwnd) && CPFocusVisualHwnd = hwnd)
        return

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
    cpFocusColors := CPPalette()
    try ctrl.Opt("+Background" cpFocusColors["focus"])
    try ctrl.SetFont("Bold c" (controlDarkMode ? cpFocusColors["accentText"] : cpFocusColors["accentFocus"]))
}

UpdateCPActiveTabHighlight(*) {
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
    if (className = "" || className = "Static")
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

CPFocusTabBar() {
    global tab, CPFocusVisualNavHwnd
    if !(IsSet(tab) && tab && tab.Hwnd)
        return false

    CPFocusVisualNavHwnd := tab.Hwnd
    try tab.Focus()
    CPSetTabFocusIndicator(true)
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
            return
        }
    }
    try ControlFocus("ahk_id " hwnd, "ahk_id " ui.Hwnd)
    try DllCall("user32\SetFocus", "ptr", hwnd, "ptr")
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
    global tab, CPFocusVisualNavHwnd
    if !(IsSet(tab) && tab && tab.Hwnd)
        return

    static TCM_GETITEMCOUNT := 0x1304
    static TCM_GETCURSEL := 0x130B
    count := SendMessage(TCM_GETITEMCOUNT, 0, 0, tab.Hwnd)
    if (count <= 0)
        return

    tabIdx := SendMessage(TCM_GETCURSEL, 0, 0, tab.Hwnd) ; zero-based
    if (tabIdx < 0)
        tabIdx := 0

    next := Mod(tabIdx + dir + count, count) + 1 ; Gui control value is one-based
    try tab.Value := next
    try tab.Focus()
    CPFocusVisualNavHwnd := tab.Hwnd
    CPSetTabFocusIndicator(true)
    CPRenderCustomTabBar(true)
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
defOverlay      := A_IsCompiled ? ".\bin\overlay.exe" : ".\bin\jrpg_overlay_C.ahk"
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
; --- Debug toggle default ---
defDebugMode := 1   ; keep current behavior; set to 0 if you want logging OFF by default
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
        ".\bin\jrpg_overlay_c.exe"
      , StrLower(A_ScriptDir "\bin\jrpg_overlay_C.exe")
      , ".\bin\jrpg_overlay.exe"
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

debugMode := Integer(Load("debugMode", defDebugMode, "cfg"))
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
ModelListRead(key, defaultsArr) {
    raw := ""
    try raw := IniRead(iniPath, "models", key, "")
    if (Trim(raw) = "") {
        return defaultsArr.Clone()
    }
    out := []
    for it in StrSplit(raw, ",") {
        s := Trim(it)
        if (s != "")
            out.Push(s)
    }
    return out.Length ? out : defaultsArr.Clone()
}
ModelListWrite(key, arr) {
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
    return out
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
def_openai_img   := ["gpt-5.5","gpt-5.4-nano","gpt-5.4-pro","gpt-4o","gpt-4o-mini"]
def_gemini_img   := ["gemini-3.1-flash-lite","gemini-3.5-flash","gemini-3.1-pro-preview","gemini-2.5-flash","gemini-2.5-flash-lite","gemini-2.5-pro"]
def_openai_tr    := ["gpt-5.5","gpt-5.4-nano","gpt-5.4-pro","gpt-4o-mini","gpt-4o"]
def_openai_audio := ["gpt-realtime-translate"]
def_gemini_audio := ["gemini-3.5-live-translate-preview"]

; --- Explanation tab defaults (provider + text models)
defExplainProvider    := "openai"
defExplainOpenAIModel := "gpt-4o-mini"            ; uses text/chat models
defExplainGeminiModel := "gemini-2.5-flash"       ; Gemini text

; load lists from INI (or defaults)
model_openai_img   := ModelListRead("openai_img",   def_openai_img)
model_gemini_img   := ModelListRead("gemini_img",   def_gemini_img)
model_openai_tr    := ModelListRead("openai_tr",    def_openai_tr)
model_openai_audio := ModelListRead("openai_audio", def_openai_audio)
model_gemini_audio := ModelListRead("gemini_audio", def_gemini_audio)
model_openai_img := ModelListMergeUnique(model_openai_img, model_openai_tr)
model_openai_tr := model_openai_img
ModelListEnsure(model_openai_audio, "gpt-realtime-translate")
ModelListEnsure(model_gemini_audio, "gemini-3.5-live-translate-preview")

SaveAll(){
    global pythonExe,audioScript,overlayAhk,imgScript,overlayTrans,captureDir
    global trModel,audioProvider,geminiAudioModel,audioTargetLang
    global imgProvider,imgModel,geminiImgModel
    global iniPath
    global capMaxKB,capMode,capRect
    global boxBgHex,bdrOutHex,bdrInHex,txtHex
    global fontName,fontSize
    global bdrOutW,bdrInW
    global model_openai_img, model_gemini_img, model_openai_tr, model_openai_audio, model_gemini_audio
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
    ModelListWrite("openai_tr",    model_openai_tr)
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

    b1.OnEvent("Click", (*) => (g.Hide(), StartTempHideWatcher("region"), _SendCapPick("region")))
    b2.OnEvent("Click", (*) => (g.Hide(), StartTempHideWatcher("window"), _SendCapPick("window")))
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

_SendCapPick(kind) {
    global capMaxKB
    ; compose a simple, future-proof key=value payload
    payload := "capcmd=pick"
            .  "|kind=" kind
            .  "|maxkb=" capMaxKB
    ok := SendOverlayCmd(payload)
    if !ok {
        Toast("Translator not found")
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
    global model_openai_img, model_gemini_img, model_openai_tr, model_openai_audio, model_gemini_audio
    global ddlIMG, ddlIMG_GM, ddlEOpenAI, ddlEGem, ddlTR, ddlA_GM

    arr := 0
    combos := []
    switch key {
        case "openai_img":
            arr := model_openai_img
            if IsSet(ddlIMG)
                combos.Push(ddlIMG)
            if IsSet(ddlEOpenAI)
                combos.Push(ddlEOpenAI)
        case "gemini_img":
            arr := model_gemini_img
            if IsSet(ddlIMG_GM)
                combos.Push(ddlIMG_GM)
            if IsSet(ddlEGem)
                combos.Push(ddlEGem)
        case "openai_tr":
            arr := model_openai_tr
            if IsSet(ddlEOpenAI)
                combos.Push(ddlEOpenAI)
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

    for combo in combos {
        modelSelectionText := Trim(combo.Text)
        if (IsObject(activeCombo) && combo.Hwnd = activeCombo.Hwnd && activePreferredText != "")
            modelSelectionText := activePreferredText
        SetComboItems(combo, arr)
        SetComboToExistingItem(combo, arr, modelSelectionText)
    }
}

AddModel(arr, key, combo) {
    new := Trim(InputBox("Add model:", "Add").Value)
    if (new = "")
        return
    for v in arr
        if (StrLower(v) = StrLower(new)) {
            MsgBox("Already in the list: " new, "Add model")
            return
        }
    arr.Push(new)            ; modifies the original array
    ModelListWrite(key, arr)
    RefreshModelCombos(key, combo, new)
    DbgCP("Model added under [" key "]: " new)
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
; GUI
; =========================
ui := Gui("+Resize +MinSize890x700", "JRPG Translator")
CPRegisterThemeMessages()

; --- Control Panel default bounds (used only if no valid [gui_bounds] exist) ---
defGuiX := 140
defGuiY := 140
defGuiW := 890
defGuiH := 680

IsValidBounds(x, y, w, h) {
    if !((x is number) && (y is number) && (w is number) && (h is number))
        return false
    if (w < 640 || h < 420)
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
btnIMG_GM_Add := ui.Add("Button", "x+6 w60", "Add")
btnIMG_GM_Del := ui.Add("Button", "x+6 w60", "Delete")

ui.Add("Text", "xm y+12 w90", "OpenAI model:")
ddlIMG := ui.Add("DropDownList", "x+m w260 0x210", model_openai_img)
imgInitIdx := ArrIndexOf(model_openai_img, imgModel)
ddlIMG.Choose(imgInitIdx ? imgInitIdx : 1)
ddlIMG.OnEvent("Change", (*) => (UpdateVars(), SaveAll(), ApplyShotSettings()))
btnIMG_Add := ui.Add("Button", "x+6 w60", "Add")
btnIMG_Del := ui.Add("Button", "x+6 w60", "Delete")

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
ui.Add(
    "Text"
  , "x+m yp+6"  ; place to the right of the button, slightly lowered
  , "Note: Requires the Translator window to be open."
)

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
btnA_GM_Add := ui.Add("Button", "x+6 w60", "Add")
btnA_GM_Del := ui.Add("Button", "x+6 w60", "Delete")

ui.Add("Text", "xm y+12", "OpenAI live model:")
ddlTR := ui.Add("DropDownList", "x+m w420 0x210", model_openai_audio) ; initial width; ResizeUI will adjust
SetComboToExistingItem(ddlTR, model_openai_audio, trModel)
btnTR_Add := ui.Add("Button", "x+6 w60", "Add")
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
ui.Add("Text", "xm y+6", "Overlay Transparency")
slTrans := ui.Add("Slider", "x+m w200 Range0-255 ToolTip")
lblTransPct := ui.Add("Text", "x+m", "100%")

ui.Add("Text", "xm y+18 w200", "Window color:")
rectBg := CPRegisterColorSwatch(ui.Add("Text", "x+m w84 h34 Border"))
ui.Add("Text", "xm y+18 w200", "Text color:")
rectTxt := CPRegisterColorSwatch(ui.Add("Text", "x+m w84 h34 Border"))
ui.Add("Text", "xm y+18 w200", "Speaker name:")
rectName := CPRegisterColorSwatch(ui.Add("Text", "x+m w84 h34 Border"))

RefreshColorSwatches()

ui.Add("Text", "xm y+22", "Font:")
ddlFont := ui.Add("ComboBox", "x+m w260", [])
ui.Add("Text", "x+14 yp", "Size:")
edFSize := ui.Add("Edit", "x+m w60 Number", fontSize)
udFSize := ui.Add("UpDown", "Range6-128", fontSize)

ui.Add("Text", "xm y+22", "Profile:")
ddlProf  := ui.Add("ComboBox", "x+m w240", [])
btnPNew  := ui.Add("Button", "x+8 w80",  "Add")
btnPSave := ui.Add("Button", "x+6 w70",  "Save")
btnPLoad := ui.Add("Button", "x+6 w70",  "Load")
btnPDel  := ui.Add("Button", "x+6 w70",  "Delete")

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

; Prompt profile events + initial list
btnPrEdit.OnEvent("Click", OpenPromptEditor)
btnPrNew.OnEvent("Click",  NewPromptProfile)
btnPrDel.OnEvent("Click",  DeletePromptProfile)
RefreshPromptProfilesList(promptProfile)


; --- Explanation: Provider + Models (independent from Screenshot Translation)
tab.UseTab(4)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer
ui.Add("Text", "xm y+6 w90", "AI Provider:")
ddlEProv := ui.Add("DropDownList", "x+m w260 0x210", ["Gemini","OpenAI"])

eProvIdx := (StrLower(explainProvider) = "gemini") ? 1 : 2
ddlEProv.Choose(eProvIdx)
ddlEProv.OnEvent("Change", (*) => (UpdateVars(), SaveAll()))

; --- Gemini row (unchanged)
ui.Add("Text", "xm y+12 w90", "Gemini model:")
ddlEGem := ui.Add("DropDownList", "x+m w260 0x210", model_gemini_img)
eGemIdx := ArrIndexOf(model_gemini_img, explainGeminiModel)
ddlEGem.Choose(eGemIdx ? eGemIdx : 1)
ddlEGem.OnEvent("Change", (*) => (UpdateVars(), SaveAll()))
btnEGem_Add := ui.Add("Button", "x+6 w60", "Add")
btnEGem_Del := ui.Add("Button", "x+6 w60", "Delete")

; --- OpenAI row (moved here; use a fresh y step so it sits below Gemini)
ui.Add("Text", "xm y+12", "OpenAI model:")
ddlEOpenAI := ui.Add("DropDownList", "x+m w260 0x210", model_openai_tr)
eOpenAIIdx := ArrIndexOf(model_openai_tr, explainOpenAIModel)
ddlEOpenAI.Choose(eOpenAIIdx ? eOpenAIIdx : 1)
ddlEOpenAI.OnEvent("Change", (*) => (UpdateVars(), SaveAll()))
btnEOpenAI_Add := ui.Add("Button", "x+6 w60", "Add")
btnEOpenAI_Del := ui.Add("Button", "x+6 w60", "Delete")


; Initialize enabled/disabled state for Explanation models
ToggleExplanationControls()
; Force a post-build sync from INI so later generic repainting canâ€™t overwrite these
SyncExplanationFromIni()

; EXPLANATION prompt profile (independent from Screenshot/Audio prompts)
ui.Add("Text", "xm y+10 w90", "Prompt:")
ddlEPr     := ui.Add("DropDownList", "x+m w260 0x210", [])
btnEPrEdit := ui.Add("Button", "x+6 w70", "Edit")
btnEPrNew  := ui.Add("Button", "x+6 w70", "Add")
btnEPrDel  := ui.Add("Button", "x+6 w70", "Delete")

; anchor to the current Sectionâ€™s left edge, keep same row spacing
btnExplainNow := ui.Add("Button", "xs y+12 w220", "Explain last jp. Text")

; Move: Save explanations checkbox â€” placed under the button row, left-aligned to the Section
saveExplChk := ui.Add("CheckBox", "xs y+10", "Save explanations to textfiles")
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

btnEOpenAI_Add.OnEvent("Click", (*) => AddModel(model_openai_img, "openai_img", ddlEOpenAI))
btnEOpenAI_Del.OnEvent("Click", (*) => DeleteModel(model_openai_img, "openai_img", ddlEOpenAI))

btnEGem_Add.OnEvent("Click", (*) => AddModel(model_gemini_img, "gemini_img", ddlEGem))
btnEGem_Del.OnEvent("Click", (*) => DeleteModel(model_gemini_img, "gemini_img", ddlEGem))


; --- Tab 5: EXPLANATION WINDOW  (UI only, not wired yet)
tab.UseTab(5)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer
; Layout parity with "Translation Window" (Tab 3), distinct control names (EW_*)
ui.Add("Text", "xm y+6", "Overlay Transparency")
slTrans_EW := ui.Add("Slider", "x+m w200 Range0-255 ToolTip")
lblTransPct_EW := ui.Add("Text", "x+m", "100%")

ui.Add("Text", "xm y+18 w200", "Window color:")
rectBg_EW := CPRegisterColorSwatch(ui.Add("Text", "x+m w84 h34 Border"))
ui.Add("Text", "xm y+18 w200", "Text color:")
rectTxt_EW := CPRegisterColorSwatch(ui.Add("Text", "x+m w84 h34 Border"))

RefreshColorSwatches_EW()

ui.Add("Text", "xm y+22", "Font:")
ddlFont_EW := ui.Add("ComboBox", "x+m w260", [])
ui.Add("Text", "x+12 yp", "Size")
edFSize_EW := ui.Add("Edit", "x+m w60 Number")
udFSize_EW := ui.Add("UpDown", "Range8-96")

; --- Explanation Profiles row ---
ui.Add("Text", "xm y+18", "Profile:")
ddlProf_EW := ui.Add("ComboBox", "x+m w220")
btnProfCreate_EW := ui.Add("Button", "x+m", "Add")
btnProfSave_EW := ui.Add("Button", "x+m", "Save")
btnProfLoad_EW := ui.Add("Button", "x+m", "Load")
btnProfDel_EW  := ui.Add("Button", "x+m", "Delete")

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
TooltipBind(cbDebug, "If enabled, Control Panel and Overlay write verbose logs") ; optional
cbDebug.OnEvent("Click", (*) => MarkDirty())

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
btnA_GM_Add   .OnEvent("Click", (*) => AddModel(model_gemini_audio, "gemini_audio", ddlA_GM))
btnA_GM_Del   .OnEvent("Click", (*) => DeleteModel(model_gemini_audio, "gemini_audio", ddlA_GM))

btnTR_Add     .OnEvent("Click", (*) => AddModel(model_openai_audio, "openai_audio", ddlTR))
btnTR_Del     .OnEvent("Click", (*) => DeleteModel(model_openai_audio, "openai_audio", ddlTR))

btnIMG_Add    .OnEvent("Click", (*) => AddModel(model_openai_img,   "openai_img",   ddlIMG))
btnIMG_Del    .OnEvent("Click", (*) => DeleteModel(model_openai_img,"openai_img",   ddlIMG))

btnIMG_GM_Add .OnEvent("Click", (*) => AddModel(model_gemini_img,   "gemini_img",   ddlIMG_GM))
btnIMG_GM_Del .OnEvent("Click", (*) => DeleteModel(model_gemini_img,"gemini_img",   ddlIMG_GM))

; initial paint + status, then start timer
Repaint()
LoadFontsIntoCombo()
LoadFontsIntoCombo_EW()   ; â† add this
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
    global ddlFont, edFSize, fontName, fontSize
    global ddlPrompt, promptProfile
    global ddlPost, imgPostproc, postCodes
	global ddlEProv, ddlEOpenAI, ddlEGem
    global explainProvider, explainOpenAIModel, explainGeminiModel, iniPath
    global model_openai_img, model_gemini_img, model_openai_tr, model_openai_audio, model_gemini_audio

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

    ddlFont.Text := fontName
    edFSize.Value := fontSize

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
        SetComboToExistingItem(ddlEGem, model_gemini_img, explainGeminiModel)
    if IsSet(ddlEOpenAI)
        SetComboToExistingItem(ddlEOpenAI, model_openai_tr, explainOpenAIModel)
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

                ; Profile row: use last Explainer profile from control.ini
        try ddlProf_EW.Text := IniRead(iniPath, "profiles", "explainer_last", "")
    }

    ToggleModelControls()
}

ToggleModelControls(){
    global ddlProv, ddlIMG, ddlIMG_GM
    prov := ddlProv.Text
    ddlIMG.Enabled    := (prov = "openai")
    ddlIMG_GM.Enabled := (prov = "gemini")
}
ToggleAudioControls(){
    global ddlAProv, ddlTR, ddlA_GM
    ap := StrLower(ddlAProv.Text)
    isOpenAI := (ap = "openai")
    ddlTR.Enabled  := isOpenAI
    ddlA_GM.Enabled := !isOpenAI
}
; NEW: Explanation tab toggles
ToggleExplanationControls(){
    global ddlEProv, ddlEOpenAI, ddlEGem
    ep := StrLower(Trim(ddlEProv.Text))
    ddlEOpenAI.Enabled := (ep = "openai")
    ddlEGem.Enabled    := (ep = "gemini")
}

; NEW: force-sync Explanation dropdowns from INI (defensive against any later repaint)
SyncExplanationFromIni(){
    global iniPath
    global ddlEProv, ddlEOpenAI, ddlEGem
    global model_openai_tr, model_gemini_img

    prov := StrLower(Trim(IniRead(iniPath, "cfg_explainer", "explainProvider", "")))
    gm   := Trim(IniRead(iniPath, "cfg_explainer", "explainGeminiModel", ""))
    om   := Trim(IniRead(iniPath, "cfg_explainer", "explainOpenAIModel", ""))

    if (prov != "")
        ddlEProv.Choose(prov = "gemini" ? 1 : 2)
    if (gm != "")
        SetComboToExistingItem(ddlEGem, model_gemini_img, gm)
    if (om != "")
        SetComboToExistingItem(ddlEOpenAI, model_openai_tr, om)
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
	global ddlFont, edFSize, fontName, fontSize
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
    SyncUnifiedWindowAppearance()
    promptProfile    := ddlPrompt.Text
    imgPostproc      := postCodes[ddlPost.Value]
	explainProvider    := ddlEProv.Text
    explainOpenAIModel := ddlEOpenAI.Text
    explainGeminiModel := ddlEGem.Text
	    ; Debug toggle
    debugMode := cbDebug.Value  ; 1/0

    ; Debug toggle
    debugMode := cbDebug.Value  ; 1/0

    ; update env so any child processes launched after this see the new state
    EnvSet "JRPG_DEBUG", (debugMode ? "1" : "0")
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

    global pad, gap
    global tab, sepAction
    global tPython,ePython,bPy,tAud,eAudio,bAud,tOv,eOverlay,bOvSel,tImg,eImg,bImgSel,tExplain,eExplain,bExplainSel
    global ddlAProv,ddlA_GM,ddlTR,ddlAudioTarget,ddlProv,ddlIMG,ddlIMG_GM
    global btnStart,btnStop,btnAudio,btnOv,btnOvClose,btnExplainerLaunch,btnExplainerClose,btnExplainNow,bSave,bClose,chkTop,chkDarkMode
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
    CPLayoutCustomTabBar()

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

        ; Rows with Add/Delete buttons
    MoveRowWithButtons(ddlTR, btnTR_Add, btnTR_Del, rightEdge, btnH)  ; place RIGHT column first
    ddlTR.GetPos(&rx,,,)                                              ; x of right comb
    MoveRowWithButtons(ddlTR,     btnTR_Add,     btnTR_Del,     rightEdge, btnH, 560)     ; capped
    MoveRowWithButtons(ddlA_GM,   btnA_GM_Add,   btnA_GM_Del,   rightEdge, btnH)          ; NEW: Gemini (Audio) now resizes
    MoveRowWithButtons(ddlIMG,     btnIMG_Add,     btnIMG_Del,     rightEdge, btnH)
    MoveRowWithButtons(ddlIMG_GM,  btnIMG_GM_Add,  btnIMG_GM_Del,  rightEdge, btnH)

    ; NEW: prompt row (combo + 3 buttons)
    btnW := 70, g := 6
    ddlPrompt.GetPos(&pcx,&pcy,,)
    pW := Max(160, rightEdge - pcx - (btnW*3 + g*3))
    ddlPrompt.Move(, , pW)
    btnPrEdit.Move(pcx + pW + g, pcy, btnW, btnH)
    btnPrNew .Move(pcx + pW + g + btnW + g, pcy, btnW, btnH)
    btnPrDel .Move(pcx + pW + g + (btnW+g)*2, pcy, btnW, btnH)

    ; AUDIO target language row
    ddlAudioTarget.GetPos(&apx,&apy,,)
    apW := Max(160, Min(420, rightEdge - apx))
    ddlAudioTarget.Move(, , apW)

    ; post-processing row (single combo)
    ddlPost.GetPos(&ppx,&ppy,,)
    ppW := Max(160, rightEdge - ppx)
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

    ; ensure ySave is defined from current client height (bottom action row)
    ui.GetClientPos(,, &cliW, &cliH)
    ySave := cliH - btnH - pad

    bSave.Move(pad, ySave, 120, btnH)
    bClose.Move(pad + 120 + 8, ySave, 120, btnH)
    bClose.GetPos(&cx,&cy,&btnW,)
    chkTop.Move(cx + btnW + 12, ySave + 6)
    chkTop.GetPos(&cpTopX, &cpTopY, &cpTopW)
    chkDarkMode.Move(cpTopX + cpTopW + 18, ySave + 6)
    CPUpdateComboArrowOverlays()

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
PickColorDialog(initHex := "FFFFFF") {
    global ui
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
    NumPut("UInt", flags, cc, (A_PtrSize=8 ? 40 : 20))
    ret := DllCall("Comdlg32\ChooseColorW", "Ptr", cc.Ptr, "Int")
    if (ret = 0)
        return ""
    gotBGR := NumGet(cc, 3*A_PtrSize, "UInt")
    gotRGB := ((gotBGR & 0xFF) << 16) | (gotBGR & 0xFF00) | ((gotBGR >> 16) & 0xFF)
    return Format("{:06X}", gotRGB)
}

PickAndApply(which) {
    global boxBgHex,txtHex,nameHex
    global rectBg,rectTxt,rectName

    colorCur := (which="bg")    ? boxBgHex
             : (which="name")  ? nameHex
             :                   txtHex

    got := PickColorDialog(colorCur)
    If (got = "")
        Return

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
    global rectBg_EW,rectTxt_EW

    colorCur := (which="bg") ? boxBgHex_EW : txtHex_EW
    got := PickColorDialog(colorCur)
    if (got = "")
        return

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
    global fontName_EW,fontSize_EW,bdrOutW_EW,bdrInW_EW,overlayTrans_EW
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
    global fontName_EW,fontSize_EW,bdrOutW_EW,bdrInW_EW,overlayTrans_EW
    global ewX,ewY,ewW,ewH
    global ui, slTrans_EW, lblTransPct_EW
    global rectBg_EW,rectTxt_EW
    global ddlFont_EW, edFSize_EW, udFSize_EW
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

        Num16(off) => NumGet(buf, off, "UShort")
        Num32(off) => NumGet(buf, off, "UInt")

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
    if (fontName != "")
        ddlFont.Text := fontName
}

LoadFontsIntoCombo_EW(){
    global ddlFont_EW, fontName_EW
    if !IsSet(ddlFont_EW) || !ddlFont_EW
        return
    EnsurePrivateFontsLoaded()
    SendMessage(0x14B, 0, 0, ddlFont_EW.Hwnd)  ; CB_RESETCONTENT
    fonts := GetInstalledFonts()
    ddlFont_EW.Add(fonts)
    if (fontName_EW != "")
        ddlFont_EW.Text := fontName_EW
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
    g.Show()
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
    g.Show()
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
    g.Show()
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
    g.Show()
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
    global fontName, fontSize
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
        "outw",   0,
        "inw",    0
    )
}

ApplyOverlayState(st) {
    global overlayTrans, boxBgHex, bdrOutHex, bdrInHex, txtHex
    global fontName, fontSize, bdrOutW, bdrInW
    global slTrans, lblTransPct
    global rectBg, rectTxt
    global ddlFont, edFSize

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
        fontName := st["font"], ddlFont.Text := fontName
    }
    if st.Has("size") {
        fontSize := Integer(st["size"]), edFSize.Value := fontSize
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
    for k in ["overlayTrans","boxBg","txt","font","size","ovX","ovY","ovW","ovH","ovDPI"] {
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
    global fontName, fontSize, bdrOutW, bdrInW
    ; ===== Vars for Explainer =====
    global overlayTrans_EW, boxBgHex_EW, bdrOutHex_EW, bdrInHex_EW, txtHex_EW, nameHex_EW
    global fontName_EW, fontSize_EW, bdrOutW_EW, bdrInW_EW

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
; Send a generic command string to the Translator overlay via WM_COPYDATA
SendOverlayCmd(s) {
    target := WinExist("ahk_exe AutoHotkey64.exe ahk_class AutoHotkey ahk_pid " ProcessExist() " ahk_title Translator")
    if !target {
        ; fallback: try any window titled exactly "Translator"
        oldMode := A_TitleMatchMode
        SetTitleMatchMode(3)
        target := WinExist("Translator")
        SetTitleMatchMode(oldMode)
        if !target
            return false
    }
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
