#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn
#NoTrayIcon
; =; === Taskbar grouping: shared AppUserModelID ===
DllCall("shell32\SetCurrentProcessExplicitAppUserModelID", "wstr", "JRPGTranslator", "int")
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

CPCreateComboArrowOverlays() {
    global ui, CPComboArrowOverlays
    CPComboArrowOverlays := []
    for cpComboHwnd in WinGetControlsHwnd("ahk_id " ui.Hwnd) {
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
        for cpApplyHwnd in WinGetControlsHwnd("ahk_id " ui.Hwnd)
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
…45138 tokens truncated… : 0),
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

    title := (kind = "jp") ? "Edit JP->EN Glossary - " prof : "Edit EN->EN Glossary - " prof
    path  := (kind = "jp") ? GlossaryJP2ENPath(prof) : GlossaryEN2ENPath(prof)

    ; ensure the profile folder exists (only when editing/saving), but do NOT auto-create other profiles
    if !DirExist(GlossaryProfileDir(prof))
        DirCreate(GlossaryProfileDir(prof))

    txt := ""
    try txt := FileExist(path) ? FileRead(path, "UTF-8") : "# One mapping per line: JP -> EN (or EN -> EN)`r`n"

   g := Gui("+Resize", title)
    edGloss := g.Add("Edit", "xm ym w700 h420 WantTab WantReturn Wrap", txt)
    btnSave  := g.Add("Button", "xm y+8 w100", "Save")
    btnClose := g.Add("Button", "x+8 yp w100", "Close")

    btnSave.OnEvent("Click", (*) => (
    SaveTextAtomic(path, edGloss.Value),
    Toast("Saved " ((kind="jp")?"JP->EN":"EN->EN") " glossary for profile '" prof "'")
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
    FileAppend("# One mapping per line: JP -> EN`r`n", GlossaryJP2ENPath(name), "UTF-8")
    FileAppend("# One mapping per line: EN -> EN`r`n", GlossaryEN2ENPath(name), "UTF-8")

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
    return Map(
        "overlayTrans", overlayTrans,
        "boxBg",  boxBgHex,
        "bdrOut", bdrOutHex,
        "bdrIn",  bdrInHex,
        "txt",    txtHex,
        "font",   fontName,
        "size",   fontSize,
        "outw",   bdrOutW,
        "inw",    bdrInW
    )
}

ApplyOverlayState(st) {
    global overlayTrans, boxBgHex, bdrOutHex, bdrInHex, txtHex
    global fontName, fontSize, bdrOutW, bdrInW
    global slTrans, lblTransPct
    global rectBg, rectOut, rectIn, rectTxt
    global ddlFont, edFSize
    global edOutW, edInW, udOutW, udInW

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
    if st.Has("bdrOut") {
        bdrOutHex := st["bdrOut"], rectOut.Opt("Background" . bdrOutHex)
    }
    if st.Has("bdrIn") {
        bdrInHex := st["bdrIn"], rectIn.Opt("Background" . bdrInHex)
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
    if st.Has("outw") {
        bdrOutW := Integer(st["outw"])
        edOutW.Value := bdrOutW
        try udOutW.Value := bdrOutW
    }

    if st.Has("inw") {
        bdrInW := Integer(st["inw"])
        edInW.Value := bdrInW
        try udInW.Value := bdrInW
    }

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
    for k in ["overlayTrans","boxBg","bdrOut","bdrIn","txt","font","size","outw","inw","ovX","ovY","ovW","ovH","ovDPI"] {
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
               . "|b_out=" bdrOutHex_EW
               . "|b_in="  bdrInHex_EW
               . "|txt="   txtHex_EW
               . "|font="  fontName_EW
               . "|size="  fontSize_EW
               . "|outw="  bdrOutW_EW
               . "|inw="   bdrInW_EW
        } else {
            s := "trans=" overlayTrans
               . "|bg="    boxBgHex
               . "|b_out=" bdrOutHex
               . "|b_in="  bdrInHex
               . "|txt="   txtHex
               . "|name="  nameHex
               . "|font="  fontName
               . "|size="  fontSize
               . "|outw="  bdrOutW
               . "|inw="   bdrInW
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

