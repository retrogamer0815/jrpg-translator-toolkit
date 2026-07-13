#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn
#NoTrayIcon
; === Taskbar grouping: shared AppUserModelID ===
DllCall("shell32\SetCurrentProcessExplicitAppUserModelID", "wstr", "JRPGTranslator", "int")
FileEncoding("UTF-8")
global SCRIPT_PID := DllCall("kernel32\GetCurrentProcessId")

; ----- App root detection (so overlay can live in .\bin) -----
global APP_ROOT := ""
; 1) CLI: --root "<path>"
for i, a in A_Args {
    if (a = "--root" && i < A_Args.Length) {
        APP_ROOT := A_Args[i+1]
        break
    }
}
; 2) Environment variable
if (!APP_ROOT)
    APP_ROOT := EnvGet("APP_ROOT")
; 3) If running from .\bin manually, use parent folder; else A_ScriptDir
if (!APP_ROOT) {
    if RegExMatch(A_ScriptDir, 'i)\\bin$') {
        APP_ROOT := SubStr(A_ScriptDir, 1, InStr(A_ScriptDir, "\",, -1)-1)
    } else {
        APP_ROOT := A_ScriptDir
    }
}
APP_ROOT := ExpandEnv(APP_ROOT)

; ===================== Helpers =====================
IsWindowTopmost(hwnd) {
    ; WS_EX_TOPMOST = 0x00000008
    if !hwnd
        return false
    oldDHW := A_DetectHiddenWindows
    DetectHiddenWindows true
    ok := WinExist("ahk_id " hwnd)
    DetectHiddenWindows oldDHW
    if !ok
        return false
    return (WinGetExStyle("ahk_id " hwnd) & 0x00000008) != 0
}

HexToBGR(hex) {
    v  := "0x" hex
    rr := (v >> 16) & 0xFF
    gg := (v >> 8)  & 0xFF
    bb :=  v        & 0xFF
    return (bb << 16) | (gg << 8) | rr
}
MakeBrush(hex) => DllCall("gdi32\CreateSolidBrush", "int", HexToBGR(hex), "ptr")

ARGB_FromHex(hex, alpha := 255) {
    local v := "0x" hex
    local r := (v >> 16) & 0xFF
    local g := (v >> 8)  & 0xFF
    local bl := v        & 0xFF
    return (alpha << 24) | (bl << 16) | (g << 8) | r
}

; ---[ RichEdit helpers ]-------------------------------------------------------
; === Color helpers ===

; === Color helpers (single source of truth) ===
NormalizeRgb(rgb) {
    ; Accept:
    ;  - numeric types: 0xRRGGBB or decimal
    ;  - "#RRGGBB", "0xRRGGBB", "RRGGBB"
    ;  - pure-digit strings: if length=6 -> treat as hex (e.g. "008040"), else decimal
    if (rgb is Number)
        return (rgb + 0) & 0xFFFFFF

    s := Trim(rgb)
    if (SubStr(s,1,1) = "#")
        s := SubStr(s,2)
    if (SubStr(s,1,2) = "0x" || SubStr(s,1,2) = "0X")
        s := SubStr(s,3)

    ; Pure digits?
    if RegExMatch(s, "^\d+$") {
        ; Exactly 6 digits: assume it's a hex RRGGBB like "008040"
        if (StrLen(s) = 6)
            return ("0x" s) + 0
        ; Otherwise treat as decimal (e.g., "1710822")
        return (s + 0) & 0xFFFFFF
    }

    ; Hex string form "RRGGBB"
    if !RegExMatch(s, "^[0-9A-Fa-f]{6}$")
        throw Error("NormalizeRgb: invalid RGB '" rgb "'")
    return ("0x" s) + 0
}

ColorToBGR(rgb) {
    v := NormalizeRgb(rgb)
    return ((v & 0xFF) << 16) | (v & 0xFF00) | ((v >> 16) & 0xFF)
}

ToHex6(rgb) {
    v := NormalizeRgb(rgb)
    return Format("{:06X}", v)
}

SetRichEditBg(ctrl, rgb) {
    static EM_SETBKGNDCOLOR := 0x0443
    if (!ctrl || !ctrl.Hwnd)
        return
    try bgr := ColorToBGR(rgb)
    catch as e {
        Dbg("SetRichEditBg: invalid rgb '" rgb "' -> " e.Message)
        return
    }
    SendMessage(EM_SETBKGNDCOLOR, 0, bgr, ctrl.Hwnd)
    DllCall("User32\InvalidateRect", "ptr", ctrl.Hwnd, "ptr", 0, "int", true)
    DllCall("User32\UpdateWindow",   "ptr", ctrl.Hwnd)
}

HideOutputCaret(*) {
    global OutputCtl
    if !(IsSet(OutputCtl) && OutputCtl && OutputCtl.Hwnd)
        return

    ; RichEdit can recreate its caret after activation/click/selection work.
    ; Hide it whenever this control currently owns keyboard focus.
    if (DllCall("user32\GetFocus", "ptr") = OutputCtl.Hwnd) {
        DllCall("user32\HideCaret", "ptr", OutputCtl.Hwnd)
        DllCall("user32\HideCaret", "ptr", 0)
        DllCall("user32\DestroyCaret")
    }
}

QueueHideOutputCaret(*) {
    HideOutputCaret()
    SetTimer(HideOutputCaret, -1)
}

StartOutputCaretSuppression(*) {
    QueueHideOutputCaret()
    SetTimer(HideOutputCaret, 25)
}

StopOutputCaretSuppression(*) {
    SetTimer(HideOutputCaret, 0)
}

FocusOverlaySink(*) {
    global FocusSink
    if (IsSet(FocusSink) && FocusSink && FocusSink.Hwnd)
        DllCall("user32\SetFocus", "ptr", FocusSink.Hwnd)
}

QueueFocusOverlaySink(*) {
    FocusOverlaySink()
    SetTimer(FocusOverlaySink, -1)
}

RefocusOverlaySink(wParam, lParam, msg, hwnd) {
    global Overlay
    if (IsSet(Overlay) && Overlay && hwnd = Overlay.Hwnd) {
        QueueHideOutputCaret()
        QueueFocusOverlaySink()
    }
}

HideOutputCaretOnFocus(wParam, lParam, msg, hwnd) {
    global OutputCtl
    if (IsSet(OutputCtl) && OutputCtl && hwnd = OutputCtl.Hwnd) {
        QueueHideOutputCaret()
        QueueFocusOverlaySink()
    }
}

BlockOutputMouse(wParam, lParam, msg, hwnd) {
    global OutputCtl
    if (IsSet(OutputCtl) && OutputCtl && hwnd = OutputCtl.Hwnd) {
        QueueHideOutputCaret()
        QueueFocusOverlaySink()
        return 0
    }
}

OutputMouseActivate(wParam, lParam, msg, hwnd) {
    global OutputCtl
    if (IsSet(OutputCtl) && OutputCtl && hwnd = OutputCtl.Hwnd) {
        QueueHideOutputCaret()
        QueueFocusOverlaySink()
        return 4 ; MA_NOACTIVATEANDEAT
    }
}

OutputHitTest(wParam, lParam, msg, hwnd) {
    global OutputCtl
    if (IsSet(OutputCtl) && OutputCtl && hwnd = OutputCtl.Hwnd)
        return -1 ; HTTRANSPARENT
}

SetTaskbarIcon(hwnd, icoPath) {
    if (!hwnd || !FileExist(icoPath))
        return false
    IMAGE_ICON := 1, LR_LOADFROMFILE := 0x10, LR_DEFAULTSIZE := 0x40
    hBig := DllCall("LoadImage", "ptr", 0, "str", icoPath, "uint", IMAGE_ICON
                  , "int", 0, "int", 0, "uint", LR_LOADFROMFILE|LR_DEFAULTSIZE, "ptr")
    hSmall := DllCall("LoadImage", "ptr", 0, "str", icoPath, "uint", IMAGE_ICON
                   , "int", 16, "int", 16, "uint", LR_LOADFROMFILE, "ptr")
    if (hBig)
        SendMessage(0x0080, 1, hBig, , "ahk_id " hwnd)
    if (hSmall)
        SendMessage(0x0080, 0, hSmall, , "ahk_id " hwnd)
    OnExit((*) => (
        hBig   && DllCall("DestroyIcon","ptr",hBig),
        hSmall && DllCall("DestroyIcon","ptr",hSmall)
    ))
    return true
}

; ---[ GDI brush helper ]-------------------------------------------------------
ReplaceBrush(&hBrush, rgb) {
    if (hBrush)
        DllCall("DeleteObject", "ptr", hBrush)
    ; RGB -> BGR for GDI
    bgr := ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
    hBrush := DllCall("CreateSolidBrush", "uint", bgr, "ptr")
}

HideOSRim(hwnd) {
    global OUTER_W, INNER_W, Overlay

    ; ---- constants ----
    static GWL_STYLE    := -16
    static GWL_EXSTYLE  := -20
    static WS_CAPTION      := 0x00C00000
    static WS_THICKFRAME   := 0x00040000
    static SWP_NOSIZE       := 0x0001
    static SWP_NOMOVE       := 0x0002
    static SWP_NOZORDER     := 0x0004
    static SWP_FRAMECHANGED := 0x0020

    ; locals (avoid #Warn shadowing)
    local s, exs
    s   := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", GWL_STYLE,   "ptr")
    exs := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", GWL_EXSTYLE, "ptr")

    if (IsWindows10()) {
        ; --- Windows 10: keep resizable rim, no classic title text
        ; NormalizeBordersForWin10() already ensures INNER_W != 0 here.
        s := (s | WS_THICKFRAME) & ~WS_CAPTION
        DllCall("SetWindowLongPtr", "ptr", hwnd, "int", GWL_STYLE,   "ptr", s,   "ptr")
        DllCall("SetWindowLongPtr", "ptr", hwnd, "int", GWL_EXSTYLE, "ptr", exs, "ptr")
        ; (No GUI caption/border toggles on Win10)
    } else {
        ; --- Windows 11+: keep WS_THICKFRAME (for resizing), remove caption.
        ; We use DWM attributes to suppress any visible rim.
        if (INNER_W = 0) {
            ; Frameless look, still resizable
            s := (s | WS_THICKFRAME) & ~WS_CAPTION
            if IsSet(Overlay)
                Overlay.Opt("-Caption +Resize")
        } else {
            ; Rim-only look (no title bar), resizable
            s := (s | WS_THICKFRAME) & ~WS_CAPTION
            if IsSet(Overlay)
                Overlay.Opt("-Caption +Resize")
        }
		
        ; --- DWM attributes to hide the white rim / keep visuals clean ---
        try {
            ; DWMWA_NCRENDERING_POLICY = 2 (ENABLED)
            ncrp := 2
            DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "int", 2, "int*", ncrp, "int", 4)

            ; DWMWA_WINDOW_CORNER_PREFERENCE = 33 (slight/default)
            corner := 1
            DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "int", 33, "int*", corner, "int", 4)

            ; DWMWA_BORDER_COLOR = 34 (use 0 when fully rimless)
            clr := (OUTER_W = 0 && INNER_W = 0) ? 0 : HexToBGR(BDR_OUT)
            DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "int", 34, "int*", clr, "int", 4)

            ; DWMWA_CAPTION_COLOR = 35 (match background-ish to avoid glow)
            capClr := HexToBGR(BDR_OUT)
            DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "int", 35, "int*", capClr, "int", 4)

            ; DWMWA_SYSTEMBACKDROP_TYPE = 38 (0/1 = off)
            noBackdrop := 1
            DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "int", 38, "int*", noBackdrop, "int", 4)

            ; DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (0 = light) to prevent white rim injection
            dark := 0
            DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "int", 20, "int*", dark, "int", 4)
        }
        DllCall("SetWindowLongPtr", "ptr", hwnd, "int", GWL_STYLE,   "ptr", s,   "ptr")
        DllCall("SetWindowLongPtr", "ptr", hwnd, "int", GWL_EXSTYLE, "ptr", exs, "ptr")
    }

    ; Recalculate non-client metrics without moving/resizing
    DllCall("user32\SetWindowPos"
        , "ptr", hwnd, "ptr", 0
        , "int", 0, "int", 0, "int", 0, "int", 0
        , "uint", SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
    return
}

IsWindows10() {
    ; Win11 has build >= 22000. Everything below is Win10/older.
    local verParts
    if (SubStr(A_OSVersion, 1, 4) = "10.0") {
        verParts := StrSplit(A_OSVersion, ".")
        return (verParts.Length >= 3 && verParts[3] < 22000)
    }
    return false
}

NormalizeBordersForWin10() {
    global OUTER_W, INNER_W
    ; Win10 safety: never allow INNER_W to be 0 (prevents resize loss)
    if (IsWindows10() && INNER_W = 0) {
        INNER_W := 1
    }
}

; --- Win11 frameless resize hit-test (keeps Win10 untouched) ---
global __HT_HOOK := false

UpdateFramelessResizeHook() {
    global __HT_HOOK, Overlay, INNER_W
    ; Only enable on Windows 11+ AND when INNER_W == 0 (truly frameless).
    if (!IsWindows10() && IsSet(Overlay) && IsObject(Overlay) && Overlay && INNER_W = 0) {
        if (!__HT_HOOK) {
            OnMessage(0x0084, NCHitTest) ; WM_NCHITTEST
            __HT_HOOK := true
        }
    } else if (__HT_HOOK) {
        ; Remove the hook so normal behavior applies elsewhere.
        OnMessage(0x0084, 0)
        __HT_HOOK := false
    }
}

NCHitTest(wParam, lParam, msg, hwnd) {
    global Overlay, INNER_W
    ; Guard: only for our overlay window, on Win11+, frameless mode.
    if (IsWindows10() || !IsSet(Overlay) || hwnd != Overlay.Hwnd || INNER_W != 0)
        return  ; let default handling run

    ; Get screen coordinates from lParam (signed 16-bit)
    sx := lParam & 0xFFFF
    sy := (lParam >> 16) & 0xFFFF
    if (sx & 0x8000)
        sx := sx - 0x10000
    if (sy & 0x8000)
        sy := sy - 0x10000

    ; Window rect in screen coords
    rc := Buffer(16, 0)
    DllCall("user32\GetWindowRect", "ptr", hwnd, "ptr", rc)
    left   := NumGet(rc, 0, "int")
    top    := NumGet(rc, 4, "int")
    right  := NumGet(rc, 8, "int")
    bottom := NumGet(rc,12, "int")
    w := right - left
    h := bottom - top

    rx := sx - left
    ry := sy - top

    ; Grip thickness: DPI-aware, minimum ~6px at 96 DPI
    dpi := GetWindowDPI(hwnd)
    grip := Max(6, Round(dpi / 16.0))  ; 96/16 = 6

    ; Determine edges
    leftEdge   := (rx < grip)
    rightEdge  := (rx >= w - grip)
    topEdge    := (ry < grip)
    bottomEdge := (ry >= h - grip)

    ; Corners first
    if (topEdge && leftEdge)
        return 13  ; HTTOPLEFT
    if (topEdge && rightEdge)
        return 14  ; HTTOPRIGHT
    if (bottomEdge && leftEdge)
        return 16  ; HTBOTTOMLEFT
    if (bottomEdge && rightEdge)
        return 17  ; HTBOTTOMRIGHT

    ; Sides
    if (topEdge)
        return 12  ; HTTOP
    if (leftEdge)
        return 10  ; HTLEFT
    if (rightEdge)
        return 11  ; HTRIGHT
    if (bottomEdge)
        return 15  ; HTBOTTOM

    ; Otherwise, don't interfere (client area stays normal so text selection works)
    return
}

GetWindowDPI(hwnd) {
    dpi := 96
    try dpi := DllCall("user32\GetDpiForWindow", "ptr", hwnd, "uint")
    return (dpi > 0) ? dpi : 96
}

PickInstalledFont(fonts*) {
    for f in fonts {
        h := DllCall("gdi32\CreateFont","int",0,"int",0,"int",0,"int",0,"int",400
            ,"uint",0,"uint",0,"uint",0,"uint",0,"uint",0,"uint",0,"uint",0,"str",f,"ptr")
        if (h) {
            DllCall("gdi32\DeleteObject","ptr",h)
            return f
        }
    }
    return "Segoe UI"
}
SafeDelete(path) {
    try if FileExist(path)
        FileDelete(path)
}
ExpandEnv(str) {
    if !str
        return ""
    capLen := 32767, buf := Buffer(capLen*2, 0)
    DllCall("Kernel32\ExpandEnvironmentStringsW","str",str,"ptr",buf,"int",capLen,"int")
    return StrGet(buf, "UTF-16")
}
ResolvePath(p) {
    if !p
        return ""
    expanded := ExpandEnv(p)
    if RegExMatch(expanded, 'i)^(?:[A-Z]:\\|\\\\)')
        return expanded
    base := (IsSet(APP_ROOT) && APP_ROOT) ? APP_ROOT : A_ScriptDir
    if RegExMatch(expanded, 'i)^(?:\./|\.\\|\.\./|\.\.\\)')
        return base "\" expanded
    return base "\" expanded
}

; --- Private fonts from .\fonts (FR_PRIVATE, process-local) ---------------------
LoadPrivateFonts(){
    static loaded := false
    if (loaded)
        return
        dir := ResolvePath(".\fonts")
    if DirExist(dir) {
        for ext in ["ttf","otf","ttc"] {
            Loop Files, dir "\*." ext, "F" {
                try DllCall("AddFontResourceEx", "str", A_LoopFileFullPath, "uint", 0x10, "ptr", 0)
            }
        }
    }
    loaded := true
}

; ===== Debug helpers (non-invasive) =====
; Read from env: JRPG_DEBUG=0 disables logging; anything else enables it (default ON if unset)
global __DBG_ENABLED := (EnvGet("JRPG_DEBUG") != "0")
global __DBG_LOG := ""   ; set after Settings folder exists

DbgInit() {
    global __DBG_ENABLED, __DBG_LOG
    if !__DBG_ENABLED
        return
    try {
        if FileExist(__DBG_LOG) {
            f := FileOpen(__DBG_LOG, "r")
            if IsObject(f) {
                sz := f.Length, f.Close()
                if (sz > 512000)  ; ~0.5MB rotate
                    FileDelete(__DBG_LOG)
            }
        }
    }
}
Dbg(msg) {
    global __DBG_ENABLED, __DBG_LOG
    if !__DBG_ENABLED
        return
    try {
        ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        FileAppend("[" ts "] OVERLAY  " msg "`r`n", __DBG_LOG, "UTF-8")
    }
}
DbgRect(tag, x, y, w, h) => Dbg(tag ": x=" x " y=" y " w=" w " h=" h)
DbgMonitors() {
    info := ""
    cnt := MonitorGetCount()
    Loop cnt {
        MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
        info .= "#" A_Index "=(" l "," t ")-(" r "," b ") "
    }
    return Trim(info)
}

EnsureBoundsOnScreen(bounds) {
    local isVisible := false
    local visibleMargin := 64
    Loop MonitorGetCount() {
        MonitorGetWorkArea(A_Index, &monL, &monT, &monR, &monB)
        if (bounds["x"] < monR - visibleMargin && bounds["x"] + bounds["w"] > monL + visibleMargin
         && bounds["y"] < monB - visibleMargin && bounds["y"] + bounds["h"] > monT + visibleMargin) {
            isVisible := true
            break
        }
    }
    if !isVisible {
        DbgRect("Clamp(not visible) before", bounds["x"], bounds["y"], bounds["w"], bounds["h"])
        bounds["x"] := 120, bounds["y"] := 120
        DbgRect("Clamp(not visible) after", bounds["x"], bounds["y"], bounds["w"], bounds["h"])
    }
    return bounds
}

; Delete a batch of files (used by delete-after-use scheduling)
DeleteFiles(list) {
    for f in list {
        try FileDelete(f)
    }
}

; Ensure an INI key exists; if not, write a default
EnsureIniDefault(iniPath, section, key, default) {
    sentinel := "__MISSING__"
    val := IniRead(iniPath, section, key, sentinel)
    if (val = sentinel)
        IniWrite(default, iniPath, section, key)
}

; === NEW: centralize glossary env export (profile-aware) ===
ExportGlossaryEnv() {
    global ControlIni
    ; Make sure profile keys exist
    EnsureIniDefault(ControlIni, "cfg", "jp2enGlossaryProfile", "default")
    EnsureIniDefault(ControlIni, "cfg", "en2enGlossaryProfile", "default")

    jpProf := Trim(IniRead(ControlIni, "cfg", "jp2enGlossaryProfile", "default"))
    if (jpProf = "")
        jpProf := "default"
    enProf := Trim(IniRead(ControlIni, "cfg", "en2enGlossaryProfile", "default"))
    if (enProf = "")
        enProf := "default"

jpPath := ResolvePath(".\Settings\glossaries\" jpProf "\jp2en.txt")
enPath := ResolvePath(".\Settings\glossaries\" enProf "\en2en.txt")

    ; Export multiple compatible names so Python/audio pick it up
    m := Map(
        "JP2EN_GLOSSARY_PATH", jpPath,
        "EN2EN_GLOSSARY_PATH", enPath,
        "GLOSSARY_JP2EN",      jpPath,
        "GLOSSARY_EN2EN",      enPath,
        "JRPG_JP2EN_GLOSSARY", jpPath,
        "JRPG_EN2EN_GLOSSARY", enPath,
        ; policy hint for downstream: apply only to English translation
        "GLOSSARY_SCOPE",      "translation_only"
    )
    for k, v in m
        EnvSet(k, v)
}

ClearAudioFile() {
    global AudioTxt
    try {
        tmpPath := AudioTxt ".tmp"
        FileAppend("", tmpPath, "UTF-8")
        FileMove(tmpPath, AudioTxt, 1)
    } catch {
        try f := FileOpen(AudioTxt, "w")
        if IsObject(f)
            f.Close()
    }
}

ClearOcrFile() {
    global OcrTxt
    try {
        tmpPath := OcrTxt ".tmp"
        FileAppend("", tmpPath, "UTF-8")
        FileMove(tmpPath, OcrTxt, 1)
    } catch {
        try f := FileOpen(OcrTxt, "w")
        if IsObject(f)
            f.Close()
    }
}

ClearExplainerFile() {
    global ExplainerTxt
    try {
        tmpPath := ExplainerTxt ".tmp"
        FileAppend("", tmpPath, "UTF-8")
        FileMove(tmpPath, ExplainerTxt, 1)
    } catch {
        try f := FileOpen(ExplainerTxt, "w")
        if IsObject(f)
            f.Close()
    }
}

HandleGuiClose(pidToClose, guiObj) {
    ; Save current bounds before exiting so next launch restores correctly
    try {
        SaveOverlayBounds()
    } catch as e {
        ; ignore
    }
    ; Exit gracefully so OnExit/Cleanup can run
    ExitApp()
}

; ================= JRPG Translation Overlay =================

; Settings folder (portable) — always rooted at app root
portableRoot := APP_ROOT "\Settings"
if !DirExist(portableRoot)
    DirCreate(portableRoot)

; --- Explainer overlay mode (optional) ---
global __EXPLAIN_MODE := !!EnvGet("EXPLAIN_MODE")
global __WIN_TITLE  := __EXPLAIN_MODE ? "Explainer" : "Translator"
Dbg("Mode: " (__EXPLAIN_MODE ? "EXPLAINER" : "TRANSLATOR"))

; If a window with the same title already exists, just activate it and exit.
oldMode := A_TitleMatchMode
SetTitleMatchMode 3 ; 3 = Match the title exactly
if WinExist(__WIN_TITLE) {
    SetTitleMatchMode oldMode ; Restore the old mod…13381 tokens truncated…      return false
    } catch {
        return false
    }

    for candidateHwnd in WinGetList() {
        try {
            title := WinGetTitle("ahk_id " candidateHwnd)
            if (title != "Translator" && title != "Explainer")
                continue
            if !DllCall("user32\IsWindowVisible", "ptr", candidateHwnd, "int")
                continue
            return candidateHwnd = overlayHwnd
        }
    }
    return false
}

OverlayShouldReceiveGlobalWheel(*) {
    global Overlay
    if !(IsSet(Overlay) && Overlay && Overlay.Hwnd)
        return false
    return IsTopVisibleJrpgOverlay(Overlay.Hwnd)
}

RegisterOverlayWheelHotkeys() {
    global Overlay
    if !(IsSet(Overlay) && Overlay && Overlay.Hwnd)
        return

    HotIf(OverlayShouldReceiveGlobalWheel)
    try Hotkey("WheelUp", ControllerWheelUp, "On")
    try Hotkey("WheelDown", ControllerWheelDown, "On")
    HotIf()
}

; === OnResize layout ===
OnResize(guiObj, minmax := "", w := 0, h := 0) {
    global OUTER_W, INNER_W, BOX_PAD
    global RectOuter, RectInner, RectPanel, OutputCtl, ClickShield, StatusCtl
    global BOX_BG, hBrushEdit, RESIZE_LOCK
    global __OcrText, __AudioText

    ; default the flag if first-time
    if !IsSet(RESIZE_LOCK)
        RESIZE_LOCK := false

    if (w=0 || h=0)
        guiObj.GetClientPos(, , &w, &h)
    if (w <= 0 || h <= 0)
        return

    RectOuter.Move(0, 0, w, h)
    ix := OUTER_W, iy := OUTER_W
    iw := Max(w - OUTER_W*2, 1), ih := Max(h - OUTER_W*2, 1)
    RectInner.Move(ix, iy, iw, ih)
    px := OUTER_W + INNER_W, py := OUTER_W + INNER_W
    pw := Max(w - 2*(OUTER_W + INNER_W), 1), ph := Max(h - 2*(OUTER_W + INNER_W), 1)
    RectPanel.Move(px, py, pw, ph)

    tx := px + BOX_PAD
    ty := py + BOX_PAD
    tw := Max(pw - BOX_PAD*2, 60)
    th := Max(ph - BOX_PAD*2, 60)
    OutputCtl.Move(tx, ty, tw, th)
    if (IsSet(ClickShield) && ClickShield) {
        ClickShield.Move(tx, ty, tw, th)
        DllCall("user32\SetWindowPos", "ptr", ClickShield.Hwnd, "ptr", 0
            , "int", 0, "int", 0, "int", 0, "int", 0
            , "uint", 0x0001 | 0x0002 | 0x0010) ; NOSIZE|NOMOVE|NOACTIVATE
    }
    if (IsSet(StatusCtl) && StatusCtl) {
        sw := 34, sh := 34
        StatusCtl.Move(tx + tw - sw - 8, ty + 6, sw, sh)
        DllCall("user32\SetWindowPos", "ptr", StatusCtl.Hwnd, "ptr", 0
            , "int", 0, "int", 0, "int", 0, "int", 0
            , "uint", 0x0001 | 0x0002 | 0x0010) ; NOSIZE|NOMOVE|NOACTIVATE
    }

    if (!RESIZE_LOCK) {
        try {
            if (!hBrushEdit)
                hBrushEdit := MakeBrush(BOX_BG)
            SetRichEditBg(OutputCtl, BOX_BG)
            DllCall("User32\RedrawWindow", "ptr", OutputCtl.Hwnd, "ptr", 0, "ptr", 0
                , "uint", 0x0001 | 0x0004 | 0x0100)
        }
        Dbg("OnResize repaint: BG=" BOX_BG " rect=(" tx "," ty "," tw "," th ")")
    } else {
        Dbg("OnResize suppressed during live drag: rect=(" tx "," ty "," tw "," th ")")
    }

    if (Trim(__AudioText) != "" && Trim(__OcrText) = "")
        SetTimer(ScrollOutputToBottom, -1)
}

GetLatestPng(dir) {
    newestFile := "", newestTime := ""
    Loop Files dir "\*.png" {
        t := A_LoopFileTimeModified
        if (StrCompare(t, newestTime) > 0)
            newestTime := t, newestFile := A_LoopFileFullPath
    }
    return newestFile
}
GetFileSize(path) {
    if !FileExist(path)
        return -1
    f := FileOpen(path, "r")
    if !IsObject(f)
        return -1
    size := f.Length, f.Close()
    return size
}
WaitForNewPngStable(dir, baselinePath, timeoutMs := 10000, settleMs := 250) {
    start := A_TickCount
    loop {
        f := GetLatestPng(dir)
        if (f && f != baselinePath) {
            s1 := GetFileSize(f)
            Sleep(settleMs)
            s2 := GetFileSize(f)
            if (s1 > 0 && s1 = s2)
                return f
        }
        if (A_TickCount - start > timeoutMs)
            return ""
        Sleep(100)
    }
}
SendShareXRegionHotkey() {
    global sharexRegionHotkey
    SendInput("{Ctrl up}{Alt up}{Shift up}")
    Sleep(15)
    if (sharexRegionHotkey = "^!F1")
        SendInput("{Ctrl down}{Alt down}{F1}{Alt up}{Ctrl up}")
    else
        Send(sharexRegionHotkey)
    Sleep(25)
}

appendLatestShot(*) {
    global ShotBuf, captureDir
    cd := ResolvePath(captureDir)
    if !DirExist(cd) {
        showText("(Config error) captureDir not found:`n" cd)
        return
    }
    ToolTip("Capturing (Ctrl+Alt+F1)…")
    SetTimer(() => ToolTip(""), -700)
    baseline := GetLatestPng(cd)
    SendShareXRegionHotkey()
    f := WaitForNewPngStable(cd, baseline, 10000, 250)
    if (!f) {
        f := GetLatestPng(captureDir)
        if (!f) {
            ToolTip("No PNG found in: " captureDir)
            SetTimer(() => ToolTip(""), -1200)
            return
        }
        ToolTip("(No fresh file) Using latest: " f)
        SetTimer(() => ToolTip(""), -1500)
    } else {
        ToolTip("Captured: " f)
        SetTimer(() => ToolTip(""), -900)
    }
    if (ShotBuf.Length && ShotBuf[ShotBuf.Length] = f) {
        ToolTip("Skipped duplicate shot.")
        SetTimer(() => ToolTip(""), -900)
        return
    }
    ShotBuf.Push(f)
    ToolTip("Buffered shots: " ShotBuf.Length)
    SetTimer(() => ToolTip(""), -700)
}

redefineRegion(*) {
    ; Open tiny picker: Region or Window (same behavior as Control Panel button)
    ; Default to region if Shift is held
    if GetKeyState("Shift", "P")
        StartPickRegion()
    else
        StartPickWindow()
}

flushTranslate(files := unset) {
    global ShotBuf, pythonExe, translatorPy, ControlIni
    local fileList, n

    px := ResolvePath(pythonExe)
    tp := ResolvePath(translatorPy)

    if IsSet(files)
        fileList := files
    else
        fileList := ShotBuf.Clone()

    n := fileList.Length
    if (n = 0) {
        ToolTip("Buffer empty.")
        SetTimer(() => ToolTip(""), -1200)
        return
    }

    if !FileExist(tp) {
        showText("(Config error) translatorPy not found:`n" tp)
        ShotBuf := []
        return
    }
    if !FileExist(px) {
        showText("(Config error) python.exe not found:`n" px)
        ShotBuf := []
        return
    }

    ToolTip()
    ShowOverlayStatus()
	; === Re-read Control Panel settings just-in-time and export to env ===
; Read the same keys the Control Panel writes into control.ini
provider       := IniRead(ControlIni, "cfg", "imgProvider",       "openai")
imgModel       := IniRead(ControlIni, "cfg", "imgModel",          "gpt-4o")
geminiImgModel := IniRead(ControlIni, "cfg", "geminiImgModel",    "gemini-2.5-flash")
promptProfile  := IniRead(ControlIni, "cfg", "promptProfile",     "default")
; Back-compat: first try imgPostproc, then legacy 'post'
imgPostproc    := IniRead(ControlIni, "cfg", "imgPostproc", IniRead(ControlIni, "cfg", "post", "tt"))

; Provider + model
EnvSet("PROVIDER", provider)
if (provider = "gemini") {
    m := geminiImgModel
    if (SubStr(m, 1, 7) != "models/")
        m := "models/" . m
    EnvSet("GEMINI_MODEL_NAME", m)
} else {
    EnvSet("MODEL_NAME", imgModel)
}

; --- Guessed-subject highlighting from Control Panel (instant effect) ---
hlGuess := IniRead(ControlIni, "cfg", "highlightGuessed", 1)
EnvSet("SHOT_ITALICIZE_GUESSED", (hlGuess ? "1" : "0"))
EnvSet("SHOT_GUESS_DELIM", Chr(0x60))  ; use backtick as the delimiter

; --- Speaker-name colorization (JP+EN; Python strips 「…」 and wraps ⟦name⟧ when ON) ---
colorSpeaker := IniRead(ControlIni, "cfg", "colorSpeaker", 1)
EnvSet("SHOT_COLOR_SPEAKER", (colorSpeaker ? "1" : "0"))

; Prompt + post-processing
EnvSet("PROMPT_PROFILE", promptProfile)
EnvSet("PROMPT_FILE", "")
EnvSet("POSTPROC_MODE", imgPostproc)

; --- Pass glossary paths (profile-aware) & policy ---
ExportGlossaryEnv()
; --- end glossary envs ---


; === end just-in-time settings ===

    args := ""
    args := ""
    for f in fileList
        args .= Format(' "{}"', f)

    cmd := Format('cmd /c chcp 65001>nul & "{}" "{}"{}', px, tp, args)
    Run(cmd, , "Hide")

    ; Schedule deletion of this batch (optional)
    if (deleteAfterUse) {
        toDelete := fileList.Clone()
        SetTimer(() => DeleteFiles(toDelete), -deleteDelayMs)
    }

    ShotBuf := []
}

oneshotTranslate(*) {
    global Cap_Mode, captureDir
    InitGDIPlus()
    if (!__GDI_Ready) {
        showText("GDI+ not available.")
        return
    }

    ; ensure we actually have a selection
    if (Cap_Mode != "region" && Cap_Mode != "window") {
        ToolTip("Pick region or window first (Ctrl+Alt+F3 or via Control Panel)."), SetTimer(() => ToolTip(""), -1400)
        return
    }

    ToolTip()
    ShowOverlayStatus()

    local f := ""
    ok := CaptureOnceToFile(&f)
    if (!ok || f = "") {
        ToolTip("No target set."), SetTimer(() => ToolTip(""), -1200)
        showText("No target set — click “Capture” to pick a region/window.")
        return
    }
    ; silent mode: no notification needed, but keep a short toast for sanity
    ; ToolTip("Captured: " f), SetTimer(() => ToolTip(""), -600)

    flushTranslate([f])
}

;------------------------------------------------------------------------------
; screenshot_translation: send ALL buffered screenshots in one LLM request
;------------------------------------------------------------------------------
FlushBufferedScreenshots(*) {
    global ShotBuf

    ; Nothing queued?
    if !(IsSet(ShotBuf)) || (ShotBuf.Length = 0) {
        ToolTip("No buffered screenshots.")
        SetTimer(() => ToolTip(""), -1200)
        return
    }

    ; Clone then clear the buffer so repeated presses don’t resend old shots.
    files := ShotBuf.Clone()    ; AHK v2 Array.Clone()
    ShotBuf := []

    try {
        ; Your pipeline already accepts an array (oneshot uses: flushTranslate([f]))
        flushTranslate(files)
        ToolTip("Sent " files.Length " screenshot(s).")
        SetTimer(() => ToolTip(""), -1200)
    } catch as ex {
        ; Restore on failure so user can retry
        ShotBuf := files
        ToolTip("Flush failed: " ex.Message)
        SetTimer(() => ToolTip(""), -1600)
    }
}

;------------------------------------------------------------------------------
; take_screenshot: native capture only (no translation)
; - saves a PNG via CaptureOnceToFile(&f)
; - adds it to ShotBuf so you can batch with flushTranslation later
;------------------------------------------------------------------------------
TakeScreenshotOnly(*) {
    global ShotBuf

    ; Make sure there is a configured selection (region/window)
    if (Cap_Mode != "region" && Cap_Mode != "window") {
        ToolTip("Pick region or window first (Ctrl+Alt+F3 or via Control Panel).")
        SetTimer(() => ToolTip(""), -1400)
        return
    }

    local f := ""
    ok := CaptureOnceToFile(&f)
    if (!ok || f = "") {
        ToolTip("No screenshot.")
        SetTimer(() => ToolTip(""), -1200)
        return
    }

    ; De-dup consecutive identical file
    if (ShotBuf.Length && ShotBuf[ShotBuf.Length] = f) {
        ToolTip("Skipped duplicate shot.")
        SetTimer(() => ToolTip(""), -900)
        return
    }

    ShotBuf.Push(f)
    ToolTip("Buffered shots: " ShotBuf.Length)
    SetTimer(() => ToolTip(""), -700)
}

; ---------- Dynamic hotkeys (load from INI and allow live reload) ----------
global __HK_REG := Map()
global __CMD_TOGGLE_EXPL := A_Temp "\JRPG_Overlay\cmd.toggle_explainer"

UnregisterAllHotkeys() {
    global __HK_REG
    ; __HK_REG maps hk => callback object. Turn each one OFF explicitly.
    for hk, cb in __HK_REG {
        try Hotkey(hk, cb, "Off")
    }
    __HK_REG := Map()
}

RegisterAllHotkeys() {
    global __HK_REG, ControlIni, __EXPLAIN_MODE
    UnregisterAllHotkeys()

    ; Read configured hotkeys (same section Control Panel writes)
    m := Map()
    if (!__EXPLAIN_MODE) {
        m["append_to_buffer"]   := IniRead(ControlIni, "hotkeys", "append_to_buffer",   "^1")
        m["flush_translation"]  := IniRead(ControlIni, "hotkeys", "flush_translation",  "^2")
        m["redefine_region"]    := IniRead(ControlIni, "hotkeys", "redefine_region",    "^3")
    	
        ; Prefer new key name; fall back to legacy entry if it exists
        tmpHk := IniRead(ControlIni, "hotkeys", "screenshot_translate", "")
        if (tmpHk = "")
            tmpHk := IniRead(ControlIni, "hotkeys", "oneshot_translate", "")
        m["oneshot_translate"]  := tmpHk
        m["recapture_region"]   := IniRead(ControlIni, "hotkeys", "recapture_region", "")
        m["take_screenshot"]    := IniRead(ControlIni, "hotkeys", "take_screenshot", "")
        m["screenshot_translation"] := IniRead(ControlIni, "hotkeys", "screenshot_translation", "")
    }
    
	; Show/Hide rows replace legacy ^0 (translator) / ^+0 (explainer)
    if (__EXPLAIN_MODE) {
         m["hide_show_explainer"] := IniRead(ControlIni, "hotkeys", "hide_show_explainer", "^+0")
     } else {
         m["hide_show_translator"] := IniRead(ControlIni, "hotkeys", "hide_show_translator", "^0")
     }
 
     ; Topmost is now configurable separately (no default here to avoid collision)
     m["toggle_top"]         := IniRead(ControlIni, "hotkeys", "toggle_top", "")
     m["toggle_audio"]       := IniRead(ControlIni, "hotkeys", "toggle_audio", "")
	
    ; Map action -> function
    fun := Map()
    fun["append_to_buffer"]  := appendLatestShot
    fun["flush_translation"] := flushTranslate
    fun["redefine_region"]   := redefineRegion
    fun["oneshot_translate"] := oneshotTranslate
    fun["toggle_top"]        := ToggleTop
    fun["toggle_audio"]      := ToggleAudioListening
	fun["recapture_region"]  := StartPickRegion
	fun["take_screenshot"]   := TakeScreenshotOnly
	fun["screenshot_translation"] := FlushBufferedScreenshots

    ; Bind Show/Hide rows to the EXACT same function as legacy ^0/^+0
    if (__EXPLAIN_MODE) {
        fun["hide_show_explainer"] := ToggleTop
    } else {
        fun["hide_show_translator"] := ToggleTop
}


    for action, hk in m {
    hk := Trim(hk)
    if (hk = "" || !fun.Has(action))
        continue

    ; --- Optional hardening: ensure no stale binding remains for this key
    if __HK_REG.Has(hk) {
        try Hotkey(hk, __HK_REG[hk], "Off")
        __HK_REG.Delete(hk)
    }
    ; ---

    try {
        cb := fun[action]
        Hotkey(hk, cb, "On")             ; bind (explicitly ON in v2)
        __HK_REG[hk] := cb               ; store callback object for proper unbind
    } catch as e {
        Dbg("Failed to bind hotkey " hk " for " action ": " e.Message)
    }
}
}

; --- command/flag files (top-level init) ---
global OVERLAY_TEMP_DIR := A_Temp "\JRPG_Overlay"
if !DirExist(OVERLAY_TEMP_DIR)
    DirCreate(OVERLAY_TEMP_DIR)

global __HK_FLAG := OVERLAY_TEMP_DIR "\hotkeys.reload"
global __CMD_TOGGLE_EXPL := OVERLAY_TEMP_DIR "\cmd.toggle_explainer"
global __CMD_ONESHOT_TRANSLATE := OVERLAY_TEMP_DIR "\cmd.oneshot_translate"
global __CMD_TAKE_SCREENSHOT := OVERLAY_TEMP_DIR "\cmd.take_screenshot"
global __CMD_SCREENSHOT_TRANSLATION := OVERLAY_TEMP_DIR "\cmd.screenshot_translation"
global __CMD_EXPLAIN_START := OVERLAY_TEMP_DIR "\cmd.explain_start"
; ------------------------------------------

CheckHotkeyReload() {
    global __HK_FLAG
    if FileExist(__HK_FLAG) {
        try FileDelete(__HK_FLAG)
        RegisterAllHotkeys()
        ToolTip("🔁 Hotkeys reloaded"), SetTimer(() => ToolTip(""), -700)
    }
}

CheckCmdSignals() {
    global __CMD_TOGGLE_EXPL, __CMD_ONESHOT_TRANSLATE, __CMD_TAKE_SCREENSHOT, __CMD_SCREENSHOT_TRANSLATION
    global __CMD_EXPLAIN_START
    global __EXPLAIN_MODE
    ; Only Explainer reacts to this command
    if (__EXPLAIN_MODE && FileExist(__CMD_TOGGLE_EXPL)) {
        try FileDelete(__CMD_TOGGLE_EXPL)
        try ToggleTop()   ; do the safe, internal toggle
    }
    if (__EXPLAIN_MODE && FileExist(__CMD_EXPLAIN_START)) {
        try FileDelete(__CMD_EXPLAIN_START)
        try ShowOverlayStatus()
    }
    ; Only Translator reacts to screenshot commands
    if (__EXPLAIN_MODE)
        return
    if FileExist(__CMD_ONESHOT_TRANSLATE) {
        try FileDelete(__CMD_ONESHOT_TRANSLATE)
        try oneshotTranslate()
    }
    if FileExist(__CMD_TAKE_SCREENSHOT) {
        try FileDelete(__CMD_TAKE_SCREENSHOT)
        try TakeScreenshotOnly()
    }
    if FileExist(__CMD_SCREENSHOT_TRANSLATION) {
        try FileDelete(__CMD_SCREENSHOT_TRANSLATION)
        try FlushBufferedScreenshots()
    }
}

; ---------- end dynamic hotkeys ----------

; ================ HOTKEYS ======================
; use dynamic registration (from INI)
RegisterAllHotkeys()
; poll for live-reload flag dropped by Control Panel
SetTimer(CheckHotkeyReload, 250)

^w::ExitApp()
^+o::ResetWindow()

ToggleAudioListening(){
    global OverlayDir, PauseFlag
    if !DirExist(OverlayDir)
        DirCreate(OverlayDir)
    if FileExist(PauseFlag) {
        FileDelete(PauseFlag)
        ToolTip("🎙 Listening: ON"), SetTimer(() => ToolTip(""), -900)
    } else {
        FileAppend("", PauseFlag, "UTF-8")
        ClearAudioFile()
        ToolTip("⏸ Listening: OFF"), SetTimer(() => ToolTip(""), -900)
    }
}

PollExplainerFile(){
    global ExplainerTxt, __LastExplainRaw, __OcrText, __AudioText
    raw := ""
    try if FileExist(ExplainerTxt)
        raw := FileRead(ExplainerTxt, "UTF-8")
    if (raw = __LastExplainRaw)
        return
    __LastExplainRaw := raw

    ; normalize newlines to CRLF (what the Edit control expects)
    norm := StrReplace(raw, "`r`n", "`n")
    norm := StrReplace(norm, "`r", "`n")
    norm := StrReplace(norm, "`n", "`r`n")
    newText := Trim(norm)

    if (newText != "") {
        __OcrText := newText     ; explainer text goes in the "OCR" slot
        __AudioText := ""        ; keep audio area empty in explain mode
    } else {
        __OcrText := ""
    }
    HideOverlayStatus()
    RenderCombined()
    ScrollOutputToTop()
}

PollOcrFile() {
    global OcrTxt, __LastOcrRaw, __OcrText, __AudioText
    raw := ""
    try if FileExist(OcrTxt)
        raw := FileRead(OcrTxt, "UTF-8")
    if (raw = __LastOcrRaw)
        return
    __LastOcrRaw := raw

    norm := StrReplace(raw, "`r`n", "`n")
    norm := StrReplace(norm, "`r", "`n")
    norm := StrReplace(norm, "`n", "`r`n")
    
    __OcrText := Trim(norm)
    __AudioText := "" ; Clear any old audio text
    HideOverlayStatus()
    RenderCombined()
    ScrollOutputToTop()
}

; Poll %TEMP%\JRPG_Overlay\audio.txt (normal translator mode)
PollAudioSubtitle(){
    global AudioTxt, __LastAudioRaw, __AudioText, __OcrText
    raw := ""
    try if FileExist(AudioTxt)
        raw := FileRead(AudioTxt, "UTF-8")
    if (raw = __LastAudioRaw)
        return
    __LastAudioRaw := raw

    ; normalize to CRLF for the Edit control
    norm := StrReplace(raw, "`r`n", "`n")
    norm := StrReplace(norm, "`r", "`n")
    norm := StrReplace(norm, "`n", "`r`n")

    __AudioText := Trim(norm)
    __OcrText := "" ; new audio clears old screenshot text
    RenderCombined()
    ScrollOutputToBottom()
}

; ======== Context menu ========
CtxCopyText(*) {
    global OutputCtl
    try A_Clipboard := OutputCtl.Text
    QueueHideOutputCaret()
    QueueFocusOverlaySink()
}
CtxClear(*) {
    global __OcrText, __AudioText
    __OcrText := ""
    __AudioText := ""
    RenderCombined()
}

; ====================== Cleanup ======================
OnExit(Cleanup)
Cleanup(*) {
    global hBrushEdit
    StopOutputCaretSuppression()
    SaveOverlayBounds()
    if (hBrushEdit)
        DllCall("gdi32\DeleteObject", "ptr", hBrushEdit)
}

