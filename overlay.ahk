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
    SetTitleMatchMode oldMode ; Restore the old mode before exiting
    WinActivate
    ExitApp
}
SetTitleMatchMode oldMode ; Restore the old mode if window wasn't found

; point debug log into Settings, then init
global __DBG_LOG := portableRoot "\debug.log"
DbgInit()
Dbg("=== START ===")
Dbg("AHK v" A_AhkVersion "  DPIAware: per-monitor (GetDpiForWindow)")
Dbg("Monitors: " MonitorGetCount() "  workareas: " DbgMonitors())
Dbg("EXPLAIN_MODE=" __EXPLAIN_MODE "  title=" __WIN_TITLE)

global ControlIni := portableRoot "\control.ini"
global BoundsIni  := portableRoot . (__EXPLAIN_MODE ? "\overlay_explainer.ini" : "\overlay_translator.ini")
LoadCfg(k, d) => IniRead(ControlIni, (__EXPLAIN_MODE ? "cfg_explainer" : "cfg"), k, d)

; NOTE: read the same keys Control Panel writes
global BOX_BG    := LoadCfg("boxBg",    "102040")
global BDR_OUT   := LoadCfg("bdrOut",   "F8F8F8")
global BDR_IN    := LoadCfg("bdrIn",    "84A9FF")
global TXT_COLOR := LoadCfg("txtColor", "FFFFFF")
global BOX_PAD   := 16
global OUTER_W   := Integer(LoadCfg("bdrOutW", 3))
global INNER_W   := Integer(LoadCfg("bdrInW",  1))
; Win10 safety: if both borders were set to 0, keep inner at 1px so resizing works
NormalizeBordersForWin10()
; Load private fonts so non-installed faces from .\fonts are usable
LoadPrivateFonts()
global FONT_NAME := LoadCfg("fontName", "Segoe UI")
global FONT_SIZE := Integer(LoadCfg("fontSize", 22))

; --- export glossary envs on startup (screenshot + audio) ---
ExportGlossaryEnv()

pythonExe    := ".\python\python.exe"
translatorPy := ".\scripts\screenshot_translator.py"
; Read from Control Panel’s INI, fall back to default
captureDir := IniRead(ControlIni, "paths", "captureDir", ".\Settings\Screenshots")
; Ensure the ShareX capture folder exists
cap := ResolvePath(captureDir)
if !DirExist(cap)
    DirCreate(cap)
captureDir := cap

; Ensure the ShareX capture folder exists
cap := ResolvePath(captureDir)
if !DirExist(cap)
    DirCreate(cap)
captureDir := cap

; Screenshot lifecycle defaults (write to INI if missing)
EnsureIniDefault(ControlIni, "paths", "deleteAfterUse", 1)      ; 1 = delete each batch after use, 0 = keep
EnsureIniDefault(ControlIni, "paths", "deleteDelayMs", 10000)  ; 10 seconds
EnsureIniDefault(ControlIni, "paths", "captureRetentionDays", 0)

; Screenshot lifecycle settings
deleteAfterUse := Integer(IniRead(ControlIni, "paths", "deleteAfterUse", 1))  ; 1 = delete each batch
deleteDelayMs  := Integer(IniRead(ControlIni, "paths", "deleteDelayMs", 10000)) ; wait 10 sec by default
retentionDays  := Integer(IniRead(ControlIni, "paths", "captureRetentionDays", 0))

; --- Native Capture (GDI+) ----------------------------------------------------
; Persisted selection + limits (same INI as Control Panel "capture" section)
global Cap_Mode    := IniRead(ControlIni, "capture", "mode",  "region")         ; "region" | "window"
global Cap_RectStr := IniRead(ControlIni, "capture", "rect",  "")               ; "x,y,w,h"
global Cap_WinTit  := IniRead(ControlIni, "capture", "winTitle", "")            ; optional convenience
global Cap_MaxKB   := Integer(IniRead(ControlIni, "capture", "maxKB", 1400))    ; default 1400 KB

; Parsed rect cache
global Cap_Rect := Map("x", 0, "y", 0, "w", 0, "h", 0)
if (Cap_RectStr != "") {
    parts := StrSplit(Cap_RectStr, ",")
    if (parts.Length = 4) {
        Cap_Rect["x"] := Integer(parts[1])
        Cap_Rect["y"] := Integer(parts[2])
        Cap_Rect["w"] := Integer(parts[3])
        Cap_Rect["h"] := Integer(parts[4])
    }
}

; Runtime state
global __GDI_Ready := false
global __Sel_Active := false
global __Sel_Gui := 0
global __Sel_Kind := ""        ; "region" | "window"

InitGDIPlus() {
    global __GDI_Ready
    if (__GDI_Ready)
        return
    static token := 0
    si := Buffer(A_PtrSize=8 ? 24 : 16, 0)
    NumPut("UInt", 1, si, 0) ; GdiplusVersion = 1 is fine for PNG
    if (DllCall("gdiplus\GdiplusStartup", "Ptr*", &token, "Ptr", si.Ptr, "Ptr", 0) = 0) {
        __GDI_Ready := true
    } else {
        __GDI_Ready := false
        Dbg("GDI+ startup failed")
    }
}

; Return a CLSID Buffer for the requested encoder without enumerating encoders.
; We only need PNG; add JPEG if you ever need it.
GetEncoderClsid(mime, &clsid) {
    clsid := Buffer(16, 0)
    ; PNG: {557CF406-1A04-11D3-9A73-0000F81EF32E}
    if (StrLower(mime) = "image/png") {
        if (DllCall("ole32\CLSIDFromString", "WStr", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid.Ptr) = 0)
            return true
        return false
    }
    ; Optional: JPEG encoder if ever needed
    ; if (StrLower(mime) = "image/jpeg") {
    ;     if (DllCall("ole32\CLSIDFromString", "WStr", "{557CF401-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid.Ptr) = 0)
    ;         return true
    ;     return false
    ; }

    return false
}

; --- Capturing ---------------------------------------------------------------

; Capture screen rect to HBITMAP
CaptureRectToBitmap(x, y, w, h) {
    hDC := DllCall("GetDC", "Ptr", 0, "Ptr")
    mDC := DllCall("CreateCompatibleDC", "Ptr", hDC, "Ptr")
    hbmp := DllCall("CreateCompatibleBitmap", "Ptr", hDC, "Int", w, "Int", h, "Ptr")
    obmp := DllCall("SelectObject", "Ptr", mDC, "Ptr", hbmp, "Ptr")
    DllCall("BitBlt", "Ptr", mDC, "Int", 0, "Int", 0, "Int", w, "Int", h, "Ptr", hDC, "Int", x, "Int", y, "UInt", 0x00CC0020)
    DllCall("SelectObject", "Ptr", mDC, "Ptr", obmp, "Ptr")
    DllCall("DeleteDC", "Ptr", mDC)
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hDC)
    return hbmp
}

; Capture a window (client area if PrintWindow fails, else full)
CaptureWindowToBitmap(hwnd) {
    ; Try PrintWindow (renders off-screen content for most windows)
    WinGetPos &wx, &wy, &ww, &wh, "ahk_id " hwnd
    if (ww <= 0 || wh <= 0)
        return 0
    hdcWin := DllCall("GetWindowDC", "Ptr", hwnd, "Ptr")
    mDC := DllCall("CreateCompatibleDC", "Ptr", hdcWin, "Ptr")
    hbmp := DllCall("CreateCompatibleBitmap", "Ptr", hdcWin, "Int", ww, "Int", wh, "Ptr")
    obmp := DllCall("SelectObject", "Ptr", mDC, "Ptr", hbmp, "Ptr")
    ok := DllCall("User32\PrintWindow", "Ptr", hwnd, "Ptr", mDC, "UInt", 0x00000002) ; PW_RENDERFULLCONTENT
    if (!ok) {
        ; fallback BitBlt from window DC
        DllCall("BitBlt", "Ptr", mDC, "Int", 0, "Int", 0, "Int", ww, "Int", wh, "Ptr", hdcWin, "Int", 0, "Int", 0, "UInt", 0x00CC0020)
    }
    DllCall("SelectObject", "Ptr", mDC, "Ptr", obmp, "Ptr")
    DllCall("DeleteDC", "Ptr", mDC)
    DllCall("ReleaseDC", "Ptr", hwnd, "Ptr", hdcWin)
    return hbmp
}

; Save HBITMAP → PNG file, optionally scaled down proportionally
; Auto reduces scale until size ≤ targetKB (with a floor so we don't go absurdly tiny)
SaveBitmapPngUnderKB(hbmp, fullW, fullH, outPath, targetKB := 1400) {
    InitGDIPlus()
    if (!__GDI_Ready)
        return false
    static pngClsid := 0
    if (!IsObject(pngClsid)) {
        if !GetEncoderClsid("image/png", &enc)
            return false
        pngClsid := enc
    }

    scale := 1.0
    minSide := 240

    Loop 10 {
        ; Create a GDI+ image from HBITMAP
        ; If scale < 1, draw into a scaled GDI+ bitmap
        DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hbmp, "Ptr", 0, "Ptr*", &src := 0)
		        ; --- ensure we scale the full image and never crop ---
        ; Some windows (e.g., Firefox fullscreen video) make PrintWindow/BitBlt
        ; produce a bitmap whose size differs from WinGetPos ww/wh.
        ; Use the actual bitmap dimensions for the source rect.
        DllCall("gdiplus\GdipGetImageWidth",  "Ptr", src, "UInt*", &realW := 0)
        DllCall("gdiplus\GdipGetImageHeight", "Ptr", src, "UInt*", &realH := 0)
        if (realW && realH) {
            fullW := realW
            fullH := realH
        }
        if (scale < 0.999) {
            dstW := Max( Round(fullW * scale), 1 )
            dstH := Max( Round(fullH * scale), 1 )
            DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", src, "Ptr*", &g := 0)
            ; make a new bitmap at target size
            DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", dstW, "Int", dstH, "Int", 0, "Int", 0x26200A, "Ptr", 0, "Ptr*", &dst := 0)
            DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", dst, "Ptr*", &g2 := 0)
            DllCall("gdiplus\GdipSetInterpolationMode", "Ptr", g2, "Int", 7)  ; HighQualityBicubic
            DllCall("gdiplus\GdipDrawImageRectRectI", "Ptr", g2, "Ptr", src, "Int", 0, "Int", 0, "Int", dstW, "Int", dstH, "Int", 0, "Int", 0, "Int", fullW, "Int", fullH, "Int", 2, "Ptr", 0, "Ptr", 0, "Ptr", 0)
            DllCall("gdiplus\GdipDeleteGraphics", "Ptr", g2)
            DllCall("gdiplus\GdipDisposeImage", "Ptr", src)
            src := dst, fullW := dstW, fullH := dstH
        }

        tmp := outPath ".tmp"
        DllCall("gdiplus\GdipSaveImageToFile", "Ptr", src, "WStr", tmp, "Ptr", pngClsid.Ptr, "Ptr", 0)
        DllCall("gdiplus\GdipDisposeImage", "Ptr", src)

        sz := FileExist(tmp) ? FileOpen(tmp, "r").Length : -1
        if (sz > 0 && sz <= targetKB * 1024) {
            try FileMove(tmp, outPath, 1)
            return true
        }
        try FileDelete(tmp)

        ; reduce scale for next loop
        if (fullW <= minSide || fullH <= minSide)
            break
        ; heuristic: shrink by ratio vs target, clamped to 0.6–0.9 per step
        factor := Sqrt( (targetKB*1024.0) / Max(sz, 1) )
        factor := Max(0.6, Min(factor, 0.9))
        scale := scale * factor
    }
    return false
}

; --- Selection UIs ------------------------------------------------------------

; --- Cancel any in-flight region/window pick: unhook, stop timers, hide bands ---
CancelAnyPick() {
    global __Sel_Active, __SelDragging, __SelTimerRunning
    global __Sel_Gui, __SelBand, __HoverBorderHwnd

    ; Stop timers (rubber-band + window hover)
    try SetTimer(DrawBand, 0)
    try SetTimer(PulseHover, 0)
    __SelTimerRunning := false

    ; Unregister both modes’ mouse hooks
    ;   0x0201 = WM_LBUTTONDOWN, 0x0202 = WM_LBUTTONUP
    try OnMessage(0x0201, Region_LButtonDown, 0)
    try OnMessage(0x0202, Region_LButtonUp,   0)
    try OnMessage(0x0201, PickWindowClick,    0)

    ; Hide/remove any visual pick UI that might be left
    if IsSet(__HoverBorderHwnd) && __HoverBorderHwnd {
        try WinHide "ahk_id " __HoverBorderHwnd
        __HoverBorderHwnd := 0
    }
    if IsSet(__SelBand) && __SelBand {
        try __SelBand.Destroy()
        __SelBand := 0
    }
    if IsSet(__Sel_Gui) && __Sel_Gui {
        try __Sel_Gui.Destroy()
        __Sel_Gui := 0
    }

    ; Reset flags
    __Sel_Active := false
    __SelDragging := false
}

; Hide Translator/Explainer while user is picking, then restore
global __HidSelf := false
global __HidOther := false

BeginPickHide() {
    global Overlay, __EXPLAIN_MODE, __HidSelf, __HidOther
    __HidSelf := false, __HidOther := false
    ; hide this overlay window
    try (Overlay.Hide(), __HidSelf := true)

    ; hide the other overlay (Explainer if we're Translator, Translator if we're Explainer)
    other := __EXPLAIN_MODE ? "Translator" : "Explainer"
    old := A_TitleMatchMode
    SetTitleMatchMode 3
    if WinExist(other) {
        try WinHide other
    __HidOther := true
    }
    SetTitleMatchMode old

    ; fail-safe: restore in 15s even if something goes wrong
    SetTimer(EndPickHide, -15000)
}

EndPickHide(*) {
    global Overlay, __EXPLAIN_MODE, __HidSelf, __HidOther
    if (__HidSelf) {
        try Overlay.Show()
        __HidSelf := false
    }
    if (__HidOther) {
        other := __EXPLAIN_MODE ? "Translator" : "Explainer"
        old := A_TitleMatchMode
        SetTitleMatchMode 3
        try WinShow other
        SetTitleMatchMode old
        __HidOther := false
    }
}

; Region rubber-band selection (drag, release to confirm)
StartPickRegion(maxKB := "", *) {
    global __Sel_Active, __Sel_Gui, __Sel_Kind, Cap_MaxKB
    global __SelBand, __SelTimerRunning
    global __Sel_sx, __Sel_sy, __Sel_ex, __Sel_ey, __SelDragging

    ; NEW: ensure no stale window/region hooks or timers are active
    CancelAnyPick()

    __Sel_Kind := "region", __Sel_Active := true

    ; Make sure window-click handler is not lingering, then register region handlers
    OnMessage(0x0201, PickWindowClick, 0)
    OnMessage(0x0201, Region_LButtonDown) ; WM_LBUTTONDOWN
    OnMessage(0x0202, Region_LButtonUp)   ; WM_LBUTTONUP

    BeginPickHide()   ; <<< hide Translator + Explainer while picking
    if (IsNumber(maxKB))
        Cap_MaxKB := Integer(maxKB)

    ; use screen pixel coords and the full virtual desktop
    CoordMode "Mouse", "Screen"
    vsx := SysGet(76)  ; SM_XVIRTUALSCREEN
    vsy := SysGet(77)  ; SM_YVIRTUALSCREEN
    vsw := SysGet(78)  ; SM_CXVIRTUALSCREEN
    vsh := SysGet(79)  ; SM_CYVIRTUALSCREEN

    ; full-sheet GUI (NOT DPI-scaled) so sizes are 1:1 with pixels
    g := Gui("-Caption +AlwaysOnTop +ToolWindow -DPIScale")
    g.BackColor := "000000"
    g.Opt("+LastFound")
    WinSetTransparent 8
    g.Show("x" vsx " y" vsy " w" vsw " h" vsh)
    __Sel_Gui := g

    ; thin border we move as the mouse drags (also NOT DPI-scaled)
    __SelBand := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    __SelBand.BackColor := "00FFFF"
    WinSetTransparent 40, __SelBand.Hwnd

    __SelDragging := false
    __Sel_sx := __Sel_sy := __Sel_ex := __Sel_ey := 0

    OnMessage(0x0201, Region_LButtonDown) ; WM_LBUTTONDOWN
    OnMessage(0x0202, Region_LButtonUp)   ; WM_LBUTTONUP

    __SelTimerRunning := true
    SetTimer(DrawBand, 15)

    ToolTip("Drag to select region… release to confirm"), SetTimer(() => ToolTip(""), -2000)
}

Region_LButtonDown(wParam := 0, lParam := 0, msg := 0, hwnd := 0) {
    global __Sel_Active, __SelDragging, __Sel_sx, __Sel_sy
    if (!__Sel_Active)
        return 0
    CoordMode "Mouse", "Screen"
    __SelDragging := true
    MouseGetPos &__Sel_sx, &__Sel_sy
    return 0
}

Region_LButtonUp(wParam := 0, lParam := 0, msg := 0, hwnd := 0) {
    global __Sel_Active, __SelDragging, __Sel_sx, __Sel_sy, __Sel_ex, __Sel_ey
    if (!__Sel_Active)
        return 0
    CoordMode "Mouse", "Screen"
    __SelDragging := false
    MouseGetPos &__Sel_ex, &__Sel_ey
    DestroySelOverlay()
    FinishRegionPick(__Sel_sx, __Sel_sy, __Sel_ex, __Sel_ey)
    return 0
}

DrawBand(*) {
    global __Sel_Active, __SelDragging, __SelBand, __Sel_sx, __Sel_sy
    if (!__Sel_Active || !__SelDragging)
        return
    CoordMode "Mouse", "Screen"
    MouseGetPos &mx, &my
    x := Min(__Sel_sx, mx), y := Min(__Sel_sy, my)
    w := Max(Abs(mx - __Sel_sx), 2), h := Max(Abs(my - __Sel_sy), 2)
    __SelBand.Show("NA x" x " y" y " w" w " h" h)
}

DestroySelOverlay() {
    global __Sel_Gui, __SelBand, __Sel_Active, __SelTimerRunning
    ; stop the draw timer
    if (__SelTimerRunning) {
        SetTimer(DrawBand, 0)
        __SelTimerRunning := false
    }
    ; don’t unregister OnMessage in v2 – just mark inactive
    __Sel_Active := false

    ; destroy band + screen GUIs
    try {
        if (IsObject(__SelBand))
            __SelBand.Destroy()
    }
    try {
        if (IsObject(__Sel_Gui))
            __Sel_Gui.Destroy()
    }
    __SelBand := 0, __Sel_Gui := 0
}

FinishRegionPick(sx, sy, ex, ey) {
    global Cap_Mode, Cap_Rect, Cap_RectStr
    x := Min(sx, ex), y := Min(sy, ey), w := Abs(ex - sx), h := Abs(ey - sy)
    if (w < 4 || h < 4) {
        ToolTip("Selection too small"), SetTimer(() => ToolTip(""), -800)
        return
    }
    Cap_Mode := "region"
    Cap_Rect["x"] := x, Cap_Rect["y"] := y, Cap_Rect["w"] := w, Cap_Rect["h"] := h
    Cap_RectStr := x "," y "," w "," h
    IniWrite(Cap_Mode,    ControlIni, "capture", "mode")
    IniWrite(Cap_RectStr, ControlIni, "capture", "rect")
    ToolTip("Region selected"), SetTimer(() => ToolTip(""), -700)
	EndPickHide()  ; <<< restore overlays after region is set
}

; Window hover highlight → click to select
StartPickWindow(maxKB := "", *) {
    global Cap_MaxKB

    ; Make sure region handlers are not lingering, then register window handler
    OnMessage(0x0201, Region_LButtonDown, 0)
    OnMessage(0x0202, Region_LButtonUp,   0)
    OnMessage(0x0201, PickWindowClick) ; left button down confirms

    BeginPickHide()   ; <<< hide Translator + Explainer while picking
    if (IsNumber(maxKB))
        Cap_MaxKB := Integer(maxKB)

    SetTimer(PulseHover, 30)
    ToolTip("Hover window and click to select…"), SetTimer(() => ToolTip(""), -1800)
}

PulseHover() {
    MouseGetPos ,, &hwnd
    if (!hwnd)
        return
    WinGetPos &x, &y, &w, &h, "ahk_id " hwnd
    ; draw a thin layered window border
    static b := 0
global __HoverBorderHwnd
if (!IsObject(b)) {
    b := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    b.BackColor := "00FFFF"
    WinSetTransparent 60, b.Hwnd
}
b.Show("NA x" x " y" y " w" w " h" h)
__HoverBorderHwnd := b.Hwnd
}

PickWindowClick(*) {
    ; Unregister the same callback we registered in StartPickWindow()
    OnMessage(0x0201, PickWindowClick, 0)
    SetTimer(PulseHover, 0)

    ; Hide the hover border if it's still shown
    global __HoverBorderHwnd
    if IsSet(__HoverBorderHwnd) && __HoverBorderHwnd {
        try WinHide "ahk_id " __HoverBorderHwnd
    }

    MouseGetPos ,, &hwnd
    if (!hwnd) {
        ToolTip("No window"), SetTimer(() => ToolTip(""), -700)
        return
    }
    tit := WinGetTitle("ahk_id " hwnd)
    global Cap_Mode, Cap_WinTit
    Cap_Mode := "window", Cap_WinTit := tit
    IniWrite(Cap_Mode,   ControlIni, "capture", "mode")
    IniWrite(Cap_WinTit, ControlIni, "capture", "winTitle")
    ToolTip("Window selected: " tit), SetTimer(() => ToolTip(""), -900)
	EndPickHide()  ; <<< restore overlays after window is set
}

; Produce a screenshot file from current selection (persisted)
CaptureOnceToFile(&outPath) {
    global Cap_Mode, Cap_Rect, Cap_MaxKB, captureDir
    InitGDIPlus()
    if (!__GDI_Ready) {
        showText("GDI+ not available.")
        return false
    }
    cd := ResolvePath(captureDir)
    if !DirExist(cd)
        DirCreate(cd)

    ts := FormatTime("", "yyyyMMdd-HHmmss")
    outPath := cd "\" ts "_cap.png"

    if (Cap_Mode = "region") {
        x := Cap_Rect["x"], y := Cap_Rect["y"], w := Cap_Rect["w"], h := Cap_Rect["h"]
        if (w <= 0 || h <= 0) {
            ToolTip("No region set. Use picker."), SetTimer(() => ToolTip(""), -1200)
            return false
        }
        hb := CaptureRectToBitmap(x, y, w, h)
        ok := SaveBitmapPngUnderKB(hb, w, h, outPath, Cap_MaxKB)
        DllCall("DeleteObject", "Ptr", hb)
        return ok
    } else if (Cap_Mode = "window") {
        MouseGetPos ,, &hw := 0
        ; if we have a title, try to find it; otherwise current under cursor
        if (Cap_WinTit != "") {
            old := A_TitleMatchMode
            SetTitleMatchMode 3
            if WinExist(Cap_WinTit)
                hw := WinExist(Cap_WinTit)
            SetTitleMatchMode old
        }
        if (!hw) {
            ToolTip("Window not found. Hover + pick again."), SetTimer(() => ToolTip(""), -1400)
            return false
        }
        WinGetPos , , &ww, &wh, "ahk_id " hw
        hb := CaptureWindowToBitmap(hw)
        if (!hb) {
            ToolTip("Capture failed"), SetTimer(() => ToolTip(""), -900)
            return false
        }
        ok := SaveBitmapPngUnderKB(hb, ww, wh, outPath, Cap_MaxKB)
        DllCall("DeleteObject", "Ptr", hb)
        return ok
    } else {
        ToolTip("No capture mode set"), SetTimer(() => ToolTip(""), -1000)
        return false
    }
}


sharexRegionHotkey := "^!F1"
sharexDefineHotkey := "^!F2"

ShotBuf := []
global IsTop := Integer(LoadCfg("winTop", !__EXPLAIN_MODE ? 1 : 0))
global hBrushEdit := 0

global OverlayDir := A_Temp "\JRPG_Overlay"
global AudioTxt     := OverlayDir . "\audio.txt"
global OcrTxt       := OverlayDir . "\ocr.txt"
global ExplainerTxt := OverlayDir . "\explainer.txt"
global PauseFlag    := OverlayDir "\audio.pause"
global __OcrText := ""
global __AudioText := ""
global __LastAudioRaw := ""
global __LastOcrRaw   := ""
global __LastExplainRaw := ""

MinW := 200, MinH := 120

; Ensure the bounds INI exists and has defaults on first run
EnsureBoundsIni(){
    global BoundsIni
    sentinel := "__MISSING__"
    if (IniRead(BoundsIni, "win", "x", sentinel) = sentinel) {
        ; First run — seed defaults
        IniWrite(120, BoundsIni, "win", "x")
        IniWrite(120, BoundsIni, "win", "y")
        IniWrite(900, BoundsIni, "win", "w")
        IniWrite(500, BoundsIni, "win", "h")
        ; If we can resolve DPI later, that's fine; 96 is a safe seed.
        IniWrite(96,  BoundsIni, "win", "dpi")
    }
}

LoadOverlayBounds(){
    global BoundsIni, MinW, MinH, Overlay
    x := IniRead(BoundsIni, "win", "x", 120)
    y := IniRead(BoundsIni, "win", "y", 120)
    w := IniRead(BoundsIni, "win", "w", 900)
    h := IniRead(BoundsIni, "win", "h", 500)
    savedDPI := IniRead(BoundsIni, "win", "dpi", 96)

    Dbg("LoadBounds ini: x=" x " y=" y " w=" w " h=" h " savedDPI=" savedDPI " (no DPI scaling; using DIP units)")
    ; Values in overlay.ini are stored in DIP (device-independent pixels); no scaling needed.
    w := Max(w, MinW), h := Max(h, MinH)

    local b := Map("x", x, "y", y, "w", w, "h", h)
    b := EnsureBoundsOnScreen(b)
    DbgRect("LoadBounds after clamp", b["x"], b["y"], b["w"], b["h"])
    return b
}
SaveOverlayBounds(){
    global Overlay, BoundsIni
        try {
        ; Save outer X/Y but CLIENT W/H (Show("w h") expects client size)
        Overlay.GetPos(&x, &y, , ,)          ; outer position
        Overlay.GetClientPos(,, &cw, &ch)    ; client width/height
        IniWrite(x,  BoundsIni, "win", "x")
        IniWrite(y,  BoundsIni, "win", "y")
        IniWrite(cw, BoundsIni, "win", "w")
        IniWrite(ch, BoundsIni, "win", "h")
        dpi := GetWindowDPI(Overlay.Hwnd)
        IniWrite(dpi, BoundsIni, "win", "dpi")
        Dbg("SaveBounds x=" x " y=" y " w(client)=" cw " h(client)=" ch " dpi=" dpi)
    }
}

; Ensure our process uses per-monitor DPI so coordinates match across displays
try DllCall("Shcore\SetProcessDpiAwareness", "Int", 2)  ; PROCESS_PER_MONITOR_DPI_AWARE
catch {
    try DllCall("User32\SetProcessDPIAware") ; legacy fallback
}

global Overlay, RectOuter, RectInner, RectPanel, OutputCtl, CtxMenu
Overlay := Gui("-Caption +Resize +MinSize" . MinW . "x" . MinH . " +0x02000000", __WIN_TITLE)  ; +WS_CLIPCHILDREN
Overlay.opt(IsTop ? "+AlwaysOnTop" : "-AlwaysOnTop")
Overlay.OnEvent("Close", HandleGuiClose.Bind(SCRIPT_PID))
Overlay.OnEvent("Escape", (*) => ExitApp())
Overlay.OnEvent("Size",   OnResize)

OnMessage(0x4A,  WM_COPYDATA)     ; WM_COPYDATA
OnMessage(0x0201, WM_LBUTTONDOWN) ; left down for dragging
OnMessage(0x0133, PaintEdit)  ; WM_CTLCOLOREDIT
OnMessage(0x0138, PaintStatic) ; WM_CTLCOLORSTATIC for RectOuter/RectInner/RectPanel
OnMessage(0x0014, EraseAnyBg)  ; WM_ERASEBKGND (Overlay/Panel/RichEdit/Inner/Outer)
OnMessage(0x0231, EnterSizeMove) ; WM_ENTERSIZEMOVE
OnMessage(0x0232, ExitSizeMove)  ; WM_EXITSIZEMOVE
OnMessage(0x8031, AfterSizeMoveRepaint) ; WM_APP+0x31 - deferred repaint after drag
OnMessage(0x0138, PaintEdit)      ; WM_CTLCOLORSTATIC (modern)
OnMessage(0x0232, SaveOnMoveSizeEnd) ; WM_EXITSIZEMOVE
OnMessage(0x020A, HandleMouseWheel)  ; WM_MOUSEWHEEL → forward to OutputCtl

RectOuter := Overlay.AddText("x0 y0 w0 h0")
RectInner := Overlay.AddText("x0 y0 w0 h0")
RectPanel := Overlay.AddText("x0 y0 w0 h0")
; NEW: use RichEdit (needs Msftedit.dll)
DllCall("LoadLibrary", "str", "Msftedit.dll", "ptr")
; Add styles as numeric flags at creation:
; +0x0004  ES_MULTILINE
; +0x0040  ES_AUTOVSCROLL
; +0x0800  ES_READONLY
; +0x200000 WS_VSCROLL
OutputCtl := Overlay.Add(
    "Custom",
    "ClassRICHEDIT50W x0 y0 w0 h0 +Theme -E0x200 -Border +0x0004 +0x0040 +0x0800"
)

; Background (avoid white)
EM_SETBKGNDCOLOR := 0x0443
SendMessage(EM_SETBKGNDCOLOR, 0, HexToBGR(BOX_BG), OutputCtl.Hwnd)

; Allow large text (256k chars)
EM_EXLIMITTEXT := 0x0435
SendMessage(EM_EXLIMITTEXT, 0, 262144, OutputCtl.Hwnd)

; Hide scrollbars (keep wheel/keyboard scrolling)
DllCall("user32\ShowScrollBar", "ptr", OutputCtl.Hwnd, "int", 1, "int", 0) ; SB_VERT = 1
DllCall("user32\ShowScrollBar", "ptr", OutputCtl.Hwnd, "int", 0, "int", 0) ; SB_HORZ = 0

GWL_STYLE := -16
WS_VSCROLL := 0x00200000, ES_DISABLENOSCROLL := 0x00002000
style := DllCall("user32\GetWindowLongPtr", "ptr", OutputCtl.Hwnd, "int", GWL_STYLE, "ptr")
DllCall("user32\SetWindowLongPtr", "ptr", OutputCtl.Hwnd, "int", GWL_STYLE, "ptr", style & ~(WS_VSCROLL|ES_DISABLENOSCROLL))
DllCall("user32\SetWindowPos", "ptr", OutputCtl.Hwnd, "ptr", 0, "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x27) ; SWP_NOSIZE|NOMOVE|NOZORDER|FRAMECHANGED

if !IsSet(gWheelHook) || !gWheelHook {
    OnMessage(0x020A, WheelToEdit) ; WM_MOUSEWHEEL
	OnMessage(0x024E, WheelToEdit) ; WM_POINTERWHEEL
    gWheelHook := true
}

CtxMenu := Menu()
CtxMenu.Add("Copy text", CtxCopyText)
CtxMenu.Add("Clear",     CtxClear)
CtxMenu.Add()
CtxMenu.Add("Always on top (toggle" . (__EXPLAIN_MODE ? " — Ctrl+Shift+X" : " — Ctrl+Shift+H") . ")", ToggleTop)
CtxMenu.Add("Reset window", ResetWindow)
CtxMenu.Add("Exit", (*) => ExitApp())
Overlay.OnEvent("ContextMenu", ShowContextMenu)

ApplyTheme()
EnsureBoundsIni()
wnd := LoadOverlayBounds()
Overlay.Show("NA x" wnd["x"] " y" wnd["y"] " w" wnd["w"] " h" wnd["h"])
HideOSRim(Overlay.Hwnd)
DbgRect("Shown", wnd["x"], wnd["y"], wnd["w"], wnd["h"])

ico := A_ScriptDir "\icon.ico"
try TraySetIcon(ico)
SetTaskbarIcon(Overlay.Hwnd, ico)

if !DirExist(OverlayDir)
    DirCreate(OverlayDir)

; start polling based on mode, clearing the relevant files first
if (__EXPLAIN_MODE) {
    ClearExplainerFile()
    SetTimer(() => PollExplainerFile(), 100)
    PollExplainerFile()
} else {
    ; In Translator mode, clear both OCR and Audio files on start
    ClearOcrFile()
    ClearAudioFile()
    SetTimer(() => PollOcrFile(), 100)
    SetTimer(() => PollAudioSubtitle(), 100)
    PollOcrFile()
    PollAudioSubtitle()
}
; watch for Control Panel “apply” signal
SetTimer(CheckCmdSignals, 120)
RenderCombined()
OnResize(Overlay)


WM_COPYDATA(wParam, lParam, msg, hwnd) {
    global Overlay
    global BOX_BG, BDR_OUT, BDR_IN, TXT_COLOR
    global FONT_NAME, FONT_SIZE, BOX_PAD, OUTER_W, INNER_W

    Critical
    settingsStr := StrGet(NumGet(lParam + A_PtrSize*2, "UPtr"))
    Dbg("WM_COPYDATA raw='" SubStr(settingsStr, 1, 240) (StrLen(settingsStr)>240 ? "...":"") "'")
; --- Debounce theme packets right after drag ends (safe version) ---
    if (IsSet(raw) && InStr(raw, "trans=") && IsSet(RESIZE_END_TICK) && (A_TickCount - RESIZE_END_TICK) < 250) {
        Dbg("WM_COPYDATA: suppressed trans packet due to recent size/move")
        return 1
    }

    local x := -1, y := -1, w := -1, h := -1
    srcDPI := 0

    for pair in StrSplit(settingsStr, "|") {
    kv := StrSplit(pair, "=")
    if (kv.Length != 2)
        continue
    key := kv[1], val := kv[2]

        if (key = "trans") {
            if IsSet(Overlay) && IsObject(Overlay)
                WinSetTransparent(val, "ahk_id " Overlay.Hwnd)
		        } else if (key = "capcmd" && val = "pick") {
            __cap_pick := 1
        } else if (key = "kind") {
            __cap_kind := val   ; "region" | "window"
        } else if (key = "maxkb") {
            __cap_maxkb := Integer(val)
        } else if (key = "bg") {
            BOX_BG := val
        } else if (key = "b_out") {
            BDR_OUT := val
        } else if (key = "b_in") {
            BDR_IN := val
        } else if (key = "txt") {
            TXT_COLOR := val
		} else if (key = "name") {
            global NAME_COLOR
            NAME_COLOR := val
        } else if (key = "font") {
            FONT_NAME := val
        } else if (key = "size") {
            FONT_SIZE := Integer(val)
        } else if (key = "pad") {
            BOX_PAD := Integer(val)
        } else if (key = "w_out" || key = "outw") {
            OUTER_W := Integer(val)
        } else if (key = "w_in"  || key = "inw") {
            INNER_W := Integer(val)
			; after both widths could have been set, this is a convenient place
            Dbg("Applied border widths OUTER_W=" OUTER_W " INNER_W=" INNER_W)
			} else if (key = "action" && val = "save_bounds") {
            Dbg("Received save_bounds command.")
            SetTimer(SaveOverlayBounds, -1) ; Save immediately after this message is processed
        } else if (key = "x") {
            x := Integer(val)
        } else if (key = "y") {
            y := Integer(val)
        } else if (key = "w") {
            w := Integer(val)
        } else if (key = "h") {
            h := Integer(val)
        } else if (key = "dpi") {
            srcDPI := Integer(val)
        }
    }

    ApplyTheme()

    if (x != -1 && y != -1 && w != -1 && h != -1) {
        Dbg("WM_COPYDATA parsed x=" x " y=" y " w=" w " h=" h " srcDPI=" srcDPI)
        if (srcDPI > 0) {
            cur := GetWindowDPI(Overlay.Hwnd)
            scale := cur / Max(srcDPI, 1)
            x := Round(x * scale), y := Round(y * scale)
            w := Max(Round(w * scale), 100), h := Max(Round(h * scale), 100)
            Dbg("WM_COPYDATA scaled-> x=" x " y=" y " w=" w " h=" h " curDPI=" cur " scale=" scale)
        }
        local bounds := Map("x", x, "y", y, "w", w, "h", h)
        local safeBounds := EnsureBoundsOnScreen(bounds)
        DbgRect("WM_COPYDATA clamp", safeBounds["x"], safeBounds["y"], safeBounds["w"], safeBounds["h"])
        Overlay.Move(safeBounds["x"], safeBounds["y"], safeBounds["w"], safeBounds["h"])
        SetTimer(() => SaveOverlayBounds(), -75)
    }
	
	    ; Handle capture picker command after parsing all pairs
    if (IsSet(__cap_pick) && __cap_pick = 1) {
        if (__cap_kind = "region")
            StartPickRegion(IsSet(__cap_maxkb) ? __cap_maxkb : "")
        else
            StartPickWindow(IsSet(__cap_maxkb) ? __cap_maxkb : "")
        return
    }


    ; Re-apply Win10 safety after any incoming changes
    NormalizeBordersForWin10()

    OnResize(Overlay)
    OnResize(Overlay)
    SetTimer(SaveOverlayBounds, -1)
    HideOSRim(Overlay.Hwnd)
    return true
}

ApplyTheme() {
    global Overlay, RectOuter, RectInner, RectPanel, OutputCtl, hBrushEdit
    global BOX_BG, BDR_OUT, BDR_IN, TXT_COLOR, FONT_NAME, FONT_SIZE, BOX_PAD

    Overlay.BackColor := BDR_OUT
    RectOuter.Opt("Background" . BDR_OUT)
    RectInner.Opt("Background" . BDR_IN)
    RectPanel.Opt("Background" . BOX_BG)
    ; Example shape – keep your function name/signature
    OnCtlColorEdit(wParam, lParam, *) {
       hdc := wParam
        ; TXT_COLOR/BOX_BG are your globals
        DllCall("SetTextColor", "ptr", hdc, "int", ((TXT_COLOR & 0xFF) << 16) | (TXT_COLOR & 0xFF00) | ((TXT_COLOR >> 16) & 0xFF))
        DllCall("SetBkColor",  "ptr", hdc, "int", ((BOX_BG   & 0xFF) << 16) | (BOX_BG   & 0xFF00) | ((BOX_BG   >> 16) & 0xFF))
        return hBrushEdit
    }


    Overlay.SetFont("s" . FONT_SIZE . " c" . TXT_COLOR, FONT_NAME)
    OutputCtl.SetFont("s" . FONT_SIZE . " c" . TXT_COLOR, FONT_NAME)

    OutputCtl.BackColor := BOX_BG
    DllCall("UxTheme\SetWindowTheme", "ptr", OutputCtl.Hwnd, "str", " ", "str", " ")

    ; Recreate the cached brush for CTL-COLOR painting
    if (hBrushEdit) {
        DllCall("gdi32\DeleteObject", "ptr", hBrushEdit)
        hBrushEdit := 0
    }
    hBrushEdit := MakeBrush(BOX_BG)

    ; Ensure the RichEdit also updates its internal background color
    try SetRichEditBg(OutputCtl, BOX_BG)

    ; Margins
    EC_LEFTMARGIN := 1, EC_RIGHTMARGIN := 2
    SendMessage(0x00D3, EC_LEFTMARGIN|EC_RIGHTMARGIN, (Integer(BOX_PAD)<<16)|Integer(BOX_PAD), OutputCtl.Hwnd)

    ; Force an immediate repaint (erase + update now)
    try DllCall("User32\RedrawWindow", "ptr", OutputCtl.Hwnd, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0004 | 0x0100) ; RDW_INVALIDATE|RDW_ERASE|RDW_UPDATENOW
		    ; ensure outer/inner/panel and edit update immediately
    RefreshAllBg()
}

RefreshAllBg() {
    global Overlay, RectOuter, RectInner, RectPanel, OutputCtl
    global BOX_BG, BDR_IN, BDR_OUT
    global hBrushOuter, hBrushInner, hBrushPanel, hBrushEdit

; ensure brush globals exist
if !IsSet(hBrushOuter) hBrushOuter := 0
if !IsSet(hBrushInner) hBrushInner := 0
if !IsSet(hBrushPanel) hBrushPanel := 0
if !IsSet(hBrushEdit)  hBrushEdit  := 0

; re-create brushes so they match the newest colors
if (hBrushOuter) {
    DllCall("gdi32\DeleteObject", "ptr", hBrushOuter)
    hBrushOuter := 0
}
if (hBrushInner) {
    DllCall("gdi32\DeleteObject", "ptr", hBrushInner)
    hBrushInner := 0
}
if (hBrushPanel) {
    DllCall("gdi32\DeleteObject", "ptr", hBrushPanel)
    hBrushPanel := 0
}
if (hBrushEdit) {
    DllCall("gdi32\DeleteObject", "ptr", hBrushEdit)
    hBrushEdit := 0
}


hBrushOuter := MakeBrush(ToHex6(BDR_OUT))
hBrushInner := MakeBrush(ToHex6(BDR_IN))
hBrushPanel := MakeBrush(ToHex6(BOX_BG))
hBrushEdit  := MakeBrush(ToHex6(BOX_BG))


    ; set GUI control backgrounds (keeps your existing logic consistent)
    RectOuter.Opt("Background" . BDR_OUT)
    RectInner.Opt("Background" . BDR_IN)
    RectPanel.Opt("Background" . BOX_BG)

    ; re-apply RichEdit BG
    SetRichEditBg(OutputCtl, BOX_BG)

    ; blast redraws so everything takes effect instantly
    for hwnd in [RectOuter.Hwnd, RectInner.Hwnd, RectPanel.Hwnd, OutputCtl.Hwnd, Overlay.Hwnd] {
        DllCall("User32\RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0
            , "uint", 0x0001 | 0x0004 | 0x0200 | 0x0100) ; INVALIDATE|ERASE|ALLCHILDREN|UPDATENOW
    }
}

WM_LBUTTONDOWN(wParam, lParam, msg, hwnd) {
    global Overlay
    if (hwnd = Overlay.Hwnd)
        PostMessage(0xA1, 2, -1, , "ahk_id " hwnd)
}

PaintEdit(wParam, lParam, *) {
    global OutputCtl, hBrushEdit, BOX_BG, TXT_COLOR
    if (lParam = OutputCtl.Hwnd) {
        if (!hBrushEdit)
            hBrushEdit := MakeBrush(BOX_BG)
        DllCall("gdi32\SetTextColor", "ptr", wParam, "int", HexToBGR(TXT_COLOR))
        DllCall("gdi32\SetBkColor",  "ptr", wParam, "int", HexToBGR(BOX_BG))
        Dbg("PaintEdit hit: BG=" BOX_BG " brush=" hBrushEdit)
        return hBrushEdit
    }
}

; Map WM_CTLCOLORSTATIC for our three rectangles so they recolor instantly.
PaintStatic(wParam, lParam, *) {
    global RectOuter, RectInner, RectPanel
    global BDR_OUT, BDR_IN, BOX_BG
    global hBrushOuter, hBrushInner, hBrushPanel

if (lParam = RectOuter.Hwnd) {
    if (!IsSet(hBrushOuter) || !hBrushOuter)
        hBrushOuter := MakeBrush(ToHex6(BDR_OUT))
    DllCall("gdi32\SetBkColor", "ptr", wParam, "int", ColorToBGR(BDR_OUT))
    return hBrushOuter
}

if (lParam = RectInner.Hwnd) {
    if (!IsSet(hBrushInner) || !hBrushInner)
        hBrushInner := MakeBrush(ToHex6(BDR_IN))
    DllCall("gdi32\SetBkColor", "ptr", wParam, "int", ColorToBGR(BDR_IN))
    return hBrushInner
}

if (lParam = RectPanel.Hwnd) {
    if (!IsSet(hBrushPanel) || !hBrushPanel)
        hBrushPanel := MakeBrush(ToHex6(BOX_BG))
    DllCall("gdi32\SetBkColor", "ptr", wParam, "int", ColorToBGR(BOX_BG))
    return hBrushPanel
}

}

; Force background erase for the RichEdit so it never flashes an old/system color.
; Fill background for Overlay, RectPanel and RichEdit with BOX_BG so no system color leaks in.
EraseAnyBg(wParam, lParam, msg, hwnd) {
    global Overlay, RectPanel, RectInner, RectOuter, OutputCtl
    global BOX_BG, BDR_IN, BDR_OUT
    global hBrushEdit, hBrushPanel, hBrushInner, hBrushOuter

    if (hwnd != Overlay.Hwnd && hwnd != RectPanel.Hwnd && hwnd != RectInner.Hwnd
        && hwnd != RectOuter.Hwnd && hwnd != OutputCtl.Hwnd)
        return

    ; choose brush by window
    brush := 0
if (hwnd = RectPanel.Hwnd) {
    if (!IsSet(hBrushPanel) || !hBrushPanel) hBrushPanel := MakeBrush(ToHex6(BOX_BG))
    brush := hBrushPanel
} else if (hwnd = RectInner.Hwnd) {
    if (!IsSet(hBrushInner) || !hBrushInner) hBrushInner := MakeBrush(ToHex6(BDR_IN))
    brush := hBrushInner
} else if (hwnd = RectOuter.Hwnd) {
    if (!IsSet(hBrushOuter) || !hBrushOuter) hBrushOuter := MakeBrush(ToHex6(BDR_OUT))
    brush := hBrushOuter
} else if (hwnd = OutputCtl.Hwnd) {
    if (!IsSet(hBrushEdit)  || !hBrushEdit)  hBrushEdit  := MakeBrush(ToHex6(BOX_BG))
    brush := hBrushEdit
}
    if (!brush)
        return

    rc := Buffer(16, 0)
    DllCall("User32\GetClientRect", "ptr", hwnd, "ptr", rc)
    DllCall("User32\FillRect", "ptr", wParam, "ptr", rc, "ptr", brush)
    Dbg("EraseAnyBg hwnd=" hwnd " filled")
    return 1
}

EnterSizeMove(wParam, lParam, msg, hwnd) {
    global Overlay, OutputCtl, RESIZE_LOCK
    if (hwnd != Overlay.Hwnd)
        return
    RESIZE_LOCK := true
    ; suspend RichEdit painting while the user is dragging the window
    SendMessage(0x000B, 0, 0, OutputCtl.Hwnd) ; WM_SETREDRAW(FALSE)
    Dbg("EnterSizeMove: redraw suspended")
    return 0
}

ExitSizeMove(wParam, lParam, msg, hwnd) {
    global Overlay, RESIZE_LOCK, RESIZE_END_TICK
    if (hwnd != Overlay.Hwnd)
        return
    RESIZE_LOCK := false
    RESIZE_END_TICK := A_TickCount    ; remember when the drag ended

    ; Defer the repaint to avoid racing the system’s final paint
    DllCall("User32\PostMessage", "ptr", Overlay.Hwnd, "uint", 0x8031, "ptr", 0, "ptr", 0) ; WM_APP+0x31
    Dbg("ExitSizeMove: posted deferred repaint")
    return 0
}

AfterSizeMoveRepaint(wParam, lParam, msg, hwnd) {
    global Overlay, OutputCtl, BOX_BG, hBrushEdit
    if (hwnd != Overlay.Hwnd)
        return

    ; Re-enable redraw for the RichEdit now that OS is done with its own pass
    SendMessage(0x000B, 1, 0, OutputCtl.Hwnd) ; WM_SETREDRAW(TRUE)

    if (!hBrushEdit)
        hBrushEdit := MakeBrush(BOX_BG)

    ; Re-apply the background (EM_SETBKGNDCOLOR inside SetRichEditBg)
    SetRichEditBg(OutputCtl, BOX_BG)

    ; Force complete erase + paint now, including children.
    DllCall("User32\RedrawWindow", "ptr", OutputCtl.Hwnd, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0004 | 0x0200 | 0x0100) ; INVALIDATE|ERASE|ALLCHILDREN|UPDATENOW

    ; One more nudge to be extra-safe on some themes:
    DllCall("User32\UpdateWindow", "ptr", OutputCtl.Hwnd)

    Dbg("AfterSizeMoveRepaint: BG re-applied=" BOX_BG " brush=" hBrushEdit)
    return 0
}

SaveOnMoveSizeEnd(wParam, lParam, msg, hwnd) {
    global Overlay
    if (hwnd = Overlay.Hwnd)
        SaveOverlayBounds()
}

showText(txt) {
    global __OcrText, __AudioText
    norm := StrReplace(txt, "`r`n", "`n")
    norm := StrReplace(norm, "`r", "`n")
    norm := StrReplace(norm, "`n", "`r`n")
    __OcrText := norm
    __AudioText := ""   ; screenshots clear audio text
    RenderCombined()
}
RenderCombined() {
    global OutputCtl, __OcrText, __AudioText
    global TXT_COLOR

    combined := ""
    if (Trim(__OcrText) != "")
        combined := __OcrText
    if (Trim(__AudioText) != "") {
        if (combined != "")
            combined .= "`r`n"
        combined .= __AudioText
    }

    ; Parse ⟦i⟧...⟦/i⟧ (italic) and ⟦name⟧...⟦/name⟧ (speaker color) markers
    plainText := ""
    spans := ParseInlineMarkers(combined) ; fills global __LastParsedPlain and returns an Array of spans

    ; Set plain text first
plainText := __LastParsedPlain
; Put text in the control
SendMessage(0x000C, 0, StrPtr(plainText), OutputCtl.Hwnd) ; WM_SETTEXT

; Reset formatting on ALL existing text: Select-All then SCF_SELECTION
EM_SETSEL        := 0x00B1
EM_SETCHARFORMAT := 0x0444
SCF_SELECTION    := 0x0001
CFM_ITALIC       := 0x00000002
CFM_COLOR        := 0x40000000

; Select-all
SendMessage(EM_SETSEL, 0, -1, OutputCtl.Hwnd)

cf := Buffer(116, 0)               ; CHARFORMAT2W
NumPut("UInt", 116, cf, 0)                             ; cbSize
NumPut("UInt", CFM_COLOR|CFM_ITALIC, cf, 4)            ; dwMask
NumPut("UInt", 0, cf, 8)                               ; dwEffects (no italic)
NumPut("UInt", HexToBGR(TXT_COLOR), cf, 20)            ; crTextColor
SendMessage(EM_SETCHARFORMAT, SCF_SELECTION, cf.Ptr, OutputCtl.Hwnd)

; Collapse selection to start
SendMessage(EM_SETSEL, 0, 0, OutputCtl.Hwnd)

    ; Apply span styles (fix CRLF→CR index mismatch in RichEdit)
    AdjustPos(idx) {
        ; RichEdit counts CRLF as one position; our spans count '\r' + '\n'
        ; Subtract the number of LF characters before idx.
        local prefix := SubStr(plainText, 1, idx)
        local lfCnt  := StrLen(prefix) - StrLen(StrReplace(prefix, "`n", ""))
        return idx - lfCnt
    }

for _, s in spans {
    start := AdjustPos(s["startPos"])    ; 0-based (RichEdit-compatible)
    end   := AdjustPos(s["endPos"])      ; 0-based exclusive
    if (end < start)
        continue

    SendMessage(EM_SETSEL, start, end, OutputCtl.Hwnd)

    mask     := 0
    effects  := 0
    colorRef := 0

    if (s["isItalic"]) {
        mask    |= CFM_ITALIC
        effects |= 0x00000002      ; CFE_ITALIC
    }
    if (s["isNameSpan"]) {
        mask    |= CFM_COLOR
        colorRef := s["nameColorRef"]
    }

    cf2 := Buffer(116, 0)
    NumPut("UInt", 116, cf2, 0)                 ; cbSize
    NumPut("UInt", mask, cf2, 4)                ; dwMask
    NumPut("UInt", effects, cf2, 8)             ; dwEffects
    if (mask & CFM_COLOR)
        NumPut("UInt", colorRef, cf2, 20)       ; crTextColor
    SendMessage(EM_SETCHARFORMAT, SCF_SELECTION, cf2.Ptr, OutputCtl.Hwnd)
}

    ; Clear selection and scroll caret into view
    SendMessage(EM_SETSEL, 0, 0, OutputCtl.Hwnd)
    SendMessage(0x00CE, 0, 0, OutputCtl.Hwnd) ; EM_SCROLLCARET
}

; --- NEW: inline marker parser for RichEdit styling ---
; Builds plain text and a list of spans with (startPos, endPos, isItalic, isNameSpan, nameColorRef)
ParseInlineMarkers(s) {
    global __LastParsedPlain
    global TXT_COLOR
    ; Name color priority: live override from WM_COPYDATA → env → fallback to TXT_COLOR
    global NAME_COLOR
    nameHex := ""
    if (IsSet(NAME_COLOR) && StrLen(Trim(NAME_COLOR)))
        nameHex := NAME_COLOR
    else {
        tmp := EnvGet("NAME_COLOR")
        if (tmp && StrLen(Trim(tmp)))
            nameHex := tmp
    }
    if (!nameHex || StrLen(Trim(nameHex)) = 0)
        nameHex := TXT_COLOR
    nameClr := HexToBGR(nameHex)

    ; helper: normalize any incoming text to CRLF *before* appending
    NormalizeNL(txt) {
        txt := StrReplace(txt, "`r`n", "`n")
        txt := StrReplace(txt, "`r", "`n")
        return StrReplace(txt, "`n", "`r`n")
    }

    out := ""
    spans := []
    pos := 1
    len := StrLen(s)
    while (pos <= len) {
        p1 := InStr(s, "⟦", false, pos)
        if (!p1) {
            out .= NormalizeNL(SubStr(s, pos))
            break
        }
        out .= NormalizeNL(SubStr(s, pos, p1 - pos))
        p2 := InStr(s, "⟧", false, p1 + 1)
        if (!p2) {
            out .= NormalizeNL(SubStr(s, p1))
            break
        }
        tag := SubStr(s, p1 + 1, p2 - p1 - 1) ; e.g., i, /i, name, /name
        if (tag = "i" || tag = "name") {
            close := "⟦/" . tag . "⟧"
            p3 := InStr(s, close, false, p2 + 1)
            if (!p3) {
                out .= NormalizeNL(SubStr(s, p1, p2 - p1 + 1))
                pos := p2 + 1
                continue
            }
            inner := NormalizeNL(SubStr(s, p2 + 1, p3 - (p2 + 1)))
            startPos := StrLen(out)        ; 0-based (before we append inner)
            out .= inner
            endPos := StrLen(out)          ; 0-based exclusive
            span := Map()
            span["startPos"] := startPos
            span["endPos"]   := endPos
            span["isItalic"] := (tag = "i")
            span["isNameSpan"] := (tag = "name")
            span["nameColorRef"] := nameClr
            spans.Push(span)
            pos := p3 + StrLen(close)
        } else {
            ; stray/unknown tag → keep literal
            out .= NormalizeNL(SubStr(s, p1, p2 - p1 + 1))
            pos := p2 + 1
        }
    }

    ; no post-loop normalization necessary—positions already match 'out'
    __LastParsedPlain := out
    return spans
}

HandleMouseWheel(wParam, lParam, msg, hwnd) {
    global Overlay, OutputCtl
    ; Only act if the wheel event hit the overlay window itself
    if (!IsSet(Overlay) || hwnd != Overlay.Hwnd)
        return
    ; Focus the edit and forward the original wheel message to it
    DllCall("user32\SetFocus", "ptr", OutputCtl.Hwnd)
    PostMessage(0x020A, wParam, lParam, OutputCtl.Hwnd)
}

; Remember the last real foreground window so we can jump back to it.
global __LastActiveHwnd := 0

ToggleTop(*) {
    global Overlay, __LastActiveHwnd
    ; Make sure nothing here clashes with globals:
    local hwnd, isTopNow, pt, x, y, h

    ; Resolve hwnd safely (Overlay may not be ready yet)
    ; --- resolve hwnd safely (Overlay may not exist yet or be hidden) ---
hwnd := 0
try hwnd := Overlay.Hwnd  ; Overlay might not be set yet

oldDHW := A_DetectHiddenWindows
DetectHiddenWindows true

if (!hwnd || !WinExist("ahk_id " hwnd)) {
    ; fallback: find by exact title
    prevTM := A_TitleMatchMode
    SetTitleMatchMode(3)  ; exact
    hwnd := WinExist("Explainer")
    SetTitleMatchMode(prevTM)

    if !hwnd {
        DetectHiddenWindows(oldDHW)
        return
    }
}

DetectHiddenWindows oldDHW
; --- end safe hwnd resolution ---

isTopNow := IsWindowTopmost(hwnd)


        if (!isTopNow) {
        ; about to turn ON: remember who was active before us
        __LastActiveHwnd := WinGetID("A")

        if (hwnd && WinExist("ahk_id " hwnd)) {
            DllCall("SetWindowPos", "ptr", hwnd, "ptr", -1    ; HWND_TOPMOST
                , "int", 0, "int", 0, "int", 0, "int", 0
                , "uint", 0x0013)                              ; NOMOVE|NOSIZE|NOACTIVATE
            try WinActivate("ahk_id " hwnd)
        } else {
            return
        }
    } else {
        ; turn OFF: drop NOTOPMOST, then shove to bottom of normal band
        if (hwnd && WinExist("ahk_id " hwnd)) {
            DllCall("SetWindowPos", "ptr", hwnd, "ptr", -2     ; HWND_NOTOPMOST
                , "int", 0, "int", 0, "int", 0, "int", 0
                , "uint", 0x0013)
            DllCall("SetWindowPos", "ptr", hwnd, "ptr", 1      ; HWND_BOTTOM
                , "int", 0, "int", 0, "int", 0, "int", 0
                , "uint", 0x0013)
        } else {
            return
        }

        ; give focus back to what you were using (RetroArch, etc.)
        if (__LastActiveHwnd && WinExist("ahk_id " __LastActiveHwnd)) {
            WinActivate("ahk_id " __LastActiveHwnd)
        } else {
            ; fallback: activate window under mouse
            pt := Buffer(8, 0), DllCall("GetCursorPos", "ptr", pt)
            x := NumGet(pt, 0, "int"), y := NumGet(pt, 4, "int")
            h := DllCall("WindowFromPoint", "int64", (y<<32)|x, "ptr")
            if (h)
                WinActivate("ahk_id " h)
        }
    }
}

ShowContextMenu(gui, ctrl, *) {
    global CtxMenu
    MouseGetPos &mx, &my
    CtxMenu.Show(mx, my)
}

ResetWindow(*) {
    global Overlay
    Overlay.Move(120, 120, 900, 500)
    SaveOverlayBounds()
    OnResize(Overlay)
}

; --- Helpers: smooth wheel / touchpad scroll for OutputCtl without visible scrollbar ---
WheelToEdit(wParam, lParam, msg, hwnd) {
    global OutputCtl
    if (!IsSet(OutputCtl) || !OutputCtl)
        return

    ; Allow wheel only if the mouse is over the output edit
    ; OR if that edit currently has keyboard focus.
    MouseGetPos ,, &winHwnd, &ctrlHwnd, 2
    hasFocus := (OutputCtl.Gui && OutputCtl.Gui.FocusedCtrl == OutputCtl)
    if (ctrlHwnd != OutputCtl.Hwnd && !hasFocus)
        return

    ; ----- scroll amount from wheel message -----
    delta := (wParam >> 16) & 0xFFFF
    if (delta >= 0x8000)
        delta -= 0x10000  ; sign-extend

    ; system "lines per notch" (usually 3)
    static SPI_GETWHEELSCROLLLINES := 0x0068
    linesPerNotch := 3
    DllCall("user32\SystemParametersInfo", "uint", SPI_GETWHEELSCROLLLINES
        , "uint", 0, "uint*", &linesPerNotch, "uint", 0)

    ; accumulate for smooth scrolling
    static remainder := 0
    remainder += delta
    WHEEL_DELTA := 120
    notches := remainder // WHEEL_DELTA
    remainder -= notches * WHEEL_DELTA

    lines := -notches * linesPerNotch  ; EM_LINESCROLL: negative is up

    ; send EM_LINESCROLL (0x00B6) to the edit control
    DllCall("user32\SendMessage", "ptr", OutputCtl.Hwnd
        , "uint", 0x00B6, "ptr", 0, "ptr", lines)

    return 0  ; consume
}

; === OnResize layout ===
OnResize(guiObj, minmax := "", w := 0, h := 0) {
    global OUTER_W, INNER_W, BOX_PAD
    global RectOuter, RectInner, RectPanel, OutputCtl
    global BOX_BG, hBrushEdit, RESIZE_LOCK

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

    ToolTip("Translating " n " shot(s)…")
    SetTimer(() => ToolTip(""), -900)
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

    ToolTip("Capture → translate…"), SetTimer(() => ToolTip(""), -700)

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
    m["append_to_buffer"]   := IniRead(ControlIni, "hotkeys", "append_to_buffer",   "^1")
    m["flush_translation"]  := IniRead(ControlIni, "hotkeys", "flush_translation",  "^2")
    m["redefine_region"]    := IniRead(ControlIni, "hotkeys", "redefine_region",    "^3")
	
    ; Prefer new key name; fall back to legacy entry if it exists
    tmpHk := IniRead(ControlIni, "hotkeys", "screenshot_translate", "")
    if (tmpHk = "")
        tmpHk := IniRead(ControlIni, "hotkeys", "oneshot_translate", "")
    m["oneshot_translate"]  := tmpHk
    
	; Show/Hide rows replace legacy ^0 (translator) / ^+0 (explainer)
    if (__EXPLAIN_MODE) {
         m["hide_show_explainer"] := IniRead(ControlIni, "hotkeys", "hide_show_explainer", "^+0")
     } else {
         m["hide_show_translator"] := IniRead(ControlIni, "hotkeys", "hide_show_translator", "^0")
     }
 
     ; Topmost is now configurable separately (no default here to avoid collision)
     m["toggle_top"]         := IniRead(ControlIni, "hotkeys", "toggle_top", "")
     m["toggle_audio"]       := IniRead(ControlIni, "hotkeys", "toggle_audio", "")
	 m["recapture_region"]   := IniRead(ControlIni, "hotkeys", "recapture_region", "")
	 m["take_screenshot"]    := IniRead(ControlIni, "hotkeys", "take_screenshot", "")
	 m["screenshot_translation"] := IniRead(ControlIni, "hotkeys", "screenshot_translation", "")
	
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
    global __CMD_TOGGLE_EXPL, __EXPLAIN_MODE
    ; Only Explainer reacts to this command
    if (!__EXPLAIN_MODE)
        return
    if FileExist(__CMD_TOGGLE_EXPL) {
        try FileDelete(__CMD_TOGGLE_EXPL)
        try ToggleTop()   ; do the safe, internal toggle
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
    RenderCombined()
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
    RenderCombined()
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
}

; ======== Context menu ========
CtxCopyText(*) {
    global OutputCtl
    ; Ensure the RichEdit has focus so WM_COPY targets the right selection
    try OutputCtl.Focus()
    ; WM_COPY (0x0301) — copies the current selection (same as Ctrl+C)
    SendMessage(0x0301, 0, 0, OutputCtl.Hwnd)
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
    SaveOverlayBounds()
    if (hBrushEdit)
        DllCall("gdi32\DeleteObject", "ptr", hBrushEdit)
}