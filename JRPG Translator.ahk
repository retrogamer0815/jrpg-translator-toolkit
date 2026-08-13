#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn
#NoTrayIcon
;@Ahk2Exe-SetVersion 0.9.5.0
;@Ahk2Exe-SetName JRPG Translator
;@Ahk2Exe-SetDescription JRPG Translator
;@Ahk2Exe-SetCopyright Copyright (c) 2025 retrogamer0815
; =; === Taskbar grouping: shared AppUserModelID ===
DllCall("shell32\SetCurrentProcessExplicitAppUserModelID", "wstr", "JRPGTranslator", "int")

; Big Box and other front ends can pass --background to initialize everything
; without showing or activating the control panel.
global CP_BACKGROUND_START := false
global CP_START_PROFILE := ""
global CP_START_TRANSLATOR := false
global CP_STUDY_START_MODE := ""
global CP_STUDY_ONLY_PROCESS := false
global APP_VERSION := "0.9.5-dev"
global PROJECT_URL := "https://github.com/retrogamer0815/jrpg-translator-toolkit"
global BUG_REPORT_URL := PROJECT_URL "/issues/new"
global WRITTEN_GUIDE_URL := PROJECT_URL "#quick-start"
global BEGINNER_VIDEO_URL := "https://youtu.be/pdZ0fBS8COc"
for __cpIndex, __cpArg in A_Args {
    __cpArgLower := StrLower(__cpArg)
    if (__cpArgLower = "--background") {
        CP_BACKGROUND_START := true
    } else if (__cpArgLower = "--open-translator") {
        CP_START_TRANSLATOR := true
    } else if (__cpArgLower = "--study-library") {
        CP_STUDY_START_MODE := "library"
        CP_STUDY_ONLY_PROCESS := true
        CP_BACKGROUND_START := true
    } else if (__cpArgLower = "--study-reader") {
        CP_STUDY_START_MODE := "reader"
        CP_STUDY_ONLY_PROCESS := true
        CP_BACKGROUND_START := true
    } else if (__cpArgLower = "--profile" && __cpIndex < A_Args.Length) {
        CP_START_PROFILE := A_Args[__cpIndex + 1]
    } else if RegExMatch(__cpArg, "i)^--profile=(.*)$", &__cpProfileMatch) {
        CP_START_PROFILE := __cpProfileMatch[1]
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
global CPControllerInputsEnabled := false
global CPControllerDpadNavigationEnabled := true
global CPControllerCaptureActive := false
global CPControllerPreviousTokens := Map()
global CPControllerBindings := Map()
global CPControllerBindingEdits := Map()
global CPControllerAssignButtons := Map()
global CPControllerDisableButtons := Map()
global CPControlsKeyboardControls := []
global CPControlsControllerControls := []
global CPControlsCurrentView := "keyboard"
global CPControllerLastDeviceName := ""
global CPControllerLastStatusText := ""
global CPControllerNavPreviousState := Map()
global CPControllerNavTargetHwnd := 0
global CPControllerNavHeldDirection := ""
global CPControllerNavNextRepeatAt := 0
global CPControllerNavHeldSince := 0
global CPControllerLastNativeNavigationAt := Map()
global CPKeyboardSliderRepeatState := Map("active", false)
global CPFontSizeAdjustState := Map("active", false)
global CPFontSizeAdjustSyncing := false
global CPMaxPngAdjustState := Map("active", false)
global CPMaxPngAdjustSyncing := false
global CPControllerColorDialogState := Map("active", false)
global CPWelcomeDialog := 0
global CPStudyLibraryWelcomeDialog := 0
global CPStudyLibraryState := 0
global CPStudyReaderState := 0
global CPStudyCandidateState := 0
global CPStudyCandidateOpening := false
global CPStudyBridgeRoot := A_Temp "\JRPG_Translator_Study_Bridge"
global CPStudyBridgeWorkspaces := Map()
global CPStudyBridgeSerial := 0

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
OnExit(StudyBridgeCleanupAll)

StudyBridgeCreateWorkspace(sbKind := "operation") {
    global CPStudyBridgeRoot, CPStudyBridgeWorkspaces, CPStudyBridgeSerial
    sbKind := RegExReplace(Trim(sbKind), "[^A-Za-z0-9_-]+", "_")
    if (sbKind = "")
        sbKind := "operation"
    DirCreate(CPStudyBridgeRoot)
    sbPid := DllCall("kernel32\GetCurrentProcessId", "uint")
    loop {
        CPStudyBridgeSerial += 1
        sbPath := CPStudyBridgeRoot "\" sbKind "_" sbPid "_"
            . A_TickCount "_" CPStudyBridgeSerial
        if !DirExist(sbPath) && !FileExist(sbPath)
            break
    }
    DirCreate(sbPath)
    CPStudyBridgeWorkspaces[StrLower(sbPath)] := sbPath
    return sbPath
}

StudyBridgeCleanupWorkspace(sbPath, *) {
    global CPStudyBridgeWorkspaces
    sbPath := Trim(sbPath)
    sbKey := StrLower(sbPath)
    ; Only paths created and registered by StudyBridgeCreateWorkspace can ever
    ; be removed here.  A bridge state therefore cannot accidentally turn a
    ; database, screenshot, or user-configured directory into a cleanup target.
    if (sbPath = "" || !CPStudyBridgeWorkspaces.Has(sbKey))
        return false
    CPStudyBridgeWorkspaces.Delete(sbKey)
    try DirDelete(sbPath, true)
    return !DirExist(sbPath)
}

StudyBridgeCleanupAll(*) {
    global CPStudyBridgeRoot, CPStudyBridgeWorkspaces
    if !IsSet(CPStudyBridgeWorkspaces) || !IsObject(CPStudyBridgeWorkspaces)
        return
    sbPaths := []
    for sbKey, sbPath in CPStudyBridgeWorkspaces
        sbPaths.Push(sbPath)
    for sbPath in sbPaths
        StudyBridgeCleanupWorkspace(sbPath)
    ; Leave no empty application-owned container after an orderly shutdown.
    try DirDelete(CPStudyBridgeRoot)
}

SendControlPanelCopyData(hwnd, payload) {
    if !hwnd
        return false
    cpPayloadBuffer := Buffer(StrPut(payload, "UTF-16") * 2, 0)
    StrPut(payload, cpPayloadBuffer, "UTF-16")
    cpCopyData := Buffer(A_PtrSize * 3, 0)
    NumPut("UPtr", 0, cpCopyData, 0)
    NumPut("UPtr", cpPayloadBuffer.Size, cpCopyData, A_PtrSize)
    NumPut("Ptr", cpPayloadBuffer.Ptr, cpCopyData, A_PtrSize * 2)
    return DllCall("user32\SendMessageW", "ptr", hwnd, "uint", 0x004A
        , "ptr", 0, "ptr", cpCopyData.Ptr, "ptr")
}

SendControlPanelCopyDataWithRetry(hwnd, payload, timeoutMs := 10000) {
    if !hwnd
        return false
    cpDeadline := A_TickCount + timeoutMs
    loop {
        if SendControlPanelCopyData(hwnd, payload)
            return true
        if (A_TickCount >= cpDeadline || !WinExist("ahk_id " hwnd))
            return false
        Sleep(60)
    }
}

if (__CP_ALREADY_RUNNING) {
    if (!CP_BACKGROUND_START || CP_START_PROFILE != ""
        || CP_STUDY_START_MODE != "") {
        __cpOldDhw := A_DetectHiddenWindows
        __cpOldTitleMode := A_TitleMatchMode
        try {
            DetectHiddenWindows true
            SetTitleMatchMode 3
            __cpExistingHwnd := WinExist("JRPG Translator")
            if (__cpExistingHwnd) {
                if (CP_START_PROFILE != "")
                    SendControlPanelCopyDataWithRetry(
                        __cpExistingHwnd, "apply_profile=" CP_START_PROFILE
                    )
                if (CP_STUDY_START_MODE = "library") {
                    SendControlPanelCopyDataWithRetry(
                        __cpExistingHwnd, "open_study_library"
                    )
                } else if (CP_STUDY_START_MODE = "reader") {
                    SendControlPanelCopyDataWithRetry(
                        __cpExistingHwnd, "open_study_reader"
                    )
                } else if (!CP_BACKGROUND_START) {
                    DllCall("user32\ShowWindow", "ptr", __cpExistingHwnd, "int", 5) ; SW_SHOW
                    try WinActivate("ahk_id " __cpExistingHwnd)
                }
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
; Keep the original folder as the non-migrating Default library. Additional
; libraries live beside it and are selected through the Study Library window.
studyLibraryDefaultDir := A_ScriptDir "\Settings\Study Library"
studyLibrariesRoot := A_ScriptDir "\Settings\Study Libraries"
studyLibrariesArchiveRoot := A_ScriptDir "\Settings\Study Libraries Archive"
studyLibraryDir := studyLibraryDefaultDir

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
    global cbDebug, debugMode, iniPath
    debugMode := cbDebug.Value ? 1 : 0
    SetDebugMode(debugMode)
    IniWrite(debugMode, iniPath, "cfg", "debugMode")
}

PromptPostprocMode(promptName := "") {
    global directModelOutput
    if directModelOutput
        return "none"

    normalized := StrLower(Trim(promptName))
    if InStr(normalized, "with_transcript") || InStr(normalized, "with_kanji_reading")
        return "tt"
    return "translation"
}

SyncPromptPostproc(promptName := "") {
    global promptProfile, imgPostproc
    if (promptName = "")
        promptName := promptProfile
    imgPostproc := PromptPostprocMode(promptName)
    return imgPostproc
}

CPOnDirectModelOutputToggle(*) {
    global cbDirectModelOutput, directModelOutput, promptProfile, imgPostproc
    directModelOutput := cbDirectModelOutput.Value ? 1 : 0
    imgPostproc := SyncPromptPostproc(promptProfile)
    SaveAll()
    ApplyShotSettings()
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
global CPThemedDialogHwnds := Map()
global CPThemedPopupButtons := Map()
global CPStudyThemedComboHwnds := Map()
global CPStudyThemedHeaderHwnds := Map()
global CPStudyThemedListHwnds := Map()
global CPStudyComboSubclassCallback := 0
global CPStudyHeaderSubclassCallback := 0
global CPStudyListSubclassCallback := 0
global CPStudyVisualOverlays := Map()
global CPStudyTransparentHwnds := Map()
global CPStudyTransparentCallback := 0
global CPThemeColorSwatchHwnds := Map()
global CPColorFocusFrame := []
global CPControllerColorGradientSliders := Map()
global CPControllerColorGradientMessageRegistered := false
global CPThemeMessagesRegistered := false
global CPComboArrowOverlays := []
global CPCapturePickerNavigation := Map()

; -------- Scrollable control-panel canvas --------
; The existing layout remains a fixed minimum design surface. When the window is
; smaller, native scrollbars expose that surface instead of compressing controls.
global CP_CANVAS_MIN_W := 890
global CP_CANVAS_MIN_H := 640
global CP_VIEWPORT_MIN_W := 640
global CP_VIEWPORT_MIN_H := 380
global CPCanvasScrollX := 0
global CPCanvasScrollY := 0
global CPCanvasScrollMaxX := 0
global CPCanvasScrollMaxY := 0
global CPCanvasViewportW := CP_CANVAS_MIN_W
global CPCanvasViewportH := CP_CANVAS_MIN_H
global CPCanvasMessagesRegistered := false
global CPCanvasFixedXHwnds := Map()
global CPCanvasFixedYHwnds := Map()
global CPCanvasClipStates := Map()
global CPCanvasSiblingClipHwnds := Map()
global CPCanvasPendingScrollX := 0
global CPCanvasPendingScrollY := 0
global CPCanvasPendingScrollValid := false
global CPCanvasScrollFlushScheduled := false
global CPPreferredViewportW := CP_CANVAS_MIN_W
global CPPreferredViewportH := CP_CANVAS_MIN_H
global CPWindowWidthSnapActive := false
global CPWindowHeightSnapActive := false
global CPWindowSnapRange := 14
global CPWindowSnapReleaseRange := 24
global CPPanelInteractiveResize := false
global CPPanelResizeOpacitySuspended := false
global CPPanelLastLayoutW := 0
global CPPanelLastLayoutH := 0
global CPTabScreenSnapshotHwnd := 0
global CPTabScreenSnapshotBitmap := 0

CPRegisterCanvasMessages() {
    global ui, CPCanvasMessagesRegistered
    if CPCanvasMessagesRegistered
        return
    OnMessage(0x0114, CPOnCanvasScroll) ; WM_HSCROLL
    OnMessage(0x0115, CPOnCanvasScroll) ; WM_VSCROLL
    OnMessage(0x020A, CPOnCanvasMouseWheel) ; WM_MOUSEWHEEL
    OnMessage(0x0231, CPOnWindowEnterSizeMove) ; WM_ENTERSIZEMOVE
    OnMessage(0x0214, CPOnWindowSizing) ; WM_SIZING
    OnMessage(0x0232, CPOnWindowExitSizeMove) ; WM_EXITSIZEMOVE
    ; Start with the non-client scrollbars hidden. SetScrollInfo will reveal
    ; either one only when the viewport is smaller than the design surface.
    try DllCall("user32\ShowScrollBar", "ptr", ui.Hwnd, "int", 3, "int", 0)
    CPCanvasMessagesRegistered := true
}

CPRegisterCanvasFixedControl(ctrl, fixedX := false, fixedY := true) {
    global CPCanvasFixedXHwnds, CPCanvasFixedYHwnds
    if !IsObject(ctrl) || !ctrl.Hwnd
        return
    if fixedX
        CPCanvasFixedXHwnds[ctrl.Hwnd] := true
    if fixedY
        CPCanvasFixedYHwnds[ctrl.Hwnd] := true
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

CPCanvasMoveChildrenDeferred(dx, dy) {
    global ui, CPCanvasFixedXHwnds, CPCanvasFixedYHwnds

    movableControls := []
    for ctrl in CPCanvasDirectControls() {
        ctrlDx := CPCanvasFixedXHwnds.Has(ctrl.Hwnd) ? 0 : dx
        ctrlDy := CPCanvasFixedYHwnds.Has(ctrl.Hwnd) ? 0 : dy
        if (ctrlDx || ctrlDy)
            movableControls.Push(Map("ctrl", ctrl, "dx", ctrlDx, "dy", ctrlDy))
    }
    if !movableControls.Length
        return true

    dpi := 96
    try dpi := Max(96, DllCall("user32\GetDpiForWindow", "ptr", ui.Hwnd, "uint"))
    scale := dpi / 96
    for entry in movableControls {
        ctrl := entry["ctrl"]
        rect := Buffer(16, 0)
        if !DllCall("user32\GetWindowRect", "ptr", ctrl.Hwnd, "ptr", rect.Ptr, "int")
            return false
        point := Buffer(8, 0)
        NumPut("int", NumGet(rect, 0, "int"), point, 0)
        NumPut("int", NumGet(rect, 4, "int"), point, 4)
        if !DllCall("user32\ScreenToClient", "ptr", ui.Hwnd, "ptr", point.Ptr, "int")
            return false

        entry["x"] := NumGet(point, 0, "int") + Round(entry["dx"] * scale)
        entry["y"] := NumGet(point, 4, "int") + Round(entry["dy"] * scale)
    }

    deferred := DllCall("user32\BeginDeferWindowPos", "int", movableControls.Length, "ptr")
    if !deferred
        return false

    static SWP_NOSIZE := 0x0001, SWP_NOZORDER := 0x0004, SWP_NOREDRAW := 0x0008
    static SWP_NOACTIVATE := 0x0010, SWP_NOOWNERZORDER := 0x0200
    moveFlags := SWP_NOSIZE | SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE | SWP_NOOWNERZORDER

    for entry in movableControls {
        ctrl := entry["ctrl"]
        deferred := DllCall("user32\DeferWindowPos", "ptr", deferred, "ptr", ctrl.Hwnd
            , "ptr", 0, "int", entry["x"], "int", entry["y"], "int", 0, "int", 0
            , "uint", moveFlags, "ptr")
        if !deferred
            return false
    }
    return DllCall("user32\EndDeferWindowPos", "ptr", deferred, "int") != 0
}

CPCanvasMoveChildren(dx, dy, redraw := true, manageRedraw := true, updateNow := true) {
    global ui, CPCanvasFixedXHwnds, CPCanvasFixedYHwnds
    if (!dx && !dy)
        return

    hwnd := ui.Hwnd
    ; WinSetTransparent makes the control panel a layered window. Suspending
    ; redraw for that entire top-level surface exposes intermediate compositor
    ; frames, so move its children as one deferred, non-redrawing batch instead.
    exStyle := hwnd ? DllCall("user32\GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr") : 0
    if (hwnd && manageRedraw && (exStyle & 0x00080000)
        && CPCanvasMoveChildrenDeferred(dx, dy)) {
        CPClipScrollableControlsToViewport(false)
        if redraw {
            redrawFlags := 0x0001 | 0x0004 | 0x0080 ; INVALIDATE|ERASE|ALLCHILDREN
            if updateNow
                redrawFlags |= 0x0100 ; UPDATENOW
            DllCall("user32\RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", redrawFlags)
        }
        return
    }

    if (hwnd && manageRedraw)
        DllCall("user32\SendMessage", "ptr", hwnd, "uint", 0x000B, "ptr", 0, "ptr", 0) ; WM_SETREDRAW
    try {
        for ctrl in CPCanvasDirectControls() {
            try {
                ctrlDx := CPCanvasFixedXHwnds.Has(ctrl.Hwnd) ? 0 : dx
                ctrlDy := CPCanvasFixedYHwnds.Has(ctrl.Hwnd) ? 0 : dy
                if (!ctrlDx && !ctrlDy)
                    continue
                ctrl.GetPos(&x, &y)
                ctrl.Move(x + ctrlDx, y + ctrlDy)
            }
        }
        CPClipScrollableControlsToViewport(false)
    } finally {
        if (hwnd && manageRedraw) {
            DllCall("user32\SendMessage", "ptr", hwnd, "uint", 0x000B, "ptr", 1, "ptr", 0)
            if redraw {
                redrawFlags := 0x0001 | 0x0004 | 0x0080 ; INVALIDATE|ERASE|ALLCHILDREN
                if updateNow
                    redrawFlags |= 0x0100 ; UPDATENOW
                DllCall("user32\RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", redrawFlags)
            }
        }
    }
}

CPClipScrollableControlsToViewport(redraw := false) {
    global ui, CPCanvasFixedYHwnds, CPCanvasClipStates, CPCanvasSiblingClipHwnds
    global CPTabBarFill, sepAction
    if !(IsSet(ui) && ui && ui.Hwnd
        && IsSet(CPTabBarFill) && CPTabBarFill
        && IsSet(sepAction) && sepAction)
        return

    headerRect := Buffer(16, 0)
    footerRect := Buffer(16, 0)
    if !DllCall("user32\GetWindowRect", "ptr", CPTabBarFill.Hwnd, "ptr", headerRect.Ptr, "int")
        return
    if !DllCall("user32\GetWindowRect", "ptr", sepAction.Hwnd, "ptr", footerRect.Ptr, "int")
        return

    clipTopPx := NumGet(headerRect, 12, "int")
    clipBottomPx := NumGet(footerRect, 4, "int")
    for scrollCtrl in CPCanvasDirectControls() {
        scrollHwnd := scrollCtrl.Hwnd
        if (CPCanvasFixedYHwnds.Has(scrollHwnd))
            continue

        ctrlRect := Buffer(16, 0)
        if !DllCall("user32\GetWindowRect", "ptr", scrollHwnd, "ptr", ctrlRect.Ptr, "int")
            continue
        ctrlLeftPx := NumGet(ctrlRect, 0, "int")
        ctrlTopPx := NumGet(ctrlRect, 4, "int")
        ctrlWidthPx := Max(0, NumGet(ctrlRect, 8, "int") - ctrlLeftPx)
        ctrlHeightPx := Max(0, NumGet(ctrlRect, 12, "int") - ctrlTopPx)
        visibleTopPx := Max(0, clipTopPx - ctrlTopPx)
        visibleBottomPx := Min(ctrlHeightPx, clipBottomPx - ctrlTopPx)
        if (visibleBottomPx < visibleTopPx)
            visibleBottomPx := visibleTopPx

        needsSiblingClip := visibleTopPx > 0 || visibleBottomPx < ctrlHeightPx
        scrollStyle := DllCall("user32\GetWindowLongPtr", "ptr", scrollHwnd, "int", -16, "ptr")
        if needsSiblingClip {
            ; Only controls crossing a sticky boundary need WS_CLIPSIBLINGS.
            ; Leaving it enabled on ordinary tab content lets hidden controls
            ; from other pages carve holes into visible dropdowns.
            if !(scrollStyle & 0x04000000) {
                DllCall("user32\SetWindowLongPtr", "ptr", scrollHwnd, "int", -16
                    , "ptr", scrollStyle | 0x04000000, "ptr")
                CPCanvasSiblingClipHwnds[scrollHwnd] := true
                DllCall("user32\SetWindowPos", "ptr", scrollHwnd, "ptr", 0
                    , "int", 0, "int", 0, "int", 0, "int", 0
                    , "uint", 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020)
            }
        } else if CPCanvasSiblingClipHwnds.Has(scrollHwnd) {
            if (scrollStyle & 0x04000000) {
                DllCall("user32\SetWindowLongPtr", "ptr", scrollHwnd, "int", -16
                    , "ptr", scrollStyle & ~0x04000000, "ptr")
                DllCall("user32\SetWindowPos", "ptr", scrollHwnd, "ptr", 0
                    , "int", 0, "int", 0, "int", 0, "int", 0
                    , "uint", 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020)
            }
            CPCanvasSiblingClipHwnds.Delete(scrollHwnd)
        }

        clipState := ctrlWidthPx "|" ctrlHeightPx "|" visibleTopPx "|" visibleBottomPx
        if (CPCanvasClipStates.Has(scrollHwnd) && CPCanvasClipStates[scrollHwnd] = clipState)
            continue

        if (visibleTopPx = 0 && visibleBottomPx = ctrlHeightPx) {
            DllCall("user32\SetWindowRgn", "ptr", scrollHwnd, "ptr", 0, "int", redraw ? 1 : 0)
        } else {
            clipRegion := DllCall("gdi32\CreateRectRgn", "int", 0, "int", visibleTopPx
                , "int", ctrlWidthPx, "int", visibleBottomPx, "ptr")
            if clipRegion {
                if !DllCall("user32\SetWindowRgn", "ptr", scrollHwnd, "ptr", clipRegion
                    , "int", redraw ? 1 : 0)
                    DllCall("gdi32\DeleteObject", "ptr", clipRegion)
            }
        }
        CPCanvasClipStates[scrollHwnd] := clipState
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

CPCanvasScrollTo(newX, newY, redraw := true, manageRedraw := true, updateNow := true) {
    global CPCanvasScrollX, CPCanvasScrollY, CPCanvasScrollMaxX, CPCanvasScrollMaxY
    global CPCanvasViewportW, CPCanvasViewportH, CP_CANVAS_MIN_W, CP_CANVAS_MIN_H

    newX := Max(0, Min(CPCanvasScrollMaxX, Round(newX)))
    newY := Max(0, Min(CPCanvasScrollMaxY, Round(newY)))
    dx := CPCanvasScrollX - newX
    dy := CPCanvasScrollY - newY
    CPCanvasSetScrollInfo(0, CP_CANVAS_MIN_W, CPCanvasViewportW, newX)
    CPCanvasSetScrollInfo(1, CP_CANVAS_MIN_H, CPCanvasViewportH, newY)
    CPCanvasMoveChildren(dx, dy, redraw, manageRedraw, updateNow)
    CPCanvasScrollX := newX
    CPCanvasScrollY := newY
    return (dx || dy)
}

CPCanvasQueueScrollTo(newX, newY) {
    global CPCanvasPendingScrollX, CPCanvasPendingScrollY, CPCanvasPendingScrollValid
    global CPCanvasScrollFlushScheduled

    CPCanvasPendingScrollX := newX
    CPCanvasPendingScrollY := newY
    CPCanvasPendingScrollValid := true
    if !CPCanvasScrollFlushScheduled {
        CPCanvasScrollFlushScheduled := true
        SetTimer(CPCanvasFlushQueuedScroll, -16)
    }
}

CPCanvasFlushQueuedScroll(*) {
    global CPCanvasPendingScrollX, CPCanvasPendingScrollY, CPCanvasPendingScrollValid
    global CPCanvasScrollFlushScheduled

    CPCanvasScrollFlushScheduled := false
    if !CPCanvasPendingScrollValid
        return
    newX := CPCanvasPendingScrollX
    newY := CPCanvasPendingScrollY
    CPCanvasPendingScrollValid := false
    ; Let Windows combine paints while the thumb is moving instead of erasing and
    ; repainting the complete control panel for every one-pixel track message.
    CPCanvasScrollTo(newX, newY, true, true, false)
}

CPCanvasCancelQueuedScroll(flush := false) {
    global CPCanvasPendingScrollX, CPCanvasPendingScrollY, CPCanvasPendingScrollValid
    global CPCanvasScrollFlushScheduled

    if CPCanvasScrollFlushScheduled
        SetTimer(CPCanvasFlushQueuedScroll, 0)
    CPCanvasScrollFlushScheduled := false
    if (flush && CPCanvasPendingScrollValid) {
        newX := CPCanvasPendingScrollX
        newY := CPCanvasPendingScrollY
        CPCanvasPendingScrollValid := false
        CPCanvasScrollTo(newX, newY)
        return
    }
    CPCanvasPendingScrollValid := false
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
    if !(IsSet(ui) && ui && hwnd = ui.Hwnd)
        return

    ; A standard window scrollbar reports lParam=0. Some themed/DPI layouts can
    ; route the same notification through a ScrollBar child instead. Accept that
    ; child, but continue ignoring trackbars and other controls which also emit
    ; WM_HSCROLL/WM_VSCROLL notifications.
    if lParam {
        scrollClass := ""
        try scrollClass := WinGetClass("ahk_id " lParam)
        if (scrollClass != "ScrollBar")
            return
    }

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
        case 8:                                      ; SB_ENDSCROLL
            CPCanvasCancelQueuedScroll(true)
            return 0
        default: return 0
    }

    if (code = 5) {
        if (bar = 0)
            CPCanvasQueueScrollTo(nextPos, CPCanvasScrollY)
        else
            CPCanvasQueueScrollTo(CPCanvasScrollX, nextPos)
        return 0
    }

    CPCanvasCancelQueuedScroll(false)
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

    if !CPMouseWheelOverCanvas()
        return

    delta := (wParam >> 16) & 0xFFFF
    if (delta & 0x8000)
        delta -= 0x10000
    if !delta
        return
    CPCanvasScrollTo(CPCanvasScrollX, CPCanvasScrollY - (delta / 120) * 48)
    return 0
}

CPMouseWheelOverCanvas(*) {
    global ui, CPCanvasScrollMaxY
    if !(IsSet(ui) && ui && ui.Hwnd && CPCanvasScrollMaxY > 0)
        return false
    if !WinActive("ahk_id " ui.Hwnd)
        return false

    mouseWindow := 0
    mouseControl := 0
    try MouseGetPos(,, &mouseWindow, &mouseControl, 2)
    if (!mouseWindow || (mouseWindow != ui.Hwnd
        && !DllCall("user32\IsChild", "ptr", ui.Hwnd, "ptr", mouseWindow, "int")))
        return false

    hoveredHwnd := mouseControl ? mouseControl : mouseWindow
    hoveredClass := ""
    try hoveredClass := WinGetClass("ahk_id " hoveredHwnd)
    if (hoveredClass = "msctls_trackbar32")
        return false
    if (hoveredHwnd && CPComboDropped(hoveredHwnd))
        return false

    focusedHwnd := CPFocusRingTargetHwnd(CPFocusedHwnd())
    if (focusedHwnd && CPComboDropped(focusedHwnd))
        return false
    return true
}

CPMouseWheelHotkey(direction, *) {
    global CPCanvasScrollX, CPCanvasScrollY
    CPCanvasScrollTo(CPCanvasScrollX, CPCanvasScrollY + direction * 48)
}

CPEnsureFocusedControlVisible(*) {
    global ui, CPCanvasScrollX, CPCanvasScrollY, CPCanvasScrollMaxX, CPCanvasScrollMaxY
    global CPCanvasFixedYHwnds, CPTabBarFill, sepAction
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
    if (IsSet(CPCanvasFixedYHwnds) && CPCanvasFixedYHwnds.Has(focusHwnd))
        return

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
    stickyRect := Buffer(16, 0)
    if (IsSet(CPTabBarFill) && CPTabBarFill && DllCall("user32\GetWindowRect", "ptr", CPTabBarFill.Hwnd, "ptr", stickyRect.Ptr, "int"))
        viewT := Max(viewT, NumGet(stickyRect, 12, "int") + 8)
    if (IsSet(sepAction) && sepAction && DllCall("user32\GetWindowRect", "ptr", sepAction.Hwnd, "ptr", stickyRect.Ptr, "int"))
        viewB := Min(viewB, NumGet(stickyRect, 4, "int") - 8)
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

CPTargetPreferredOuterWidthPx() {
    global ui, CPPreferredViewportW
    if !(IsSet(ui) && ui && ui.Hwnd && CPPreferredViewportW > 0)
        return 0

    windowRect := Buffer(16, 0)
    clientRect := Buffer(16, 0)
    if !DllCall("user32\GetWindowRect", "ptr", ui.Hwnd, "ptr", windowRect.Ptr, "int")
        return 0
    if !DllCall("user32\GetClientRect", "ptr", ui.Hwnd, "ptr", clientRect.Ptr, "int")
        return 0

    outerWidthPx := NumGet(windowRect, 8, "int") - NumGet(windowRect, 0, "int")
    clientWidthPx := NumGet(clientRect, 8, "int")
    ui.GetClientPos(,, &logicalClientW, &logicalClientH)
    logicalToPhysical := clientWidthPx / Max(1, logicalClientW)
    targetClientWidthPx := Round(CPPreferredViewportW * logicalToPhysical)
    return targetClientWidthPx + outerWidthPx - clientWidthPx
}

CPTargetPreferredOuterHeightPx() {
    global ui, CPPreferredViewportH
    if !(IsSet(ui) && ui && ui.Hwnd && CPPreferredViewportH > 0)
        return 0

    windowRect := Buffer(16, 0)
    clientRect := Buffer(16, 0)
    if !DllCall("user32\GetWindowRect", "ptr", ui.Hwnd, "ptr", windowRect.Ptr, "int")
        return 0
    if !DllCall("user32\GetClientRect", "ptr", ui.Hwnd, "ptr", clientRect.Ptr, "int")
        return 0

    outerHeightPx := NumGet(windowRect, 12, "int") - NumGet(windowRect, 4, "int")
    clientHeightPx := NumGet(clientRect, 12, "int")
    ui.GetClientPos(,, &logicalClientW, &logicalClientH)
    logicalToPhysical := clientHeightPx / Max(1, logicalClientH)
    targetClientHeightPx := Round(CPPreferredViewportH * logicalToPhysical)
    return targetClientHeightPx + outerHeightPx - clientHeightPx
}

CPWindowSnapShouldApply(distance, &active) {
    global CPWindowSnapRange, CPWindowSnapReleaseRange
    if !active {
        if (distance > CPWindowSnapRange)
            return false
        active := true
    } else if (distance > CPWindowSnapReleaseRange) {
        active := false
        return false
    }
    return true
}

CPOnWindowEnterSizeMove(wParam, lParam, msg, hwnd) {
    global ui, CPWindowWidthSnapActive, CPWindowHeightSnapActive
    global CPPanelInteractiveResize, CPPanelResizeOpacitySuspended
    global controlPanelOpacity
    if (IsSet(ui) && ui && hwnd = ui.Hwnd) {
        CPWindowWidthSnapActive := false
        CPWindowHeightSnapActive := false
        CPPanelInteractiveResize := true
        CPPanelResizeOpacitySuspended := false

        ; A top-level window using alpha transparency is a layered window.
        ; Windows recomposes every moved child HWND separately during a live
        ; resize, which produces unavoidable flashing below 100% opacity.
        ; Temporarily remove only that layered style for the drag; the exact
        ; configured opacity is restored immediately when resizing finishes.
        if (controlPanelOpacity < 100) {
            try {
                WinSetTransparent("Off", "ahk_id " hwnd)
                CPPanelResizeOpacitySuspended := true
                DllCall("user32\RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0
                    , "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100)
            }
        }
    }
}

CPOnWindowSizing(wParam, lParam, msg, hwnd) {
    global ui, CPWindowWidthSnapActive, CPWindowHeightSnapActive
    if !(IsSet(ui) && ui && hwnd = ui.Hwnd && lParam)
        return

    proposedLeft := NumGet(lParam, 0, "int")
    proposedTop := NumGet(lParam, 4, "int")
    proposedRight := NumGet(lParam, 8, "int")
    proposedBottom := NumGet(lParam, 12, "int")
    snapped := false

    ; Left/right edges and corners snap width while preserving the opposite edge.
    if (wParam = 1 || wParam = 2 || wParam = 4 || wParam = 5
        || wParam = 7 || wParam = 8) {
        targetOuterWidth := CPTargetPreferredOuterWidthPx()
        widthDistance := Abs(proposedRight - proposedLeft - targetOuterWidth)
        if (targetOuterWidth > 0 && CPWindowSnapShouldApply(widthDistance, &CPWindowWidthSnapActive)) {
            if (wParam = 1 || wParam = 4 || wParam = 7)
                NumPut("int", proposedRight - targetOuterWidth, lParam, 0)
            else
                NumPut("int", proposedLeft + targetOuterWidth, lParam, 8)
            snapped := true
        }
    }

    ; Top/bottom edges and corners snap height while preserving the opposite edge.
    if (wParam = 3 || wParam = 4 || wParam = 5
        || wParam = 6 || wParam = 7 || wParam = 8) {
        targetOuterHeight := CPTargetPreferredOuterHeightPx()
        heightDistance := Abs(proposedBottom - proposedTop - targetOuterHeight)
        if (targetOuterHeight > 0 && CPWindowSnapShouldApply(heightDistance, &CPWindowHeightSnapActive)) {
            if (wParam = 3 || wParam = 4 || wParam = 5)
                NumPut("int", proposedBottom - targetOuterHeight, lParam, 4)
            else
                NumPut("int", proposedTop + targetOuterHeight, lParam, 12)
            snapped := true
        }
    }

    if snapped
        return true
}

CPOnWindowExitSizeMove(wParam, lParam, msg, hwnd) {
    global ui, CPWindowWidthSnapActive, CPWindowHeightSnapActive
    global CPPanelInteractiveResize
    if (IsSet(ui) && ui && hwnd = ui.Hwnd) {
        CPWindowWidthSnapActive := false
        CPWindowHeightSnapActive := false
        CPPanelInteractiveResize := false
        ; Let the non-client resize loop finish before the one full cleanup
        ; repaint. Live WM_SIZE passes intentionally avoid synchronous erases.
        SetTimer(CPFinalizeInteractiveResize, -1)
    }
}

CPFinalizeInteractiveResize(*) {
    global ui, CPPanelResizeOpacitySuspended
    if !(IsSet(ui) && ui)
        return
    try hwnd := ui.Hwnd
    catch
        return
    if !hwnd
        return
    try DllCall("user32\RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100) ; INVALIDATE|ERASE|ALLCHILDREN|UPDATENOW
    if CPPanelResizeOpacitySuspended
        SetTimer(CPRestoreOpacityAfterResize, -25)
}

CPRestoreOpacityAfterResize(*) {
    global ui, CPPanelResizeOpacitySuspended
    if !CPPanelResizeOpacitySuspended
        return
    CPPanelResizeOpacitySuspended := false
    if !(IsSet(ui) && ui)
        return
    try hwnd := ui.Hwnd
    catch
        return
    if !hwnd
        return

    CPApplyControlPanelOpacity(false)
    ; Re-adding alpha transparency needs only one compositor update; do not
    ; erase the already complete child-control frame again.
    try DllCall("user32\RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0020 | 0x0080) ; INVALIDATE|NOERASE|ALLCHILDREN
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
    if IsSet(CPThemeBrushWindow) && CPThemeBrushWindow
        try DllCall("gdi32\DeleteObject", "ptr", CPThemeBrushWindow)
    if IsSet(CPThemeBrushSurface) && CPThemeBrushSurface
        try DllCall("gdi32\DeleteObject", "ptr", CPThemeBrushSurface)
    if IsSet(CPThemeBrushFocus) && CPThemeBrushFocus
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
        CPComboArrowOverlays.Push(Map("combo", cpComboCtrl, "arrow", cpArrowCtrl, "state", ""))
    }
    CPUpdateComboArrowOverlays()
    CPClipScrollableControlsToViewport(false)
    CPMaintainFooterZOrder()
}

CPUpdateComboArrowOverlays(force := false, *) {
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
            cpArrowNextState := "hidden"
            cpArrowCurrentState := cpArrowEntry.Has("state") ? cpArrowEntry["state"] : ""
            if (force || cpArrowCurrentState != cpArrowNextState) {
                cpArrowCtrl.Visible := false
                cpArrowEntry["state"] := cpArrowNextState
            }
            continue
        }

        cpArrowCombo.GetPos(&cpArrowX, &cpArrowY, &cpArrowW, &cpArrowH)
        cpArrowWidth := Min(26, Max(20, Floor(cpArrowH * 0.85)))
        cpArrowEnabled := DllCall("user32\IsWindowEnabled", "ptr", cpArrowCombo.Hwnd, "int")
        cpArrowNextState := "visible|" cpArrowX "|" cpArrowY "|" cpArrowW "|" cpArrowH "|" cpArrowEnabled
        cpArrowCurrentState := cpArrowEntry.Has("state") ? cpArrowEntry["state"] : ""
        if (!force && cpArrowCurrentState = cpArrowNextState)
            continue

        cpArrowCtrl.Move(cpArrowX + cpArrowW - cpArrowWidth, cpArrowY + 1, cpArrowWidth - 1, Max(1, cpArrowH - 2))
        cpArrowCtrl.Opt("+Background" cpArrowColors["surface"])
        cpArrowCtrl.SetFont("s9 c" (cpArrowEnabled ? cpArrowColors["text"] : cpArrowColors["muted"]))
        cpArrowCtrl.Visible := true
        try DllCall("user32\SetWindowPos", "ptr", cpArrowCtrl.Hwnd, "ptr", 0
             , "int", 0, "int", 0, "int", 0, "int", 0, "uint", SWP_KEEP_GEOMETRY)
        try cpArrowCtrl.Redraw()
        cpArrowEntry["state"] := cpArrowNextState
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
        CPAllowDarkModeForWindow(cpPartHwnd, darkMode)
        if darkMode {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpPartHwnd, "wstr", "DarkMode_Explorer", "ptr", 0)
        } else {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpPartHwnd, "ptr", 0, "ptr", 0)
        }
        try DllCall("user32\InvalidateRect", "ptr", cpPartHwnd, "ptr", 0, "int", 1)
    }
}

CPPrepareStudyCombo(cpComboHwnd) {
    global CPStudyThemedComboHwnds, CPStudyComboSubclassCallback
    if !cpComboHwnd || CPStudyThemedComboHwnds.Has(cpComboHwnd)
        return
    cpComboCtrl := 0
    try cpComboCtrl := GuiCtrlFromHwnd(cpComboHwnd)
    if !IsObject(cpComboCtrl)
        return
    ; Owner-draw only changes the appearance. It retains the native ComboBox,
    ; its items, focus behavior, keyboard navigation and controller wiring.
    try cpComboCtrl.Opt("+0x10") ; CBS_OWNERDRAWFIXED
    if !CPStudyComboSubclassCallback
        ; This callback performs GUI/map lookups and custom GDI painting.  A
        ; Fast callback reuses whichever AutoHotkey thread Windows interrupts,
        ; which is unsafe when a native control sends messages while another
        ; GUI event or timer is active.  Give it an ordinary callback thread so
        ; a delayed paint cannot corrupt the interpreter after a long idle.
        CPStudyComboSubclassCallback := CallbackCreate(CPStudyComboWindowProc, "", 4)
    cpComboOriginalProc := DllCall(
        "user32\SetWindowLongPtr", "ptr", cpComboHwnd, "int", -4,
        "ptr", CPStudyComboSubclassCallback, "ptr"
    )
    if cpComboOriginalProc
        CPStudyThemedComboHwnds[cpComboHwnd] := cpComboOriginalProc
}

CPPrepareStudyListHeader(cpHeaderHwnd) {
    global CPStudyThemedHeaderHwnds
    if !cpHeaderHwnd || CPStudyThemedHeaderHwnds.Has(cpHeaderHwnd)
        return
    CPStudyThemedHeaderHwnds[cpHeaderHwnd] := true
}

CPPrepareStudyListView(cpListHwnd, cpHeaderHwnd) {
    global CPStudyThemedListHwnds, CPStudyListSubclassCallback
    CPPrepareStudyListHeader(cpHeaderHwnd)
    if !cpListHwnd || CPStudyThemedListHwnds.Has(cpListHwnd)
        return
    if !CPStudyListSubclassCallback
        CPStudyListSubclassCallback := CallbackCreate(CPStudyListWindowProc, "", 4)
    cpListOriginalProc := DllCall(
        "user32\SetWindowLongPtr", "ptr", cpListHwnd, "int", -4,
        "ptr", CPStudyListSubclassCallback, "ptr"
    )
    if cpListOriginalProc
        CPStudyThemedListHwnds[cpListHwnd] := cpListOriginalProc
}

CPStudyListWindowProc(cpListHwnd, cpListMsg, cpListWParam, cpListLParam) {
    global CPStudyThemedListHwnds, CPStudyThemedHeaderHwnds
    if (cpListMsg = 0x002B && cpListLParam) { ; WM_DRAWITEM
        cpListDrawType := NumGet(cpListLParam, 0, "uint")
        cpListDrawHwndOffset := A_PtrSize = 8 ? 24 : 20
        cpListDrawHwnd := NumGet(cpListLParam, cpListDrawHwndOffset, "ptr")
        if (cpListDrawType = 100
            && IsSet(CPStudyThemedHeaderHwnds)
            && IsObject(CPStudyThemedHeaderHwnds)
            && CPStudyThemedHeaderHwnds.Has(cpListDrawHwnd)) {
            try {
                if CPThemeDrawHeaderItem(cpListLParam)
                    return 1
            }
        }
    }
    cpListMapReady := IsSet(CPStudyThemedListHwnds)
        && IsObject(CPStudyThemedListHwnds)
    cpListOriginalProc := cpListMapReady && CPStudyThemedListHwnds.Has(cpListHwnd)
        ? CPStudyThemedListHwnds[cpListHwnd] : 0
    cpListResult := cpListOriginalProc
        ? DllCall("user32\CallWindowProcW", "ptr", cpListOriginalProc,
            "ptr", cpListHwnd, "uint", cpListMsg, "uptr", cpListWParam,
            "ptr", cpListLParam, "ptr")
        : DllCall("user32\DefWindowProcW", "ptr", cpListHwnd,
            "uint", cpListMsg, "uptr", cpListWParam, "ptr", cpListLParam, "ptr")
    if (cpListMsg = 0x0082 && cpListMapReady
        && CPStudyThemedListHwnds.Has(cpListHwnd))
        CPStudyThemedListHwnds.Delete(cpListHwnd)
    return cpListResult
}

CPStudyComboWindowProc(cpComboHwnd, cpComboMsg, cpComboWParam, cpComboLParam) {
    global CPStudyThemedComboHwnds
    ; Native controls can finish WM_NCDESTROY after AutoHotkey has started
    ; releasing globals during ExitApp.  Treat a missing map as an already
    ; completed teardown instead of raising a last-moment script error.
    cpComboMapReady := IsSet(CPStudyThemedComboHwnds)
        && IsObject(CPStudyThemedComboHwnds)
    cpComboOriginalProc := cpComboMapReady && CPStudyThemedComboHwnds.Has(cpComboHwnd)
        ? CPStudyThemedComboHwnds[cpComboHwnd] : 0
    cpComboResult := cpComboOriginalProc
        ? DllCall("user32\CallWindowProcW", "ptr", cpComboOriginalProc,
            "ptr", cpComboHwnd, "uint", cpComboMsg, "uptr", cpComboWParam,
            "ptr", cpComboLParam, "ptr")
        : DllCall("user32\DefWindowProcW", "ptr", cpComboHwnd,
            "uint", cpComboMsg, "uptr", cpComboWParam, "ptr", cpComboLParam,
            "ptr")
    try {
        if (cpComboMsg = 0x000F) { ; WM_PAINT
            cpComboDc := DllCall("user32\GetDC", "ptr", cpComboHwnd, "ptr")
            if cpComboDc {
                try CPThemePaintStudyComboArrow(cpComboHwnd, cpComboDc)
                DllCall("user32\ReleaseDC", "ptr", cpComboHwnd, "ptr", cpComboDc)
            }
        } else if (cpComboMsg = 0x0317 || cpComboMsg = 0x0318) { ; WM_PRINT/CLIENT
            if cpComboWParam
                CPThemePaintStudyComboArrow(cpComboHwnd, cpComboWParam)
        }
    }
    if (cpComboMsg = 0x0082 && cpComboMapReady
        && CPStudyThemedComboHwnds.Has(cpComboHwnd))
        CPStudyThemedComboHwnds.Delete(cpComboHwnd)
    return cpComboResult
}

CPThemePaintStudyComboArrow(cpComboHwnd, cpComboDc) {
    global controlDarkMode
    cpComboRect := Buffer(16, 0)
    if !DllCall("user32\GetClientRect", "ptr", cpComboHwnd, "ptr", cpComboRect)
        return
    cpComboRight := NumGet(cpComboRect, 8, "int")
    cpComboBottom := NumGet(cpComboRect, 12, "int")
    cpComboDpi := GetWindowDPI(cpComboHwnd)
    cpComboArrowW := Max(20, Round(24 * cpComboDpi / 96))
    cpComboArrowRect := Buffer(16, 0)
    NumPut("int", Max(1, cpComboRight - cpComboArrowW), "int", 1,
        "int", Max(1, cpComboRight - 1), "int", Max(1, cpComboBottom - 1),
        cpComboArrowRect, 0)
    cpComboColors := CPPalette(controlDarkMode)
    cpComboBrush := DllCall("gdi32\CreateSolidBrush", "uint",
        CPColorRef(cpComboColors["surfaceAlt"]), "ptr")
    try DllCall("user32\FillRect", "ptr", cpComboDc,
        "ptr", cpComboArrowRect, "ptr", cpComboBrush)
    finally DllCall("gdi32\DeleteObject", "ptr", cpComboBrush)
    cpComboEnabled := DllCall("user32\IsWindowEnabled", "ptr", cpComboHwnd, "int")
    DllCall("gdi32\SetTextColor", "ptr", cpComboDc, "uint", CPColorRef(
        cpComboEnabled ? cpComboColors["text"] : cpComboColors["muted"]
    ))
    DllCall("gdi32\SetBkMode", "ptr", cpComboDc, "int", 1)
    cpComboFont := SendMessage(0x0031, 0, 0, cpComboHwnd)
    cpComboOldFont := cpComboFont
        ? DllCall("gdi32\SelectObject", "ptr", cpComboDc, "ptr", cpComboFont, "ptr") : 0
    DllCall("user32\DrawTextW", "ptr", cpComboDc, "wstr", Chr(9662),
        "int", -1, "ptr", cpComboArrowRect,
        "uint", 0x0020 | 0x0001 | 0x0004 | 0x0800)
    if cpComboOldFont
        DllCall("gdi32\SelectObject", "ptr", cpComboDc, "ptr", cpComboOldFont)
    cpComboPen := DllCall("gdi32\CreatePen", "int", 0, "int", 1,
        "uint", CPColorRef(cpComboColors["border"]), "ptr")
    cpComboOldPen := DllCall("gdi32\SelectObject", "ptr", cpComboDc,
        "ptr", cpComboPen, "ptr")
    DllCall("gdi32\MoveToEx", "ptr", cpComboDc,
        "int", Max(1, cpComboRight - cpComboArrowW), "int", 1, "ptr", 0)
    DllCall("gdi32\LineTo", "ptr", cpComboDc,
        "int", Max(1, cpComboRight - cpComboArrowW), "int", cpComboBottom - 1)
    DllCall("gdi32\SelectObject", "ptr", cpComboDc, "ptr", cpComboOldPen)
    DllCall("gdi32\DeleteObject", "ptr", cpComboPen)
}

CPStudyTransparentWindowProc(cpOverlayHwnd, cpOverlayMsg, cpOverlayWParam, cpOverlayLParam) {
    global CPStudyTransparentHwnds
    if (cpOverlayMsg = 0x0084) ; WM_NCHITTEST: pass through to native header below
        return -1 ; HTTRANSPARENT
    if (cpOverlayMsg = 0x0014) ; WM_ERASEBKGND
        return 1
    if (cpOverlayMsg = 0x000F) { ; WM_PAINT
        cpOverlayPaint := Buffer(A_PtrSize = 8 ? 72 : 64, 0)
        cpOverlayDc := DllCall("user32\BeginPaint", "ptr", cpOverlayHwnd,
            "ptr", cpOverlayPaint, "ptr")
        if cpOverlayDc
            CPThemePaintStudyHeaderOverlay(cpOverlayHwnd, cpOverlayDc)
        DllCall("user32\EndPaint", "ptr", cpOverlayHwnd, "ptr", cpOverlayPaint)
        return 0
    }
    if (cpOverlayMsg = 0x0318 && cpOverlayWParam) { ; WM_PRINTCLIENT
        CPThemePaintStudyHeaderOverlay(cpOverlayHwnd, cpOverlayWParam)
        return 0
    }
    cpOverlayMapReady := IsSet(CPStudyTransparentHwnds)
        && IsObject(CPStudyTransparentHwnds)
    cpOverlayOriginalProc := (cpOverlayMapReady
        && CPStudyTransparentHwnds.Has(cpOverlayHwnd)
        && IsObject(CPStudyTransparentHwnds[cpOverlayHwnd]))
        ? CPStudyTransparentHwnds[cpOverlayHwnd]["proc"] : 0
    cpOverlayResult := cpOverlayOriginalProc
        ? DllCall("user32\CallWindowProcW", "ptr", cpOverlayOriginalProc,
            "ptr", cpOverlayHwnd, "uint", cpOverlayMsg, "uptr", cpOverlayWParam,
            "ptr", cpOverlayLParam, "ptr")
        : DllCall("user32\DefWindowProcW", "ptr", cpOverlayHwnd,
            "uint", cpOverlayMsg, "uptr", cpOverlayWParam,
            "ptr", cpOverlayLParam, "ptr")
    if (cpOverlayMsg = 0x0082 && cpOverlayMapReady
        && CPStudyTransparentHwnds.Has(cpOverlayHwnd))
        CPStudyTransparentHwnds.Delete(cpOverlayHwnd)
    return cpOverlayResult
}

CPThemePaintStudyHeaderOverlay(cpOverlayHwnd, cpOverlayDc) {
    global CPStudyTransparentHwnds, controlDarkMode
    if !IsSet(CPStudyTransparentHwnds) || !IsObject(CPStudyTransparentHwnds)
        || !CPStudyTransparentHwnds.Has(cpOverlayHwnd)
        return
    cpOverlayData := CPStudyTransparentHwnds[cpOverlayHwnd]
    cpOverlayColors := CPPalette(controlDarkMode)
    cpOverlayRect := Buffer(16, 0)
    DllCall("user32\GetClientRect", "ptr", cpOverlayHwnd, "ptr", cpOverlayRect)
    cpOverlayBrush := DllCall("gdi32\CreateSolidBrush", "uint",
        CPColorRef(cpOverlayColors["surfaceAlt"]), "ptr")
    DllCall("user32\FillRect", "ptr", cpOverlayDc, "ptr", cpOverlayRect,
        "ptr", cpOverlayBrush)
    DllCall("gdi32\DeleteObject", "ptr", cpOverlayBrush)
    DllCall("gdi32\SetTextColor", "ptr", cpOverlayDc, "uint",
        CPColorRef(cpOverlayColors["text"]))
    DllCall("gdi32\SetBkMode", "ptr", cpOverlayDc, "int", 1)
    cpOverlayFont := SendMessage(0x0031, 0, 0, cpOverlayHwnd)
    cpOverlayOldFont := cpOverlayFont
        ? DllCall("gdi32\SelectObject", "ptr", cpOverlayDc, "ptr", cpOverlayFont, "ptr") : 0
    cpOverlayTextRect := Buffer(16, 0)
    NumPut("int", 7, "int", 0,
        "int", Max(7, NumGet(cpOverlayRect, 8, "int") - 4),
        "int", NumGet(cpOverlayRect, 12, "int"), cpOverlayTextRect, 0)
    DllCall("user32\DrawTextW", "ptr", cpOverlayDc,
        "wstr", cpOverlayData["text"], "int", -1, "ptr", cpOverlayTextRect,
        "uint", 0x0020 | 0x0004 | 0x0800 | 0x8000)
    if cpOverlayOldFont
        DllCall("gdi32\SelectObject", "ptr", cpOverlayDc, "ptr", cpOverlayOldFont)
    cpOverlayPen := DllCall("gdi32\CreatePen", "int", 0, "int", 1,
        "uint", CPColorRef(cpOverlayColors["border"]), "ptr")
    cpOverlayOldPen := DllCall("gdi32\SelectObject", "ptr", cpOverlayDc,
        "ptr", cpOverlayPen, "ptr")
    cpOverlayRight := NumGet(cpOverlayRect, 8, "int") - 1
    cpOverlayBottom := NumGet(cpOverlayRect, 12, "int") - 1
    DllCall("gdi32\MoveToEx", "ptr", cpOverlayDc, "int", cpOverlayRight,
        "int", 0, "ptr", 0)
    DllCall("gdi32\LineTo", "ptr", cpOverlayDc, "int", cpOverlayRight,
        "int", cpOverlayBottom + 1)
    DllCall("gdi32\MoveToEx", "ptr", cpOverlayDc, "int", 0,
        "int", cpOverlayBottom, "ptr", 0)
    DllCall("gdi32\LineTo", "ptr", cpOverlayDc, "int", cpOverlayRight + 1,
        "int", cpOverlayBottom)
    DllCall("gdi32\SelectObject", "ptr", cpOverlayDc, "ptr", cpOverlayOldPen)
    DllCall("gdi32\DeleteObject", "ptr", cpOverlayPen)
}

CPPrepareTransparentStudyOverlay(cpOverlayHwnd) {
    global CPStudyTransparentHwnds, CPStudyTransparentCallback
    if !cpOverlayHwnd || CPStudyTransparentHwnds.Has(cpOverlayHwnd)
        return
    if !CPStudyTransparentCallback
        CPStudyTransparentCallback := CallbackCreate(CPStudyTransparentWindowProc, "", 4)
    cpOverlayOriginalProc := DllCall(
        "user32\SetWindowLongPtr", "ptr", cpOverlayHwnd, "int", -4,
        "ptr", CPStudyTransparentCallback, "ptr"
    )
    if cpOverlayOriginalProc
        CPStudyTransparentHwnds[cpOverlayHwnd] := Map(
            "proc", cpOverlayOriginalProc, "text", ""
        )
}

CPEnsureStudyVisualOverlays(cpStudyGui) {
    global CPStudyVisualOverlays
    cpStudyRoot := 0
    try cpStudyRoot := IsObject(cpStudyGui) ? cpStudyGui.Hwnd : 0
    if !cpStudyRoot || !DllCall("user32\IsWindow", "ptr", cpStudyRoot, "int")
        return
    if !CPStudyVisualOverlays.Has(cpStudyRoot) {
        cpStudyEntry := Map("gui", cpStudyGui, "arrows", [], "headers", [])
        for cpStudyControlHwnd in WinGetControlsHwnd("ahk_id " cpStudyRoot) {
            cpStudyClass := WinGetClass("ahk_id " cpStudyControlHwnd)
            cpStudyControl := 0
            try cpStudyControl := GuiCtrlFromHwnd(cpStudyControlHwnd)
            if !IsObject(cpStudyControl)
                continue
            if (cpStudyClass = "ComboBox") {
                cpStudyArrow := cpStudyGui.Add(
                    "Text", "x0 y0 w1 h1 Hidden Center +0x100 +0x200", Chr(9662)
                )
                cpStudyArrow.Cursor := "Hand"
                cpStudyArrow.OnEvent("Click", CPComboArrowClick.Bind(cpStudyControlHwnd))
                cpStudyEntry["arrows"].Push(Map(
                    "combo", cpStudyControl, "arrow", cpStudyArrow, "last", ""
                ))
            } else if (cpStudyClass = "SysListView32") {
                cpStudyHeaderHwnd := SendMessage(0x101F, 0, 0, cpStudyControlHwnd)
                cpStudyHeaderCount := cpStudyHeaderHwnd
                    ? SendMessage(0x1200, 0, 0, cpStudyHeaderHwnd) : 0
                cpStudyLabels := []
                loop Max(0, cpStudyHeaderCount) {
                    cpStudyLabelHwnd := DllCall(
                        "user32\CreateWindowExW", "uint", 0,
                        "wstr", "Static", "wstr", "",
                        "uint", 0x40000000 | 0x04000000,
                        "int", 0, "int", 0, "int", 1, "int", 1,
                        "ptr", cpStudyHeaderHwnd, "ptr", 0,
                        "ptr", DllCall("kernel32\GetModuleHandleW", "ptr", 0, "ptr"),
                        "ptr", 0, "ptr"
                    )
                    if !cpStudyLabelHwnd
                        continue
                    cpStudyHeaderFont := SendMessage(0x0031, 0, 0, cpStudyHeaderHwnd)
                    if cpStudyHeaderFont
                        SendMessage(0x0030, cpStudyHeaderFont, 0, cpStudyLabelHwnd)
                    CPPrepareTransparentStudyOverlay(cpStudyLabelHwnd)
                    cpStudyLabels.Push(Map("hwnd", cpStudyLabelHwnd, "last", ""))
                }
                cpStudyEntry["headers"].Push(Map(
                    "list", cpStudyControl,
                    "headerHwnd", cpStudyHeaderHwnd,
                    "labels", cpStudyLabels
                ))
            }
        }
        CPStudyVisualOverlays[cpStudyRoot] := cpStudyEntry
    }
    CPUpdateStudyVisualOverlays(cpStudyRoot)
    SetTimer(CPUpdateAllStudyVisualOverlays, 125)
}

CPStudyHeaderText(cpHeaderHwnd, cpHeaderIndex) {
    if !cpHeaderHwnd || !DllCall("user32\IsWindow", "ptr", cpHeaderHwnd, "int")
        return ""
    cpStudyTextBuffer := Buffer(1024, 0)
    cpStudyHeaderItem := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
    NumPut("uint", 0x0002, cpStudyHeaderItem, 0)
    NumPut("ptr", cpStudyTextBuffer.Ptr, cpStudyHeaderItem, 8)
    NumPut("int", 511, cpStudyHeaderItem, A_PtrSize = 8 ? 24 : 16)
    if SendMessage(0x120B, cpHeaderIndex, cpStudyHeaderItem.Ptr, cpHeaderHwnd)
        return StrGet(cpStudyTextBuffer, "UTF-16")
    return ""
}

CPStudyLiveControlHwnd(cpStudyControl) {
    cpStudyHwnd := 0
    try cpStudyHwnd := IsObject(cpStudyControl) ? cpStudyControl.Hwnd : 0
    return (cpStudyHwnd
        && DllCall("user32\IsWindow", "ptr", cpStudyHwnd, "int"))
        ? cpStudyHwnd : 0
}

CPUpdateStudyVisualOverlays(cpStudyRoot) {
    global CPStudyVisualOverlays, CPStudyTransparentHwnds, controlDarkMode
    if !CPStudyVisualOverlays.Has(cpStudyRoot)
        return
    cpStudyEntry := CPStudyVisualOverlays[cpStudyRoot]
    cpStudyColors := CPPalette(controlDarkMode)
    static SWP_KEEP_GEOMETRY := 0x0001 | 0x0002 | 0x0010

    for cpStudyArrowEntry in cpStudyEntry["arrows"] {
        cpStudyCombo := cpStudyArrowEntry["combo"]
        cpStudyArrow := cpStudyArrowEntry["arrow"]
        cpStudyComboHwnd := CPStudyLiveControlHwnd(cpStudyCombo)
        cpStudyArrowHwnd := CPStudyLiveControlHwnd(cpStudyArrow)
        if !cpStudyComboHwnd || !cpStudyArrowHwnd
            continue
        cpStudyShow := DllCall("user32\IsWindowVisible", "ptr", cpStudyComboHwnd, "int")
        if !cpStudyShow {
            if (cpStudyArrowEntry["last"] != "hidden") {
                try cpStudyArrow.Visible := false
                cpStudyArrowEntry["last"] := "hidden"
            }
            continue
        }
        try cpStudyCombo.GetPos(&cpStudyX, &cpStudyY, &cpStudyW, &cpStudyH)
        catch
            continue
        cpStudyArrowW := Min(28, Max(20, Floor(cpStudyH * 0.85)))
        cpStudyEnabled := DllCall("user32\IsWindowEnabled", "ptr", cpStudyComboHwnd, "int")
        cpStudyArrowColor := cpStudyEnabled
            ? cpStudyColors["text"] : cpStudyColors["muted"]
        cpStudyArrowState := (cpStudyX "|" cpStudyY "|" cpStudyW "|" cpStudyH
            "|" cpStudyArrowColor "|" cpStudyColors["surfaceAlt"])
        if (cpStudyArrowEntry["last"] != cpStudyArrowState) {
            try {
                cpStudyArrow.Move(cpStudyX + cpStudyW - cpStudyArrowW,
                    cpStudyY + 1, cpStudyArrowW - 1, Max(1, cpStudyH - 2))
                cpStudyArrow.Opt("+Background" cpStudyColors["surfaceAlt"])
                cpStudyArrow.SetFont("s9 c" cpStudyArrowColor)
                cpStudyArrow.Visible := true
            } catch {
                continue
            }
            try DllCall("user32\SetWindowPos", "ptr", cpStudyArrowHwnd, "ptr", 0,
                "int", 0, "int", 0, "int", 0, "int", 0,
                "uint", SWP_KEEP_GEOMETRY)
            cpStudyArrowEntry["last"] := cpStudyArrowState
        }
    }

    for cpStudyHeaderEntry in cpStudyEntry["headers"] {
        cpStudyHeader := cpStudyHeaderEntry["headerHwnd"]
        cpStudyHeaderLive := (cpStudyHeader
            && DllCall("user32\IsWindow", "ptr", cpStudyHeader, "int"))
        if !cpStudyHeaderLive
            continue
        if !DllCall("user32\IsWindowVisible", "ptr", cpStudyHeader, "int") {
            for cpStudyLabelEntry in cpStudyHeaderEntry["labels"] {
                cpStudyLabelHwnd := cpStudyLabelEntry["hwnd"]
                if !cpStudyLabelHwnd || !DllCall("user32\IsWindow", "ptr", cpStudyLabelHwnd, "int")
                    continue
                if (cpStudyLabelEntry["last"] != "hidden") {
                    try DllCall("user32\ShowWindow", "ptr",
                        cpStudyLabelHwnd, "int", 0)
                    cpStudyLabelEntry["last"] := "hidden"
                }
            }
            continue
        }
        cpStudyHeaderRect := Buffer(16, 0)
        DllCall("user32\GetClientRect", "ptr", cpStudyHeader, "ptr", cpStudyHeaderRect)
        cpStudyHeaderH := NumGet(cpStudyHeaderRect, 12, "int")
        for cpStudyIndex, cpStudyLabelEntry in cpStudyHeaderEntry["labels"] {
            cpStudyLabelHwnd := cpStudyLabelEntry["hwnd"]
            if !cpStudyLabelHwnd || !DllCall("user32\IsWindow", "ptr", cpStudyLabelHwnd, "int")
                continue
            cpStudyItemRect := Buffer(16, 0)
            if !SendMessage(0x1207, cpStudyIndex - 1,
                cpStudyItemRect.Ptr, cpStudyHeader) {
                if (cpStudyLabelEntry["last"] != "hidden") {
                    try DllCall("user32\ShowWindow", "ptr", cpStudyLabelHwnd, "int", 0)
                    cpStudyLabelEntry["last"] := "hidden"
                }
                continue
            }
            cpStudyLeft := NumGet(cpStudyItemRect, 0, "int")
            cpStudyRight := NumGet(cpStudyItemRect, 8, "int")
            cpStudyItemW := cpStudyRight - cpStudyLeft
            if (cpStudyItemW <= 0) {
                if (cpStudyLabelEntry["last"] != "hidden") {
                    try DllCall("user32\ShowWindow", "ptr", cpStudyLabelHwnd, "int", 0)
                    cpStudyLabelEntry["last"] := "hidden"
                }
                continue
            }
            cpStudyLabelText := CPStudyHeaderText(cpStudyHeader, cpStudyIndex - 1)
            cpStudyLabelState := (cpStudyLeft "|0"
                "|" cpStudyItemW "|" cpStudyHeaderH "|" cpStudyLabelText
                "|" cpStudyColors["surfaceAlt"] "|" cpStudyColors["text"])
            if (cpStudyLabelEntry["last"] != cpStudyLabelState) {
                if CPStudyTransparentHwnds.Has(cpStudyLabelHwnd)
                    CPStudyTransparentHwnds[cpStudyLabelHwnd]["text"] := cpStudyLabelText
                try DllCall("user32\SetWindowPos", "ptr", cpStudyLabelHwnd, "ptr", 0,
                    "int", cpStudyLeft, "int", 0, "int", cpStudyItemW,
                    "int", cpStudyHeaderH, "uint", 0x0010 | 0x0040)
                try DllCall("user32\InvalidateRect", "ptr", cpStudyLabelHwnd,
                    "ptr", 0, "int", 1)
                cpStudyLabelEntry["last"] := cpStudyLabelState
            }
        }
    }
}

CPUpdateAllStudyVisualOverlays(*) {
    global CPStudyVisualOverlays
    cpStudyDeadRoots := []
    for cpStudyRoot, cpStudyEntry in CPStudyVisualOverlays {
        if !DllCall("user32\IsWindow", "ptr", cpStudyRoot, "int")
            cpStudyDeadRoots.Push(cpStudyRoot)
        else {
            try CPUpdateStudyVisualOverlays(cpStudyRoot)
            catch as cpStudyOverlayError
                DbgCP("Study visual overlay update failed: " cpStudyOverlayError.Message)
        }
    }
    for cpStudyRoot in cpStudyDeadRoots
        CPStudyVisualOverlays.Delete(cpStudyRoot)
    if (CPStudyVisualOverlays.Count = 0)
        SetTimer(CPUpdateAllStudyVisualOverlays, 0)
}

CPStudyHeaderWindowProc(cpHeaderHwnd, cpHeaderMsg, cpHeaderWParam, cpHeaderLParam) {
    global CPStudyThemedHeaderHwnds
    cpHeaderMapReady := IsSet(CPStudyThemedHeaderHwnds)
        && IsObject(CPStudyThemedHeaderHwnds)
    cpHeaderOriginalProc := cpHeaderMapReady && CPStudyThemedHeaderHwnds.Has(cpHeaderHwnd)
        ? CPStudyThemedHeaderHwnds[cpHeaderHwnd] : 0
    cpHeaderResult := cpHeaderOriginalProc
        ? DllCall("user32\CallWindowProcW", "ptr", cpHeaderOriginalProc,
            "ptr", cpHeaderHwnd, "uint", cpHeaderMsg, "uptr", cpHeaderWParam,
            "ptr", cpHeaderLParam, "ptr")
        : DllCall("user32\DefWindowProcW", "ptr", cpHeaderHwnd,
            "uint", cpHeaderMsg, "uptr", cpHeaderWParam, "ptr", cpHeaderLParam,
            "ptr")
    try {
        if (cpHeaderMsg = 0x000F) { ; WM_PAINT
            cpHeaderDc := DllCall("user32\GetDC", "ptr", cpHeaderHwnd, "ptr")
            if cpHeaderDc {
                try CPThemePaintAllHeaders(cpHeaderHwnd, cpHeaderDc)
                DllCall("user32\ReleaseDC", "ptr", cpHeaderHwnd, "ptr", cpHeaderDc)
            }
        } else if (cpHeaderMsg = 0x0317 || cpHeaderMsg = 0x0318) {
            if cpHeaderWParam
                CPThemePaintAllHeaders(cpHeaderHwnd, cpHeaderWParam)
        }
    }
    if (cpHeaderMsg = 0x0082 && cpHeaderMapReady
        && CPStudyThemedHeaderHwnds.Has(cpHeaderHwnd))
        CPStudyThemedHeaderHwnds.Delete(cpHeaderHwnd)
    return cpHeaderResult
}

CPThemePaintAllHeaders(cpHeaderHwnd, cpHeaderDc) {
    cpHeaderCount := SendMessage(0x1200, 0, 0, cpHeaderHwnd)
    loop Max(0, cpHeaderCount) {
        cpHeaderIndex := A_Index - 1
        cpHeaderRect := Buffer(16, 0)
        if SendMessage(0x1207, cpHeaderIndex, cpHeaderRect.Ptr, cpHeaderHwnd)
            CPThemePaintHeader(cpHeaderHwnd, cpHeaderIndex, cpHeaderDc,
                cpHeaderRect.Ptr, 0)
    }
}

CPApplyThemeToControl(cpThemeHwnd, darkMode := -1) {
    global controlDarkMode, CPStudyThemedComboHwnds
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
    CPAllowDarkModeForWindow(cpThemeHwnd, darkMode)

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
            cpFieldTheme := "DarkMode_CFD"
            if (cpThemeClass = "ComboBox" && CPStudyThemedComboHwnds.Has(cpThemeHwnd))
                cpFieldTheme := "DarkMode_Explorer"
            if (cpThemeClass = "Edit") {
                cpFieldStyle := DllCall(
                    "user32\GetWindowLongPtr", "ptr", cpThemeHwnd, "int", -16, "ptr"
                )
                ; Explorer supplies dark native scrollbars for multiline/read-only
                ; viewer fields; the CFD edit theme otherwise leaves a light gutter.
                if ((cpFieldStyle & 0x0004) || (cpFieldStyle & 0x00200000))
                    cpFieldTheme := "DarkMode_Explorer"
            }
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "wstr", cpFieldTheme, "ptr", 0)
        } else {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "ptr", 0, "ptr", 0)
        }
        if (cpThemeClass = "ComboBox")
            CPThemeComboParts(cpThemeHwnd, darkMode)
    } else if (cpThemeClass = "SysListView32") {
        if darkMode {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "wstr", "DarkMode_Explorer", "ptr", 0)
        } else {
            try DllCall("uxtheme\SetWindowTheme", "ptr", cpThemeHwnd, "ptr", 0, "ptr", 0)
        }
        ; Explorer theming alone is not sufficient on every Windows build,
        ; especially when a ListView was created before its standalone window
        ; was registered for dark painting. Set its three native colors too.
        cpListBack := CPColorRef(cpThemeColors["surface"])
        cpListText := CPColorRef(cpThemeColors["text"])
        try SendMessage(0x1001, 0, cpListBack, cpThemeHwnd) ; LVM_SETBKCOLOR
        try SendMessage(0x1026, 0, cpListBack, cpThemeHwnd) ; LVM_SETTEXTBKCOLOR
        try SendMessage(0x1024, 0, cpListText, cpThemeHwnd) ; LVM_SETTEXTCOLOR
        cpListHeader := 0
        try cpListHeader := SendMessage(0x101F, 0, 0, cpThemeHwnd) ; LVM_GETHEADER
        if cpListHeader {
            CPAllowDarkModeForWindow(cpListHeader, darkMode)
            if darkMode {
                ; Keep the native light header for reliable contrast. Windows does
                ; not expose a safe independent header-text color setting here.
                try DllCall("uxtheme\SetWindowTheme", "ptr", cpListHeader, "wstr", "DarkMode_Explorer", "ptr", 0)
            } else {
                try DllCall("uxtheme\SetWindowTheme", "ptr", cpListHeader, "ptr", 0, "ptr", 0)
            }
            try DllCall("user32\InvalidateRect", "ptr", cpListHeader, "ptr", 0, "int", 1)
        }
    } else if (cpThemeClass = "msctls_trackbar32" || cpThemeClass = "msctls_updown32"
        || cpThemeClass = "SysTabControl32") {
        if (cpThemeClass = "SysTabControl32")
            try cpThemeCtrl.SetFont("c" cpThemeColors["text"])
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

CPAllowDarkModeForWindow(cpThemeHwnd, darkMode) {
    if cpThemeHwnd
        try DllCall("uxtheme\#133", "ptr", cpThemeHwnd, "int", darkMode ? 1 : 0, "int")
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
        , CPThemedDialogHwnds
    if !controlDarkMode || !(IsSet(ui) && ui && ui.Hwnd)
        return
    cpThemeRoot := 0
    cpThemeParentRoot := 0
    try cpThemeRoot := DllCall("user32\GetAncestor", "ptr", lParam, "uint", 2, "ptr") ; GA_ROOT
    try cpThemeParentRoot := DllCall("user32\GetAncestor", "ptr", parentHwnd, "uint", 2, "ptr")
    cpThemeOwner := DllCall("user32\GetWindow", "ptr", parentHwnd, "uint", 4, "ptr") ; GW_OWNER
    cpThemeRootTracked := false
    cpThemeParentTracked := false
    try cpThemeRootTracked := (cpThemeRoot
        && IsSet(CPThemedDialogHwnds) && IsObject(CPThemedDialogHwnds)
        && CPThemedDialogHwnds.Has(cpThemeRoot))
    try cpThemeParentTracked := (cpThemeParentRoot
        && IsSet(CPThemedDialogHwnds) && IsObject(CPThemedDialogHwnds)
        && CPThemedDialogHwnds.Has(cpThemeParentRoot))
    if (parentHwnd != ui.Hwnd
        && cpThemeOwner != ui.Hwnd
        && !DllCall("user32\IsChild", "ptr", ui.Hwnd, "ptr", lParam, "int")
        && !cpThemeRootTracked
        && !cpThemeParentTracked)
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

CPThemedWindowDestroyed(wParam, lParam, msg, hwnd) {
    global CPThemedDialogHwnds, CPStudyThemedHeaderHwnds
        , CPStudyVisualOverlays, CPThemedPopupButtons
    try {
        if IsSet(CPThemedDialogHwnds) && IsObject(CPThemedDialogHwnds)
            && CPThemedDialogHwnds.Has(hwnd)
            CPThemedDialogHwnds.Delete(hwnd)
    }
    ; Do not clear the ComboBox/ListView/overlay subclass maps here.  Their
    ; window procedures still need the saved original WNDPROC while processing
    ; WM_NCDESTROY and remove their own entry only after CallWindowProc returns.
    ; Clearing it from this general OnMessage hook can make a subclass fall
    ; through to DefWindowProc mid-destruction, corrupting native control state.
    try {
        if IsSet(CPStudyThemedHeaderHwnds) && IsObject(CPStudyThemedHeaderHwnds)
            && CPStudyThemedHeaderHwnds.Has(hwnd)
            CPStudyThemedHeaderHwnds.Delete(hwnd)
    }
    try {
        if IsSet(CPStudyVisualOverlays) && IsObject(CPStudyVisualOverlays)
            && CPStudyVisualOverlays.Has(hwnd)
            CPStudyVisualOverlays.Delete(hwnd)
    }
    try {
        if IsSet(CPThemedPopupButtons) && IsObject(CPThemedPopupButtons)
            && CPThemedPopupButtons.Has(hwnd)
            CPThemedPopupButtons.Delete(hwnd)
    }
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
        , CPThemedPopupButtons
    if !lParam
        return
    cpDrawType := NumGet(lParam, 0, "uint")
    if (cpDrawType = 100) ; ODT_HEADER
        return CPThemeDrawHeaderItem(lParam)
    if (cpDrawType = 4) { ; ODT_BUTTON
        cpButtonHwndOffset := (A_PtrSize = 8) ? 24 : 20
        cpButtonHwnd := NumGet(lParam, cpButtonHwndOffset, "ptr")
        if cpButtonHwnd && IsSet(CPThemedPopupButtons)
            && IsObject(CPThemedPopupButtons)
            && CPThemedPopupButtons.Has(cpButtonHwnd)
            return CPThemeDrawPopupButton(lParam)
        return
    }
    if (cpDrawType != 3) ; ODT_COMBOBOX
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

CPThemeDrawPopupButton(cpButtonDraw) {
    global controlDarkMode, CPThemedPopupButtons
    cpButtonHwndOffset := (A_PtrSize = 8) ? 24 : 20
    cpButtonHdcOffset := (A_PtrSize = 8) ? 32 : 24
    cpButtonRectOffset := (A_PtrSize = 8) ? 40 : 28
    cpButtonHwnd := NumGet(cpButtonDraw, cpButtonHwndOffset, "ptr")
    cpButtonHdc := NumGet(cpButtonDraw, cpButtonHdcOffset, "ptr")
    if !cpButtonHwnd || !cpButtonHdc || !CPThemedPopupButtons.Has(cpButtonHwnd)
        return

    cpButtonState := NumGet(cpButtonDraw, 16, "uint")
    cpButtonPressed := (cpButtonState & 0x0001) != 0 ; ODS_SELECTED
    cpButtonDisabled := (cpButtonState & 0x0002) != 0
        || (cpButtonState & 0x0004) != 0
    cpButtonFocused := (cpButtonState & 0x0010) != 0 ; ODS_FOCUS
    cpButtonColors := CPPalette(controlDarkMode)
    cpButtonBackHex := cpButtonPressed
        ? cpButtonColors["accentFocus"]
        : (cpButtonFocused ? cpButtonColors["focus"] : cpButtonColors["surface"])
    cpButtonTextHex := cpButtonDisabled
        ? cpButtonColors["muted"]
        : (cpButtonPressed ? cpButtonColors["accentText"] : cpButtonColors["text"])

    cpButtonBackBrush := DllCall(
        "gdi32\CreateSolidBrush", "uint", CPColorRef(cpButtonBackHex), "ptr"
    )
    try DllCall(
        "user32\FillRect", "ptr", cpButtonHdc,
        "ptr", cpButtonDraw + cpButtonRectOffset, "ptr", cpButtonBackBrush
    )
    finally DllCall("gdi32\DeleteObject", "ptr", cpButtonBackBrush)

    cpButtonBorderBrush := DllCall(
        "gdi32\CreateSolidBrush", "uint", CPColorRef(cpButtonColors["border"]), "ptr"
    )
    try DllCall(
        "user32\FrameRect", "ptr", cpButtonHdc,
        "ptr", cpButtonDraw + cpButtonRectOffset, "ptr", cpButtonBorderBrush
    )
    finally DllCall("gdi32\DeleteObject", "ptr", cpButtonBorderBrush)

    DllCall(
        "gdi32\SetTextColor", "ptr", cpButtonHdc,
        "uint", CPColorRef(cpButtonTextHex)
    )
    DllCall("gdi32\SetBkMode", "ptr", cpButtonHdc, "int", 1) ; TRANSPARENT
    cpButtonFont := SendMessage(0x0031, 0, 0, cpButtonHwnd) ; WM_GETFONT
    cpButtonOldFont := cpButtonFont
        ? DllCall("gdi32\SelectObject", "ptr", cpButtonHdc, "ptr", cpButtonFont, "ptr") : 0
    cpButtonTextRect := Buffer(16, 0)
    cpButtonLeft := NumGet(cpButtonDraw, cpButtonRectOffset, "int") + 10
    cpButtonTop := NumGet(cpButtonDraw, cpButtonRectOffset + 4, "int")
    cpButtonRight := NumGet(cpButtonDraw, cpButtonRectOffset + 8, "int") - 8
    cpButtonBottom := NumGet(cpButtonDraw, cpButtonRectOffset + 12, "int")
    if cpButtonPressed
        cpButtonLeft += 1, cpButtonTop += 1
    NumPut(
        "int", cpButtonLeft, "int", cpButtonTop,
        "int", cpButtonRight, "int", cpButtonBottom,
        cpButtonTextRect, 0
    )
    DllCall(
        "user32\DrawTextW", "ptr", cpButtonHdc,
        "wstr", CPThemedPopupButtons[cpButtonHwnd], "int", -1,
        "ptr", cpButtonTextRect.Ptr,
        "uint", 0x0020 | 0x0004 | 0x0800 | 0x8000
    ) ; SINGLELINE | VCENTER | NOPREFIX | END_ELLIPSIS
    if cpButtonOldFont
        DllCall("gdi32\SelectObject", "ptr", cpButtonHdc, "ptr", cpButtonOldFont, "ptr")
    if cpButtonFocused
        DllCall(
            "user32\DrawFocusRect", "ptr", cpButtonHdc,
            "ptr", cpButtonDraw + cpButtonRectOffset
        )
    return true
}

CPThemeDrawHeaderItem(cpHeaderDraw) {
    global controlDarkMode, CPStudyThemedHeaderHwnds
    cpHeaderHwndOffset := A_PtrSize = 8 ? 24 : 20
    cpHeaderDcOffset := A_PtrSize = 8 ? 32 : 24
    cpHeaderRectOffset := A_PtrSize = 8 ? 40 : 28
    cpHeaderHwnd := NumGet(cpHeaderDraw, cpHeaderHwndOffset, "ptr")
    if !cpHeaderHwnd || !CPStudyThemedHeaderHwnds.Has(cpHeaderHwnd)
        return
    cpHeaderDc := NumGet(cpHeaderDraw, cpHeaderDcOffset, "ptr")
    cpHeaderIndex := NumGet(cpHeaderDraw, 8, "uint")
    cpHeaderState := NumGet(cpHeaderDraw, 16, "uint")
    if !cpHeaderDc
        return

    cpHeaderColors := CPPalette(controlDarkMode)
    cpHeaderBackHex := (cpHeaderState & 0x0040)
        ? cpHeaderColors["focus"] : cpHeaderColors["surfaceAlt"]
    cpHeaderBrush := DllCall(
        "gdi32\CreateSolidBrush", "uint", CPColorRef(cpHeaderBackHex), "ptr"
    )
    try DllCall(
        "user32\FillRect", "ptr", cpHeaderDc,
        "ptr", cpHeaderDraw + cpHeaderRectOffset, "ptr", cpHeaderBrush
    )
    finally DllCall("gdi32\DeleteObject", "ptr", cpHeaderBrush)

    cpHeaderTextBuffer := Buffer(1024, 0)
    cpHeaderTextItem := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
    NumPut("uint", 0x0002, cpHeaderTextItem, 0) ; HDI_TEXT
    NumPut("ptr", cpHeaderTextBuffer.Ptr, cpHeaderTextItem, 8)
    NumPut("int", 511, cpHeaderTextItem, A_PtrSize = 8 ? 24 : 16)
    cpHeaderText := ""
    if SendMessage(0x120B, cpHeaderIndex, cpHeaderTextItem.Ptr, cpHeaderHwnd)
        cpHeaderText := StrGet(cpHeaderTextBuffer, "UTF-16")

    DllCall(
        "gdi32\SetTextColor", "ptr", cpHeaderDc,
        "uint", CPColorRef(cpHeaderColors["text"])
    )
    DllCall("gdi32\SetBkMode", "ptr", cpHeaderDc, "int", 1) ; TRANSPARENT
    cpHeaderFont := SendMessage(0x0031, 0, 0, cpHeaderHwnd) ; WM_GETFONT
    cpHeaderOldFont := cpHeaderFont
        ? DllCall("gdi32\SelectObject", "ptr", cpHeaderDc, "ptr", cpHeaderFont, "ptr") : 0
    cpHeaderTextRect := Buffer(16, 0)
    cpHeaderLeft := NumGet(cpHeaderDraw, cpHeaderRectOffset, "int") + 7
    cpHeaderTop := NumGet(cpHeaderDraw, cpHeaderRectOffset + 4, "int")
    cpHeaderRight := NumGet(cpHeaderDraw, cpHeaderRectOffset + 8, "int") - 4
    cpHeaderBottom := NumGet(cpHeaderDraw, cpHeaderRectOffset + 12, "int")
    NumPut(
        "int", cpHeaderLeft, "int", cpHeaderTop,
        "int", cpHeaderRight, "int", cpHeaderBottom,
        cpHeaderTextRect, 0
    )
    DllCall(
        "user32\DrawTextW", "ptr", cpHeaderDc, "wstr", cpHeaderText, "int", -1,
        "ptr", cpHeaderTextRect, "uint", 0x0020 | 0x0004 | 0x0800 | 0x8000
    ) ; SINGLELINE | VCENTER | NOPREFIX | END_ELLIPSIS
    if cpHeaderOldFont
        DllCall("gdi32\SelectObject", "ptr", cpHeaderDc, "ptr", cpHeaderOldFont)

    cpHeaderPen := DllCall(
        "gdi32\CreatePen", "int", 0, "int", 1,
        "uint", CPColorRef(cpHeaderColors["border"]), "ptr"
    )
    cpHeaderOldPen := DllCall(
        "gdi32\SelectObject", "ptr", cpHeaderDc, "ptr", cpHeaderPen, "ptr"
    )
    DllCall("gdi32\MoveToEx", "ptr", cpHeaderDc, "int", cpHeaderRight + 3,
        "int", cpHeaderTop, "ptr", 0)
    DllCall("gdi32\LineTo", "ptr", cpHeaderDc, "int", cpHeaderRight + 3,
        "int", cpHeaderBottom)
    DllCall("gdi32\MoveToEx", "ptr", cpHeaderDc, "int", cpHeaderLeft - 7,
        "int", cpHeaderBottom - 1, "ptr", 0)
    DllCall("gdi32\LineTo", "ptr", cpHeaderDc, "int", cpHeaderRight + 4,
        "int", cpHeaderBottom - 1)
    DllCall("gdi32\SelectObject", "ptr", cpHeaderDc, "ptr", cpHeaderOldPen)
    DllCall("gdi32\DeleteObject", "ptr", cpHeaderPen)
    return true
}

CPThemeNotify(wParam, lParam, msg, parentHwnd) {
    global CPStudyThemedHeaderHwnds
    if !lParam
        return
    cpNotifyHeader := NumGet(lParam, 0, "ptr")
    if !cpNotifyHeader || !CPStudyThemedHeaderHwnds.Has(cpNotifyHeader)
        return
    cpNotifyCode := NumGet(lParam, A_PtrSize = 8 ? 16 : 8, "int")
    if (cpNotifyCode != -12) ; NM_CUSTOMDRAW
        return
    cpNotifyStage := NumGet(lParam, A_PtrSize = 8 ? 24 : 12, "uint")
    if (cpNotifyStage = 0x00000001) ; CDDS_PREPAINT
        return 0x00000020 ; CDRF_NOTIFYITEMDRAW
    if (cpNotifyStage != 0x00010001) ; CDDS_ITEMPREPAINT
        return
    cpNotifyDc := NumGet(lParam, A_PtrSize = 8 ? 32 : 16, "ptr")
    cpNotifyRect := lParam + (A_PtrSize = 8 ? 40 : 20)
    cpNotifyItem := NumGet(lParam, A_PtrSize = 8 ? 56 : 36, "uptr")
    cpNotifyState := NumGet(lParam, A_PtrSize = 8 ? 64 : 40, "uint")
    CPThemePaintHeader(
        cpNotifyHeader, cpNotifyItem, cpNotifyDc, cpNotifyRect, cpNotifyState
    )
    return 0x00000004 ; CDRF_SKIPDEFAULT
}

CPThemePaintHeader(cpHeaderHwnd, cpHeaderIndex, cpHeaderDc, cpHeaderRect, cpHeaderState) {
    global controlDarkMode
    if !cpHeaderDc
        return
    cpHeaderColors := CPPalette(controlDarkMode)
    cpHeaderBackHex := (cpHeaderState & 0x0040)
        ? cpHeaderColors["focus"] : cpHeaderColors["surfaceAlt"]
    cpHeaderBrush := DllCall(
        "gdi32\CreateSolidBrush", "uint", CPColorRef(cpHeaderBackHex), "ptr"
    )
    try DllCall(
        "user32\FillRect", "ptr", cpHeaderDc, "ptr", cpHeaderRect,
        "ptr", cpHeaderBrush
    )
    finally DllCall("gdi32\DeleteObject", "ptr", cpHeaderBrush)

    cpHeaderTextBuffer := Buffer(1024, 0)
    cpHeaderTextItem := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
    NumPut("uint", 0x0002, cpHeaderTextItem, 0) ; HDI_TEXT
    NumPut("ptr", cpHeaderTextBuffer.Ptr, cpHeaderTextItem, 8)
    NumPut("int", 511, cpHeaderTextItem, A_PtrSize = 8 ? 24 : 16)
    cpHeaderText := ""
    if SendMessage(0x120B, cpHeaderIndex, cpHeaderTextItem.Ptr, cpHeaderHwnd)
        cpHeaderText := StrGet(cpHeaderTextBuffer, "UTF-16")

    DllCall("gdi32\SetTextColor", "ptr", cpHeaderDc,
        "uint", CPColorRef(cpHeaderColors["text"]))
    DllCall("gdi32\SetBkMode", "ptr", cpHeaderDc, "int", 1)
    cpHeaderFont := SendMessage(0x0031, 0, 0, cpHeaderHwnd)
    cpHeaderOldFont := cpHeaderFont
        ? DllCall("gdi32\SelectObject", "ptr", cpHeaderDc, "ptr", cpHeaderFont, "ptr") : 0
    cpHeaderTextRect := Buffer(16, 0)
    cpHeaderLeft := NumGet(cpHeaderRect, 0, "int") + 7
    cpHeaderTop := NumGet(cpHeaderRect, 4, "int")
    cpHeaderRight := NumGet(cpHeaderRect, 8, "int") - 4
    cpHeaderBottom := NumGet(cpHeaderRect, 12, "int")
    NumPut("int", cpHeaderLeft, "int", cpHeaderTop,
        "int", cpHeaderRight, "int", cpHeaderBottom, cpHeaderTextRect, 0)
    DllCall("user32\DrawTextW", "ptr", cpHeaderDc, "wstr", cpHeaderText,
        "int", -1, "ptr", cpHeaderTextRect,
        "uint", 0x0020 | 0x0004 | 0x0800 | 0x8000)
    if cpHeaderOldFont
        DllCall("gdi32\SelectObject", "ptr", cpHeaderDc, "ptr", cpHeaderOldFont)

    cpHeaderPen := DllCall("gdi32\CreatePen", "int", 0, "int", 1,
        "uint", CPColorRef(cpHeaderColors["border"]), "ptr")
    cpHeaderOldPen := DllCall("gdi32\SelectObject", "ptr", cpHeaderDc,
        "ptr", cpHeaderPen, "ptr")
    DllCall("gdi32\MoveToEx", "ptr", cpHeaderDc, "int", cpHeaderRight + 3,
        "int", cpHeaderTop, "ptr", 0)
    DllCall("gdi32\LineTo", "ptr", cpHeaderDc, "int", cpHeaderRight + 3,
        "int", cpHeaderBottom)
    DllCall("gdi32\MoveToEx", "ptr", cpHeaderDc, "int", cpHeaderLeft - 7,
        "int", cpHeaderBottom - 1, "ptr", 0)
    DllCall("gdi32\LineTo", "ptr", cpHeaderDc, "int", cpHeaderRight + 4,
        "int", cpHeaderBottom - 1)
    DllCall("gdi32\SelectObject", "ptr", cpHeaderDc, "ptr", cpHeaderOldPen)
    DllCall("gdi32\DeleteObject", "ptr", cpHeaderPen)
}

CPRegisterThemeMessages() {
    global CPThemeMessagesRegistered
    if CPThemeMessagesRegistered
        return
    for cpThemeMsg in [0x0133, 0x0134, 0x0135, 0x0138]
        OnMessage(cpThemeMsg, CPThemeCtlColor)
    OnMessage(0x002B, CPThemeDrawItem)
    OnMessage(0x002C, CPThemeMeasureItem)
    OnMessage(0x0082, CPThemedWindowDestroyed) ; WM_NCDESTROY
    OnExit(CPShutdownThemeMessages)
    CPThemeMessagesRegistered := true
}

CPUnregisterThemeMessages(*) {
    global CPThemeMessagesRegistered
    if !(IsSet(CPThemeMessagesRegistered) && CPThemeMessagesRegistered)
        return
    ; Stop paint/destruction callbacks before ExitApp tears down GUI controls and
    ; their global tracking maps. An in-flight callback remains independently safe.
    for cpThemeMsg in [0x0133, 0x0134, 0x0135, 0x0138]
        try OnMessage(cpThemeMsg, CPThemeCtlColor, 0)
    try OnMessage(0x002B, CPThemeDrawItem, 0)
    try OnMessage(0x002C, CPThemeMeasureItem, 0)
    try OnMessage(0x0082, CPThemedWindowDestroyed, 0)
    CPThemeMessagesRegistered := false
}

CPUnsubclassStudyControls() {
    global CPStudyThemedComboHwnds, CPStudyThemedListHwnds
        , CPStudyThemedHeaderHwnds, CPStudyTransparentHwnds
        , CPStudyVisualOverlays

    ; Remove our native WNDPROC hooks while the tracking maps and saved original
    ; procedures still exist. Otherwise Windows may deliver a final destruction
    ; message after AutoHotkey has begun releasing globals during ExitApp.
    SetTimer(CPUpdateAllStudyVisualOverlays, 0)
    if IsSet(CPStudyThemedComboHwnds) && IsObject(CPStudyThemedComboHwnds) {
        for cpComboHwnd, cpComboOriginalProc in CPStudyThemedComboHwnds {
            if cpComboHwnd && cpComboOriginalProc
                && DllCall("user32\IsWindow", "ptr", cpComboHwnd, "int")
                try DllCall("user32\SetWindowLongPtr", "ptr", cpComboHwnd,
                    "int", -4, "ptr", cpComboOriginalProc, "ptr")
        }
        CPStudyThemedComboHwnds.Clear()
    }
    if IsSet(CPStudyThemedListHwnds) && IsObject(CPStudyThemedListHwnds) {
        for cpListHwnd, cpListOriginalProc in CPStudyThemedListHwnds {
            if cpListHwnd && cpListOriginalProc
                && DllCall("user32\IsWindow", "ptr", cpListHwnd, "int")
                try DllCall("user32\SetWindowLongPtr", "ptr", cpListHwnd,
                    "int", -4, "ptr", cpListOriginalProc, "ptr")
        }
        CPStudyThemedListHwnds.Clear()
    }
    if IsSet(CPStudyTransparentHwnds) && IsObject(CPStudyTransparentHwnds) {
        for cpOverlayHwnd, cpOverlayData in CPStudyTransparentHwnds {
            cpOverlayOriginalProc := IsObject(cpOverlayData) && cpOverlayData.Has("proc")
                ? cpOverlayData["proc"] : 0
            if cpOverlayHwnd && cpOverlayOriginalProc
                && DllCall("user32\IsWindow", "ptr", cpOverlayHwnd, "int")
                try DllCall("user32\SetWindowLongPtr", "ptr", cpOverlayHwnd,
                    "int", -4, "ptr", cpOverlayOriginalProc, "ptr")
        }
        CPStudyTransparentHwnds.Clear()
    }
    if IsSet(CPStudyThemedHeaderHwnds) && IsObject(CPStudyThemedHeaderHwnds)
        CPStudyThemedHeaderHwnds.Clear()
    if IsSet(CPStudyVisualOverlays) && IsObject(CPStudyVisualOverlays)
        CPStudyVisualOverlays.Clear()
}

CPShutdownThemeMessages(*) {
    CPUnsubclassStudyControls()
    CPUnregisterThemeMessages()
    CPDestroyThemeBrushes()
}

CPApplyControlPanelTheme(forceRedraw := true) {
    global ui, controlDarkMode, CPTabBarFill, CPFooterFill, CPTabRenderedState, hkConflictText
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
    if (IsSet(CPFooterFill) && CPFooterFill)
        try CPFooterFill.Opt("+Background" cpApplyColors["window"])
    if (IsSet(hkConflictText) && hkConflictText)
        try hkConflictText.SetFont("c" (cpApplyDark ? "FF7B72" : "FF0000"))

    CPTabRenderedState := ""
    CPRenderCustomTabBar(true)
    CPUpdateComboArrowOverlays()
    CPMaintainFooterZOrder()
    if forceRedraw
        try DllCall("user32\RedrawWindow", "ptr", ui.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400)
}

CPOnDarkModeToggle(*) {
    global chkDarkMode, controlDarkMode, iniPath
    controlDarkMode := chkDarkMode.Value ? 1 : 0
    IniWrite(controlDarkMode, iniPath, "cfg_control", "darkMode")
    CPApplyControlPanelTheme()
    CPApplyOpenStudyWindowThemes()
}

CPClampControlPanelOpacity(value) {
    return Max(70, Min(100, Round(value)))
}

CPApplyControlPanelOpacity(persist := false) {
    global ui, slControlOpacity, lblControlOpacityPct
    global controlPanelOpacity, iniPath

    controlPanelOpacity := CPClampControlPanelOpacity(controlPanelOpacity)
    if (IsSet(slControlOpacity) && slControlOpacity) {
        try {
            if (slControlOpacity.Value != controlPanelOpacity)
                slControlOpacity.Value := controlPanelOpacity
        }
    }
    if (IsSet(lblControlOpacityPct) && lblControlOpacityPct)
        try lblControlOpacityPct.Text := controlPanelOpacity "%"

    if (persist)
        IniWrite(controlPanelOpacity, iniPath, "cfg_control", "opacity")

    if (!IsSet(ui) || !ui || !ui.Hwnd)
        return

    targetWindow := "ahk_id " ui.Hwnd
    try {
        if (controlPanelOpacity >= 100)
            WinSetTransparent("Off", targetWindow)
        else
            WinSetTransparent(Round(controlPanelOpacity * 255 / 100), targetWindow)
    }
}

CPOnControlPanelOpacityChange(ctrl, *) {
    global controlPanelOpacity
    controlPanelOpacity := CPClampControlPanelOpacity(ctrl.Value)
    CPApplyControlPanelOpacity(true)
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
        hSmall && DllCall("DestroyIcon","ptr",hSmall),
        0 ; a nonzero OnExit return cancels shutdown and leaves a hidden process
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
gameProfilesDir := appDir "\game_profiles"
try DirCreate(gameProfilesDir)
studyLibraryDir := StudyLibraryConfiguredDirectory()

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
            ; While the control-panel workflow owns the foreground, native
            ; shoulder buttons change tabs instead. Suppress a matching
            ; keyboard-emulated/JoyToKey hotkey so it cannot also request an
            ; explanation in the background.
            Hotkey(newHK, CPExplainLastHotkey)
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

CPSetWindowCloaked(cpPanelHwnd, cpCloaked) {
    if !cpPanelHwnd
        return false
    cpCloakValue := Buffer(4, 0)
    NumPut("int", cpCloaked ? 1 : 0, cpCloakValue, 0)
    try return DllCall(
        "dwmapi\DwmSetWindowAttribute",
        "ptr", cpPanelHwnd,
        "uint", 13, ; DWMWA_CLOAK
        "ptr", cpCloakValue.Ptr,
        "uint", 4,
        "int"
    ) = 0
    return false
}

CPShowControlPanelReady(activate := true) {
    global ui, controlDarkMode
    if !(IsSet(ui) && ui && ui.Hwnd)
        return

    ; Keep the window out of DWM composition until its complete dark/light frame
    ; has been presented once. This prevents the native default (white) surface
    ; and title bar from appearing between Show and the custom-control repaint.
    cpPanelCloaked := CPSetWindowCloaked(ui.Hwnd, true)
    ui.Show("NA")
    CPApplyDarkTitleBar(ui.Hwnd, controlDarkMode)
    try DllCall(
        "user32\RedrawWindow",
        "ptr", ui.Hwnd,
        "ptr", 0,
        "ptr", 0,
        "uint", 0x0001 | 0x0020 | 0x0080 | 0x0100 | 0x0400
        ; INVALIDATE | NOERASE | ALLCHILDREN | UPDATENOW | FRAME
    )
    if cpPanelCloaked {
        try DllCall("dwmapi\DwmFlush")
        CPSetWindowCloaked(ui.Hwnd, false)
    }
    ; Applying transparency while the window is hidden/cloaked is not reliable:
    ; DWM may present the first visible frame as fully opaque and only notice the
    ; layered alpha after the window next moves.  Reapply it after the prepared
    ; window is genuinely visible and uncloaked.
    CPApplyControlPanelOpacity(false)
    if activate
        try WinActivate("ahk_id " ui.Hwnd)
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
            CPShowControlPanelReady(true)
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

CPExplainLastHotkey(*) {
    if CPControllerNavigationTarget()
        return
    SafeCall(ExplainNow)
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
    CPClipScrollableControlsToViewport(false)
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

CPMaintainFooterZOrder() {
    global CPFooterFill, sepAction, btnOv, btnOvClose, btnAudio
    global btnExplainerLaunch, btnExplainerClose, bClose
    global chkTop, chkDarkMode, txtControlOpacity, slControlOpacity, lblControlOpacityPct
    if !(IsSet(CPFooterFill) && CPFooterFill && CPFooterFill.Hwnd)
        return

    static SWP_KEEP_GEOMETRY := 0x0001 | 0x0002 | 0x0010 ; NOSIZE | NOMOVE | NOACTIVATE
    ; The opaque fill sits above scrollable page controls. Footer controls are
    ; then raised above the fill so content can never bleed through this region.
    try DllCall("user32\SetWindowPos", "ptr", CPFooterFill.Hwnd, "ptr", 0
        , "int", 0, "int", 0, "int", 0, "int", 0, "uint", SWP_KEEP_GEOMETRY)
    for footerZCtrl in [sepAction, btnOv, btnOvClose, btnAudio, btnExplainerLaunch
        , btnExplainerClose, bClose, chkTop, chkDarkMode, txtControlOpacity
        , slControlOpacity, lblControlOpacityPct] {
        if (footerZCtrl && footerZCtrl.Hwnd)
            try DllCall("user32\SetWindowPos", "ptr", footerZCtrl.Hwnd, "ptr", 0
                , "int", 0, "int", 0, "int", 0, "int", 0, "uint", SWP_KEEP_GEOMETRY)
    }
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
    CPChangeTabPageAtomically(cpTabIndex, false)
}

CPShowTabScreenSnapshot(*) {
    global ui, controlPanelOpacity
    global CPTabScreenSnapshotHwnd, CPTabScreenSnapshotBitmap

    ; WinSetTransparent turns the control panel into a layered top-level window.
    ; Windows can briefly omit that whole window while native tab pages exchange
    ; visibility. A child cover cannot help because it disappears with its parent.
    ; Capture the final composited pixels from the screen and hold them in a
    ; separate, non-activating owned window until the new page has been painted.
    if (controlPanelOpacity >= 100)
        return false
    if !(IsSet(ui) && ui && ui.Hwnd)
        return false
    if !DllCall("user32\IsWindowVisible", "ptr", ui.Hwnd, "int")
        return false

    cpRect := Buffer(16, 0)
    if !DllCall("user32\GetWindowRect", "ptr", ui.Hwnd, "ptr", cpRect.Ptr, "int")
        return false
    cpX := NumGet(cpRect, 0, "int")
    cpY := NumGet(cpRect, 4, "int")
    cpW := NumGet(cpRect, 8, "int") - cpX
    cpH := NumGet(cpRect, 12, "int") - cpY
    if (cpW < 1 || cpH < 1)
        return false

    cpScreenDc := DllCall("user32\GetDC", "ptr", 0, "ptr")
    if !cpScreenDc
        return false
    cpMemoryDc := DllCall("gdi32\CreateCompatibleDC", "ptr", cpScreenDc, "ptr")
    cpBitmap := cpMemoryDc
        ? DllCall("gdi32\CreateCompatibleBitmap", "ptr", cpScreenDc
            , "int", cpW, "int", cpH, "ptr")
        : 0
    if !cpMemoryDc || !cpBitmap {
        if cpMemoryDc
            DllCall("gdi32\DeleteDC", "ptr", cpMemoryDc)
        DllCall("user32\ReleaseDC", "ptr", 0, "ptr", cpScreenDc)
        return false
    }

    cpOldObject := DllCall("gdi32\SelectObject", "ptr", cpMemoryDc, "ptr", cpBitmap, "ptr")
    cpCaptured := DllCall("gdi32\BitBlt", "ptr", cpMemoryDc
        , "int", 0, "int", 0, "int", cpW, "int", cpH
        , "ptr", cpScreenDc, "int", cpX, "int", cpY
        , "uint", 0x40CC0020, "int") ; CAPTUREBLT | SRCCOPY
    if cpOldObject
        DllCall("gdi32\SelectObject", "ptr", cpMemoryDc, "ptr", cpOldObject, "ptr")
    DllCall("gdi32\DeleteDC", "ptr", cpMemoryDc)
    DllCall("user32\ReleaseDC", "ptr", 0, "ptr", cpScreenDc)
    if !cpCaptured {
        DllCall("gdi32\DeleteObject", "ptr", cpBitmap)
        return false
    }

    cpSnapshotHwnd := CPTabScreenSnapshotHwnd
    if !(cpSnapshotHwnd && DllCall("user32\IsWindow", "ptr", cpSnapshotHwnd, "int")) {
        cpModule := DllCall("kernel32\GetModuleHandleW", "ptr", 0, "ptr")
        cpSnapshotHwnd := DllCall("user32\CreateWindowExW"
            , "uint", 0x080000A0 ; WS_EX_NOACTIVATE|TOOLWINDOW|TRANSPARENT
            , "str", "Static", "ptr", 0
            , "uint", 0x8800004E ; WS_POPUP|DISABLED|SS_BITMAP|SS_REALSIZECONTROL
            , "int", cpX, "int", cpY, "int", cpW, "int", cpH
            , "ptr", ui.Hwnd, "ptr", 0, "ptr", cpModule, "ptr", 0, "ptr")
        if !cpSnapshotHwnd {
            DllCall("gdi32\DeleteObject", "ptr", cpBitmap)
            return false
        }
        CPTabScreenSnapshotHwnd := cpSnapshotHwnd
    }

    cpPreviousBitmap := DllCall("user32\SendMessageW", "ptr", cpSnapshotHwnd
        , "uint", 0x0172, "ptr", 0, "ptr", cpBitmap, "ptr") ; STM_SETIMAGE
    if cpPreviousBitmap
        DllCall("gdi32\DeleteObject", "ptr", cpPreviousBitmap)
    CPTabScreenSnapshotBitmap := cpBitmap

    DllCall("user32\SetWindowPos", "ptr", cpSnapshotHwnd, "ptr", 0
        , "int", cpX, "int", cpY, "int", cpW, "int", cpH
        , "uint", 0x0010 | 0x0040) ; NOACTIVATE | SHOWWINDOW
    DllCall("user32\RedrawWindow", "ptr", cpSnapshotHwnd, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0100) ; INVALIDATE | UPDATENOW
    try DllCall("dwmapi\DwmFlush")
    return true
}

CPHideTabScreenSnapshot(*) {
    global CPTabScreenSnapshotHwnd, CPTabScreenSnapshotBitmap

    cpSnapshotHwnd := CPTabScreenSnapshotHwnd
    if (cpSnapshotHwnd && DllCall("user32\IsWindow", "ptr", cpSnapshotHwnd, "int")) {
        DllCall("user32\ShowWindow", "ptr", cpSnapshotHwnd, "int", 0) ; SW_HIDE
        cpBitmap := DllCall("user32\SendMessageW", "ptr", cpSnapshotHwnd
            , "uint", 0x0172, "ptr", 0, "ptr", 0, "ptr") ; STM_SETIMAGE
        if cpBitmap
            DllCall("gdi32\DeleteObject", "ptr", cpBitmap)
    } else if CPTabScreenSnapshotBitmap {
        DllCall("gdi32\DeleteObject", "ptr", CPTabScreenSnapshotBitmap)
    }
    CPTabScreenSnapshotBitmap := 0
}

CPChangeTabPageAtomically(cpTabIndex, cpShowNavFocus := unset) {
    global ui, tab, CPTabBarHasNavFocus
    if !(IsSet(ui) && ui && ui.Hwnd && IsSet(tab) && tab && tab.Hwnd)
        return false

    if IsSet(cpShowNavFocus)
        CPTabBarHasNavFocus := cpShowNavFocus ? true : false

    cpCurrentTab := 0
    try cpCurrentTab := tab.Value
    if (cpCurrentTab = cpTabIndex) {
        CPRenderCustomTabBar()
        return false
    }

    cpSnapshotVisible := CPShowTabScreenSnapshot()
    cpChanged := false
    try {
        try tab.Value := cpTabIndex
        try cpChanged := tab.Value = cpTabIndex
        if cpChanged {
            CPCanvasScrollTo(0, 0, false, false, false)
            CPRedrawActiveTabPage()
            CPUpdateComboArrowOverlays()
            CPRenderCustomTabBar(true)
            ; Commit the completed parent frame while the independent copy of the
            ; previous frame still masks any layered-window composition gap.
            if cpSnapshotVisible
                try DllCall("dwmapi\DwmFlush")
        }
    } finally {
        if cpSnapshotVisible
            CPHideTabScreenSnapshot()
    }
    return cpChanged
}

CPRedrawActiveTabPage(*) {
    global ui, tab
    if !(IsSet(ui) && ui && IsSet(tab) && tab)
        return
    try hwnd := ui.Hwnd
    catch
        return
    if !hwnd
        return

    ; AHK shows/hides the native tab-page controls after the selected page
    ; changes. Invalidate every child once after that visibility update so a
    ; layered panel does not leave the new controls blank until mouse-over.
    try tab.Redraw()
    try DllCall("user32\RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0020 | 0x0080 | 0x0100) ; INVALIDATE|NOERASE|ALLCHILDREN|UPDATENOW
}

CPMouseTabClick(*) {
    global ui, CPTabVisiblePages, CPTabButtons, CPFocusVisualNavHwnd, eCapMax
    if !(IsSet(ui) && ui && ui.Hwnd && IsSet(CPTabButtons) && IsObject(CPTabButtons))
        return

    ; Use the actual child window beneath the pointer. Coordinate comparisons are
    ; unsafe here because MouseGetPos and GetWindowRect can use different origins.
    MouseGetPos &cpMouseX, &cpMouseY, &cpMouseWindowHwnd, &cpMouseControlHwnd, 2
    if (cpMouseWindowHwnd != ui.Hwnd || !cpMouseControlHwnd)
        return

    ; A direct click on the Maximum PNG field means the user intends ordinary
    ; mouse/keyboard editing, not controller-navigation adjustment. Release the
    ; navigation ownership so typed digits remain available after that click.
    if (IsSet(eCapMax) && IsObject(eCapMax) && cpMouseControlHwnd = eCapMax.Hwnd) {
        if CPMaxPngAdjustActive()
            CPMaxPngAdjustFinish(true)
        CPFocusVisualNavHwnd := 0
        SetTimer(UpdateCPFocusRing, -1)
        return
    }

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
    ; Keep it single-line so hidden native labels can never wrap below the cover.
    try tab.Opt("-Wrap +0x04000000") ; WS_CLIPSIBLINGS keeps native painting below the custom row.
    CPTabBarFill := ui.Add("Text", "x0 y0 w1 h1 Hidden BackgroundF0F0F0 +0x100 +0x04000000")
    CPRegisterCanvasFixedControl(CPTabBarFill, true, true)
    for cpTabPage in CPTabVisiblePages {
        cpTabLabel := tabNames[cpTabPage]
        cpTabCtrl := ui.Add("Text", "x0 y0 h30 Hidden Center Border +0x100 +0x200 +0x04000000 BackgroundF3F3F3 c202020", cpTabLabel)
        CPRegisterCanvasFixedControl(cpTabCtrl, true, true)
        cpTabCtrl.GetPos(,, &cpTabNaturalW)
        CPTabNaturalWidths.Push(Max(46, cpTabNaturalW + 18))
        cpTabCtrl.Cursor := "Hand"
        ; Mouse tab selection is handled centrally by CPMouseTabClick. Registering
        ; an additional Click callback here made a single click run two complete
        ; page-update paths, which was especially visible on a layered window.
        CPTabButtons.Push(cpTabCtrl)
    }
    CPConfigurePreferredPanelWidth()
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
        CPUpdateFontSizeControllerHint()
        CPUpdateMaxPngControllerHint()
        CPRestoreFocusVisual()
        CPSetTabFocusIndicator(false)
        return
    }

    hwnd := CPFocusRingTargetHwnd(CPFocusedHwnd())
    if !hwnd {
        CPUpdateFontSizeControllerHint()
        CPUpdateMaxPngControllerHint()
        CPRestoreFocusVisual()
        CPSetTabFocusIndicator(false)
        return
    }
    CPUpdateFontSizeControllerHint(hwnd)
    CPUpdateMaxPngControllerHint(hwnd)
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
    cb := CallbackCreate(CPEnumFocusableProc)
    try DllCall("user32\EnumChildWindows", "ptr", ui.Hwnd, "ptr", cb, "ptr", 0)
    finally CallbackFree(cb)
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

CPNavControlsViewTab(dir) {
    global rbControlsKeyboard, rbControlsController
    global hkEdits, CPControllerBindingEdits, hotkeyActions
    if (dir != "Left" && dir != "Right" && dir != "Down")
        return false
    if !(IsSet(rbControlsKeyboard) && rbControlsKeyboard
        && IsSet(rbControlsController) && rbControlsController)
        return false

    focusedHwnd := CPFocusedHwnd()
    if (focusedHwnd != rbControlsKeyboard.Hwnd
        && focusedHwnd != rbControlsController.Hwnd)
        return false

    if (dir = "Down") {
        inputMap := focusedHwnd = rbControlsKeyboard.Hwnd
            ? hkEdits : CPControllerBindingEdits
        for actionKey in hotkeyActions {
            if !inputMap.Has(actionKey)
                continue
            inputCtrl := inputMap[actionKey]
            if (inputCtrl && inputCtrl.Hwnd
                && DllCall("user32\IsWindowVisible", "ptr", inputCtrl.Hwnd, "int")) {
                CPSetFocusHwnd(inputCtrl.Hwnd)
                return true
            }
        }
        ; Do not let a missing view control unexpectedly drop focus into the
        ; global footer; remain on the active view tab instead.
        return true
    }

    ; Windows normally moves within a radio group with Left/Right, but the
    ; control-panel navigation hotkeys intercept those keys first. Switch the
    ; view explicitly so keyboard arrows and native controller D-pad input
    ; follow the same predictable two-tab behavior.
    if (focusedHwnd = rbControlsKeyboard.Hwnd && dir = "Right") {
        CPSetControlsView("controller", true)
        CPSetFocusHwnd(rbControlsController.Hwnd)
    } else if (focusedHwnd = rbControlsController.Hwnd && dir = "Left") {
        CPSetControlsView("keyboard", true)
        CPSetFocusHwnd(rbControlsKeyboard.Hwnd)
    }
    ; Consume both boundary directions so focus cannot fall out of this row.
    return true
}

CPFontSizeAdjustActive() {
    global CPFontSizeAdjustState
    return IsObject(CPFontSizeAdjustState)
        && CPFontSizeAdjustState.Has("active")
        && CPFontSizeAdjustState["active"]
}

CPFontSizeAdjustKindForHwnd(hwnd) {
    global edFSize, udFSize, edFSize_EW, udFSize_EW
    if !hwnd
        return ""
    try {
        if ((IsSet(edFSize) && IsObject(edFSize) && hwnd = edFSize.Hwnd)
         || (IsSet(udFSize) && IsObject(udFSize) && hwnd = udFSize.Hwnd))
            return "translator"
        if ((IsSet(edFSize_EW) && IsObject(edFSize_EW) && hwnd = edFSize_EW.Hwnd)
         || (IsSet(udFSize_EW) && IsObject(udFSize_EW) && hwnd = udFSize_EW.Hwnd))
            return "explainer"
    }
    return ""
}

CPUpdateFontSizeControllerHint(hwnd := 0) {
    global txtFontSizeHint, txtFontSizeHint_EW, CPFocusVisualNavHwnd, controlDarkMode
    static cpLastHintState := ""

    cpSizeKind := CPFontSizeAdjustKindForHwnd(hwnd)
    cpHasNavFocus := (cpSizeKind != "")
        && IsSet(CPFocusVisualNavHwnd) && (CPFocusVisualNavHwnd = hwnd)
    cpSizeActive := cpHasNavFocus && CPFontSizeAdjustActive()
    cpHintState := (cpHasNavFocus ? cpSizeKind : "hidden") "|" cpSizeActive "|" (controlDarkMode ? 1 : 0)
    if (cpHintState = cpLastHintState)
        return
    cpLastHintState := cpHintState

    if IsSet(txtFontSizeHint)
        try txtFontSizeHint.Visible := false
    if IsSet(txtFontSizeHint_EW)
        try txtFontSizeHint_EW.Visible := false
    if !cpHasNavFocus
        return

    cpHintCtrl := (cpSizeKind = "translator") ? txtFontSizeHint : txtFontSizeHint_EW
    cpHintColors := CPPalette(controlDarkMode)
    cpHintCtrl.Text := cpSizeActive ? "< A >" : "A"
    cpHintCtrl.Opt("+Background" (cpSizeActive ? cpHintColors["accentFocus"] : cpHintColors["focus"]))
    cpHintCtrl.SetFont("Bold c" (cpSizeActive ? cpHintColors["accentText"]
        : (controlDarkMode ? cpHintColors["accentText"] : cpHintColors["accentFocus"])))
    cpHintCtrl.Visible := true
    try cpHintCtrl.Redraw()
}

CPFontSizeAdjustSyncControls(cpSizeState, cpSizeValue) {
    global CPFontSizeAdjustSyncing
    CPFontSizeAdjustSyncing := true
    try {
        cpSizeState["edit"].Value := cpSizeValue
        cpSizeState["spinner"].Value := cpSizeValue
    } finally {
        CPFontSizeAdjustSyncing := false
    }
}

CPFontSizeAdjustStart(hwnd) {
    global CPFontSizeAdjustState
    global edFSize, udFSize, fontSize
    global edFSize_EW, udFSize_EW, fontSize_EW

    cpSizeKind := CPFontSizeAdjustKindForHwnd(hwnd)
    if (cpSizeKind = "")
        return false
    if CPMaxPngAdjustActive()
        CPMaxPngAdjustFinish(true)

    if (cpSizeKind = "translator") {
        cpSizeEdit := edFSize
        cpSizeSpinner := udFSize
        cpSizeOriginal := Max(6, Min(128, Integer(fontSize)))
        cpSizeMin := 6
        cpSizeMax := 128
    } else {
        cpSizeEdit := edFSize_EW
        cpSizeSpinner := udFSize_EW
        cpSizeOriginal := Max(6, Min(200, Integer(fontSize_EW)))
        cpSizeMin := 6
        cpSizeMax := 200
    }
    cpSizeValue := cpSizeOriginal
    try {
        cpSizePendingText := Trim(cpSizeEdit.Value)
        if (cpSizePendingText != "")
            cpSizeValue := Max(cpSizeMin, Min(cpSizeMax, Integer(cpSizePendingText)))
    }

    CPFontSizeAdjustState := Map(
        "active", true,
        "kind", cpSizeKind,
        "edit", cpSizeEdit,
        "spinner", cpSizeSpinner,
        "original", cpSizeOriginal,
        "value", cpSizeValue,
        "min", cpSizeMin,
        "max", cpSizeMax,
        "step", 1)
    CPFontSizeAdjustSyncControls(CPFontSizeAdjustState, cpSizeValue)
    CPSetFocusHwnd(cpSizeEdit.Hwnd)
    UpdateCPFocusRing()
    return true
}

CPFontSizeAdjustHandleDirection(cpSizeDirection) {
    global CPFontSizeAdjustState, fontSize, fontSize_EW
    if !CPFontSizeAdjustActive()
        return false

    cpSizeFocusedKind := CPFontSizeAdjustKindForHwnd(CPFocusedHwnd())
    if (cpSizeFocusedKind != CPFontSizeAdjustState["kind"]) {
        CPFontSizeAdjustFinish(true)
        return false
    }
    cpSizeStep := CPFontSizeAdjustState.Has("step") ? CPFontSizeAdjustState["step"] : 1
    if (cpSizeDirection = "Up" || cpSizeDirection = "Right")
        cpSizeDelta := cpSizeStep
    else if (cpSizeDirection = "Down" || cpSizeDirection = "Left")
        cpSizeDelta := -cpSizeStep
    else
        return true

    cpSizeValue := Max(CPFontSizeAdjustState["min"]
        , Min(CPFontSizeAdjustState["max"], CPFontSizeAdjustState["value"] + cpSizeDelta))
    if (cpSizeValue = CPFontSizeAdjustState["value"])
        return true

    CPFontSizeAdjustState["value"] := cpSizeValue
    if (CPFontSizeAdjustState["kind"] = "translator")
        fontSize := cpSizeValue
    else
        fontSize_EW := cpSizeValue
    CPFontSizeAdjustSyncControls(CPFontSizeAdjustState, cpSizeValue)
    SendOverlayTheme()
    return true
}

CPFontSizeAdjustFinish(cpSizeSave := true) {
    global CPFontSizeAdjustState, fontSize, fontSize_EW
    if !CPFontSizeAdjustActive()
        return false

    cpSizeState := CPFontSizeAdjustState
    cpSizeValue := cpSizeSave ? cpSizeState["value"] : cpSizeState["original"]
    if (cpSizeState["kind"] = "translator")
        fontSize := cpSizeValue
    else
        fontSize_EW := cpSizeValue

    CPFontSizeAdjustState := Map("active", false)
    CPFontSizeAdjustSyncControls(cpSizeState, cpSizeValue)
    if cpSizeSave {
        SaveAll()
        DbgCP("Font size controller adjustment saved: " cpSizeState["kind"] " -> " cpSizeValue)
    } else {
        DbgCP("Font size controller adjustment canceled: " cpSizeState["kind"] " -> " cpSizeValue)
    }
    SendOverlayTheme()
    CPSetFocusHwnd(cpSizeState["edit"].Hwnd)
    UpdateCPFocusRing()
    return true
}

CPFontSizeAdjustCancel() {
    return CPFontSizeAdjustFinish(false)
}

CPMaxPngAdjustActive() {
    global CPMaxPngAdjustState
    return IsObject(CPMaxPngAdjustState)
        && CPMaxPngAdjustState.Has("active")
        && CPMaxPngAdjustState["active"]
}

CPMaxPngAdjustForHwnd(hwnd) {
    global eCapMax
    if !hwnd
        return false
    try return IsSet(eCapMax) && IsObject(eCapMax) && hwnd = eCapMax.Hwnd
    return false
}

CPUpdateMaxPngControllerHint(hwnd := 0) {
    global txtCapMaxHint, CPFocusVisualNavHwnd, controlDarkMode
    static cpLastMaxPngHintState := ""

    cpHasNavFocus := CPMaxPngAdjustForHwnd(hwnd)
        && IsSet(CPFocusVisualNavHwnd) && (CPFocusVisualNavHwnd = hwnd)
    cpAdjustActive := cpHasNavFocus && CPMaxPngAdjustActive()
    cpHintState := (cpHasNavFocus ? "shown" : "hidden") "|" cpAdjustActive "|" (controlDarkMode ? 1 : 0)
    if (cpHintState = cpLastMaxPngHintState)
        return
    cpLastMaxPngHintState := cpHintState

    if IsSet(txtCapMaxHint)
        try txtCapMaxHint.Visible := false
    if !cpHasNavFocus
        return

    cpHintColors := CPPalette(controlDarkMode)
    txtCapMaxHint.Text := cpAdjustActive ? "< A >" : "A"
    txtCapMaxHint.Opt("+Background" (cpAdjustActive ? cpHintColors["accentFocus"] : cpHintColors["focus"]))
    txtCapMaxHint.SetFont("Bold c" (cpAdjustActive ? cpHintColors["accentText"]
        : (controlDarkMode ? cpHintColors["accentText"] : cpHintColors["accentFocus"])))
    txtCapMaxHint.Visible := true
    try txtCapMaxHint.Redraw()
}

CPMaxPngAdjustSyncValue(cpMaxPngValue) {
    global CPMaxPngAdjustSyncing, eCapMax
    CPMaxPngAdjustSyncing := true
    try eCapMax.Value := cpMaxPngValue
    finally CPMaxPngAdjustSyncing := false
}

CPMaxPngAdjustStart(hwnd) {
    global CPMaxPngAdjustState, capMaxKB, eCapMax
    if !CPMaxPngAdjustForHwnd(hwnd)
        return false
    if CPFontSizeAdjustActive()
        CPFontSizeAdjustFinish(true)

    cpMaxPngOriginal := Max(100, Min(10000, Integer(capMaxKB)))
    cpMaxPngValue := cpMaxPngOriginal
    try {
        cpPendingText := Trim(eCapMax.Value)
        if (cpPendingText != "")
            cpMaxPngValue := Max(100, Min(10000, Integer(cpPendingText)))
    }

    CPMaxPngAdjustState := Map(
        "active", true,
        "original", cpMaxPngOriginal,
        "value", cpMaxPngValue,
        "min", 100,
        "max", 10000,
        "stepHorizontal", 100,
        "stepVertical", 500)
    CPMaxPngAdjustSyncValue(cpMaxPngValue)
    CPSetFocusHwnd(eCapMax.Hwnd)
    UpdateCPFocusRing()
    return true
}

CPMaxPngAdjustHandleDirection(cpMaxPngDirection) {
    global CPMaxPngAdjustState
    if !CPMaxPngAdjustActive()
        return false

    if !CPMaxPngAdjustForHwnd(CPFocusedHwnd()) {
        CPMaxPngAdjustFinish(true)
        return false
    }
    cpMaxPngHorizontalStep := CPMaxPngAdjustState.Has("stepHorizontal")
        ? CPMaxPngAdjustState["stepHorizontal"] : 100
    cpMaxPngVerticalStep := CPMaxPngAdjustState.Has("stepVertical")
        ? CPMaxPngAdjustState["stepVertical"] : 500
    switch cpMaxPngDirection {
        case "Right": cpMaxPngDelta := cpMaxPngHorizontalStep
        case "Left": cpMaxPngDelta := -cpMaxPngHorizontalStep
        case "Up": cpMaxPngDelta := cpMaxPngVerticalStep
        case "Down": cpMaxPngDelta := -cpMaxPngVerticalStep
        default: return true
    }

    cpMaxPngValue := Max(CPMaxPngAdjustState["min"]
        , Min(CPMaxPngAdjustState["max"], CPMaxPngAdjustState["value"] + cpMaxPngDelta))
    if (cpMaxPngValue = CPMaxPngAdjustState["value"])
        return true

    CPMaxPngAdjustState["value"] := cpMaxPngValue
    CPMaxPngAdjustSyncValue(cpMaxPngValue)
    return true
}

CPMaxPngAdjustFinish(cpMaxPngSave := true) {
    global CPMaxPngAdjustState, capMaxKB, eCapMax
    if !CPMaxPngAdjustActive()
        return false

    cpMaxPngState := CPMaxPngAdjustState
    cpMaxPngValue := cpMaxPngSave ? cpMaxPngState["value"] : cpMaxPngState["original"]
    capMaxKB := cpMaxPngValue
    CPMaxPngAdjustState := Map("active", false)
    CPMaxPngAdjustSyncValue(cpMaxPngValue)
    if cpMaxPngSave {
        SaveAll()
        DbgCP("Max PNG size controller adjustment saved: " cpMaxPngValue " KB")
    } else {
        DbgCP("Max PNG size controller adjustment canceled: " cpMaxPngValue " KB")
    }
    CPSetFocusHwnd(eCapMax.Hwnd)
    UpdateCPFocusRing()
    return true
}

CPMaxPngAdjustCancel() {
    return CPMaxPngAdjustFinish(false)
}

CPNavMove(dir, *) {
    global eCapMax, btnCapPick

    if CPFontSizeAdjustHandleDirection(dir)
        return
    if CPMaxPngAdjustHandleDirection(dir)
        return

    if CPForwardIfComboOpen(dir)
        return

    if CPNavControlsViewTab(dir)
        return

    curHwnd := CPFocusedHwnd()
    if (dir = "Down") {
        try {
            if (curHwnd = eCapMax.Hwnd) {
                CPSetFocusHwnd(btnCapPick.Hwnd)
                return
            }
        }
    }
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

CPControllerKeyboardMirrorActive(command) {
    ; JoyToKey can mirror the same physical controller press as an arrow,
    ; Enter, Space, or Escape key. Native controller navigation owns that
    ; press while the corresponding physical control is held, so discard only
    ; the mirrored keyboard event and leave genuine keyboard input untouched.
    global CPControllerLastNativeNavigationAt
    static cpMirrorGraceUntil := Map()
    static cpMirrorReleaseGraceMs := 180
    static cpNativeActionGraceMs := 650

    cpMirrorNow := A_TickCount
    if (CPControllerLastNativeNavigationAt.Has(command)
     && cpMirrorNow - CPControllerLastNativeNavigationAt[command] <= cpNativeActionGraceMs)
        return true

    snapshot := CPControllerReadSnapshot()
    if !snapshot["connected"]
        return false

    cpMirrorState := CPControllerNavigationState(snapshot)
    if !cpMirrorState.Has(command)
        return false

    if cpMirrorState[command] {
        cpMirrorGraceUntil[command] := cpMirrorNow + cpMirrorReleaseGraceMs
        return true
    }
    return cpMirrorGraceUntil.Has(command)
        && cpMirrorNow <= cpMirrorGraceUntil[command]
}

CPKeyboardSliderRepeatStop(*) {
    global CPKeyboardSliderRepeatState
    try SetTimer(CPKeyboardSliderRepeatTick, 0)
    CPKeyboardSliderRepeatState := Map("active", false)
}

CPKeyboardSliderRepeatRelease(direction, *) {
    global CPKeyboardSliderRepeatState
    if (CPKeyboardSliderRepeatState.Has("active")
     && CPKeyboardSliderRepeatState["active"]
     && CPKeyboardSliderRepeatState.Has("direction")
     && CPKeyboardSliderRepeatState["direction"] = direction)
        CPKeyboardSliderRepeatStop()
}

CPKeyboardSliderRepeatStart(direction, stepCallback) {
    global CPKeyboardSliderRepeatState
    cpKeyboardRepeatTarget := CPControllerNavigationTarget()
    if (!cpKeyboardRepeatTarget
     || !CPControllerAcceleratedSliderRepeatActive(cpKeyboardRepeatTarget, direction))
        return false

    if (CPKeyboardSliderRepeatState.Has("active")
     && CPKeyboardSliderRepeatState["active"]) {
        if (CPKeyboardSliderRepeatState.Has("direction")
         && CPKeyboardSliderRepeatState["direction"] = direction
         && CPKeyboardSliderRepeatState.Has("targetHwnd")
         && CPKeyboardSliderRepeatState["targetHwnd"] = cpKeyboardRepeatTarget)
            return true
        CPKeyboardSliderRepeatStop()
    }

    try stepCallback.Call()
    catch {
        return false
    }
    cpKeyboardRepeatNow := A_TickCount
    CPKeyboardSliderRepeatState := Map(
        "active", true,
        "direction", direction,
        "targetHwnd", cpKeyboardRepeatTarget,
        "step", stepCallback,
        "heldSince", cpKeyboardRepeatNow,
        "nextAt", cpKeyboardRepeatNow + 280)
    SetTimer(CPKeyboardSliderRepeatTick, 15)
    return true
}

CPKeyboardSliderRepeatTick(*) {
    global CPKeyboardSliderRepeatState
    cpKeyboardRepeatState := CPKeyboardSliderRepeatState
    if !(cpKeyboardRepeatState.Has("active") && cpKeyboardRepeatState["active"])
        return CPKeyboardSliderRepeatStop()

    cpKeyboardRepeatTarget := CPControllerNavigationTarget()
    if (!cpKeyboardRepeatTarget
     || cpKeyboardRepeatTarget != cpKeyboardRepeatState["targetHwnd"]
     || !CPControllerAcceleratedSliderRepeatActive(cpKeyboardRepeatTarget
        , cpKeyboardRepeatState["direction"]))
        return CPKeyboardSliderRepeatStop()

    cpKeyboardRepeatNow := A_TickCount
    if (cpKeyboardRepeatNow < cpKeyboardRepeatState["nextAt"])
        return

    cpKeyboardHeldMs := cpKeyboardRepeatNow - cpKeyboardRepeatState["heldSince"]
    cpKeyboardRepeatSteps := cpKeyboardHeldMs >= 1100 ? 4 : (cpKeyboardHeldMs >= 650 ? 2 : 1)
    try {
        Loop cpKeyboardRepeatSteps
            cpKeyboardRepeatState["step"].Call()
    } catch {
        return CPKeyboardSliderRepeatStop()
    }
    CPKeyboardSliderRepeatState["nextAt"] := cpKeyboardRepeatNow + 55
}

CPNavUp(*) {
    if CPControllerKeyboardMirrorActive("Up")
        return
    CPNavMove("Up")
}

CPNavDown(*) {
    if CPControllerKeyboardMirrorActive("Down")
        return
    CPNavMove("Down")
}

CPNavLeft(*) {
    if CPControllerKeyboardMirrorActive("Left")
        return
    if CPKeyboardSliderRepeatStart("Left", CPNavMove.Bind("Left"))
        return
    CPNavMove("Left")
}

CPNavRight(*) {
    if CPControllerKeyboardMirrorActive("Right")
        return
    if CPKeyboardSliderRepeatStart("Right", CPNavMove.Bind("Right"))
        return
    CPNavMove("Right")
}

CPNavActivate(keyName := "Enter", *) {
    if CPFontSizeAdjustActive() {
        CPFontSizeAdjustFinish(true)
        return
    }
    if CPMaxPngAdjustActive() {
        CPMaxPngAdjustFinish(true)
        return
    }
    hwnd := CPFocusedHwnd()
    if (keyName = "Enter" && CPFontSizeAdjustStart(hwnd))
        return
    if (keyName = "Enter" && CPMaxPngAdjustStart(hwnd))
        return
    if (keyName = "Enter" && CPIsColorSwatchControl(hwnd)) {
        ; Let the controller polling callback finish before the modal editor opens.
        SetTimer(CPAdjustColorSwatchWithController.Bind(hwnd), -1)
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
    if CPControllerKeyboardMirrorActive("Activate")
        return
    CPNavActivate("Enter")
}

CPNavSpace(*) {
    if CPControllerKeyboardMirrorActive("Activate")
        return
    CPNavActivate("Space")
}

CPNavCancelCurrent() {
    if CPFontSizeAdjustCancel()
        return true
    if CPMaxPngAdjustCancel()
        return true
    cpNavFocusedHwnd := CPFocusedHwnd()
    if (cpNavFocusedHwnd && CPHwndIsCombo(cpNavFocusedHwnd)
     && CPComboDropped(cpNavFocusedHwnd)) {
        CPShowCombo(cpNavFocusedHwnd, false)
        return true
    }
    return false
}

CPNavEscape(*) {
    if CPControllerKeyboardMirrorActive("Cancel")
        return
    CPNavCancelCurrent()
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
    try tab.Focus()
    CPFocusVisualNavHwnd := tab.Hwnd
    CPChangeTabPageAtomically(CPTabVisiblePages[nextVisibleIndex], true)
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
    try Hotkey("$Right Up", CPKeyboardSliderRepeatRelease.Bind("Right"), "On")
    try Hotkey("$Left Up", CPKeyboardSliderRepeatRelease.Bind("Left"), "On")
    try Hotkey("$Enter", CPNavEnter, "On")
    try Hotkey("$NumpadEnter", CPNavEnter, "On")
    try Hotkey("$Space", CPNavSpace, "On")
    try Hotkey("$PgUp", CPNavPrevTab, "On")
    try Hotkey("$PgDn", CPNavNextTab, "On")
    try Hotkey("~LButton", CPMouseTabClick, "On")
    try Hotkey("$Esc", CPNavEscape, "On")
    HotIfWinActive()

    ; Child controls such as combo boxes can consume WM_MOUSEWHEEL before the
    ; parent GUI sees it. A cursor-scoped hotkey keeps canvas scrolling available
    ; immediately, while leaving open dropdowns and the opacity slider alone.
    HotIf(CPMouseWheelOverCanvas)
    try Hotkey("$WheelUp", CPMouseWheelHotkey.Bind(-1), "On")
    try Hotkey("$WheelDown", CPMouseWheelHotkey.Bind(1), "On")
    HotIf()

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
defUseTerminologyOverrides := 1

; NEW: default prompt profile name
defPromptProfile := "default_en"
; EXPLAIN: default prompt profile
defExplainPromptProfile := "default_en"
; Advanced override for showing the screenshot model response unchanged.
defDirectModelOutput := 0

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
global gAudioTestPid := 0
global gAudioTestResultPath := ""
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

; Migrate the former per-translation deletion option. Captures now remain
; available for explanations throughout the current session and are cleared
; safely from an exact-path ledger the next time JRPG Translator starts.
__clearShotsMissing := "__MISSING__"
__clearShotsRaw := IniRead(iniPath, "paths", "clearScreenshotsOnStartup", __clearShotsMissing)
if (__clearShotsRaw = __clearShotsMissing) {
    clearScreenshotsOnStartup := Integer(IniRead(iniPath, "paths", "deleteAfterUse", 0)) ? 1 : 0
    IniWrite(clearScreenshotsOnStartup, iniPath, "paths", "clearScreenshotsOnStartup")
} else {
    clearScreenshotsOnStartup := Integer(__clearShotsRaw) ? 1 : 0
}
; Older overlays must never resume the former ten-second deletion timer.
IniWrite(0, iniPath, "paths", "deleteAfterUse")
CleanupScreenshotsFromPriorSession(clearScreenshotsOnStartup)

; --- NEW: Native capture settings (safe defaults) ---
capMaxKB   := Integer(IniRead(iniPath, "capture", "maxKB", 1400))     ; cap file size in KB
capMode    := IniRead(iniPath, "capture", "mode", "region")           ; "region" or "window"
capWinInfo := IniRead(iniPath, "capture", "winTitle", "")             ; window title (fallback)
capRect    := IniRead(iniPath, "capture", "rect", "")                 ; "x,y,w,h" once selected

showPathsTab := Integer(Load("showPathsTab", defShowPathsTab, "cfg")) ? 1 : 0
debugMode := Integer(Load("debugMode", defDebugMode, "cfg"))
directModelOutput := Integer(Load("directModelOutput", defDirectModelOutput, "cfg")) ? 1 : 0
SetDebugMode(debugMode)
controlDarkMode := Integer(Load("darkMode", 0, "cfg_control")) ? 1 : 0
controlPanelOpacity := 100
try controlPanelOpacity := Integer(Load("opacity", 100, "cfg_control"))
controlPanelOpacity := CPClampControlPanelOpacity(controlPanelOpacity)
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

; Screenshot extraction follows the selected prompt name unless the advanced
; direct-output override is enabled.
imgPostproc      := PromptPostprocMode(promptProfile)
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
useTerminologyOverrides := Integer(Load("useTerminologyOverrides", defUseTerminologyOverrides)) ? 1 : 0
EnvSet("USE_TERMINOLOGY_OVERRIDES", useTerminologyOverrides ? "1" : "0")

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
    global iniPath, debugMode, showPathsTab, directModelOutput
    global capMaxKB,capMode,capRect
    global boxBgHex,bdrOutHex,bdrInHex,txtHex
    global fontName,fontSize,fontBold
    global fontBold_EW
    global bdrOutW,bdrInW
    global model_openai_img, model_gemini_img, model_openai_explain, model_gemini_explain
    global model_openai_audio, model_gemini_audio
    global promptProfile, imgPostproc, chkDel, chkTop, chkDarkMode, controlDarkMode, controlPanelOpacity
    global useTerminologyOverrides

    SyncUnifiedWindowAppearance()
    imgPostproc := SyncPromptPostproc(promptProfile)
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
    IniWrite(controlPanelOpacity, iniPath, "cfg_control", "opacity")
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
    IniWrite(directModelOutput, iniPath, "cfg", "directModelOutput")
    IniWrite(useTerminologyOverrides, iniPath, "cfg", "useTerminologyOverrides")
	; Screenshot cleanup happens once on the next application start.
    IniWrite(chkDel.Value ? 1 : 0, iniPath, "paths", "clearScreenshotsOnStartup")
    IniWrite(0, iniPath, "paths", "deleteAfterUse") ; retired compatibility key

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

SetCapMaxKBFromUI(v) {
    global CPMaxPngAdjustSyncing, CPMaxPngAdjustState, CPFocusVisualNavHwnd
    global capMaxKB, eCapMax
    if CPMaxPngAdjustSyncing
        return


    ; When controller/arrow navigation owns this focused edit, raw digits can
    ; be duplicates emitted by JoyToKey (for example D-pad Down -> Numpad2).
    ; Only the explicit A/Enter adjustment path may change the value in this
    ; state. Programmatic adjustment updates are already covered by the syncing
    ; guard above. Mouse clicks and ordinary Tab focus release/bypass ownership.
    cpControllerOwnsMaxPng := false
    try cpControllerOwnsMaxPng := IsSet(eCapMax) && IsObject(eCapMax)
        && CPFocusVisualNavHwnd = eCapMax.Hwnd
        && CPFocusedHwnd() = eCapMax.Hwnd
    if cpControllerOwnsMaxPng {
        cpProtectedMaxPng := capMaxKB
        if CPMaxPngAdjustActive()
            cpProtectedMaxPng := CPMaxPngAdjustState["value"]
        CPMaxPngAdjustSyncValue(cpProtectedMaxPng)
        return
    }
    SetCapMaxKB(v)
}

EnsureOverlayDir(){
    global overlayDir
    if !DirExist(overlayDir)
        DirCreate(overlayDir)
}

CPNormalizeApiSecret(value) {
    value := Trim(value)
    if (StrLen(value) >= 2) {
        firstChar := SubStr(value, 1, 1)
        lastChar := SubStr(value, -1)
        if ((firstChar = Chr(34) && lastChar = Chr(34))
          || (firstChar = "'" && lastChar = "'"))
            value := SubStr(value, 2, StrLen(value) - 2)
    }
    return Trim(value)
}

CPDotEnvValue(cpEnvText, keyName) {
    if (cpEnvText = "")
        return ""
    if RegExMatch(cpEnvText, "im)^\x{FEFF}?\s*" keyName "\s*=\s*(.*)$", &keyMatch)
        return CPNormalizeApiSecret(keyMatch[1])
    return ""
}

CPApiKeyConfigured(provider) {
    global envPath
    provider := StrLower(Trim(provider))
    if (provider = "gemini") {
        keyNames := ["GEMINI_API_KEY", "GOOGLE_API_KEY", "GEMINI_LOCAL_KEY", "GOOGLE_LOCAL_KEY"]
        fileVarName := "GEMINI_API_KEY_FILE"
    } else if (provider = "openai") {
        keyNames := ["OPENAI_API_KEY", "OPENAI_LOCAL_KEY", "OPENAI_API_KEY_LOCAL", "OPENAI_KEY"]
        fileVarName := "OPENAI_API_KEY_FILE"
    } else {
        return true
    }

    cpEnvText := ""
    for candidateEnvPath in [envPath, A_ScriptDir "\.env"] {
        if FileExist(candidateEnvPath) {
            try cpEnvText := FileRead(candidateEnvPath, "UTF-8")
            break
        }
    }

    for keyName in keyNames {
        if (CPNormalizeApiSecret(EnvGet(keyName)) != ""
          || CPDotEnvValue(cpEnvText, keyName) != "")
            return true
    }

    keyFile := CPNormalizeApiSecret(EnvGet(fileVarName))
    if (keyFile = "")
        keyFile := CPDotEnvValue(cpEnvText, fileVarName)
    if (keyFile != "") {
        keyFile := ResolvePath(keyFile)
        if FileExist(keyFile) {
            try if (Trim(FileRead(keyFile, "UTF-8")) != "")
                return true
        }
    }
    return false
}

CPMissingApiKeyText(provider) {
    provider := StrLower(Trim(provider))
    if (provider = "gemini")
        return "Gemini API key missing.`n`nAdd it in the API Keys tab, or set GEMINI_API_KEY (or GOOGLE_API_KEY) in Windows Environment Variables and restart JRPG Translator."
    return "OpenAI API key missing.`n`nAdd it in the API Keys tab, or set OPENAI_API_KEY in Windows Environment Variables and restart JRPG Translator."
}

CPAtomicWriteOverlayMessage(path, text) {
    tmpPath := path ".tmp"
    try {
        if FileExist(tmpPath)
            FileDelete(tmpPath)
        FileAppend(text, tmpPath, "UTF-8")
        FileMove(tmpPath, path, true)
        return true
    } catch as ex {
        try if FileExist(tmpPath)
            FileDelete(tmpPath)
        DbgCP("Could not write overlay message: " ex.Message)
        return false
    }
}

CPShowMissingApiKey(provider, target := "translator") {
    global overlayDir
    title := (target = "explainer") ? "Explainer" : "Translator"
    hwnd := CPOverlayWindowHwnd(title)
    if !hwnd {
        if (target = "explainer")
            LaunchExplainerOverlay()
        else
            LaunchOverlay()
        hwnd := CPOverlayWindowHwnd(title)
    }
    if hwnd
        try ShowWindowNoActivate(hwnd)

    EnsureOverlayDir()
    targetName := (target = "explainer") ? "explainer" : "translator"
    payloadPath := overlayDir "\message." targetName ".txt"
    signalPath := overlayDir "\cmd.show_" targetName "_message"
    if CPAtomicWriteOverlayMessage(payloadPath, CPMissingApiKeyText(provider)) {
        try {
            if FileExist(signalPath)
                FileDelete(signalPath)
            FileAppend("", signalPath, "UTF-8")
        } catch as ex {
            DbgCP("Could not signal overlay message: " ex.Message)
        }
    }
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

CPWriteExplainerTerminalResult(message, requestId) {
    global overlayDir
    EnsureOverlayDir()
    CPAtomicWriteOverlayMessage(overlayDir "\explainer.txt", message)
    CPAtomicWriteOverlayMessage(overlayDir "\explainer.done", requestId)
}

ScreenshotCleanupLedgerPath() {
    return A_ScriptDir "\Settings\screenshot_cleanup_pending.txt"
}

CleanupScreenshotsFromPriorSession(shouldDelete) {
    ledgerPath := ScreenshotCleanupLedgerPath()
    if !FileExist(ledgerPath)
        return

    entries := ""
    try entries := FileRead(ledgerPath, "UTF-8")

    if shouldDelete {
        seen := Map()
        for rawPath in StrSplit(entries, "`n", "`r") {
            screenshotPath := Trim(StrReplace(rawPath, Chr(0xFEFF)))
            if (screenshotPath = "" || seen.Has(StrLower(screenshotPath)))
                continue
            seen[StrLower(screenshotPath)] := true

            ; The ledger contains exact paths written by the capture overlay.
            ; Restrict cleanup to absolute PNG files and never delete folders.
            if !RegExMatch(screenshotPath, 'i)^(?:[A-Z]:\\|\\\\).+\.png$')
                continue
            try {
                if FileExist(screenshotPath) && !InStr(FileGetAttrib(screenshotPath), "D")
                    FileDelete(screenshotPath)
            }
        }
    }

    ; Disabled cleanup also discards the pending list so old captures cannot be
    ; removed unexpectedly if the setting is enabled again much later.
    try FileDelete(ledgerPath)
}

; =========================
; Audio state helper
; =========================
AudioIsRunning(allowRecoveryScan := false) {
    global gPidAudio
    if (gPidAudio && ProcessExist(gPidAudio))
        return true

    if gPidAudio
        gPidAudio := 0

    ; WMI is only needed to adopt a process which this instance did not launch,
    ; such as one left behind by an earlier crash. Normal status polling relies
    ; exclusively on the PID returned by Run().
    if allowRecoveryScan {
        pids := AudioPidsByScript()
        if (pids.Length) {
            gPidAudio := pids[1]
            return true
        }
    }

    return false
}

UpdateStatus(allowRecoveryScan := false){
    global btnAudio
    running := AudioIsRunning(allowRecoveryScan)
    if (IsSet(btnAudio) && IsObject(btnAudio))
        btnAudio.Text := running ? "Audio Translation On" : "Audio Translation Off"
}

_UpdateStatus(){
    UpdateStatus()
}

; === Path-only dirty state & autosave helpers ===
global pathsDirty := false

UpdatePathsDirtyState(*) {
    global pathsDirty, btnSavePaths
    global ePython, eOverlay, eImg, eAudio, eExplain
    global pythonExe, overlayAhk, imgScript, audioScript, explainScript
    if !(IsSet(btnSavePaths) && IsSet(ePython) && IsSet(eOverlay)
      && IsSet(eImg) && IsSet(eAudio) && IsSet(eExplain))
        return

    pathsDirty := (
        ePython.Value != pythonExe
        || eOverlay.Value != overlayAhk
        || eImg.Value != imgScript
        || eAudio.Value != audioScript
        || eExplain.Value != explainScript
    )
    btnSavePaths.Enabled := pathsDirty
    btnSavePaths.Text := pathsDirty ? "Save paths *" : "Save paths"
}

ClearPathsDirty() {
    global pathsDirty, btnSavePaths
    pathsDirty := false
    if IsSet(btnSavePaths) && IsObject(btnSavePaths) {
        btnSavePaths.Enabled := false
        btnSavePaths.Text := "Save paths"
    }
}

EditedPathsAreValid() {
    global ePython, eOverlay, eImg, eAudio, eExplain
    pathSpecs := [
        ["Python executable", ePython.Value],
        ["Overlay executable", eOverlay.Value],
        ["Screenshot translator", eImg.Value],
        ["Audio translator", eAudio.Value],
        ["Explainer", eExplain.Value]
    ]
    missing := ""
    for pathSpec in pathSpecs {
        rawPath := Trim(pathSpec[2])
        if (rawPath = "" || !FileExist(ResolvePath(rawPath)))
            missing .= (missing = "" ? "" : "`n") "- " pathSpec[1] ": " (rawPath = "" ? "(empty)" : rawPath)
    }
    if (missing = "")
        return true
    answer := MsgBox(
        "These entries do not currently point to files:`n`n" missing
        . "`n`nSave them anyway?",
        "Save paths",
        "YesNo Icon!"
    )
    return answer = "Yes"
}

SaveEditedPaths(showToast := true) {
    global pythonExe, overlayAhk, imgScript, audioScript, explainScript
    global ePython, eOverlay, eImg, eAudio, eExplain, iniPath
    if !EditedPathsAreValid()
        return false

    pythonExe := ePython.Value
    overlayAhk := eOverlay.Value
    imgScript := eImg.Value
    audioScript := eAudio.Value
    explainScript := eExplain.Value

    IniWrite(pythonExe, iniPath, "cfg", "pythonExe")
    IniWrite(overlayAhk, iniPath, "cfg", "overlayAhk")
    IniWrite(imgScript, iniPath, "cfg", "imgScript")
    IniWrite(audioScript, iniPath, "cfg", "audioScript")
    IniWrite(explainScript, iniPath, "cfg", "explainScript")
    ClearPathsDirty()
    if showToast
        Toast("Paths saved")
    DbgCP("Edited paths saved")
    return true
}

ConfirmUnsavedPaths() {
    global pathsDirty
    if !pathsDirty
        return true
    answer := MsgBox(
        "The Paths tab contains unsaved edits.",
        "Unsaved paths",
        "YesNoCancel Icon!"
    )
    if (answer = "Cancel")
        return false
    if (answer = "Yes")
        return SaveEditedPaths(false)
    ClearPathsDirty()
    return true
}

ExitControlPanel(*) {
    if !ConfirmUnsavedPaths()
        return
    SetTimer(_UpdateStatus, 0)
    SetTimer(UpdateCPFocusRing, 0)
    SetTimer(UpdateCPActiveTabHighlight, 0)
    SavePanelBounds()
    ExitApp()
}

; Use this for changes that should immediately persist.
AutoPersist(){
    UpdateVars()
    SaveAll()
}

; -------- Browse handlers --------
BrowsePythonExe(*) {
    global pythonExe, ePython, iniPath
    sel := FileSelect(3,, "Select python.exe", "Programs (*.exe)")
    if (sel != "") {
        pythonExe := sel
        ePython.Value := sel
        IniWrite(pythonExe, iniPath, "cfg", "pythonExe")
        UpdatePathsDirtyState()
        DbgCP("BrowsePythonExe -> " sel)
    }
}
BrowseAudioScript(*) {
    global audioScript, eAudio, iniPath
    sel := FileSelect(3,, "Select audio subtitle .py", "Python (*.py)")
    if (sel != "") {
        audioScript := sel
        eAudio.Value := sel
        IniWrite(audioScript, iniPath, "cfg", "audioScript")
        UpdatePathsDirtyState()
        DbgCP("BrowseAudioScript -> " sel)
    }
}
BrowseOverlayAhk(*) {
    global overlayAhk, eOverlay, iniPath
    sel := FileSelect(3,, "Select overlay .exe", "AutoHotkey (*.exe)")
    if (sel != "") {
        overlayAhk := sel
        eOverlay.Value := sel
        IniWrite(overlayAhk, iniPath, "cfg", "overlayAhk")
        UpdatePathsDirtyState()
        DbgCP("BrowseOverlayAhk -> " sel)
    }
}
BrowseImageScript(*) {
    global imgScript, eImg, iniPath
    sel := FileSelect(3,, "Select image translator .py", "Python (*.py)")
    if (sel != "") {
        imgScript := sel
        eImg.Value := sel
        IniWrite(imgScript, iniPath, "cfg", "imgScript")
        UpdatePathsDirtyState()
        DbgCP("BrowseImageScript -> " sel)
    }
}

BrowseCaptureDir(*) {
    global captureDir
    sel := DirSelect(, , "Select screenshot folder")
    if (sel != "")
        captureDir := sel, SaveAll(), Repaint(), DbgCP("BrowseCaptureDir -> " sel)
}

BrowseExplainScript(*) {
    global explainScript, eExplain, iniPath
    sel := FileSelect(3,, "Select explainer .py", "Python (*.py)")
    if (sel != "") {
        explainScript := sel
        eExplain.Value := sel
        IniWrite(explainScript, iniPath, "cfg", "explainScript")
        UpdatePathsDirtyState()
        DbgCP("BrowseExplainScript -> " sel)
    }
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
    global __OldPickSeq := IniRead(iniPath, "capture", "pickSeq", "")
    global __CapWatchActive := true


    ; hide the Control Panel now
    try ui.Hide()

    ; start polling the INI every 150ms for a completed selection
    SetTimer(WatchCapDone, 150)

    ; Controller adjustment can take a while; Esc/confirm restores immediately.
    SetTimer(CapWatchFallback, -120000)
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
    global ui, iniPath, __HideWatchKind, __OldMode, __OldRect, __OldTit, __OldIniMTime, __OldPickSeq

    curMode := IniRead(iniPath, "capture", "mode", "")
	curPickSeq := IniRead(iniPath, "capture", "pickSeq", "")
	if (curPickSeq != "" && curPickSeq != __OldPickSeq) {
		FinishCapWatch()
		return
	}
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
CapturePickerStyleFocusedButton(buttons, focusedHwnd) {
    global controlDarkMode

    colors := CPPalette(controlDarkMode)
    for buttonCtrl in buttons {
        try buttonCtrl.Opt("-Default")
        try CPApplyThemeToControl(buttonCtrl.Hwnd, controlDarkMode)
        try buttonCtrl.SetFont("Norm c" colors["text"])
    }

    for buttonCtrl in buttons {
        if (buttonCtrl.Hwnd != focusedHwnd)
            continue
        try buttonCtrl.Opt("+Default")
        try buttonCtrl.SetFont("Bold c" (controlDarkMode ? colors["accentText"] : colors["accentFocus"]))
        try buttonCtrl.Redraw()
        break
    }
}

CapturePickerRefreshFocus(guiHwnd, buttons, focusState, *) {
    if !DllCall("user32\IsWindow", "ptr", guiHwnd, "int")
        return

    focusedHwnd := DllCall("user32\GetFocus", "ptr")
    if (focusState["last"] = focusedHwnd)
        return
    focusState["last"] := focusedHwnd
    CapturePickerStyleFocusedButton(buttons, focusedHwnd)
}

CapturePickerRegisterNavigation(buttons, navigationState) {
    global CPCapturePickerNavigation
    static messageRegistered := false

    if !messageRegistered {
        OnMessage(0x0100, CapturePickerOnKeyDown) ; WM_KEYDOWN
        messageRegistered := true
    }

    for buttonCtrl in buttons
        CPCapturePickerNavigation[buttonCtrl.Hwnd] := navigationState
}

CapturePickerUnregisterNavigation(buttons) {
    global CPCapturePickerNavigation

    for buttonCtrl in buttons {
        buttonHwnd := buttonCtrl.Hwnd
        if CPCapturePickerNavigation.Has(buttonHwnd)
            CPCapturePickerNavigation.Delete(buttonHwnd)
    }
}

CapturePickerOnKeyDown(wParam, lParam, msg, hwnd) {
    global CPCapturePickerNavigation

    focusedHwnd := DllCall("user32\GetFocus", "ptr")
    if !CPCapturePickerNavigation.Has(focusedHwnd)
        return

    navigationState := CPCapturePickerNavigation[focusedHwnd]
    regionHwnd := navigationState["region"]
    windowHwnd := navigationState["window"]
    cancelHwnd := navigationState["cancel"]
    targetHwnd := 0

    if (focusedHwnd = regionHwnd) {
        navigationState["top"] := regionHwnd
        if (wParam = 0x27)       ; Right
            targetHwnd := windowHwnd
        else if (wParam = 0x28)  ; Down
            targetHwnd := cancelHwnd
    } else if (focusedHwnd = windowHwnd) {
        navigationState["top"] := windowHwnd
        if (wParam = 0x25)       ; Left
            targetHwnd := regionHwnd
        else if (wParam = 0x28)  ; Down
            targetHwnd := cancelHwnd
    } else if (focusedHwnd = cancelHwnd) {
        if (wParam = 0x26)       ; Up
            targetHwnd := navigationState["top"]
    }

    if targetHwnd {
        DllCall("user32\SetFocus", "ptr", targetHwnd, "ptr")
        CapturePickerStyleFocusedButton(navigationState["buttons"], targetHwnd)
        return 0
    }

    ; A controller confirmation is delivered to owned dialogs as Enter. Treat
    ; keyboard Enter/Space the same way, while a real pointer click continues
    ; through the ordinary Click event with the default "mouse" source.
    if (wParam = 0x0D || wParam = 0x20) { ; Enter or Space
        navigationState["activationSource"] := "controller"
        SendMessage(0x00F5, 0, 0, focusedHwnd) ; BM_CLICK
        return 0
    }

    ; Consume unused arrows so native dialog order cannot move sideways.
    if (wParam >= 0x25 && wParam <= 0x28)
        return 0
}

CapturePickerClose(g, guiHwnd, focusTimer, buttons, *) {
    SetTimer(focusTimer, 0)
    if !DllCall("user32\IsWindow", "ptr", guiHwnd, "int")
        return

    CapturePickerUnregisterNavigation(buttons)
    try g.Destroy()
}

CapturePickerChoose(g, guiHwnd, focusTimer, buttons, navigationState, mode, *) {
    inputSource := navigationState.Has("activationSource")
        ? navigationState["activationSource"] : "mouse"
    CapturePickerClose(g, guiHwnd, focusTimer, buttons)
    _SendCapPick(mode, inputSource)
}

; --- Native capture mode picker (clamped to virtual desktop, near mouse) ---
OpenCapturePicker(*) {
    global ui, controlDarkMode

    CoordMode "Mouse", "Screen"
    MouseGetPos &mx, &my

    ; virtual desktop across all monitors
    vsx := SysGet(76)  ; SM_XVIRTUALSCREEN
    vsy := SysGet(77)  ; SM_YVIRTUALSCREEN
    vsw := SysGet(78)  ; SM_CXVIRTUALSCREEN
    vsh := SysGet(79)  ; SM_CYVIRTUALSCREEN

    g := Gui("+Owner" ui.Hwnd " +ToolWindow -Caption +AlwaysOnTop -DPIScale")
    colors := CPPalette(controlDarkMode)
    g.BackColor := colors["window"]
    g.MarginX := 16, g.MarginY := 16
    g.SetFont("s10 c" colors["text"], "Segoe UI")

    g.SetFont("s12 Bold c" colors["text"])
    g.Add("Text", "w336 c" colors["text"], "Select capture mode")
    g.SetFont("s10 Norm c" colors["text"])

    b1 := g.Add("Button", "xm y+12 w160 h46", "Region")
    b2 := g.Add("Button", "x+m w160 h46", "Window")
    bCancel := g.Add("Button", "xm y+12 w336 h46", "Cancel") ; 160 + 16 (margin) + 160
    buttons := [b1, b2, bCancel]
    focusState := Map("last", -1)
    navigationState := Map(
        "region", b1.Hwnd,
        "window", b2.Hwnd,
        "cancel", bCancel.Hwnd,
        "top", b1.Hwnd,
        "buttons", buttons,
        "activationSource", "mouse"
    )
    focusTimer := CapturePickerRefreshFocus.Bind(g.Hwnd, buttons, focusState)
    CapturePickerRegisterNavigation(buttons, navigationState)

    b1.OnEvent("Click", CapturePickerChoose.Bind(g, g.Hwnd, focusTimer, buttons, navigationState, "region"))
    b2.OnEvent("Click", CapturePickerChoose.Bind(g, g.Hwnd, focusTimer, buttons, navigationState, "window"))
    bCancel.OnEvent("Click", CapturePickerClose.Bind(g, g.Hwnd, focusTimer, buttons))
    g.OnEvent("Escape", CapturePickerClose.Bind(g, g.Hwnd, focusTimer, buttons))
    g.OnEvent("Close", CapturePickerClose.Bind(g, g.Hwnd, focusTimer, buttons))

    g.Show("AutoSize NoActivate x" vsx+vsw " y" vsy+vsh)
    g.GetPos(, &dlgW, &dlgH)
    x := mx - Floor(dlgW/2)
    y := my + 20
    x := Max(vsx, Min(x, vsx + vsw - dlgW))
    y := Max(vsy, Min(y, vsy + vsh - dlgH))
    g.Move(x, y)
    g.Show()
    CPApplyOwnedDialogTheme(g)
    b1.Focus()
    CapturePickerRefreshFocus(g.Hwnd, buttons, focusState)
    SetTimer(focusTimer, 60)

    ; clean up if left open
    SetTimer(CapturePickerClose.Bind(g, g.Hwnd, focusTimer, buttons), -120000)
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

_SendCapPick(kind, inputSource := "mouse") {
    global capMaxKB
    ; compose a simple, future-proof key=value payload
    payload := "capcmd=pick"
            .  "|kind=" kind
            .  "|maxkb=" capMaxKB
            .  "|regionpreset=" ((kind = "region" && inputSource != "mouse") ? 1 : 0)

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
    global ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt, imgPostproc
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
    imgPostproc := SyncPromptPostproc(ddlPrompt.Text)
    EnvSet("POSTPROC_MODE", imgPostproc)

    ; --- Highlight guessed subjects ---
    EnvSet("SHOT_ITALICIZE_GUESSED", chkGuess.Value ? "1" : "0")
    EnvSet("SHOT_GUESS_DELIM", Chr(0x60))  ; literal backtick

    ; --- Speaker name color toggle (JP+EN; Python strips ã€Œâ€¦ã€ when ON) ---
    global chkName
    EnvSet("SHOT_COLOR_SPEAKER", chkName.Value ? "1" : "0")
}

CPSyncExplanationSelectionFromControls() {
    global explainProvider, explainOpenAIModel, explainGeminiModel
    global ddlEProv, ddlEOpenAI, ddlEGem

    prov := StrLower(Trim(explainProvider))
    if IsSet(ddlEProv) {
        candidate := StrLower(Trim(ddlEProv.Text))
        if (candidate = "gemini" || candidate = "openai")
            prov := candidate
    }
    if (prov != "gemini" && prov != "openai")
        prov := "openai"

    if IsSet(ddlEOpenAI) {
        candidate := Trim(ddlEOpenAI.Text)
        if (candidate != "")
            explainOpenAIModel := candidate
    }
    if IsSet(ddlEGem) {
        candidate := Trim(ddlEGem.Text)
        if (candidate != "")
            explainGeminiModel := candidate
    }
    explainProvider := prov
    return prov
}

ExplainNow(*) {
    static explainRunning := false
    global pythonExe, explainScript
    global explainProvider, explainOpenAIModel, explainGeminiModel
    global debugMode, explainsDir, studyLibraryDir, overlayDir
    px := ResolvePath(pythonExe)
    ex := ResolvePath(explainScript)
    if explainRunning {
        Toast("An explanation is already being generated")
        return
    }
    if !(FileExist(px) && FileExist(ex)) {
        MsgBox("Set valid paths for python.exe and explainer script first.`n`npythonExe:`n" px "`n`nexplainer:`n" ex, "Missing", 48)
        return
    }

    ; The visible Explanation controls are authoritative. This prevents a
    ; defensive INI/UI refresh from leaving the internal provider stale.
    prov := CPSyncExplanationSelectionFromControls()
    if !CPApiKeyConfigured(prov) {
        CPShowMissingApiKey(prov, "explainer")
        DbgCP("ExplainNow blocked: " prov " API key is missing")
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
	
	; --- Explanation archive wiring for explainer.py ---
    saveExpl := Integer(IniRead(iniPath, "cfg", "saveExplains", 0))
    saveStudyLibrary := Integer(IniRead(iniPath, "cfg", "saveStudyLibrary", 0))
    saveStudyScreenshots := Integer(IniRead(iniPath, "cfg", "studyLibraryScreenshots", 1))

    ; Pass environment variables to explainer.py
    ; SAVE_EXPLAINS: "1" to archive each explanation; "0" to skip (default)
    ; EXPLAIN_SAVE_DIR: directory where time-stamped files are written. An active
    ; unified Profile gets its own safely named subfolder; no active Profile keeps
    ; the original root-folder behavior.
    ; SETTINGS_DIR: optional hint for Python's fallback resolution
    explainSaveDir := explainsDir
    activeGameProfile := Trim(IniRead(iniPath, "game_profiles", "active", ""))
    if (activeGameProfile != "") {
        safeProfileFolder := GameProfileSafeName(activeGameProfile)
        if (safeProfileFolder != "")
            explainSaveDir := explainsDir "\" safeProfileFolder
    }
    EnvSet "SAVE_EXPLAINS", (saveExpl ? "1" : "0")
    EnvSet "EXPLAIN_SAVE_DIR", explainSaveDir
    EnvSet "SETTINGS_DIR", A_ScriptDir "\Settings"
    EnvSet "SAVE_STUDY_LIBRARY", (saveStudyLibrary ? "1" : "0")
    EnvSet "STUDY_LIBRARY_SCREENSHOTS", (saveStudyScreenshots ? "1" : "0")
    EnvSet "STUDY_LIBRARY_DIR", studyLibraryDir
    EnvSet "STUDY_LIBRARY_PROFILE", activeGameProfile
    EnvSet "EXPLAIN_PROMPT_PROFILE", Trim(ddlEPr.Text)
    
    requestId := A_NowUTC "-" A_TickCount
    safeRequestId := RegExReplace(requestId, "[^0-9A-Za-z_-]", "_")
    errFile := A_Temp "\JRPG_Explain_" safeRequestId ".err"
    try FileDelete(errFile)
    EnvSet("JRPG_REQUEST_ID", requestId)
    EnvSet("JRPG_MODEL_TIMEOUT_SECONDS", "75")
    EnvSet("EXPLAIN_ERROR_FILE", errFile)
    cmd := Format('"{1}" "{2}"', px, ex)
    DbgCP("ExplainNow -> " cmd)
    Toast("Generating explanation…")
    SignalExplainerBusy()
    explainRunning := true
    pid := 0
    try Run(cmd, A_ScriptDir, "Hide", &pid)
    catch as launchError {
        explainRunning := false
        EnvSet("EXPLAIN_ERROR_FILE", "")
        msg := "Explanation could not start.`n`nThe bundled Python process could not be launched.`n`nDetails: " launchError.Message
        CPWriteExplainerTerminalResult(msg, requestId)
        Toast("Explanation failed")
        MsgBox(msg, "Explain failed", 16)
        DbgCP("ExplainNow launch error: " msg)
        return
    }

    startedAt := A_TickCount
    timedOut := false
    while ProcessExist(pid) {
        if ((A_TickCount - startedAt) >= 120000) {
            timedOut := true
            try ProcessClose(pid)
            break
        }
        Sleep(50)
    }

    err := (FileExist(errFile) ? Trim(FileRead(errFile, "UTF-8")) : "")
    try FileDelete(errFile)
    donePath := overlayDir "\explainer.done"
    doneToken := ""
    try if FileExist(donePath)
        doneToken := Trim(FileRead(donePath, "UTF-8"))
    explainRunning := false
    EnvSet("EXPLAIN_ERROR_FILE", "")

    if timedOut {
        msg := "Explanation timed out.`n`nThe selected model did not finish within two minutes. Please try again, check the internet connection, or select another model."
        CPWriteExplainerTerminalResult(msg, requestId)
        Toast("Explanation failed")
        MsgBox(msg, "Explain failed", 16)
        DbgCP("ExplainNow timeout")
        return
    }

    exitCode := (doneToken = requestId && err = "") ? 0 : 1
    if (exitCode != 0 && err = "") {
        err := "Explanation failed.`n`nThe model process ended without returning a result. Please try again. If this repeats, check the selected provider, model, API key, and internet connection."
        CPWriteExplainerTerminalResult(err, requestId)
    }

    if (exitCode = 0) {
        Toast("Explanation updated")
        DbgCP("ExplainNow OK")
    } else {
        msg := err
        MsgBox(msg, "Explain failed", 16)
        DbgCP("ExplainNow ERR: " msg)
    }
}

StudyLibrarySafeName(slName) {
    slName := Trim(slName)
    if (slName = "" || StrLen(slName) > 80)
        return ""
    if RegExMatch(slName, '[\\/:*?"<>|]')
        return ""
    if RegExMatch(slName, "[\. ]$")
        return ""
    if RegExMatch(slName, 'i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$')
        return ""
    if (StrLower(slName) = "default")
        return ""
    return slName
}

StudyLibraryDirectoryForName(slName) {
    global studyLibraryDefaultDir, studyLibrariesRoot
    if (StrLower(Trim(slName)) = "default")
        return studyLibraryDefaultDir
    slSafeName := StudyLibrarySafeName(slName)
    return slSafeName != "" ? studyLibrariesRoot "\" slSafeName : ""
}

StudyLibraryListNames() {
    global studyLibrariesRoot
    slNames := ["Default"]
    if !DirExist(studyLibrariesRoot)
        return slNames
    slNameText := ""
    Loop Files studyLibrariesRoot "\*", "D" {
        if (StrLower(A_LoopFileName) = "default")
            continue
        slNameText .= A_LoopFileName "`n"
    }
    slNameText := RTrim(slNameText, "`n")
    if (slNameText != "") {
        slNameText := Sort(slNameText)
        Loop Parse slNameText, "`n", "`r"
            if (A_LoopField != "")
                slNames.Push(A_LoopField)
    }
    return slNames
}

StudyLibraryConfiguredName() {
    global iniPath
    slName := Trim(IniRead(iniPath, "study_library", "active", "Default"))
    if (StrLower(slName) = "default")
        return "Default"
    slSafeName := StudyLibrarySafeName(slName)
    slDirectory := StudyLibraryDirectoryForName(slSafeName)
    if (slSafeName = "" || slDirectory = "" || !DirExist(slDirectory))
        return "Default"
    return slSafeName
}

StudyLibraryConfiguredDirectory() {
    return StudyLibraryDirectoryForName(StudyLibraryConfiguredName())
}

StudyLibraryActivateName(slName) {
    global studyLibraryDir, iniPath
    if (StrLower(Trim(slName)) = "default")
        slName := "Default"
    else
        slName := StudyLibrarySafeName(slName)
    slDirectory := StudyLibraryDirectoryForName(slName)
    if (slDirectory = "")
        return false
    if (slName != "Default" && !DirExist(slDirectory))
        return false
    if (slName = "Default")
        DirCreate(slDirectory)
    studyLibraryDir := slDirectory
    IniWrite(slName, iniPath, "study_library", "active")
    return true
}

StudyLibraryApplyProfileSelection(slProfileName, &slWarnings) {
    global CPStudyLibraryState, CPStudyReaderState, studyLibraryDir
    slProfileName := Trim(slProfileName)
    if (StrLower(slProfileName) = "default") {
        slName := "Default"
    } else {
        slName := StudyLibrarySafeName(slProfileName)
        if (slName = "") {
            GameProfileAppendWarning(
                &slWarnings,
                "Study Library '" slProfileName
                    . "' has an invalid name; the current library was kept."
            )
            return false
        }
    }
    slDirectory := StudyLibraryDirectoryForName(slName)
    if (slName != "Default" && !DirExist(slDirectory)) {
        GameProfileAppendWarning(
            &slWarnings,
            "Study Library '" slName
                . "' was not found; the current library was kept."
        )
        return false
    }
    if (StrLower(StudyLibraryConfiguredName()) = StrLower(slName)) {
        studyLibraryDir := StudyLibraryConfiguredDirectory()
        EnvSet("STUDY_LIBRARY_DIR", studyLibraryDir)
        return true
    }
    slApplied := false
    if StudyLibraryStateAlive(CPStudyLibraryState) {
        slApplied := StudyLibrarySwitchTo(CPStudyLibraryState, slName, false)
    } else {
        if IsObject(CPStudyReaderState)
            StudyReaderClose(CPStudyReaderState)
        slApplied := StudyLibraryActivateName(slName)
    }
    if !slApplied {
        GameProfileAppendWarning(
            &slWarnings,
            "Study Library '" slName
                . "' could not be selected; the current library was kept."
        )
        return false
    }
    studyLibraryDir := StudyLibraryConfiguredDirectory()
    EnvSet("STUDY_LIBRARY_DIR", studyLibraryDir)
    return true
}

StudyLibraryRefreshLibrarySelector(slState, slSelect := "") {
    if !StudyLibraryStateAlive(slState)
        return
    if (slSelect = "")
        slSelect := slState.Has("libraryName")
            ? slState["libraryName"] : StudyLibraryConfiguredName()
    slNames := StudyLibraryListNames()
    slState["suspendLibrary"] := true
    try {
        slState["libraryDdl"].Delete()
        slState["libraryDdl"].Add(slNames)
        slIndex := 1
        for slCandidateIndex, slCandidate in slNames {
            if (StrLower(slCandidate) = StrLower(slSelect)) {
                slIndex := slCandidateIndex
                slSelect := slCandidate
                break
            }
        }
        slState["libraryDdl"].Choose(slIndex)
        slState["libraryNames"] := slNames
        slState["libraryName"] := slSelect
    } finally {
        slState["suspendLibrary"] := false
    }
}

StudyLibrarySwitchTo(slState, slName, slAnnounce := true) {
    global CPStudyReaderState, CPStudyCandidateState
    if !StudyLibraryStateAlive(slState)
        return false
    if !StudyLibraryActivateName(slName) {
        StudyLibraryRefreshLibrarySelector(slState)
        return false
    }
    slName := StudyLibraryConfiguredName()
    slDirectory := StudyLibraryConfiguredDirectory()
    if IsObject(CPStudyReaderState)
        StudyReaderClose(CPStudyReaderState)
    if IsObject(CPStudyCandidateState)
        StudyCandidatesClose(CPStudyCandidateState)
    slState["libraryName"] := slName
    slState["database"] := slDirectory "\study_library.db"
    slState["storage"] := 0
    slState["storageLastTick"] := 0
    slState["storageButton"].Enabled := false
    slState["storageButton"].Text := "Storage..."
    slState["search"].Value := ""
    slState["profileMode"] := "all", slState["profileFilter"] := ""
    slState["chapterMode"] := "all", slState["chapterFilter"] := ""
    slState["speakerMode"] := "all", slState["speakerFilter"] := ""
    slState["tagMode"] := "all", slState["tagFilter"] := ""
    slState["ankiMode"] := "all", slState["dateMode"] := "all"
    slState["gui"].Title := "JRPG Translator - Study Library — " slName
    StudyLibraryRefreshLibrarySelector(slState, slName)
    StudyLibraryRunBridge(slState, "ensure")
    StudyLibraryRefresh(slState)
    StudyLibraryRefreshStorage(slState, true)
    if slAnnounce
        Toast("Study Library: " slName)
    return true
}

StudyLibraryLibraryChanged(slState, *) {
    if !StudyLibraryStateAlive(slState)
        return
    if (slState.Has("suspendLibrary") && slState["suspendLibrary"])
        return
    slName := Trim(slState["libraryDdl"].Text)
    if (slName = "" || StrLower(slName) = StrLower(slState["libraryName"]))
        return
    StudyLibrarySwitchTo(slState, slName)
}

StudyLibraryCreateNew(slState, slDialog, slNameEdit, slManager := 0, *) {
    global studyLibrariesRoot
    slName := Trim(slNameEdit.Value)
    slSafeName := StudyLibrarySafeName(slName)
    if (slSafeName = "") {
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "Enter a name up to 80 characters without any of these characters:"
                . "`n`n\ / : * ? `" < > |`n`nThe name Default is reserved.",
            "New Study Library"
        )
        slNameEdit.Focus()
        return
    }
    slDirectory := StudyLibraryDirectoryForName(slSafeName)
    if DirExist(slDirectory) {
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "A Study Library named '" slSafeName "' already exists.",
            "New Study Library"
        )
        slNameEdit.Focus()
        return
    }
    try {
        DirCreate(studyLibrariesRoot)
        DirCreate(slDirectory)
    } catch as slCreateError {
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "The Study Library folder could not be created:`n`n"
                . slCreateError.Message,
            "New Study Library", "ok", "error"
        )
        return
    }
    try slDialog.Destroy()
    StudyLibrarySwitchTo(slState, slSafeName)
    if IsObject(slManager)
        StudyLibraryRefreshManager(slManager, slSafeName)
}

StudyLibraryOpenNew(slState, slManager := 0, *) {
    if !StudyLibraryStateAlive(slState)
        return
    slOwnerHwnd := slState["gui"].Hwnd
    if IsObject(slManager) {
        try slOwnerHwnd := slManager["gui"].Hwnd
    }
    slDialog := Gui(
        "+Owner" slOwnerHwnd " +OwnDialogs",
        "New Study Library"
    )
    slDialog.MarginX := 18, slDialog.MarginY := 16
    slDialog.SetFont("s10", "Segoe UI")
    slDialog.Add("Text", "xm ym w450", "Create a separate Study Library")
        .SetFont("s11 Bold")
    slDialog.Add(
        "Text", "xm y+8 w450 h42 cGray",
        "New explanations will be saved to this library after it is created. "
            . "Existing explanations remain in their current library."
    )
    slNameEdit := slDialog.Add("Edit", "xm y+12 w450", "")
    try slNameEdit.SetCueBanner("For example: Dragon Quest or PC Engine")
    slCreate := slDialog.Add("Button", "xm y+14 w110 Default", "Create")
    slCancel := slDialog.Add("Button", "x+10 yp w100", "Cancel")
    slCreate.OnEvent(
        "Click",
        StudyLibraryCreateNew.Bind(
            slState, slDialog, slNameEdit, slManager
        )
    )
    slCancel.OnEvent("Click", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Close", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slDialog)
    slNameEdit.Focus()
}

StudyLibraryManagerSelectedName(slManager) {
    try slRow := slManager["list"].GetNext()
    catch
        return ""
    return slRow > 0 ? Trim(slManager["list"].GetText(slRow, 1)) : ""
}

StudyLibraryManagerUpdateActions(slManager, *) {
    slName := StudyLibraryManagerSelectedName(slManager)
    slHasSelection := slName != ""
    slIsDefault := StrLower(slName) = "default"
    slIsActive := slHasSelection
        && StrLower(slName) = StrLower(slManager["ownerState"]["libraryName"])
    slManager["switchButton"].Enabled := slHasSelection && !slIsActive
    slManager["renameButton"].Enabled := slHasSelection && !slIsDefault
    slManager["archiveButton"].Enabled := slHasSelection && !slIsDefault
    slManager["openButton"].Enabled := slHasSelection
}

StudyLibraryManagerSummary(slState, slName, slOutputDir) {
    return StudyLibraryManagerSummaryDirectory(
        slState, StudyLibraryDirectoryForName(slName), slOutputDir
    )
}

StudyLibraryManagerSummaryDirectory(slState, slDirectory, slOutputDir) {
    slDatabase := slDirectory "\study_library.db"
    slSummaryState := Map(
        "database", slDatabase,
        "outputDir", slOutputDir
    )
    if !StudyLibraryRunBridge(slSummaryState, "storage")
        return Map("sources", 0, "explanations", 0, "bytes", 0)
    slRows := StudyLibraryReadRows(slOutputDir "\storage.tsv")
    if !slRows.Length || slRows[1].Length < 13
        return Map("sources", 0, "explanations", 0, "bytes", 0)
    return Map(
        "sources", Integer(slRows[1][12]),
        "explanations", Integer(slRows[1][13]),
        "bytes", Integer(slRows[1][9])
    )
}

StudyLibraryRefreshManager(slManager, slPrefer := "") {
    slState := slManager["ownerState"]
    if !StudyLibraryStateAlive(slState)
        return
    if (slPrefer = "")
        slPrefer := StudyLibraryManagerSelectedName(slManager)
    slNames := StudyLibraryListNames()
    slManager["list"].Delete()
    slSelectedRow := 0
    for slIndex, slName in slNames {
        slSummary := StudyLibraryManagerSummary(
            slState, slName, slManager["outputDir"]
        )
        slStatus := StrLower(slName) = StrLower(slState["libraryName"])
            ? "Active" : ""
        slManager["list"].Add(
            "", slName, slStatus, slSummary["sources"],
            slSummary["explanations"],
            StudyLibraryFormatBytes(slSummary["bytes"])
        )
        if (StrLower(slName) = StrLower(slPrefer))
            slSelectedRow := slIndex
    }
    Loop 5
        slManager["list"].ModifyCol(A_Index, "AutoHdr")
    if slNames.Length {
        if !slSelectedRow
            slSelectedRow := 1
        slManager["list"].Modify(slSelectedRow, "Select Focus Vis")
    }
    StudyLibraryManagerUpdateActions(slManager)
}

StudyLibraryManagerSwitch(slManager, *) {
    slName := StudyLibraryManagerSelectedName(slManager)
    if (slName = "")
        return
    if StudyLibrarySwitchTo(slManager["ownerState"], slName)
        StudyLibraryRefreshManager(slManager, slName)
}

StudyLibraryManagerNew(slManager, *) {
    slState := slManager["ownerState"]
    StudyLibraryOpenNew(slState, slManager)
}

StudyLibraryManagerOpenFolder(slManager, *) {
    slName := StudyLibraryManagerSelectedName(slManager)
    slDirectory := StudyLibraryDirectoryForName(slName)
    if (slDirectory != "" && DirExist(slDirectory))
        Run('explorer.exe "' slDirectory '"')
}

StudyLibraryManagerRenameApply(
    slManager, slDialog, slNameEdit, slOldName, *
) {
    global CPStudyReaderState
    slNewName := StudyLibrarySafeName(Trim(slNameEdit.Value))
    if (slNewName = "") {
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "Enter a name up to 80 characters without any of these characters:"
                . "`n`n\ / : * ? `" < > |`n`nThe name Default is reserved.",
            "Rename Study Library"
        )
        slNameEdit.Focus()
        return
    }
    if (StrLower(slNewName) = StrLower(slOldName)) {
        try slDialog.Destroy()
        return
    }
    slOldDirectory := StudyLibraryDirectoryForName(slOldName)
    slNewDirectory := StudyLibraryDirectoryForName(slNewName)
    if DirExist(slNewDirectory) {
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "A Study Library named '" slNewName "' already exists.",
            "Rename Study Library"
        )
        slNameEdit.Focus()
        return
    }
    slState := slManager["ownerState"]
    slWasActive := StrLower(slOldName) = StrLower(slState["libraryName"])
    if slWasActive
        StudyLibrarySwitchTo(slState, "Default", false)
    try {
        DirMove(slOldDirectory, slNewDirectory)
    } catch as slRenameError {
        if slWasActive
            StudyLibrarySwitchTo(slState, slOldName, false)
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "The Study Library could not be renamed:`n`n"
                . slRenameError.Message,
            "Rename Study Library", "ok", "error"
        )
        return
    }
    if slWasActive
        StudyLibrarySwitchTo(slState, slNewName, false)
    else
        StudyLibraryRefreshLibrarySelector(slState)
    try slDialog.Destroy()
    StudyLibraryRefreshManager(slManager, slNewName)
    Toast("Renamed Study Library to " slNewName)
}

StudyLibraryManagerRename(slManager, *) {
    slOldName := StudyLibraryManagerSelectedName(slManager)
    if (slOldName = "" || StrLower(slOldName) = "default")
        return
    slDialog := Gui(
        "+Owner" slManager["gui"].Hwnd " +OwnDialogs",
        "Rename Study Library"
    )
    slDialog.MarginX := 18, slDialog.MarginY := 16
    slDialog.SetFont("s10", "Segoe UI")
    slDialog.Add("Text", "xm ym w420", "Rename '" slOldName "'")
        .SetFont("s11 Bold")
    slNameEdit := slDialog.Add("Edit", "xm y+12 w420", slOldName)
    slRename := slDialog.Add("Button", "xm y+14 w110 Default", "Rename")
    slCancel := slDialog.Add("Button", "x+10 yp w100", "Cancel")
    slRename.OnEvent(
        "Click",
        StudyLibraryManagerRenameApply.Bind(
            slManager, slDialog, slNameEdit, slOldName
        )
    )
    slCancel.OnEvent("Click", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Close", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slDialog)
    slNameEdit.Focus()
    Send("^a")
}

StudyLibraryManagerArchive(slManager, *) {
    global studyLibrariesArchiveRoot
    slName := StudyLibraryManagerSelectedName(slManager)
    if (slName = "" || StrLower(slName) = "default")
        return
    slArchiveMessage := "Archive the Study Library '" slName "'?`n`n"
        . "It will disappear from the library selector, but its database, "
        . "screenshots, backups, and Trash will be moved intact to:`n`n"
        . studyLibrariesArchiveRoot
    if (GlossaryOwnedMessage(
        slManager["gui"].Hwnd, slArchiveMessage,
        "Archive Study Library", "yesno", "warning"
    ) != 6)
        return
    slState := slManager["ownerState"]
    slWasActive := StrLower(slName) = StrLower(slState["libraryName"])
    if slWasActive
        StudyLibrarySwitchTo(slState, "Default", false)
    slSource := StudyLibraryDirectoryForName(slName)
    DirCreate(studyLibrariesArchiveRoot)
    slStamp := FormatTime(, "yyyyMMdd-HHmmss")
    slDestination := studyLibrariesArchiveRoot "\" slName "_" slStamp
    try {
        DirMove(slSource, slDestination)
    } catch as slArchiveError {
        if slWasActive
            StudyLibrarySwitchTo(slState, slName, false)
        GlossaryOwnedMessage(
            slManager["gui"].Hwnd,
            "The Study Library could not be archived:`n`n"
                . slArchiveError.Message,
            "Archive Study Library", "ok", "error"
        )
        return
    }
    StudyLibraryRefreshLibrarySelector(slState)
    StudyLibraryRefreshManager(slManager, "Default")
    Toast("Archived Study Library: " slName)
}

StudyLibraryArchiveEntries() {
    global studyLibrariesArchiveRoot
    slEntries := []
    if !DirExist(studyLibrariesArchiveRoot)
        return slEntries
    slFolderNames := ""
    Loop Files studyLibrariesArchiveRoot "\*", "D"
        slFolderNames .= A_LoopFileName "`n"
    slFolderNames := RTrim(slFolderNames, "`n")
    if (slFolderNames = "")
        return slEntries
    slFolderNames := Sort(slFolderNames, "R")
    Loop Parse slFolderNames, "`n", "`r" {
        slFolderName := A_LoopField
        if (slFolderName = "")
            continue
        slName := slFolderName
        slArchivedAt := "Unknown"
        if RegExMatch(
            slFolderName, "^(.*)_([0-9]{8})-([0-9]{6})$", &slMatch
        ) {
            slName := slMatch[1]
            slDate := slMatch[2], slTime := slMatch[3]
            slArchivedAt := SubStr(slDate, 1, 4) "-"
                . SubStr(slDate, 5, 2) "-" SubStr(slDate, 7, 2) " "
                . SubStr(slTime, 1, 2) ":" SubStr(slTime, 3, 2) ":"
                . SubStr(slTime, 5, 2)
        } else {
            try slArchivedAt := FormatTime(
                FileGetTime(studyLibrariesArchiveRoot "\" slFolderName, "M"),
                "yyyy-MM-dd HH:mm:ss"
            )
        }
        slEntries.Push(Map(
            "name", slName,
            "folderName", slFolderName,
            "path", studyLibrariesArchiveRoot "\" slFolderName,
            "archivedAt", slArchivedAt
        ))
    }
    return slEntries
}

StudyLibraryArchiveSelectedEntry(slArchive) {
    try slRow := slArchive["list"].GetNext()
    catch
        return 0
    if (slRow <= 0 || slRow > slArchive["entries"].Length)
        return 0
    return slArchive["entries"][slRow]
}

StudyLibraryArchiveUpdateActions(slArchive, *) {
    slEntry := StudyLibraryArchiveSelectedEntry(slArchive)
    slEnabled := IsObject(slEntry)
    slArchive["restoreButton"].Enabled := slEnabled
    slArchive["openButton"].Enabled := slEnabled
}

StudyLibraryRefreshArchives(slArchive, slPreferPath := "") {
    if !StudyLibraryStateAlive(slArchive["ownerState"])
        return
    if (slPreferPath = "") {
        slCurrent := StudyLibraryArchiveSelectedEntry(slArchive)
        if IsObject(slCurrent)
            slPreferPath := slCurrent["path"]
    }
    slEntries := StudyLibraryArchiveEntries()
    slArchive["entries"] := slEntries
    slArchive["list"].Delete()
    slSelectedRow := 0
    for slIndex, slEntry in slEntries {
        slSummary := StudyLibraryManagerSummaryDirectory(
            slArchive["ownerState"], slEntry["path"], slArchive["outputDir"]
        )
        slArchive["list"].Add(
            "", slEntry["name"], slEntry["archivedAt"],
            slSummary["sources"], slSummary["explanations"],
            StudyLibraryFormatBytes(slSummary["bytes"])
        )
        if (StrLower(slEntry["path"]) = StrLower(slPreferPath))
            slSelectedRow := slIndex
    }
    Loop 5
        slArchive["list"].ModifyCol(A_Index, "AutoHdr")
    if slEntries.Length {
        if !slSelectedRow
            slSelectedRow := 1
        slArchive["list"].Modify(slSelectedRow, "Select Focus Vis")
        slArchive["emptyText"].Text := ""
    } else {
        slArchive["emptyText"].Text := "No archived Study Libraries."
    }
    StudyLibraryArchiveUpdateActions(slArchive)
}

StudyLibraryArchiveOpenFolder(slArchive, *) {
    slEntry := StudyLibraryArchiveSelectedEntry(slArchive)
    if IsObject(slEntry) && DirExist(slEntry["path"])
        Run('explorer.exe "' slEntry["path"] '"')
}

StudyLibraryArchiveRestoreApply(
    slArchive, slDialog, slNameEdit, slEntry, *
) {
    slName := StudyLibrarySafeName(Trim(slNameEdit.Value))
    if (slName = "") {
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "Enter a name up to 80 characters without any of these characters:"
                . "`n`n\ / : * ? `" < > |`n`nThe name Default is reserved.",
            "Restore Study Library"
        )
        slNameEdit.Focus()
        return
    }
    slDestination := StudyLibraryDirectoryForName(slName)
    if (slDestination = "" || DirExist(slDestination)) {
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "A current Study Library named '" slName "' already exists."
                . "`n`nChoose a different name before restoring this archive.",
            "Restore Study Library"
        )
        slNameEdit.Focus()
        return
    }
    try {
        DirCreate(RegExReplace(slDestination, "\\[^\\]+$"))
        DirMove(slEntry["path"], slDestination)
    } catch as slRestoreError {
        GlossaryOwnedMessage(
            slDialog.Hwnd,
            "The Study Library could not be restored:`n`n"
                . slRestoreError.Message,
            "Restore Study Library", "ok", "error"
        )
        return
    }
    slManager := slArchive["manager"]
    slState := slArchive["ownerState"]
    StudyLibraryRefreshLibrarySelector(slState)
    StudyLibraryRefreshManager(slManager, slName)
    try slDialog.Destroy()
    StudyLibraryRefreshArchives(slArchive)
    Toast("Restored Study Library: " slName)
}

StudyLibraryArchiveRestore(slArchive, *) {
    slEntry := StudyLibraryArchiveSelectedEntry(slArchive)
    if !IsObject(slEntry)
        return
    slSuggestedName := slEntry["name"]
    if DirExist(StudyLibraryDirectoryForName(slSuggestedName))
        slSuggestedName .= " Restored"
    slDialog := Gui(
        "+Owner" slArchive["gui"].Hwnd " +OwnDialogs",
        "Restore Study Library"
    )
    slDialog.MarginX := 18, slDialog.MarginY := 16
    slDialog.SetFont("s10", "Segoe UI")
    slDialog.Add("Text", "xm ym w460", "Restore '" slEntry["name"] "'")
        .SetFont("s11 Bold")
    slDialog.Add(
        "Text", "xm y+8 w460 h42 cGray",
        "The complete archived folder will return to the active library list. "
            . "No existing library will be overwritten."
    )
    slNameEdit := slDialog.Add("Edit", "xm y+12 w460", slSuggestedName)
    slRestore := slDialog.Add("Button", "xm y+14 w110 Default", "Restore")
    slCancel := slDialog.Add("Button", "x+10 yp w100", "Cancel")
    slRestore.OnEvent(
        "Click",
        StudyLibraryArchiveRestoreApply.Bind(
            slArchive, slDialog, slNameEdit, slEntry
        )
    )
    slCancel.OnEvent("Click", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Close", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slDialog)
    slNameEdit.Focus()
    Send("^a")
}

StudyLibraryOpenArchives(slManager, *) {
    global studyLibrariesArchiveRoot
    slState := slManager["ownerState"]
    if !StudyLibraryStateAlive(slState)
        return
    slOutputDir := slManager["outputDir"] "\archives"
    DirCreate(slOutputDir)
    slGui := Gui(
        "+Owner" slManager["gui"].Hwnd " +OwnDialogs",
        "Archived Study Libraries"
    )
    slGui.MarginX := 16, slGui.MarginY := 14
    slGui.SetFont("s10", "Segoe UI")
    slGui.Add("Text", "xm ym w690", "Archived Study Libraries")
        .SetFont("s11 Bold")
    slHint := slGui.Add(
        "Text", "xm y+5 w690 h42 cGray",
        "Archives retain their complete database and media. Restore returns "
            . "an archive to the active library list; nothing here is removed "
            . "automatically."
    )
    CPRegisterMutedControl(slHint)
    slList := slGui.Add(
        "ListView", "xm y+10 w690 h230 Grid -Multi",
        ["Library", "Archived", "Sources", "Explanations", "Storage"]
    )
    slEmpty := slGui.Add("Text", "xm y+5 w690 h24 cGray", "")
    CPRegisterMutedControl(slEmpty)
    slRestore := slGui.Add("Button", "xm y+7 w100 Disabled", "Restore...")
    slOpen := slGui.Add("Button", "x+8 yp w110 Disabled", "Open Folder")
    slClose := slGui.Add("Button", "x+8 yp w82 Default", "Close")
    slArchive := Map(
        "gui", slGui,
        "manager", slManager,
        "ownerState", slState,
        "outputDir", slOutputDir,
        "list", slList,
        "entries", [],
        "emptyText", slEmpty,
        "restoreButton", slRestore,
        "openButton", slOpen
    )
    slList.OnEvent("ItemFocus", StudyLibraryArchiveUpdateActions.Bind(slArchive))
    slList.OnEvent("ItemSelect", StudyLibraryArchiveUpdateActions.Bind(slArchive))
    slList.OnEvent("DoubleClick", StudyLibraryArchiveRestore.Bind(slArchive))
    slRestore.OnEvent("Click", StudyLibraryArchiveRestore.Bind(slArchive))
    slOpen.OnEvent("Click", StudyLibraryArchiveOpenFolder.Bind(slArchive))
    slClose.OnEvent("Click", StudyLibraryCloseDialog.Bind(slGui))
    slGui.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slGui))
    slGui.OnEvent("Close", StudyLibraryCloseDialog.Bind(slGui))
    StudyLibraryRefreshArchives(slArchive)
    slGui.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slGui)
}

StudyLibraryOpenManager(slState, *) {
    if !StudyLibraryStateAlive(slState)
        return
    slOutputDir := slState["outputDir"] "\manager"
    DirCreate(slOutputDir)
    slGui := Gui(
        "+Owner" slState["gui"].Hwnd " +OwnDialogs",
        "Study Libraries"
    )
    slGui.MarginX := 16, slGui.MarginY := 14
    slGui.SetFont("s10", "Segoe UI")
    slTitle := slGui.Add("Text", "xm ym w690", "Manage Study Libraries")
    slTitle.SetFont("s11 Bold")
    slHint := slGui.Add(
        "Text", "xm y+5 w690 h42 cGray",
        "Each library has its own database, screenshots, backups, and Trash. "
            . "Default preserves the original Study Library folder and cannot "
            . "be renamed or archived."
    )
    CPRegisterMutedControl(slHint)
    slList := slGui.Add(
        "ListView", "xm y+10 w690 h230 Grid -Multi",
        ["Library", "Status", "Sources", "Explanations", "Storage"]
    )
    slNew := slGui.Add("Button", "xm y+12 w82", "New...")
    slSwitch := slGui.Add("Button", "x+8 yp w82 Disabled", "Switch")
    slRename := slGui.Add("Button", "x+8 yp w90 Disabled", "Rename...")
    slArchive := slGui.Add("Button", "x+8 yp w90 Disabled", "Archive...")
    slArchived := slGui.Add("Button", "x+8 yp w96", "Archived...")
    slOpen := slGui.Add("Button", "x+8 yp w110 Disabled", "Open Folder")
    slClose := slGui.Add("Button", "x+8 yp w82 Default", "Close")
    slManager := Map(
        "gui", slGui,
        "ownerState", slState,
        "outputDir", slOutputDir,
        "title", slTitle,
        "hint", slHint,
        "list", slList,
        "newButton", slNew,
        "switchButton", slSwitch,
        "renameButton", slRename,
        "archiveButton", slArchive,
        "archivedButton", slArchived,
        "openButton", slOpen,
        "closeButton", slClose
    )
    slList.OnEvent("ItemFocus", StudyLibraryManagerUpdateActions.Bind(slManager))
    slList.OnEvent("ItemSelect", StudyLibraryManagerUpdateActions.Bind(slManager))
    slList.OnEvent("DoubleClick", StudyLibraryManagerSwitch.Bind(slManager))
    slNew.OnEvent("Click", StudyLibraryManagerNew.Bind(slManager))
    slSwitch.OnEvent("Click", StudyLibraryManagerSwitch.Bind(slManager))
    slRename.OnEvent("Click", StudyLibraryManagerRename.Bind(slManager))
    slArchive.OnEvent("Click", StudyLibraryManagerArchive.Bind(slManager))
    slArchived.OnEvent("Click", StudyLibraryOpenArchives.Bind(slManager))
    slOpen.OnEvent("Click", StudyLibraryManagerOpenFolder.Bind(slManager))
    slClose.OnEvent("Click", StudyLibraryCloseDialog.Bind(slGui))
    slGui.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slGui))
    slGui.OnEvent("Close", StudyLibraryCloseDialog.Bind(slGui))
    StudyLibraryRefreshManager(slManager, slState["libraryName"])
    slGui.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slGui)
}

StudyLibrarySaveToggleChanged(*) {
    global saveLibraryChk, saveLibraryScreenshotsChk, iniPath
    enabled := saveLibraryChk.Value ? 1 : 0
    IniWrite(enabled, iniPath, "cfg", "saveStudyLibrary")
    saveLibraryScreenshotsChk.Enabled := enabled
}

StudyLibraryHexDecode(slHex) {
    if (slHex = "")
        return ""
    slByteCount := Floor(StrLen(slHex) / 2)
    if (slByteCount <= 0)
        return ""
    slBuffer := Buffer(slByteCount, 0)
    Loop slByteCount {
        slByte := Integer("0x" SubStr(slHex, (A_Index - 1) * 2 + 1, 2))
        NumPut("UChar", slByte, slBuffer, A_Index - 1)
    }
    return StrGet(slBuffer, slByteCount, "UTF-8")
}

StudyLibraryHexEncode(slText) {
    if (slText = "")
        return ""
    slByteCount := StrPut(slText, "UTF-8") - 1
    if (slByteCount <= 0)
        return ""
    slBuffer := Buffer(slByteCount + 1, 0)
    StrPut(slText, slBuffer, "UTF-8")
    slHex := ""
    Loop slByteCount
        slHex .= Format("{:02x}", NumGet(slBuffer, A_Index - 1, "UChar"))
    return slHex
}

StudyLibraryReadRows(slPath) {
    slRows := []
    if !FileExist(slPath)
        return slRows
    try slText := FileRead(slPath, "UTF-8")
    catch
        return slRows
    Loop Parse slText, "`n", "`r" {
        if (A_LoopField = "")
            continue
        slRows.Push(StrSplit(A_LoopField, "`t"))
    }
    return slRows
}

StudyAnkiConfigPath() {
    return A_ScriptDir "\Settings\anki.ini"
}

StudyAnkiProfileSection(saProfile) {
    ; Keep the human-readable profile name in the values as well, so the safe
    ; INI section cannot silently link the wrong profile after a name collision.
    return "profile_" GameProfileSafeName(saProfile)
}

StudyAnkiEmptyMapping(saProfile := "") {
    return Map(
        "profile", Trim(saProfile),
        "deck", "",
        "addDeck", "",
        "vocabularyDeck", "",
        "model", "Basic",
        "japaneseField", "Front",
        "explanationField", "Back"
    )
}

StudyAnkiIncludeScreenshotDefault() {
    return IniRead(StudyAnkiConfigPath(), "settings", "includeScreenshot", "1") = "1"
}

StudyAnkiSaveIncludeScreenshot(saEnabled) {
    DirCreate(A_ScriptDir "\Settings")
    IniWrite(saEnabled ? "1" : "0", StudyAnkiConfigPath(),
        "settings", "includeScreenshot")
}

StudyAnkiLoadMapping(saProfile) {
    saProfile := Trim(saProfile)
    saMapping := StudyAnkiEmptyMapping(saProfile)
    if (saProfile = "")
        return saMapping
    saPath := StudyAnkiConfigPath()
    saSection := StudyAnkiProfileSection(saProfile)
    saStoredProfile := Trim(IniRead(saPath, saSection, "profile", ""))
    if (saStoredProfile != "" && StrLower(saStoredProfile) = StrLower(saProfile)) {
        saMapping["deck"] := Trim(IniRead(saPath, saSection, "deck", ""))
        saMapping["addDeck"] := Trim(IniRead(saPath, saSection, "addDeck", ""))
        saMapping["vocabularyDeck"] := Trim(
            IniRead(saPath, saSection, "vocabularyDeck", "")
        )
        saMapping["model"] := Trim(IniRead(saPath, saSection, "model", "Basic"))
        saMapping["japaneseField"] := Trim(
            IniRead(saPath, saSection, "japaneseField", "Front")
        )
        saMapping["explanationField"] := Trim(
            IniRead(saPath, saSection, "explanationField", "Back")
        )
        return saMapping
    }

    ; A unified Profile is portable through the LaunchBox plugin. Import its
    ; optional Anki mapping the first time this installation needs it.
    saProfilePath := GameProfilePath(saProfile)
    if FileExist(saProfilePath) {
        saDeck := Trim(IniRead(saProfilePath, "anki", "deck", ""))
        if (saDeck != "") {
            saMapping["deck"] := saDeck
            saMapping["addDeck"] := Trim(
                IniRead(saProfilePath, "anki", "addDeck", "")
            )
            saMapping["vocabularyDeck"] := Trim(
                IniRead(saProfilePath, "anki", "vocabularyDeck", "")
            )
            saMapping["model"] := Trim(
                IniRead(saProfilePath, "anki", "model", "Basic")
            )
            saMapping["japaneseField"] := Trim(
                IniRead(saProfilePath, "anki", "japaneseField", "Front")
            )
            saMapping["explanationField"] := Trim(
                IniRead(saProfilePath, "anki", "explanationField", "Back")
            )
        }
    }
    return saMapping
}

StudyAnkiSaveMapping(saMapping, saWriteUnifiedProfile := true) {
    saProfile := Trim(saMapping["profile"])
    if (saProfile = "")
        return false
    saPath := StudyAnkiConfigPath()
    saSection := StudyAnkiProfileSection(saProfile)
    DirCreate(A_ScriptDir "\Settings")
    IniWrite(saProfile, saPath, saSection, "profile")
    IniWrite(saMapping["deck"], saPath, saSection, "deck")
    IniWrite(
        saMapping.Has("addDeck") ? saMapping["addDeck"] : "",
        saPath, saSection, "addDeck"
    )
    IniWrite(
        saMapping.Has("vocabularyDeck") ? saMapping["vocabularyDeck"] : "",
        saPath, saSection, "vocabularyDeck"
    )
    IniWrite(saMapping["model"], saPath, saSection, "model")
    IniWrite(saMapping["japaneseField"], saPath, saSection, "japaneseField")
    IniWrite(
        saMapping["explanationField"], saPath, saSection,
        "explanationField"
    )
    if saWriteUnifiedProfile {
        saProfilePath := GameProfilePath(saProfile)
        if FileExist(saProfilePath) {
            IniWriteRetry(saMapping["deck"], saProfilePath, "anki", "deck")
            IniWriteRetry(
                saMapping.Has("addDeck") ? saMapping["addDeck"] : "",
                saProfilePath, "anki", "addDeck"
            )
            IniWriteRetry(
                saMapping.Has("vocabularyDeck")
                    ? saMapping["vocabularyDeck"] : "",
                saProfilePath, "anki", "vocabularyDeck"
            )
            IniWriteRetry(saMapping["model"], saProfilePath, "anki", "model")
            IniWriteRetry(
                saMapping["japaneseField"], saProfilePath,
                "anki", "japaneseField"
            )
            IniWriteRetry(
                saMapping["explanationField"], saProfilePath,
                "anki", "explanationField"
            )
        }
    }
    return true
}

StudyAnkiStatusLabel(saStatus, saManualAddedAt := "") {
    saStatus := StrLower(Trim(saStatus))
    if (saStatus = "found")
        return "Found in Anki"
    if (saStatus = "not_found")
        return saManualAddedAt != ""
            ? "Not found (manual marker)" : "Not found"
    if (saManualAddedAt != "")
        return "Added (manual)"
    return "Not checked"
}

StudyAnkiRunBridge(saState, saAction) {
    global pythonExe
    saPython := ResolvePath(pythonExe)
    saBridge := A_ScriptDir "\scripts\anki_bridge.py"
    if !(FileExist(saPython) && FileExist(saBridge)) {
        MsgBox(
            "The Anki integration needs valid Python and bridge paths.`n`n"
            . "Python:`n" saPython "`n`nBridge:`n" saBridge,
            "Study Library - Anki",
            "OK Icon!"
        )
        return false
    }
    saCommand := Format(
        '"{1}" "{2}" {3} --db "{4}" --output-dir "{5}"',
        saPython, saBridge, saAction, saState["database"], saState["outputDir"]
    )
    DbgCP("Anki bridge -> " saCommand)
    try saExitCode := RunWait(saCommand, A_ScriptDir, "Hide")
    catch as saError {
        MsgBox(saError.Message, "Study Library - Anki", "OK Iconx")
        return false
    }
    return saExitCode = 0
}

StudyAnkiReadStatus(saState) {
    saRows := StudyLibraryReadRows(saState["outputDir"] "\anki_status.tsv")
    if !saRows.Length
        return Map("code", "error", "version", "", "message", "No Anki status was returned.")
    saRow := saRows[1]
    return Map(
        "code", saRow.Length >= 1 ? saRow[1] : "error",
        "version", saRow.Length >= 2 ? saRow[2] : "",
        "message", saRow.Length >= 3 ? StudyLibraryHexDecode(saRow[3]) : ""
    )
}

StudyAnkiExactDeckSuggestion(saProfile, saDecks) {
    saProfile := StrLower(Trim(saProfile))
    if (saProfile = "")
        return ""
    for saDeck in saDecks {
        saParts := StrSplit(saDeck, "::")
        saCandidate := ""
        for saPart in saParts {
            saCandidate .= (saCandidate = "" ? "" : "::") saPart
            if (StrLower(Trim(saPart)) = saProfile)
                return saCandidate
        }
    }
    return ""
}

StudyAnkiChooseText(saDdl, saValues, saWanted, saFallback := 1) {
    saTarget := 0
    for saIndex, saValue in saValues {
        if (StrLower(Trim(saValue)) = StrLower(Trim(saWanted))) {
            saTarget := saIndex
            break
        }
    }
    if (!saTarget && saValues.Length && saFallback > 0)
        saTarget := Max(1, Min(saValues.Length, saFallback))
    if saTarget
        saDdl.Choose(saTarget)
    else
        saDdl.Choose(0)
    return saTarget
}

StudyAnkiProfiles(saState) {
    saProfiles := []
    saSeen := Map()
    for saChoice in saState["profileChoices"] {
        if (saChoice["mode"] != "profile")
            continue
        saProfile := Trim(saChoice["value"])
        saKey := StrLower(saProfile)
        if (saProfile != "" && !saSeen.Has(saKey)) {
            saSeen[saKey] := true
            saProfiles.Push(saProfile)
        }
    }
    return saProfiles
}

StudyAnkiSetFields(saDialogState, saJapaneseWanted := "", saExplanationWanted := "") {
    saModelIndex := saDialogState["modelDdl"].Value
    saFields := []
    if (saModelIndex >= 1 && saModelIndex <= saDialogState["models"].Length) {
        saModel := saDialogState["models"][saModelIndex]
        if saDialogState["modelFields"].Has(saModel)
            saFields := saDialogState["modelFields"][saModel]
    }
    saDialogState["fieldNames"] := saFields
    for saDdl in [saDialogState["japaneseDdl"], saDialogState["explanationDdl"]] {
        saDdl.Delete()
        if saFields.Length
            saDdl.Add(saFields)
        saDdl.Enabled := saFields.Length > 0
    }
    StudyAnkiChooseText(
        saDialogState["japaneseDdl"], saFields,
        saJapaneseWanted != "" ? saJapaneseWanted : "Front"
    )
    StudyAnkiChooseText(
        saDialogState["explanationDdl"], saFields,
        saExplanationWanted != "" ? saExplanationWanted : "Back",
        saFields.Length >= 2 ? 2 : 1
    )
}

StudyAnkiModelChanged(saDialogState, *) {
    StudyAnkiSetFields(saDialogState)
}

StudyAnkiApplyProfileMapping(saDialogState, *) {
    saIndex := saDialogState["profileDdl"].Value
    if (saIndex < 1 || saIndex > saDialogState["profiles"].Length)
        return
    saProfile := saDialogState["profiles"][saIndex]
    saMapping := StudyAnkiLoadMapping(saProfile)
    saDeck := saMapping["deck"]
    if (saDeck = "")
        saDeck := StudyAnkiExactDeckSuggestion(saProfile, saDialogState["decks"])
    StudyAnkiChooseText(
        saDialogState["deckDdl"], saDialogState["decks"], saDeck, 0
    )
    StudyAnkiChooseText(
        saDialogState["modelDdl"], saDialogState["models"],
        saMapping["model"] != "" ? saMapping["model"] : "Basic"
    )
    StudyAnkiSetFields(
        saDialogState, saMapping["japaneseField"],
        saMapping["explanationField"]
    )
}

StudyAnkiCurrentMapping(saDialogState) {
    saProfileIndex := saDialogState["profileDdl"].Value
    saDeckIndex := saDialogState["deckDdl"].Value
    saModelIndex := saDialogState["modelDdl"].Value
    saJapaneseIndex := saDialogState["japaneseDdl"].Value
    saExplanationIndex := saDialogState["explanationDdl"].Value
    if (saProfileIndex < 1 || saProfileIndex > saDialogState["profiles"].Length
        || saDeckIndex < 1 || saDeckIndex > saDialogState["decks"].Length
        || saModelIndex < 1 || saModelIndex > saDialogState["models"].Length
        || saJapaneseIndex < 1 || saJapaneseIndex > saDialogState["fieldNames"].Length
        || saExplanationIndex < 1 || saExplanationIndex > saDialogState["fieldNames"].Length)
        return false
    saProfile := saDialogState["profiles"][saProfileIndex]
    saExisting := StudyAnkiLoadMapping(saProfile)
    return Map(
        "profile", saProfile,
        "deck", saDialogState["decks"][saDeckIndex],
        "addDeck", saExisting.Has("addDeck") ? saExisting["addDeck"] : "",
        "vocabularyDeck", saExisting.Has("vocabularyDeck")
            ? saExisting["vocabularyDeck"] : "",
        "model", saDialogState["models"][saModelIndex],
        "japaneseField", saDialogState["fieldNames"][saJapaneseIndex],
        "explanationField", saDialogState["fieldNames"][saExplanationIndex]
    )
}

StudyAnkiUpdateConnectionText(saDialogState, saStatus) {
    saCode := saStatus["code"]
    if (saCode = "connected") {
        saText := "Connected to AnkiConnect"
        if (saStatus["version"] != "")
            saText .= " (API " saStatus["version"] ")"
    } else
        saText := saStatus["message"] != "" ? saStatus["message"] : "Anki is unavailable."
    saDialogState["status"].Value := saText
    saConnected := saCode = "connected"
    for saControl in [
        saDialogState["profileDdl"], saDialogState["deckDdl"],
        saDialogState["modelDdl"], saDialogState["japaneseDdl"],
        saDialogState["explanationDdl"], saDialogState["saveButton"],
        saDialogState["refreshButton"]
    ]
        saControl.Enabled := saConnected
}

StudyAnkiDiscover(saDialogState, *) {
    if !StudyAnkiRunBridge(saDialogState["libraryState"], "discover") {
        saDialogState["status"].Value :=
            "The Anki check failed. Please try Test connection again."
        return
    }
    saLibraryState := saDialogState["libraryState"]
    saStatus := StudyAnkiReadStatus(saLibraryState)
    saDialogState["decks"] := []
    for saRow in StudyLibraryReadRows(saLibraryState["outputDir"] "\anki_decks.tsv") {
        if saRow.Length
            saDialogState["decks"].Push(StudyLibraryHexDecode(saRow[1]))
    }
    saDialogState["models"] := []
    saDialogState["modelFields"] := Map()
    for saRow in StudyLibraryReadRows(saLibraryState["outputDir"] "\anki_models.tsv") {
        if (saRow.Length < 1)
            continue
        saModel := StudyLibraryHexDecode(saRow[1])
        saField := saRow.Length >= 2 ? StudyLibraryHexDecode(saRow[2]) : ""
        if !saDialogState["modelFields"].Has(saModel) {
            saDialogState["models"].Push(saModel)
            saDialogState["modelFields"][saModel] := []
        }
        if (saField != "")
            saDialogState["modelFields"][saModel].Push(saField)
    }
    saDialogState["deckDdl"].Delete()
    saDialogState["modelDdl"].Delete()
    if saDialogState["decks"].Length
        saDialogState["deckDdl"].Add(saDialogState["decks"])
    if saDialogState["models"].Length
        saDialogState["modelDdl"].Add(saDialogState["models"])
    StudyAnkiUpdateConnectionText(saDialogState, saStatus)
    if (saStatus["code"] = "connected")
        StudyAnkiApplyProfileMapping(saDialogState)
}

StudyAnkiSaveDialogMapping(saDialogState, saAnnounce := true, *) {
    saMapping := StudyAnkiCurrentMapping(saDialogState)
    if !IsObject(saMapping) {
        if saAnnounce
            MsgBox(
                "Select a Profile, deck, note type, and fields first.",
                "Study Library - Anki", "OK Icon!"
            )
        return false
    }
    StudyAnkiSaveMapping(saMapping)
    if saAnnounce
        saDialogState["status"].Value := "Mapping saved for " saMapping["profile"] "."
    return saMapping
}

StudyAnkiRefreshLinks(saDialogState, *) {
    global CPStudyReaderState
    saMapping := StudyAnkiSaveDialogMapping(saDialogState, false)
    if !IsObject(saMapping) {
        MsgBox(
            "Select a Profile, deck, note type, and fields first.",
            "Study Library - Anki", "OK Icon!"
        )
        return
    }
    saEnvironment := Map(
        "STUDY_ANKI_PROFILE", saMapping["profile"],
        "STUDY_ANKI_DECK", saMapping["deck"],
        "STUDY_ANKI_MODEL", saMapping["model"],
        "STUDY_ANKI_JAPANESE_FIELD", saMapping["japaneseField"],
        "STUDY_ANKI_EXPLANATION_FIELD", saMapping["explanationField"]
    )
    for saName, saValue in saEnvironment
        EnvSet(saName, saValue)
    try saRan := StudyAnkiRunBridge(saDialogState["libraryState"], "refresh")
    finally {
        for saName, saValue in saEnvironment
            EnvSet(saName, "")
    }
    if !saRan
        return
    saStatus := StudyAnkiReadStatus(saDialogState["libraryState"])
    StudyAnkiUpdateConnectionText(saDialogState, saStatus)
    if (saStatus["code"] != "connected")
        return
    saRows := StudyLibraryReadRows(
        saDialogState["libraryState"]["outputDir"] "\anki_refresh.tsv"
    )
    if saRows.Length && saRows[1].Length >= 4 {
        saRow := saRows[1]
        saDialogState["status"].Value := "Checked " saRow[1] " explanation"
            . (saRow[1] = "1" ? "" : "s") ": " saRow[2] " found, "
            . saRow[3] " not found."
        if (Integer(saRow[4]) > 0)
            saDialogState["status"].Value .= " " saRow[4] " duplicate match(es)."
    }
    saLibraryState := saDialogState["libraryState"]
    StudyLibraryRefresh(saLibraryState)
    if (CPStudyReaderState && CPStudyReaderState.Has("gui")) {
        try StudyLibraryLoadGroup(
            CPStudyReaderState, CPStudyReaderState["currentGroupId"],
            CPStudyReaderState["currentVersion"]
        )
    }
}

StudyLibraryOpenAnki(slState, *) {
    global iniPath
    saProfiles := StudyAnkiProfiles(slState)
    if !saProfiles.Length {
        MsgBox(
            "No named Study Library Profiles are available yet.`n`n"
            . "Generate an explanation while a unified Profile is active first.",
            "Study Library - Anki", "OK Icon!"
        )
        return
    }
    saGui := Gui(
        "+Owner" slState["gui"].Hwnd " +OwnDialogs",
        "Study Library - Anki"
    )
    saGui.MarginX := 18, saGui.MarginY := 16
    saGui.SetFont("s10", "Segoe UI")
    saGui.Add("Text", "xm ym w590", "Read-only Anki link check").SetFont("s11 Bold")
    saGui.Add(
        "Text", "xm y+6 w590 h42",
        "This checks for exact normalized Japanese matches and does not change "
        . "anything in Anki. A parent deck also includes its subdecks."
    )
    saStatusText := saGui.Add("Text", "xm y+12 w430 h42 cGray", "Checking AnkiConnect...")
    CPRegisterMutedControl(saStatusText)
    saTest := saGui.Add("Button", "x+10 yp w140 h30", "Test connection")
    ; Leave the profile row below the taller connection-status text as well as
    ; the shorter Test button, so the themed DropDownList border cannot overlap.
    saGui.Add("Text", "xm y+22 w130", "Study Profile:")
    saProfileDdl := saGui.Add("DropDownList", "x+10 yp-4 w450 0x210", saProfiles)
    saGui.Add("Text", "xm y+14 w130", "Anki deck / parent:")
    saDeckDdl := saGui.Add("DropDownList", "x+10 yp-4 w450 0x210", [])
    saGui.Add("Text", "xm y+14 w130", "Note type:")
    saModelDdl := saGui.Add("DropDownList", "x+10 yp-4 w450 0x210", [])
    saGui.Add("Text", "xm y+14 w130", "Japanese field:")
    saJapaneseDdl := saGui.Add("DropDownList", "x+10 yp-4 w450 0x210", [])
    saGui.Add("Text", "xm y+14 w130", "Explanation field:")
    saExplanationDdl := saGui.Add("DropDownList", "x+10 yp-4 w450 0x210", [])
    saRefresh := saGui.Add("Button", "xm y+18 w190 h32 Default", "Refresh Anki status")
    saSave := saGui.Add("Button", "x+10 yp w130 h32", "Save mapping")
    saClose := saGui.Add("Button", "x+10 yp w100 h32", "Close")
    saState := Map(
        "gui", saGui,
        "libraryState", slState,
        "status", saStatusText,
        "testButton", saTest,
        "profileDdl", saProfileDdl,
        "deckDdl", saDeckDdl,
        "modelDdl", saModelDdl,
        "japaneseDdl", saJapaneseDdl,
        "explanationDdl", saExplanationDdl,
        "refreshButton", saRefresh,
        "saveButton", saSave,
        "profiles", saProfiles,
        "decks", [],
        "models", [],
        "modelFields", Map(),
        "fieldNames", []
    )
    saActiveProfile := Trim(IniRead(iniPath, "game_profiles", "active", ""))
    StudyAnkiChooseText(saProfileDdl, saProfiles, saActiveProfile)
    saProfileDdl.OnEvent("Change", StudyAnkiApplyProfileMapping.Bind(saState))
    saModelDdl.OnEvent("Change", StudyAnkiModelChanged.Bind(saState))
    saTest.OnEvent("Click", StudyAnkiDiscover.Bind(saState))
    saRefresh.OnEvent("Click", StudyAnkiRefreshLinks.Bind(saState))
    saSave.OnEvent("Click", StudyAnkiSaveDialogMapping.Bind(saState, true))
    saClose.OnEvent("Click", StudyLibraryCloseDialog.Bind(saGui))
    saGui.OnEvent("Escape", StudyLibraryCloseDialog.Bind(saGui))
    saGui.OnEvent("Close", StudyLibraryCloseDialog.Bind(saGui))
    saGui.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(saGui)
    StudyAnkiDiscover(saState)
}

StudyLibraryAddSelectedToAnki(slState, *) {
    slSelectedIds := StudyLibrarySelectedGroupIds(slState)
    if (slSelectedIds.Length != 1)
        return
    if (slState["currentGroupId"] != slSelectedIds[1])
        StudyLibraryLoadGroup(slState, slSelectedIds[1])
    if (slState["currentGroupId"] <= 0 || !slState["versions"].Length)
        return
    ; The Library and Reader detail states expose the same loaded explanation,
    ; version, media and Anki fields. Reuse the reviewed Reader workflow so
    ; duplicate checks, mappings, screenshots and HTML formatting stay shared.
    StudyReaderOpenAnkiAddDialog(slState)
}

StudyLibraryShowAnkiMenu(slState, *) {
    slChoices := []
    slActions := []
    if (StudyLibrarySelectedGroupIds(slState).Length = 1
        && slState["currentGroupId"] > 0
        && slState["versions"].Length) {
        slChoices.Push("Add selected explanation...")
        slActions.Push(StudyLibraryAddSelectedToAnki.Bind(slState))
    }
    slChoices.Push("Review Anki candidates...")
    slActions.Push(StudyLibraryOpenCandidates.Bind(slState))
    slChoices.Push("Anki connection and link check...")
    slActions.Push(StudyLibraryOpenAnki.Bind(slState))
    slChoice := CPThemedChoicePopup(
        slState["gui"].Hwnd, slState["ankiButton"].Hwnd, slChoices
    )
    if (slChoice >= 1 && slChoice <= slActions.Length)
        slActions[slChoice].Call()
}

StudyAnkiDeckScope(saParentDeck, saDecks) {
    saScoped := []
    saParentFolded := StrLower(Trim(saParentDeck))
    if (saParentFolded = "")
        return saScoped
    saPrefix := saParentFolded "::"
    for saDeck in saDecks {
        saDeckFolded := StrLower(Trim(saDeck))
        if (saDeckFolded = saParentFolded
            || SubStr(saDeckFolded, 1, StrLen(saPrefix)) = saPrefix)
            saScoped.Push(saDeck)
    }
    return saScoped
}

StudyAnkiTextInList(saValues, saWanted) {
    saWanted := StrLower(Trim(saWanted))
    for saValue in saValues {
        if (StrLower(Trim(saValue)) = saWanted)
            return true
    }
    return false
}

StudyAnkiReadAddResult(saState) {
    saRows := StudyLibraryReadRows(saState["outputDir"] "\anki_add.tsv")
    if !saRows.Length
        return Map(
            "code", "error", "noteId", "",
            "message", "Anki did not return a result for this entry."
        )
    saRow := saRows[1]
    return Map(
        "code", saRow.Length >= 1 ? saRow[1] : "error",
        "noteId", saRow.Length >= 2 ? saRow[2] : "",
        "message", saRow.Length >= 3
            ? StudyLibraryHexDecode(saRow[3]) : ""
    )
}

StudyReaderWriteAnkiReviewFile(saPath, saText, saOwnerHwnd) {
    saFile := 0
    try {
        try FileDelete(saPath)
        saFile := FileOpen(saPath, "w", "UTF-8-RAW")
        if !saFile
            throw Error("The temporary review file could not be created.")
        saFile.Write(saText)
        saFile.Close()
        return true
    } catch as saWriteError {
        try saFile.Close()
        CPThemedOwnedMessage(
            saOwnerHwnd,
            "The reviewed Anki entry could not be prepared.`n`n"
                . saWriteError.Message,
            "Add to Anki", "ok", "error"
        )
        return false
    }
}

StudyReaderCloseAnkiAddDialog(saAddState, *) {
    if saAddState["closed"]
        return
    saAddState["closed"] := true
    try saAddState["gui"].Destroy()
    saOwnerHwnd := saAddState["readerState"]["gui"].Hwnd
    if saOwnerHwnd && DllCall("user32\IsWindow", "ptr", saOwnerHwnd, "int") {
        try DllCall("user32\EnableWindow", "ptr", saOwnerHwnd, "int", 1)
        try WinActivate("ahk_id " saOwnerHwnd)
    }
}

StudyReaderRefreshAfterAnkiAdd(saReaderState) {
    global CPStudyLibraryState
    saGroupId := saReaderState["currentGroupId"]
    saVersion := saReaderState["currentVersion"]
    ; The reviewed-add dialog can be owned by either the Reader or the Library.
    ; Reload its owner once, then refresh the other open view if there is one.
    try StudyLibraryLoadGroup(saReaderState, saGroupId, saVersion)
    if (CPStudyLibraryState && CPStudyLibraryState.Has("gui")) {
        if (CPStudyLibraryState != saReaderState) {
            try StudyLibraryRefresh(CPStudyLibraryState)
            if (CPStudyLibraryState["currentGroupId"] = saGroupId)
                try StudyLibraryLoadGroup(CPStudyLibraryState, saGroupId, saVersion)
        }
    }
}

StudyReaderCurrentScreenshot(saReaderState) {
    saIndex := saReaderState["mediaIndex"]
    if (saIndex < 1 || saIndex > saReaderState["media"].Length)
        return ""
    saPath := saReaderState["media"][saIndex]["path"]
    return FileExist(saPath) ? saPath : ""
}

StudyReaderAnkiScreenshotPreferenceChanged(saAddState, *) {
    if (saAddState["closed"] || !saAddState["includeScreenshot"].Enabled)
        return
    StudyAnkiSaveIncludeScreenshot(saAddState["includeScreenshot"].Value != 0)
}

StudyReaderStripGeneratedExample(saBack) {
    ; Generated examples always occupy a clearly labelled final section.  This
    ; also lets regeneration replace a manually adjusted generated example
    ; without disturbing the reviewed vocabulary definition above it.
    return RTrim(
        RegExReplace(
            saBack,
            "s)(?:\R{2}|^)Example sentence:\R.*\z",
            ""
        ),
        " `t`r`n"
    )
}

StudyReaderHasGeneratedExample(saBack) {
    return RegExMatch(
        saBack,
        "s)(?:\R{2}|^)Example sentence:\R.*\z"
    ) != 0
}

StudyReaderGenerateVocabularyExample(saAddState, *) {
    global pythonExe, explainProvider, explainOpenAIModel, explainGeminiModel
    if saAddState["closed"] || saAddState["cardKind"] != "vocabulary"
        return

    saProvider := CPSyncExplanationSelectionFromControls()
    if !CPApiKeyConfigured(saProvider) {
        saProviderLabel := saProvider = "gemini" ? "Gemini" : "OpenAI"
        CPThemedOwnedMessage(
            saAddState["gui"].Hwnd,
            saProviderLabel " API key missing.`n`nAdd it in the API Keys tab, "
                . "or set the corresponding Windows environment variable and "
                . "restart JRPG Translator.",
            "Generate example sentence", "ok", "error"
        )
        return
    }

    saPython := ResolvePath(pythonExe)
    saExampleScript := A_ScriptDir "\scripts\example_sentence.py"
    if !(FileExist(saPython) && FileExist(saExampleScript)) {
        CPThemedOwnedMessage(
            saAddState["gui"].Hwnd,
            "The configured Python runtime or example-sentence helper could not be found.",
            "Generate example sentence", "ok", "error"
        )
        return
    }

    saFront := Trim(saAddState["front"].Value, " `t`r`n")
    saCurrentBack := saAddState["back"].Value
    saDefinition := StudyReaderStripGeneratedExample(saCurrentBack)
    if (saFront = "" || Trim(saDefinition) = "") {
        saAddState["status"].Value := "Enter the vocabulary term and definition first."
        return
    }

    saReaderState := saAddState["readerState"]
    saOutputDir := saReaderState["outputDir"]
    saTermFile := saOutputDir "\example_term.txt"
    saDefinitionFile := saOutputDir "\example_definition.txt"
    saContextFile := saOutputDir "\example_context.txt"
    saResultFile := saOutputDir "\example_sentence.tsv"
    saOutFile := saOutputDir "\example_sentence_out.txt"
    saErrFile := saOutputDir "\example_sentence_err.txt"
    saContext := saReaderState["sections"].Length
        ? saReaderState["sections"][1]["content"] : ""
    if !StudyReaderWriteAnkiReviewFile(
        saTermFile, saFront, saAddState["gui"].Hwnd
    ) || !StudyReaderWriteAnkiReviewFile(
        saDefinitionFile, saDefinition, saAddState["gui"].Hwnd
    ) || !StudyReaderWriteAnkiReviewFile(
        saContextFile, saContext, saAddState["gui"].Hwnd
    )
        return
    for saTempFile in [saResultFile, saOutFile, saErrFile]
        try FileDelete(saTempFile)

    saEnvironmentNames := [
        "EXPLAIN_PROVIDER", "EXPLAIN_MODEL", "GEMINI_EXPLAIN_MODEL",
        "EXAMPLE_TERM_FILE", "EXAMPLE_DEFINITION_FILE",
        "EXAMPLE_CONTEXT_FILE", "EXAMPLE_RESULT_FILE",
        "SETTINGS_DIR", "PYTHONIOENCODING"
    ]
    saOldEnvironment := Map()
    for saEnvironmentName in saEnvironmentNames
        saOldEnvironment[saEnvironmentName] := EnvGet(saEnvironmentName)

    saExampleButton := saAddState["exampleButton"]
    saExampleButton.Enabled := false
    saAddState["addButton"].Enabled := false
    saAddState["status"].Value := "Generating a learner-friendly example sentence..."
    try DllCall(
        "user32\RedrawWindow", "ptr", saAddState["gui"].Hwnd,
        "ptr", 0, "ptr", 0, "uint", 0x185
    )
    saExitCode := -1
    try {
        EnvSet("EXPLAIN_PROVIDER", saProvider)
        if (saProvider = "gemini") {
            saModel := explainGeminiModel
            if (SubStr(saModel, 1, 7) != "models/")
                saModel := "models/" saModel
            EnvSet("GEMINI_EXPLAIN_MODEL", saModel)
            EnvSet("EXPLAIN_MODEL", "")
        } else {
            EnvSet("EXPLAIN_MODEL", explainOpenAIModel)
            EnvSet("GEMINI_EXPLAIN_MODEL", "")
        }
        EnvSet("EXAMPLE_TERM_FILE", saTermFile)
        EnvSet("EXAMPLE_DEFINITION_FILE", saDefinitionFile)
        EnvSet("EXAMPLE_CONTEXT_FILE", saContextFile)
        EnvSet("EXAMPLE_RESULT_FILE", saResultFile)
        EnvSet("SETTINGS_DIR", A_ScriptDir "\Settings")
        EnvSet("PYTHONIOENCODING", "utf-8")
        saCommand := Format(
            'cmd /c chcp 65001>nul & "{1}" "{2}" 1>"{3}" 2>"{4}"',
            saPython, saExampleScript, saOutFile, saErrFile
        )
        DbgCP("Vocabulary example -> " saCommand)
        saExitCode := RunWait(saCommand, , "Hide")
    } finally {
        for saEnvironmentName, saEnvironmentValue in saOldEnvironment
            EnvSet(saEnvironmentName, saEnvironmentValue)
    }

    saExampleButton.Enabled := true
    saAddState["addButton"].Enabled := true
    saRows := StudyLibraryReadRows(saResultFile)
    if (saExitCode != 0 || !saRows.Length || saRows[1].Length < 4
        || saRows[1][1] != "ok") {
        saError := FileExist(saErrFile)
            ? Trim(FileRead(saErrFile, "UTF-8")) : ""
        saOutput := FileExist(saOutFile)
            ? Trim(FileRead(saOutFile, "UTF-8")) : ""
        saAddState["status"].Value := "Example generation failed."
        CPThemedOwnedMessage(
            saAddState["gui"].Hwnd,
            "The example sentence could not be generated.`n`n"
                . (saError != "" ? saError : saOutput),
            "Generate example sentence", "ok", "error", 650
        )
        return
    }

    saPlain := StudyLibraryHexDecode(saRows[1][2])
    saReading := StudyLibraryHexDecode(saRows[1][3])
    saTranslation := StudyLibraryHexDecode(saRows[1][4])
    saExampleBlock := "Example sentence:`r`n" saPlain
        . "`r`n`r`n" saReading "`r`n`r`n" saTranslation
    saNewBack := saDefinition
        . (saDefinition != "" ? "`r`n`r`n" : "")
        . saExampleBlock
    saAddState["back"].Value := saNewBack
    saAddState["exampleGenerated"] := true
    saExampleButton.Text := "Regenerate example…"
    ; Put the newly generated section in view immediately for review.
    try {
        saAddState["back"].Focus()
        saEnd := StrLen(saNewBack)
        SendMessage(0x00B1, saEnd, saEnd, saAddState["back"].Hwnd) ; EM_SETSEL
        SendMessage(0x00B7, 0, 0, saAddState["back"].Hwnd) ; EM_SCROLLCARET
    }
    saAddState["status"].Value := "Example generated. Review it before adding."
}

StudyReaderAddReviewedAnkiNote(saAddState, *) {
    if saAddState["closed"]
        return
    saDeckIndex := saAddState["deckDdl"].Value
    if (saDeckIndex < 1 || saDeckIndex > saAddState["decks"].Length) {
        saAddState["status"].Value := "Select a destination deck."
        return
    }
    saDeck := saAddState["decks"][saDeckIndex]
    saFront := Trim(saAddState["front"].Value, " `t`r`n")
    saBack := Trim(saAddState["back"].Value, " `t`r`n")
    if (saFront = "" || saBack = "") {
        saAddState["status"].Value := "Both the Japanese and explanation fields are required."
        return
    }
    saMapping := saAddState["mapping"]
    saCardKind := saAddState.Has("cardKind")
        ? saAddState["cardKind"] : "explanation"
    saIsVocabulary := saCardKind = "vocabulary"
    saIncludeScreenshot := saAddState["includeScreenshot"].Enabled
        && saAddState["includeScreenshot"].Value
        && FileExist(saAddState["screenshotPath"])
    saConfirm := "Add this "
        . (saIsVocabulary ? "vocabulary" : "explanation")
        . " entry to Anki?`n`n"
        . "Deck: " saDeck "`n"
        . "Note type: " saMapping["model"] "`n"
        . "Fields: " saMapping["japaneseField"] " / "
        . saMapping["explanationField"] "`n"
        . "Source screenshot: " (saIncludeScreenshot
            ? "Include compact JPEG on the back" : "Do not include")
    if (CPThemedOwnedMessage(
        saAddState["gui"].Hwnd, saConfirm, "Add to Anki",
        "yesno", "info", 560
    ) != "Yes")
        return

    saReaderState := saAddState["readerState"]
    saFrontPath := saReaderState["outputDir"] "\anki_review_front.txt"
    saBackPath := saReaderState["outputDir"] "\anki_review_back.txt"
    if !StudyReaderWriteAnkiReviewFile(
        saFrontPath, saFront, saAddState["gui"].Hwnd
    ) || !StudyReaderWriteAnkiReviewFile(
        saBackPath, saBack, saAddState["gui"].Hwnd
    )
        return

    saEnvironment := Map(
        "STUDY_ANKI_PROFILE", saMapping["profile"],
        "STUDY_ANKI_PARENT_DECK", saMapping["deck"],
        "STUDY_ANKI_DECK", saDeck,
        "STUDY_ANKI_MODEL", saMapping["model"],
        "STUDY_ANKI_JAPANESE_FIELD", saMapping["japaneseField"],
        "STUDY_ANKI_EXPLANATION_FIELD", saMapping["explanationField"],
        "STUDY_ANKI_GROUP_ID", saReaderState["currentGroupId"],
        "STUDY_ANKI_VERSION", saReaderState["currentVersion"],
        "STUDY_ANKI_FRONT_PATH", saFrontPath,
        "STUDY_ANKI_BACK_PATH", saBackPath,
        "STUDY_ANKI_CARD_KIND", saCardKind,
        "STUDY_ANKI_INCLUDE_SCREENSHOT", saIncludeScreenshot ? "1" : "0",
        "STUDY_ANKI_SCREENSHOT_PATH", saIncludeScreenshot
            ? saAddState["screenshotPath"] : ""
    )
    saAddState["addButton"].Enabled := false
    saAddState["status"].Value := "Checking for duplicates and adding to Anki..."
    for saName, saValue in saEnvironment
        EnvSet(saName, saValue)
    try saRan := StudyAnkiRunBridge(saReaderState, "add-note")
    finally {
        for saName, saValue in saEnvironment
            EnvSet(saName, "")
    }
    if !saRan {
        saAddState["status"].Value := "The Anki bridge could not be started. Nothing was added to Anki."
        saAddState["addButton"].Enabled := true
        return
    }
    saResult := StudyAnkiReadAddResult(saReaderState)
    saCode := StrLower(saResult["code"])
    if (saCode = "added" || saCode = "duplicate"
        || saCode = "added_unlinked") {
        if saIsVocabulary
            saMapping["vocabularyDeck"] := saDeck
        else
            saMapping["addDeck"] := saDeck
        StudyAnkiSaveMapping(saMapping)
        StudyReaderCloseAnkiAddDialog(saAddState)
        if !saIsVocabulary
            StudyReaderRefreshAfterAnkiAdd(saReaderState)
        if (saCode = "added") {
            CPThemedOwnedMessage(
                saReaderState["gui"].Hwnd,
                saResult["message"] != "" ? saResult["message"]
                    : "The reviewed " (saIsVocabulary
                        ? "vocabulary card" : "explanation")
                        . " was added to Anki.",
                "Added to Anki", "ok", "info"
            )
        } else if (saCode = "duplicate") {
            CPThemedOwnedMessage(
                saReaderState["gui"].Hwnd,
                "Nothing new was added to Anki.`n`n" saResult["message"],
                "Already in Anki", "ok", "info"
            )
        } else {
            CPThemedOwnedMessage(
                saReaderState["gui"].Hwnd, saResult["message"],
                "Added, but not linked", "ok", "warning"
            )
        }
        return
    }
    saAddState["status"].Value := saResult["message"] != ""
        ? saResult["message"] : "Anki rejected the entry. Nothing was added to Anki."
    saAddState["addButton"].Enabled := true
}

StudyReaderOpenAnkiAddDialog(srState, *) {
    StudyReaderOpenReviewedAnkiDialog(srState, "explanation", "", "")
}

StudyReaderControlText(srControl) {
    if !(IsObject(srControl) && srControl.Hwnd)
        return ""
    srLength := DllCall(
        "user32\GetWindowTextLengthW", "ptr", srControl.Hwnd, "int"
    )
    if (srLength <= 0)
        return ""
    srBuffer := Buffer((srLength + 1) * 2, 0)
    DllCall(
        "user32\GetWindowTextW", "ptr", srControl.Hwnd,
        "ptr", srBuffer.Ptr, "int", srLength + 1
    )
    return StrGet(srBuffer, srLength, "UTF-16")
}

StudyReaderSelectedExplanationText(srState) {
    if !(srState.Has("explanation") && srState["explanation"].Hwnd)
        return ""
    srStart := Buffer(4, 0), srEnd := Buffer(4, 0)
    DllCall(
        "user32\SendMessageW", "ptr", srState["explanation"].Hwnd,
        "uint", 0x00B0, "ptr", srStart.Ptr, "ptr", srEnd.Ptr
    ) ; EM_GETSEL
    srStartIndex := NumGet(srStart, 0, "UInt")
    srEndIndex := NumGet(srEnd, 0, "UInt")
    if (srEndIndex <= srStartIndex)
        return ""
    srText := StudyReaderControlText(srState["explanation"])
    return Trim(
        SubStr(srText, srStartIndex + 1, srEndIndex - srStartIndex),
        " `t`r`n"
    )
}

StudyReaderInferVocabularyCard(srSelection) {
    srBack := Trim(srSelection, " `t`r`n")
    srLine := ""
    Loop Parse srBack, "`n", "`r" {
        if (Trim(A_LoopField) != "") {
            srLine := Trim(A_LoopField)
            break
        }
    }
    srLine := RegExReplace(srLine, "^\s*(?:[*•●▪◦]|[-–—]\s+)")
    srDelimiter := 0
    if RegExMatch(srLine, "\s*[—–―]\s*|\s+-\s+", &srDash)
        srDelimiter := srDash.Pos(0)
    srHead := srDelimiter > 0
        ? Trim(SubStr(srLine, 1, srDelimiter - 1)) : ""
    if InStr(srHead, "→")
        srHead := Trim(RegExReplace(srHead, "^.*→\s*"))
    srFront := StudyReaderRemoveJapaneseReadings(srHead)
    ; This is only the isolated term heading, so kana-only parentheticals are
    ; readings as well (for example a custom prompt's katakana spelling note).
    srFront := RegExReplace(
        srFront, "[\(（][ぁ-ゖァ-ヺー・･ /]+[\)）]", ""
    )
    srFront := Trim(srFront, " `t`r`n*•●▪◦")
    srConfident := srFront != ""
        && RegExMatch(srFront, "[ぁ-ゖァ-ヺ一-龯々〆ヵヶ]")
    return Map(
        "front", srConfident ? srFront : "",
        "back", srBack,
        "confident", srConfident
    )
}

StudyReaderOpenVocabularyAnkiDialog(srState, *) {
    srSelection := StudyReaderSelectedExplanationText(srState)
    if (srSelection = "") {
        CPThemedOwnedMessage(
            srState["gui"].Hwnd,
            "Select one complete vocabulary entry in the explanation text "
                . "first, then choose Add selected vocabulary again.",
            "Add vocabulary to Anki", "ok", "info"
        )
        return
    }
    srCard := StudyReaderInferVocabularyCard(srSelection)
    StudyReaderOpenReviewedAnkiDialog(
        srState, "vocabulary", srCard["front"], srCard["back"],
        srCard["confident"]
    )
}

StudyReaderShowAnkiMenu(srState, *) {
    if ((srState.Has("editing") && srState["editing"])
        || srState["currentGroupId"] <= 0)
        return
    srChoices := [
        "Add full explanation...",
        "Add selected vocabulary..."
    ]
    srActions := [
        StudyReaderOpenAnkiAddDialog.Bind(srState),
        StudyReaderOpenVocabularyAnkiDialog.Bind(srState)
    ]
    srChoice := CPThemedChoicePopup(
        srState["gui"].Hwnd, srState["addAnkiButton"].Hwnd, srChoices
    )
    if (srChoice >= 1 && srChoice <= srActions.Length)
        srActions[srChoice].Call()
}

StudyReaderOpenReviewedAnkiDialog(
    srState, saCardKind, saReviewedFront, saReviewedBack,
    saInferenceConfident := true
) {
    saIsVocabulary := saCardKind = "vocabulary"
    if ((srState.Has("editing") && srState["editing"])
        || srState["currentGroupId"] <= 0)
        return
    if (!saIsVocabulary && srState["currentAnkiStatus"] = "found") {
        CPThemedOwnedMessage(
            srState["gui"].Hwnd,
            "An exact normalized Japanese match is currently marked as "
                . "present in Anki.`n`n"
                . "If you deleted an older card in Anki, return to the Study "
                . "Library, choose Anki... > Anki connection and link check..., "
                . "then select Refresh Anki status "
                . "before trying again.",
            "Already in Anki", "ok", "info"
        )
        return
    }
    saProfile := Trim(srState["currentProfile"])
    if (saProfile = "") {
        CPThemedOwnedMessage(
            srState["gui"].Hwnd,
            "This explanation has no named Profile. Assign a Profile before "
                . "adding it to Anki so its deck mapping is unambiguous.",
            "Add to Anki", "ok", "warning"
        )
        return
    }
    saMapping := StudyAnkiLoadMapping(saProfile)
    if (saMapping["deck"] = "" || saMapping["model"] = ""
        || saMapping["japaneseField"] = ""
        || saMapping["explanationField"] = "") {
        CPThemedOwnedMessage(
            srState["gui"].Hwnd,
            "No complete Anki mapping is saved for '" saProfile "'.`n`n"
                . "Open Study Library > Libraries... > Anki setup first.",
            "Add to Anki", "ok", "warning"
        )
        return
    }

    if !StudyAnkiRunBridge(srState, "discover")
        return
    saStatus := StudyAnkiReadStatus(srState)
    if (saStatus["code"] != "connected") {
        CPThemedOwnedMessage(
            srState["gui"].Hwnd, saStatus["message"],
            "Anki unavailable", "ok", "warning"
        )
        return
    }
    saAllDecks := []
    for saRow in StudyLibraryReadRows(srState["outputDir"] "\anki_decks.tsv") {
        if saRow.Length
            saAllDecks.Push(StudyLibraryHexDecode(saRow[1]))
    }
    saDecks := StudyAnkiDeckScope(saMapping["deck"], saAllDecks)
    if !saDecks.Length {
        CPThemedOwnedMessage(
            srState["gui"].Hwnd,
            "The mapped Anki parent deck was not found:`n`n" saMapping["deck"],
            "Add to Anki", "ok", "warning"
        )
        return
    }
    saModelFields := Map()
    for saRow in StudyLibraryReadRows(srState["outputDir"] "\anki_models.tsv") {
        if (saRow.Length < 1)
            continue
        saModel := StudyLibraryHexDecode(saRow[1])
        saField := saRow.Length >= 2 ? StudyLibraryHexDecode(saRow[2]) : ""
        if !saModelFields.Has(saModel)
            saModelFields[saModel] := []
        if (saField != "")
            saModelFields[saModel].Push(saField)
    }
    if (!saModelFields.Has(saMapping["model"])
        || !StudyAnkiTextInList(
            saModelFields[saMapping["model"]], saMapping["japaneseField"]
        ) || !StudyAnkiTextInList(
            saModelFields[saMapping["model"]], saMapping["explanationField"]
        )) {
        CPThemedOwnedMessage(
            srState["gui"].Hwnd,
            "The saved note type or field mapping no longer exists in Anki.`n`n"
                . "Open the Anki setup and update the mapping first.",
            "Add to Anki", "ok", "warning"
        )
        return
    }

    saFront := saIsVocabulary ? saReviewedFront
        : StudyReaderRemoveJapaneseReadings(srState["source"].Value)
    saBack := saIsVocabulary ? saReviewedBack
        : (srState["sections"].Length
            ? srState["sections"][1]["content"] : "")
    saGui := Gui(
        "+Owner" srState["gui"].Hwnd " +OwnDialogs +AlwaysOnTop",
        saIsVocabulary ? "Add vocabulary to Anki" : "Add explanation to Anki"
    )
    saGui.MarginX := 16, saGui.MarginY := 14
    saGui.SetFont("s10", "Segoe UI")
    saGui.Add("Text", "xm ym w560", "Review Anki entry").SetFont("s11 Bold")
    saHint := saGui.Add(
        "Text", "xm y+4 w560 h34 cGray",
        saIsVocabulary
            ? "The selected entry was parsed locally. Review the inferred term "
                . "and definition before adding; both fields remain editable."
            : "Nothing is sent until you confirm Add to Anki. The Japanese field starts "
                . "without attached kana readings; both fields remain editable."
    )
    CPRegisterMutedControl(saHint)
    saGui.Add(
        "Text", "xm y+6 w560 h36",
        "Profile: " saProfile "    Note type: " saMapping["model"] "`n"
            . "Fields: " saMapping["japaneseField"] " / "
            . saMapping["explanationField"]
    )
    saGui.Add("Text", "xm y+6 w120", "Destination deck:")
    saDeckDdl := saGui.Add("DropDownList", "x+8 yp-4 w430 0x210", saDecks)
    saPreferredDeck := saIsVocabulary
        && saMapping.Has("vocabularyDeck")
        && saMapping["vocabularyDeck"] != ""
            ? saMapping["vocabularyDeck"]
            : (saMapping["addDeck"] != ""
                ? saMapping["addDeck"] : saMapping["deck"])
    StudyAnkiChooseText(saDeckDdl, saDecks, saPreferredDeck)
    saGui.Add("Text", "xm y+8 w560", saMapping["japaneseField"] " (Japanese)")
    saFrontEdit := saGui.Add(
        "Edit", "xm y+3 w560 h58 Multi VScroll", saFront
    )
    saBackLabel := saGui.Add(
        "Text", "xm y+7 w560 h20",
        saMapping["explanationField"] " (explanation)"
    )
    saBackLabel.GetPos(&saBackLabelX, &saBackLabelY, &saBackLabelW, &saBackLabelH)
    saExampleButton := 0
    if saIsVocabulary {
        saExampleButton := saGui.Add(
            "Button", "x386 y" (saBackLabelY - 5) " w190 h28",
            StudyReaderHasGeneratedExample(saBack)
                ? "Regenerate example…" : "Generate example…"
        )
    }
    saBackEdit := saGui.Add(
        "Edit", "xm y" (saBackLabelY + 23) " w560 h118 Multi VScroll", saBack
    )
    saScreenshotPath := StudyReaderCurrentScreenshot(srState)
    saHasScreenshot := saScreenshotPath != ""
    saGui.Add("Text", "xm y+7 w560", "Source screenshot")
    saPreviewFrame := saGui.Add(
        "Text", "xm y+3 w132 h78 Border Background101010", ""
    )
    saPreviewFrame.GetPos(&saFrameX, &saFrameY, &saFrameW, &saFrameH)
    saPreviewPicture := saGui.Add(
        "Picture", "x" (saFrameX + 3) " y" (saFrameY + 3) " w1 h1 Hidden", ""
    )
    if saHasScreenshot {
        saNativeW := 0, saNativeH := 0
        StudyLibraryImageDimensions(saScreenshotPath, &saNativeW, &saNativeH)
        if (saNativeW <= 0 || saNativeH <= 0)
            saNativeW := 126, saNativeH := 72
        saPreviewScale := Min(126 / saNativeW, 72 / saNativeH)
        saPreviewW := Max(1, Round(saNativeW * saPreviewScale))
        saPreviewH := Max(1, Round(saNativeH * saPreviewScale))
        saPreviewX := saFrameX + 3 + Floor((126 - saPreviewW) / 2)
        saPreviewY := saFrameY + 3 + Floor((72 - saPreviewH) / 2)
        saPreviewDpi := GetWindowDPI(saGui.Hwnd)
        saPreviewLoadW := Max(1, Round(saPreviewW * saPreviewDpi / 96))
        saPreviewLoadH := Max(1, Round(saPreviewH * saPreviewDpi / 96))
        try {
            saPreviewPicture.Value := "*w" saPreviewLoadW " *h"
                . saPreviewLoadH " " saScreenshotPath
            saPreviewPicture.Move(saPreviewX, saPreviewY, saPreviewW, saPreviewH)
            saPreviewPicture.Visible := true
        }
    }
    saIncludeScreenshot := saGui.Add(
        "CheckBox", "x160 y" saFrameY " w400 h24",
        "Include current screenshot on the card back"
    )
    saIncludeScreenshot.Enabled := saHasScreenshot
    saIncludeScreenshot.Value := saHasScreenshot
        && StudyAnkiIncludeScreenshotDefault() ? 1 : 0
    saScreenshotHint := saGui.Add(
        "Text", "x160 y" (saFrameY + 28) " w400 h45 cGray",
        saHasScreenshot
            ? "A compact JPEG copy (maximum 1280 px) is created for Anki. "
                . "The Study Library original remains unchanged."
            : "No source screenshot is saved for this explanation version."
    )
    CPRegisterMutedControl(saScreenshotHint)
    saStatusText := saGui.Add(
        "Text", "xm y+9 w560 h28 cGray",
        !saIsVocabulary || saInferenceConfident
            ? "Exact normalized Japanese duplicates are rejected before creation."
            : "The term could not be inferred confidently. Enter the front field manually."
    )
    CPRegisterMutedControl(saStatusText)
    saAddButton := saGui.Add("Button", "xm y+8 w110 h32 Default", "Add to Anki")
    saCancelButton := saGui.Add("Button", "x+10 yp w100 h32", "Cancel")
    saAddState := Map(
        "gui", saGui,
        "readerState", srState,
        "mapping", saMapping,
        "cardKind", saCardKind,
        "decks", saDecks,
        "deckDdl", saDeckDdl,
        "front", saFrontEdit,
        "back", saBackEdit,
        "screenshotPath", saScreenshotPath,
        "includeScreenshot", saIncludeScreenshot,
        "status", saStatusText,
        "addButton", saAddButton,
        "exampleButton", saExampleButton,
        "exampleGenerated", StudyReaderHasGeneratedExample(saBack),
        "closed", false
    )
    saIncludeScreenshot.OnEvent(
        "Click", StudyReaderAnkiScreenshotPreferenceChanged.Bind(saAddState)
    )
    saAddButton.OnEvent("Click", StudyReaderAddReviewedAnkiNote.Bind(saAddState))
    if saIsVocabulary
        saExampleButton.OnEvent(
            "Click", StudyReaderGenerateVocabularyExample.Bind(saAddState)
        )
    saCancelButton.OnEvent("Click", StudyReaderCloseAnkiAddDialog.Bind(saAddState))
    saGui.OnEvent("Escape", StudyReaderCloseAnkiAddDialog.Bind(saAddState))
    saGui.OnEvent("Close", StudyReaderCloseAnkiAddDialog.Bind(saAddState))
    try DllCall("user32\EnableWindow", "ptr", srState["gui"].Hwnd, "int", 0)
    saGui.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(saGui)
    saFrontEdit.Focus()
}

StudyCandidatesRecommendationDefaults() {
    return Map(
        "learnerLevel", "intermediate",
        "selectionStyle", "balanced",
        "focusVocabulary", true,
        "focusGrammar", true,
        "focusNaturalPhrasing", true,
        "focusReading", true,
        "additionalCriteria", ""
    )
}

StudyCandidatesRecommendationLoadSettings() {
    global iniPath
    scDefaults := StudyCandidatesRecommendationDefaults()
    scSettings := Map()
    scLevel := StrLower(Trim(IniRead(
        iniPath, "ankiRecommendations", "learnerLevel",
        scDefaults["learnerLevel"]
    )))
    if !(scLevel = "beginner" || scLevel = "intermediate"
        || scLevel = "advanced")
        scLevel := scDefaults["learnerLevel"]
    scStyle := StrLower(Trim(IniRead(
        iniPath, "ankiRecommendations", "selectionStyle",
        scDefaults["selectionStyle"]
    )))
    if !(scStyle = "selective" || scStyle = "balanced"
        || scStyle = "generous")
        scStyle := scDefaults["selectionStyle"]
    scAdditionalHex := Trim(IniRead(
        iniPath, "ankiRecommendations", "additionalCriteriaHex", ""
    ))
    scAdditional := ""
    try scAdditional := StudyLibraryHexDecode(scAdditionalHex)
    scSettings["learnerLevel"] := scLevel
    scSettings["selectionStyle"] := scStyle
    for scKey in [
        "focusVocabulary", "focusGrammar",
        "focusNaturalPhrasing", "focusReading"
    ]
        scSettings[scKey] := Trim(IniRead(
            iniPath, "ankiRecommendations", scKey,
            scDefaults[scKey] ? "1" : "0"
        )) != "0"
    scSettings["additionalCriteria"] := SubStr(scAdditional, 1, 1000)
    return scSettings
}

StudyCandidatesRecommendationSaveSettings(scSettings) {
    global iniPath
    IniWrite(
        scSettings["learnerLevel"], iniPath,
        "ankiRecommendations", "learnerLevel"
    )
    IniWrite(
        scSettings["selectionStyle"], iniPath,
        "ankiRecommendations", "selectionStyle"
    )
    for scKey in [
        "focusVocabulary", "focusGrammar",
        "focusNaturalPhrasing", "focusReading"
    ]
        IniWrite(
            scSettings[scKey] ? 1 : 0, iniPath,
            "ankiRecommendations", scKey
        )
    IniWrite(
        StudyLibraryHexEncode(SubStr(
            scSettings["additionalCriteria"], 1, 1000
        )),
        iniPath, "ankiRecommendations", "additionalCriteriaHex"
    )
}

StudyCandidatesRecommendationFocusCsv(scSettings) {
    scFocus := ""
    for scPair in [
        ["focusVocabulary", "vocabulary"],
        ["focusGrammar", "grammar"],
        ["focusNaturalPhrasing", "natural_phrasing"],
        ["focusReading", "reading"]
    ] {
        if !scSettings[scPair[1]]
            continue
        scFocus .= (scFocus = "" ? "" : ",") scPair[2]
    }
    return scFocus != ""
        ? scFocus : "vocabulary,grammar,natural_phrasing,reading"
}

StudyCandidatesRecommendationCounts(scState, scIncludeAssessed := false) {
    scSentenceKeys := Map(), scVocabularyKeys := Map()
    for scCandidate in scState["allSentences"] {
        if (scIncludeAssessed || scCandidate["recommendation"] < 0)
            scSentenceKeys[scCandidate["recommendationKey"]] := true
    }
    for scCandidate in scState["allVocabulary"] {
        if (scIncludeAssessed || scCandidate["recommendation"] < 0)
            scVocabularyKeys[scCandidate["recommendationKey"]] := true
    }
    return Map(
        "sentences", scSentenceKeys.Count,
        "vocabulary", scVocabularyKeys.Count,
        "total", scSentenceKeys.Count + scVocabularyKeys.Count
    )
}

StudyCandidatesRecommendationLevelIndex(scLevel) {
    return scLevel = "beginner" ? 1 : (scLevel = "advanced" ? 3 : 2)
}

StudyCandidatesRecommendationStyleIndex(scStyle) {
    return scStyle = "selective" ? 1 : (scStyle = "generous" ? 3 : 2)
}

StudyCandidatesRecommendationSyncBasicSettings(scDialogState) {
    scDialogState["settings"]["learnerLevel"] := [
        "beginner", "intermediate", "advanced"
    ][scDialogState["levelDdl"].Value]
    scDialogState["settings"]["selectionStyle"] := [
        "selective", "balanced", "generous"
    ][scDialogState["styleDdl"].Value]
}

StudyCandidatesRecommendationDialogClose(scDialogState, scDialog, scResult, *) {
    if scDialogState["closed"]
        return
    ; Learner level and selection style describe which cached assessment set the
    ; review table should show. Preserve those choices even when the user closes
    ; this dialog without starting an API request.
    StudyCandidatesRecommendationSyncBasicSettings(scDialogState)
    StudyCandidatesRecommendationSaveSettings(
        scDialogState["settings"]
    )
    scDialogState["result"] := scResult
    scDialogState["closed"] := true
    try scDialog.Destroy()
}

StudyCandidatesRecommendationAdvancedDraft(scAdvanced) {
    return Map(
        "learnerLevel", scAdvanced["settings"]["learnerLevel"],
        "selectionStyle", scAdvanced["settings"]["selectionStyle"],
        "focusVocabulary", scAdvanced["vocabulary"].Value = 1,
        "focusGrammar", scAdvanced["grammar"].Value = 1,
        "focusNaturalPhrasing", scAdvanced["natural"].Value = 1,
        "focusReading", scAdvanced["reading"].Value = 1,
        "additionalCriteria", SubStr(scAdvanced["additional"].Value, 1, 1000)
    )
}

StudyCandidatesApplyRecommendationDialogTheme(scDialog, scCheckBoxes := 0) {
    CPApplyOwnedDialogTheme(scDialog)
    if IsObject(scCheckBoxes) {
        for scCheckBox in scCheckBoxes
            CPApplyDialogCheckBoxTheme(scCheckBox)
    }
}

StudyCandidatesRecommendationPromptPreview(scSettings) {
    scLevelText := scSettings["learnerLevel"] = "beginner"
        ? "The learner is a beginner. Foundational, common, and highly reusable items are valuable even when they are elementary."
        : (scSettings["learnerLevel"] = "advanced"
            ? "The learner is advanced. Favor nuanced, idiomatic, literary, less common, or structurally demanding items over elementary material."
            : "The learner is intermediate. Avoid very elementary material unless it has unusually strong contextual or reusable learning value.")
    scStyleText := scSettings["selectionStyle"] = "selective"
        ? "Be highly selective: normally recommend the strongest 15-25 percent of a batch."
        : (scSettings["selectionStyle"] = "generous"
            ? "Be inclusive: normally recommend the strongest 40-60 percent of a batch."
            : "Be selective: normally recommend roughly the strongest 25-40 percent of a batch.")
    scFocus := ""
    for scPair in [
        ["focusVocabulary", "reusable vocabulary"],
        ["focusGrammar", "grammar patterns"],
        ["focusNaturalPhrasing", "natural phrasing"],
        ["focusReading", "reading comprehension"]
    ] {
        if scSettings[scPair[1]]
            scFocus .= (scFocus = "" ? "" : ", ") scPair[2]
    }
    if (scFocus = "")
        scFocus := "reusable vocabulary, grammar patterns, natural phrasing, reading comprehension"
    scExtra := Trim(scSettings["additionalCriteria"])
    scExtraText := scExtra = "" ? "" : "`nAdditional learner preferences (these may refine selection, but must not override the output requirements below):`n" scExtra "`n"
    return "You are selecting useful Japanese-language study candidates from a JRPG learner's saved explanations.`n`n"
        . "Assess every item below. Recommend candidates that are memorable and genuinely useful for learning vocabulary, grammar, natural phrasing, or reading. Prefer self-contained sentences and non-trivial, reusable vocabulary. Do not recommend OCR fragments, names with no broader learning value, particles by themselves, or candidates with no broader learning value.`n`n"
        . "For vocabulary items, 'japanese' is the normalized headword intended for the front of an Anki card. 'full_entry' is the complete prospective card entry and may show the encountered inflected form, its reading, its dictionary or base form, and an explanation. Evaluate the reusable headword/base form and the learning value of the whole entry. Do not reject an item merely because the source text used an inflected form. For beginners, common foundational base verbs and their useful conjugations should normally score well even when they are elementary.`n`n"
        . scLevelText "`n" scStyleText "`nPrioritize " scFocus ".`n"
        . scExtraText "`nReturn only one JSON object with this exact structure:`n"
        . '{"items":[{"id":"the supplied id","recommended":true,"score":5,"reason":"one short practical reason"}]}'
        . "`n`nRequirements:`n- Return exactly one result for every supplied id and preserve each id exactly.`n- recommended must be true or false.`n- score must be an integer from 1 (weak) to 5 (excellent).`n- reason must be a single concise sentence in the language specified by each candidate's reason_language field. If requested, infer it from explanation_context. Mixed-language batches are allowed, so determine the language separately for every candidate.`n- Do not rewrite or translate the candidates.`n`nCandidates:`n[candidate data, reason_language, and a short language-context excerpt are inserted here]"
}

StudyCandidatesRecommendationPreviewClose(scPreview, scDialog, *) {
    scPreview["closed"] := true
    try scDialog.Destroy()
}

StudyCandidatesRecommendationShowPrompt(scAdvanced, *) {
    scSettings := StudyCandidatesRecommendationAdvancedDraft(scAdvanced)
    scOwner := scAdvanced["gui"].Hwnd
    scDialog := Gui(
        "+Owner" scOwner " +AlwaysOnTop +Resize",
        "Recommendation prompt preview"
    )
    scDialog.MarginX := 16, scDialog.MarginY := 14
    scDialog.SetFont("s10", "Segoe UI")
    scDialog.Add(
        "Text", "xm ym w700 h38",
        "This is the effective prompt. The candidate data and locked JSON rules cannot be edited here."
    )
    scEditor := scDialog.Add(
        "Edit", "xm y+8 w700 r24 ReadOnly Multi VScroll",
        StudyCandidatesRecommendationPromptPreview(scSettings)
    )
    scClose := scDialog.Add("Button", "xm y+12 w110 Default", "Close")
    scPreview := Map("closed", false)
    scClose.OnEvent(
        "Click", StudyCandidatesRecommendationPreviewClose.Bind(
            scPreview, scDialog
        )
    )
    scDialog.OnEvent(
        "Escape", StudyCandidatesRecommendationPreviewClose.Bind(
            scPreview, scDialog
        )
    )
    scDialog.OnEvent(
        "Close", StudyCandidatesRecommendationPreviewClose.Bind(
            scPreview, scDialog
        )
    )
    scDialog.Show("Hide AutoSize")
    StudyCandidatesApplyRecommendationDialogTheme(scDialog)
    try DllCall("user32\EnableWindow", "ptr", scOwner, "int", 0)
    try {
        scDialog.Show("Center")
        ; Native button faces and Edit scrollbars finish initialization only
        ; after the owned window becomes visible.
        StudyCandidatesApplyRecommendationDialogTheme(scDialog)
        scClose.Focus()
        while !scPreview["closed"]
            Sleep(25)
    } finally {
        try scDialog.Destroy()
        try DllCall("user32\EnableWindow", "ptr", scOwner, "int", 1)
        try WinActivate("ahk_id " scOwner)
    }
}

StudyCandidatesRecommendationAdvancedRestore(scAdvanced, *) {
    scDefaults := StudyCandidatesRecommendationDefaults()
    scAdvanced["settings"]["learnerLevel"] := scDefaults["learnerLevel"]
    scAdvanced["settings"]["selectionStyle"] := scDefaults["selectionStyle"]
    scAdvanced["vocabulary"].Value := 1
    scAdvanced["grammar"].Value := 1
    scAdvanced["natural"].Value := 1
    scAdvanced["reading"].Value := 1
    scAdvanced["additional"].Value := ""
}

StudyCandidatesRecommendationAdvancedClose(scAdvanced, scApply, *) {
    if scAdvanced["closed"]
        return
    if scApply {
        scDraft := StudyCandidatesRecommendationAdvancedDraft(scAdvanced)
        if !(scDraft["focusVocabulary"] || scDraft["focusGrammar"]
            || scDraft["focusNaturalPhrasing"] || scDraft["focusReading"]) {
            CPThemedOwnedMessage(
                scAdvanced["gui"].Hwnd,
                "Select at least one study focus.",
                "Recommendation preferences", "ok", "warning", 430
            )
            return
        }
        for scKey, scValue in scDraft
            scAdvanced["settings"][scKey] := scValue
    }
    scAdvanced["applied"] := scApply
    scAdvanced["closed"] := true
    try scAdvanced["gui"].Destroy()
}

StudyCandidatesRecommendationCustomize(scDialogState, *) {
    ; Customize and its prompt preview must use the selections currently
    ; visible in the confirmation dialog, even before they are saved.
    StudyCandidatesRecommendationSyncBasicSettings(scDialogState)
    scSettings := scDialogState["settings"]
    scOwner := scDialogState["gui"].Hwnd
    scDialog := Gui(
        "+Owner" scOwner " +AlwaysOnTop",
        "Recommendation preferences"
    )
    scDialog.MarginX := 18, scDialog.MarginY := 16
    scDialog.SetFont("s10", "Segoe UI")
    scDialog.Add("Text", "xm ym w560", "Study focus").SetFont("s11 Bold")
    scVocabulary := scDialog.Add(
        "CheckBox", "xm y+10 w260",
        "Reusable vocabulary"
    )
    scGrammar := scDialog.Add(
        "CheckBox", "x+12 yp w260",
        "Grammar patterns"
    )
    scNatural := scDialog.Add(
        "CheckBox", "xm y+8 w260",
        "Natural phrasing"
    )
    scReading := scDialog.Add(
        "CheckBox", "x+12 yp w260",
        "Reading comprehension"
    )
    scVocabulary.Value := scSettings["focusVocabulary"] ? 1 : 0
    scGrammar.Value := scSettings["focusGrammar"] ? 1 : 0
    scNatural.Value := scSettings["focusNaturalPhrasing"] ? 1 : 0
    scReading.Value := scSettings["focusReading"] ? 1 : 0
    scDialog.Add(
        "Text", "xm y+16 w560",
        "Additional selection guidance (optional)"
    ).SetFont("s10 Bold")
    scAdditional := scDialog.Add(
        "Edit", "xm y+7 w560 r4 Multi VScroll Limit1000",
        scSettings["additionalCriteria"]
    )
    scHint := scDialog.Add(
        "Text", "xm y+6 w560 h36 cGray",
        "Use this for preferences such as avoiding proper names or favoring dialogue. The output format remains protected."
    )
    CPRegisterMutedControl(scHint)
    scView := scDialog.Add("Button", "xm y+14 w150", "View full prompt...")
    scRestore := scDialog.Add("Button", "x+10 yp w130", "Restore defaults")
    scOK := scDialog.Add("Button", "x+26 yp w100 Default", "OK")
    scCancel := scDialog.Add("Button", "x+10 yp w100", "Cancel")
    scAdvanced := Map(
        "gui", scDialog,
        "settings", scSettings,
        "vocabulary", scVocabulary,
        "grammar", scGrammar,
        "natural", scNatural,
        "reading", scReading,
        "additional", scAdditional,
        "closed", false,
        "applied", false
    )
    scView.OnEvent(
        "Click", StudyCandidatesRecommendationShowPrompt.Bind(scAdvanced)
    )
    scRestore.OnEvent(
        "Click", StudyCandidatesRecommendationAdvancedRestore.Bind(scAdvanced)
    )
    scOK.OnEvent(
        "Click", StudyCandidatesRecommendationAdvancedClose.Bind(
            scAdvanced, true
        )
    )
    scCancel.OnEvent(
        "Click", StudyCandidatesRecommendationAdvancedClose.Bind(
            scAdvanced, false
        )
    )
    scDialog.OnEvent(
        "Escape", StudyCandidatesRecommendationAdvancedClose.Bind(
            scAdvanced, false
        )
    )
    scDialog.OnEvent(
        "Close", StudyCandidatesRecommendationAdvancedClose.Bind(
            scAdvanced, false
        )
    )
    scDialog.Show("Hide AutoSize")
    scRecommendationChecks := [
        scVocabulary, scGrammar, scNatural, scReading
    ]
    StudyCandidatesApplyRecommendationDialogTheme(
        scDialog, scRecommendationChecks
    )
    try DllCall("user32\EnableWindow", "ptr", scOwner, "int", 0)
    try {
        scDialog.Show("Center")
        StudyCandidatesApplyRecommendationDialogTheme(
            scDialog, scRecommendationChecks
        )
        scVocabulary.Focus()
        while !scAdvanced["closed"]
            Sleep(25)
    } finally {
        try scDialog.Destroy()
        try DllCall("user32\EnableWindow", "ptr", scOwner, "int", 1)
        try WinActivate("ahk_id " scOwner)
    }
    if scAdvanced["applied"] {
        scDialogState["levelDdl"].Choose(
            StudyCandidatesRecommendationLevelIndex(
                scSettings["learnerLevel"]
            )
        )
        scDialogState["styleDdl"].Choose(
            StudyCandidatesRecommendationStyleIndex(
                scSettings["selectionStyle"]
            )
        )
    }
}

StudyCandidatesRecommendationConfirm(
    scState, scProviderLabel, scModel, scCounts, scRegenerate := false
) {
    global controlDarkMode
    scOwner := scState["gui"].Hwnd
    scSettings := StudyCandidatesRecommendationLoadSettings()
    scDialog := Gui(
        "+Owner" scOwner " +AlwaysOnTop",
        "Generate recommendations"
    )
    scDialog.MarginX := 18, scDialog.MarginY := 16
    scColors := CPPalette(controlDarkMode)
    scDialog.BackColor := scColors["window"]
    scDialog.SetFont("s10 c" scColors["text"], "Segoe UI")
    scDialog.Add(
        "Text", "xm ym w620 c" scColors["text"],
        (scRegenerate ? "Regenerate" : "Generate cached")
            . " AI recommendations for " scCounts["total"]
            . " candidate" (scCounts["total"] = 1 ? "" : "s") "?"
    ).SetFont("s11 Bold c" scColors["text"])
    scBreakdown := scCounts["sentences"] " sentence"
        . (scCounts["sentences"] = 1 ? "" : "s") " and "
        . scCounts["vocabulary"] " vocabulary candidate"
        . (scCounts["vocabulary"] = 1 ? "" : "s")
    scDialog.Add(
        "Text", "xm y+7 w620 c" scColors["text"], scBreakdown
    )
    scInfo := scDialog.Add(
        "Text", "xm y+12 w620 h42 cGray",
        "Provider: " scProviderLabel "`nModel: " scModel
    )
    CPRegisterMutedControl(scInfo)
    scDialog.Add("Text", "xm y+12 w120", "Learner level:")
    scLevelDdl := scDialog.Add(
        "DropDownList", "x+10 yp-3 w190 0x210 Choose"
            . StudyCandidatesRecommendationLevelIndex(
                scSettings["learnerLevel"]
            ),
        ["Beginner", "Intermediate", "Advanced"]
    )
    scDialog.Add("Text", "x+24 yp+3 w110", "Selection style:")
    scStyleDdl := scDialog.Add(
        "DropDownList", "x+10 yp-3 w150 0x210 Choose"
            . StudyCandidatesRecommendationStyleIndex(
                scSettings["selectionStyle"]
            ),
        ["Selective", "Balanced", "Generous"]
    )
    scHintText := scRegenerate
        ? "The existing ratings for these candidates will be replaced using the selected settings. Existing Anki matches and globally ignored vocabulary stay excluded."
        : "Only candidates that have never been assessed are sent. Existing Anki matches and globally ignored vocabulary stay excluded. Results are cached, so refreshing or reopening this window uses no API calls."
    scHint := scDialog.Add(
        "Text", "xm y+14 w620 h58 cGray", scHintText
    )
    CPRegisterMutedControl(scHint)
    scCustomize := scDialog.Add("Button", "xm y+14 w130", "Customize...")
    scGenerate := scDialog.Add("Button", "x+166 yp w110 Default", "Generate")
    scCancel := scDialog.Add("Button", "x+10 yp w100", "Close")
    scDialogState := Map(
        "gui", scDialog,
        "settings", scSettings,
        "levelDdl", scLevelDdl,
        "styleDdl", scStyleDdl,
        "closed", false,
        "result", false
    )
    scCustomize.OnEvent(
        "Click", StudyCandidatesRecommendationCustomize.Bind(scDialogState)
    )
    scGenerate.OnEvent(
        "Click", StudyCandidatesRecommendationDialogClose.Bind(
            scDialogState, scDialog, true
        )
    )
    scCancel.OnEvent(
        "Click", StudyCandidatesRecommendationDialogClose.Bind(
            scDialogState, scDialog, false
        )
    )
    scDialog.OnEvent(
        "Escape", StudyCandidatesRecommendationDialogClose.Bind(
            scDialogState, scDialog, false
        )
    )
    scDialog.OnEvent(
        "Close", StudyCandidatesRecommendationDialogClose.Bind(
            scDialogState, scDialog, false
        )
    )
    scDialog.Show("Hide AutoSize")
    StudyCandidatesApplyRecommendationDialogTheme(scDialog)
    try DllCall("user32\EnableWindow", "ptr", scOwner, "int", 0)
    try {
        scDialog.Show("Center")
        ; Reapply after first show so Windows paints the final dark button
        ; faces, owner-drawn combo arrows and drop-down list scrollbars.
        StudyCandidatesApplyRecommendationDialogTheme(scDialog)
        scGenerate.Focus()
        while !scDialogState["closed"]
            Sleep(25)
    } finally {
        try scDialog.Destroy()
        try DllCall("user32\EnableWindow", "ptr", scOwner, "int", 1)
        try WinActivate("ahk_id " scOwner)
    }
    return scDialogState["result"] ? scSettings : 0
}

StudyCandidatesRunBridge(scState, scAction, scFront := "") {
    global pythonExe, explainProvider, explainOpenAIModel, explainGeminiModel
    scPython := ResolvePath(pythonExe)
    scBridge := A_ScriptDir "\scripts\anki_candidates.py"
    if !(FileExist(scPython) && FileExist(scBridge)) {
        CPThemedOwnedMessage(
            scState["gui"].Hwnd,
            "Review for Anki needs valid Python and candidate bridge paths.`n`n"
                . "Python:`n" scPython "`n`nBridge:`n" scBridge,
            "Review for Anki", "ok", "warning"
        )
        return false
    }
    scCommand := Format(
        '"{1}" "{2}" {3} --db "{4}" --output-dir "{5}" '
            . '--anki-config "{6}" --preferences-db "{7}"',
        scPython, scBridge, scAction, scState["database"],
        scState["outputDir"], A_ScriptDir "\Settings\anki.ini",
        A_ScriptDir "\Settings\anki_candidate_preferences.db"
    )
    if (scAction = "snapshot" || scAction = "generate-recommendations") {
        scProvider := CPSyncExplanationSelectionFromControls()
        scModel := scProvider = "gemini"
            ? explainGeminiModel : explainOpenAIModel
        scSettings := scState.Has("recommendationSettings")
            ? scState["recommendationSettings"]
            : StudyCandidatesRecommendationLoadSettings()
        if (scAction = "snapshot")
            scCommand .= " --scope " scState["scope"]
        scCommand .= ' --provider "' scProvider '" --model "' scModel '"'
            . ' --learner-level "' scSettings["learnerLevel"] '"'
            . ' --selection-style "' scSettings["selectionStyle"] '"'
            . ' --focus-areas "'
            . StudyCandidatesRecommendationFocusCsv(scSettings) '"'
            . ' --additional-criteria-hex "'
            . StudyLibraryHexEncode(scSettings["additionalCriteria"]) '"'
        if (scAction = "generate-recommendations"
            && scState.Has("forceRecommendations")
            && scState["forceRecommendations"])
            scCommand .= " --force"
    }
    else if (scAction = "finish-review")
        scCommand .= ' --reviewed-through "' scState["snapshotThrough"] '"'
    else if (scAction = "hide-vocabulary" || scAction = "restore-vocabulary")
        scCommand .= ' --front-hex "' StudyLibraryHexEncode(scFront) '"'
    if !StudyCandidatesBeginWork(scState)
        return false
    try {
        try scExitCode := RunWait(scCommand, A_ScriptDir, "Hide")
        catch as scError {
            if (!scState.Has("closeRequested") || !scState["closeRequested"])
                && StudyCandidatesGuiAlive(scState) {
                CPThemedOwnedMessage(
                    scState["gui"].Hwnd,
                    "The Anki candidate list could not be updated.`n`n" scError.Message,
                    "Review for Anki", "ok", "warning"
                )
            }
            return false
        }
    } finally {
        StudyCandidatesEndWork(scState)
    }
    if (scExitCode != 0) {
        scDetails := ""
        if (scAction = "generate-recommendations") {
            scErrorPath := scState["outputDir"] "\candidate_recommendation_error.txt"
            try scDetails := Trim(FileRead(scErrorPath, "UTF-8"))
        }
        if (!scState.Has("closeRequested") || !scState["closeRequested"])
            && StudyCandidatesGuiAlive(scState) {
            CPThemedOwnedMessage(
                scState["gui"].Hwnd,
                "The Anki candidate list could not be updated (bridge exit "
                    . scExitCode ")." (scDetails != "" ? "`n`n" scDetails : ""),
                "Review for Anki", "ok", "warning"
            )
        }
        return false
    }
    return true
}

StudyCandidatesGuiAlive(scState) {
    if !IsObject(scState) || !scState.Has("gui") || !scState["gui"]
        return false
    try return DllCall(
        "user32\IsWindow", "ptr", scState["gui"].Hwnd, "int"
    ) != 0
    catch
        return false
}

StudyCandidatesBeginWork(scState) {
    if !StudyCandidatesGuiAlive(scState)
        || (scState.Has("closeRequested") && scState["closeRequested"])
        return false
    scState["busyCount"] := (scState.Has("busyCount")
        ? scState["busyCount"] : 0) + 1
    return true
}

StudyCandidatesEndWork(scState) {
    if !IsObject(scState)
        return
    scState["busyCount"] := Max(0, (scState.Has("busyCount")
        ? scState["busyCount"] : 1) - 1)
    if (scState["busyCount"] = 0 && scState.Has("closeRequested")
        && scState["closeRequested"])
        SetTimer(StudyCandidatesFinalizeClose.Bind(scState), -1)
}

StudyCandidatesSelected(scState) {
    scVocabulary := scState["tabs"].Value = 2
    scList := scVocabulary ? scState["vocabularyList"] : scState["sentenceList"]
    scRows := scVocabulary ? scState["vocabulary"] : scState["sentences"]
    scRow := scList.GetNext(0, "F")
    if !scRow
        scRow := scList.GetNext()
    if (scRow < 1 || scRow > scRows.Length)
        return 0
    return scRows[scRow]
}

StudyCandidatesUpdateActions(scState, *) {
    if !StudyCandidatesGuiAlive(scState)
        return
    scSelected := StudyCandidatesSelected(scState)
    scEnabled := IsObject(scSelected)
    scHidden := scState["scope"] = "hidden"
    scVocabulary := scState["tabs"].Value = 2
    scState["openButton"].Enabled := scEnabled && !scHidden
    scState["addButton"].Enabled := scEnabled && !scHidden
    scState["finishButton"].Enabled := scState["scope"] = "new"
        && scState["snapshotThrough"] != ""
    scState["triageButton"].Text := scHidden
        ? "Restore vocabulary" : "Ignore vocabulary..."
    scState["triageButton"].Enabled := scEnabled && scVocabulary
    scHasCandidates := scState["allSentences"].Length
        + scState["allVocabulary"].Length > 0
    scState["recommendButton"].Enabled := !scHidden && scHasCandidates
    scState["recommendButton"].Text := scHasCandidates
        && scState["unassessedCount"] < 1
        ? "Regenerate recommendations..."
        : "Generate recommendations..."
    scStatusText := scState.Has("baseStatus") ? scState["baseStatus"] : ""
    scAiStatusText := scState.Has("noticeStatus")
        ? scState["noticeStatus"] : ""
    if scEnabled && scSelected.Has("recommendation")
        && scSelected["recommendation"] >= 0 {
        scAiText := scSelected["recommendation"] = 1
            ? "Recommended" : "Not recommended"
        if (scSelected["score"] > 0)
            scAiText .= " (" scSelected["score"] "/5)"
        if (scSelected["reason"] != "")
            scAiText .= " — " scSelected["reason"]
        scAiStatusText := "AI: " scAiText
    } else if scEnabled
        scAiStatusText := "AI: Not yet assessed."
    scState["status"].Value := scStatusText
    scState["aiStatus"].Value := scAiStatusText
}

StudyCandidatesAiResult(scCandidate) {
    scRecommendation := scCandidate["recommendation"]
    if (scRecommendation < 0)
        return "Unassessed"
    scScore := scCandidate["score"] > 0
        ? " " scCandidate["score"] "/5" : ""
    return scRecommendation = 1 ? "★" scScore : "No ·" scScore
}

StudyCandidatesFilteredRows(scRows, scAiFilter) {
    scFiltered := []
    for scCandidate in scRows {
        scRecommendation := scCandidate["recommendation"]
        if (scAiFilter = 2 && scRecommendation != 1)
            continue
        if (scAiFilter = 3 && scRecommendation != 0)
            continue
        if (scAiFilter = 4 && scRecommendation >= 0)
            continue
        scFiltered.Push(scCandidate)
    }
    ; Within the assessed-only views, show the strongest candidates first while
    ; retaining the original date order among entries with an equal score.
    if (scAiFilter != 2 && scAiFilter != 3)
        return scFiltered
    scSorted := []
    Loop 5 {
        scWantedScore := 6 - A_Index
        for scCandidate in scFiltered
            if (scCandidate["score"] = scWantedScore)
                scSorted.Push(scCandidate)
    }
    for scCandidate in scFiltered
        if (scCandidate["score"] < 1 || scCandidate["score"] > 5)
            scSorted.Push(scCandidate)
    return scSorted
}

StudyCandidatesApplyAiFilter(scState) {
    if !StudyCandidatesGuiAlive(scState)
        return
    scState["sentences"] := StudyCandidatesFilteredRows(
        scState["allSentences"], scState["aiFilter"]
    )
    scState["vocabulary"] := StudyCandidatesFilteredRows(
        scState["allVocabulary"], scState["aiFilter"]
    )
    scState["sentenceList"].Delete()
    scState["vocabularyList"].Delete()
    for scCandidate in scState["sentences"]
        scState["sentenceList"].Add(
            "", StudyCandidatesAiResult(scCandidate), scCandidate["date"],
            scCandidate["profile"], scCandidate["source"],
            scCandidate["versions"]
        )
    for scCandidate in scState["vocabulary"]
        scState["vocabularyList"].Add(
            "", StudyCandidatesAiResult(scCandidate), scCandidate["display"],
            scCandidate["meaning"], scCandidate["occurrences"],
            scCandidate["profile"], scCandidate["date"]
        )

    scScopeText := scState["scope"] = "new"
        ? "New since last review"
        : (scState["scope"] = "hidden" ? "Ignored vocabulary" : "All backlog")
    scAiText := scState["aiFilter"] = 2
        ? "Recommended" : (scState["aiFilter"] = 3
            ? "Not recommended" : (scState["aiFilter"] = 4
                ? "Not yet assessed" : ""))
    if (scAiText != "")
        scScopeText .= " / " scAiText
    scState["baseStatus"] := scScopeText ": "
        . scState["sentences"].Length " sentence"
        . (scState["sentences"].Length = 1 ? "" : "s") ", "
        . scState["vocabulary"].Length " vocabulary candidate"
        . (scState["vocabulary"].Length = 1 ? "" : "s") ".  "
        . scState["ankiMessage"]
    if (scState["scope"] != "hidden")
        scState["baseStatus"] .= "  " scState["unassessedCount"]
            . " not yet assessed by AI."

    if scState["sentences"].Length
        scState["sentenceList"].Modify(1, "Select Vis")
    if scState["vocabulary"].Length
        scState["vocabularyList"].Modify(1, "Select Vis")
    scActiveList := scState["tabs"].Value = 2
        ? scState["vocabularyList"] : scState["sentenceList"]
    scActiveRows := scState["tabs"].Value = 2
        ? scState["vocabulary"] : scState["sentences"]
    if scActiveRows.Length {
        scActiveList.Modify(1, "Focus Vis")
        scActiveList.Focus()
    }
    StudyCandidatesUpdateActions(scState)
}

StudyCandidatesReadSnapshot(scState) {
    if !StudyCandidatesGuiAlive(scState)
        || (scState.Has("closeRequested") && scState["closeRequested"])
        return false
    scState["allSentences"] := []
    scState["allVocabulary"] := []

    for scRow in StudyLibraryReadRows(
        scState["outputDir"] "\candidate_sentences.tsv"
    ) {
        if (scRow.Length < 9)
            continue
        scCandidate := Map(
            "groupId", Integer(scRow[1]),
            "profile", StudyLibraryHexDecode(scRow[2]),
            "date", StudyLibraryHexDecode(scRow[3]),
            "source", StudyLibraryHexDecode(scRow[4]),
            "versions", Integer(scRow[5]),
            "recommendationKey", scRow[6],
            "recommendation", Integer(scRow[7]),
            "score", Integer(scRow[8]),
            "reason", StudyLibraryHexDecode(scRow[9])
        )
        scState["allSentences"].Push(scCandidate)
    }
    for scRow in StudyLibraryReadRows(
        scState["outputDir"] "\candidate_vocabulary.tsv"
    ) {
        if (scRow.Length < 12)
            continue
        scCandidate := Map(
            "groupId", Integer(scRow[1]),
            "profile", StudyLibraryHexDecode(scRow[2]),
            "date", StudyLibraryHexDecode(scRow[3]),
            "display", StudyLibraryHexDecode(scRow[4]),
            "front", StudyLibraryHexDecode(scRow[5]),
            "back", StudyLibraryHexDecode(scRow[6]),
            "meaning", StudyLibraryHexDecode(scRow[7]),
            "occurrences", Integer(scRow[8]),
            "recommendationKey", scRow[9],
            "recommendation", Integer(scRow[10]),
            "score", Integer(scRow[11]),
            "reason", StudyLibraryHexDecode(scRow[12])
        )
        scState["allVocabulary"].Push(scCandidate)
    }
    scStatusRows := StudyLibraryReadRows(
        scState["outputDir"] "\candidate_status.tsv"
    )
    if scStatusRows.Length && scStatusRows[1].Length >= 6 {
        scRow := scStatusRows[1]
        scState["snapshotThrough"] := StudyLibraryHexDecode(scRow[1])
        scState["reviewedThrough"] := StudyLibraryHexDecode(scRow[2])
        scState["ankiCode"] := StudyLibraryHexDecode(scRow[3])
        scState["ankiMessage"] := StudyLibraryHexDecode(scRow[4])
    }
    scState["unassessedCount"] := 0
    for scCandidate in scState["allSentences"]
        if (scCandidate["recommendation"] < 0)
            scState["unassessedCount"] += 1
    for scCandidate in scState["allVocabulary"]
        if (scCandidate["recommendation"] < 0)
            scState["unassessedCount"] += 1
    scState["noticeStatus"] := ""
    if scState.Has("recommendationMessage")
        && scState["recommendationMessage"] != "" {
        scState["noticeStatus"] := scState["recommendationMessage"]
        scState["recommendationMessage"] := ""
    }
    StudyCandidatesApplyAiFilter(scState)
    return true
}

StudyCandidatesRefresh(scState, *) {
    if !StudyCandidatesGuiAlive(scState)
        || (scState.Has("closeRequested") && scState["closeRequested"])
        || (scState.Has("busyCount") && scState["busyCount"] > 0)
        return
    try {
        scScopeValue := scState["scopeDdl"].Value
        scState["scope"] := scScopeValue = 3
            ? "hidden" : (scScopeValue = 2 ? "all" : "new")
        if (scState["scope"] = "hidden" && scState["tabs"].Value != 2)
            scState["tabs"].Choose(2)
        scState["status"].Value := "Building local candidate lists..."
        scState["refreshButton"].Enabled := false
        if StudyCandidatesRunBridge(scState, "snapshot")
            && !scState["closeRequested"]
            StudyCandidatesReadSnapshot(scState)
    } finally {
        if StudyCandidatesGuiAlive(scState) && !scState["closeRequested"]
            scState["refreshButton"].Enabled := true
    }
}

StudyCandidatesScopeChanged(scState, *) {
    StudyCandidatesRefresh(scState)
}

StudyCandidatesAiFilterChanged(scState, *) {
    scState["aiFilter"] := scState["aiFilterDdl"].Value
    StudyCandidatesApplyAiFilter(scState)
}

StudyCandidatesTabChanged(scState, *) {
    StudyCandidatesUpdateActions(scState)
    scList := scState["tabs"].Value = 2
        ? scState["vocabularyList"] : scState["sentenceList"]
    try scList.Focus()
}

StudyCandidatesGenerateRecommendations(scState, *) {
    global explainProvider, explainOpenAIModel, explainGeminiModel
    if !StudyCandidatesGuiAlive(scState)
        || (scState.Has("closeRequested") && scState["closeRequested"])
        || (scState.Has("busyCount") && scState["busyCount"] > 0)
        return
    scCandidateCount := scState["allSentences"].Length
        + scState["allVocabulary"].Length
    if (scState["scope"] = "hidden" || scCandidateCount < 1)
        return
    scProvider := CPSyncExplanationSelectionFromControls()
    scModel := scProvider = "gemini"
        ? explainGeminiModel : explainOpenAIModel
    if !CPApiKeyConfigured(scProvider) {
        CPThemedOwnedMessage(
            scState["gui"].Hwnd,
            CPMissingApiKeyText(scProvider),
            "Generate recommendations", "ok", "warning", 620
        )
        return
    }
    scProviderLabel := scProvider = "gemini" ? "Gemini" : "OpenAI"
    scRegenerate := scState["unassessedCount"] < 1
    scCounts := StudyCandidatesRecommendationCounts(scState, scRegenerate)
    if (scCounts["total"] < 1)
        return
    scRecommendationSettings := StudyCandidatesRecommendationConfirm(
        scState, scProviderLabel, scModel, scCounts, scRegenerate
    )
    if !IsObject(scRecommendationSettings) {
        ; Closing the generation dialog still commits its learner preferences.
        ; Rebuild the local snapshot so the tables immediately show the cache
        ; belonging to that selection (or Unassessed if none exists yet).
        scState["recommendationSettings"] :=
            StudyCandidatesRecommendationLoadSettings()
        if StudyCandidatesRunBridge(scState, "snapshot")
            && !scState["closeRequested"]
            StudyCandidatesReadSnapshot(scState)
        return
    }
    scState["recommendationSettings"] := scRecommendationSettings
    scState["forceRecommendations"] := scRegenerate

    ; A changed learner level, model or recommendation preference represents a
    ; different assessment profile. Rebuild the snapshot so cached ratings from
    ; another profile are neither reused nor skipped by generation.
    if !StudyCandidatesRunBridge(scState, "snapshot") {
        scState["forceRecommendations"] := false
        return
    }
    if scState["closeRequested"] {
        scState["forceRecommendations"] := false
        return
    }
    StudyCandidatesReadSnapshot(scState)
    scCounts := StudyCandidatesRecommendationCounts(scState)
    if (scCounts["total"] < 1 && !scRegenerate) {
        scState["forceRecommendations"] := false
        return
    }

    scState["baseStatus"] := "Generating recommendations with "
        . scProviderLabel " / " scModel "..."
    scState["status"].Value := scState["baseStatus"]
    scState["refreshButton"].Enabled := false
    scState["recommendButton"].Enabled := false
    try DllCall("user32\UpdateWindow", "ptr", scState["gui"].Hwnd)
    try {
        if StudyCandidatesRunBridge(scState, "generate-recommendations") {
            if scState["closeRequested"]
                return
            scRows := StudyLibraryReadRows(
                scState["outputDir"] "\candidate_recommendation_status.tsv"
            )
            if scRows.Length && scRows[1].Length >= 7
                scState["recommendationMessage"] := StudyLibraryHexDecode(
                    scRows[1][7]
                )
            StudyCandidatesRefresh(scState)
        }
    } finally {
        scState["forceRecommendations"] := false
        if StudyCandidatesGuiAlive(scState) && !scState["closeRequested"] {
            scState["refreshButton"].Enabled := true
            StudyCandidatesUpdateActions(scState)
        }
    }
}

StudyCandidatesReaderSequence(scState) {
    scRows := scState["tabs"].Value = 2
        ? scState["vocabulary"] : scState["sentences"]
    scSequence := []
    scSeen := Map()
    for scCandidate in scRows {
        scId := scCandidate["groupId"]
        if scSeen.Has(scId)
            continue
        scSeen[scId] := true
        scSequence.Push(Map(
            "id", scId,
            "addedToAnkiAt", "",
            "ankiStatus", "not_checked"
        ))
    }
    return scSequence
}

StudyCandidatesOpenSelected(scState, *) {
    if (scState["scope"] = "hidden")
        return
    scCandidate := StudyCandidatesSelected(scState)
    if !IsObject(scCandidate)
        return
    OpenStudyReader(
        scCandidate["groupId"], 0, StudyCandidatesReaderSequence(scState)
    )
}

StudyCandidatesAddSelected(scState, *) {
    global CPStudyReaderState
    if (scState["scope"] = "hidden")
        return
    scCandidate := StudyCandidatesSelected(scState)
    if !IsObject(scCandidate)
        return
    OpenStudyReader(
        scCandidate["groupId"], 0, StudyCandidatesReaderSequence(scState)
    )
    if !(CPStudyReaderState && CPStudyReaderState.Has("gui"))
        return
    if (scState["tabs"].Value = 2) {
        StudyReaderOpenReviewedAnkiDialog(
            CPStudyReaderState, "vocabulary",
            scCandidate["front"], scCandidate["back"]
        )
    } else
        StudyReaderOpenAnkiAddDialog(CPStudyReaderState)
}

StudyCandidatesFinishReview(scState, *) {
    if !StudyCandidatesGuiAlive(scState)
        || (scState.Has("closeRequested") && scState["closeRequested"])
        || (scState.Has("busyCount") && scState["busyCount"] > 0)
        return
    if (scState["scope"] != "new" || scState["snapshotThrough"] = "")
        return
    scMessage := "Mark everything included in this candidate snapshot as reviewed?`n`n"
        . "This does not add, edit, hide, or delete anything. Explanations generated "
        . "after this snapshot will remain under New since last review."
    if (CPThemedOwnedMessage(
        scState["gui"].Hwnd, scMessage, "Finish candidate review",
        "yesno", "info", 580
    ) != "Yes")
        return
    if StudyCandidatesRunBridge(scState, "finish-review")
        StudyCandidatesRefresh(scState)
}

StudyCandidatesTriageSelected(scState, *) {
    if !StudyCandidatesGuiAlive(scState)
        || (scState.Has("closeRequested") && scState["closeRequested"])
        || (scState.Has("busyCount") && scState["busyCount"] > 0)
        return
    if (scState["tabs"].Value != 2)
        return
    scCandidate := StudyCandidatesSelected(scState)
    if !IsObject(scCandidate)
        return
    scHidden := scState["scope"] = "hidden"
    if scHidden {
        scMessage := "Restore “" scCandidate["front"] "” to vocabulary candidate "
            . "lists in every Study Library?"
        scTitle := "Restore vocabulary candidate"
        scAction := "restore-vocabulary"
    } else {
        scMessage := "Ignore “" scCandidate["front"] "” in future vocabulary "
            . "candidate lists across every Study Library?`n`n"
            . "The explanation and screenshots stay unchanged. You can restore "
            . "the term later from Show: Ignored vocabulary."
        scTitle := "Ignore vocabulary candidate"
        scAction := "hide-vocabulary"
    }
    if (CPThemedOwnedMessage(
        scState["gui"].Hwnd, scMessage, scTitle, "yesno", "info", 580
    ) != "Yes")
        return
    if StudyCandidatesRunBridge(scState, scAction, scCandidate["front"])
        StudyCandidatesRefresh(scState)
}

StudyCandidatesContextMenu(
    scState, scGui, scControl, scRow, scIsRightClick, scX, scY
) {
    ; Gui.ContextMenu also fires for the keyboard context-menu key. Restrict
    ; this handler to the two candidate tables, but support both invocation
    ; methods so keyboard users receive the same commands as mouse users.
    try scControlHwnd := scControl.Hwnd
    catch
        return
    scSentenceHwnd := scState["sentenceList"].Hwnd
    scVocabularyHwnd := scState["vocabularyList"].Hwnd
    if (scControlHwnd != scSentenceHwnd
        && scControlHwnd != scVocabularyHwnd)
        return

    scVocabulary := scControlHwnd = scVocabularyHwnd
    scExpectedTab := scVocabulary ? 2 : 1
    if (scState["tabs"].Value != scExpectedTab)
        scState["tabs"].Choose(scExpectedTab)

    if (scRow > 0) {
        ; These ListViews are single-selection tables. Select the row under the
        ; pointer before building the menu, as Explorer and spreadsheet tools do.
        scControl.Modify(0, "-Select")
        scControl.Modify(scRow, "Select Focus Vis")
        StudyCandidatesUpdateActions(scState)
    }

    scCandidate := StudyCandidatesSelected(scState)
    scHasSelection := IsObject(scCandidate)
    scHidden := scState["scope"] = "hidden"
    scMenuItems := []
    scHasRowCommands := false

    if (scHasSelection && !scHidden) {
        scMenuItems.Push(Map(
            "label", "Open in Reader",
            "action", StudyCandidatesOpenSelected.Bind(scState)
        ))
        scMenuItems.Push(Map(
            "label", "Add to Anki...",
            "action", StudyCandidatesAddSelected.Bind(scState)
        ))
        scHasRowCommands := true
    }
    if (scHasSelection && scVocabulary) {
        if scHasRowCommands
            scMenuItems.Push(Map("separator", true))
        scTriageLabel := scHidden
            ? "Restore vocabulary" : "Ignore vocabulary..."
        scMenuItems.Push(Map(
            "label", scTriageLabel,
            "action", StudyCandidatesTriageSelected.Bind(scState)
        ))
        scHasRowCommands := true
    }
    if (scState["scope"] = "new" && scState["snapshotThrough"] != "") {
        if scHasRowCommands
            scMenuItems.Push(Map("separator", true))
        scMenuItems.Push(Map(
            "label", "Finish review",
            "action", StudyCandidatesFinishReview.Bind(scState)
        ))
        scHasRowCommands := true
    }
    if (scState["recommendButton"].Enabled) {
        if scHasRowCommands
            scMenuItems.Push(Map("separator", true))
        scMenuItems.Push(Map(
            "label", scState["recommendButton"].Text,
            "action", StudyCandidatesGenerateRecommendations.Bind(scState)
        ))
        scHasRowCommands := true
    }
    if scHasRowCommands
        scMenuItems.Push(Map("separator", true))
    scMenuItems.Push(Map(
        "label", "Refresh",
        "action", StudyCandidatesRefresh.Bind(scState),
        "enabled", !(scState.Has("busyCount") && scState["busyCount"] > 0)
    ))
    scChoice := CPThemedContextPopup(
        scGui.Hwnd, scX, scY, scMenuItems
    )
    if (scChoice >= 1 && scChoice <= scMenuItems.Length
        && scMenuItems[scChoice].Has("action"))
        scMenuItems[scChoice]["action"].Call()
}

StudyCandidatesResize(scState, scGui, scMinMax, scWidth, scHeight) {
    if (scMinMax = -1 || scWidth < 680 || scHeight < 420)
        return
    scMargin := 14, scGap := 10
    scScopeLabelW := 46, scScopeW := 190, scRefreshW := 72
    scRecommendW := 218, scAiLabelW := 24, scAiFilterW := 160
    scState["scopeLabel"].Move(scMargin, scMargin + 5, scScopeLabelW, 22)
    scState["scopeDdl"].Move(
        scMargin + scScopeLabelW, scMargin, scScopeW, 30
    )
    scState["refreshButton"].Move(
        scMargin + scScopeLabelW + scScopeW + scGap,
        scMargin, scRefreshW, 30
    )
    scState["recommendButton"].Move(
        scMargin + scScopeLabelW + scScopeW + scGap + scRefreshW + scGap,
        scMargin, scRecommendW, 30
    )
    scAiX := scMargin + scScopeLabelW + scScopeW + scGap + scRefreshW
        + scGap + scRecommendW + scGap
    scState["aiFilterLabel"].Move(scAiX, scMargin + 5, scAiLabelW, 22)
    scState["aiFilterDdl"].Move(
        scAiX + scAiLabelW, scMargin, scAiFilterW, 30
    )
    scTabY := 54
    scBottomH := 120
    scTabH := Max(232, scHeight - scTabY - scBottomH - scMargin)
    scTabW := Max(640, scWidth - scMargin * 2)
    scState["tabs"].Move(scMargin, scTabY, scTabW, scTabH)
    scListX := scMargin + 8, scListY := scTabY + 31
    scListW := Max(620, scTabW - 16), scListH := Max(192, scTabH - 39)
    scState["sentenceList"].Move(scListX, scListY, scListW, scListH)
    scState["vocabularyList"].Move(scListX, scListY, scListW, scListH)
    scActionY := scTabY + scTabH + 8
    scState["openButton"].Move(scMargin, scActionY, 120, 30)
    scState["addButton"].Move(scMargin + 130, scActionY, 125, 30)
    scState["finishButton"].Move(scMargin + 265, scActionY, 112, 30)
    scState["triageButton"].Move(scMargin + 387, scActionY, 140, 30)
    scState["closeButton"].Move(scMargin + 537, scActionY, 82, 30)
    scState["status"].Move(
        scMargin, scActionY + 37, Max(620, scWidth - scMargin * 2), 22
    )
    scState["aiStatus"].Move(
        scMargin, scActionY + 59, Max(620, scWidth - scMargin * 2), 50
    )
}

StudyCandidatesFinalizeClose(scState, *) {
    global CPStudyCandidateState
    if !IsObject(scState)
        return
    if scState.Has("closeFinalized") && scState["closeFinalized"]
        return
    scState["closeFinalized"] := true
    try CPListViewSaveColumnOrder(
        scState["sentenceList"], "study_candidate_view",
        "sentenceColumnOrder", 5
    )
    try CPListViewSaveColumnOrder(
        scState["vocabularyList"], "study_candidate_view",
        "vocabularyColumnOrder", 6
    )
    try scState["gui"].Destroy()
    if scState.Has("outputDir")
        try StudyBridgeCleanupWorkspace(scState["outputDir"])
    if IsObject(CPStudyCandidateState)
        && ObjPtr(CPStudyCandidateState) = ObjPtr(scState)
        CPStudyCandidateState := 0
}

StudyCandidatesClose(scState, *) {
    global CPStudyCandidateState
    if !IsObject(scState)
        return
    if scState.Has("closeRequested") && scState["closeRequested"]
        return
    scState["closeRequested"] := true
    ; Remove this closing instance from the public slot immediately, allowing a
    ; newly selected Study Library to open its own candidate window while the
    ; old bridge invocation winds down in the background.
    if IsObject(CPStudyCandidateState)
        && ObjPtr(CPStudyCandidateState) = ObjPtr(scState)
        CPStudyCandidateState := 0
    ; RunWait yields to GUI events, so Close can arrive while the Python bridge
    ; is active. Hide immediately but keep the native controls alive until the
    ; interrupted refresh thread returns and can finish without touching a
    ; destroyed HWND.
    if scState.Has("busyCount") && scState["busyCount"] > 0 {
        try scState["gui"].Hide()
        return
    }
    StudyCandidatesFinalizeClose(scState)
}

StudyCandidatesInitialRefresh(scState, *) {
    ; Run the Python/Anki snapshot only after the button event which created the
    ; native window has returned.  This prevents Windows owner-draw callbacks
    ; from being nested inside a half-finished GUI construction thread.
    if !StudyCandidatesGuiAlive(scState)
        || (scState.Has("closeRequested") && scState["closeRequested"])
        return
    StudyCandidatesRefresh(scState)
    if StudyCandidatesGuiAlive(scState)
        && !(scState.Has("closeRequested") && scState["closeRequested"])
        CPApplyOwnedDialogTheme(scState["gui"])
}

StudyLibraryOpenCandidates(slState, *) {
    global CPStudyCandidateState, CPStudyCandidateOpening, controlDarkMode
    if CPStudyCandidateOpening
        return
    if IsObject(CPStudyCandidateState) {
        if !StudyCandidatesGuiAlive(CPStudyCandidateState)
            || (CPStudyCandidateState.Has("closeRequested")
                && CPStudyCandidateState["closeRequested"]) {
            ; A destroyed native HWND must never be reused merely because its
            ; old AHK state object is still reachable.
            CPStudyCandidateState := 0
        }
    }
    if IsObject(CPStudyCandidateState) && CPStudyCandidateState.Has("gui") {
        try {
            CPStudyCandidateState["gui"].Show()
            WinActivate("ahk_id " CPStudyCandidateState["gui"].Hwnd)
            StudyCandidatesRefresh(CPStudyCandidateState)
            return
        }
        CPStudyCandidateState := 0
    }
    CPStudyCandidateOpening := true
    try {
    scOutputDir := StudyBridgeCreateWorkspace("anki_candidates")
    scGui := Gui(
        "+Owner" slState["gui"].Hwnd " +OwnDialogs +Resize +MinSize680x420",
        "Study Library - Review for Anki"
    )
    scGui.MarginX := 14, scGui.MarginY := 14
    scGui.SetFont("s10", "Segoe UI")
    scColors := CPPalette(controlDarkMode)
    scGui.BackColor := scColors["window"]
    CPApplyOwnedDialogTheme(scGui)
    scScopeLabel := scGui.Add("Text", "x14 y19 w46", "Show:")
    scScopeDdl := scGui.Add(
        "DropDownList", "x60 y14 w190 0x210",
        ["New since last review", "All backlog", "Ignored vocabulary"]
    )
    scScopeDdl.Choose(1)
    scRefresh := scGui.Add("Button", "x260 y14 w72 h30", "Refresh")
    scRecommend := scGui.Add(
        "Button", "x342 y14 w218 h30 Disabled", "Generate recommendations..."
    )
    scAiFilterLabel := scGui.Add("Text", "x570 y19 w24", "AI:")
    scAiFilterDdl := scGui.Add(
        "DropDownList", "x594 y14 w160 0x210",
        ["All statuses", "Recommended only", "Not recommended", "Not yet assessed"]
    )
    scAiFilterDdl.Choose(1)
    scTabs := scGui.Add("Tab3", "x14 y54 w812 h438", ["Sentences", "Vocabulary"])
    scTabs.SetFont("c" scColors["text"])
    scTabs.UseTab("Sentences")
    scSentenceList := scGui.Add(
        "ListView", "x22 y85 w796 h399 Grid Background"
            . scColors["surface"] " c" scColors["text"],
        ["AI result", "Date generated", "Profile", "Japanese source", "Versions"]
    )
    SendMessage(0x1036, 0x10000, 0x10000, scSentenceList.Hwnd)
    scSentenceList.ModifyCol(1, 88), scSentenceList.ModifyCol(2, 125)
    scSentenceList.ModifyCol(3, 125), scSentenceList.ModifyCol(4, 380)
    scSentenceList.ModifyCol(5, 70)
    scTabs.UseTab("Vocabulary")
    scVocabularyList := scGui.Add(
        "ListView", "x22 y85 w796 h399 Grid Background"
            . scColors["surface"] " c" scColors["text"],
        ["AI result", "Vocabulary", "Meaning / context", "Occurrences", "Profile", "Latest"]
    )
    SendMessage(0x1036, 0x10000, 0x10000, scVocabularyList.Hwnd)
    scVocabularyList.ModifyCol(1, 88), scVocabularyList.ModifyCol(2, 190)
    scVocabularyList.ModifyCol(3, 260), scVocabularyList.ModifyCol(4, 82)
    scVocabularyList.ModifyCol(5, 115), scVocabularyList.ModifyCol(6, 125)
    CPListViewRestoreColumnOrder(
        scSentenceList, "study_candidate_view", "sentenceColumnOrder", 5
    )
    CPListViewRestoreColumnOrder(
        scVocabularyList, "study_candidate_view", "vocabularyColumnOrder", 6
    )
    scTabs.UseTab()
    scOpen := scGui.Add("Button", "x14 y500 w120 h30 Disabled", "Open in Reader")
    scAdd := scGui.Add("Button", "x144 y500 w125 h30 Disabled", "Add to Anki...")
    scFinish := scGui.Add("Button", "x279 y500 w112 h30 Disabled", "Finish review")
    scTriage := scGui.Add(
        "Button", "x401 y500 w140 h30 Disabled", "Ignore vocabulary..."
    )
    scClose := scGui.Add("Button", "x551 y500 w82 h30", "Close")
    scStatus := scGui.Add(
        "Text", "x14 y537 w812 h24 cGray", "Building local candidate lists..."
    )
    CPRegisterMutedControl(scStatus)
    scAiStatus := scGui.Add("Text", "x14 y559 w812 h50 cGray", "")
    CPRegisterMutedControl(scAiStatus)
    scState := Map(
        "gui", scGui,
        "libraryState", slState,
        "database", slState["database"],
        "outputDir", scOutputDir,
        "scopeLabel", scScopeLabel,
        "scopeDdl", scScopeDdl,
        "refreshButton", scRefresh,
        "recommendButton", scRecommend,
        "aiFilterLabel", scAiFilterLabel,
        "aiFilterDdl", scAiFilterDdl,
        "tabs", scTabs,
        "sentenceList", scSentenceList,
        "vocabularyList", scVocabularyList,
        "openButton", scOpen,
        "addButton", scAdd,
        "finishButton", scFinish,
        "triageButton", scTriage,
        "closeButton", scClose,
        "status", scStatus,
        "aiStatus", scAiStatus,
        "scope", "new",
        "aiFilter", 1,
        "snapshotThrough", "",
        "reviewedThrough", "",
        "ankiCode", "",
        "ankiMessage", "",
        "baseStatus", "",
        "noticeStatus", "",
        "recommendationMessage", "",
        "busyCount", 0,
        "closeRequested", false,
        "closeFinalized", false,
        "unassessedCount", 0,
        "sentences", [],
        "vocabulary", [],
        "allSentences", [],
        "allVocabulary", []
    )
    CPStudyCandidateState := scState
    scScopeDdl.OnEvent("Change", StudyCandidatesScopeChanged.Bind(scState))
    scAiFilterDdl.OnEvent(
        "Change", StudyCandidatesAiFilterChanged.Bind(scState)
    )
    scRefresh.OnEvent("Click", StudyCandidatesRefresh.Bind(scState))
    scRecommend.OnEvent(
        "Click", StudyCandidatesGenerateRecommendations.Bind(scState)
    )
    scTabs.OnEvent("Change", StudyCandidatesTabChanged.Bind(scState))
    scSentenceList.OnEvent("ItemSelect", StudyCandidatesUpdateActions.Bind(scState))
    scVocabularyList.OnEvent("ItemSelect", StudyCandidatesUpdateActions.Bind(scState))
    scSentenceList.OnEvent("DoubleClick", StudyCandidatesOpenSelected.Bind(scState))
    scVocabularyList.OnEvent("DoubleClick", StudyCandidatesOpenSelected.Bind(scState))
    scOpen.OnEvent("Click", StudyCandidatesOpenSelected.Bind(scState))
    scAdd.OnEvent("Click", StudyCandidatesAddSelected.Bind(scState))
    scFinish.OnEvent("Click", StudyCandidatesFinishReview.Bind(scState))
    scTriage.OnEvent("Click", StudyCandidatesTriageSelected.Bind(scState))
    scClose.OnEvent("Click", StudyCandidatesClose.Bind(scState))
    scGui.OnEvent("Escape", StudyCandidatesClose.Bind(scState))
    scGui.OnEvent("Close", StudyCandidatesClose.Bind(scState))
    scGui.OnEvent("Size", StudyCandidatesResize.Bind(scState))
    scGui.OnEvent("ContextMenu", StudyCandidatesContextMenu.Bind(scState))
    scGui.Show("Hide w840 h575")
    CPApplyOwnedDialogTheme(scGui)
    scGui.Show("Center")
    ; Several Windows 10/11 native controls finalize their non-client pieces
    ; (button faces, combo arrow and scrollbars) only after becoming visible.
    ; A final immediate pass makes those pieces dark without changing layout.
    CPApplyOwnedDialogTheme(scGui)
    ; Defer the synchronous bridge snapshot until this creation callback has
    ; unwound.  The visible status text provides immediate feedback meanwhile.
    SetTimer(StudyCandidatesInitialRefresh.Bind(scState), -1)
    } finally {
        CPStudyCandidateOpening := false
    }
}

StudyLibraryFormatBytes(slBytes) {
    try slBytes := Max(0, Integer(slBytes))
    catch
        slBytes := 0
    if (slBytes < 1024)
        return slBytes " B"
    slUnits := ["KB", "MB", "GB", "TB"]
    slValue := slBytes / 1024
    for slUnit in slUnits {
        if (slValue < 1024 || slUnit = "TB") {
            slDecimals := slValue < 10 ? 1 : 0
            return Round(slValue, slDecimals) " " slUnit
        }
        slValue /= 1024
    }
    return slBytes " B"
}

StudyLibraryReadStorage(slState) {
    slRows := StudyLibraryReadRows(slState["outputDir"] "\storage.tsv")
    if !slRows.Length || slRows[1].Length < 13
        return false
    slRow := slRows[1]
    return Map(
        "databaseBytes", Integer(slRow[1]),
        "mediaBytes", Integer(slRow[2]),
        "mediaFiles", Integer(slRow[3]),
        "backupBytes", Integer(slRow[4]),
        "backupFiles", Integer(slRow[5]),
        "trashBytes", Integer(slRow[6]),
        "trashFiles", Integer(slRow[7]),
        "otherBytes", Integer(slRow[8]),
        "totalBytes", Integer(slRow[9]),
        "freeBytes", Integer(slRow[10]),
        "volumeBytes", Integer(slRow[11]),
        "sourceCount", Integer(slRow[12]),
        "explanationCount", Integer(slRow[13])
    )
}

StudyLibraryRefreshStorage(slState, slForce := false, *) {
    if !StudyLibraryStateAlive(slState)
        return false
    if (slState.Has("storageRefreshing") && slState["storageRefreshing"])
        return false
    slLastTick := slState.Has("storageLastTick") ? slState["storageLastTick"] : 0
    if (!slForce && slLastTick && A_TickCount - slLastTick < 30000) {
        if slState.Has("storage") && IsObject(slState["storage"])
            slState["storageButton"].Text := "Storage: "
                . StudyLibraryFormatBytes(slState["storage"]["totalBytes"])
        return true
    }

    slState["storageRefreshing"] := true
    try {
        if !StudyLibraryRunBridge(slState, "storage")
            return false
        if !StudyLibraryStateAlive(slState)
            return false
        slStorage := StudyLibraryReadStorage(slState)
        if !IsObject(slStorage)
            return false
        slState["storage"] := slStorage
        slState["storageLastTick"] := A_TickCount
        slState["storageButton"].Text := "Storage: "
            . StudyLibraryFormatBytes(slStorage["totalBytes"])
        slState["storageButton"].Enabled := true
        return true
    } finally {
        if IsObject(slState)
            slState["storageRefreshing"] := false
    }
}

StudyLibraryUpdateStorageDialog(slState, slControls) {
    if !(slState.Has("storage") && IsObject(slState["storage"]))
        return
    slStorage := slState["storage"]
    slControls["entries"].Value := slStorage["sourceCount"] " source"
        . (slStorage["sourceCount"] = 1 ? "" : "s") " • "
        . slStorage["explanationCount"] " explanation"
        . (slStorage["explanationCount"] = 1 ? "" : "s")
    slControls["database"].Value := StudyLibraryFormatBytes(
        slStorage["databaseBytes"]
    )
    slControls["media"].Value := StudyLibraryFormatBytes(
        slStorage["mediaBytes"]
    ) "  (" slStorage["mediaFiles"] " file"
        . (slStorage["mediaFiles"] = 1 ? "" : "s") ")"
    slControls["backups"].Value := StudyLibraryFormatBytes(
        slStorage["backupBytes"]
    ) "  (" slStorage["backupFiles"] " file"
        . (slStorage["backupFiles"] = 1 ? "" : "s") ")"
    slControls["trash"].Value := StudyLibraryFormatBytes(
        slStorage["trashBytes"]
    ) "  (" slStorage["trashFiles"] " file"
        . (slStorage["trashFiles"] = 1 ? "" : "s") ")"
    slControls["other"].Value := StudyLibraryFormatBytes(
        slStorage["otherBytes"]
    )
    slControls["total"].Value := StudyLibraryFormatBytes(
        slStorage["totalBytes"]
    )
    slControls["free"].Value := StudyLibraryFormatBytes(
        slStorage["freeBytes"]
    ) " of " StudyLibraryFormatBytes(slStorage["volumeBytes"])

    if (slStorage["freeBytes"] < 2 * 1024 * 1024 * 1024) {
        slNotice := "Disk space is running low. Consider moving or cleaning "
            . "Study Library media, backups, or Trash."
    } else if (slStorage["sourceCount"] >= 5000) {
        slNotice := "This is a large library. It remains safe to use, but opening, "
            . "refreshing, and broad searches may become slower."
    } else {
        slNotice := "There is no fixed Study Library limit. Screenshots normally "
            . "account for most of its disk usage."
    }
    slControls["notice"].Value := slNotice
}

StudyLibraryRefreshStorageDialog(slState, slControls, *) {
    if StudyLibraryRefreshStorage(slState, true)
        StudyLibraryUpdateStorageDialog(slState, slControls)
}

StudyLibraryOpenStorageFolder(slState, *) {
    SplitPath(slState["database"],, &slDirectory)
    if (slDirectory != "")
        Run('explorer.exe "' slDirectory '"')
}

StudyLibraryOpenStorage(slState, *) {
    if !StudyLibraryRefreshStorage(slState, true)
        return
    SplitPath(slState["database"],, &slDirectory)
    slDialog := Gui(
        "+Owner" slState["gui"].Hwnd " +OwnDialogs",
        "Study Library - Storage - " slState["libraryName"]
    )
    slDialog.MarginX := 18, slDialog.MarginY := 16
    slDialog.SetFont("s10", "Segoe UI")
    slDialog.Add(
        "Text", "xm ym w540",
        "Current Study Library: " slState["libraryName"]
    )
        .SetFont("s11 Bold")
    slPath := slDialog.Add(
        "Edit", "xm y+8 w540 h30 ReadOnly", slDirectory
    )

    slControls := Map()
    slRows := [
        ["entries", "Saved entries"],
        ["database", "Database"],
        ["media", "Active screenshots"],
        ["backups", "Recovery backups"],
        ["trash", "Trash"],
        ["other", "Other files"],
        ["total", "Total Study Library"],
        ["free", "Free disk space"]
    ]
    for slIndex, slRow in slRows {
        slYOption := slIndex = 1 ? "xm y+16" : "xm y+10"
        slLabel := slDialog.Add("Text", slYOption " w180", slRow[2] ":")
        if (slRow[1] = "total")
            slLabel.SetFont("Bold")
        slValue := slDialog.Add("Text", "x+10 yp w350 Right", "")
        if (slRow[1] = "total")
            slValue.SetFont("Bold")
        slControls[slRow[1]] := slValue
    }
    slNotice := slDialog.Add("Text", "xm y+18 w540 h46 cGray", "")
    CPRegisterMutedControl(slNotice)
    slControls["notice"] := slNotice

    slOpenFolder := slDialog.Add("Button", "xm y+14 w160", "Open Library Folder")
    slRefresh := slDialog.Add("Button", "x+10 yp w100", "Refresh")
    slClose := slDialog.Add("Button", "x+10 yp w100 Default", "Close")
    slOpenFolder.OnEvent("Click", StudyLibraryOpenStorageFolder.Bind(slState))
    slRefresh.OnEvent(
        "Click", StudyLibraryRefreshStorageDialog.Bind(slState, slControls)
    )
    slClose.OnEvent("Click", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Close", StudyLibraryCloseDialog.Bind(slDialog))
    StudyLibraryUpdateStorageDialog(slState, slControls)
    slDialog.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slDialog)
}

StudyLibraryRefreshAll(slState, *) {
    StudyLibraryRefresh(slState)
    StudyLibraryRefreshStorage(slState, true)
}

StudyLibraryRunBridge(slState, slAction, slGroupId := 0, slVersion := 0) {
    global pythonExe
    slPython := ResolvePath(pythonExe)
    slBridge := A_ScriptDir "\scripts\study_library.py"
    if !(FileExist(slPython) && FileExist(slBridge)) {
        MsgBox(
            "The Study Library viewer needs valid Python and bridge paths.`n`n"
            . "Python:`n" slPython "`n`nBridge:`n" slBridge,
            "Study Library",
            "OK Icon!"
        )
        return false
    }

    slCommand := Format(
        '"{1}" "{2}" {3} --db "{4}" --output-dir "{5}"',
        slPython, slBridge, slAction, slState["database"], slState["outputDir"]
    )
    if (slAction = "detail" || slAction = "set-metadata"
        || slAction = "set-anki" || slAction = "save-edit"
        || slAction = "revert-edit" || slAction = "remove-version") {
        slCommand .= " --group-id " Integer(slGroupId)
    }
    if (slAction = "detail" || slAction = "save-edit"
        || slAction = "revert-edit" || slAction = "remove-version")
        slCommand .= " --version " Integer(slVersion)
    DbgCP("Study Library bridge -> " slCommand)
    slBridgeOperation := (slAction = "set-metadata" || slAction = "bulk-metadata"
        || slAction = "set-anki" || slAction = "save-edit"
        || slAction = "revert-edit" || slAction = "remove-version")
        ? "updated" : "read"
    try slExitCode := RunWait(slCommand, A_ScriptDir, "Hide")
    catch as slBridgeError {
        MsgBox(
            "The Study Library could not be " slBridgeOperation ".`n`n"
            . slBridgeError.Message,
            "Study Library",
            "OK Iconx"
        )
        return false
    }
    if (slExitCode != 0) {
        MsgBox(
            "The Study Library could not be " slBridgeOperation
            . " (bridge exit " slExitCode ").",
            "Study Library",
            "OK Iconx"
        )
        return false
    }
    return true
}

StudyLibraryClearDetail(slState, slMessage := "Select an explanation.") {
    slState["currentGroupId"] := 0
    slState["versions"] := []
    slState["sections"] := []
    slState["media"] := []
    slState["mediaIndex"] := 0
    slState["currentVersion"] := 0
    slState["currentChapter"] := ""
    slState["currentSpeaker"] := ""
    slState["currentTags"] := ""
    slState["currentAddedToAnkiAt"] := ""
    slState["currentAnkiStatus"] := "not_checked"
    slState["currentAnkiNoteId"] := 0
    slState["currentAnkiCheckedAt"] := ""
    slState["currentProfile"] := ""
    slState["currentProvider"] := ""
    slState["currentModel"] := ""
    slState["currentPrompt"] := ""
    slState["currentManuallyEditedAt"] := ""
    slState["detailTitle"].Value := slMessage
    slState["metadata"].Value := ""
    slState["source"].Value := ""
    slState["explanation"].Value := ""
    slState["versionDdl"].Delete()
    slState["versionDdl"].Enabled := false
    if slState.Has("editDetailsButton")
        slState["editDetailsButton"].Enabled := false
    if slState.Has("removeVersionButton")
        slState["removeVersionButton"].Enabled := false
    StudyLibrarySyncVersionNavigation(slState)
    if slState.Has("studyButton")
        slState["studyButton"].Enabled := false
    if slState.Has("editExplanationButton")
        slState["editExplanationButton"].Enabled := false
    if slState.Has("copyButton")
        slState["copyButton"].Enabled := false
    if slState.Has("addAnkiButton")
        slState["addAnkiButton"].Enabled := false
    slState["sectionDdl"].Delete()
    slState["sectionDdl"].Add(["Full explanation"])
    slState["sectionDdl"].Choose(1)
    slState["sectionDdl"].Enabled := false
    if slState.Has("sectionNavCallback")
        slState["sectionNavCallback"].Call()
    slState["picture"].Visible := false
    try slState["picture"].Value := ""
    slState["imageInfo"].Value := "No source screenshot selected."
    slState["previousImage"].Enabled := false
    slState["nextImage"].Enabled := false
    slState["openImage"].Enabled := false
}

StudyLibrarySyncVersionNavigation(slState) {
    slCount := slState["versions"].Length
    slIndex := slState["versionDdl"].Value
    slLabel := ""
    if (slIndex >= 1 && slIndex <= slCount) {
        slVersionEntry := slState["versions"][slIndex]
        slLabel := Format(
            "v{1:02}  •  {2}",
            slVersionEntry["version"],
            slVersionEntry["created"]
        )
        if slVersionEntry["preferred"]
            slLabel .= "  (latest)"
        if slVersionEntry.Has("manuallyEdited") && slVersionEntry["manuallyEdited"]
            slLabel .= "  •  manually edited"
    }
    if slState.Has("versionView")
        slState["versionView"].Value := slLabel
    if slState.Has("previousVersion")
        slState["previousVersion"].Enabled := slIndex > 1
    if slState.Has("nextVersion")
        slState["nextVersion"].Enabled := slIndex >= 1 && slIndex < slCount
    if slState.Has("newVersionButton")
        slState["newVersionButton"].Enabled := slCount > 0
    if slState.Has("removeVersionButton") {
        slState["removeVersionButton"].Text := slCount = 1
            ? "Remove explanation..."
            : "Remove version..."
        slState["removeVersionButton"].Enabled := slCount > 0
    }
}

StudyLibraryStepVersion(slState, slDirection, *) {
    if (slState.Has("editing") && slState["editing"])
        return
    slIndex := slState["versionDdl"].Value
    slTargetIndex := slIndex + slDirection
    if (slIndex < 1 || slTargetIndex < 1
        || slTargetIndex > slState["versions"].Length)
        return
    slSectionKey := "full"
    if slState.Has("entryIds") {
        slSectionIndex := slState["sectionDdl"].Value
        if (slSectionIndex >= 1
            && slSectionIndex <= slState["sections"].Length)
            slSectionKey := slState["sections"][slSectionIndex]["key"]
    }
    StudyLibraryLoadGroup(
        slState,
        slState["currentGroupId"],
        slState["versions"][slTargetIndex]["version"],
        slSectionKey
    )
}

StudyLibraryImageDimensions(slPath, &slWidth, &slHeight) {
    slWidth := 0, slHeight := 0
    try {
        slImageType := 0
        slHandle := LoadPicture(slPath, "", &slImageType)
        if !slHandle
            return false
        if (slImageType = 0) {
            slBitmapInfo := Buffer(32, 0)
            if DllCall("gdi32\GetObjectW", "ptr", slHandle, "int", slBitmapInfo.Size,
                "ptr", slBitmapInfo.Ptr, "int") {
                slWidth := Abs(NumGet(slBitmapInfo, 4, "Int"))
                slHeight := Abs(NumGet(slBitmapInfo, 8, "Int"))
            }
            DllCall("gdi32\DeleteObject", "ptr", slHandle)
        } else {
            DllCall("user32\DestroyIcon", "ptr", slHandle)
        }
    }
    return (slWidth > 0 && slHeight > 0)
}

StudyLibraryShowImage(slState, *) {
    if !(slState.Has("gui") && slState["gui"] && slState["gui"].Hwnd)
        return
    slMedia := slState["media"]
    slIndex := slState["mediaIndex"]
    if (slIndex < 1 || slIndex > slMedia.Length) {
        slState["picture"].Visible := false
        slState["imageInfo"].Value := "No source screenshot saved for this version."
        slState["previousImage"].Enabled := false
        slState["nextImage"].Enabled := false
        slState["openImage"].Enabled := false
        return
    }

    slItem := slMedia[slIndex]
    slPath := slItem["path"]
    if !FileExist(slPath) {
        slState["picture"].Visible := false
        slState["imageInfo"].Value := "The saved source screenshot is missing."
        slState["previousImage"].Enabled := (slIndex > 1)
        slState["nextImage"].Enabled := (slIndex < slMedia.Length)
        slState["openImage"].Enabled := false
        return
    }

    slArea := slState["imageArea"]
    slNativeW := slItem["width"], slNativeH := slItem["height"]
    if (slNativeW <= 0 || slNativeH <= 0) {
        StudyLibraryImageDimensions(slPath, &slNativeW, &slNativeH)
        slItem["width"] := slNativeW, slItem["height"] := slNativeH
    }
    if (slNativeW <= 0 || slNativeH <= 0) {
        slNativeW := slArea["w"], slNativeH := slArea["h"]
    }
    slScale := Min(slArea["w"] / slNativeW, slArea["h"] / slNativeH)
    slDisplayW := Max(1, Round(slNativeW * slScale))
    slDisplayH := Max(1, Round(slNativeH * slScale))
    slX := slArea["x"] + Floor((slArea["w"] - slDisplayW) / 2)
    slY := slArea["y"] + Floor((slArea["h"] - slDisplayH) / 2)
    try {
        ; Picture load dimensions are physical pixels, while GUI Move uses
        ; DPI-scaled logical coordinates. Supplying logical values directly
        ; shrinks previews severely on high-DPI displays.
        slDpi := GetWindowDPI(slState["gui"].Hwnd)
        slLoadW := Max(1, Round(slDisplayW * slDpi / 96))
        slLoadH := Max(1, Round(slDisplayH * slDpi / 96))
        slState["picture"].Value := "*w" slLoadW " *h" slLoadH " " slPath
        ; Assigning a new bitmap can resize the Picture control. Restore the
        ; intended logical bounds after loading it.
        slState["picture"].Move(slX, slY, slDisplayW, slDisplayH)
        slState["picture"].Visible := true
    } catch {
        slState["picture"].Visible := false
    }
    slState["imageInfo"].Value := "Screenshot " slIndex " of " slMedia.Length
    slState["previousImage"].Enabled := (slIndex > 1)
    slState["nextImage"].Enabled := (slIndex < slMedia.Length)
    slState["openImage"].Enabled := true
}

StudyLibraryPreviousImage(slState, *) {
    if (slState["mediaIndex"] > 1) {
        slState["mediaIndex"] -= 1
        StudyLibraryShowImage(slState)
    }
}

StudyLibraryNextImage(slState, *) {
    if (slState["mediaIndex"] < slState["media"].Length) {
        slState["mediaIndex"] += 1
        StudyLibraryShowImage(slState)
    }
}

StudyLibraryOpenImage(slState, *) {
    slIndex := slState["mediaIndex"]
    if (slIndex < 1 || slIndex > slState["media"].Length)
        return
    slPath := slState["media"][slIndex]["path"]
    if FileExist(slPath)
        try Run('"' slPath '"')
}

StudyLibraryUpdateMetadataText(slState) {
    if (slState["currentGroupId"] <= 0) {
        slState["metadata"].Value := ""
        return
    }
    slProfile := slState.Has("currentProfile") ? slState["currentProfile"] : ""
    slProvider := slState.Has("currentProvider") ? slState["currentProvider"] : ""
    slModel := slState.Has("currentModel") ? slState["currentModel"] : ""
    slPrompt := slState.Has("currentPrompt") ? slState["currentPrompt"] : ""
    slState["metadata"].Value := "Profile: "
        . (slProfile != "" ? slProfile : "Unsorted")
        . "    Chapter: "
        . (slState["currentChapter"] != "" ? slState["currentChapter"] : "—")
        . "    Speaker: "
        . (slState["currentSpeaker"] != "" ? slState["currentSpeaker"] : "—")
        . "`r`nTags: "
        . (slState["currentTags"] != "" ? slState["currentTags"] : "—")
        . "    Anki: "
        . StudyAnkiStatusLabel(
            slState["currentAnkiStatus"], slState["currentAddedToAnkiAt"]
        )
        . "    Model: " slProvider " / " slModel "    Prompt: " slPrompt
}

StudyLibraryLoadGroup(
    slState, slGroupId, slVersion := 0, slPreferredSectionKey := ""
) {
    if (slGroupId <= 0)
        return StudyLibraryClearDetail(slState)
    if !StudyLibraryRunBridge(slState, "detail", slGroupId, slVersion)
        return

    slOutputDir := slState["outputDir"]
    slDetailRows := StudyLibraryReadRows(slOutputDir "\detail.tsv")
    if !slDetailRows.Length
        return StudyLibraryClearDetail(slState, "This explanation is no longer available.")
    slDetail := slDetailRows[1]
    if (slDetail.Length < 10)
        return StudyLibraryClearDetail(slState, "This explanation could not be read.")

    slState["currentGroupId"] := Integer(slDetail[1])
    slSelectedVersion := Integer(slDetail[3])
    slProfile := StudyLibraryHexDecode(slDetail[5])
    slCreated := StudyLibraryHexDecode(slDetail[6])
    slProvider := StudyLibraryHexDecode(slDetail[7])
    slModel := StudyLibraryHexDecode(slDetail[8])
    slPrompt := StudyLibraryHexDecode(slDetail[9])
    slChapter := slDetail.Length >= 11 ? StudyLibraryHexDecode(slDetail[11]) : ""
    slSpeaker := slDetail.Length >= 12 ? StudyLibraryHexDecode(slDetail[12]) : ""
    slTags := slDetail.Length >= 13 ? StudyLibraryHexDecode(slDetail[13]) : ""
    slAddedToAnkiAt := slDetail.Length >= 14 ? StudyLibraryHexDecode(slDetail[14]) : ""
    slManuallyEditedAt := slDetail.Length >= 15
        ? StudyLibraryHexDecode(slDetail[15]) : ""
    slAnkiStatus := slDetail.Length >= 16
        ? StudyLibraryHexDecode(slDetail[16]) : "not_checked"
    slAnkiNoteId := slDetail.Length >= 17 ? Integer(slDetail[17]) : 0
    slAnkiCheckedAt := slDetail.Length >= 18
        ? StudyLibraryHexDecode(slDetail[18]) : ""
    slState["currentVersion"] := slSelectedVersion
    slState["currentChapter"] := slChapter
    slState["currentSpeaker"] := slSpeaker
    slState["currentTags"] := slTags
    slState["currentAddedToAnkiAt"] := slAddedToAnkiAt
    slState["currentAnkiStatus"] := slAnkiStatus
    slState["currentAnkiNoteId"] := slAnkiNoteId
    slState["currentAnkiCheckedAt"] := slAnkiCheckedAt
    slState["currentProfile"] := slProfile
    slState["currentProvider"] := slProvider
    slState["currentModel"] := slModel
    slState["currentPrompt"] := slPrompt
    slState["currentManuallyEditedAt"] := slManuallyEditedAt

    slState["suspend"] := true
    slState["versions"] := []
    slVersionLabels := []
    slVersionIndex := 1
    for slVersionRow in StudyLibraryReadRows(slOutputDir "\versions.tsv") {
        if (slVersionRow.Length < 4)
            continue
        slVersionNumber := Integer(slVersionRow[2])
        slVersionDate := StudyLibraryHexDecode(slVersionRow[4])
        slPreferred := Integer(slVersionRow[3])
        slManuallyEditedAt := slVersionRow.Length >= 5
            ? StudyLibraryHexDecode(slVersionRow[5]) : ""
        slVersionEntry := Map(
            "id", Integer(slVersionRow[1]),
            "version", slVersionNumber,
            "preferred", slPreferred,
            "created", slVersionDate,
            "manuallyEdited", slManuallyEditedAt != ""
        )
        slState["versions"].Push(slVersionEntry)
        slVersionLabel := Format("v{1:02}  •  {2}", slVersionNumber, slVersionDate)
        if slPreferred
            slVersionLabel .= "  (latest)"
        if (slManuallyEditedAt != "")
            slVersionLabel .= "  •  manually edited"
        slVersionLabels.Push(slVersionLabel)
        if (slVersionNumber = slSelectedVersion)
            slVersionIndex := slVersionLabels.Length
    }
    slState["versionDdl"].Delete()
    if slVersionLabels.Length {
        slState["versionDdl"].Add(slVersionLabels)
        slState["versionDdl"].Choose(slVersionIndex)
        slState["versionDdl"].Enabled := true
    } else {
        slState["versionDdl"].Enabled := false
    }
    if slState.Has("editDetailsButton")
        slState["editDetailsButton"].Enabled := true
    if slState.Has("removeVersionButton")
        slState["removeVersionButton"].Enabled := slVersionLabels.Length > 0
    if slState.Has("studyButton")
        slState["studyButton"].Enabled := true
    if slState.Has("editExplanationButton")
        slState["editExplanationButton"].Enabled := true
    if slState.Has("copyButton")
        slState["copyButton"].Enabled := true
    slState["suspend"] := false
    StudyLibrarySyncVersionNavigation(slState)

    slSourcePath := slOutputDir "\source.txt"
    slExplanationPath := slOutputDir "\explanation.txt"
    slSourceText := FileExist(slSourcePath) ? FileRead(slSourcePath, "UTF-8") : ""
    slExplanationText := FileExist(slExplanationPath)
        ? FileRead(slExplanationPath, "UTF-8") : ""
    slState["source"].Value := slSourceText
    slState["suspend"] := true
    slState["sections"] := [
        Map("key", "full", "heading", "Full explanation", "content", slExplanationText)
    ]
    slSectionLabels := ["Full explanation"]
    for slSectionRow in StudyLibraryReadRows(slOutputDir "\sections.tsv") {
        if (slSectionRow.Length < 4)
            continue
        slSectionKey := StudyLibraryHexDecode(slSectionRow[2])
        ; A parser fallback is identical to the permanent Full explanation item.
        if (slSectionKey = "full")
            continue
        slSectionHeading := StudyLibraryHexDecode(slSectionRow[3])
        slSectionFile := StudyLibraryHexDecode(slSectionRow[4])
        slSectionPath := slOutputDir "\" slSectionFile
        slSectionText := FileExist(slSectionPath)
            ? FileRead(slSectionPath, "UTF-8") : ""
        slState["sections"].Push(Map(
            "key", slSectionKey,
            "heading", slSectionHeading,
            "content", slSectionText
        ))
        slSectionLabels.Push(slSectionHeading)
    }
    slState["sectionDdl"].Delete()
    slState["sectionDdl"].Add(slSectionLabels)
    slTargetSectionIndex := 1
    if (slPreferredSectionKey != "" && slPreferredSectionKey != "full") {
        for slIndex, slSection in slState["sections"] {
            if (slSection["key"] = slPreferredSectionKey) {
                slTargetSectionIndex := slIndex
                break
            }
        }
    }
    slState["sectionDdl"].Choose(slTargetSectionIndex)
    slState["sectionDdl"].Enabled := (slSectionLabels.Length > 1)
    ; Paint only the requested section. Previously the loader displayed the
    ; full explanation first and the Reader restored the section afterward,
    ; causing a brief but visible flash during entry/version navigation.
    slState["explanation"].Value :=
        slState["sections"][slTargetSectionIndex]["content"]
    slState["suspend"] := false
    if slState.Has("sectionNavCallback")
        slState["sectionNavCallback"].Call()
    slState["detailTitle"].Value := "Saved explanation"
    StudyLibraryUpdateMetadataText(slState)

    slState["media"] := []
    for slMediaRow in StudyLibraryReadRows(slOutputDir "\media.tsv") {
        if (slMediaRow.Length < 7)
            continue
        slMediaPath := StudyLibraryHexDecode(slMediaRow[2])
        slState["media"].Push(Map(
            "order", Integer(slMediaRow[1]),
            "path", slMediaPath,
            "exists", Integer(slMediaRow[7]),
            "width", 0,
            "height", 0
        ))
    }
    slState["mediaIndex"] := slState["media"].Length ? 1 : 0
    StudyLibraryShowImage(slState)
    if slState.Has("entryNavCallback")
        slState["entryNavCallback"].Call()
}

StudyLibraryGroupFocused(slState, slList, slRow, *) {
    if slState["suspend"] || slRow < 1 || slRow > slState["groups"].Length
        return
    slGroupId := StudyLibraryRowGroupId(slState, slRow)
    if (slGroupId <= 0)
        return
    StudyLibraryLoadGroup(slState, slGroupId)
    StudyLibraryUpdateSelectionActions(slState)
}

StudyLibraryVersionChanged(slState, *) {
    if slState["suspend"]
        return
    slIndex := slState["versionDdl"].Value
    if (slIndex < 1 || slIndex > slState["versions"].Length)
        return
    slSectionKey := "full"
    if slState.Has("entryIds") {
        slSectionIndex := slState["sectionDdl"].Value
        if (slSectionIndex >= 1
            && slSectionIndex <= slState["sections"].Length)
            slSectionKey := slState["sections"][slSectionIndex]["key"]
    }
    StudyLibraryLoadGroup(
        slState,
        slState["currentGroupId"],
        slState["versions"][slIndex]["version"],
        slSectionKey
    )
}

StudyLibrarySectionChanged(slState, *) {
    if slState["suspend"]
        return
    slIndex := slState["sectionDdl"].Value
    if (slIndex < 1 || slIndex > slState["sections"].Length)
        return
    slState["explanation"].Value := slState["sections"][slIndex]["content"]
    try {
        DllCall("user32\SendMessageW", "ptr", slState["explanation"].Hwnd,
            "uint", 0x00B1, "ptr", 0, "ptr", 0) ; EM_SETSEL
        DllCall("user32\SendMessageW", "ptr", slState["explanation"].Hwnd,
            "uint", 0x00B7, "ptr", 0, "ptr", 0) ; EM_SCROLLCARET
    }
    if slState.Has("sectionNavCallback")
        slState["sectionNavCallback"].Call()
}

StudyReaderSyncSectionNavigation(srState, *) {
    srSectionCount := Max(0, srState["sections"].Length - 1)
    srIndex := srState["sectionDdl"].Value
    if (srIndex < 1 || srIndex > srState["sections"].Length)
        srIndex := 1
    srHasExplanation := srState["currentGroupId"] > 0
    srState["fullSectionButton"].Enabled := srHasExplanation
    srState["previousSection"].Enabled := srSectionCount > 0
    srState["nextSection"].Enabled := srSectionCount > 0
    if !srHasExplanation {
        srState["sectionStatus"].Value := "No explanation selected"
    } else if (srIndex = 1) {
        srState["sectionStatus"].Value := "Full explanation"
    } else {
        srHeading := srState["sections"][srIndex]["heading"]
        srState["sectionStatus"].Value := srHeading "  ·  "
            . (srIndex - 1) " of " srSectionCount
    }
}

StudyReaderSelectFullSection(srState, *) {
    if (srState.Has("editing") && srState["editing"])
        return
    if (srState["currentGroupId"] <= 0 || !srState["sections"].Length)
        return
    srState["sectionDdl"].Choose(1)
    StudyLibrarySectionChanged(srState)
}

StudyReaderStepSection(srState, srDirection, *) {
    if (srState.Has("editing") && srState["editing"])
        return
    srSectionCount := Max(0, srState["sections"].Length - 1)
    if (srSectionCount <= 0)
        return
    srIndex := srState["sectionDdl"].Value
    if (srIndex <= 1) {
        srIndex := srDirection > 0 ? 2 : srState["sections"].Length
    } else if (srDirection > 0) {
        srIndex := srIndex < srState["sections"].Length ? srIndex + 1 : 2
    } else {
        srIndex := srIndex > 2 ? srIndex - 1 : srState["sections"].Length
    }
    srState["sectionDdl"].Choose(srIndex)
    StudyLibrarySectionChanged(srState)
}

StudyReaderSetEntrySequence(srState, srGroups, srGroupId) {
    srIds := []
    srSeen := Map()
    if IsObject(srGroups) {
        for srGroup in srGroups {
            if !(IsObject(srGroup) && srGroup.Has("id"))
                continue
            srId := Integer(srGroup["id"])
            if (srId <= 0 || srSeen.Has(srId))
                continue
            srSeen[srId] := true
            srIds.Push(srId)
        }
    }
    if !srIds.Length && srGroupId > 0
        srIds.Push(Integer(srGroupId))
    srState["entryIds"] := srIds
    StudyReaderSyncEntryNavigation(srState)
}

StudyReaderEntryIndex(srState, srGroupId := 0) {
    if (srGroupId <= 0)
        srGroupId := srState["currentGroupId"]
    for srIndex, srId in srState["entryIds"] {
        if (srId = srGroupId)
            return srIndex
    }
    return 0
}

StudyReaderSyncEntryNavigation(srState, *) {
    global iniPath
    srHasEntry := srState["currentGroupId"] > 0
    srIndex := StudyReaderEntryIndex(srState)
    if (srHasEntry && !srIndex) {
        srState["entryIds"].Push(srState["currentGroupId"])
        srIndex := srState["entryIds"].Length
    }
    srCount := srState["entryIds"].Length
    ; Entry browsing wraps at either end, so both controls remain useful as
    ; long as there is more than one saved explanation.
    srCanStep := srHasEntry && srIndex > 0 && srCount > 1
    srState["previousEntry"].Enabled := srCanStep
    srState["nextEntry"].Enabled := srCanStep
    srState["entryStatus"].Value := srHasEntry && srIndex
        ? srIndex " of " srCount : "No entry"
    srState["suspendAnki"] := true
    srFoundInAnki := srHasEntry && srState["currentAnkiStatus"] = "found"
    srState["ankiCheck"].Enabled := srHasEntry && !srFoundInAnki
    srState["ankiCheck"].Value := srHasEntry
        && (srFoundInAnki || srState["currentAddedToAnkiAt"] != "") ? 1 : 0
    srState["ankiCheck"].Text := srFoundInAnki
        ? "Found in Anki (exact match)"
        : "Added to Anki (manual)  (Alt+A)"
    if srState.Has("addAnkiButton")
        ; Vocabulary cards remain available even if the complete explanation
        ; already has an exact Anki match.
        srState["addAnkiButton"].Enabled := srHasEntry
            && !(srState.Has("editing") && srState["editing"])
    srState["suspendAnki"] := false
    if srHasEntry
        try IniWrite(srState["currentGroupId"], iniPath,
            "study_reader_view", "lastGroupId")
}

StudyReaderRememberEntryView(srState) {
    global iniPath
    srGroupId := srState["currentGroupId"]
    if (srGroupId <= 0)
        return
    try IniWrite(srGroupId, iniPath, "study_reader_view", "lastGroupId")
    srSectionKey := "full"
    srSectionIndex := srState["sectionDdl"].Value
    if (srSectionIndex >= 1 && srSectionIndex <= srState["sections"].Length)
        srSectionKey := srState["sections"][srSectionIndex]["key"]
    srFirstLine := 0
    try srFirstLine := DllCall(
        "user32\SendMessageW", "ptr", srState["explanation"].Hwnd,
        "uint", 0x00CE, "ptr", 0, "ptr", 0, "int" ; EM_GETFIRSTVISIBLELINE
    )
    srState["entryViewMemory"][srGroupId] := Map(
        "version", srState["currentVersion"],
        "sectionKey", srSectionKey,
        "firstLine", Max(0, srFirstLine)
    )
}

StudyReaderRestoreEntryView(srState, srGroupId) {
    if !srState["entryViewMemory"].Has(srGroupId)
        return
    srMemory := srState["entryViewMemory"][srGroupId]
    srTargetSection := 1
    for srIndex, srSection in srState["sections"] {
        if (srSection["key"] = srMemory["sectionKey"]) {
            srTargetSection := srIndex
            break
        }
    }
    srState["sectionDdl"].Choose(srTargetSection)
    StudyLibrarySectionChanged(srState)
    srFirstLine := srMemory["firstLine"]
    if (srFirstLine > 0) {
        try DllCall(
            "user32\SendMessageW", "ptr", srState["explanation"].Hwnd,
            "uint", 0x00B6, "ptr", 0, "ptr", srFirstLine ; EM_LINESCROLL
        )
    }
}

StudyReaderStepEntry(srState, srDirection, *) {
    if (srState.Has("editing") && srState["editing"])
        return
    srIndex := StudyReaderEntryIndex(srState)
    srCount := srState["entryIds"].Length
    if (srIndex <= 0 || srCount <= 1)
        return
    srSectionKey := "full"
    srSectionIndex := srState["sectionDdl"].Value
    if (srSectionIndex >= 1 && srSectionIndex <= srState["sections"].Length)
        srSectionKey := srState["sections"][srSectionIndex]["key"]
    srTargetIndex := srIndex + srDirection
    if (srTargetIndex < 1)
        srTargetIndex := srCount
    else if (srTargetIndex > srCount)
        srTargetIndex := 1
    StudyReaderRememberEntryView(srState)
    srTargetId := srState["entryIds"][srTargetIndex]
    srTargetVersion := 0
    if srState["entryViewMemory"].Has(srTargetId)
        srTargetVersion := srState["entryViewMemory"][srTargetId]["version"]
    ; Supply the section to the loader so the target section is painted
    ; directly, without briefly showing Full explanation first.
    StudyLibraryLoadGroup(srState, srTargetId, srTargetVersion, srSectionKey)
}

StudyReaderCurrentSectionKey(srState) {
    srSectionKey := "full"
    srSectionIndex := srState["sectionDdl"].Value
    if (srSectionIndex >= 1 && srSectionIndex <= srState["sections"].Length)
        srSectionKey := srState["sections"][srSectionIndex]["key"]
    return srSectionKey
}

StudyReaderRefreshMainExplainPromptList() {
    global ddlEPr
    if !IsSet(ddlEPr)
        return
    srPreviousMainPrompt := Trim(ddlEPr.Text)
    RefreshExplainPromptProfilesList(srPreviousMainPrompt)
    ; If an externally deleted prompt was active in the Explanation tab,
    ; persist the safe fallback selected by the shared refresh routine.
    if (Trim(ddlEPr.Text) != srPreviousMainPrompt)
        ExplainPromptChanged()
}

StudyReaderRefreshNewVersionPromptList(srNewState, srSelect := "") {
    srPromptList := ListExplainPromptProfiles()
    srPromptCombo := srNewState["prompt"]
    srPromptCombo.Delete()
    if !srPromptList.Length {
        try srPromptCombo.Text := ""
        return
    }
    srPromptCombo.Add(srPromptList)
    srPromptIndex := srSelect != ""
        ? ArrayIndexOf(srPromptList, srSelect) : 0
    if !srPromptIndex
        srPromptIndex := 1
    srPromptCombo.Choose(srPromptIndex)
}

StudyReaderSaveNewVersionPrompt(
    srNewState, srEditorGui, srEditor, srPath, srName, *
) {
    try {
        if FileExist(srPath)
            FileCopy(srPath, srPath ".bak", true)
        srFile := FileOpen(srPath, "w", "UTF-8")
        if !srFile
            throw Error("The prompt file could not be opened for writing.")
        srFile.Write(srEditor.Value)
        srFile.Close()
        StudyReaderRefreshNewVersionPromptList(srNewState, srName)
        StudyReaderRefreshMainExplainPromptList()
        Toast("Saved EXPLAIN prompt: " srName)
    } catch as srPromptError {
        CPThemedOwnedMessage(
            srEditorGui.Hwnd,
            "The explanation prompt could not be saved.`n`n"
                . srPromptError.Message,
            "Edit Explanation Prompt", "ok", "error"
        )
    }
}

StudyReaderEditNewVersionPrompt(srNewState, *) {
    srName := Trim(srNewState["prompt"].Text)
    if (srName = "")
        srName := "default"
    srPath := ExplainProfilePath(srName)
    srText := ""
    try srText := FileExist(srPath) ? FileRead(srPath, "UTF-8") : ""

    srEditorGui := Gui(
        "+Resize +Owner" srNewState["gui"].Hwnd,
        "Edit Explanation Prompt - " srName
    )
    srEditor := srEditorGui.Add(
        "Edit", "xm ym w680 h420 WantTab WantReturn Wrap", srText
    )
    srSave := srEditorGui.Add("Button", "xm y+8 w100", "Save")
    srClose := srEditorGui.Add("Button", "x+8 yp w100", "Close")
    srSave.OnEvent(
        "Click",
        StudyReaderSaveNewVersionPrompt.Bind(
            srNewState, srEditorGui, srEditor, srPath, srName
        )
    )
    srClose.OnEvent("Click", (*) => srEditorGui.Destroy())
    srEditorGui.OnEvent("Size", (gui, mm, w, h) => (
        srEditor.Move(, , Max(300, w - 40), Max(180, h - 90)),
        srButtonY := h - 52,
        srSave.Move(20, srButtonY, 100, 32),
        srClose.Move(128, srButtonY, 100, 32)
    ))
    CPShowTextEditorDialog(
        srEditorGui, srEditor, srNewState["gui"].Hwnd
    )
}

StudyReaderAddNewVersionPrompt(srNewState, *) {
    srInput := CPThemedInputBox(
        "Enter a name for the new EXPLANATION prompt:",
        "New EXPLAIN prompt", "", "", 420, srNewState["gui"].Hwnd
    )
    if (srInput.Result = "Cancel")
        return
    srName := Trim(srInput.Value)
    if (srName = "") {
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd, "Please enter a non-empty name.",
            "New EXPLAIN prompt", "ok", "warning"
        )
        return
    }
    srName := RegExReplace(srName, '[\\/:*?"<>|]+', "_")
    if RegExMatch(srName, 'i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$')
        srName := "_" srName
    srPath := ExplainProfilePath(srName)
    if FileExist(srPath) {
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd,
            "A prompt with that name already exists.",
            "New EXPLAIN prompt", "ok", "warning"
        )
        return
    }
    FileAppend(
        "You are a friendly tutor for learners of Japanese.`r`n`r`n"
            . "Japanese:`r`n{jp}",
        srPath, "UTF-8"
    )
    StudyReaderRefreshNewVersionPromptList(srNewState, srName)
    StudyReaderRefreshMainExplainPromptList()
}

StudyReaderDeleteNewVersionPrompt(srNewState, *) {
    srName := Trim(srNewState["prompt"].Text)
    if (srName = "") {
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd, "No EXPLAIN prompt selected.",
            "Delete EXPLAIN prompt", "ok", "warning"
        )
        return
    }
    if (CPThemedOwnedMessage(
        srNewState["gui"].Hwnd,
        "Delete EXPLAIN prompt '" srName "'?",
        "Delete EXPLAIN prompt", "yesno", "warning"
    ) != "Yes")
        return
    try FileDelete(ExplainProfilePath(srName))
    StudyReaderRefreshNewVersionPromptList(srNewState)
    StudyReaderRefreshMainExplainPromptList()
}

StudyReaderNewVersionProviderChanged(srNewState, *) {
    srProvider := StrLower(Trim(srNewState["provider"].Text))
    srNewState["geminiModel"].Enabled := srProvider = "gemini"
    srNewState["openAIModel"].Enabled := srProvider = "openai"
}

StudyReaderCloseNewVersionDialog(srNewState, *) {
    if srNewState.Has("closed") && srNewState["closed"]
        return
    srNewState["closed"] := true
    try srNewState["readerState"]["newVersionDialog"] := 0
    try srNewState["gui"].Destroy()
}

StudyReaderGenerateNewVersion(srNewState, *) {
    global pythonExe, explainScript, debugMode, CPStudyLibraryState
    srReaderState := srNewState["readerState"]
    if (srReaderState["currentGroupId"] <= 0)
        return
    srProvider := StrLower(Trim(srNewState["provider"].Text))
    if (srProvider != "gemini" && srProvider != "openai")
        return
    srModel := Trim(
        srProvider = "gemini"
            ? srNewState["geminiModel"].Text
            : srNewState["openAIModel"].Text
    )
    srPromptName := Trim(srNewState["prompt"].Text)
    if (srModel = "" || srPromptName = "") {
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd,
            "Select a model and an explanation prompt first.",
            "Generate new explanation version", "ok", "warning"
        )
        return
    }
    if !CPApiKeyConfigured(srProvider) {
        srProviderLabel := srProvider = "gemini" ? "Gemini" : "OpenAI"
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd,
            srProviderLabel " API key missing.`n`nAdd it in the API Keys tab, "
                . "or set the corresponding Windows environment variable and "
                . "restart JRPG Translator.",
            "Generate new explanation version", "ok", "error"
        )
        return
    }

    srPython := ResolvePath(pythonExe)
    srExplainer := ResolvePath(explainScript)
    if !(FileExist(srPython) && FileExist(srExplainer)) {
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd,
            "The configured Python runtime or explainer script could not be found.",
            "Generate new explanation version", "ok", "error"
        )
        return
    }
    srSourceText := Trim(srReaderState["source"].Value)
    srSourceFile := srReaderState["outputDir"] "\\new_version_source.txt"
    srPathsFile := srReaderState["outputDir"] "\\new_version_media.txt"
    try FileDelete(srSourceFile)
    try FileDelete(srPathsFile)
    srSourceHandle := FileOpen(srSourceFile, "w", "UTF-8-RAW")
    if !srSourceHandle {
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd,
            "The archived Japanese source could not be prepared.",
            "Generate new explanation version", "ok", "error"
        )
        return
    }
    srSourceHandle.Write(srSourceText)
    srSourceHandle.Close()
    srMediaHandle := FileOpen(srPathsFile, "w", "UTF-8-RAW")
    if srMediaHandle {
        for srMediaItem in srReaderState["media"] {
            srMediaPath := srMediaItem["path"]
            if FileExist(srMediaPath)
                srMediaHandle.Write(srMediaPath "`n")
        }
        srMediaHandle.Close()
    }
    srPromptPath := ExplainProfilePath(srPromptName)
    if !FileExist(srPromptPath) {
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd,
            "The selected explanation prompt no longer exists.",
            "Generate new explanation version", "ok", "error"
        )
        return
    }

    SplitPath(srReaderState["database"],, &srStudyDir)
    srEnvironmentNames := [
        "EXPLAIN_PROVIDER", "EXPLAIN_MODEL", "GEMINI_EXPLAIN_MODEL",
        "EXPLAIN_PROMPT_FILE", "EXPLAIN_PROMPT_PROFILE",
        "EXPLAIN_SOURCE_TEXT_FILE", "EXPLAIN_SOURCE_PATHS_FILE",
        "EXPLAIN_UPDATE_OVERLAY", "SAVE_EXPLAINS", "SAVE_STUDY_LIBRARY",
        "STUDY_LIBRARY_SCREENSHOTS", "STUDY_LIBRARY_DIR",
        "STUDY_LIBRARY_PROFILE", "SETTINGS_DIR", "JRPG_DEBUG",
        "PYTHONIOENCODING"
    ]
    srOldEnvironment := Map()
    for srEnvironmentName in srEnvironmentNames
        srOldEnvironment[srEnvironmentName] := EnvGet(srEnvironmentName)

    srOutFile := srReaderState["outputDir"] "\\new_version_out.txt"
    srErrFile := srReaderState["outputDir"] "\\new_version_err.txt"
    try FileDelete(srOutFile)
    try FileDelete(srErrFile)
    srOldVersionCount := srReaderState["versions"].Length
    srSectionKey := StudyReaderCurrentSectionKey(srReaderState)
    srNewState["generateButton"].Enabled := false
    srNewState["cancelButton"].Enabled := false
    srNewState["status"].Value := "Generating another explanation..."
    try DllCall(
        "user32\\RedrawWindow", "ptr", srNewState["gui"].Hwnd,
        "ptr", 0, "ptr", 0, "uint", 0x185
    )
    Toast("Generating another explanation…")
    srExitCode := -1
    try {
        EnvSet("EXPLAIN_PROVIDER", srProvider)
        if (srProvider = "gemini") {
            srGeminiModel := srModel
            if (SubStr(srGeminiModel, 1, 7) != "models/")
                srGeminiModel := "models/" srGeminiModel
            EnvSet("GEMINI_EXPLAIN_MODEL", srGeminiModel)
            EnvSet("EXPLAIN_MODEL", "")
        } else {
            EnvSet("EXPLAIN_MODEL", srModel)
            EnvSet("GEMINI_EXPLAIN_MODEL", "")
        }
        EnvSet("EXPLAIN_PROMPT_FILE", srPromptPath)
        EnvSet("EXPLAIN_PROMPT_PROFILE", srPromptName)
        EnvSet("EXPLAIN_SOURCE_TEXT_FILE", srSourceFile)
        EnvSet("EXPLAIN_SOURCE_PATHS_FILE", srPathsFile)
        EnvSet("EXPLAIN_UPDATE_OVERLAY", "0")
        EnvSet("SAVE_EXPLAINS", "0")
        EnvSet("SAVE_STUDY_LIBRARY", "1")
        EnvSet("STUDY_LIBRARY_SCREENSHOTS", "1")
        EnvSet("STUDY_LIBRARY_DIR", srStudyDir)
        EnvSet("STUDY_LIBRARY_PROFILE", srReaderState["currentProfile"])
        EnvSet("SETTINGS_DIR", A_ScriptDir "\\Settings")
        EnvSet("JRPG_DEBUG", debugMode ? "1" : "0")
        EnvSet("PYTHONIOENCODING", "utf-8")
        srCommand := Format(
            'cmd /c chcp 65001>nul & "{1}" "{2}" 1>"{3}" 2>"{4}"',
            srPython, srExplainer, srOutFile, srErrFile
        )
        DbgCP("Study Reader new version -> " srCommand)
        srExitCode := RunWait(srCommand, , "Hide")
    } finally {
        for srEnvironmentName, srEnvironmentValue in srOldEnvironment
            EnvSet(srEnvironmentName, srEnvironmentValue)
    }
    srOutput := FileExist(srOutFile)
        ? Trim(FileRead(srOutFile, "UTF-8")) : ""
    srError := FileExist(srErrFile)
        ? Trim(FileRead(srErrFile, "UTF-8")) : ""
    if (srExitCode != 0) {
        srNewState["generateButton"].Enabled := true
        srNewState["cancelButton"].Enabled := true
        srNewState["status"].Value := "Generation failed."
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd,
            "The new version could not be generated.`n`n"
                . (srError != "" ? srError : srOutput),
            "Generate new explanation version", "ok", "error", 650
        )
        return
    }

    StudyLibraryLoadGroup(
        srReaderState, srReaderState["currentGroupId"], 0, srSectionKey
    )
    if (srReaderState["versions"].Length <= srOldVersionCount) {
        srNewState["generateButton"].Enabled := true
        srNewState["cancelButton"].Enabled := true
        srNewState["status"].Value := "The Study Library was not updated."
        CPThemedOwnedMessage(
            srNewState["gui"].Hwnd,
            "The model returned successfully, but no new Study Library version "
                . "was found.`n`n" . (srError != "" ? srError : srOutput),
            "Generate new explanation version", "ok", "error", 650
        )
        return
    }
    if (CPStudyLibraryState && CPStudyLibraryState.Has("gui")) {
        try StudyLibraryRefresh(CPStudyLibraryState)
        if (CPStudyLibraryState["currentGroupId"]
            = srReaderState["currentGroupId"])
            try StudyLibraryLoadGroup(
                CPStudyLibraryState, srReaderState["currentGroupId"], 0
            )
    }
    StudyReaderCloseNewVersionDialog(srNewState)
    Toast("New explanation version added")
}

StudyReaderOpenNewVersionDialog(srState, *) {
    global model_openai_explain, model_gemini_explain
    global explainProvider, explainOpenAIModel, explainGeminiModel
    global explainPromptProfile
    if ((srState.Has("editing") && srState["editing"])
        || srState["currentGroupId"] <= 0)
        return
    if (srState.Has("newVersionDialog")
        && IsObject(srState["newVersionDialog"])) {
        try {
            WinActivate("ahk_id " srState["newVersionDialog"]["gui"].Hwnd)
            return
        }
    }

    srNewGui := Gui(
        "+Owner" srState["gui"].Hwnd " +OwnDialogs",
        "Generate new explanation version"
    )
    srNewGui.MarginX := 18, srNewGui.MarginY := 16
    srNewGui.SetFont("s10", "Segoe UI")
    srHeading := srNewGui.Add(
        "Text", "xm ym w680 h28", "Generate new explanation version"
    )
    srHeading.SetFont("s12 Bold")
    srNewGui.Add(
        "Text", "xm y+4 w680 h42",
        "Creates another version from the archived Japanese and source "
            . "screenshots. Existing versions remain unchanged."
    )
    srNewGui.Add("Text", "xm y+12 w110", "AI provider:")
    srProvider := srNewGui.Add(
        "DropDownList", "x+8 yp-4 w300 0x210", ["Gemini", "OpenAI"]
    )
    srInitialProvider := StrLower(Trim(srState["currentProvider"]))
    if (srInitialProvider != "gemini" && srInitialProvider != "openai")
        srInitialProvider := StrLower(Trim(explainProvider))
    if (srInitialProvider != "gemini" && srInitialProvider != "openai")
        srInitialProvider := "gemini"
    srProvider.Choose(srInitialProvider = "gemini" ? 1 : 2)

    srNewGui.Add("Text", "xm y+13 w110", "Gemini model:")
    srGeminiModel := srNewGui.Add(
        "DropDownList", "x+8 yp-4 w300 0x210", model_gemini_explain
    )
    srGeminiDefault := srInitialProvider = "gemini"
        ? Trim(srState["currentModel"]) : explainGeminiModel
    if (SubStr(srGeminiDefault, 1, 7) = "models/")
        srGeminiDefault := SubStr(srGeminiDefault, 8)
    SetComboToExistingItem(
        srGeminiModel, model_gemini_explain, srGeminiDefault
    )

    srNewGui.Add("Text", "xm y+13 w110", "OpenAI model:")
    srOpenAIModel := srNewGui.Add(
        "DropDownList", "x+8 yp-4 w300 0x210", model_openai_explain
    )
    srOpenAIDefault := srInitialProvider = "openai"
        ? Trim(srState["currentModel"]) : explainOpenAIModel
    SetComboToExistingItem(
        srOpenAIModel, model_openai_explain, srOpenAIDefault
    )

    srNewGui.Add("Text", "xm y+13 w110", "Prompt:")
    srPrompt := srNewGui.Add(
        "DropDownList", "x+8 yp-4 w300 0x210", []
    )
    srPromptEdit := srNewGui.Add("Button", "x+8 yp w70", "Edit")
    srPromptAdd := srNewGui.Add("Button", "x+6 yp w70", "Add")
    srPromptDelete := srNewGui.Add("Button", "x+6 yp w70", "Delete")
    srStatus := srNewGui.Add("Text", "xm y+18 w680 h24", "Ready.")
    srGenerate := srNewGui.Add(
        "Button", "xm y+10 w190 h34 Default", "Generate new version"
    )
    srCancel := srNewGui.Add("Button", "x+10 yp w120 h34", "Cancel")

    srNewState := Map(
        "gui", srNewGui,
        "readerState", srState,
        "provider", srProvider,
        "geminiModel", srGeminiModel,
        "openAIModel", srOpenAIModel,
        "prompt", srPrompt,
        "status", srStatus,
        "generateButton", srGenerate,
        "cancelButton", srCancel,
        "closed", false
    )
    srState["newVersionDialog"] := srNewState
    srPromptDefault := Trim(srState["currentPrompt"])
    if (srPromptDefault = "")
        srPromptDefault := explainPromptProfile
    StudyReaderRefreshNewVersionPromptList(srNewState, srPromptDefault)
    StudyReaderNewVersionProviderChanged(srNewState)

    srProvider.OnEvent(
        "Change", StudyReaderNewVersionProviderChanged.Bind(srNewState)
    )
    srPromptEdit.OnEvent(
        "Click", StudyReaderEditNewVersionPrompt.Bind(srNewState)
    )
    srPromptAdd.OnEvent(
        "Click", StudyReaderAddNewVersionPrompt.Bind(srNewState)
    )
    srPromptDelete.OnEvent(
        "Click", StudyReaderDeleteNewVersionPrompt.Bind(srNewState)
    )
    srGenerate.OnEvent(
        "Click", StudyReaderGenerateNewVersion.Bind(srNewState)
    )
    srCancel.OnEvent(
        "Click", StudyReaderCloseNewVersionDialog.Bind(srNewState)
    )
    srNewGui.OnEvent(
        "Escape", StudyReaderCloseNewVersionDialog.Bind(srNewState)
    )
    srNewGui.OnEvent(
        "Close", StudyReaderCloseNewVersionDialog.Bind(srNewState)
    )
    srNewGui.Show("Hide AutoSize")
    CPApplyOwnedDialogTheme(srNewGui)
    srNewGui.Show("AutoSize Center")
}

StudyReaderAnkiChanged(srState, *) {
    global CPStudyLibraryState
    if ((srState.Has("editing") && srState["editing"])
        || srState["suspendAnki"] || srState["currentGroupId"] <= 0
        || srState["currentAnkiStatus"] = "found")
        return
    srDesired := srState["ankiCheck"].Value ? 1 : 0
    srPrevious := srState["currentAddedToAnkiAt"]
    srState["ankiCheck"].Enabled := false
    EnvSet("STUDY_LIBRARY_ADDED_TO_ANKI", srDesired ? "1" : "0")
    srUpdated := false
    try {
        srUpdated := StudyLibraryRunBridge(
            srState, "set-anki", srState["currentGroupId"]
        )
    } finally {
        EnvSet("STUDY_LIBRARY_ADDED_TO_ANKI", "")
    }
    if !srUpdated {
        srState["currentAddedToAnkiAt"] := srPrevious
        StudyReaderSyncEntryNavigation(srState)
        return
    }
    srMutationRows := StudyLibraryReadRows(srState["outputDir"] "\mutation.tsv")
    if (srMutationRows.Length && srMutationRows[1].Length >= 5)
        srState["currentAddedToAnkiAt"] := StudyLibraryHexDecode(
            srMutationRows[1][5]
        )
    else
        srState["currentAddedToAnkiAt"] := srDesired ? "Added" : ""
    StudyLibraryUpdateMetadataText(srState)
    StudyReaderSyncEntryNavigation(srState)
    if (CPStudyLibraryState && CPStudyLibraryState.Has("gui"))
        try StudyLibraryRefresh(CPStudyLibraryState)
}

StudyReaderToggleAnki(srState, *) {
    if ((srState.Has("editing") && srState["editing"])
        || srState["currentGroupId"] <= 0
        || srState["currentAnkiStatus"] = "found")
        return
    srState["ankiCheck"].Value := srState["ankiCheck"].Value ? 0 : 1
    StudyReaderAnkiChanged(srState)
}

StudyReaderSetEditing(srState, srEditing) {
    srState["editing"] := srEditing
    srState["explanation"].Opt(srEditing ? "-ReadOnly" : "+ReadOnly")
    srState["copyButton"].Visible := !srEditing
    srState["addAnkiButton"].Visible := !srEditing
    srState["editExplanationButton"].Visible := !srEditing
    srState["saveEditButton"].Visible := srEditing
    srState["cancelEditButton"].Visible := srEditing
    for srControl in [
        srState["previousEntry"], srState["nextEntry"], srState["ankiCheck"],
        srState["previousVersion"], srState["nextVersion"],
        srState["newVersionButton"],
        srState["fullSectionButton"], srState["previousSection"],
        srState["nextSection"]
    ]
        srControl.Enabled := !srEditing
    if srEditing {
        srState["explanationLabel"].Value := srState["editMode"] = "section"
            ? "Editing section (heading protected)"
            : "Editing full explanation (section headings affect navigation)"
        srState["explanation"].Focus()
    } else {
        srState["explanationLabel"].Value := "Explanation"
        StudyLibrarySyncVersionNavigation(srState)
        StudyReaderSyncEntryNavigation(srState)
        StudyReaderSyncSectionNavigation(srState)
        srState["copyButton"].Enabled := srState["currentGroupId"] > 0
        srState["addAnkiButton"].Enabled := srState["currentGroupId"] > 0
        srState["editExplanationButton"].Enabled := srState["currentGroupId"] > 0
    }
    try {
        srState["gui"].GetClientPos(,, &srClientW, &srClientH)
        StudyReaderResize(srState, srState["gui"], 0, srClientW, srClientH)
    }
}

StudyReaderBeginSectionEdit(srState, *) {
    srIndex := srState["sectionDdl"].Value
    if (srState["currentGroupId"] <= 0 || srIndex <= 1
        || srIndex > srState["sections"].Length)
        return
    srSection := srState["sections"][srIndex]
    srState["editMode"] := "section"
    srState["editSectionIndex"] := srIndex
    srState["editSectionKey"] := srSection["key"]
    srState["editOriginalText"] := srSection["content"]
    srState["explanation"].Value := srSection["content"]
    StudyReaderSetEditing(srState, true)
}

StudyReaderBeginFullEdit(srState, *) {
    if (srState["currentGroupId"] <= 0 || !srState["sections"].Length)
        return
    srWarning := "Full-explanation editing also allows section headings to be changed. "
        . "Those headings are used for the Reader's section navigation, so deleting "
        . "or renaming them may reduce the available sections.`n`nContinue?"
    if (CPThemedOwnedMessage(
        srState["gui"].Hwnd, srWarning, "Edit full explanation",
        "yesno", "warning"
    ) != "Yes")
        return
    srState["editMode"] := "full"
    srState["editSectionIndex"] := 1
    srState["editSectionKey"] := "full"
    srState["editOriginalText"] := srState["sections"][1]["content"]
    srState["explanation"].Value := srState["sections"][1]["content"]
    StudyReaderSetEditing(srState, true)
}

StudyReaderRemoveJapaneseReadings(srText) {
    ; Remove only kana parentheticals attached to a word containing kanji.
    ; Ordinary explanatory parentheses, Latin abbreviations and translations
    ; therefore remain untouched.
    srReadingPattern := "([一-龯々〆ヵヶ][一-龯々〆ヵヶぁ-ゖー]*)[ \t]*[\(（][ぁ-ゖー・ \t]+[\)）]"
    return RegExReplace(srText, srReadingPattern, "$1")
}

StudyReaderRestoreCopyButton(srState, srSerial, *) {
    try {
        if (!srState.Has("copyFeedbackSerial")
            || srState["copyFeedbackSerial"] != srSerial)
            return
        srState["copyButton"].Text := "Copy..."
    }
}

StudyReaderCopyText(srState, srText, *) {
    srText := Trim(srText, " `t`r`n")
    if (srText = "")
        return
    A_Clipboard := srText
    srSerial := srState.Has("copyFeedbackSerial")
        ? srState["copyFeedbackSerial"] + 1 : 1
    srState["copyFeedbackSerial"] := srSerial
    srState["copyButton"].Text := "Copied!"
    SetTimer(StudyReaderRestoreCopyButton.Bind(srState, srSerial), -1100)
}

StudyReaderCopyJapaneseWithoutReadings(srState, srText, *) {
    StudyReaderCopyText(srState, StudyReaderRemoveJapaneseReadings(srText))
}

StudyReaderShowCopyMenu(srState, *) {
    if ((srState.Has("editing") && srState["editing"])
        || srState["currentGroupId"] <= 0
        || !srState["sections"].Length)
        return
    srSectionIndex := srState["sectionDdl"].Value
    if (srSectionIndex < 1 || srSectionIndex > srState["sections"].Length)
        return
    srCurrentSection := srState["sections"][srSectionIndex]
    srChoices := ["Copy current section", "Copy full explanation"]
    srActions := [
        StudyReaderCopyText.Bind(srState, srCurrentSection["content"]),
        StudyReaderCopyText.Bind(srState, srState["sections"][1]["content"])
    ]
    if (srCurrentSection["key"] = "original") {
        srChoices.Push("Copy Japanese without readings")
        srActions.Push(StudyReaderCopyJapaneseWithoutReadings.Bind(
            srState, srCurrentSection["content"]
        ))
    }
    srChoice := CPThemedChoicePopup(
        srState["gui"].Hwnd, srState["copyButton"].Hwnd, srChoices
    )
    if (srChoice >= 1 && srChoice <= srActions.Length)
        srActions[srChoice].Call()
}

StudyReaderShowEditMenu(srState, *) {
    if ((srState.Has("editing") && srState["editing"])
        || srState["currentGroupId"] <= 0)
        return
    srChoices := []
    srActions := []
    srSectionIndex := srState["sectionDdl"].Value
    if (srSectionIndex > 1 && srSectionIndex <= srState["sections"].Length) {
        srChoices.Push("Edit current section...")
        srActions.Push(StudyReaderBeginSectionEdit.Bind(srState))
    }
    srChoices.Push("Edit full explanation...")
    srActions.Push(StudyReaderBeginFullEdit.Bind(srState))
    if (srState.Has("currentManuallyEditedAt")
        && srState["currentManuallyEditedAt"] != "") {
        srChoices.Push("Revert manual edits...")
        srActions.Push(StudyReaderRevertEdit.Bind(srState))
    }
    srChoice := CPThemedChoicePopup(
        srState["gui"].Hwnd, srState["editExplanationButton"].Hwnd,
        srChoices
    )
    if (srChoice >= 1 && srChoice <= srActions.Length)
        srActions[srChoice].Call()
}

StudyReaderCancelEdit(srState, *) {
    if !(srState.Has("editing") && srState["editing"])
        return
    StudyReaderSetEditing(srState, false)
    StudyLibrarySectionChanged(srState)
}

StudyReaderWriteEditFile(srState, srText) {
    srPath := srState["outputDir"] "\manual_edit.txt"
    try FileDelete(srPath)
    srFile := 0
    try {
        srFile := FileOpen(srPath, "w", "UTF-8-RAW")
        if !srFile
            throw Error("The temporary edit file could not be created.")
        srFile.Write(srText)
        srFile.Close()
        return true
    } catch as srWriteError {
        try srFile.Close()
        CPThemedOwnedMessage(
            srState["gui"].Hwnd,
            "The edited explanation could not be prepared for saving.`n`n"
            . srWriteError.Message,
            "Edit explanation", "ok", "error"
        )
        return false
    }
}

StudyReaderRestoreSectionAfterReload(srState, srSectionKey) {
    srTargetIndex := 1
    if (srSectionKey != "full") {
        for srIndex, srSection in srState["sections"] {
            if (srSection["key"] = srSectionKey) {
                srTargetIndex := srIndex
                break
            }
        }
    }
    srState["sectionDdl"].Choose(srTargetIndex)
    StudyLibrarySectionChanged(srState)
}

StudyReaderReloadAfterEdit(srState, srSectionKey) {
    global CPStudyLibraryState
    srGroupId := srState["currentGroupId"]
    srVersion := srState["currentVersion"]
    StudyLibraryLoadGroup(srState, srGroupId, srVersion)
    StudyReaderRestoreSectionAfterReload(srState, srSectionKey)
    if (CPStudyLibraryState && CPStudyLibraryState.Has("gui")
        && CPStudyLibraryState["currentGroupId"] = srGroupId) {
        try StudyLibraryLoadGroup(CPStudyLibraryState, srGroupId, srVersion)
    }
}

StudyReaderSaveEdit(srState, *) {
    if !(srState.Has("editing") && srState["editing"])
        return
    srEditedText := srState["explanation"].Value
    srSectionKey := srState["editSectionKey"]
    if (srState["editMode"] = "full") {
        if (Trim(srEditedText) = "") {
            CPThemedOwnedMessage(
                srState["gui"].Hwnd, "An explanation cannot be saved empty.",
                "Edit explanation", "ok", "warning"
            )
            return
        }
        srUpdatedFullText := srEditedText
    } else {
        srOldSectionText := srState["editOriginalText"]
        srFullText := srState["sections"][1]["content"]
        if (srOldSectionText = "") {
            CPThemedOwnedMessage(
                srState["gui"].Hwnd,
                "This section cannot be replaced safely. Use Edit full explanation instead.",
                "Edit explanation", "ok", "warning"
            )
            return
        }
        srFirstPos := InStr(srFullText, srOldSectionText, true)
        srSecondPos := srFirstPos
            ? InStr(srFullText, srOldSectionText, true,
                srFirstPos + StrLen(srOldSectionText))
            : 0
        if (!srFirstPos || srSecondPos) {
            CPThemedOwnedMessage(
                srState["gui"].Hwnd,
                "The selected section could not be matched uniquely in the full "
                . "explanation. No change was made; use Edit full explanation instead.",
                "Edit explanation", "ok", "warning"
            )
            return
        }
        srUpdatedFullText := SubStr(srFullText, 1, srFirstPos - 1)
            . srEditedText
            . SubStr(srFullText, srFirstPos + StrLen(srOldSectionText))
    }
    if (srUpdatedFullText = srState["sections"][1]["content"])
        return StudyReaderCancelEdit(srState)
    if !StudyReaderWriteEditFile(srState, srUpdatedFullText)
        return

    srState["saveEditButton"].Enabled := false
    srSaved := StudyLibraryRunBridge(
        srState, "save-edit", srState["currentGroupId"], srState["currentVersion"]
    )
    srState["saveEditButton"].Enabled := true
    if !srSaved
        return
    StudyReaderSetEditing(srState, false)
    StudyReaderReloadAfterEdit(srState, srSectionKey)
}

StudyReaderRevertEdit(srState, *) {
    if (srState["currentGroupId"] <= 0
        || !srState.Has("currentManuallyEditedAt")
        || srState["currentManuallyEditedAt"] = "")
        return
    srQuestion := "Discard all manual changes to this generated version and restore "
        . "the untouched model output?"
    if (CPThemedOwnedMessage(
        srState["gui"].Hwnd, srQuestion, "Revert manual edits",
        "yesno", "warning"
    ) != "Yes")
        return
    srSectionIndex := srState["sectionDdl"].Value
    srSectionKey := (srSectionIndex >= 1
        && srSectionIndex <= srState["sections"].Length)
        ? srState["sections"][srSectionIndex]["key"] : "full"
    if !StudyLibraryRunBridge(
        srState, "revert-edit", srState["currentGroupId"], srState["currentVersion"]
    )
        return
    StudyReaderReloadAfterEdit(srState, srSectionKey)
}

StudyReaderBindHotkeys(srState) {
    srContext := "ahk_id " srState["gui"].Hwnd
    srHotkeys := Map(
        "!Left", StudyReaderStepEntry.Bind(srState, -1),
        "!Right", StudyReaderStepEntry.Bind(srState, 1),
        "^Left", StudyReaderStepSection.Bind(srState, -1),
        "^Right", StudyReaderStepSection.Bind(srState, 1),
        "!Home", StudyReaderSelectFullSection.Bind(srState),
        "!a", StudyReaderToggleAnki.Bind(srState)
    )
    HotIfWinActive(srContext)
    for srKey, srCallback in srHotkeys
        Hotkey(srKey, srCallback, "On")
    HotIf()
    srState["hotkeyContext"] := srContext
    srState["hotkeys"] := srHotkeys
}

StudyReaderUnbindHotkeys(srState) {
    if !(srState.Has("hotkeys") && srState.Has("hotkeyContext"))
        return
    HotIfWinActive(srState["hotkeyContext"])
    for srKey, srCallback in srState["hotkeys"]
        try Hotkey(srKey, "Off")
    HotIf()
    srState["hotkeys"] := Map()
}

StudyReaderRedraw(srState, srErase := true, *) {
    StudyLibraryRedraw(srState, srErase)
}

StudyReaderResize(srState, srGui, srMinMax, srWidth, srHeight) {
    if (srMinMax = -1 || srWidth < 600 || srHeight < 400)
        return
    srHwnd := srGui.Hwnd
    srVisible := srHwnd
        && DllCall("user32\IsWindowVisible", "ptr", srHwnd, "int")
    srInteractive := srState.Has("interactiveResize")
        && srState["interactiveResize"]
    if srInteractive {
        ; Native Edit controls and the screenshot preview are expensive to
        ; continuously move and repaint inside Windows' modal sizing loop.
        ; Keep the existing reader layout stable while the frame is dragged;
        ; WM_EXITSIZEMOVE performs one exact layout at the released size.
        srState["deferredResizeWidth"] := srWidth
        srState["deferredResizeHeight"] := srHeight
        return
    }
    if srVisible
        DllCall("user32\SendMessageW", "ptr", srHwnd, "uint", 0x000B,
            "ptr", 0, "ptr", 0) ; WM_SETREDRAW off
    try {
        srMargin := 14, srGap := 14
        srContextW := Max(270, Min(420, Round(srWidth * 0.33)))
        srLeftW := Max(300, srWidth - srMargin * 2 - srGap - srContextW)
        srContextX := srMargin + srLeftW + srGap
        srBottom := srHeight - srMargin

        srState["detailTitle"].Move(srMargin, srMargin, srLeftW, 26)
        srEntryY := srMargin + 31
        srState["previousEntry"].Move(srMargin, srEntryY, 34, 30)
        srState["entryStatus"].Move(srMargin + 40, srEntryY, 100, 30)
        srState["nextEntry"].Move(srMargin + 146, srEntryY, 34, 30)
        srState["ankiCheck"].Move(
            srMargin + 190, srEntryY + 4, Max(120, srLeftW - 190), 24
        )
        srToolbarY := srEntryY + 36
        srVersionLabelW := 58, srArrowW := 34, srNavGap := 6
        srNewVersionW := 104
        srVersionW := Min(
            300,
            Max(110, srLeftW - srVersionLabelW - srArrowW * 2
                - srNavGap * 2 - srNewVersionW - 14)
        )
        srState["versionLabel"].Move(srMargin, srToolbarY + 4, srVersionLabelW, 22)
        srVersionPrevX := srMargin + srVersionLabelW
        srState["previousVersion"].Move(
            srVersionPrevX, srToolbarY, srArrowW, 30
        )
        srVersionViewX := srVersionPrevX + srArrowW + srNavGap
        srState["versionView"].Move(
            srVersionViewX, srToolbarY, srVersionW, 30
        )
        srState["nextVersion"].Move(
            srVersionViewX + srVersionW + srNavGap,
            srToolbarY, srArrowW, 30
        )
        srState["newVersionButton"].Move(
            srVersionViewX + srVersionW + srNavGap + srArrowW + 8,
            srToolbarY, srNewVersionW, 30
        )
        srSectionNavY := srToolbarY + 36
        srFullButtonW := 128
        srState["fullSectionButton"].Move(
            srMargin, srSectionNavY, srFullButtonW, 30
        )
        srPreviousX := srMargin + srFullButtonW + 8
        srState["previousSection"].Move(
            srPreviousX, srSectionNavY, srArrowW, 30
        )
        srSectionStatusX := srPreviousX + srArrowW + srNavGap
        ; Wide enough for the longest expected localized section heading, but
        ; no longer stretches across the complete Reader on large displays.
        srSectionStatusW := Min(
            420,
            Max(70, srLeftW - (srSectionStatusX - srMargin) - srArrowW - srNavGap)
        )
        srState["sectionStatus"].Move(
            srSectionStatusX, srSectionNavY, srSectionStatusW, 30
        )
        srNextX := srSectionStatusX + srSectionStatusW + srNavGap
        srState["nextSection"].Move(srNextX, srSectionNavY, srArrowW, 30)
        srExplanationLabelY := srSectionNavY + 35
        srAddAnkiButtonW := 108, srCopyButtonW := 88, srEditButtonW := 88
        srSaveButtonW := 72, srCancelButtonW := 78
        srEditButtonY := srExplanationLabelY - 3
        srEditButtonX := srMargin + srLeftW - srEditButtonW
        srCopyButtonX := srEditButtonX - 6 - srCopyButtonW
        srAddAnkiButtonX := srCopyButtonX - 6 - srAddAnkiButtonW
        srState["addAnkiButton"].Move(
            srAddAnkiButtonX, srEditButtonY, srAddAnkiButtonW, 27
        )
        srState["copyButton"].Move(
            srCopyButtonX, srEditButtonY, srCopyButtonW, 27
        )
        srState["editExplanationButton"].Move(
            srEditButtonX, srEditButtonY, srEditButtonW, 27
        )
        srCancelButtonX := srMargin + srLeftW - srCancelButtonW
        srSaveButtonX := srCancelButtonX - 6 - srSaveButtonW
        srState["saveEditButton"].Move(
            srSaveButtonX, srEditButtonY, srSaveButtonW, 27
        )
        srState["cancelEditButton"].Move(
            srCancelButtonX, srEditButtonY, srCancelButtonW, 27
        )
        srLabelReserve := srState.Has("editing") && srState["editing"]
            ? srSaveButtonW + srCancelButtonW + 14
            : srAddAnkiButtonW + srCopyButtonW + srEditButtonW + 20
        srState["explanationLabel"].Move(
            srMargin, srExplanationLabelY, Max(60, srLeftW - srLabelReserve), 22
        )
        srExplanationY := srExplanationLabelY + 28
        srState["explanation"].Move(
            srMargin, srExplanationY, srLeftW,
            Max(80, srBottom - srExplanationY)
        )

        srState["contextTitle"].Move(srContextX, srMargin, srContextW, 26)
        srMetadataY := srMargin + 31
        srState["metadata"].Move(srContextX, srMetadataY, srContextW, 72)
        srImageLabelY := srMetadataY + 76
        srState["imageLabel"].Move(srContextX, srImageLabelY, srContextW, 20)
        srImageY := srImageLabelY + 20
        srMaxImageH := Max(90, srBottom - srImageY - 27 - 5 - 20 - 90)
        srDesiredImageH := Round((srBottom - srImageY) * 0.43)
        srImageH := Max(90, Min(300, Min(srMaxImageH, srDesiredImageH)))
        srState["imageFrame"].Move(srContextX, srImageY, srContextW, srImageH)
        srState["imageArea"] := Map(
            "x", srContextX + 4,
            "y", srImageY + 4,
            "w", Max(1, srContextW - 8),
            "h", Max(1, srImageH - 8)
        )
        srNavY := srImageY + srImageH + 5
        srImageInfoW := Max(70, srContextW - 200)
        srState["previousImage"].Move(srContextX, srNavY, 34, 27)
        srState["imageInfo"].Move(
            srContextX + 40, srNavY + 4, srImageInfoW, 22
        )
        srState["nextImage"].Move(
            srContextX + 40 + srImageInfoW + 6, srNavY, 34, 27
        )
        srState["openImage"].Move(
            srContextX + srContextW - 112, srNavY, 112, 27
        )
        srSourceLabelY := srNavY + 32
        srState["sourceLabel"].Move(srContextX, srSourceLabelY, srContextW, 20)
        srSourceY := srSourceLabelY + 20
        srState["source"].Move(
            srContextX, srSourceY, srContextW,
            Max(60, srBottom - srSourceY)
        )
        if srVisible
            StudyLibraryQueueImageLayout(srState, 80)
    } finally {
        if srVisible {
            DllCall("user32\SendMessageW", "ptr", srHwnd, "uint", 0x000B,
                "ptr", 1, "ptr", 0) ; WM_SETREDRAW on
            StudyReaderRedraw(srState, true)
        }
    }
    if srVisible
        SetTimer(srState["redrawCallback"], -160)
}

StudyReaderSaveBounds(srState) {
    global iniPath
    if !(srState.Has("gui") && srState["gui"] && srState["gui"].Hwnd)
        return
    try {
        if (WinGetMinMax("ahk_id " srState["gui"].Hwnd) != 0)
            return
        srState["gui"].GetPos(&srX, &srY)
        srState["gui"].GetClientPos(,, &srW, &srH)
        if (srW < 600 || srH < 400)
            return
        IniWrite(srX, iniPath, "study_reader_view", "x")
        IniWrite(srY, iniPath, "study_reader_view", "y")
        IniWrite(srW, iniPath, "study_reader_view", "w")
        IniWrite(srH, iniPath, "study_reader_view", "h")
    }
}

StudyReaderClampPosition(srGui, &srX, &srY) {
    srGui.GetPos(,, &srOuterW, &srOuterH)
    srDpi := GetWindowDPI(srGui.Hwnd)
    srScale := 96 / Max(96, srDpi)
    srBestMonitor := 0, srBestArea := 0
    try {
        Loop MonitorGetCount() {
            MonitorGetWorkArea(A_Index, &srLeft, &srTop, &srRight, &srBottom)
            srLeft := Ceil(srLeft * srScale), srTop := Ceil(srTop * srScale)
            srRight := Floor(srRight * srScale), srBottom := Floor(srBottom * srScale)
            srIntersectionW := Max(0,
                Min(srX + srOuterW, srRight) - Max(srX, srLeft)
            )
            srIntersectionH := Max(0,
                Min(srY + srOuterH, srBottom) - Max(srY, srTop)
            )
            srArea := srIntersectionW * srIntersectionH
            if (srArea > srBestArea) {
                srBestArea := srArea
                srBestMonitor := A_Index
            }
        }
    }
    if !srBestMonitor {
        try srBestMonitor := MonitorGetPrimary()
        catch
            srBestMonitor := 1
    }
    try {
        MonitorGetWorkArea(
            srBestMonitor, &srLeft, &srTop, &srRight, &srBottom
        )
        srLeft := Ceil(srLeft * srScale), srTop := Ceil(srTop * srScale)
        srRight := Floor(srRight * srScale), srBottom := Floor(srBottom * srScale)
        srX := Min(Max(srX, srLeft), Max(srLeft, srRight - srOuterW))
        srY := Min(Max(srY, srTop), Max(srTop, srBottom - srOuterH))
    }
}

StudyWindowRevealFinished(studyGui, positionOptions := "") {
    if !(IsObject(studyGui) && studyGui.Hwnd)
        return
    ; Native ListView/combo/button surfaces are not fully composed while their
    ; parent is hidden. Paint one complete frame just beyond the virtual desktop
    ; so the first on-screen frame is already themed and populated.
    studyGui.GetPos(,, &studyWarmW, &studyWarmH)
    studyVirtualLeft := DllCall("user32\GetSystemMetrics", "int", 76, "int")
    studyVirtualTop := DllCall("user32\GetSystemMetrics", "int", 77, "int")
    studyVirtualW := DllCall("user32\GetSystemMetrics", "int", 78, "int")
    studyWarmX := studyVirtualLeft + studyVirtualW + Max(100, studyWarmW)
    studyGui.Show("NA x" studyWarmX " y" studyVirtualTop)
    ; Several native controls create their final themed surfaces only on the
    ; first visible Show. Reapply the palette while that Show is safely
    ; off-screen so buttons, combo boxes and scrollbars do not retain their
    ; default light appearance when moved on-screen.
    CPApplyOwnedDialogTheme(studyGui)
    try DllCall("user32\RedrawWindow", "ptr", studyGui.Hwnd, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400)
    try DllCall("dwmapi\DwmFlush")
    studyGui.Hide()

    ; Establish the final bounds while the window is still hidden. Revealing it
    ; without another geometry change prevents the first-frame position jump.
    finalOptions := "Hide"
    if (Trim(positionOptions) != "")
        finalOptions .= " " positionOptions
    studyGui.Show(finalOptions)
    ; All painting and native theming has already happened while hidden. Avoid
    ; another full parent/child erase here: on some Windows/DPI combinations it
    ; exposes a white intermediate surface during the first composition frame.
    studyGui.Show("NA")
}

StudyWindowEnableLiveResize(studyState, *) {
    if !StudyLibraryStateAlive(studyState)
        return
    if !(studyState.Has("sizeCallback") && studyState["sizeCallback"])
        return
    try studyState["gui"].OnEvent("Size", studyState["sizeCallback"])
}

StudyWindowSizeMoveMessage(wParam, lParam, msg, hwnd) {
    global CPStudyLibraryState, CPStudyReaderState
    for studyState in [CPStudyLibraryState, CPStudyReaderState] {
        if !StudyLibraryStateAlive(studyState)
            continue
        if (studyState["gui"].Hwnd != hwnd)
            continue
        if (msg = 0x0231) { ; WM_ENTERSIZEMOVE
            studyState["interactiveResize"] := true
            ; Windows enters the same modal loop for a title-bar move and a
            ; border resize. Remember the client dimensions so a plain move
            ; does not trigger the expensive final layout/redraw on exit.
            studyStartW := 0, studyStartH := 0
            try studyState["gui"].GetClientPos(
                , , &studyStartW, &studyStartH
            )
            studyState["sizeMoveStartWidth"] := studyStartW
            studyState["sizeMoveStartHeight"] := studyStartH
            studyState["imageLayoutLastTick"] := 0
            try SetTimer(studyState["imageLayoutCallback"], 0)
            studyState["imageLayoutQueued"] := false
            try SetTimer(studyState["redrawCallback"], 0)
        } else if (msg = 0x0232) { ; WM_EXITSIZEMOVE
            studyState["interactiveResize"] := false
            studyFinalW := 0, studyFinalH := 0
            try studyState["gui"].GetClientPos(, , &studyFinalW, &studyFinalH)
            studySizeChanged := !studyState.Has("sizeMoveStartWidth")
                || !studyState.Has("sizeMoveStartHeight")
                || studyFinalW != studyState["sizeMoveStartWidth"]
                || studyFinalH != studyState["sizeMoveStartHeight"]
            if !studySizeChanged
                break
            ; The Size callback deliberately left all child controls untouched
            ; during the drag. Lay them out once at the final client size now.
            try {
                studyState["sizeCallback"].Call(
                    studyState["gui"], 0, studyFinalW, studyFinalH
                )
            }
            ; Finish with the exact final preview size and one complete frame.
            try SetTimer(studyState["imageLayoutCallback"], 0)
            studyState["imageLayoutQueued"] := false
            StudyLibraryShowImage(studyState)
            StudyLibraryRedraw(studyState, true)
        }
        break
    }
}

StudyReaderClose(srState, *) {
    global CPStudyReaderState
    if (srState.Has("editing") && srState["editing"]
        && CPThemedOwnedMessage(
            srState["gui"].Hwnd,
            "Discard the unsaved explanation edit?",
            "Unsaved explanation edit", "yesno", "warning"
        ) != "Yes")
        return
    try SetTimer(srState["imageLayoutCallback"], 0)
    try SetTimer(srState["redrawCallback"], 0)
    try StudyReaderRememberEntryView(srState)
    try StudyReaderUnbindHotkeys(srState)
    try StudyReaderSaveBounds(srState)
    if (srState.Has("newVersionDialog")
        && IsObject(srState["newVersionDialog"]))
        try StudyReaderCloseNewVersionDialog(srState["newVersionDialog"])
    try srState["gui"].Destroy()
    CPStudyReaderState := 0
    StudyStandaloneMaybeExit()
}

StudyLibraryOpenSelectedReader(slState, *) {
    if (slState["currentGroupId"] <= 0)
        return
    OpenStudyReader(
        slState["currentGroupId"], slState["currentVersion"],
        StudyLibraryGroupsInDisplayOrder(slState)
    )
}

StudyLibraryOpenReaderFromList(slState, slList, slRow, *) {
    if (slRow < 1 || slRow > slState["groups"].Length)
        return
    slGroupId := StudyLibraryRowGroupId(slState, slRow)
    if (slGroupId <= 0)
        return
    slList.Modify(slRow, "Select Focus Vis")
    OpenStudyReader(slGroupId, 0, StudyLibraryGroupsInDisplayOrder(slState))
}

StudyLibraryReaderSnapshot() {
    global studyLibraryDir
    srOutputDir := StudyBridgeCreateWorkspace("reader_startup")
    srState := Map(
        "database", studyLibraryDir "\study_library.db",
        "outputDir", srOutputDir
    )
    srEnvironment := Map(
        "STUDY_LIBRARY_QUERY", "",
        "STUDY_LIBRARY_PROFILE_MODE", "all",
        "STUDY_LIBRARY_PROFILE_FILTER", "",
        "STUDY_LIBRARY_CHAPTER_MODE", "all",
        "STUDY_LIBRARY_CHAPTER_FILTER", "",
        "STUDY_LIBRARY_SPEAKER_MODE", "all",
        "STUDY_LIBRARY_SPEAKER_FILTER", "",
        "STUDY_LIBRARY_TAG_MODE", "all",
        "STUDY_LIBRARY_TAG_FILTER", "",
        "STUDY_LIBRARY_ANKI_MODE", "all",
        "STUDY_LIBRARY_DATE_MODE", "all",
        "STUDY_LIBRARY_DATE_FROM", "",
        "STUDY_LIBRARY_DATE_TO", ""
    )
    srPreviousEnvironment := Map()
    for srName, srValue in srEnvironment {
        srPreviousEnvironment[srName] := EnvGet(srName)
        EnvSet(srName, srValue)
    }
    srSucceeded := false
    try {
        if StudyLibraryRunBridge(srState, "ensure")
            srSucceeded := StudyLibraryRunBridge(srState, "snapshot")
    }
    finally {
        for srName, srValue in srPreviousEnvironment
            EnvSet(srName, srValue)
    }
    srGroups := []
    try {
        if srSucceeded {
            for srRow in StudyLibraryReadRows(srOutputDir "\groups.tsv") {
                if (srRow.Length < 6)
                    continue
                srGroups.Push(Map(
                    "id", Integer(srRow[1]),
                    "addedToAnkiAt", srRow.Length >= 10
                        ? StudyLibraryHexDecode(srRow[10]) : "",
                    "ankiStatus", srRow.Length >= 11
                        ? StudyLibraryHexDecode(srRow[11]) : "not_checked"
                ))
            }
        }
    } finally {
        StudyBridgeCleanupWorkspace(srOutputDir)
    }
    return srGroups
}

OpenStandaloneStudyReader(*) {
    global iniPath
    srGroups := StudyLibraryReaderSnapshot()
    if !srGroups.Length {
        ; The Library provides a useful empty-state message and gives a new
        ; user access to all filters once explanations have been generated.
        OpenStandaloneStudyLibrary()
        return
    }
    srLastGroupId := 0
    try srLastGroupId := Integer(IniRead(
        iniPath, "study_reader_view", "lastGroupId", 0
    ))
    srTargetId := 0
    if (srLastGroupId > 0) {
        for srGroup in srGroups {
            if (srGroup["id"] = srLastGroupId) {
                srTargetId := srLastGroupId
                break
            }
        }
    }
    if !srTargetId {
        for srGroup in srGroups {
            if (srGroup["ankiStatus"] != "found"
                && srGroup["addedToAnkiAt"] = "") {
                srTargetId := srGroup["id"]
                break
            }
        }
    }
    if !srTargetId
        srTargetId := srGroups[1]["id"]
    OpenStudyReader(srTargetId, 0, srGroups)
}

OpenStudyReader(srGroupId, srVersion := 0, srGroups := 0) {
    global CPStudyReaderState, studyLibraryDir, iniPath
    if (srGroupId <= 0)
        return
    if (CPStudyReaderState && CPStudyReaderState.Has("gui")) {
        try {
            if IsObject(srGroups)
                StudyReaderSetEntrySequence(
                    CPStudyReaderState, srGroups, srGroupId
                )
            StudyLibraryLoadGroup(CPStudyReaderState, srGroupId, srVersion)
            CPStudyReaderState["gui"].Show()
            WinActivate("ahk_id " CPStudyReaderState["gui"].Hwnd)
            return
        }
    }

    DirCreate(studyLibraryDir)
    srOutputDir := StudyBridgeCreateWorkspace("reader")
    srGui := Gui(
        "+Resize +MinSize720x480 +OwnDialogs",
        "JRPG Translator - Study Reader"
    )
    srGui.MarginX := 14, srGui.MarginY := 14
    srGui.SetFont("s10", "Segoe UI")

    srDetailTitle := srGui.Add(
        "Text", "x14 y14 w650 h26 +0x200", "Saved explanation"
    )
    srDetailTitle.SetFont("s11 Bold")
    srPreviousEntry := srGui.Add("Button", "x14 y45 w34 h30 Disabled", "‹")
    srEntryStatus := srGui.Add(
        "Text", "x54 y45 w100 h30 Border Center +0x200", "No entry"
    )
    srNextEntry := srGui.Add("Button", "x160 y45 w34 h30 Disabled", "›")
    srAnkiCheck := srGui.Add(
        "CheckBox", "x204 y49 w210 h24 Disabled", "Added to Anki  (Alt+A)"
    )
    srVersionLabel := srGui.Add("Text", "x14 y85 w58", "Version:")
    ; Keep the original list hidden as the version data model. The visible
    ; read-only field and arrow buttons provide simpler navigation.
    srVersionDdl := srGui.Add(
        "DropDownList", "x0 y0 w1 h1 Hidden 0x210", []
    )
    srVersionView := srGui.Add(
        "Edit", "x72 y81 w220 h30 ReadOnly", ""
    )
    srPreviousVersion := srGui.Add(
        "Button", "x298 y81 w34 h30 Disabled", "‹"
    )
    srNextVersion := srGui.Add(
        "Button", "x338 y81 w34 h30 Disabled", "›"
    )
    srNewVersionButton := srGui.Add(
        "Button", "x380 y81 w104 h30 Disabled", "New version..."
    )
    ; The loader still uses this hidden list as its section index/data model.
    ; The Reader exposes the compact permanent navigation controls below.
    srSectionLabel := srGui.Add("Text", "x0 y0 w1 h1 Hidden", "Section:")
    srSectionDdl := srGui.Add(
        "DropDownList", "x0 y0 w1 h1 Hidden 0x210", ["Full explanation"]
    )
    srFullSectionButton := srGui.Add(
        "Button", "x14 y117 w128 h30 Disabled", "Full explanation"
    )
    srPreviousSection := srGui.Add("Button", "x150 y117 w34 h30 Disabled", "‹")
    srSectionStatus := srGui.Add(
        "Text",
        "x190 y117 w434 h30 Border Center +0x200 +0x4000",
        "No explanation selected"
    )
    srNextSection := srGui.Add("Button", "x630 y117 w34 h30 Disabled", "›")
    srExplanationLabel := srGui.Add(
        "Text", "x14 y152 w470 h22", "Explanation"
    )
    srExplanationLabel.SetFont("Bold")
    srCopyButton := srGui.Add(
        "Button", "x482 y149 w88 h27 Disabled", "Copy..."
    )
    srAddAnkiButton := srGui.Add(
        "Button", "x368 y149 w108 h27 Disabled", "Add to Anki..."
    )
    srEditExplanationButton := srGui.Add(
        "Button", "x576 y149 w88 h27 Disabled", "Edit..."
    )
    srSaveEditButton := srGui.Add(
        "Button", "x504 y149 w72 h27 Hidden", "Save"
    )
    srCancelEditButton := srGui.Add(
        "Button", "x582 y149 w82 h27 Hidden", "Cancel"
    )
    srExplanation := srGui.Add(
        "Edit", "x14 y180 w650 h484 ReadOnly Multi VScroll"
    )
    srExplanation.SetFont("s11", "Segoe UI")

    srContextTitle := srGui.Add("Text", "x678 y14 w380 h26", "Context")
    srContextTitle.SetFont("s11 Bold")
    srMetadata := srGui.Add("Text", "x678 y45 w380 h72 cGray", "")
    CPRegisterMutedControl(srMetadata)
    srImageLabel := srGui.Add(
        "Text", "x678 y121 w380 h20", "Source screenshot"
    )
    srImageFrame := srGui.Add(
        "Text", "x678 y141 w380 h250 Border Background101010", ""
    )
    srPicture := srGui.Add("Picture", "x682 y145 w1 h1 Hidden", "")
    srPreviousImage := srGui.Add("Button", "x678 y396 w34 h27", "‹")
    srNextImage := srGui.Add("Button", "x718 y396 w34 h27", "›")
    srImageInfo := srGui.Add(
        "Text", "x760 y400 w178", "No source screenshot selected."
    )
    srOpenImage := srGui.Add(
        "Button", "x946 y396 w112 h27", "Open full image"
    )
    srSourceLabel := srGui.Add(
        "Text", "x678 y428 w380 h20", "Original Japanese"
    )
    srSource := srGui.Add(
        "Edit", "x678 y448 w380 h216 ReadOnly Multi VScroll"
    )

    srState := Map(
        "gui", srGui,
        "database", studyLibraryDir "\study_library.db",
        "outputDir", srOutputDir,
        "detailTitle", srDetailTitle,
        "previousEntry", srPreviousEntry,
        "entryStatus", srEntryStatus,
        "nextEntry", srNextEntry,
        "ankiCheck", srAnkiCheck,
        "entryIds", [],
        "entryViewMemory", Map(),
        "suspendAnki", false,
        "versionLabel", srVersionLabel,
        "versionDdl", srVersionDdl,
        "versionView", srVersionView,
        "previousVersion", srPreviousVersion,
        "nextVersion", srNextVersion,
        "newVersionButton", srNewVersionButton,
        "versions", [],
        "sectionLabel", srSectionLabel,
        "sectionDdl", srSectionDdl,
        "sections", [],
        "fullSectionButton", srFullSectionButton,
        "previousSection", srPreviousSection,
        "sectionStatus", srSectionStatus,
        "nextSection", srNextSection,
        "explanationLabel", srExplanationLabel,
        "copyButton", srCopyButton,
        "addAnkiButton", srAddAnkiButton,
        "copyFeedbackSerial", 0,
        "editExplanationButton", srEditExplanationButton,
        "saveEditButton", srSaveEditButton,
        "cancelEditButton", srCancelEditButton,
        "explanation", srExplanation,
        "contextTitle", srContextTitle,
        "metadata", srMetadata,
        "imageLabel", srImageLabel,
        "imageFrame", srImageFrame,
        "picture", srPicture,
        "previousImage", srPreviousImage,
        "nextImage", srNextImage,
        "imageInfo", srImageInfo,
        "openImage", srOpenImage,
        "imageArea", Map("x", 682, "y", 145, "w", 372, "h", 242),
        "media", [],
        "mediaIndex", 0,
        "imageLayoutQueued", false,
        "imageLayoutLastTick", 0,
        "interactiveResize", false,
        "sourceLabel", srSourceLabel,
        "source", srSource,
        "currentGroupId", 0,
        "currentVersion", 0,
        "currentChapter", "",
        "currentSpeaker", "",
        "currentTags", "",
        "currentAddedToAnkiAt", "",
        "currentAnkiStatus", "not_checked",
        "currentAnkiNoteId", 0,
        "currentAnkiCheckedAt", "",
        "currentProfile", "",
        "currentProvider", "",
        "currentModel", "",
        "currentPrompt", "",
        "currentManuallyEditedAt", "",
        "editing", false,
        "editMode", "",
        "editSectionIndex", 1,
        "editSectionKey", "full",
        "editOriginalText", "",
        "newVersionDialog", 0,
        "suspend", false
    )
    srState["imageLayoutCallback"] := StudyLibraryShowImage.Bind(srState)
    srState["redrawCallback"] := StudyReaderRedraw.Bind(srState)
    srState["sizeCallback"] := StudyReaderResize.Bind(srState)
    srState["sectionNavCallback"] := StudyReaderSyncSectionNavigation.Bind(srState)
    srState["entryNavCallback"] := StudyReaderSyncEntryNavigation.Bind(srState)
    StudyReaderSetEntrySequence(srState, srGroups, srGroupId)
    CPStudyReaderState := srState

    srPreviousEntry.OnEvent("Click", StudyReaderStepEntry.Bind(srState, -1))
    srNextEntry.OnEvent("Click", StudyReaderStepEntry.Bind(srState, 1))
    srAnkiCheck.OnEvent("Click", StudyReaderAnkiChanged.Bind(srState))
    srPreviousVersion.OnEvent(
        "Click", StudyLibraryStepVersion.Bind(srState, -1)
    )
    srNextVersion.OnEvent(
        "Click", StudyLibraryStepVersion.Bind(srState, 1)
    )
    srNewVersionButton.OnEvent(
        "Click", StudyReaderOpenNewVersionDialog.Bind(srState)
    )
    srFullSectionButton.OnEvent(
        "Click", StudyReaderSelectFullSection.Bind(srState)
    )
    srPreviousSection.OnEvent(
        "Click", StudyReaderStepSection.Bind(srState, -1)
    )
    srNextSection.OnEvent(
        "Click", StudyReaderStepSection.Bind(srState, 1)
    )
    srCopyButton.OnEvent(
        "Click", StudyReaderShowCopyMenu.Bind(srState)
    )
    srAddAnkiButton.OnEvent(
        "Click", StudyReaderShowAnkiMenu.Bind(srState)
    )
    srEditExplanationButton.OnEvent(
        "Click", StudyReaderShowEditMenu.Bind(srState)
    )
    srSaveEditButton.OnEvent("Click", StudyReaderSaveEdit.Bind(srState))
    srCancelEditButton.OnEvent("Click", StudyReaderCancelEdit.Bind(srState))
    srPreviousImage.OnEvent("Click", StudyLibraryPreviousImage.Bind(srState))
    srNextImage.OnEvent("Click", StudyLibraryNextImage.Bind(srState))
    srOpenImage.OnEvent("Click", StudyLibraryOpenImage.Bind(srState))
    srGui.OnEvent("Escape", StudyReaderClose.Bind(srState))
    srGui.OnEvent("Close", StudyReaderClose.Bind(srState))
    StudyReaderBindHotkeys(srState)

    srGui.Show("Hide w1100 h700")
    srDpi := GetWindowDPI(srGui.Hwnd)
    srLogicalScale := 96 / Max(96, srDpi)
    srMaxW := Max(720, Floor(A_ScreenWidth * srLogicalScale) - 40)
    srMaxH := Max(480, Floor(A_ScreenHeight * srLogicalScale) - 60)
    srDefaultW := Max(720, Min(1200, srMaxW - 40))
    srDefaultH := Max(480, Min(760, srMaxH - 40))
    srSavedW := IniRead(iniPath, "study_reader_view", "w", srDefaultW)
    srSavedH := IniRead(iniPath, "study_reader_view", "h", srDefaultH)
    try srSavedW := Integer(srSavedW)
    catch
        srSavedW := srDefaultW
    try srSavedH := Integer(srSavedH)
    catch
        srSavedH := srDefaultH
    srSavedW := Max(720, Min(srMaxW, srSavedW))
    srSavedH := Max(480, Min(srMaxH, srSavedH))
    srGui.Show("Hide w" srSavedW " h" srSavedH)
    CPApplyOwnedDialogTheme(srGui)
    srGui.GetClientPos(,, &srClientW, &srClientH)
    StudyReaderResize(srState, srGui, 0, srClientW, srClientH)
    StudyLibraryRunBridge(srState, "ensure")
    StudyLibraryLoadGroup(srState, srGroupId, srVersion)
    ; Loading enables the initially disabled navigation/Anki controls. Reapply
    ; their native dark state after that transition so Windows does not repaint
    ; them with the default light button/check-box theme.
    CPApplyOwnedDialogTheme(srGui)
    ; The explicit hidden sizing pass above queued image/repaint work intended
    ; for interactive resizing. The loaded image is already laid out
    ; synchronously, so do not let those startup timers rebuild the visible UI.
    try SetTimer(srState["imageLayoutCallback"], 0)
    try SetTimer(srState["redrawCallback"], 0)

    srSavedX := IniRead(iniPath, "study_reader_view", "x", "__missing__")
    srSavedY := IniRead(iniPath, "study_reader_view", "y", "__missing__")
    if ((srSavedX is number) && (srSavedY is number)) {
        srSavedX := Integer(srSavedX), srSavedY := Integer(srSavedY)
        StudyReaderClampPosition(srGui, &srSavedX, &srSavedY)
        StudyWindowRevealFinished(srGui, "x" srSavedX " y" srSavedY)
    } else {
        StudyWindowRevealFinished(srGui, "Center")
    }
    WinActivate("ahk_id " srGui.Hwnd)
    ; Show/activation can post one final native Size notification. Wait until
    ; that startup queue has drained before enabling normal interactive resize.
    SetTimer(StudyWindowEnableLiveResize.Bind(srState), -300)
}

StudyLibraryCloseDialog(slDialog, *) {
    try slDialog.Destroy()
}

StudyLibrarySelectedGroupIds(slState) {
    slIds := []
    if !(slState.Has("list") && slState.Has("groups"))
        return slIds
    slRow := 0
    Loop {
        slRow := slState["list"].GetNext(slRow)
        if !slRow
            break
        slGroupId := StudyLibraryRowGroupId(slState, slRow)
        if (slGroupId > 0)
            slIds.Push(slGroupId)
    }
    return slIds
}

StudyLibraryRowGroupId(slState, slRow) {
    try return Integer(
        slState["list"].GetText(slRow, slState["columns"].Length + 1)
    )
    return 0
}

StudyLibraryGroupsInDisplayOrder(slState) {
    slGroups := []
    slGroupsById := Map()
    for slGroup in slState["groups"]
        slGroupsById[slGroup["id"]] := slGroup
    Loop slState["list"].GetCount() {
        slGroupId := StudyLibraryRowGroupId(slState, A_Index)
        if slGroupsById.Has(slGroupId)
            slGroups.Push(slGroupsById[slGroupId])
    }
    return slGroups
}

StudyLibraryFindGroupRow(slState, slGroupId) {
    if (slGroupId <= 0)
        return 0
    Loop slState["list"].GetCount() {
        if (StudyLibraryRowGroupId(slState, A_Index) = slGroupId)
            return A_Index
    }
    return 0
}

StudyLibraryGroupIdSelected(slIds, slGroupId) {
    for slId in slIds {
        if (slId = slGroupId)
            return true
    }
    return false
}

StudyLibraryUpdateSelectionActions(slState, *) {
    if !(slState.Has("editDetailsButton") && slState["editDetailsButton"])
        return
    slCount := StudyLibrarySelectedGroupIds(slState).Length
    slState["editDetailsButton"].Text := slCount > 1
        ? "Edit " slCount " selected..." : "Edit details..."
    slState["editDetailsButton"].Enabled := slCount > 0
    if slState.Has("studyButton")
        slState["studyButton"].Enabled := slCount = 1
}

StudyLibraryEditDetails(slState, *) {
    slIds := StudyLibrarySelectedGroupIds(slState)
    if !slIds.Length
        return
    if (slIds.Length > 1)
        return StudyLibraryOpenBulkDetails(slState, slIds)
    if (slState["currentGroupId"] != slIds[1])
        StudyLibraryLoadGroup(slState, slIds[1])
    if (slState["currentGroupId"] <= 0)
        return
    slDialog := Gui(
        "+Owner" slState["gui"].Hwnd " +OwnDialogs",
        "Study Library - Edit details"
    )
    slDialog.MarginX := 18, slDialog.MarginY := 16
    slDialog.SetFont("s10", "Segoe UI")
    slDialog.Add("Text", "xm ym w110", "Chapter / section:")
    slChapterEdit := slDialog.Add(
        "Edit", "x+10 yp-4 w360", slState["currentChapter"]
    )
    slDialog.Add("Text", "xm y+18 w110", "Speaker:")
    slSpeakerEdit := slDialog.Add(
        "Edit", "x+10 yp-4 w360", slState["currentSpeaker"]
    )
    slDialog.Add("Text", "xm y+18 w110", "Tags:")
    slTagsEdit := slDialog.Add(
        "Edit", "x+10 yp-4 w360", slState["currentTags"]
    )
    slAddedToAnki := slDialog.Add(
        "CheckBox", "xm y+18 w480", "Added to Anki"
    )
    slAddedToAnki.Value := slState["currentAddedToAnkiAt"] != "" ? 1 : 0
    slHint := slDialog.Add(
        "Text", "xm y+12 w480 h38 cGray",
        "Separate multiple tags with commas. Chapter, speaker and tags are "
        . "included in the Study Library search."
    )
    CPRegisterMutedControl(slHint)
    slSaveButton := slDialog.Add("Button", "xm y+14 w120 Default", "Save details")
    slCancelButton := slDialog.Add("Button", "x+10 yp w120", "Cancel")
    slSaveButton.OnEvent(
        "Click",
        StudyLibrarySaveDetails.Bind(
            slState, slDialog, slChapterEdit, slSpeakerEdit, slTagsEdit,
            slAddedToAnki
        )
    )
    slCancelButton.OnEvent("Click", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Close", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slDialog)
    slChapterEdit.Focus()
}

StudyLibraryBulkValueModeChanged(slModeDdl, slValueEdit, slEnabledValues, *) {
    slValueEdit.Enabled := InStr("," slEnabledValues ",", "," slModeDdl.Value ",")
}

StudyLibraryBulkOwnedMessage(slOwnerHwnd, slMessage, slIcon := "warning") {
    if !slOwnerHwnd || !DllCall("user32\IsWindow", "ptr", slOwnerHwnd, "int")
        slOwnerHwnd := DllCall("user32\GetForegroundWindow", "ptr")
    slFlags := 0x00010000 ; MB_SETFOREGROUND
    slFlags |= slIcon = "error" ? 0x00000010 : 0x00000030
    return DllCall(
        "user32\MessageBoxW", "ptr", slOwnerHwnd,
        "wstr", slMessage, "wstr", "Study Library - Bulk edit",
        "uint", slFlags, "int"
    )
}

StudyLibraryOpenBulkDetails(slState, slIds) {
    slCount := slIds.Length
    slDialog := Gui(
        "+Owner" slState["gui"].Hwnd " +OwnDialogs",
        "Study Library - Bulk edit"
    )
    slDialog.MarginX := 18, slDialog.MarginY := 16
    slDialog.SetFont("s10", "Segoe UI")
    slHeading := slDialog.Add(
        "Text", "xm ym w540", slCount " explanations selected"
    )
    slHeading.SetFont("s11 Bold")
    slHint := slDialog.Add(
        "Text", "xm y+7 w540 h42 cGray",
        "Only the changes chosen below are applied. Every field starts at "
        . "Keep existing so unrelated metadata remains untouched."
    )
    CPRegisterMutedControl(slHint)

    slDialog.Add("Text", "xm y+14 w92", "Chapter:")
    slChapterMode := slDialog.Add(
        "DropDownList", "x+8 yp-4 w145 0x210", ["Keep existing", "Set to...", "Clear"]
    )
    slChapterMode.Choose(1)
    slChapterEdit := slDialog.Add("Edit", "x+10 yp w285 Disabled")

    slDialog.Add("Text", "xm y+16 w92", "Speaker:")
    slSpeakerMode := slDialog.Add(
        "DropDownList", "x+8 yp-4 w145 0x210", ["Keep existing", "Set to...", "Clear"]
    )
    slSpeakerMode.Choose(1)
    slSpeakerEdit := slDialog.Add("Edit", "x+10 yp w285 Disabled")

    slDialog.Add("Text", "xm y+16 w92", "Tags:")
    slTagsMode := slDialog.Add(
        "DropDownList", "x+8 yp-4 w145 0x210",
        ["Keep existing", "Add tags", "Remove tags", "Replace all tags", "Clear all tags"]
    )
    slTagsMode.Choose(1)
    slTagsEdit := slDialog.Add("Edit", "x+10 yp w285 Disabled")

    slDialog.Add("Text", "xm y+16 w92", "Added to Anki:")
    slAnkiMode := slDialog.Add(
        "DropDownList", "x+8 yp-4 w210 0x210",
        ["Keep existing", "Mark as added", "Mark as not added"]
    )
    slAnkiMode.Choose(1)

    slTagHint := slDialog.Add(
        "Text", "xm y+13 w540 h38 cGray",
        "Separate multiple tags with commas. Add and Remove preserve all other "
        . "tags; Replace all tags intentionally overwrites them."
    )
    CPRegisterMutedControl(slTagHint)
    slApply := slDialog.Add(
        "Button", "xm y+12 w170 Default", "Apply to " slCount " entries"
    )
    slCancel := slDialog.Add("Button", "x+10 yp w110", "Cancel")

    slChapterMode.OnEvent(
        "Change", StudyLibraryBulkValueModeChanged.Bind(
            slChapterMode, slChapterEdit, "2"
        )
    )
    slSpeakerMode.OnEvent(
        "Change", StudyLibraryBulkValueModeChanged.Bind(
            slSpeakerMode, slSpeakerEdit, "2"
        )
    )
    slTagsMode.OnEvent(
        "Change", StudyLibraryBulkValueModeChanged.Bind(
            slTagsMode, slTagsEdit, "2,3,4"
        )
    )
    slApply.OnEvent(
        "Click", StudyLibraryApplyBulkDetails.Bind(
            slState, slDialog, slIds,
            slChapterMode, slChapterEdit, slSpeakerMode, slSpeakerEdit,
            slTagsMode, slTagsEdit, slAnkiMode
        )
    )
    slCancel.OnEvent("Click", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Close", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slDialog)
    slChapterMode.Focus()
}

StudyLibraryApplyBulkDetails(
    slState, slDialog, slIds,
    slChapterMode, slChapterEdit, slSpeakerMode, slSpeakerEdit,
    slTagsMode, slTagsEdit, slAnkiMode, *
) {
    global CPStudyReaderState
    slChapterModes := ["keep", "set", "clear"]
    slSpeakerModes := ["keep", "set", "clear"]
    slTagsModes := ["keep", "add", "remove", "replace", "clear"]
    slAnkiModes := ["keep", "1", "0"]
    slChapterOperation := slChapterModes[slChapterMode.Value]
    slSpeakerOperation := slSpeakerModes[slSpeakerMode.Value]
    slTagsOperation := slTagsModes[slTagsMode.Value]
    slAnkiOperation := slAnkiModes[slAnkiMode.Value]
    slChapterValue := Trim(slChapterEdit.Value)
    slSpeakerValue := Trim(slSpeakerEdit.Value)
    slTagsValue := Trim(slTagsEdit.Value)

    if (slChapterOperation = "keep" && slSpeakerOperation = "keep"
        && slTagsOperation = "keep" && slAnkiOperation = "keep") {
        StudyLibraryBulkOwnedMessage(
            slDialog.Hwnd, "Choose at least one metadata change first."
        )
        return
    }
    if (slChapterOperation = "set" && slChapterValue = "") {
        StudyLibraryBulkOwnedMessage(
            slDialog.Hwnd, "Enter the chapter that should be set."
        )
        slChapterEdit.Focus()
        return
    }
    if (slSpeakerOperation = "set" && slSpeakerValue = "") {
        StudyLibraryBulkOwnedMessage(
            slDialog.Hwnd, "Enter the speaker that should be set."
        )
        slSpeakerEdit.Focus()
        return
    }
    if (slTagsOperation = "add" || slTagsOperation = "remove"
        || slTagsOperation = "replace") && slTagsValue = "" {
        StudyLibraryBulkOwnedMessage(
            slDialog.Hwnd, "Enter at least one tag for the selected tag operation."
        )
        slTagsEdit.Focus()
        return
    }

    slIdText := ""
    for slId in slIds
        slIdText .= (slIdText = "" ? "" : ",") slId
    EnvSet("STUDY_LIBRARY_GROUP_IDS", slIdText)
    EnvSet("STUDY_LIBRARY_BULK_CHAPTER_MODE", slChapterOperation)
    EnvSet("STUDY_LIBRARY_BULK_CHAPTER", slChapterValue)
    EnvSet("STUDY_LIBRARY_BULK_SPEAKER_MODE", slSpeakerOperation)
    EnvSet("STUDY_LIBRARY_BULK_SPEAKER", slSpeakerValue)
    EnvSet("STUDY_LIBRARY_BULK_TAGS_MODE", slTagsOperation)
    EnvSet("STUDY_LIBRARY_BULK_TAGS", slTagsValue)
    EnvSet("STUDY_LIBRARY_BULK_ANKI_MODE", slAnkiOperation)
    slUpdated := false
    try {
        slUpdated := StudyLibraryRunBridge(slState, "bulk-metadata")
    } finally {
        for slName in [
            "STUDY_LIBRARY_GROUP_IDS",
            "STUDY_LIBRARY_BULK_CHAPTER_MODE", "STUDY_LIBRARY_BULK_CHAPTER",
            "STUDY_LIBRARY_BULK_SPEAKER_MODE", "STUDY_LIBRARY_BULK_SPEAKER",
            "STUDY_LIBRARY_BULK_TAGS_MODE", "STUDY_LIBRARY_BULK_TAGS",
            "STUDY_LIBRARY_BULK_ANKI_MODE"
        ]
            EnvSet(slName, "")
    }
    if !slUpdated
        return

    try slDialog.Destroy()
    if (CPStudyReaderState && CPStudyReaderState.Has("currentGroupId")
        && StudyLibraryGroupIdSelected(
            slIds, CPStudyReaderState["currentGroupId"]
        )) {
        slReaderGroup := CPStudyReaderState["currentGroupId"]
        slReaderVersion := CPStudyReaderState["currentVersion"]
        StudyReaderRememberEntryView(CPStudyReaderState)
        StudyLibraryLoadGroup(
            CPStudyReaderState, slReaderGroup, slReaderVersion
        )
        StudyReaderRestoreEntryView(CPStudyReaderState, slReaderGroup)
    }
    StudyLibraryRefresh(slState)
}

StudyLibrarySaveDetails(
    slState, slDialog, slChapterEdit, slSpeakerEdit, slTagsEdit,
    slAddedToAnki, *
) {
    EnvSet("STUDY_LIBRARY_CHAPTER", Trim(slChapterEdit.Value))
    EnvSet("STUDY_LIBRARY_SPEAKER", Trim(slSpeakerEdit.Value))
    EnvSet("STUDY_LIBRARY_TAGS", Trim(slTagsEdit.Value))
    EnvSet("STUDY_LIBRARY_ADDED_TO_ANKI", slAddedToAnki.Value ? "1" : "0")
    try {
        if !StudyLibraryRunBridge(
            slState, "set-metadata", slState["currentGroupId"]
        )
            return
    } finally {
        EnvSet("STUDY_LIBRARY_CHAPTER", "")
        EnvSet("STUDY_LIBRARY_SPEAKER", "")
        EnvSet("STUDY_LIBRARY_TAGS", "")
        EnvSet("STUDY_LIBRARY_ADDED_TO_ANKI", "")
    }
    try slDialog.Destroy()
    StudyLibraryRefresh(slState)
}

StudyLibraryRemoveVersion(slState, *) {
    slIndex := slState["versionDdl"].Value
    if (slState["currentGroupId"] <= 0 || slIndex < 1
        || slIndex > slState["versions"].Length)
        return
    slVersion := slState["versions"][slIndex]["version"]
    slOnlyVersion := slState["versions"].Length = 1
    slQuestion := slOnlyVersion
        ? "This is the only version. Removing it will also remove this source "
            . "from the Study Library."
        : "Remove version v" Format("{:02}", slVersion)
            . " from this saved explanation?"
    slQuestion .= "`n`nA database backup is created first, and saved screenshots "
        . "are moved into the Study Library Trash folder. Plain-text copies are "
        . "not deleted."
    slDialogTitle := slOnlyVersion ? "Remove explanation" : "Remove explanation version"
    if (MsgBox(slQuestion, slDialogTitle, "YesNo Icon! Default2") != "Yes")
        return
    if !StudyLibraryRunBridge(
        slState, "remove-version", slState["currentGroupId"], slVersion
    )
        return
    slMutationRows := StudyLibraryReadRows(slState["outputDir"] "\mutation.tsv")
    slBackupPath := ""
    if slMutationRows.Length && slMutationRows[1].Length >= 3
        slBackupPath := StudyLibraryHexDecode(slMutationRows[1][3])
    StudyLibraryRefresh(slState)
    StudyLibraryRefreshStorage(slState, true)
    slResult := slOnlyVersion
        ? "The explanation was removed from the Study Library."
        : "The selected version was removed from the Study Library."
    if (slBackupPath != "")
        slResult .= "`n`nRecovery backup:`n" slBackupPath
    MsgBox(slResult, "Study Library", "OK Iconi")
}

StudyLibraryContextMenu(
    slState, slGui, slControl, slRow, slIsRightClick, slX, slY
) {
    try slControlHwnd := slControl.Hwnd
    catch
        return
    if (slControlHwnd != slState["list"].Hwnd)
        return

    if (slRow > 0) {
        slClickedId := StudyLibraryRowGroupId(slState, slRow)
        if (slClickedId > 0) {
            slSelectedIds := StudyLibrarySelectedGroupIds(slState)
            ; Preserve an existing multi-selection when its row is right-clicked.
            ; A click on any other row starts a new single-row selection.
            if !StudyLibraryGroupIdSelected(slSelectedIds, slClickedId) {
                slControl.Modify(0, "-Select")
                slControl.Modify(slRow, "Select Focus Vis")
            } else
                slControl.Modify(slRow, "Focus Vis")
            if (slState["currentGroupId"] != slClickedId)
                StudyLibraryLoadGroup(slState, slClickedId)
            StudyLibraryUpdateSelectionActions(slState)
        }
    }

    slSelectedIds := StudyLibrarySelectedGroupIds(slState)
    slSelectedCount := slSelectedIds.Length
    slMenuItems := []
    if (slSelectedCount > 0) {
        if (slSelectedCount = 1) {
            slMenuItems.Push(Map(
                "label", "Open in Reader...",
                "action", StudyLibraryOpenSelectedReader.Bind(slState)
            ))
            slMenuItems.Push(Map(
                "label", "Add selected explanation to Anki...",
                "action", StudyLibraryAddSelectedToAnki.Bind(slState)
            ))
        }
        slEditLabel := slSelectedCount > 1
            ? "Edit " slSelectedCount " selected..." : "Edit details..."
        slMenuItems.Push(Map(
            "label", slEditLabel,
            "action", StudyLibraryEditDetails.Bind(slState)
        ))
        if (slSelectedCount = 1 && slState["versions"].Length > 0) {
            slRemoveLabel := slState["versions"].Length = 1
                ? "Remove explanation..." : "Remove version..."
            slMenuItems.Push(Map(
                "label", slRemoveLabel,
                "action", StudyLibraryRemoveVersion.Bind(slState)
            ))
        }
        slMenuItems.Push(Map("separator", true))
    }
    slMenuItems.Push(Map(
        "label", "Filters...",
        "action", StudyLibraryOpenFilters.Bind(slState)
    ))
    slMenuItems.Push(Map(
        "label", "Columns...",
        "action", StudyLibraryOpenColumns.Bind(slState)
    ))
    slChoice := CPThemedContextPopup(
        slGui.Hwnd, slX, slY, slMenuItems
    )
    if (slChoice >= 1 && slChoice <= slMenuItems.Length
        && slMenuItems[slChoice].Has("action"))
        slMenuItems[slChoice]["action"].Call()
}

StudyLibraryValueChoices(
    slPath, slAllLabel, slEmptyLabel, slCurrentMode := "all", slCurrentValue := ""
) {
    slLabels := [slAllLabel, slEmptyLabel]
    slChoices := [
        Map("mode", "all", "value", ""),
        Map("mode", "empty", "value", "")
    ]
    slCurrentFound := (slCurrentMode != "value")
    for slRow in StudyLibraryReadRows(slPath) {
        if !slRow.Length
            continue
        slValue := StudyLibraryHexDecode(slRow[1])
        if (slValue = "")
            continue
        slLabels.Push(slValue)
        slChoices.Push(Map("mode", "value", "value", slValue))
        if (slCurrentMode = "value" && slValue = slCurrentValue)
            slCurrentFound := true
    }
    ; Keep an active value available even if another filter currently excludes
    ; every matching row or the value was removed in another process.
    if !slCurrentFound && slCurrentValue != "" {
        slLabels.Push(slCurrentValue)
        slChoices.Push(Map("mode", "value", "value", slCurrentValue))
    }
    return Map("labels", slLabels, "choices", slChoices)
}

StudyLibraryChoiceIndex(slChoices, slMode, slValue := "") {
    for slIndex, slChoice in slChoices {
        if (slChoice["mode"] = slMode && slChoice["value"] = slValue)
            return slIndex
    }
    return 1
}

StudyLibraryFilterDisplay(slMode, slValue, slAny := "Any", slEmpty := "Not set") {
    if (slMode = "empty")
        return slEmpty
    if (slMode = "value" && slValue != "")
        return slValue
    return slAny
}

StudyLibraryDateModeIndex(slMode) {
    if (slMode = "today")
        return 2
    if (slMode = "yesterday")
        return 3
    if (slMode = "last24")
        return 4
    if (slMode = "last7")
        return 5
    if (slMode = "custom")
        return 6
    return 1
}

StudyLibraryDateModeFromIndex(slIndex) {
    return slIndex = 2 ? "today"
        : (slIndex = 3 ? "yesterday"
        : (slIndex = 4 ? "last24"
        : (slIndex = 5 ? "last7"
        : (slIndex = 6 ? "custom" : "all"))))
}

StudyLibraryCompactDateTime(slStamp) {
    if (StrLen(slStamp) < 12)
        return ""
    return SubStr(slStamp, 1, 4) "-" SubStr(slStamp, 5, 2) "-"
        . SubStr(slStamp, 7, 2) " " SubStr(slStamp, 9, 2) ":"
        . SubStr(slStamp, 11, 2)
}

StudyLibraryBridgeDateTime(slStamp) {
    if (StrLen(slStamp) < 14)
        return ""
    return SubStr(slStamp, 1, 4) "-" SubStr(slStamp, 5, 2) "-"
        . SubStr(slStamp, 7, 2) " " SubStr(slStamp, 9, 2) ":"
        . SubStr(slStamp, 11, 2) ":" SubStr(slStamp, 13, 2)
}

StudyLibraryDateFilterDisplay(slState) {
    slMode := slState["dateMode"]
    if (slMode = "today")
        return "Today"
    if (slMode = "yesterday")
        return "Yesterday"
    if (slMode = "last24")
        return "Last 24 hours"
    if (slMode = "last7")
        return "Last 7 days"
    if (slMode = "custom")
        return StudyLibraryCompactDateTime(slState["dateFrom"])
            . " – " StudyLibraryCompactDateTime(slState["dateTo"])
    return "Any time"
}

StudyLibraryDateControlsChanged(
    slDateDdl, slFromDate, slFromTime, slToDate, slToTime, *
) {
    slEnabled := slDateDdl.Value = 6
    for slControl in [slFromDate, slFromTime, slToDate, slToTime]
        slControl.Enabled := slEnabled
}

StudyLibraryDialogFilterCount(slState) {
    slCount := 0
    for slMode in [
        slState["profileMode"], slState["chapterMode"], slState["speakerMode"],
        slState["tagMode"], slState["ankiMode"], slState["dateMode"]
    ] {
        if (slMode != "all")
            slCount += 1
    }
    return slCount
}

StudyLibraryColumnFilterActive(slState, slColumnKey) {
    if (slColumnKey = "updated")
        return slState["dateMode"] != "all"
    if (slColumnKey = "profile")
        return slState["profileMode"] != "all"
    if (slColumnKey = "chapter")
        return slState["chapterMode"] != "all"
    if (slColumnKey = "speaker")
        return slState["speakerMode"] != "all"
    if (slColumnKey = "tags")
        return slState["tagMode"] != "all"
    if (slColumnKey = "anki")
        return slState["ankiMode"] != "all"
    return false
}

StudyLibraryUpdateFilterIndicators(slState) {
    slCount := StudyLibraryDialogFilterCount(slState)
    slState["filterButton"].Text := slCount
        ? "Filters (" slCount ")..." : "Filters..."
    StudyLibraryApplyColumns(slState)
}

StudyLibraryApplyFilters(
    slState, slDialog, slProfileDdl, slChapterDdl, slSpeakerDdl, slTagDdl, slAnkiDdl,
    slDateDdl,
    slFromDate, slFromTime, slToDate, slToTime, *
) {
    slDateMode := StudyLibraryDateModeFromIndex(slDateDdl.Value)
    ; The picker displays minutes, so include the complete selected final minute.
    slFromStamp := SubStr(slFromDate.Value, 1, 8) SubStr(slFromTime.Value, 9, 4) "00"
    slToStamp := SubStr(slToDate.Value, 1, 8) SubStr(slToTime.Value, 9, 4) "59"
    if (slDateMode = "custom" && Integer(slToStamp) < Integer(slFromStamp)) {
        MsgBox(
            "The end of the generated-date range must not be before its start.",
            "Study Library - Filters",
            "OK Icon!"
        )
        return
    }
    slProfileIndex := slProfileDdl.Value
    slChapterIndex := slChapterDdl.Value
    slSpeakerIndex := slSpeakerDdl.Value
    slTagIndex := slTagDdl.Value
    if (slProfileIndex >= 1 && slProfileIndex <= slState["profileChoices"].Length) {
        slChoice := slState["profileChoices"][slProfileIndex]
        slState["profileMode"] := slChoice["mode"]
        slState["profileFilter"] := slChoice["value"]
    }
    if (slChapterIndex >= 1 && slChapterIndex <= slState["chapterChoices"].Length) {
        slChoice := slState["chapterChoices"][slChapterIndex]
        slState["chapterMode"] := slChoice["mode"]
        slState["chapterFilter"] := slChoice["value"]
    }
    if (slSpeakerIndex >= 1 && slSpeakerIndex <= slState["speakerChoices"].Length) {
        slChoice := slState["speakerChoices"][slSpeakerIndex]
        slState["speakerMode"] := slChoice["mode"]
        slState["speakerFilter"] := slChoice["value"]
    }
    if (slTagIndex >= 1 && slTagIndex <= slState["tagChoices"].Length) {
        slChoice := slState["tagChoices"][slTagIndex]
        slState["tagMode"] := slChoice["mode"]
        slState["tagFilter"] := slChoice["value"]
    }
    slAnkiIndex := slAnkiDdl.Value
    slState["ankiMode"] := slAnkiIndex = 2
        ? "found" : (slAnkiIndex = 3 ? "not-found"
            : (slAnkiIndex = 4 ? "not-checked" : "all"))
    slState["dateMode"] := slDateMode
    if (slDateMode = "custom") {
        slState["dateFrom"] := slFromStamp
        slState["dateTo"] := slToStamp
    }
    try slDialog.Destroy()
    StudyLibraryRefresh(slState)
}

StudyLibraryOpenFilters(slState, *) {
    slDialog := Gui(
        "+Owner" slState["gui"].Hwnd " +OwnDialogs",
        "Study Library - Filters"
    )
    slDialog.MarginX := 18, slDialog.MarginY := 16
    slDialog.SetFont("s10", "Segoe UI")
    slDialog.Add("Text", "xm ym w90", "Profile:")
    slProfileDdl := slDialog.Add(
        "DropDownList", "x+10 yp-4 w310 0x210", slState["profileLabels"]
    )
    slProfileDdl.Choose(StudyLibraryChoiceIndex(
        slState["profileChoices"], slState["profileMode"], slState["profileFilter"]
    ))
    slDialog.Add("Text", "xm y+18 w90", "Chapter:")
    slChapterDdl := slDialog.Add(
        "DropDownList", "x+10 yp-4 w310 0x210", slState["chapterLabels"]
    )
    slChapterDdl.Choose(StudyLibraryChoiceIndex(
        slState["chapterChoices"], slState["chapterMode"], slState["chapterFilter"]
    ))
    slDialog.Add("Text", "xm y+18 w90", "Speaker:")
    slSpeakerDdl := slDialog.Add(
        "DropDownList", "x+10 yp-4 w310 0x210", slState["speakerLabels"]
    )
    slSpeakerDdl.Choose(StudyLibraryChoiceIndex(
        slState["speakerChoices"], slState["speakerMode"], slState["speakerFilter"]
    ))
    slDialog.Add("Text", "xm y+18 w90", "Tag:")
    slTagDdl := slDialog.Add(
        "DropDownList", "x+10 yp-4 w310 0x210", slState["tagLabels"]
    )
    slTagDdl.Choose(StudyLibraryChoiceIndex(
        slState["tagChoices"], slState["tagMode"], slState["tagFilter"]
    ))
    slDialog.Add("Text", "xm y+18 w90", "Anki status:")
    slAnkiDdl := slDialog.Add(
        "DropDownList", "x+10 yp-4 w310 0x210",
        ["All", "Found in Anki", "Not found", "Not checked"]
    )
    slAnkiDdl.Choose(
        slState["ankiMode"] = "found" ? 2
            : (slState["ankiMode"] = "not-found" ? 3
                : (slState["ankiMode"] = "not-checked" ? 4 : 1))
    )
    slDialog.Add("Text", "xm y+18 w90", "Date generated:")
    slDateDdl := slDialog.Add(
        "DropDownList", "x+10 yp-4 w310 0x210",
        ["Any time", "Today", "Yesterday", "Last 24 hours", "Last 7 days",
            "Custom range..."]
    )
    slDateDdl.Choose(StudyLibraryDateModeIndex(slState["dateMode"]))
    slDialog.Add("Text", "xm y+16 w90", "From:")
    slFromDate := slDialog.Add(
        "DateTime", "x+10 yp-4 w140 Choose" slState["dateFrom"], "yyyy-MM-dd"
    )
    slFromTime := slDialog.Add(
        "DateTime", "x+8 yp w90 Choose" slState["dateFrom"], "HH:mm"
    )
    slDialog.Add("Text", "xm y+16 w90", "To:")
    slToDate := slDialog.Add(
        "DateTime", "x+10 yp-4 w140 Choose" slState["dateTo"], "yyyy-MM-dd"
    )
    slToTime := slDialog.Add(
        "DateTime", "x+8 yp w90 Choose" slState["dateTo"], "HH:mm"
    )
    slDateDdl.OnEvent("Change", StudyLibraryDateControlsChanged.Bind(
        slDateDdl, slFromDate, slFromTime, slToDate, slToTime
    ))
    StudyLibraryDateControlsChanged(
        slDateDdl, slFromDate, slFromTime, slToDate, slToTime
    )
    slHint := slDialog.Add(
        "Text", "xm y+12 w410 h38 cGray",
        "These filters work together with the search field."
    )
    CPRegisterMutedControl(slHint)
    slApply := slDialog.Add("Button", "xm y+12 w120 Default", "Apply filters")
    slClear := slDialog.Add("Button", "x+10 yp w100", "Clear all")
    slCancel := slDialog.Add("Button", "x+10 yp w100", "Cancel")
    slApply.OnEvent("Click", StudyLibraryApplyFilters.Bind(
        slState, slDialog, slProfileDdl, slChapterDdl, slSpeakerDdl, slTagDdl, slAnkiDdl,
        slDateDdl, slFromDate, slFromTime, slToDate, slToTime
    ))
    slClear.OnEvent("Click", StudyLibraryClearFiltersAndClose.Bind(slState, slDialog))
    slCancel.OnEvent("Click", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Close", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slDialog)
}

StudyLibraryClearFilters(slState, *) {
    slState["profileMode"] := "all"
    slState["profileFilter"] := ""
    slState["chapterMode"] := "all"
    slState["chapterFilter"] := ""
    slState["speakerMode"] := "all"
    slState["speakerFilter"] := ""
    slState["tagMode"] := "all"
    slState["tagFilter"] := ""
    slState["ankiMode"] := "all"
    slState["dateMode"] := "all"
    slState["search"].Value := ""
    StudyLibraryRefresh(slState)
}

StudyLibraryClearFiltersAndClose(slState, slDialog, *) {
    try slDialog.Destroy()
    StudyLibraryClearFilters(slState)
}

CPListViewReadColumnOrder(cpList, cpColumnCount) {
    cpOrder := []
    if !(cpList && cpList.Hwnd && cpColumnCount > 0)
        return cpOrder
    cpOrderBuffer := Buffer(cpColumnCount * 4, 0)
    try cpSucceeded := DllCall(
        "user32\SendMessageW", "ptr", cpList.Hwnd,
        "uint", 0x103B, ; LVM_GETCOLUMNORDERARRAY
        "uptr", cpColumnCount, "ptr", cpOrderBuffer.Ptr, "ptr"
    )
    catch
        return cpOrder
    if !cpSucceeded
        return cpOrder
    Loop cpColumnCount
        cpOrder.Push(
            NumGet(cpOrderBuffer, (A_Index - 1) * 4, "Int") + 1
        )
    return cpOrder
}

CPListViewNormalizeColumnOrder(cpOrder, cpColumnCount, cpFixedLast := 0) {
    cpNormalized := []
    cpSeen := Map()
    if IsObject(cpOrder) {
        for cpColumn in cpOrder {
            try cpColumn := Integer(cpColumn)
            catch
                continue
            if (cpColumn < 1 || cpColumn > cpColumnCount
                || cpColumn = cpFixedLast || cpSeen.Has(cpColumn))
                continue
            cpNormalized.Push(cpColumn)
            cpSeen[cpColumn] := true
        }
    }
    Loop cpColumnCount {
        cpColumn := A_Index
        if (cpColumn = cpFixedLast || cpSeen.Has(cpColumn))
            continue
        cpNormalized.Push(cpColumn)
        cpSeen[cpColumn] := true
    }
    if (cpFixedLast >= 1 && cpFixedLast <= cpColumnCount)
        cpNormalized.Push(cpFixedLast)
    return cpNormalized
}

CPListViewParseColumnOrder(cpRaw, cpColumnCount, cpFixedLast := 0) {
    cpParsed := []
    Loop Parse String(cpRaw), "," {
        try cpParsed.Push(Integer(Trim(A_LoopField)))
    }
    return CPListViewNormalizeColumnOrder(
        cpParsed, cpColumnCount, cpFixedLast
    )
}

CPListViewSetColumnOrder(cpList, cpOrder) {
    if !(cpList && cpList.Hwnd && IsObject(cpOrder) && cpOrder.Length)
        return false
    cpOrderBuffer := Buffer(cpOrder.Length * 4, 0)
    for cpIndex, cpColumn in cpOrder
        NumPut("Int", cpColumn - 1, cpOrderBuffer, (cpIndex - 1) * 4)
    try return !!DllCall(
        "user32\SendMessageW", "ptr", cpList.Hwnd,
        "uint", 0x103A, ; LVM_SETCOLUMNORDERARRAY
        "uptr", cpOrder.Length, "ptr", cpOrderBuffer.Ptr, "ptr"
    )
    catch
        return false
}

CPListViewRestoreColumnOrder(
    cpList, cpSection, cpKey, cpColumnCount, cpFixedLast := 0
) {
    global iniPath
    cpRaw := IniRead(iniPath, cpSection, cpKey, "")
    cpOrder := CPListViewParseColumnOrder(
        cpRaw, cpColumnCount, cpFixedLast
    )
    return CPListViewSetColumnOrder(cpList, cpOrder)
}

CPListViewSaveColumnOrder(
    cpList, cpSection, cpKey, cpColumnCount, cpFixedLast := 0
) {
    global iniPath
    cpOrder := CPListViewReadColumnOrder(cpList, cpColumnCount)
    if !cpOrder.Length
        return false
    cpOrder := CPListViewNormalizeColumnOrder(
        cpOrder, cpColumnCount, cpFixedLast
    )
    cpSerialized := ""
    for cpColumn in cpOrder {
        if (cpColumn = cpFixedLast)
            continue
        cpSerialized .= (cpSerialized = "" ? "" : ",") cpColumn
    }
    IniWrite(cpSerialized, iniPath, cpSection, cpKey)
    return true
}

StudyLibraryCreateColumns() {
    global iniPath
    slVisibleValue := StrLower(IniRead(
        iniPath, "study_library_view", "visibleColumns",
        "updated,profile,chapter,speaker,tags,source,grammar,versions,anki"
    ))
    ; Introduce the new overview column once on upgrades, while preserving any
    ; later choice to hide it through Columns...
    if !Integer(IniRead(
        iniPath, "study_library_view", "keyGrammarColumnIntroduced", 0
    )) {
        if !InStr("," slVisibleValue ",", ",grammar,")
            slVisibleValue .= (slVisibleValue = "" ? "" : ",") "grammar"
        IniWrite(1, iniPath, "study_library_view", "keyGrammarColumnIntroduced")
    }
    slVisibleRaw := "," slVisibleValue ","
    slDefinitions := [
        ["updated", "Date generated", 125, false],
        ["profile", "Profile", 110, false],
        ["chapter", "Chapter", 120, false],
        ["speaker", "Speaker", 100, false],
        ["tags", "Tags", 150, false],
        ["source", "Japanese source", 260, true],
        ["grammar", "Key grammar", 190, false],
        ["versions", "Versions", 72, false],
        ["anki", "Anki status", 120, false]
    ]
    slWidthUnits := IniRead(
        iniPath, "study_library_view", "columnWidthUnits", ""
    )
    slUseSavedWidths := slWidthUnits = "logical-v1"
    slColumns := []
    for slIndex, slDefinition in slDefinitions {
        slKey := slDefinition[1]
        slDefaultWidth := slDefinition[3]
        slWidth := slUseSavedWidths
            ? Integer(IniRead(
                iniPath, "study_library_view", "columnWidth_" slKey, slDefaultWidth
            ))
            : slDefaultWidth
        slRequired := slDefinition[4]
        slColumns.Push(Map(
            "index", slIndex,
            "key", slKey,
            "label", slDefinition[2],
            "width", Max(40, Min(600, slWidth)),
            "required", slRequired,
            "visible", slRequired || InStr(slVisibleRaw, "," slKey ",")
        ))
    }
    return slColumns
}

StudyLibraryCaptureColumnWidths(slState, slPersist := false) {
    global iniPath
    for slColumn in slState["columns"] {
        if !slColumn["visible"]
            continue
        try slPhysicalWidth := DllCall(
            "user32\SendMessageW",
            "ptr", slState["list"].Hwnd,
            "uint", 0x101D, ; LVM_GETCOLUMNWIDTH
            "ptr", slColumn["index"] - 1,
            "ptr", 0,
            "ptr"
        )
        catch
            continue
        slDpi := GetWindowDPI(slState["gui"].Hwnd)
        slWidth := Round(slPhysicalWidth * 96 / Max(96, slDpi))
        if (slWidth >= 40)
            slColumn["width"] := slWidth
        if slPersist
            IniWrite(
                slColumn["width"], iniPath, "study_library_view",
                "columnWidth_" slColumn["key"]
            )
    }
    if slPersist
        IniWrite(
            "logical-v1", iniPath, "study_library_view", "columnWidthUnits"
        )
}

StudyLibraryApplyColumns(slState) {
    for slColumn in slState["columns"] {
        slWidth := slColumn["visible"] ? slColumn["width"] : 0
        slTitle := slColumn["label"]
        if StudyLibraryColumnFilterActive(slState, slColumn["key"])
            slTitle .= "  ▼"
        slState["list"].ModifyCol(
            slColumn["index"], slWidth, slTitle
        )
    }
    StudyLibraryEnsureInternalColumnHidden(slState)
    StudyLibraryApplyHeaderIndicators(slState)
}

StudyLibraryEnsureInternalColumnHidden(slState, *) {
    if !StudyLibraryStateAlive(slState)
        return
    slInternalIndex := slState["columns"].Length
    slState["lockingInternalColumn"] := true
    try DllCall(
        "user32\SendMessageW", "ptr", slState["list"].Hwnd,
        "uint", 0x101E, ; LVM_SETCOLUMNWIDTH
        "ptr", slInternalIndex, "ptr", 0, "ptr"
    )
    finally slState["lockingInternalColumn"] := false
}

StudyLibraryApplyHeaderIndicators(slState) {
    if !StudyLibraryStateAlive(slState)
        return
    try {
        if !slState["headerHwnd"]
            slState["headerHwnd"] := DllCall(
                "user32\SendMessageW", "ptr", slState["list"].Hwnd,
                "uint", 0x101F, "ptr", 0, "ptr", 0, "ptr"
            ) ; LVM_GETHEADER
        if (!slState["headerHwnd"])
            return
        slItemSize := A_PtrSize = 8 ? 72 : 48
        slFormatOffset := 12 + 2 * A_PtrSize
        for slColumn in slState["columns"] {
            slItem := Buffer(slItemSize, 0)
            NumPut("uint", 0x00000004, slItem, 0) ; HDI_FORMAT
            if !DllCall(
                "user32\SendMessageW", "ptr", slState["headerHwnd"],
                "uint", 0x1203, "ptr", slColumn["index"] - 1,
                "ptr", slItem, "ptr"
            ) ; HDM_GETITEMA (format/image fields are encoding-independent)
                continue
            slFormat := NumGet(slItem, slFormatOffset, "int")
            slFormat &= ~(0x00000800 | 0x00000400 | 0x00000200)
            if (slColumn["index"] = slState["sortColumn"])
                slFormat |= slState["sortDirection"] = "asc"
                    ? 0x00000400 : 0x00000200 ; HDF_SORTUP / HDF_SORTDOWN
            NumPut("uint", 0x00000004, slItem, 0) ; HDI_FORMAT
            NumPut("int", slFormat, slItem, slFormatOffset)
            DllCall(
                "user32\SendMessageW", "ptr", slState["headerHwnd"],
                "uint", 0x1204, "ptr", slColumn["index"] - 1,
                "ptr", slItem, "ptr"
            ) ; HDM_SETITEMA
        }
        ; The final column stores the database group ID used to keep row
        ; selection correct after sorting. It is implementation data, not UI.
        ; HDF_FIXEDWIDTH prevents the zero-width divider from being dragged.
        slInternalIndex := slState["columns"].Length
        slItem := Buffer(slItemSize, 0)
        NumPut("uint", 0x00000004, slItem, 0) ; HDI_FORMAT
        if DllCall(
            "user32\SendMessageW", "ptr", slState["headerHwnd"],
            "uint", 0x1203, "ptr", slInternalIndex,
            "ptr", slItem, "ptr"
        ) {
            slFormat := NumGet(slItem, slFormatOffset, "int")
            slFormat |= 0x00000100 ; HDF_FIXEDWIDTH
            slFormat &= ~(0x00000800 | 0x00000400 | 0x00000200)
            NumPut("uint", 0x00000004, slItem, 0)
            NumPut("int", slFormat, slItem, slFormatOffset)
            DllCall(
                "user32\SendMessageW", "ptr", slState["headerHwnd"],
                "uint", 0x1204, "ptr", slInternalIndex,
                "ptr", slItem, "ptr"
            )
        }
        StudyLibraryEnsureInternalColumnHidden(slState)
    }
}

StudyLibraryRedrawList(slState, *) {
    if !StudyLibraryStateAlive(slState)
        return
    try DllCall(
        "user32\RedrawWindow", "ptr", slState["list"].Hwnd,
        "ptr", 0, "ptr", 0,
        "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100
    ) ; RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW
}

StudyLibraryAutoSizeColumn(slState, slColumnIndex, *) {
    if !StudyLibraryStateAlive(slState)
        return
    if (slColumnIndex < 0 || slColumnIndex >= slState["columns"].Length)
        return
    try {
        DllCall(
            "user32\SendMessageW", "ptr", slState["list"].Hwnd,
            "uint", 0x101E, ; LVM_SETCOLUMNWIDTH
            "ptr", slColumnIndex,
            "ptr", -2, ; LVSCW_AUTOSIZE_USEHEADER: header or cells, whichever is wider
            "ptr"
        )
        StudyLibraryCaptureColumnWidths(slState)
        StudyLibraryRedrawList(slState)
    }
}

StudyLibraryHeaderNotify(slState, wParam, lParam, slMsg, slHwnd) {
    if (!StudyLibraryStateAlive(slState) || !lParam)
        return
    try {
        if !slState["headerHwnd"]
            slState["headerHwnd"] := DllCall(
                "user32\SendMessageW", "ptr", slState["list"].Hwnd,
                "uint", 0x101F, "ptr", 0, "ptr", 0, "ptr"
            ) ; LVM_GETHEADER
        if (NumGet(lParam, 0, "Ptr") != slState["headerHwnd"])
            return
        slCode := NumGet(lParam, A_PtrSize * 2, "Int")
        slColumnIndex := -1
        if (slCode = -300 || slCode = -320 || slCode = -301 || slCode = -321
            || slCode = -305 || slCode = -325 || slCode = -306 || slCode = -326
            || slCode = -307 || slCode = -327 || slCode = -308 || slCode = -328) {
            slItemOffset := A_PtrSize = 8 ? 24 : 12
            slColumnIndex := NumGet(lParam, slItemOffset, "Int")
        }
        slInternalIndex := slState["columns"].Length
        if (slColumnIndex = slInternalIndex
            && !slState["lockingInternalColumn"]) {
            ; Reject every interactive path that could reveal the ID column.
            if (slCode = -300 || slCode = -320  ; HDN_ITEMCHANGING A/W
                || slCode = -305 || slCode = -325 ; HDN_DIVIDERDBLCLICK A/W
                || slCode = -306 || slCode = -326 ; HDN_BEGINTRACK A/W
                || slCode = -308 || slCode = -328) { ; HDN_TRACK A/W
                SetTimer(
                    StudyLibraryEnsureInternalColumnHidden.Bind(slState), -1
                )
                return true
            }
            if (slCode = -301 || slCode = -321
                || slCode = -307 || slCode = -327) {
                SetTimer(
                    StudyLibraryEnsureInternalColumnHidden.Bind(slState), -1
                )
                return true
            }
        }
        ; Header tracking does not reliably erase old vertical grid lines on a
        ; dark native ListView. Mark the table dirty while tracking, then force
        ; one complete repaint after Windows commits the new column width.
        if (slCode = -308 || slCode = -328
            || slCode = -301 || slCode = -321) { ; HDN_TRACK/ITEMCHANGED A/W
            DllCall(
                "user32\InvalidateRect", "ptr", slState["list"].Hwnd,
                "ptr", 0, "int", true
            )
        }
        if (slCode = -307 || slCode = -327) ; HDN_ENDTRACK A/W
            SetTimer(slState["listRedrawCallback"], -1)
        if (slCode = -305 || slCode = -325) { ; HDN_DIVIDERDBLCLICK A/W
            ; Run after the native double-click handler. Windows first applies
            ; content-only auto sizing; this replaces it with the documented
            ; header-or-content mode.
            SetTimer(
                StudyLibraryAutoSizeColumn.Bind(slState, slColumnIndex), -1
            )
        }
    }
}

StudyLibraryApplyCurrentSort(slState) {
    if !StudyLibraryStateAlive(slState)
        return
    try slState["list"].ModifyCol(
        slState["sortColumn"],
        slState["sortDirection"] = "asc" ? "Sort" : "SortDesc"
    )
    StudyLibraryApplyHeaderIndicators(slState)
}

StudyLibraryColumnClicked(slState, slList, slColumn) {
    if (!StudyLibraryStateAlive(slState) || slState["refreshing"])
        return
    if (slState["sortColumn"] = slColumn)
        slState["sortDirection"] := slState["sortDirection"] = "asc"
            ? "desc" : "asc"
    else {
        slState["sortColumn"] := slColumn
        slState["sortDirection"] := "asc"
    }
    ; Run after the native column-click notification has completed so the
    ; stored direction and the visible arrow always match the final row order.
    SetTimer(StudyLibraryApplyCurrentSort.Bind(slState), -1)
}

StudyLibrarySaveColumns(slState, slDialog, slChecks, *) {
    global iniPath
    StudyLibraryCaptureColumnWidths(slState)
    slVisible := ""
    for slIndex, slColumn in slState["columns"] {
        if !slColumn["required"]
            slColumn["visible"] := slChecks[slIndex].Value ? true : false
        if slColumn["visible"]
            slVisible .= (slVisible = "" ? "" : ",") slColumn["key"]
    }
    IniWrite(slVisible, iniPath, "study_library_view", "visibleColumns")
    StudyLibraryApplyColumns(slState)
    try slDialog.Destroy()
}

StudyLibraryOpenColumns(slState, *) {
    StudyLibraryCaptureColumnWidths(slState)
    slDialog := Gui(
        "+Owner" slState["gui"].Hwnd " +OwnDialogs",
        "Study Library - Columns"
    )
    slDialog.MarginX := 18, slDialog.MarginY := 16
    slDialog.SetFont("s10", "Segoe UI")
    slDialog.Add("Text", "xm ym w320", "Choose which columns are shown:")
    slChecks := []
    for slColumn in slState["columns"] {
        slCheck := slDialog.Add("CheckBox", "xm y+9 w300", slColumn["label"])
        slCheck.Value := slColumn["visible"] ? 1 : 0
        if slColumn["required"]
            slCheck.Enabled := false
        slChecks.Push(slCheck)
    }
    slHint := slDialog.Add(
        "Text", "xm y+12 w320 h38 cGray",
        "Column widths are saved when the Study Library closes. Japanese source remains visible."
    )
    CPRegisterMutedControl(slHint)
    slApply := slDialog.Add("Button", "xm y+12 w110 Default", "Apply")
    slCancel := slDialog.Add("Button", "x+10 yp w100", "Cancel")
    slApply.OnEvent("Click", StudyLibrarySaveColumns.Bind(
        slState, slDialog, slChecks
    ))
    slCancel.OnEvent("Click", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Escape", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.OnEvent("Close", StudyLibraryCloseDialog.Bind(slDialog))
    slDialog.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(slDialog)
}

StudyLibraryQueueImageLayout(slState, slDelay := 24) {
    if !StudyLibraryStateAlive(slState)
        return
    if (slState.Has("imageLayoutQueued") && slState["imageLayoutQueued"])
        return
    slState["imageLayoutQueued"] := true
    SetTimer(slState["imageLayoutCallback"], -slDelay)
}

StudyLibraryRunQueuedImageLayout(slState, *) {
    if !IsObject(slState)
        return
    slState["imageLayoutQueued"] := false
    if StudyLibraryStateAlive(slState)
        StudyLibraryShowImage(slState)
}

StudyLibraryStateAlive(slState) {
    if !(IsObject(slState) && slState.Has("gui"))
        return false
    if (slState.Has("closed") && slState["closed"])
        return false
    try {
        slHwnd := slState["gui"].Hwnd
        return slHwnd && DllCall("user32\IsWindow", "ptr", slHwnd, "int")
    }
    return false
}

StudyLibraryRefresh(slState, *) {
    if !StudyLibraryStateAlive(slState)
        return
    if (slState.Has("refreshing") && slState["refreshing"])
        return
    slState["refreshing"] := true
    try {
    slPreviousGroup := slState["currentGroupId"]
    EnvSet("STUDY_LIBRARY_QUERY", Trim(slState["search"].Value))
    EnvSet("STUDY_LIBRARY_PROFILE_MODE", slState["profileMode"])
    EnvSet("STUDY_LIBRARY_PROFILE_FILTER", slState["profileFilter"])
    EnvSet("STUDY_LIBRARY_CHAPTER_MODE", slState["chapterMode"])
    EnvSet("STUDY_LIBRARY_CHAPTER_FILTER", slState["chapterFilter"])
    EnvSet("STUDY_LIBRARY_SPEAKER_MODE", slState["speakerMode"])
    EnvSet("STUDY_LIBRARY_SPEAKER_FILTER", slState["speakerFilter"])
    EnvSet("STUDY_LIBRARY_TAG_MODE", slState["tagMode"])
    EnvSet("STUDY_LIBRARY_TAG_FILTER", slState["tagFilter"])
    EnvSet("STUDY_LIBRARY_ANKI_MODE", slState["ankiMode"])
    EnvSet("STUDY_LIBRARY_DATE_MODE", slState["dateMode"])
    EnvSet("STUDY_LIBRARY_DATE_FROM", StudyLibraryBridgeDateTime(slState["dateFrom"]))
    EnvSet("STUDY_LIBRARY_DATE_TO", StudyLibraryBridgeDateTime(slState["dateTo"]))
    if !StudyLibraryRunBridge(slState, "snapshot")
        return
    if !StudyLibraryStateAlive(slState)
        return

    slOutputDir := slState["outputDir"]
    slState["suspend"] := true
    slProfileLabels := ["All Profiles", "Unsorted (no Profile)"]
    slProfileChoices := [
        Map("mode", "all", "value", ""),
        Map("mode", "unsorted", "value", "")
    ]
    for slProfileRow in StudyLibraryReadRows(slOutputDir "\profiles.tsv") {
        if !slProfileRow.Length
            continue
        slProfileName := StudyLibraryHexDecode(slProfileRow[1])
        if (slProfileName = "")
            continue
        slProfileLabels.Push(slProfileName)
        slProfileChoices.Push(Map("mode", "profile", "value", slProfileName))
    }
    slState["profileLabels"] := slProfileLabels
    slState["profileChoices"] := slProfileChoices

    slChapterData := StudyLibraryValueChoices(
        slOutputDir "\chapters.tsv", "Any chapter", "Chapter not set",
        slState["chapterMode"], slState["chapterFilter"]
    )
    slSpeakerData := StudyLibraryValueChoices(
        slOutputDir "\speakers.tsv", "Any speaker", "Speaker not set",
        slState["speakerMode"], slState["speakerFilter"]
    )
    slTagData := StudyLibraryValueChoices(
        slOutputDir "\tags.tsv", "Any tag", "No tags",
        slState["tagMode"], slState["tagFilter"]
    )
    slState["chapterLabels"] := slChapterData["labels"]
    slState["chapterChoices"] := slChapterData["choices"]
    slState["speakerLabels"] := slSpeakerData["labels"]
    slState["speakerChoices"] := slSpeakerData["choices"]
    slState["tagLabels"] := slTagData["labels"]
    slState["tagChoices"] := slTagData["choices"]
    StudyLibraryUpdateFilterIndicators(slState)

    slState["groups"] := []
    slState["list"].Delete()
    slTotalVersions := 0
    slNotAddedCount := 0
    slSelectedRow := 0
    for slGroupRow in StudyLibraryReadRows(slOutputDir "\groups.tsv") {
        if (slGroupRow.Length < 6)
            continue
        slGroupId := Integer(slGroupRow[1])
        slGroupProfile := StudyLibraryHexDecode(slGroupRow[2])
        slUpdated := StudyLibraryHexDecode(slGroupRow[3])
        slVersionCount := Integer(slGroupRow[4])
        slPreview := StudyLibraryHexDecode(slGroupRow[5])
        slChapter := slGroupRow.Length >= 7 ? StudyLibraryHexDecode(slGroupRow[7]) : ""
        slSpeaker := slGroupRow.Length >= 8 ? StudyLibraryHexDecode(slGroupRow[8]) : ""
        slTags := slGroupRow.Length >= 9 ? StudyLibraryHexDecode(slGroupRow[9]) : ""
        slAddedToAnkiAt := slGroupRow.Length >= 10
            ? StudyLibraryHexDecode(slGroupRow[10]) : ""
        slAnkiStatus := slGroupRow.Length >= 11
            ? StudyLibraryHexDecode(slGroupRow[11]) : "not_checked"
        slAnkiNoteId := slGroupRow.Length >= 12 ? Integer(slGroupRow[12]) : 0
        slAnkiCheckedAt := slGroupRow.Length >= 13
            ? StudyLibraryHexDecode(slGroupRow[13]) : ""
        slKeyGrammar := slGroupRow.Length >= 14
            ? StudyLibraryHexDecode(slGroupRow[14]) : ""
        slEntry := Map(
            "id", slGroupId,
            "profile", slGroupProfile,
            "updated", slUpdated,
            "versions", slVersionCount,
            "preview", slPreview,
            "chapter", slChapter,
            "speaker", slSpeaker,
            "tags", slTags,
            "grammar", slKeyGrammar,
            "addedToAnkiAt", slAddedToAnkiAt,
            "ankiStatus", slAnkiStatus,
            "ankiNoteId", slAnkiNoteId,
            "ankiCheckedAt", slAnkiCheckedAt
        )
        slState["groups"].Push(slEntry)
        slState["list"].Add(
            "",
            slUpdated,
            slGroupProfile != "" ? slGroupProfile : "Unsorted",
            slChapter,
            slSpeaker,
            slTags,
            slPreview,
            slKeyGrammar,
            slVersionCount,
            StudyAnkiStatusLabel(slAnkiStatus, slAddedToAnkiAt),
            slGroupId
        )
        slTotalVersions += slVersionCount
        if (slAnkiStatus != "found" && slAddedToAnkiAt = "")
            slNotAddedCount += 1
        if (slGroupId = slPreviousGroup)
            slSelectedRow := slState["groups"].Length
    }
    StudyLibraryApplyCurrentSort(slState)
    slSelectedRow := StudyLibraryFindGroupRow(slState, slPreviousGroup)
    slState["suspend"] := false

    if slState["groups"].Length {
        if !slSelectedRow
            slSelectedRow := 1
        slState["list"].Modify(slSelectedRow, "Select Focus Vis")
        StudyLibraryLoadGroup(
            slState, StudyLibraryRowGroupId(slState, slSelectedRow)
        )
        StudyLibraryUpdateSelectionActions(slState)
        slState["status"].Value := slState["groups"].Length " source"
            . (slState["groups"].Length = 1 ? "" : "s") " • " slTotalVersions
            . " explanation" (slTotalVersions = 1 ? "" : "s") " • "
            . slNotAddedCount " not linked to Anki"
    } else {
        slSearchActive := (Trim(slState["search"].Value) != "")
            || (slState["profileMode"] != "all")
            || (slState["chapterMode"] != "all")
            || (slState["speakerMode"] != "all")
            || (slState["tagMode"] != "all")
            || (slState["ankiMode"] != "all")
            || (slState["dateMode"] != "all")
        StudyLibraryClearDetail(
            slState,
            slSearchActive ? "No explanations match the current filters."
                : "No explanations have been saved yet."
        )
        slState["status"].Value := slSearchActive
            ? "No matching explanations" : "Study Library is empty"
        StudyLibraryUpdateSelectionActions(slState)
    }
    } finally {
        if IsObject(slState)
            slState["refreshing"] := false
    }
    if StudyLibraryStateAlive(slState)
        SetTimer(StudyLibraryRefreshStorage.Bind(slState, false), -10)
}

StudyLibraryRedraw(slState, slErase := true, *) {
    if !StudyLibraryStateAlive(slState)
        return
    ; A full parent + child erase is useful after a discrete maximize/restore,
    ; but doing it synchronously for every mouse-move is the source of visible
    ; flashing in the large ListView/Edit regions. Interactive resizing uses a
    ; non-erasing invalidation and receives one complete repaint on exit.
    slFlags := slErase
        ? (0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400)
        ; Clear the old child positions, but leave painting asynchronous so
        ; successive mouse movements are composed into one frame.
        : (0x0001 | 0x0004 | 0x0080)
    try DllCall(
        "user32\RedrawWindow",
        "ptr", slState["gui"].Hwnd,
        "ptr", 0,
        "ptr", 0,
        "uint", slFlags
    )
}

StudyLibraryResize(slState, slGui, slMinMax, slWidth, slHeight) {
    if (slMinMax = -1 || slWidth < 500 || slHeight < 350)
        return
    slHwnd := slGui.Hwnd
    slVisible := slHwnd
        && DllCall("user32\IsWindowVisible", "ptr", slHwnd, "int")
    slInteractive := slState.Has("interactiveResize")
        && slState["interactiveResize"]
    if slInteractive {
        ; Defer the heavy ListView, Edit, and bitmap layout until the user
        ; releases the sizing border. The frame remains responsive and newly
        ; exposed space uses the window's normal background instead of forcing
        ; every native child to erase and repaint for every mouse pixel.
        slState["deferredResizeWidth"] := slWidth
        slState["deferredResizeHeight"] := slHeight
        return
    }
    if slVisible
        DllCall("user32\SendMessageW", "ptr", slHwnd, "uint", 0x000B,
            "ptr", 0, "ptr", 0) ; WM_SETREDRAW off
    try {
    slMargin := 14, slGap := 12, slToolbarH := 30, slStatusH := 30
    slState["layoutWidth"] := slWidth
    slButtonW := 78
    slToolbarY := slMargin
    slLibraryLabelW := 50, slLibraryW := 150, slNewLibraryW := 82
    slLibraryX := slMargin + slLibraryLabelW
    slNewLibraryX := slLibraryX + slLibraryW + 8
    slSearchX := slNewLibraryX + slNewLibraryW + slGap
    slSearchW := Min(
        520,
        Max(100, slWidth - slSearchX - slMargin
            - slButtonW * 2 - slGap * 2)
    )
    slSearchButtonX := slSearchX + slSearchW + slGap
    slRefreshX := slSearchButtonX + slButtonW + slGap
    slState["libraryLabel"].Move(
        slMargin, slToolbarY + 4, slLibraryLabelW, 22
    )
    slState["libraryDdl"].Move(
        slLibraryX, slToolbarY, slLibraryW, slToolbarH
    )
    slState["newLibraryButton"].Move(
        slNewLibraryX, slToolbarY, slNewLibraryW, slToolbarH
    )
    slState["search"].Move(slSearchX, slToolbarY, slSearchW, slToolbarH)
    slState["searchButton"].Move(slSearchButtonX, slToolbarY, slButtonW, slToolbarH)
    slState["refreshButton"].Move(slRefreshX, slToolbarY, slButtonW, slToolbarH)

    slFilterY := slToolbarY + slToolbarH + 8
    slFilterButtonW := 110, slColumnsW := 92
    slColumnsX := slMargin + slFilterButtonW + 8
    slEditX := slColumnsX + slColumnsW + 8
    slStudyX := slEditX + 140 + 8
    slAnkiX := slStudyX + 132 + 8
    slState["filterButton"].Move(
        slMargin, slFilterY, slFilterButtonW, slToolbarH
    )
    slState["columnsButton"].Move(slColumnsX, slFilterY, slColumnsW, slToolbarH)
    slState["editDetailsButton"].Move(slEditX, slFilterY, 140, slToolbarH)
    slState["studyButton"].Move(slStudyX, slFilterY, 132, slToolbarH)
    slState["ankiButton"].Move(slAnkiX, slFilterY, 82, slToolbarH)

    slContentY := slFilterY + slToolbarH + slGap
    slContentBottom := slHeight - slMargin - slStatusH
    slContentH := Max(240, slContentBottom - slContentY - slGap)
    ; The Study Reader now owns full-explanation reading. Keep this management
    ; view predictable: a compact fixed-width context pane and a flexible table.
    slPaneGap := 14
    slAvailableW := Max(1, slWidth - slMargin * 2 - slPaneGap)
    slRightW := Max(300, Min(325, slAvailableW - 250))
    slLeftW := Max(250, slAvailableW - slRightW)
    slRightX := slMargin + slLeftW + slPaneGap
    slState["list"].Move(slMargin, slContentY, slLeftW, slContentH)

    slDetailY := slContentY
    slState["detailTitle"].Move(slRightX, slDetailY, slRightW, 24)
    slVersionY := slDetailY + 27
    slState["versionLabel"].Move(slRightX, slVersionY + 4, 58, 22)
    slArrowW := 34, slVersionGap := 6
    slPreviousVersionX := slRightX + 58
    slState["previousVersion"].Move(
        slPreviousVersionX, slVersionY, slArrowW, 30
    )
    slVersionViewX := slPreviousVersionX + slArrowW + slVersionGap
    slVersionW := Max(100, slRightW - 58 - slArrowW * 2 - slVersionGap * 2)
    slState["versionView"].Move(
        slVersionViewX, slVersionY, slVersionW, 30
    )
    slState["nextVersion"].Move(
        slVersionViewX + slVersionW + slVersionGap,
        slVersionY, slArrowW, 30
    )
    slRemoveY := slVersionY + 34
    slRemoveW := 154
    slRemoveX := slRightX + 58
    slState["removeVersionButton"].Move(
        slRemoveX, slRemoveY, slRemoveW, 30
    )
    slMetaY := slRemoveY + 34
    slCompactDetails := slContentH < 400
    slMetaH := slCompactDetails ? 38 : 56
    slState["metadata"].Move(slRightX, slMetaY, slRightW, slMetaH)
    slImageLabelY := slMetaY + slMetaH + 2
    slState["imageLabel"].Move(slRightX, slImageLabelY, slRightW, 20)
    slImageY := slImageLabelY + 20
    slMinImageH := slCompactDetails ? 50 : 70
    slMinSourceH := slCompactDetails ? 30 : 42
    slRemainingForDetails := Max(90, slContentBottom - slImageY)
    slMaxImageH := Max(
        slMinImageH,
        slContentBottom - slImageY - 31 - 20 - slMinSourceH
    )
    slImageH := Min(
        260, slMaxImageH,
        Max(slMinImageH, Round(slRemainingForDetails * 0.55))
    )
    slState["imageFrame"].Move(slRightX, slImageY, slRightW, slImageH)
    slState["imageArea"] := Map(
        "x", slRightX + 4,
        "y", slImageY + 4,
        "w", Max(1, slRightW - 8),
        "h", Max(1, slImageH - 8)
    )
    slNavY := slImageY + slImageH + 5
    slImageInfoW := Max(70, slRightW - 200)
    slState["previousImage"].Move(slRightX, slNavY, 34, 27)
    slState["imageInfo"].Move(
        slRightX + 40, slNavY + 4, slImageInfoW, 22
    )
    slState["nextImage"].Move(
        slRightX + 40 + slImageInfoW + 6, slNavY, 34, 27
    )
    slState["openImage"].Move(slRightX + slRightW - 112, slNavY, 112, 27)

    slSourceLabelY := slNavY + 31
    slState["sourceLabel"].Move(slRightX, slSourceLabelY, slRightW, 20)
    slSourceY := slSourceLabelY + 20
    slSourceH := Max(slMinSourceH, slContentBottom - slSourceY)
    slState["source"].Move(slRightX, slSourceY, slRightW, slSourceH)
    slStorageW := 150
    slFooterY := slContentBottom + 3
    slState["status"].Move(
        slMargin, slFooterY + 4,
        Max(100, slWidth - slMargin * 2 - slStorageW - 10), 22
    )
    slState["storageButton"].Move(
        slWidth - slMargin - slStorageW, slFooterY, slStorageW, 27
    )
    if slVisible
        StudyLibraryQueueImageLayout(slState, 80)
    } finally {
        if slVisible {
            DllCall("user32\SendMessageW", "ptr", slHwnd, "uint", 0x000B,
                "ptr", 1, "ptr", 0) ; WM_SETREDRAW on
            StudyLibraryRedraw(slState, true)
        }
    }
    ; Maximize/restore composition can finish after the Size event returns.
    ; Coalesce one final repaint once Windows has settled on the new bounds.
    if slVisible
        SetTimer(slState["redrawCallback"], -160)
}

StudyLibrarySaveBounds(slState) {
    global iniPath
    if !(slState.Has("gui") && slState["gui"] && slState["gui"].Hwnd)
        return
    try {
        ; Keep the last useful restored bounds rather than replacing them with
        ; minimized or maximized coordinates.
        if (WinGetMinMax("ahk_id " slState["gui"].Hwnd) != 0)
            return
        slState["gui"].GetPos(&slX, &slY)
        slState["gui"].GetClientPos(,, &slW, &slH)
        if (slW < 640 || slH < 440)
            return
        IniWrite(slX, iniPath, "study_library_view", "x")
        IniWrite(slY, iniPath, "study_library_view", "y")
        IniWrite(slW, iniPath, "study_library_view", "w")
        IniWrite(slH, iniPath, "study_library_view", "h")
    }
}

StudyLibraryClose(slState, *) {
    global CPStudyLibraryState, CPStudyCandidateState
    if !IsObject(slState)
        return
    slState["closed"] := true
    try SetTimer(slState["imageLayoutCallback"], 0)
    try SetTimer(slState["redrawCallback"], 0)
    try SetTimer(slState["listRedrawCallback"], 0)
    try StudyLibraryCaptureColumnWidths(slState, true)
    try CPListViewSaveColumnOrder(
        slState["list"], "study_library_view", "columnOrder",
        slState["columns"].Length + 1, slState["columns"].Length + 1
    )
    try StudyLibrarySaveBounds(slState)
    try OnMessage(0x004E, slState["headerNotifyCallback"], 0)
    if IsObject(CPStudyCandidateState)
        try StudyCandidatesClose(CPStudyCandidateState)
    try slState["gui"].Destroy()
    CPStudyLibraryState := 0
    StudyStandaloneMaybeExit()
}

StudyStandaloneMaybeExit(*) {
    global CP_STUDY_ONLY_PROCESS, CPStudyLibraryState, CPStudyReaderState
    if !CP_STUDY_ONLY_PROCESS
        return
    if IsObject(CPStudyLibraryState) || IsObject(CPStudyReaderState)
        return
    SetTimer((*) => ExitApp(), -25)
}

SaveStudyLibraryWelcomePreference(dontShowAgain) {
    global iniPath
    try IniWriteRetry(
        dontShowAgain.Value ? 1 : 0,
        iniPath,
        "study_library",
        "welcomeDismissed"
    )
}

CloseStudyLibraryWelcome(dlg, dontShowAgain, *) {
    global CPStudyLibraryWelcomeDialog
    SaveStudyLibraryWelcomePreference(dontShowAgain)
    try dlg.Destroy()
    CPStudyLibraryWelcomeDialog := 0
}

OpenStudyLibraryWelcomeAnkiConnect(*) {
    try Run("https://ankiweb.net/shared/info/2055492159")
    catch as ex
        MsgBox(
            "The AnkiConnect page could not be opened:`n`n" ex.Message,
            "Study Library",
            "OK Icon!"
        )
}

OpenStudyLibraryWelcomeFolder(slState, *) {
    slDirectory := ""
    if (IsObject(slState) && slState.Has("database"))
        SplitPath(slState["database"],, &slDirectory)
    if (slDirectory = "")
        slDirectory := StudyLibraryConfiguredDirectory()
    try {
        if !DirExist(slDirectory)
            DirCreate(slDirectory)
        Run('explorer.exe "' slDirectory '"')
    } catch as ex {
        MsgBox(
            "The Study Library folder could not be opened:`n`n" ex.Message,
            "Study Library",
            "OK Icon!"
        )
    }
}

ShowStudyLibraryWelcome(slState, *) {
    global iniPath, CPStudyLibraryWelcomeDialog

    if !IsObject(slState) || !slState.Has("gui")
        return
    try {
        if !slState["gui"].Hwnd
            return
    } catch {
        return
    }
    if Integer(IniRead(
        iniPath, "study_library", "welcomeDismissed", 0
    ))
        return

    if (IsObject(CPStudyLibraryWelcomeDialog)
        && CPStudyLibraryWelcomeDialog.Has("gui")) {
        try {
            CPStudyLibraryWelcomeDialog["gui"].Show()
            WinActivate(
                "ahk_id " CPStudyLibraryWelcomeDialog["gui"].Hwnd
            )
            return
        }
    }

    dlg := Gui(
        "+Owner" slState["gui"].Hwnd " +AlwaysOnTop +OwnDialogs",
        "Welcome to the Study Library"
    )
    dlg.MarginX := 20
    dlg.MarginY := 18
    dlg.SetFont("s10", "Segoe UI")

    dlg.SetFont("s17 Bold")
    dlg.Add("Text", "xm w670", "Welcome to the Study Library")
    dlg.SetFont("s10 Norm")
    dlg.Add(
        "Text", "xm y+10 w670",
        "Explanations generated while playing can be saved here together "
        . "with their original Japanese and optional source screenshots."
    )
    dlg.Add(
        "Text", "xm y+12 w670",
        "Review and organize them later with Profiles, Libraries, filters, "
        . "chapters, speakers and tags. Open an entry in the Study Reader "
        . "to browse its sections, edit mistakes, copy text or create a new "
        . "explanation version."
    )
    dlg.Add(
        "Text", "xm y+12 w670",
        "Anki integration can check existing cards, send full explanations "
        . "or selected vocabulary, include compact screenshots, generate "
        . "example sentences and rank useful Anki candidates."
    )
    dlg.SetFont("s10 Bold")
    dlg.Add("Text", "xm y+16 w670", "Anki setup")
    dlg.SetFont("s10 Norm")
    dlg.Add(
        "Text", "xm y+5 w670",
        "Install AnkiConnect in desktop Anki, restart Anki and keep it "
        . "running while using Anki features. Then use Anki... > Anki link "
        . "check in this window to connect a Study Profile to a deck."
    )
    dlg.Add(
        "Text", "xm y+12 w670 cGray",
        "Study Libraries are stored locally in the application's Settings "
        . "folder, so include that folder in your backups. AI features send "
        . "requests to your configured provider and may incur normal API usage costs."
    )

    dontShowAgain := dlg.Add(
        "CheckBox", "xm y+16 Checked",
        "Don't show this introduction again"
    )
    btnAnkiConnect := dlg.Add(
        "Button", "xm y+20 w170", "Open AnkiConnect Page"
    )
    btnOpenFolder := dlg.Add(
        "Button", "x+10 yp w180", "Open Study Library Folder"
    )
    btnClose := dlg.Add("Button", "x+10 yp w110 Default", "Close")

    CPStudyLibraryWelcomeDialog := Map(
        "gui", dlg,
        "dontShowAgain", dontShowAgain
    )
    btnAnkiConnect.OnEvent(
        "Click", OpenStudyLibraryWelcomeAnkiConnect
    )
    btnOpenFolder.OnEvent(
        "Click", OpenStudyLibraryWelcomeFolder.Bind(slState)
    )
    closeCallback := CloseStudyLibraryWelcome.Bind(
        dlg, dontShowAgain
    )
    btnClose.OnEvent("Click", closeCallback)
    dlg.OnEvent("Escape", closeCallback)
    dlg.OnEvent("Close", closeCallback)

    dlg.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(dlg)
    try btnClose.Focus()
}

OpenStudyLibrary(*) {
    OpenStudyLibraryWindow(false)
}

OpenStandaloneStudyLibrary(*) {
    OpenStudyLibraryWindow(true)
}

OpenStudyLibraryWindow(slStandalone := false) {
    global CPStudyLibraryState, studyLibraryDir, ui, iniPath, controlDarkMode
    if (CPStudyLibraryState && CPStudyLibraryState.Has("gui")) {
        try {
            CPStudyLibraryState["gui"].Show()
            WinActivate("ahk_id " CPStudyLibraryState["gui"].Hwnd)
            StudyLibraryRefresh(CPStudyLibraryState)
            ; A programmatically selected ListView row remains an inactive gray
            ; selection until the table owns keyboard focus. Without this,
            ; Windows consumes the first double-click merely to focus it.
            CPStudyLibraryState["list"].Focus()
            return
        }
    }

    slActiveLibrary := StudyLibraryConfiguredName()
    StudyLibraryActivateName(slActiveLibrary)
    DirCreate(studyLibraryDir)
    slOutputDir := StudyBridgeCreateWorkspace("library")
    slGuiOptions := "+Resize +MinSize640x440 +OwnDialogs"
    if !slStandalone
        slGuiOptions .= " +Owner" ui.Hwnd
    slGui := Gui(
        slGuiOptions,
        "JRPG Translator - Study Library — " slActiveLibrary
    )
    slGui.MarginX := 14, slGui.MarginY := 14
    slGui.SetFont("s10", "Segoe UI")
    ; Establish the window palette before creating native child controls. If a
    ; ListView is born with the default Windows palette, its first visible frame
    ; can flash white even though the later dark-theme pass is done while hidden.
    slInitialColors := CPPalette(controlDarkMode)
    slGui.BackColor := slInitialColors["window"]
    CPApplyOwnedDialogTheme(slGui)

    slColumns := StudyLibraryCreateColumns()
    slColumnLabels := []
    for slColumn in slColumns
        slColumnLabels.Push(slColumn["label"])
    ; A permanently hidden ID column keeps selections mapped to the correct
    ; database row even after the visible table is sorted.
    slColumnLabels.Push("")
    slNowStamp := FormatTime(, "yyyyMMddHHmmss")
    slTodayStart := SubStr(slNowStamp, 1, 8) "000000"

    slLibraryLabel := slGui.Add("Text", "x14 y18 w50", "Library:")
    slLibraryDdl := slGui.Add("DropDownList", "x64 y14 w150 0x210", [])
    slNewLibraryButton := slGui.Add("Button", "x222 y14 w82", "Libraries...")
    slSearch := slGui.Add("Edit", "x316 y14 w328")
    try slSearch.SetCueBanner(
        "Japanese, explanation, grammar, chapter, speaker or tags"
    )
    slSearchButton := slGui.Add("Button", "x656 y14 w78 Default", "Search")
    slRefreshButton := slGui.Add("Button", "x746 y14 w78", "Refresh")
    slFilterButton := slGui.Add("Button", "x14 y52 w110 h30", "Filters...")
    slColumnsButton := slGui.Add("Button", "x132 y52 w92 h30", "Columns...")
    slEditDetailsButton := slGui.Add("Button", "x232 y52 w140 h30 Disabled", "Edit details...")
    slStudyButton := slGui.Add(
        "Button", "x380 y52 w132 h30 Disabled", "Open in Reader..."
    )
    slAnkiButton := slGui.Add("Button", "x520 y52 w82 h30", "Anki...")
    slListOptions := "x14 y94 w430 h530 Grid Multi"
        . " Background" slInitialColors["surface"]
        . " c" slInitialColors["text"]
    slList := slGui.Add(
        "ListView", slListOptions,
        slColumnLabels
    )
    ; The table is the largest repainted native surface in this window. Native
    ; ListView double-buffering removes row/grid flashing during live resizing
    ; without changing its dark-mode colors or selection behavior.
    SendMessage(0x1036, 0x10000, 0x10000, slList.Hwnd)
    slList.ModifyCol(slColumns.Length + 1, 0)
    slDetailTitle := slGui.Add("Text", "x468 y94 w638 h24 +0x200", "Select an explanation.")
    slDetailTitle.SetFont("s11 Bold")
    slVersionLabel := slGui.Add("Text", "x468 y125 w58", "Version:")
    ; Keep the original list hidden as the version data model. The visible
    ; read-only field and arrow buttons provide simpler navigation.
    slVersionDdl := slGui.Add(
        "DropDownList", "x0 y0 w1 h1 Hidden 0x210", []
    )
    slVersionView := slGui.Add(
        "Edit", "x526 y121 w300 h30 ReadOnly", ""
    )
    slRemoveVersionButton := slGui.Add(
        "Button", "x526 y155 w154 h30 Disabled", "Remove version..."
    )
    slPreviousVersion := slGui.Add(
        "Button", "x688 y155 w34 h30 Disabled", "‹"
    )
    slNextVersion := slGui.Add(
        "Button", "x728 y155 w34 h30 Disabled", "›"
    )
    slMetadata := slGui.Add("Text", "x468 y155 w638 h38 cGray", "")
    CPRegisterMutedControl(slMetadata)
    slImageLabel := slGui.Add("Text", "x468 y195 w638 h20", "Source screenshot")
    slImageFrame := slGui.Add("Text", "x468 y215 w638 h190 Border Background101010", "")
    slPicture := slGui.Add("Picture", "x472 y219 w1 h1 Hidden", "")
    slPreviousImage := slGui.Add("Button", "x468 y410 w34 h27", "‹")
    slNextImage := slGui.Add("Button", "x508 y410 w34 h27", "›")
    slImageInfo := slGui.Add("Text", "x550 y414 w300", "No source screenshot selected.")
    slOpenImage := slGui.Add("Button", "x994 y410 w112 h27", "Open full image")
    slSourceLabel := slGui.Add("Text", "x468 y443 w638 h20", "Original Japanese")
    slSource := slGui.Add("Edit", "x468 y463 w638 h90 ReadOnly Multi VScroll")
    ; These hidden controls retain the shared loading/parsing state used by the
    ; Study Reader without occupying space in the management window.
    slSectionLabel := slGui.Add("Text", "x0 y0 w1 h1 Hidden", "Section:")
    slSectionDdl := slGui.Add("DropDownList", "x0 y0 w1 Hidden", ["Full explanation"])
    slExplanation := slGui.Add("Edit", "x0 y0 w1 h1 Hidden ReadOnly Multi")
    slStatus := slGui.Add("Text", "x14 y736 w1092 h22 cGray", "Loading Study Library...")
    CPRegisterMutedControl(slStatus)
    slStorageButton := slGui.Add(
        "Button", "x956 y732 w150 h27 Disabled", "Storage..."
    )

    slState := Map(
        "gui", slGui,
        "database", studyLibraryDir "\study_library.db",
        "outputDir", slOutputDir,
        "libraryName", slActiveLibrary,
        "libraryNames", [],
        "libraryLabel", slLibraryLabel,
        "libraryDdl", slLibraryDdl,
        "newLibraryButton", slNewLibraryButton,
        "suspendLibrary", false,
        "profileLabels", ["All Profiles"],
        "profileChoices", [],
        "profileMode", "all",
        "profileFilter", "",
        "search", slSearch,
        "searchButton", slSearchButton,
        "refreshButton", slRefreshButton,
        "filterButton", slFilterButton,
        "columnsButton", slColumnsButton,
        "chapterLabels", ["Any chapter", "Chapter not set"],
        "chapterChoices", [
            Map("mode", "all", "value", ""),
            Map("mode", "empty", "value", "")
        ],
        "chapterMode", "all",
        "chapterFilter", "",
        "speakerLabels", ["Any speaker", "Speaker not set"],
        "speakerChoices", [
            Map("mode", "all", "value", ""),
            Map("mode", "empty", "value", "")
        ],
        "speakerMode", "all",
        "speakerFilter", "",
        "tagLabels", ["Any tag", "No tags"],
        "tagChoices", [
            Map("mode", "all", "value", ""),
            Map("mode", "empty", "value", "")
        ],
        "tagMode", "all",
        "tagFilter", "",
        "ankiMode", "all",
        "dateMode", "all",
        "dateFrom", slTodayStart,
        "dateTo", slNowStamp,
        "list", slList,
        "columns", slColumns,
        "sortColumn", 1,
        "sortDirection", "desc",
        "headerHwnd", 0,
        "lockingInternalColumn", false,
        "imageLayoutQueued", false,
        "imageLayoutLastTick", 0,
        "interactiveResize", false,
        "layoutWidth", 900,
        "standaloneWindow", slStandalone,
        "groups", [],
        "detailTitle", slDetailTitle,
        "studyButton", slStudyButton,
        "ankiButton", slAnkiButton,
        "editDetailsButton", slEditDetailsButton,
        "versionLabel", slVersionLabel,
        "versionDdl", slVersionDdl,
        "versionView", slVersionView,
        "previousVersion", slPreviousVersion,
        "nextVersion", slNextVersion,
        "removeVersionButton", slRemoveVersionButton,
        "versions", [],
        "metadata", slMetadata,
        "imageLabel", slImageLabel,
        "imageFrame", slImageFrame,
        "picture", slPicture,
        "previousImage", slPreviousImage,
        "nextImage", slNextImage,
        "imageInfo", slImageInfo,
        "openImage", slOpenImage,
        "imageArea", Map("x", 472, "y", 219, "w", 630, "h", 182),
        "media", [],
        "mediaIndex", 0,
        "sourceLabel", slSourceLabel,
        "source", slSource,
        "sectionLabel", slSectionLabel,
        "sectionDdl", slSectionDdl,
        "sections", [],
        "explanation", slExplanation,
        "status", slStatus,
        "storageButton", slStorageButton,
        "storage", 0,
        "storageLastTick", 0,
        "storageRefreshing", false,
        "currentGroupId", 0,
        "currentVersion", 0,
        "currentChapter", "",
        "currentSpeaker", "",
        "currentTags", "",
        "currentAddedToAnkiAt", "",
        "currentAnkiStatus", "not_checked",
        "currentAnkiNoteId", 0,
        "currentAnkiCheckedAt", "",
        "suspend", false,
        "closed", false,
        "refreshing", false
    )
    slState["imageLayoutCallback"] := StudyLibraryRunQueuedImageLayout.Bind(slState)
    slState["redrawCallback"] := StudyLibraryRedraw.Bind(slState)
    slState["listRedrawCallback"] := StudyLibraryRedrawList.Bind(slState)
    slState["sizeCallback"] := StudyLibraryResize.Bind(slState)
    slState["headerNotifyCallback"] := StudyLibraryHeaderNotify.Bind(slState)
    CPStudyLibraryState := slState

    OnMessage(0x004E, slState["headerNotifyCallback"])

    slLibraryDdl.OnEvent("Change", StudyLibraryLibraryChanged.Bind(slState))
    slNewLibraryButton.OnEvent("Click", StudyLibraryOpenManager.Bind(slState))
    slSearchButton.OnEvent("Click", StudyLibraryRefresh.Bind(slState))
    slRefreshButton.OnEvent("Click", StudyLibraryRefreshAll.Bind(slState))
    slStorageButton.OnEvent("Click", StudyLibraryOpenStorage.Bind(slState))
    slFilterButton.OnEvent("Click", StudyLibraryOpenFilters.Bind(slState))
    slColumnsButton.OnEvent("Click", StudyLibraryOpenColumns.Bind(slState))
    slList.OnEvent("ItemFocus", StudyLibraryGroupFocused.Bind(slState))
    slList.OnEvent("ItemSelect", StudyLibraryUpdateSelectionActions.Bind(slState))
    slList.OnEvent("DoubleClick", StudyLibraryOpenReaderFromList.Bind(slState))
    slList.OnEvent("ColClick", StudyLibraryColumnClicked.Bind(slState))
    slStudyButton.OnEvent("Click", StudyLibraryOpenSelectedReader.Bind(slState))
    slAnkiButton.OnEvent("Click", StudyLibraryShowAnkiMenu.Bind(slState))
    slPreviousVersion.OnEvent(
        "Click", StudyLibraryStepVersion.Bind(slState, -1)
    )
    slNextVersion.OnEvent(
        "Click", StudyLibraryStepVersion.Bind(slState, 1)
    )
    slEditDetailsButton.OnEvent("Click", StudyLibraryEditDetails.Bind(slState))
    slRemoveVersionButton.OnEvent("Click", StudyLibraryRemoveVersion.Bind(slState))
    slSectionDdl.OnEvent("Change", StudyLibrarySectionChanged.Bind(slState))
    slPreviousImage.OnEvent("Click", StudyLibraryPreviousImage.Bind(slState))
    slNextImage.OnEvent("Click", StudyLibraryNextImage.Bind(slState))
    slOpenImage.OnEvent("Click", StudyLibraryOpenImage.Bind(slState))
    slGui.OnEvent("Escape", StudyLibraryClose.Bind(slState))
    slGui.OnEvent("Close", StudyLibraryClose.Bind(slState))
    slGui.OnEvent("ContextMenu", StudyLibraryContextMenu.Bind(slState))

    slGui.Show("Hide w900 h600")
    slDpi := GetWindowDPI(slGui.Hwnd)
    slLogicalScale := 96 / Max(96, slDpi)
    slMaxW := Max(640, Floor(A_ScreenWidth * slLogicalScale) - 60)
    slMaxH := Max(440, Floor(A_ScreenHeight * slLogicalScale) - 80)
    ; Preserve the current adaptive first-run size, but allow a larger manually
    ; resized library to reopen at the user's preferred dimensions.
    slDefaultW := Max(640, Min(1180, slMaxW))
    slDefaultH := Max(440, Min(700, slMaxH))
    slSavedW := IniRead(
        iniPath, "study_library_view", "w", slDefaultW
    )
    slSavedH := IniRead(
        iniPath, "study_library_view", "h", slDefaultH
    )
    try slSavedW := Integer(slSavedW)
    catch
        slSavedW := slDefaultW
    try slSavedH := Integer(slSavedH)
    catch
        slSavedH := slDefaultH
    slSavedW := Max(640, Min(slMaxW, slSavedW))
    slSavedH := Max(440, Min(slMaxH, slSavedH))
    slGui.Show("Hide w" slSavedW " h" slSavedH)
    CPApplyOwnedDialogTheme(slGui)
    StudyLibraryRefreshLibrarySelector(slState, slActiveLibrary)
    StudyLibraryApplyColumns(slState)
    CPListViewRestoreColumnOrder(
        slList, "study_library_view", "columnOrder",
        slColumns.Length + 1, slColumns.Length + 1
    )
    slGui.GetClientPos(,, &slClientW, &slClientH)
    StudyLibraryResize(slState, slGui, 0, slClientW, slClientH)
    StudyLibraryRunBridge(slState, "ensure")
    StudyLibraryRefresh(slState)
    CPApplyOwnedDialogTheme(slGui)
    ; Theme application can recreate native header state, so restore the
    ; filter and sort indicators once the final theme pass is complete.
    StudyLibraryApplyHeaderIndicators(slState)
    ; Complete startup image/layout work while hidden. Otherwise the deferred
    ; resize timers can repaint and reposition controls just after first show.
    try SetTimer(slState["imageLayoutCallback"], 0)
    slState["imageLayoutQueued"] := false
    try SetTimer(slState["redrawCallback"], 0)
    StudyLibraryShowImage(slState)
    slSavedX := IniRead(
        iniPath, "study_library_view", "x", "__missing__"
    )
    slSavedY := IniRead(
        iniPath, "study_library_view", "y", "__missing__"
    )
    if ((slSavedX is number) && (slSavedY is number)) {
        slSavedX := Integer(slSavedX), slSavedY := Integer(slSavedY)
        StudyReaderClampPosition(slGui, &slSavedX, &slSavedY)
        StudyWindowRevealFinished(slGui, "x" slSavedX " y" slSavedY)
    } else
        StudyWindowRevealFinished(slGui, "Center")
    WinActivate("ahk_id " slGui.Hwnd)
    ; Give the populated table focus immediately so its initial selected row is
    ; active and the first double-click reaches the ListView event handler.
    slList.Focus()
    ; The Study Library has its own concise first-run introduction. Keeping its
    ; preference separate from the main setup guide also covers users who open
    ; the standalone Library before ever showing the control panel.
    SetTimer(ShowStudyLibraryWelcome.Bind(slState), -250)
    ; Delay live resizing for the same reason as in the Reader. In particular,
    ; this prevents the native ListView from rebuilding against a briefly
    ; erased right-hand surface immediately after first show.
    SetTimer(StudyWindowEnableLiveResize.Bind(slState), -300)
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
    global ddlAProv, ddlA_GM, ddlTR, ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt
    for c in [ddlAProv, ddlA_GM, ddlTR, ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt]
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
    CPApplyOwnedDialogTheme(dlg)
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

; ===================== Direct controller inputs =====================
CPControllerTokenDisplay(token) {
    static labels := Map(
        "X:DPAD_UP", "D-pad Up", "X:DPAD_DOWN", "D-pad Down",
        "X:DPAD_LEFT", "D-pad Left", "X:DPAD_RIGHT", "D-pad Right",
        "X:START", "Menu / Start", "X:BACK", "View / Back",
        "X:LS", "Left stick click", "X:RS", "Right stick click",
        "X:LB", "Left shoulder", "X:RB", "Right shoulder",
        "X:A", "A / Cross", "X:B", "B / Circle",
        "X:X", "X / Square", "X:Y", "Y / Triangle",
        "X:LT", "Left trigger", "X:RT", "Right trigger"
    )
    if labels.Has(token)
        return labels[token]
    if InStr(token, "J:") = 1
        return SubStr(token, 3)
    return token
}

CPJoystickNativeCapabilities(controllerId, forget := false) {
    ; AutoHotkey's legacy joystick lookups can enter unstable driver code when
    ; a device disappears between enumeration and a state query. Read and cache
    ; both the device name and its native axis ranges through WinMM instead.
    static cachedCapabilities := Map()
    controllerId := Integer(controllerId)
    if forget {
        if cachedCapabilities.Has(controllerId)
            cachedCapabilities.Delete(controllerId)
        return false
    }
    if cachedCapabilities.Has(controllerId)
        return cachedCapabilities[controllerId]

    ; JOYCAPSW is 728 bytes: two WORDs, a 32-character product name,
    ; 19 UINT capability fields, a 32-character registry key, and a
    ; 260-character OEM driver name.  szPname begins at byte offset 4.
    caps := Buffer(728, 0)
    result := 1
    try result := DllCall("winmm\joyGetDevCapsW"
        , "uptr", controllerId - 1, "ptr", caps.Ptr, "uint", caps.Size, "uint")
    if (result != 0)
        return false

    controllerName := ""
    try controllerName := Trim(StrGet(caps.Ptr + 4, 32, "UTF-16"))
    if (controllerName = "")
        controllerName := "Joystick " controllerId

    axisRanges := Map(
        "X", [NumGet(caps, 68, "uint"), NumGet(caps, 72, "uint")],
        "Y", [NumGet(caps, 76, "uint"), NumGet(caps, 80, "uint")],
        "Z", [NumGet(caps, 84, "uint"), NumGet(caps, 88, "uint")],
        "R", [NumGet(caps, 104, "uint"), NumGet(caps, 108, "uint")],
        "U", [NumGet(caps, 112, "uint"), NumGet(caps, 116, "uint")],
        "V", [NumGet(caps, 120, "uint"), NumGet(caps, 124, "uint")]
    )
    capabilities := Map("name", controllerName, "axisRanges", axisRanges)
    cachedCapabilities[controllerId] := capabilities
    return capabilities
}

CPJoystickNativeName(controllerId, forget := false) {
    capabilities := CPJoystickNativeCapabilities(controllerId, forget)
    if !IsObject(capabilities)
        return forget ? "" : "Joystick " Integer(controllerId)
    return capabilities["name"]
}

CPJoystickNativeAxisRanges(controllerId) {
    capabilities := CPJoystickNativeCapabilities(controllerId)
    return IsObject(capabilities) ? capabilities["axisRanges"] : false
}

CPJoystickNativeSnapshot(controllerId) {
    ; One joyGetPosEx call returns all 32 buttons and the POV hat together.
    ; This avoids repeatedly crossing AutoHotkey's legacy GetKeyState joystick
    ; path (previously up to 34 calls per device on every 20 ms timer tick).
    static JOY_RETURNALL := 0xFF
    stateBuffer := Buffer(52, 0) ; JOYINFOEX: thirteen DWORD values.
    NumPut("uint", stateBuffer.Size, stateBuffer, 0)
    NumPut("uint", JOY_RETURNALL, stateBuffer, 4)
    result := 1
    try result := DllCall("winmm\joyGetPosEx"
        , "uint", Integer(controllerId) - 1, "ptr", stateBuffer.Ptr, "uint")
    if (result != 0)
        return false
    return Map(
        "axes", Map(
            "X", NumGet(stateBuffer, 8, "uint"),
            "Y", NumGet(stateBuffer, 12, "uint"),
            "Z", NumGet(stateBuffer, 16, "uint"),
            "R", NumGet(stateBuffer, 20, "uint"),
            "U", NumGet(stateBuffer, 24, "uint"),
            "V", NumGet(stateBuffer, 28, "uint")
        ),
        "buttons", NumGet(stateBuffer, 32, "uint"),
        "pov", NumGet(stateBuffer, 40, "uint")
    )
}

CPControllerAppendLegacyState(snapshot, controllerId, legacyState) {
    controllerName := CPJoystickNativeName(controllerId)
    if !snapshot["connected"] {
        snapshot["connected"] := true
        snapshot["name"] := controllerName
    }
    pressedButtons := legacyState["buttons"]
    Loop 32 {
        buttonIndex := A_Index
        if (pressedButtons & (1 << (buttonIndex - 1)))
            snapshot["tokens"]["J:Button " buttonIndex] := true
    }
    pov := legacyState["pov"]
    if (pov < 36000) {
        if (pov >= 31500 || pov <= 4500)
            snapshot["tokens"]["J:D-pad Up"] := true
        if (pov >= 4500 && pov <= 13500)
            snapshot["tokens"]["J:D-pad Right"] := true
        if (pov >= 13500 && pov <= 22500)
            snapshot["tokens"]["J:D-pad Down"] := true
        if (pov >= 22500 && pov <= 31500)
            snapshot["tokens"]["J:D-pad Left"] := true
    }
}

CPControllerReadSnapshot() {
    static xinputButtons := [
        [0x0001, "X:DPAD_UP"], [0x0002, "X:DPAD_DOWN"],
        [0x0004, "X:DPAD_LEFT"], [0x0008, "X:DPAD_RIGHT"],
        [0x0010, "X:START"], [0x0020, "X:BACK"],
        [0x0040, "X:LS"], [0x0080, "X:RS"],
        [0x0100, "X:LB"], [0x0200, "X:RB"],
        [0x1000, "X:A"], [0x2000, "X:B"],
        [0x4000, "X:X"], [0x8000, "X:Y"]
    ]
    ; XInput is cheap enough to check every poll.  WinMM device discovery is
    ; deliberately throttled: repeatedly probing every possible legacy slot
    ; while a controller/driver is disappearing can destabilize vendor drivers.
    static controllerBackend := ""
    static legacyControllerIds := []
    static xinputDisconnectStartedMs := 0
    static nextLegacyDiscoveryMs := 0
    static DISCONNECT_GRACE_MS := 750
    static LEGACY_DISCOVERY_INTERVAL_MS := 1000
    snapshot := Map("connected", false, "name", "", "tokens", Map())
    nowMs := DllCall("kernel32\GetTickCount64", "uint64")

    Loop 4 {
        userIndex := A_Index - 1
        state := CPXInputGetState(userIndex)
        if !IsObject(state)
            continue
        if !snapshot["connected"] {
            snapshot["connected"] := true
            snapshot["name"] := "XInput controller " A_Index
        }
        buttons := state["buttons"]
        for buttonInfo in xinputButtons {
            if (buttons & buttonInfo[1])
                snapshot["tokens"][buttonInfo[2]] := true
        }
        if (state["leftTrigger"] >= 128)
            snapshot["tokens"]["X:LT"] := true
        if (state["rightTrigger"] >= 128)
            snapshot["tokens"]["X:RT"] := true
    }
    if snapshot["connected"] {
        controllerBackend := "xinput"
        legacyControllerIds := []
        xinputDisconnectStartedMs := 0
        nextLegacyDiscoveryMs := 0
        return snapshot
    }

    ; Do not immediately reinterpret an unplugged XInput device as a legacy
    ; joystick.  Give Windows and the controller driver time to settle first;
    ; a reconnect is still noticed immediately by the XInput loop above.
    if (controllerBackend = "xinput") {
        if !xinputDisconnectStartedMs
            xinputDisconnectStartedMs := nowMs
        if ((nowMs - xinputDisconnectStartedMs) < DISCONNECT_GRACE_MS)
            return snapshot
        controllerBackend := ""
        legacyControllerIds := []
        xinputDisconnectStartedMs := 0
        nextLegacyDiscoveryMs := 0
    } else {
        xinputDisconnectStartedMs := 0
    }

    ; DualSense and other non-XInput devices are exposed through the legacy
    ; joystick API. Once found, only the known device slots are polled at the
    ; normal 20 ms rate. Button numbering follows Windows' controller panel.
    if (controllerBackend = "legacy" && legacyControllerIds.Length) {
        stillConnectedIds := []
        for controllerId in legacyControllerIds {
            legacyState := CPJoystickNativeSnapshot(controllerId)
            if !IsObject(legacyState) {
                CPJoystickNativeName(controllerId, true)
                continue
            }
            stillConnectedIds.Push(controllerId)
            CPControllerAppendLegacyState(snapshot, controllerId, legacyState)
        }
        legacyControllerIds := stillConnectedIds
        if snapshot["connected"]
            return snapshot

        ; The active legacy controller disappeared.  Reset immediately, then
        ; defer broad discovery briefly instead of hammering all driver slots.
        controllerBackend := ""
        nextLegacyDiscoveryMs := nowMs + DISCONNECT_GRACE_MS
        return snapshot
    }

    if (nowMs < nextLegacyDiscoveryMs)
        return snapshot
    nextLegacyDiscoveryMs := nowMs + LEGACY_DISCOVERY_INTERVAL_MS

    detectedLegacyIds := []
    Loop 16 {
        controllerId := A_Index
        legacyState := CPJoystickNativeSnapshot(controllerId)
        if !IsObject(legacyState) {
            CPJoystickNativeName(controllerId, true)
            continue
        }
        detectedLegacyIds.Push(controllerId)
        CPControllerAppendLegacyState(snapshot, controllerId, legacyState)
    }
    if snapshot["connected"] {
        controllerBackend := "legacy"
        legacyControllerIds := detectedLegacyIds
    }
    return snapshot
}

CPControllerUpdateStatus(snapshot := 0) {
    global txtControllerStatus, CPControllerInputsEnabled, CPControllerLastStatusText
    if !IsSet(txtControllerStatus)
        return
    if !CPControllerInputsEnabled {
        newStatus := "Bindings off; navigation remains active."
    } else {
        if !IsObject(snapshot)
            snapshot := CPControllerReadSnapshot()
        newStatus := snapshot["connected"]
            ? "Detected: " snapshot["name"]
            : "No compatible controller detected."
    }
    if (newStatus != CPControllerLastStatusText) {
        txtControllerStatus.Value := newStatus
        CPControllerLastStatusText := newStatus
    }
}

CPControllerLoadBindings() {
    global iniPath, hotkeyActions, CPControllerBindings, CPControllerBindingEdits
    CPControllerBindings := Map()
    for actionKey in hotkeyActions {
        token := Trim(IniRead(iniPath, "controller_inputs", actionKey, ""))
        CPControllerBindings[actionKey] := token
        if CPControllerBindingEdits.Has(actionKey)
            CPControllerBindingEdits[actionKey].Value := token = "" ? "Disabled" : CPControllerTokenDisplay(token)
    }
}

CPControllerSaveBinding(action, token) {
    global iniPath, CPControllerBindings, CPControllerBindingEdits
    CPControllerBindings[action] := token
    IniWrite(token, iniPath, "controller_inputs", action)
    if CPControllerBindingEdits.Has(action)
        CPControllerBindingEdits[action].Value := token = "" ? "Disabled" : CPControllerTokenDisplay(token)
}

CPControllerCaptureDialog(action) {
    global ui, hotkeyLabels, controlDarkMode, CPControllerCaptureActive
    result := ""
    closed := false
    armed := false
    dlg := Gui("+Owner" ui.Hwnd " +MinSize500x190", "Assign controller button")
    dlg.MarginX := 22, dlg.MarginY := 18
    dlg.SetFont("s10", "Segoe UI")
    dlg.BackColor := CPPalette(controlDarkMode)["window"]
    dlg.Add("Text", "xm w500", "Action: " hotkeyLabels[action])
    dlg.Add("Text", "xm y+14 w500", "Release all buttons, then press the controller button to assign.")
    status := dlg.Add("Text", "xm y+18 w500 cGray", "Waiting for a controller...")
    cancel := dlg.Add("Button", "xm y+22 w120", "Cancel")
    cancel.OnEvent("Click", (*) => (closed := true))
    dlg.OnEvent("Close", (*) => (closed := true))
    dlg.OnEvent("Escape", (*) => (closed := true))

    CPControllerCaptureActive := true
    try {
        dlg.Show("w550 h210 Center")
        CPApplyOwnedDialogTheme(dlg)
        cancel.Focus()
        while !closed {
            snapshot := CPControllerReadSnapshot()
            if snapshot["connected"] {
                if !armed {
                    status.Value := "Detected " snapshot["name"] ". Release all buttons to arm capture."
                    if (snapshot["tokens"].Count = 0) {
                        armed := true
                        status.Value := "Ready. Press one controller button."
                    }
                } else if (result = "" && snapshot["tokens"].Count > 0) {
                    for token, _ in snapshot["tokens"] {
                        result := token
                        status.Value := "Assigned " CPControllerTokenDisplay(token) ". Release it to finish."
                        break
                    }
                } else if (result != "" && snapshot["tokens"].Count = 0) {
                    ; Keep exclusive ownership of the captured press until it
                    ; is released. Otherwise the same held shoulder reaches the
                    ; main navigation poll after this dialog closes and changes
                    ; the selected settings tab as an unintended second action.
                    closed := true
                }
            } else {
                status.Value := "No compatible controller detected."
            }
            Sleep(20)
        }
    } finally {
        try dlg.Destroy()
        CPControllerResetNavigation()
        CPControllerCaptureActive := false
    }
    return result
}

CPControllerAssign(action, *) {
    global hotkeyActions, hotkeyLabels, CPControllerBindings
    token := CPControllerCaptureDialog(action)
    if (token = "")
        return
    for otherAction in hotkeyActions {
        if (otherAction != action && CPControllerBindings.Has(otherAction)
         && CPControllerBindings[otherAction] = token) {
            response := MsgBox(
                CPControllerTokenDisplay(token) " is currently assigned to '"
                hotkeyLabels[otherAction] "'.`n`nMove it to '" hotkeyLabels[action] "'?",
                "Controller assignment", "YesNo Icon?")
            if (response != "Yes")
                return
            CPControllerSaveBinding(otherAction, "")
            break
        }
    }
    CPControllerSaveBinding(action, token)
}

CPControllerDisable(action, *) {
    CPControllerSaveBinding(action, "")
}

CPControllerSetEnabled(enabled, persist := true) {
    global iniPath, CPControllerInputsEnabled, CPControllerPreviousTokens
    global CPControllerLastDeviceName, CPControllerLastStatusText, cbControllerInputsEnabled
    CPControllerInputsEnabled := enabled ? true : false
    if IsSet(cbControllerInputsEnabled)
        cbControllerInputsEnabled.Value := CPControllerInputsEnabled ? 1 : 0
    if persist
        IniWrite(CPControllerInputsEnabled ? 1 : 0, iniPath, "controller_inputs", "enabled")
    CPControllerPreviousTokens := Map()
    CPControllerLastDeviceName := ""
    CPControllerLastStatusText := ""
    ; Basic foreground navigation is always available. When optional action
    ; bindings are off, the poll returns before reading a controller unless the
    ; control panel or one of its dialogs is actually in the foreground.
    SetTimer(CPControllerPoll, 20)
    CPControllerUpdateStatus()
}

CPControllerEnabledChanged(*) {
    global cbControllerInputsEnabled
    CPControllerSetEnabled(cbControllerInputsEnabled.Value != 0)
}

CPControllerSetDpadNavigationEnabled(enabled, persist := true) {
    global iniPath, CPControllerDpadNavigationEnabled
    global cbControllerDpadNavigationEnabled
    CPControllerDpadNavigationEnabled := enabled ? true : false
    if IsSet(cbControllerDpadNavigationEnabled)
        cbControllerDpadNavigationEnabled.Value := CPControllerDpadNavigationEnabled ? 1 : 0
    if persist
        IniWrite(CPControllerDpadNavigationEnabled ? 1 : 0, iniPath, "controller_inputs", "dpad_navigation")
    CPControllerResetNavigation()
}

CPControllerDpadNavigationChanged(*) {
    global cbControllerDpadNavigationEnabled
    CPControllerSetDpadNavigationEnabled(cbControllerDpadNavigationEnabled.Value != 0)
}

CPOverlayWindowHwnd(title) {
    oldMode := A_TitleMatchMode
    oldHidden := A_DetectHiddenWindows
    hwnd := 0
    try {
        SetTitleMatchMode(3)
        DetectHiddenWindows(true)
        hwnd := WinExist(title)
    } finally {
        DetectHiddenWindows(oldHidden)
        SetTitleMatchMode(oldMode)
    }
    return hwnd
}

CPControllerWriteOverlayCommand(commandName) {
    global overlayDir
    try {
        EnsureOverlayDir()
        commandPath := overlayDir "\" commandName
        if FileExist(commandPath)
            FileDelete(commandPath)
        FileAppend("", commandPath, "UTF-8")
        return true
    } catch as ex {
        DbgCP("Controller command failed for " commandName ": " ex.Message)
        return false
    }
}

CPControllerTranslatorAction(commandName) {
    if !CPOverlayWindowHwnd("Translator")
        LaunchOverlay()
    if !CPOverlayWindowHwnd("Translator") {
        Toast("Translator is not running")
        return
    }
    try ApplyShotSettings()
    CPControllerWriteOverlayCommand(commandName)
}

CPControllerToggleTranslator() {
    if CPOverlayWindowHwnd("Translator")
        CPControllerWriteOverlayCommand("cmd.toggle_translator")
    else
        LaunchOverlay()
}

CPControllerToggleExplainer() {
    if CPOverlayWindowHwnd("Explainer")
        CPControllerWriteOverlayCommand("cmd.toggle_explainer")
    else
        LaunchExplainerOverlay()
}

CPControllerDispatchAction(action) {
    try {
        switch action {
            case "screenshot_translate":
                CPControllerTranslatorAction("cmd.oneshot_translate")
            case "explain_last_translation":
                ExplainNow()
            case "hide_show_translator":
                CPControllerToggleTranslator()
            case "hide_show_explainer":
                CPControllerToggleExplainer()
            case "hide_show_control_panel":
                ToggleControlPanel()
            case "take_screenshot":
                CPControllerTranslatorAction("cmd.take_screenshot")
            case "screenshot_translation":
                CPControllerTranslatorAction("cmd.screenshot_translation")
            case "launch_explainer_request":
                CP_LaunchExplainerRequest()
            case "recapture_region":
                CPControllerTranslatorAction("cmd.recapture_region")
            case "start_stop_audio":
                StartStopAudio()
        }
    } catch as ex {
        DbgCP("Direct controller action failed for " action ": " ex.Message)
        Toast("Controller action failed")
    }
}

CPControllerNavigationTarget() {
    global ui
    if !(IsSet(ui) && ui && ui.Hwnd)
        return 0
    if !DllCall("user32\IsWindowVisible", "ptr", ui.Hwnd, "int")
        return 0

    cpNavForeground := DllCall("user32\GetForegroundWindow", "ptr")
    if !cpNavForeground
        return 0

    ; Owned dialogs are part of the control-panel workflow. Follow the owner
    ; chain instead of matching titles so every existing and future dialog can
    ; use the same default controller navigation safely.
    cpNavOwner := cpNavForeground
    Loop 16 {
        if (cpNavOwner = ui.Hwnd)
            return cpNavForeground
        cpNavOwner := DllCall("user32\GetWindow", "ptr", cpNavOwner, "uint", 4, "ptr") ; GW_OWNER
        if !cpNavOwner
            break
    }
    return 0
}

CPControllerNavigationState(snapshot) {
    global CPControllerDpadNavigationEnabled
    cpNavTokens := snapshot["tokens"]
    cpNavIsXInput := InStr(snapshot["name"], "XInput controller ") = 1
    cpNavIsPlayStation := RegExMatch(snapshot["name"], "i)(DualSense|DualShock|Wireless Controller|PlayStation)")
    cpNavConfirmToken := cpNavIsPlayStation ? "J:Button 2" : "J:Button 1"
    cpNavCancelToken := cpNavIsPlayStation ? "J:Button 3" : "J:Button 2"
    cpNavDpadEnabled := CPControllerDpadNavigationEnabled

    return Map(
        "Up", cpNavDpadEnabled && cpNavTokens.Has(cpNavIsXInput ? "X:DPAD_UP" : "J:D-pad Up"),
        "Down", cpNavDpadEnabled && cpNavTokens.Has(cpNavIsXInput ? "X:DPAD_DOWN" : "J:D-pad Down"),
        "Left", cpNavDpadEnabled && cpNavTokens.Has(cpNavIsXInput ? "X:DPAD_LEFT" : "J:D-pad Left"),
        "Right", cpNavDpadEnabled && cpNavTokens.Has(cpNavIsXInput ? "X:DPAD_RIGHT" : "J:D-pad Right"),
        "PreviousTab", cpNavTokens.Has(cpNavIsXInput ? "X:LB" : "J:Button 5"),
        "NextTab", cpNavTokens.Has(cpNavIsXInput ? "X:RB" : "J:Button 6"),
        "Activate", cpNavTokens.Has(cpNavIsXInput ? "X:A" : cpNavConfirmToken),
        "Cancel", cpNavTokens.Has(cpNavIsXInput ? "X:B" : cpNavCancelToken)
    )
}

CPControllerResetNavigation(targetHwnd := 0, navState := 0) {
    global CPControllerNavPreviousState, CPControllerNavTargetHwnd
    global CPControllerNavHeldDirection, CPControllerNavNextRepeatAt, CPControllerNavHeldSince
    CPControllerNavTargetHwnd := targetHwnd
    CPControllerNavPreviousState := IsObject(navState) ? navState : Map()
    CPControllerNavHeldDirection := ""
    CPControllerNavNextRepeatAt := 0
    CPControllerNavHeldSince := 0
}

CPControllerSendDialogKey(keyName) {
    ; A send level above the dialog hotkeys' default input level lets existing
    ; custom keyboard navigation handle the event. Dialogs without custom
    ; handlers simply receive the native key.
    try {
        SendLevel(1)
        SendEvent("{" keyName "}")
    } finally {
        SendLevel(0)
    }
}

CPControllerDispatchNavigation(command, targetHwnd) {
    global ui, CPControllerLastNativeNavigationAt

    if (targetHwnd = ui.Hwnd) {
        CPControllerLastNativeNavigationAt[command] := A_TickCount
        switch command {
            case "Up", "Down", "Left", "Right":
                CPNavMove(command)
            case "PreviousTab":
                CPNavSwitchTab(-1)
            case "NextTab":
                CPNavSwitchTab(1)
            case "Activate":
                ; Bypass the keyboard mirror guard: this command originated
                ; from the native controller path itself.
                CPNavActivate("Enter")
            case "Cancel":
                CPNavCancelCurrent()
        }
        return
    }

    if CPControllerColorDispatch(command, targetHwnd)
        return

    switch command {
        case "Up", "Down", "Left", "Right":
            CPControllerSendDialogKey(command)
        case "Activate":
            CPControllerSendDialogKey("Enter")
        case "Cancel":
            CPControllerSendDialogKey("Esc")
    }
}

CPControllerHandleNavigation(snapshot, targetHwnd) {
    global CPControllerNavPreviousState, CPControllerNavTargetHwnd
    global CPControllerNavHeldDirection, CPControllerNavNextRepeatAt, CPControllerNavHeldSince
    cpNavState := CPControllerNavigationState(snapshot)

    ; Opening the panel with a controller button must not immediately activate
    ; whichever control receives focus. The first frame establishes a baseline.
    if (targetHwnd != CPControllerNavTargetHwnd || CPControllerNavPreviousState.Count = 0) {
        CPControllerResetNavigation(targetHwnd, cpNavState)
        return
    }

    cpNavDirections := ["Up", "Down", "Left", "Right"]
    cpNavNewDirection := ""
    for cpNavDirection in cpNavDirections {
        if (cpNavState[cpNavDirection] && !CPControllerNavPreviousState[cpNavDirection]) {
            cpNavNewDirection := cpNavDirection
            break
        }
    }

    if (cpNavNewDirection != "") {
        CPControllerDispatchNavigation(cpNavNewDirection, targetHwnd)
        CPControllerNavHeldDirection := cpNavNewDirection
        CPControllerNavHeldSince := A_TickCount
        cpNavAcceleratedSliderRepeat := CPControllerAcceleratedSliderRepeatActive(targetHwnd, cpNavNewDirection)
        CPControllerNavNextRepeatAt := A_TickCount + (cpNavAcceleratedSliderRepeat ? 280 : 340)
    } else if (CPControllerNavHeldDirection != "" && cpNavState[CPControllerNavHeldDirection]) {
        if (A_TickCount >= CPControllerNavNextRepeatAt) {
            ; Tab changes are discrete controller actions. Repeating a held
            ; D-pad direction here makes a single deliberate press skip tabs,
            ; especially when a mapper also mirrors it as an arrow key.
            cpNavFocusedHwnd := targetHwnd = ui.Hwnd ? CPFocusedHwnd() : 0
            cpNavAcceleratedSliderRepeat := CPControllerAcceleratedSliderRepeatActive(
                targetHwnd, CPControllerNavHeldDirection)
            cpNavRepeatSteps := 1
            cpNavRepeatDelay := 90
            if cpNavAcceleratedSliderRepeat {
                cpNavHeldMs := A_TickCount - CPControllerNavHeldSince
                cpNavRepeatSteps := cpNavHeldMs >= 1100 ? 4 : (cpNavHeldMs >= 650 ? 2 : 1)
                cpNavRepeatDelay := 55
            }
            if !(cpNavFocusedHwnd && CPHwndIsTab(cpNavFocusedHwnd)) {
                Loop cpNavRepeatSteps
                    CPControllerDispatchNavigation(CPControllerNavHeldDirection, targetHwnd)
            }
            CPControllerNavNextRepeatAt := A_TickCount + cpNavRepeatDelay
        }
    } else {
        CPControllerNavHeldDirection := ""
        CPControllerNavNextRepeatAt := 0
        CPControllerNavHeldSince := 0
    }

    if (cpNavState["Activate"] && !CPControllerNavPreviousState["Activate"])
        CPControllerDispatchNavigation("Activate", targetHwnd)
    if (cpNavState["Cancel"] && !CPControllerNavPreviousState["Cancel"])
        CPControllerDispatchNavigation("Cancel", targetHwnd)
    ; Shoulder navigation is intentionally edge-triggered: holding a button
    ; must never skip several settings tabs. Owned dialogs continue suppressing
    ; gameplay actions but do not reinterpret the shoulders as dialog input.
    if (targetHwnd = ui.Hwnd) {
        if (cpNavState["PreviousTab"] && !CPControllerNavPreviousState["PreviousTab"])
            CPControllerDispatchNavigation("PreviousTab", targetHwnd)
        if (cpNavState["NextTab"] && !CPControllerNavPreviousState["NextTab"])
            CPControllerDispatchNavigation("NextTab", targetHwnd)
    }

    CPControllerNavPreviousState := cpNavState
}

CPControllerPoll(*) {
    global CPControllerInputsEnabled, CPControllerCaptureActive
    global CPControllerPreviousTokens, CPControllerBindings, hotkeyActions
    global CPControllerLastDeviceName, CPOverlayAdjustState
    if CPControllerCaptureActive
        return
    if ((CPOverlayAdjustState.Has("active") && CPOverlayAdjustState["active"])
     || CPControllerModalFlagActive()) {
        CPControllerPreviousTokens := Map()
        CPControllerResetNavigation()
        return
    }

    cpNavTarget := CPControllerNavigationTarget()
    if (!CPControllerInputsEnabled && !cpNavTarget) {
        CPControllerResetNavigation()
        return
    }

    snapshot := CPControllerReadSnapshot()
    CPControllerUpdateStatus(snapshot)
    if !snapshot["connected"] {
        CPControllerPreviousTokens := Map()
        CPControllerLastDeviceName := ""
        CPControllerResetNavigation(cpNavTarget)
        return
    }

    if cpNavTarget {
        CPControllerHandleNavigation(snapshot, cpNavTarget)
        CPControllerPreviousTokens := snapshot["tokens"]
        CPControllerLastDeviceName := snapshot["name"]
        return
    }

    CPControllerResetNavigation()
    if !CPControllerInputsEnabled
        return

    ; On connection or controller changes, use the first frame only as the
    ; baseline so a button already being held cannot trigger unexpectedly.
    if (snapshot["name"] != CPControllerLastDeviceName) {
        CPControllerLastDeviceName := snapshot["name"]
        CPControllerPreviousTokens := snapshot["tokens"]
        return
    }

    for actionKey in hotkeyActions {
        if !CPControllerBindings.Has(actionKey)
            continue
        token := CPControllerBindings[actionKey]
        if (token != "" && snapshot["tokens"].Has(token)
         && !CPControllerPreviousTokens.Has(token)) {
            CPControllerDispatchAction(actionKey)
        }
    }
    CPControllerPreviousTokens := snapshot["tokens"]
}

CPAddControlsViewControl(viewName, ctrl) {
    global CPControlsKeyboardControls, CPControlsControllerControls
    if (viewName = "controller")
        CPControlsControllerControls.Push(ctrl)
    else
        CPControlsKeyboardControls.Push(ctrl)
    return ctrl
}

CPSetControlsView(viewName, persist := true, *) {
    global iniPath, CPControlsCurrentView, CPControlsKeyboardControls, CPControlsControllerControls
    global rbControlsKeyboard, rbControlsController
    if (viewName != "controller")
        viewName := "keyboard"
    CPControlsCurrentView := viewName
    if IsSet(rbControlsKeyboard)
        rbControlsKeyboard.Value := viewName = "keyboard" ? 1 : 0
    if IsSet(rbControlsController)
        rbControlsController.Value := viewName = "controller" ? 1 : 0
    for ctrl in CPControlsKeyboardControls
        try ctrl.Visible := viewName = "keyboard"
    for ctrl in CPControlsControllerControls
        try ctrl.Visible := viewName = "controller"
    if persist
        IniWrite(viewName, iniPath, "controller_inputs", "view")
}
; =====================================================================

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
    newModel := Trim(CPThemedInputBox("Add model:", "Add", "", "", 360).Value)
    return AddModelValue(arr, key, combo, newModel)
}

CPApplyOwnedDialogTheme(dlg) {
    global controlDarkMode, CPThemedDialogHwnds, CPThemeBrushWindow
    if !(IsObject(dlg) && dlg.Hwnd)
        return

    CPThemedDialogHwnds[dlg.Hwnd] := true
    dialogColors := CPPalette(controlDarkMode)
    dlg.BackColor := dialogColors["window"]
    CPApplyDarkTitleBar(dlg.Hwnd, controlDarkMode)
    try CPSetPreferredAppDarkMode(controlDarkMode, dlg.Hwnd)
    CPAllowDarkModeForWindow(dlg.Hwnd, controlDarkMode)
    CPApplyWindowScrollbarTheme(dlg.Hwnd, controlDarkMode)
    if !CPThemeBrushWindow
        CPRefreshThemeBrushes()

    ; Dialogs are often built and themed while still hidden so their first
    ; visible frame is complete. Enumerate their child controls explicitly;
    ; otherwise native tabs, buttons and tables can retain the light theme.
    cpDialogControls := []
    cpDialogOldDetectHidden := A_DetectHiddenWindows
    try {
        DetectHiddenWindows true
        cpDialogControls := WinGetControlsHwnd("ahk_id " dlg.Hwnd)
    } finally {
        DetectHiddenWindows cpDialogOldDetectHidden
    }
    try {
        cpDialogTitle := WinGetTitle("ahk_id " dlg.Hwnd)
        cpDialogIsStudy := InStr(cpDialogTitle, "Study") > 0
            || InStr(cpDialogTitle, "Recommendation") > 0
            || cpDialogTitle = "Generate recommendations"
        for controlHwnd in cpDialogControls {
            if cpDialogIsStudy {
                cpDialogControlClass := WinGetClass("ahk_id " controlHwnd)
                if (cpDialogControlClass = "ComboBox") {
                    CPPrepareStudyCombo(controlHwnd)
                } else if (cpDialogControlClass = "SysListView32") {
                    cpDialogHeader := SendMessage(0x101F, 0, 0, controlHwnd)
                    CPPrepareStudyListHeader(cpDialogHeader)
                }
            }
            CPApplyThemeToControl(controlHwnd, controlDarkMode)
        }
    }
    try DllCall("user32\RedrawWindow", "ptr", dlg.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x185)
}

CPApplyOpenStudyWindowThemes() {
    global CPStudyLibraryState, CPStudyReaderState, CPStudyCandidateState
    for cpStudyState in [
        CPStudyLibraryState, CPStudyReaderState, CPStudyCandidateState
    ] {
        if !(IsObject(cpStudyState) && cpStudyState.Has("gui"))
            continue
        try {
            CPApplyOwnedDialogTheme(cpStudyState["gui"])
            if cpStudyState.Has("redrawCallback")
                cpStudyState["redrawCallback"].Call()
        }
    }
}

CPThemedDialogFinish(cpDialogState, cpDialog, cpResult, *) {
    if cpDialogState["closed"]
        return
    cpDialogState["result"] := cpResult
    cpDialogState["closed"] := true
    try cpDialog.Destroy()
}

CPThemedOwnedMessage(ownerHwnd, message, dialogTitle := ""
    , buttons := "ok", icon := "warning", dialogWidth := 520) {
    global controlDarkMode

    cpOwnerValid := ownerHwnd
        && DllCall("user32\IsWindow", "ptr", ownerHwnd, "int")
    cpDefaultResult := buttons = "yesno" ? "No" : "OK"
    cpDialogState := Map("closed", false, "result", cpDefaultResult)
    cpOptions := "+AlwaysOnTop"
    if cpOwnerValid
        cpOptions .= " +Owner" ownerHwnd
    cpDialog := Gui(cpOptions, dialogTitle)
    cpDialog.MarginX := 18
    cpDialog.MarginY := 16
    cpColors := CPPalette(controlDarkMode)
    cpDialog.BackColor := cpColors["window"]
    cpDialog.SetFont("s10 c" cpColors["text"], "Segoe UI")

    switch icon {
        case "error":
            cpIconGlyph := Chr(0x2715), cpIconColor := "F85149"
        case "info":
            cpIconGlyph := "i", cpIconColor := "58A6FF"
        default:
            cpIconGlyph := Chr(0x26A0), cpIconColor := "D29922"
    }
    cpIcon := cpDialog.Add(
        "Text", "xm ym w42 h50 Center c" cpIconColor, cpIconGlyph
    )
    cpIcon.SetFont("s24 Bold c" cpIconColor, "Segoe UI Symbol")
    cpMessageWidth := Max(300, dialogWidth - 78)
    cpDialog.Add(
        "Text", "x+12 yp+2 w" cpMessageWidth " c" cpColors["text"],
        message
    )

    if (buttons = "yesno") {
        cpYes := cpDialog.Add("Button", "xm y+18 w100", "Yes")
        cpNo := cpDialog.Add("Button", "x+10 yp w100 Default", "No")
        cpYes.OnEvent(
            "Click", CPThemedDialogFinish.Bind(
                cpDialogState, cpDialog, "Yes"
            )
        )
        cpNo.OnEvent(
            "Click", CPThemedDialogFinish.Bind(
                cpDialogState, cpDialog, "No"
            )
        )
        cpInitialButton := cpNo
    } else {
        cpOK := cpDialog.Add("Button", "xm y+18 w110 Default", "OK")
        cpOK.OnEvent(
            "Click", CPThemedDialogFinish.Bind(
                cpDialogState, cpDialog, "OK"
            )
        )
        cpInitialButton := cpOK
    }
    cpDialog.OnEvent(
        "Escape", CPThemedDialogFinish.Bind(
            cpDialogState, cpDialog, cpDefaultResult
        )
    )
    cpDialog.OnEvent(
        "Close", CPThemedDialogFinish.Bind(
            cpDialogState, cpDialog, cpDefaultResult
        )
    )

    cpDialog.Show("Hide AutoSize")
    CPApplyOwnedDialogTheme(cpDialog)
    if cpOwnerValid
        try DllCall("user32\EnableWindow", "ptr", ownerHwnd, "int", 0)
    try {
        cpDialog.Show("Center")
        try cpInitialButton.Focus()
        while !cpDialogState["closed"]
            Sleep(25)
    } finally {
        try cpDialog.Destroy()
        if cpOwnerValid
            try DllCall("user32\EnableWindow", "ptr", ownerHwnd, "int", 1)
        if cpOwnerValid && DllCall("user32\IsWindow", "ptr", ownerHwnd, "int")
            try WinActivate("ahk_id " ownerHwnd)
    }
    return cpDialogState["result"]
}

CPThemedChoicePopupFinish(cpPopupState, cpPopup, cpResult, *) {
    if cpPopupState["closed"]
        return
    cpPopupState["result"] := cpResult
    cpPopupState["closed"] := true
    try cpPopup.Destroy()
}

CPThemedChoicePopupRegistry() {
    static cpRegistry := Map()
    return cpRegistry
}

CPThemedChoicePopupPaintFocus(cpPopupState, cpFocusIndex) {
    cpColors := cpPopupState["colors"]
    for cpRowIndex, cpRow in cpPopupState["rows"] {
        cpRowBack := cpRowIndex = cpFocusIndex
            ? cpColors["focus"] : cpColors["surface"]
        try cpRow.Opt("+Background" cpRowBack)
    }
    cpPopupState["focusIndex"] := cpFocusIndex
}

CPThemedChoicePopupFocus(cpPopupState, cpFocusIndex) {
    cpRowCount := cpPopupState["rows"].Length
    if !cpRowCount
        return
    if (cpFocusIndex < 1)
        cpFocusIndex := cpRowCount
    else if (cpFocusIndex > cpRowCount)
        cpFocusIndex := 1
    CPThemedChoicePopupPaintFocus(cpPopupState, cpFocusIndex)
    try cpPopupState["rows"][cpFocusIndex].Focus()
}

CPThemedChoicePopupOnKeyDown(wParam, lParam, msg, hwnd) {
    cpRootHwnd := DllCall("user32\GetAncestor", "ptr", hwnd, "uint", 2, "ptr")
    cpRegistry := CPThemedChoicePopupRegistry()
    if !cpRootHwnd || !cpRegistry.Has(cpRootHwnd)
        return
    cpPopupState := cpRegistry[cpRootHwnd]
    if cpPopupState["closed"]
        return
    switch wParam {
        case 0x26, 0x25: ; Up / Left
            CPThemedChoicePopupFocus(
                cpPopupState, cpPopupState["focusIndex"] - 1
            )
            return true
        case 0x28, 0x27: ; Down / Right
            CPThemedChoicePopupFocus(
                cpPopupState, cpPopupState["focusIndex"] + 1
            )
            return true
        case 0x0D, 0x20: ; Enter / Space
            cpPopupResult := cpPopupState.Has("resultValues")
                ? cpPopupState["resultValues"][cpPopupState["focusIndex"]]
                : cpPopupState["focusIndex"]
            CPThemedChoicePopupFinish(
                cpPopupState, cpPopupState["gui"],
                cpPopupResult
            )
            return true
    }
}

CPThemedChoicePopupRegister(cpPopupState) {
    static cpKeyHandlerRegistered := false
    if !cpKeyHandlerRegistered {
        OnMessage(0x0100, CPThemedChoicePopupOnKeyDown) ; WM_KEYDOWN
        cpKeyHandlerRegistered := true
    }
    CPThemedChoicePopupRegistry()[cpPopupState["gui"].Hwnd] := cpPopupState
}

CPThemedChoicePopupUnregister(cpPopupState) {
    cpRegistry := CPThemedChoicePopupRegistry()
    try {
        cpPopupHwnd := cpPopupState["gui"].Hwnd
        if cpRegistry.Has(cpPopupHwnd)
            cpRegistry.Delete(cpPopupHwnd)
    }
}

CPThemedChoicePopupWatch(cpPopupState, cpPopupHwnd, cpPopup, *) {
    if cpPopupState["closed"]
        return
    if !DllCall("user32\IsWindow", "ptr", cpPopupHwnd, "int")
        return CPThemedChoicePopupFinish(cpPopupState, cpPopup, 0)
    if !WinActive("ahk_id " cpPopupHwnd)
        CPThemedChoicePopupFinish(cpPopupState, cpPopup, 0)
}

CPThemedChoicePopup(ownerHwnd, anchorHwnd, choices) {
    global controlDarkMode
    if !choices.Length
        return 0

    cpOwnerValid := ownerHwnd
        && DllCall("user32\IsWindow", "ptr", ownerHwnd, "int")
    cpPopupState := Map("closed", false, "result", 0)
    cpOptions := "+ToolWindow -Caption +Border +AlwaysOnTop"
    if cpOwnerValid
        cpOptions .= " +Owner" ownerHwnd
    cpPopup := Gui(cpOptions)
    cpPopup.MarginX := 3
    cpPopup.MarginY := 3
    cpColors := CPPalette(controlDarkMode)
    cpPopup.BackColor := cpColors["window"]
    cpPopup.SetFont("s10 c" cpColors["text"], "Segoe UI")
    cpPopupState["gui"] := cpPopup
    cpPopupState["colors"] := cpColors
    cpPopupState["rows"] := []
    cpPopupState["focusIndex"] := 1
    for cpChoiceIndex, cpChoiceText in choices {
        cpPosition := cpChoiceIndex = 1 ? "xm ym" : "xm y+1"
        cpChoiceRow := cpPopup.Add(
            "Text",
            cpPosition " w250 h30 0x300 +Tabstop Border"
                . " Background" cpColors["surface"]
                . " c" cpColors["text"],
            "  " cpChoiceText
        )
        cpChoiceRow.OnEvent(
            "Click", CPThemedChoicePopupFinish.Bind(
                cpPopupState, cpPopup, cpChoiceIndex
            )
        )
        cpPopupState["rows"].Push(cpChoiceRow)
    }
    cpPopup.OnEvent(
        "Escape", CPThemedChoicePopupFinish.Bind(cpPopupState, cpPopup, 0)
    )
    cpPopup.OnEvent(
        "Close", CPThemedChoicePopupFinish.Bind(cpPopupState, cpPopup, 0)
    )

    cpPopup.Show("Hide AutoSize")
    CPApplyOwnedDialogTheme(cpPopup)
    ; CPApplyOwnedDialogTheme uses the normal window background for generic
    ; text controls. Restore the menu-row surfaces afterward.
    CPThemedChoicePopupPaintFocus(cpPopupState, 1)
    cpPopup.GetPos(,, &cpPopupW, &cpPopupH)
    cpAnchorRect := Buffer(16, 0)
    if !(anchorHwnd && DllCall(
        "user32\GetWindowRect", "ptr", anchorHwnd,
        "ptr", cpAnchorRect.Ptr, "int"
    )) {
        CoordMode("Mouse", "Screen")
        MouseGetPos(&cpX, &cpY)
        cpAnchorTop := cpY
    } else {
        cpAnchorLeft := NumGet(cpAnchorRect, 0, "int")
        cpAnchorTop := NumGet(cpAnchorRect, 4, "int")
        cpAnchorRight := NumGet(cpAnchorRect, 8, "int")
        cpAnchorBottom := NumGet(cpAnchorRect, 12, "int")
        cpX := cpAnchorRight - cpPopupW
        cpY := cpAnchorBottom
    }
    cpVirtualX := SysGet(76), cpVirtualY := SysGet(77)
    cpVirtualW := SysGet(78), cpVirtualH := SysGet(79)
    cpX := Max(cpVirtualX, Min(cpX, cpVirtualX + cpVirtualW - cpPopupW))
    if (cpY + cpPopupH > cpVirtualY + cpVirtualH)
        cpY := Max(cpVirtualY, cpAnchorTop - cpPopupH)
    cpPopup.Show("x" cpX " y" cpY)
    CPThemedChoicePopupRegister(cpPopupState)
    CPThemedChoicePopupFocus(cpPopupState, 1)
    try WinActivate("ahk_id " cpPopup.Hwnd)
    cpWatch := CPThemedChoicePopupWatch.Bind(
        cpPopupState, cpPopup.Hwnd, cpPopup
    )
    SetTimer(cpWatch, 80)
    try {
        while !cpPopupState["closed"]
            Sleep(20)
    } finally {
        SetTimer(cpWatch, 0)
        CPThemedChoicePopupUnregister(cpPopupState)
        try cpPopup.Destroy()
    }
    return cpPopupState["result"]
}

CPThemedContextPopup(ownerHwnd, clientX, clientY, items) {
    global controlDarkMode
    if !items.Length
        return 0

    cpOwnerValid := ownerHwnd
        && DllCall("user32\IsWindow", "ptr", ownerHwnd, "int")
    cpPopupState := Map(
        "closed", false,
        "result", 0,
        "rows", [],
        "resultValues", [],
        "disabledRows", [],
        "separators", [],
        "focusIndex", 1
    )
    cpOptions := "+ToolWindow -Caption +Border +AlwaysOnTop"
    if cpOwnerValid
        cpOptions .= " +Owner" ownerHwnd
    cpPopup := Gui(cpOptions)
    cpPopup.MarginX := 4
    cpPopup.MarginY := 4
    cpPopup.SetFont("s9", "Segoe UI")
    cpColors := CPPalette(controlDarkMode)
    cpPopup.BackColor := cpColors["surface"]
    cpPopupState["gui"] := cpPopup
    cpPopupState["colors"] := cpColors

    cpLongestLabel := 0
    for cpItem in items {
        if cpItem.Has("label")
            cpLongestLabel := Max(cpLongestLabel, StrLen(cpItem["label"]))
    }
    cpMenuWidth := Max(220, Min(390, 42 + cpLongestLabel * 8))
    cpHasPreviousRow := false
    for cpItemIndex, cpItem in items {
        if (cpItem.Has("separator") && cpItem["separator"]) {
            cpSeparatorPosition := cpHasPreviousRow ? "xm y+4" : "xm ym"
            cpSeparator := cpPopup.Add(
                "Text",
                cpSeparatorPosition " w" cpMenuWidth " h1"
                    . " Background" cpColors["border"]
            )
            cpPopupState["separators"].Push(cpSeparator)
            cpHasPreviousRow := false
            continue
        }
        if !cpItem.Has("label")
            continue
        cpEnabled := !cpItem.Has("enabled") || cpItem["enabled"]
        cpPosition := cpHasPreviousRow ? "xm y+1" : "xm y+4"
        cpRowOptions := cpPosition " w" cpMenuWidth " h29 0x200"
            . " Background" cpColors["surface"]
            . " c" (cpEnabled ? cpColors["text"] : cpColors["muted"])
        if cpEnabled
            cpRowOptions .= " +Tabstop"
        cpChoiceRow := cpPopup.Add(
            "Text", cpRowOptions, "  " cpItem["label"]
        )
        if cpEnabled {
            cpPopupState["rows"].Push(cpChoiceRow)
            cpPopupState["resultValues"].Push(cpItemIndex)
            cpChoiceRow.OnEvent(
                "Click", CPThemedChoicePopupFinish.Bind(
                    cpPopupState, cpPopup, cpItemIndex
                )
            )
        } else
            cpPopupState["disabledRows"].Push(cpChoiceRow)
        cpHasPreviousRow := true
    }
    if !cpPopupState["rows"].Length
        return 0

    cpPopup.OnEvent(
        "Escape", CPThemedChoicePopupFinish.Bind(cpPopupState, cpPopup, 0)
    )
    cpPopup.OnEvent(
        "Close", CPThemedChoicePopupFinish.Bind(cpPopupState, cpPopup, 0)
    )
    cpPopup.Show("Hide AutoSize")
    CPApplyOwnedDialogTheme(cpPopup)
    for cpDisabledRow in cpPopupState["disabledRows"]
        try cpDisabledRow.Opt(
            "+Background" cpColors["surface"] " c" cpColors["muted"]
        )
    for cpSeparator in cpPopupState["separators"]
        try cpSeparator.Opt("+Background" cpColors["border"])
    CPThemedChoicePopupPaintFocus(cpPopupState, 1)
    cpPopup.GetPos(,, &cpPopupW, &cpPopupH)

    cpScreenPoint := Buffer(8, 0)
    NumPut("int", clientX, "int", clientY, cpScreenPoint, 0)
    if cpOwnerValid
        try DllCall(
            "user32\ClientToScreen", "ptr", ownerHwnd,
            "ptr", cpScreenPoint.Ptr, "int"
        )
    cpX := NumGet(cpScreenPoint, 0, "int")
    cpY := NumGet(cpScreenPoint, 4, "int")
    cpVirtualX := SysGet(76), cpVirtualY := SysGet(77)
    cpVirtualW := SysGet(78), cpVirtualH := SysGet(79)
    cpX := Max(cpVirtualX, Min(cpX, cpVirtualX + cpVirtualW - cpPopupW))
    cpY := Max(cpVirtualY, Min(cpY, cpVirtualY + cpVirtualH - cpPopupH))

    cpPopup.Show("x" cpX " y" cpY)
    CPThemedChoicePopupRegister(cpPopupState)
    CPThemedChoicePopupFocus(cpPopupState, 1)
    try WinActivate("ahk_id " cpPopup.Hwnd)
    cpWatch := CPThemedChoicePopupWatch.Bind(
        cpPopupState, cpPopup.Hwnd, cpPopup
    )
    SetTimer(cpWatch, 80)
    try {
        while !cpPopupState["closed"]
            Sleep(20)
    } finally {
        SetTimer(cpWatch, 0)
        CPThemedChoicePopupUnregister(cpPopupState)
        try cpPopup.Destroy()
    }
    return cpPopupState["result"]
}

CPThemedInputBox(promptText, dialogTitle, infoText := "", initialValue := ""
    , dialogWidth := 420, ownerHwnd := 0) {
    global ui
    result := {Result: "Cancel", Value: ""}
    closed := false
    contentWidth := Max(280, dialogWidth - 32)

    inputOwner := ownerHwnd
    if !(inputOwner && DllCall("user32\IsWindow", "ptr", inputOwner, "int"))
        inputOwner := ui.Hwnd
    dlg := Gui("+Owner" inputOwner " +AlwaysOnTop", dialogTitle)
    dlg.MarginX := 16
    dlg.MarginY := 14
    dlg.Add("Text", "xm w" contentWidth, promptText)
    if (infoText != "")
        dlg.Add("Text", "xm y+8 w" contentWidth, infoText)
    editor := dlg.Add("Edit", "xm y+12 w" contentWidth, initialValue)
    btnOK := dlg.Add("Button", "xm y+14 w120 Default", "OK")
    btnCancel := dlg.Add("Button", "x+10 yp w120", "Cancel")

    btnOK.OnEvent("Click", (*) => (
        result.Result := "OK",
        result.Value := editor.Value,
        closed := true,
        dlg.Destroy()
    ))
    btnCancel.OnEvent("Click", (*) => (closed := true, dlg.Destroy()))
    dlg.OnEvent("Escape", (*) => (closed := true, dlg.Destroy()))
    dlg.OnEvent("Close", (*) => (closed := true))

    dlg.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(dlg)
    try editor.Focus()
    while !closed
        Sleep(30)
    return result
}

CPTextEditorClose(dlg, *) {
    try dlg.Destroy()
}

CPShowTextEditorDialog(dlg, editorCtrl, ownerHwnd := 0) {
    global ui
    ; Ownership lets the controller-navigation poll recognize this editor as
    ; part of the control-panel workflow. B / Circle is dispatched as Escape,
    ; which uses the same close path as the editor's Close button.
    editorOwner := ownerHwnd
    if !(editorOwner && DllCall("user32\IsWindow", "ptr", editorOwner, "int"))
        editorOwner := ui.Hwnd
    try dlg.Opt("+Owner" editorOwner)
    dlg.OnEvent("Escape", CPTextEditorClose.Bind(dlg))
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

CPApplyDialogCheckBoxTheme(checkCtrl) {
    global controlDarkMode
    colors := CPPalette(controlDarkMode)
    try checkCtrl.SetFont("c" colors["text"])
    if controlDarkMode {
        ; Unthemed checkbox text honors WM_CTLCOLORBTN. The Explorer checkbox
        ; theme otherwise keeps a black caption on these owned dark dialogs.
        try DllCall("uxtheme\SetWindowTheme", "ptr", checkCtrl.Hwnd, "wstr", "", "wstr", "")
    } else {
        try DllCall("uxtheme\SetWindowTheme", "ptr", checkCtrl.Hwnd, "ptr", 0, "ptr", 0)
    }
    try DllCall("user32\InvalidateRect", "ptr", checkCtrl.Hwnd, "ptr", 0, "int", 1)
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

    selectedAudioProvider := (StrLower(audioProvider) = "gemini") ? "gemini" : "openai"
    if !CPApiKeyConfigured(selectedAudioProvider) {
        CPShowMissingApiKey(selectedAudioProvider, "translator")
        DbgCP("StartAudio blocked: " selectedAudioProvider " API key is missing")
        UpdateStatus()
        return false
    }

    EnvSet("AUDIO_PROVIDER", selectedAudioProvider)
    EnvSet("TEXT_PROVIDER",  selectedAudioProvider)
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
    }

    ; Run() normally provides the definitive PID. Query WMI only once as a
    ; recovery fallback for unusual launcher/process hand-off behavior.
    if !started {
        pids := AudioPidsByScript()
        if (pids.Length) {
            gPidAudio := pids[1]
            started := true
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
    static queryInProgress := false
    ap := ResolvePath(audioScript)
    apL := StrLower(StrReplace(ap, "/", "\"))   ; normalize & lower
    out := []
    if queryInProgress
        return out

    queryInProgress := true
    previousCritical := A_IsCritical
    Critical(50)
    try {
        wm := ComObjGet("winmgmts:")
        processes := wm.ExecQuery("Select ProcessId,CommandLine from Win32_Process Where Name='python.exe'")
        for p in processes {
            cmd := p.CommandLine ? p.CommandLine : ""
            cmdL := StrLower(StrReplace(cmd, "/", "\"))  ; normalize & lower
            if (InStr(cmdL, apL) && !InStr(cmdL, "--list-speakers"))
                out.Push(p.ProcessId)
        }
    } catch as ex {
        DbgCP("AudioPidsByScript WMI query failed: " ex.Message)
    } finally {
        queryInProgress := false
        Critical(previousCritical)
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

    ; Capture the known PID and perform at most one recovery scan. The wait loop
    ; below checks these PIDs directly instead of repeatedly querying WMI.
    knownPids := Map()
    if (gPidAudio && ProcessExist(gPidAudio))
        knownPids[gPidAudio] := true
    for pid in AudioPidsByScript() {
        knownPids[pid] := true
    }
    for pid, _ in knownPids
        try ProcessClose(pid)

    stopped := false
    Loop 20 {
        remainingAlive := false
        for pid, _ in knownPids {
            if ProcessExist(pid) {
                remainingAlive := true
                try ProcessClose(pid)
            }
        }
        if !remainingAlive {
            stopped := true
            break
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

ModelPickerControlAlive(control) {
    try {
        hwnd := control.Hwnd
        return hwnd && DllCall("user32\IsWindow", "ptr", hwnd, "int")
    } catch {
        return false
    }
}

ModelPickerControlsAlive(controls*) {
    for control in controls
        if !ModelPickerControlAlive(control)
            return false
    return true
}

ModelPickerFinish(pickerDialog, pickerState, selection := "", *) {
    if pickerState["closed"]
        return
    pickerState["result"] := selection
    pickerState["closed"] := true
    try pickerDialog.Destroy()
}

ModelPickerPopulate(modelListBox, modelStatus, modelAddButton, catalog, existingModels) {
    if !ModelPickerControlsAlive(modelListBox, modelStatus, modelAddButton)
        return 0
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
    if !ModelPickerControlAlive(modelListBox)
        return
    chosenModel := Trim(modelListBox.Text)
    if (chosenModel = "") {
        SoundBeep(1100, 80)
        return
    }
    finishPicker.Call(chosenModel)
}

ModelPickerActivate(modelListBox, modelAddButton, refreshButton, cancelButton, finishPicker, *) {
    if !ModelPickerControlsAlive(modelListBox, modelAddButton, refreshButton, cancelButton)
        return
    focusHwnd := DllCall("user32\GetFocus", "ptr")
    if (focusHwnd = modelListBox.Hwnd || focusHwnd = modelAddButton.Hwnd) {
        if modelAddButton.Enabled
            ModelPickerAccept(modelListBox, finishPicker)
    } else if (focusHwnd = refreshButton.Hwnd) {
        SendMessage(0x00F5, 0, 0, refreshButton.Hwnd) ; BM_CLICK
    } else if (focusHwnd = cancelButton.Hwnd) {
        SendMessage(0x00F5, 0, 0, cancelButton.Hwnd)
    }
}

ModelPickerNavigate(direction, modelListBox, modelAddButton, refreshButton, cancelButton, *) {
    if !ModelPickerControlsAlive(modelListBox, modelAddButton, refreshButton, cancelButton)
        return
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

ModelPickerRefresh(provider, purpose, existingModels, pickerDialog, modelListBox, modelStatus, modelAddButton, refreshButton, pickerState, *) {
    if pickerState["closed"] || pickerState["refreshing"]
        return
    if !ModelPickerControlsAlive(pickerDialog, modelListBox, modelStatus, modelAddButton, refreshButton)
        return

    pickerState["refreshing"] := true
    try {
        previousStatus := modelStatus.Text
        modelListBox.Enabled := false
        modelAddButton.Enabled := false
        refreshButton.Enabled := false
        modelStatus.Text := "Refreshing the online model catalog..."

        refreshedCatalog := ModelCatalogQueryWithFeedback(provider, purpose, true)

        ; The network query yields to the GUI. The user may close the picker while
        ; it is running, so never resume against controls that no longer exist.
        if pickerState["closed"]
            return
        if !ModelPickerControlsAlive(pickerDialog, modelListBox, modelStatus, modelAddButton, refreshButton)
            return

        refreshButton.Enabled := true
        if !refreshedCatalog["ok"] {
            modelStatus.Text := previousStatus
            modelListBox.Enabled := true
            modelAddButton.Enabled := Trim(modelListBox.Text) != ""
            MsgBox("The online model list could not be refreshed.`n`n" refreshedCatalog["error"], "Browse models", 48)
            return
        }

        ModelPickerPopulate(modelListBox, modelStatus, modelAddButton, refreshedCatalog, existingModels)
        if pickerState["closed"] || !ModelPickerControlAlive(pickerDialog)
            return
        CPApplyOwnedDialogTheme(pickerDialog)
        if !ModelPickerControlsAlive(modelListBox, refreshButton)
            return
        if modelListBox.Enabled
            modelListBox.Focus()
        else
            refreshButton.Focus()
    } finally {
        pickerState["refreshing"] := false
    }
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
    pickerState := Map("closed", false, "refreshing", false, "result", "")
    pickerDialog := Gui("+Owner" ui.Hwnd " +AlwaysOnTop", "Browse " providerName " models")
    pickerDialog.MarginX := 18, pickerDialog.MarginY := 16
    pickerDialog.SetFont("s10", "Segoe UI")
    pickerDialog.Add("Text", "xm w560", "Select one model to add to " purposeLabel ".")
    pickerDialog.Add("Text", "xm y+5 w560", "The list contains compatible models available to the configured API key.")
    modelListBox := pickerDialog.Add("ListBox", "xm y+12 w560 r16")
    modelStatus := pickerDialog.Add("Text", "xm y+8 w560 h42")
    modelAddButton := pickerDialog.Add("Button", "xm y+12 w120", "Add model")
    refreshButton := pickerDialog.Add("Button", "x+8 w100", "Refresh")
    cancelButton := pickerDialog.Add("Button", "x+8 w100", "Cancel")

    finishPicker := ModelPickerFinish.Bind(pickerDialog, pickerState)
    modelAddButton.OnEvent("Click", ModelPickerAccept.Bind(modelListBox, finishPicker))
    modelListBox.OnEvent("DoubleClick", ModelPickerAccept.Bind(modelListBox, finishPicker))
    refreshButton.OnEvent("Click", ModelPickerRefresh.Bind(provider, purpose, existingModels, pickerDialog, modelListBox, modelStatus, modelAddButton, refreshButton, pickerState))
    cancelButton.OnEvent("Click", (*) => finishPicker.Call(""))
    pickerDialog.OnEvent("Escape", (*) => finishPicker.Call(""))
    pickerDialog.OnEvent("Close", (*) => finishPicker.Call(""))

    ModelPickerPopulate(modelListBox, modelStatus, modelAddButton, catalog, existingModels)
    pickerHotIf := "ahk_id " pickerDialog.Hwnd
    pickerArrowHotkeys := Map("$Up", "Up", "$Down", "Down", "$Left", "Left", "$Right", "Right")
    HotIfWinActive(pickerHotIf)
    for keyName, direction in pickerArrowHotkeys
        try Hotkey(keyName, ModelPickerNavigate.Bind(direction, modelListBox, modelAddButton, refreshButton, cancelButton), "On")
    try Hotkey("$Enter", ModelPickerActivate.Bind(modelListBox, modelAddButton, refreshButton, cancelButton, finishPicker), "On")
    try Hotkey("$NumpadEnter", ModelPickerActivate.Bind(modelListBox, modelAddButton, refreshButton, cancelButton, finishPicker), "On")
    HotIfWinActive()

    try {
        pickerDialog.Show("AutoSize Center")
        CPApplyOwnedDialogTheme(pickerDialog)
        if modelListBox.Enabled
            modelListBox.Focus()
        else
            refreshButton.Focus()
        while !pickerState["closed"]
            Sleep(30)
    } finally {
        HotIfWinActive(pickerHotIf)
        for keyName, direction in pickerArrowHotkeys
            try Hotkey(keyName, "Off")
        try Hotkey("$Enter", "Off")
        try Hotkey("$NumpadEnter", "Off")
        HotIfWinActive()
    }
    return pickerState["result"]
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
    imgPostproc := SyncPromptPostproc(promptProfile)
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

        DbgCP("LaunchOverlay Run EXCEPTION: " ex.Message "  cmd=" cmd " wd=" ovDir)
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
            DbgCP("WinSetTransparent failed on 'Translator': " ex.Message)
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

    prov := CPSyncExplanationSelectionFromControls()

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
    DbgCP(
        "LaunchExplainerOverlay run='" ov "' provider=" prov " model="
            . (prov = "gemini" ? explainGeminiModel : explainOpenAIModel)
    )

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
            DbgCP("WinSetTransparent failed: " ex.Message)
        }
        ; start periodic watcher to persist movement/resizes
        StartExplainerBoundsWatcher()
    }
    SetTitleMatchMode oldMode

    SendOverlayTheme("Explainer")
}

ToastDestroy(cpToastGui) {
    cpToastHwnd := 0
    try cpToastHwnd := IsObject(cpToastGui) ? cpToastGui.Hwnd : 0
    if cpToastHwnd && DllCall("user32\IsWindow", "ptr", cpToastHwnd, "int")
        try cpToastGui.Destroy()
}

Toast(msg){
    static g := 0
    ToastDestroy(g)
    g := Gui("+ToolWindow -Caption +AlwaysOnTop +E0x20")
    g.BackColor := "101825"
    g.MarginX := 12, g.MarginY := 8
    g.SetFont("s11", "Segoe UI")
    g.Add("Text", "cWhite", msg)
    g.Show("AutoSize NoActivate x20 y20")
    SetTimer(ToastDestroy.Bind(g), -1100)
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

CPControllerModalFlagActive() {
    global CP_OVERLAY_ADJUST_FLAG
    if !FileExist(CP_OVERLAY_ADJUST_FLAG)
        return false
    ownerPid := 0
    try ownerPid := Integer(Trim(FileRead(CP_OVERLAY_ADJUST_FLAG, "UTF-8")))
    if (ownerPid && ProcessExist(ownerPid))
        return true
    try FileDelete(CP_OVERLAY_ADJUST_FLAG)
    return false
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

CPJoystickAxis(controllerId, axisName, snapshot := false, axisRanges := false) {
    if !IsObject(snapshot)
        snapshot := CPJoystickNativeSnapshot(controllerId)
    if !(IsObject(snapshot) && snapshot.Has("axes") && snapshot["axes"].Has(axisName))
        return ""
    if !IsObject(axisRanges)
        axisRanges := CPJoystickNativeAxisRanges(controllerId)
    if !(IsObject(axisRanges) && axisRanges.Has(axisName))
        return ""

    axisRange := axisRanges[axisName]
    minimum := axisRange[1] + 0.0
    maximum := axisRange[2] + 0.0
    if (maximum <= minimum)
        return 50.0
    normalized := (snapshot["axes"][axisName] - minimum) * 100.0 / (maximum - minimum)
    return Max(0.0, Min(100.0, normalized))
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
        "packet", NumGet(stateBuffer, 0, "uint"),
        "buttons", NumGet(stateBuffer, 4, "ushort"),
        "leftTrigger", NumGet(stateBuffer, 6, "uchar"),
        "rightTrigger", NumGet(stateBuffer, 7, "uchar"),
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
        legacyState := CPJoystickNativeSnapshot(controllerId)
        if !IsObject(legacyState) {
            CPJoystickNativeName(controllerId, true)
            continue
        }
        controllerName := CPJoystickNativeName(controllerId)

        rightAxes := CPControllerRightAxes(controllerName)
        axisRanges := CPJoystickNativeAxisRanges(controllerId)
        baseline := Map()
        for axisName in ["X", "Y", "Z", "R", "U", "V"] {
            value := CPJoystickAxis(controllerId, axisName, legacyState, axisRanges)
            baseline[axisName] := (value = "") ? 50.0 : value
        }
        controllers.Push(Map(
            "type", "legacy",
            "id", controllerId,
            "name", controllerName,
            "rightX", rightAxes[1],
            "rightY", rightAxes[2],
            "axisRanges", axisRanges,
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

    current := CPJoystickNativeSnapshot(controller["id"])
    if !IsObject(current)
        return false
    return Map(
        "moveX", CPOverlayAdjustAxis(controller, "X", current),
        "moveY", CPOverlayAdjustAxis(controller, "Y", current),
        "sizeX", CPOverlayAdjustAxis(controller, controller["rightX"], current),
        "sizeY", CPOverlayAdjustAxis(controller, controller["rightY"], current)
    )
}

CPReadOverlayAdjustButtons(controller) {
    if (controller["type"] = "xinput") {
        currentState := CPXInputGetState(controller["id"])
        if !IsObject(currentState)
            return false
        buttonBits := currentState["buttons"]
        return Map(
            "confirm", !!(buttonBits & 0x1000),
            "cancel", !!(buttonBits & 0x2000)
        )
    }

    controllerName := controller["name"]
    isPlayStationController := RegExMatch(controllerName
        , "i)(DualSense|DualShock|Wireless Controller|PlayStation)")
    confirmButtonIndex := isPlayStationController ? 2 : 1
    cancelButtonIndex := isPlayStationController ? 3 : 2
    currentState := CPJoystickNativeSnapshot(controller["id"])
    if !IsObject(currentState)
        return false
    buttonBits := currentState["buttons"]
    return Map(
        "confirm", !!(buttonBits & (1 << (confirmButtonIndex - 1))),
        "cancel", !!(buttonBits & (1 << (cancelButtonIndex - 1)))
    )
}

CPOverlayAdjustControllerKey(controller) {
    return controller["type"] ":" controller["id"]
}

CPInitializeOverlayAdjustButtonStates() {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    buttonStates := Map()
    for candidateController in state["controllers"] {
        candidateButtons := CPReadOverlayAdjustButtons(candidateController)
        if IsObject(candidateButtons)
            buttonStates[CPOverlayAdjustControllerKey(candidateController)] := candidateButtons
    }
    state["controllerButtonStates"] := buttonStates
}

CPHandlePendingOverlayAdjustControllerButtons() {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("controllerButtonStates") && IsObject(state["controllerButtonStates"]))
        state["controllerButtonStates"] := Map()
    buttonStates := state["controllerButtonStates"]

    for candidateController in state["controllers"] {
        candidateKey := CPOverlayAdjustControllerKey(candidateController)
        currentButtons := CPReadOverlayAdjustButtons(candidateController)
        if !IsObject(currentButtons)
            continue
        if !buttonStates.Has(candidateKey) {
            buttonStates[candidateKey] := currentButtons
            continue
        }

        previousButtons := buttonStates[candidateKey]
        buttonStates[candidateKey] := currentButtons
        if (currentButtons["cancel"] && !previousButtons["cancel"]) {
            CPOverlayAdjustCancel()
            return true
        }
        if (A_TickCount >= state["acceptAfter"]
            && currentButtons["confirm"] && !previousButtons["confirm"]) {
            CPOverlayAdjustConfirm()
            return true
        }
    }
    return false
}

CPHandleOverlayAdjustControllerButtons(controller) {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    currentButtons := CPReadOverlayAdjustButtons(controller)
    if !IsObject(currentButtons)
        return false

    if !(state.Has("controllerButtons") && IsObject(state["controllerButtons"])) {
        state["controllerButtons"] := currentButtons
        return false
    }

    previousButtons := state["controllerButtons"]
    state["controllerButtons"] := currentButtons
    if (currentButtons["cancel"] && !previousButtons["cancel"]) {
        CPOverlayAdjustCancel()
        return true
    }
    if (A_TickCount >= state["acceptAfter"]
        && currentButtons["confirm"] && !previousButtons["confirm"]) {
        CPOverlayAdjustConfirm()
        return true
    }
    return false
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
    controllerKey := CPOverlayAdjustControllerKey(bestController)
    initialButtons := (state.Has("controllerButtonStates")
        && state["controllerButtonStates"].Has(controllerKey))
        ? state["controllerButtonStates"][controllerKey]
        : CPReadOverlayAdjustButtons(bestController)
    state["controllerButtons"] := IsObject(initialButtons)
        ? initialButtons : Map("confirm", false, "cancel", false)
    CPUpdateOverlayAdjustHud(true)
    return true
}

CPOverlayAdjustAxis(controller, axisName, snapshot := false) {
    value := CPJoystickAxis(controller["id"], axisName, snapshot, controller["axisRanges"])
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
    hudText := hud.Add("Text", "w540 h78 Center", "")
    state["hud"] := hud
    state["hudText"] := hudText
    state["hudReady"] := false
    CPUpdateOverlayAdjustHud(true)
    hud.Show("NA AutoSize")
    state["hudReady"] := true
    CPPositionOverlayAdjustHud()
}

CPPositionOverlayAdjustHud() {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("active") && state["active"] && state.Has("hud") && state["hud"]
        && state.Has("hudReady") && state["hudReady"])
        return

    hudHwnd := 0
    try hudHwnd := state["hud"].Hwnd
    if !(hudHwnd && DllCall("user32\IsWindow", "ptr", hudHwnd, "int"))
        return

    workLeft := 0, workTop := 0, workRight := A_ScreenWidth, workBottom := A_ScreenHeight
    monitorHandle := DllCall("user32\MonitorFromWindow", "ptr", state["hwnd"]
        , "uint", 2, "ptr")
    monitorInfo := Buffer(40, 0)
    NumPut("uint", 40, monitorInfo, 0)
    if (monitorHandle && DllCall("user32\GetMonitorInfoW", "ptr", monitorHandle
        , "ptr", monitorInfo.Ptr, "int")) {
        workLeft := NumGet(monitorInfo, 20, "int")
        workTop := NumGet(monitorInfo, 24, "int")
        workRight := NumGet(monitorInfo, 28, "int")
        workBottom := NumGet(monitorInfo, 32, "int")
    }

    windowRect := Buffer(16, 0)
    if !DllCall("user32\GetWindowRect", "ptr", hudHwnd, "ptr", windowRect.Ptr, "int")
        return
    hudW := NumGet(windowRect, 8, "int") - NumGet(windowRect, 0, "int")
    hudH := NumGet(windowRect, 12, "int") - NumGet(windowRect, 4, "int")
    if (hudW <= 0 || hudH <= 0)
        return
    margin := 12
    hudX := Round(workLeft + (workRight - workLeft - hudW) / 2)
    hudY := workTop + margin
    maximumX := workRight - hudW - margin
    maximumY := workBottom - hudH - margin
    hudX := (maximumX >= workLeft + margin)
        ? Max(workLeft + margin, Min(hudX, maximumX)) : workLeft
    hudY := (maximumY >= workTop + margin)
        ? Max(workTop + margin, Min(hudY, maximumY)) : workTop
    DllCall("user32\SetWindowPos", "ptr", hudHwnd, "ptr", 0
        , "int", hudX, "int", hudY, "int", 0, "int", 0
        , "uint", 0x0015)
}

CPUpdateOverlayAdjustHud(force := false) {
    global CPOverlayAdjustState
    state := CPOverlayAdjustState
    if !(state.Has("active") && state["active"] && state.Has("hudText") && state["hudText"])
        return
    hudTextControl := state["hudText"]
    hudTextHwnd := 0
    try hudTextHwnd := hudTextControl.Hwnd
    if !(hudTextHwnd && DllCall("user32\IsWindow", "ptr", hudTextHwnd, "int"))
        return
    now := A_TickCount
    if (!force && state.Has("lastHudTick") && now - state["lastHudTick"] < 100)
        return
    state["lastHudTick"] := now

    controllerLine := "Move either analog stick to select a controller"
    if (state.Has("controller") && IsObject(state["controller"]))
        controllerLine := state["controller"]["name"]
    try hudTextControl.Text := "Adjusting " state["title"] " | " controllerLine
        . "`nLeft stick or arrows: move | Right stick or hold Screenshot + Translate + arrows: resize"
        . "`nA / Cross / Enter saves | B / Circle / Esc cancels | " state["w"] " x " state["h"]
    catch
        return
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
        if CPHandlePendingOverlayAdjustControllerButtons()
            return
        CPDetectOverlayAdjustController()
        CPUpdateOverlayAdjustHud()
        return
    }

    controller := state["controller"]
    if CPHandleOverlayAdjustControllerButtons(controller)
        return
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
    CPInitializeOverlayAdjustButtonStates()
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
    state["active"] := false
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
    try state.Delete("hudText")
    try state.Delete("hud")
    try state.Delete("hudReady")
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
; Unified profiles
; =========================
GameProfileSafeName(name) {
    name := Trim(name)
    name := RegExReplace(name, '[\\/:*?"<>|]+', "_")
    name := RegExReplace(name, "[\. ]+$", "")
    if RegExMatch(name, 'i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$')
        name := "_" name
    return name
}

GameProfilePath(name) {
    global gameProfilesDir
    return gameProfilesDir "\" GameProfileSafeName(name) ".ini"
}

ListGameProfiles() {
    global gameProfilesDir
    out := []
    if DirExist(gameProfilesDir) {
        Loop Files gameProfilesDir "\*.ini", "F"
            out.Push(RegExReplace(A_LoopFileName, "i)\.ini$", ""))
    }
    if out.Length {
        textList := ""
        for item in out
            textList .= item "`n"
        textList := Sort(RTrim(textList, "`n"))
        out := StrSplit(textList, "`n")
    }
    return out
}

GameProfileBoundsPath(title) {
    global appDir
    return appDir "\overlay_" (title = "Explainer" ? "explainer" : "translator") ".ini"
}

GameProfileSnapshotBounds(title) {
    if CPFindExactWindow(title) {
        try SendOverlayCmdTo(title, "action=save_bounds")
        Sleep(120)
    }

    path := GameProfileBoundsPath(title)
    bounds := Map()
    for key in ["x", "y", "w", "h", "dpi"] {
        value := IniRead(path, "win", key, "")
        if (value != "" && IsNumber(value))
            bounds[key] := Integer(value)
    }
    return bounds
}

GameProfileWriteBounds(profilePath, section, bounds) {
    for key in ["x", "y", "w", "h", "dpi"] {
        if bounds.Has(key)
            IniWriteRetry(bounds[key], profilePath, section, key)
    }
}

GameProfileReadInt(path, section, key, fallback, minValue := "", maxValue := "") {
    value := IniRead(path, section, key, fallback)
    if !IsNumber(value)
        value := fallback
    value := Integer(value)
    if (minValue != "")
        value := Max(value, minValue)
    if (maxValue != "")
        value := Min(value, maxValue)
    return value
}

GameProfileReadColor(path, section, key, fallback) {
    value := StrUpper(Trim(IniRead(path, section, key, fallback)))
    return RegExMatch(value, "i)^[0-9a-f]{6}$") ? value : fallback
}

GameProfileAppendWarning(&warnings, message) {
    warnings .= (warnings = "" ? "" : "`n") "- " message
}

GameProfileSave(name, announce := true) {
    global iniPath
    global CPControllerDpadNavigationEnabled
    global useTerminologyOverrides
    global ddlPrompt, ddlEPr, ddlJPG, ddlENG
    global chkGuess, chkName, chkOpenTW, chkTop_TW, chkOpenEW, chkTop_EW
    global overlayTrans, boxBgHex, txtHex, nameHex, fontName, fontSize, fontBold
    global overlayTrans_EW, boxBgHex_EW, txtHex_EW, fontName_EW, fontSize_EW, fontBold_EW

    name := GameProfileSafeName(name)
    if (name = "") {
        MsgBox("Please enter a non-empty profile name.", "Profiles", "OK Icon!")
        return false
    }

    path := GameProfilePath(name)
    try {
        if FileExist(path)
            FileCopy(path, path ".bak", true)

        promptName := Trim(ddlPrompt.Text)
        explainPromptName := Trim(ddlEPr.Text)
        jpProfile := Trim(ddlJPG.Text)
        tlProfile := Trim(ddlENG.Text)

        IniWriteRetry(1, path, "profile", "schemaVersion")
        IniWriteRetry(name, path, "profile", "name")
        IniWriteRetry(FormatTime(, "yyyy-MM-dd HH:mm:ss"), path, "profile", "updated")

        IniWriteRetry(promptName, path, "screenshot", "promptProfile")
        IniWriteRetry(IniRead(iniPath, "capture", "mode", "region"), path, "screenshot", "captureMode")
        IniWriteRetry(IniRead(iniPath, "capture", "rect", ""), path, "screenshot", "captureRect")
        IniWriteRetry(IniRead(iniPath, "capture", "winTitle", ""), path, "screenshot", "captureWindowTitle")
        IniWriteRetry(chkGuess.Value ? 1 : 0, path, "screenshot", "highlightGuessed")
        IniWriteRetry(chkName.Value ? 1 : 0, path, "screenshot", "colorSpeaker")

        IniWriteRetry(explainPromptName, path, "explanation", "promptProfile")
        IniWriteRetry(
            StudyLibraryConfiguredName(), path, "study_library", "name"
        )
        ankiMapping := StudyAnkiLoadMapping(name)
        if (ankiMapping["deck"] != "") {
            IniWriteRetry(ankiMapping["deck"], path, "anki", "deck")
            IniWriteRetry(
                ankiMapping.Has("addDeck") ? ankiMapping["addDeck"] : "",
                path, "anki", "addDeck"
            )
            IniWriteRetry(
                ankiMapping.Has("vocabularyDeck")
                    ? ankiMapping["vocabularyDeck"] : "",
                path, "anki", "vocabularyDeck"
            )
            IniWriteRetry(ankiMapping["model"], path, "anki", "model")
            IniWriteRetry(
                ankiMapping["japaneseField"], path,
                "anki", "japaneseField"
            )
            IniWriteRetry(
                ankiMapping["explanationField"], path,
                "anki", "explanationField"
            )
        }
        IniWriteRetry(useTerminologyOverrides ? 1 : 0, path, "terminology", "enabled")
        IniWriteRetry(jpProfile, path, "terminology", "jp2tlProfile")
        IniWriteRetry(tlProfile, path, "terminology", "tl2tlProfile")

        IniWriteRetry(overlayTrans, path, "translator", "overlayTrans")
        IniWriteRetry(boxBgHex, path, "translator", "boxBg")
        IniWriteRetry(txtHex, path, "translator", "txtColor")
        IniWriteRetry(nameHex, path, "translator", "nameColor")
        IniWriteRetry(fontName, path, "translator", "fontName")
        IniWriteRetry(fontSize, path, "translator", "fontSize")
        IniWriteRetry(fontBold ? 1 : 0, path, "translator", "fontBold")
        IniWriteRetry(chkOpenTW.Value ? 1 : 0, path, "translator", "openOnLaunch")
        IniWriteRetry(chkTop_TW.Value ? 1 : 0, path, "translator", "alwaysOnTop")
        GameProfileWriteBounds(path, "translator", GameProfileSnapshotBounds("Translator"))

        IniWriteRetry(overlayTrans_EW, path, "explainer", "overlayTrans")
        IniWriteRetry(boxBgHex_EW, path, "explainer", "boxBg")
        IniWriteRetry(txtHex_EW, path, "explainer", "txtColor")
        IniWriteRetry(fontName_EW, path, "explainer", "fontName")
        IniWriteRetry(fontSize_EW, path, "explainer", "fontSize")
        IniWriteRetry(fontBold_EW ? 1 : 0, path, "explainer", "fontBold")
        IniWriteRetry(chkOpenEW.Value ? 1 : 0, path, "explainer", "openOnLaunch")
        IniWriteRetry(chkTop_EW.Value ? 1 : 0, path, "explainer", "alwaysOnTop")
        GameProfileWriteBounds(path, "explainer", GameProfileSnapshotBounds("Explainer"))

        for key in ["x", "y", "w", "h"] {
            value := IniRead(iniPath, "explainer_bounds", key, "")
            if (value != "" && IsNumber(value))
                IniWriteRetry(Integer(value), path, "explainer_control_bounds", key)
        }

        IniWriteRetry(CPControllerDpadNavigationEnabled ? 1 : 0, path, "controls", "dpadNavigation")
        IniWriteRetry(name, iniPath, "game_profiles", "active")
        DbgCP("Saved unified profile: " name)
        if announce
            Toast("Profile saved: " name)
        return true
    } catch as ex {
        MsgBox("Could not save profile:`n" ex.Message, "Profiles", "OK Iconx")
        return false
    }
}

GameProfileReadBounds(profilePath, section) {
    values := Map()
    for key in ["x", "y", "w", "h", "dpi"] {
        value := IniRead(profilePath, section, key, "")
        if (value != "" && IsNumber(value))
            values[key] := Integer(value)
    }
    if !(values.Has("x") && values.Has("y") && values.Has("w") && values.Has("h"))
        return 0

    if !values.Has("dpi")
        values["dpi"] := 96
    return values
}

GameProfileEnsureBoundsVisible(values) {
    if !IsObject(values)
        return values

    visibleMargin := 64
    dpi := Max(values.Has("dpi") ? values["dpi"] : 96, 96)
    physicalW := Max(Round(values["w"] * dpi / 96), visibleMargin)
    physicalH := Max(Round(values["h"] * dpi / 96), visibleMargin)
    isVisible := false
    Loop MonitorGetCount() {
        MonitorGetWorkArea(A_Index, &left, &top, &right, &bottom)
        if (values["x"] < right - visibleMargin && values["x"] + physicalW > left + visibleMargin
         && values["y"] < bottom - visibleMargin && values["y"] + physicalH > top + visibleMargin) {
            isVisible := true
            break
        }
    }
    if !isVisible {
        DbgCP("Profile bounds were off-screen; resetting origin to 120,120.")
        values["x"] := 120
        values["y"] := 120
    }
    return values
}

GameProfileApplyBounds(title, profilePath, section) {
    ini := GameProfileBoundsPath(title)
    values := GameProfileEnsureBoundsVisible(GameProfileReadBounds(profilePath, section))
    if !IsObject(values)
        return 0

    for key, value in values
        IniWriteRetry(value, ini, "win", key)

    command := "x=" values["x"] "|y=" values["y"] "|w=" values["w"] "|h=" values["h"] "|dpi=" values["dpi"]
    try SendOverlayCmdTo(title, command)
    return values
}

GameProfileApply(name, announce := true) {
    global iniPath
    global promptProfile, explainPromptProfile, imgPostproc, jp2enGlossaryProfile, en2enGlossaryProfile
    global useTerminologyOverrides
    global capMode, capRect, capWinInfo
    global overlayTrans, boxBgHex, txtHex, nameHex, fontName, fontSize, fontBold
    global overlayTrans_EW, boxBgHex_EW, txtHex_EW, fontName_EW, fontSize_EW, fontBold_EW
    global ddlPrompt, ddlEPr, ddlJPG, ddlENG, chkUseTerminologyOverrides
    global chkGuess, chkName, chkOpenTW, chkTop_TW, chkOpenEW, chkTop_EW
    global slTrans, lblTransPct, rectBg, rectTxt, rectName, ddlFont, edFSize, udFSize, chkFontBold
    global slTrans_EW, lblTransPct_EW, rectBg_EW, rectTxt_EW, ddlFont_EW, edFSize_EW, udFSize_EW, chkFontBold_EW
    global promptsDir, explainPromptsDir
    global ewX, ewY, ewW, ewH, ew_lastX, ew_lastY, ew_lastW, ew_lastH, ew_bounds_watch_running

    path := GameProfilePath(name)
    if !FileExist(path) {
        if announce
            MsgBox("Profile not found:`n" path, "Profiles", "OK Icon!")
        else
            DbgCP("Requested Profile was not found: " path)
        return false
    }
    if (GameProfileReadInt(path, "profile", "schemaVersion", 0) != 1) {
        if announce
            MsgBox("This profile uses an unsupported format.", "Profiles", "OK Icon!")
        else
            DbgCP("Requested Profile uses an unsupported format: " path)
        return false
    }

    warnings := ""
    candidate := Trim(IniRead(path, "screenshot", "promptProfile", promptProfile))
    if (candidate != "" && FileExist(promptsDir "\" candidate ".txt"))
        promptProfile := candidate
    else
        GameProfileAppendWarning(&warnings, "Screenshot prompt '" candidate "' was not found; the current prompt was kept.")

    candidate := Trim(IniRead(path, "explanation", "promptProfile", explainPromptProfile))
    if (candidate != "" && FileExist(explainPromptsDir "\" candidate ".txt"))
        explainPromptProfile := candidate
    else
        GameProfileAppendWarning(&warnings, "Explanation prompt '" candidate "' was not found; the current prompt was kept.")

    jpGlossaryList := ListGlossaryProfiles("jp")
    candidate := Trim(IniRead(path, "terminology", "jp2tlProfile", jp2enGlossaryProfile))
    if ArrHas(jpGlossaryList, candidate)
        jp2enGlossaryProfile := candidate
    else
        GameProfileAppendWarning(&warnings, "JP -> TL glossary '" candidate "' was not found; the current selection was kept.")
    enGlossaryList := ListGlossaryProfiles("en")
    candidate := Trim(IniRead(path, "terminology", "tl2tlProfile", en2enGlossaryProfile))
    if ArrHas(enGlossaryList, candidate)
        en2enGlossaryProfile := candidate
    else
        GameProfileAppendWarning(&warnings, "TL -> TL glossary '" candidate "' was not found; the current selection was kept.")
    ; Profiles created before this toggle existed used terminology overrides
    ; unconditionally, so a missing key intentionally falls back to enabled.
    useTerminologyOverrides := GameProfileReadInt(path, "terminology", "enabled", 1) ? 1 : 0

    ; Profiles saved before Study Libraries existed intentionally have no key.
    ; In that case, preserve the user's current library selection.
    profileLibraryMissing := "__JRPG_MISSING_STUDY_LIBRARY__"
    profileLibrary := IniRead(
        path, "study_library", "name", profileLibraryMissing
    )
    if (profileLibrary != profileLibraryMissing)
        StudyLibraryApplyProfileSelection(profileLibrary, &warnings)

    profileAnkiDeck := Trim(IniRead(path, "anki", "deck", ""))
    if (profileAnkiDeck != "") {
        profileAnkiMapping := Map(
            "profile", name,
            "deck", profileAnkiDeck,
            "addDeck", Trim(IniRead(path, "anki", "addDeck", "")),
            "vocabularyDeck", Trim(
                IniRead(path, "anki", "vocabularyDeck", "")
            ),
            "model", Trim(IniRead(path, "anki", "model", "Basic")),
            "japaneseField", Trim(
                IniRead(path, "anki", "japaneseField", "Front")
            ),
            "explanationField", Trim(
                IniRead(path, "anki", "explanationField", "Back")
            )
        )
        StudyAnkiSaveMapping(profileAnkiMapping, false)
    }

    capMode := StrLower(Trim(IniRead(path, "screenshot", "captureMode", capMode)))
    if (capMode != "window")
        capMode := "region"
    capRect := Trim(IniRead(path, "screenshot", "captureRect", capRect))
    capWinInfo := IniRead(path, "screenshot", "captureWindowTitle", capWinInfo)

    overlayTrans := GameProfileReadInt(path, "translator", "overlayTrans", overlayTrans, 0, 255)
    boxBgHex := GameProfileReadColor(path, "translator", "boxBg", boxBgHex)
    txtHex := GameProfileReadColor(path, "translator", "txtColor", txtHex)
    nameHex := GameProfileReadColor(path, "translator", "nameColor", nameHex)
    fontName := IniRead(path, "translator", "fontName", fontName)
    fontSize := GameProfileReadInt(path, "translator", "fontSize", fontSize, 6, 128)
    fontBold := GameProfileReadInt(path, "translator", "fontBold", fontBold) ? 1 : 0

    overlayTrans_EW := GameProfileReadInt(path, "explainer", "overlayTrans", overlayTrans_EW, 0, 255)
    boxBgHex_EW := GameProfileReadColor(path, "explainer", "boxBg", boxBgHex_EW)
    txtHex_EW := GameProfileReadColor(path, "explainer", "txtColor", txtHex_EW)
    fontName_EW := IniRead(path, "explainer", "fontName", fontName_EW)
    fontSize_EW := GameProfileReadInt(path, "explainer", "fontSize", fontSize_EW, 6, 200)
    fontBold_EW := GameProfileReadInt(path, "explainer", "fontBold", fontBold_EW) ? 1 : 0
    SyncUnifiedWindowAppearance()

    RefreshPromptProfilesList(promptProfile)
    imgPostproc := SyncPromptPostproc(promptProfile)
    RefreshExplainPromptProfilesList(explainPromptProfile)
    RefreshGlossaryProfilesList(jp2enGlossaryProfile, en2enGlossaryProfile)
    chkUseTerminologyOverrides.Value := useTerminologyOverrides
    EnvSet("USE_TERMINOLOGY_OVERRIDES", useTerminologyOverrides ? "1" : "0")

    chkGuess.Value := GameProfileReadInt(path, "screenshot", "highlightGuessed", chkGuess.Value) ? 1 : 0
    chkName.Value := GameProfileReadInt(path, "screenshot", "colorSpeaker", chkName.Value) ? 1 : 0
    chkOpenTW.Value := GameProfileReadInt(path, "translator", "openOnLaunch", chkOpenTW.Value) ? 1 : 0
    chkTop_TW.Value := GameProfileReadInt(path, "translator", "alwaysOnTop", chkTop_TW.Value) ? 1 : 0
    chkOpenEW.Value := GameProfileReadInt(path, "explainer", "openOnLaunch", chkOpenEW.Value) ? 1 : 0
    chkTop_EW.Value := GameProfileReadInt(path, "explainer", "alwaysOnTop", chkTop_EW.Value) ? 1 : 0

    profileDpadNavigation := GameProfileReadInt(path, "controls", "dpadNavigation", -1)
    if (profileDpadNavigation >= 0)
        CPControllerSetDpadNavigationEnabled(profileDpadNavigation != 0, true)

    slTrans.Value := overlayTrans
    lblTransPct.Value := Round(overlayTrans / 255 * 100) "%"
    rectBg.Opt("Background" boxBgHex)
    rectTxt.Opt("Background" txtHex)
    rectName.Opt("Background" nameHex)
    try ddlFont.Text := fontName
    try edFSize.Value := fontSize, udFSize.Value := fontSize
    chkFontBold.Value := fontBold

    slTrans_EW.Value := overlayTrans_EW
    lblTransPct_EW.Value := Round(overlayTrans_EW / 255 * 100) "%"
    rectBg_EW.Opt("Background" boxBgHex_EW)
    rectTxt_EW.Opt("Background" txtHex_EW)
    try ddlFont_EW.Text := fontName_EW
    try edFSize_EW.Value := fontSize_EW, udFSize_EW.Value := fontSize_EW
    chkFontBold_EW.Value := fontBold_EW

    IniWriteRetry(capMode, iniPath, "capture", "mode")
    IniWriteRetry(capRect, iniPath, "capture", "rect")
    IniWriteRetry(capWinInfo, iniPath, "capture", "winTitle")
    IniWriteRetry(explainPromptProfile, iniPath, "cfg", "explainPromptProfile")
    IniWriteRetry(jp2enGlossaryProfile, iniPath, "cfg", "jp2enGlossaryProfile")
    IniWriteRetry(en2enGlossaryProfile, iniPath, "cfg", "en2enGlossaryProfile")
    IniWriteRetry(useTerminologyOverrides, iniPath, "cfg", "useTerminologyOverrides")
    IniWriteRetry(chkGuess.Value, iniPath, "cfg", "highlightGuessed")
    IniWriteRetry(chkName.Value, iniPath, "cfg", "colorSpeaker")
    IniWriteRetry(chkOpenTW.Value, iniPath, "cfg", "openTranslatorOnLaunch")
    IniWriteRetry(chkTop_TW.Value, iniPath, "cfg", "winTop")
    IniWriteRetry(chkOpenEW.Value, iniPath, "cfg", "openExplainerOnLaunch")
    IniWriteRetry(chkTop_EW.Value, iniPath, "cfg_explainer", "winTop")
    IniWriteRetry(name, iniPath, "game_profiles", "active")

    ; The control panel also tracks the Explainer's outer bounds. Pause its
    ; watcher while applying so stale live coordinates cannot overwrite the
    ; profile before the overlay finishes moving.
    explainerHwndBefore := CPFindExactWindow("Explainer")
    watcherWasRunning := ew_bounds_watch_running
    if watcherWasRunning
        StopExplainerBoundsWatcher()

    explainerBounds := GameProfileEnsureBoundsVisible(GameProfileReadBounds(path, "explainer"))
    if IsObject(explainerBounds) {
        ewX := explainerBounds["x"]
        ewY := explainerBounds["y"]
    }
    controlW := IniRead(path, "explainer_control_bounds", "w", "")
    controlH := IniRead(path, "explainer_control_bounds", "h", "")
    if (controlW != "" && IsNumber(controlW))
        ewW := Integer(controlW)
    else if IsObject(explainerBounds)
        ewW := Round(explainerBounds["w"] * explainerBounds["dpi"] / 96)
    if (controlH != "" && IsNumber(controlH))
        ewH := Integer(controlH)
    else if IsObject(explainerBounds)
        ewH := Round(explainerBounds["h"] * explainerBounds["dpi"] / 96)
    ew_lastX := ewX, ew_lastY := ewY, ew_lastW := ewW, ew_lastH := ewH

    SaveAll()
    RefreshColorSwatches()
    RefreshColorSwatches_EW()
    GameProfileApplyBounds("Translator", path, "translator")
    appliedExplainerBounds := GameProfileApplyBounds("Explainer", path, "explainer")
    if IsObject(appliedExplainerBounds) {
        ewX := appliedExplainerBounds["x"]
        ewY := appliedExplainerBounds["y"]
        ew_lastX := ewX, ew_lastY := ewY
        IniWriteRetry(ewX, iniPath, "explainer_bounds", "x")
        IniWriteRetry(ewY, iniPath, "explainer_bounds", "y")
    }
    SendOverlayTheme()

    for title, topmost in Map("Translator", chkTop_TW.Value, "Explainer", chkTop_EW.Value) {
        hwnd := CPFindExactWindow(title)
        if hwnd
            try WinSetAlwaysOnTop(topmost ? 1 : 0, "ahk_id " hwnd)
    }

    if (watcherWasRunning || explainerHwndBefore)
        SetTimer(StartExplainerBoundsWatcher, -350)

    DbgCP("Applied unified profile: " name)
    if announce
        Toast("Profile applied: " name)
    if (warnings != "") {
        if announce
            MsgBox("The profile was applied with these exceptions:`n`n" warnings, "Profiles", "OK Icon!")
        else
            DbgCP("Profile applied with exceptions: " StrReplace(warnings, "`n", " | "))
    }
    return true
}

CPApplyExternalProfile(name) {
    name := GameProfileSafeName(name)
    if (name = "")
        return
    if GameProfileApply(name, false)
        RefreshGameProfilesList(name)
}

CPControlPanelCopyData(wParam, lParam, msg, hwnd) {
    if !lParam
        return 0
    payloadPtr := NumGet(lParam + A_PtrSize * 2, "UPtr")
    if !payloadPtr
        return 0
    payload := StrGet(payloadPtr, "UTF-16")
    if (payload = "open_study_library") {
        SetTimer(OpenStandaloneStudyLibrary, -10)
        return 1
    }
    if (payload = "open_study_reader") {
        SetTimer(OpenStandaloneStudyReader, -10)
        return 1
    }
    prefix := "apply_profile="
    if (SubStr(payload, 1, StrLen(prefix)) != prefix)
        return 0
    profileName := SubStr(payload, StrLen(prefix) + 1)
    SetTimer(CPApplyExternalProfile.Bind(profileName), -10)
    return 1
}

RefreshGameProfilesList(select := "") {
    global ddlGameProfile, iniPath
    list := ListGameProfiles()
    ddlGameProfile.Delete()
    if list.Length {
        ddlGameProfile.Add(list)
        if (select = "")
            select := IniRead(iniPath, "game_profiles", "active", "")
        index := ArrayIndexOf(list, select)
        ddlGameProfile.Choose(index ? index : 1)
    } else {
        ddlGameProfile.Text := ""
    }
    GameProfileUpdateSummary()
}

GameProfileUpdateSummary(*) {
    global ddlGameProfile, txtGameProfileState, iniPath
    if !(IsSet(ddlGameProfile) && IsSet(txtGameProfileState))
        return
    name := Trim(ddlGameProfile.Text)
    active := IniRead(iniPath, "game_profiles", "active", "")
    if (name = "")
        txtGameProfileState.Value := "No profiles have been created yet."
    else
        txtGameProfileState.Value := "Selected: " name (name = active ? " (last saved or applied)" : "")
}

CreateGameProfile(*) {
    global ddlGameProfile
    input := CPThemedInputBox("Enter a name for the new profile:", "Create Profile", "", "", 400)
    if (input.Result != "OK")
        return
    name := GameProfileSafeName(input.Value)
    if (name = "") {
        MsgBox("Please enter a non-empty profile name.", "Profiles", "OK Icon!")
        return
    }
    path := GameProfilePath(name)
    if FileExist(path) && MsgBox("Profile '" name "' already exists. Overwrite it?", "Profiles", "YesNo Icon!") != "Yes"
        return
    if GameProfileSave(name)
        RefreshGameProfilesList(name)
}

SaveSelectedGameProfile(*) {
    global ddlGameProfile
    name := Trim(ddlGameProfile.Text)
    if (name = "") {
        CreateGameProfile()
        return
    }
    if GameProfileSave(name)
        RefreshGameProfilesList(name)
}

ApplySelectedGameProfile(*) {
    global ddlGameProfile
    name := Trim(ddlGameProfile.Text)
    if (name = "") {
        MsgBox("Select a profile first.", "Profiles", "OK Icon!")
        return
    }
    if GameProfileApply(name)
        RefreshGameProfilesList(name)
}

DeleteSelectedGameProfile(*) {
    global ddlGameProfile, iniPath
    name := Trim(ddlGameProfile.Text)
    if (name = "")
        return
    if MsgBox("Delete profile '" name "'?`n`nPrompts, terminology files, and overlay settings will not be deleted.", "Profiles", "YesNo Icon!") != "Yes"
        return
    path := GameProfilePath(name)
    try FileDelete(path)
    if (IniRead(iniPath, "game_profiles", "active", "") = name)
        IniWriteRetry("", iniPath, "game_profiles", "active")
    RefreshGameProfilesList()
}

; =========================
; GUI
; =========================
cpUiCreateOptions := "+Resize +MinSize" CP_VIEWPORT_MIN_W "x" CP_VIEWPORT_MIN_H
    . " +0x300000 +ToolWindow +E0x08000000" ; WS_EX_NOACTIVATE
; Every launch begins with non-activating tool-window styles. Even AutoHotkey's
; hidden measurement Show can otherwise leak a short ordinary white frame to
; DWM. Normal launches restore the application-window styles immediately before
; their intentional final reveal; background/Study launches stay hidden.
ui := Gui(cpUiCreateOptions, "JRPG Translator")
; Cloak immediately—before even AutoHotkey's first hidden measurement Show.
; Keeping this same window out of composition throughout initialization avoids
; a second warm-up presentation and lets the final reveal be genuinely singular.
CPSetWindowCloaked(ui.Hwnd, true)
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

CPConstrainPanelBoundsToWorkArea(&x, &y, &clientW, &clientH, nonClientW, nonClientH) {
    global ui, CP_VIEWPORT_MIN_W, CP_VIEWPORT_MIN_H

    originalX := x, originalY := y
    originalW := clientW, originalH := clientH
    proposedRight := x + clientW + nonClientW
    proposedBottom := y + clientH + nonClientH
    bestMonitor := 0
    bestArea := -1
    ; Gui.Show/GetPos/GetClientPos use AutoHotkey's DPI-scaled logical units,
    ; while MonitorGetWorkArea returns physical desktop pixels. Convert every
    ; monitor rectangle before comparing or clamping the saved GUI rectangle.
    panelDpi := GetWindowDPI(ui.Hwnd)
    physicalToLogical := 96 / Max(96, panelDpi)

    try {
        Loop MonitorGetCount() {
            MonitorGetWorkArea(A_Index, &left, &top, &right, &bottom)
            left := Ceil(left * physicalToLogical)
            top := Ceil(top * physicalToLogical)
            right := Floor(right * physicalToLogical)
            bottom := Floor(bottom * physicalToLogical)
            intersectionW := Max(0, Min(proposedRight, right) - Max(x, left))
            intersectionH := Max(0, Min(proposedBottom, bottom) - Max(y, top))
            intersectionArea := intersectionW * intersectionH
            if (intersectionArea > bestArea) {
                bestArea := intersectionArea
                bestMonitor := A_Index
            }
        }
    }

    if !bestMonitor {
        try bestMonitor := MonitorGetPrimary()
        catch
            bestMonitor := 1
    }

    try {
        MonitorGetWorkArea(bestMonitor, &workLeft, &workTop, &workRight, &workBottom)
        workLeft := Ceil(workLeft * physicalToLogical)
        workTop := Ceil(workTop * physicalToLogical)
        workRight := Floor(workRight * physicalToLogical)
        workBottom := Floor(workBottom * physicalToLogical)
    }
    catch {
        workLeft := 0, workTop := 0
        workRight := Floor(A_ScreenWidth * physicalToLogical)
        workBottom := Floor(A_ScreenHeight * physicalToLogical)
    }

    workW := Max(1, workRight - workLeft)
    workH := Max(1, workBottom - workTop)
    maxClientW := Max(1, workW - nonClientW)
    maxClientH := Max(1, workH - nonClientH)
    minClientW := Min(CP_VIEWPORT_MIN_W, maxClientW)
    minClientH := Min(CP_VIEWPORT_MIN_H, maxClientH)
    clientW := Max(minClientW, Min(clientW, maxClientW))
    clientH := Max(minClientH, Min(clientH, maxClientH))

    outerW := clientW + nonClientW
    outerH := clientH + nonClientH
    x := Min(Max(x, workLeft), Max(workLeft, workRight - outerW))
    y := Min(Max(y, workTop), Max(workTop, workBottom - outerH))

    return (x != originalX || y != originalY
        || clientW != originalW || clientH != originalH)
}

CPApplyInitialPanelLayout(*) {
    global ui
    if !(IsSet(ui) && ui && ui.Hwnd)
        return
    try {
        ui.GetClientPos(,, &clientW, &clientH)
        ResizeUI(ui, 0, clientW, clientH)
    }
}

ui.MarginX := pad, ui.MarginY := pad
ui.SetFont("s10", "Segoe UI")
ui.BackColor := CPPalette(controlDarkMode)["window"]

; The native tab remains as a page host and focus proxy. A custom tab bar is
; created after all page controls so its styling and geometry are predictable.
tabNames := ["Screenshot Translation","Audio Translation","Translation Window","Explanation","Explanation Window","Terminology Overrides","Profiles","Controls","API Keys","Paths"]
CPTabVisiblePages := [1, 2, 3, 4, 5, 6, 7, 8, 9]
if showPathsTab
    CPTabVisiblePages.Push(10)
tab := ui.Add("Tab", "xm ym w760 h420 Buttons -Wrap", tabNames)
CPRegisterCanvasFixedControl(tab, false, true)

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

; Keep captures for the current session, then clear them at the next startup.
clearScreenshotsOnStartup := Integer(IniRead(iniPath, "paths", "clearScreenshotsOnStartup", 0)) ? 1 : 0
chkDel := ui.Add("Checkbox", "xm y+16", "Clear screenshots on startup")
chkDel.Value := clearScreenshotsOnStartup
; Persist to control.ini immediately when toggled
chkDel.OnEvent("Click", (*) => IniWrite(chkDel.Value ? 1 : 0, iniPath, "paths", "clearScreenshotsOnStartup"))

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

; --- Max size + Capture picker (native path, non-breaking) ---
; Pin this row to the left-column baseline under the screenshot-cleanup checkbox
chkDel.GetPos(&delX, &delY, , &delH)
leftBaseY := delY + delH + 16  ; spacing under the delete checkbox

ui.Add("Text", Format("xm y{} w160", leftBaseY), "Maximum PNG size (KB):")
eCapMax := ui.Add("Edit", "x+m yp w80 Number", capMaxKB)
txtCapMaxHint := ui.Add("Text", "x+8 yp w44 h23 Hidden Center Border +0x200", "A")
eCapMax.OnEvent("Change", (*) => SetCapMaxKBFromUI(eCapMax.Value))

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
btnAudioTest := ui.Add("Button", "x+6 w90", "Test Audio")

txtAudioHelp := ui.Add(
    "Text",
    "xm y+10 w700 h76",
    "Live Audio Translation captures the selected Windows output and translates spoken Japanese as it plays.`n"
    . "Play any audible sound, then select `"Test Audio.`" If no sound is detected, verify the selected device "
    . "or try another audio driver in the game or emulator, such as XAudio instead of WASAPI in RetroArch. "
    . "No API request is made."
)
CPRegisterMutedControl(txtAudioHelp)
txtAudioTestStatus := ui.Add(
    "Text",
    "xm y+4 w700 h24",
    "Test status: Not tested"
)

; Live translation provider
lblLiveTranslation := ui.Add("Text", "xm y+12 w180", "Live translation")
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
btnAudioTest.OnEvent("Click", TestAudioInput)
ddlSpeaker.OnEvent("Change", SpeakerChanged)

; --- Tab 3: TRANSLATION WINDOW
tab.UseTab(3)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer
twLabelX := pad + 16
twLabelW := 140
twControlW := 280
twSwatchW := 84
twSwatchX := twLabelX + twLabelW + pad + twControlW - twSwatchW

ui.Add("Text", "x" twLabelX " y+10 w" twLabelW, "Overlay Opacity:")
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
txtFontSizeHint := ui.Add("Text", "x+8 yp w44 h23 Hidden Center Border +0x200", "A")
chkFontBold := ui.Add("CheckBox", "x+8 yp+2", "Bold")
chkFontBold.Value := fontBold

twControlX := twLabelX + twLabelW + pad
btnMoveResize := ui.Add("Button", "x" twControlX " y+28 w180", "Move / Resize")
txtMoveResize := ui.Add("Text", "x" twControlX " y+5 w590 cGray"
    , "Left stick or arrows move; right stick or Screenshot + Translate + arrows resize. Enter saves, Esc cancels.")
CPRegisterMutedControl(txtMoveResize)
btnMoveResize.OnEvent("Click", StartOverlayAdjustment.Bind("Translator"))

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
btnOpenStudyLibrary := ui.Add("Button", "x+8 yp w200", "Open Study Library...")

; Migrate the existing text-archive preference conservatively. Existing users who
; already enabled text archives keep both outputs; a new installation starts with
; the Study Library enabled and plain-text copies optional.
saveExplRaw := IniRead(iniPath, "cfg", "saveExplains", "__missing__")
saveExplVal := (saveExplRaw = "__missing__") ? 0 : Integer(saveExplRaw)
saveLibraryRaw := IniRead(iniPath, "cfg", "saveStudyLibrary", "__missing__")
if (saveLibraryRaw = "__missing__") {
    saveLibraryVal := (saveExplRaw = "__missing__") ? 1 : saveExplVal
    IniWrite(saveLibraryVal, iniPath, "cfg", "saveStudyLibrary")
} else {
    saveLibraryVal := Integer(saveLibraryRaw)
}
saveLibraryScreenshotsVal := Integer(IniRead(
    iniPath, "cfg", "studyLibraryScreenshots", 1
))

saveLibraryChk := ui.Add("CheckBox", "xs y+16", "Save explanations to Study Library")
saveLibraryChk.Value := saveLibraryVal ? 1 : 0
saveLibraryScreenshotsChk := ui.Add(
    "CheckBox", "xs y+10", "Include source screenshots in Study Library"
)
saveLibraryScreenshotsChk.Value := saveLibraryScreenshotsVal ? 1 : 0
saveLibraryScreenshotsChk.Enabled := saveLibraryChk.Value ? true : false
saveExplChk := ui.Add("CheckBox", "xs y+10", "Save plain-text copies")
saveExplChk.Value := saveExplVal ? 1 : 0

txtExplainSaveInfo := ui.Add("Text"
    , "xm y+4 w760 h42 cGray"
    , "Study Library entries are grouped by the active unified Profile; without one, they are kept under Unsorted. Optional plain-text copies continue using the 'Settings\\Explanations' folder and its Profile subfolders."
)
CPRegisterMutedControl(txtExplainSaveInfo)
saveLibraryChk.OnEvent("Click", StudyLibrarySaveToggleChanged)
saveLibraryScreenshotsChk.OnEvent("Click", (*) => IniWrite(
    saveLibraryScreenshotsChk.Value ? 1 : 0, iniPath, "cfg", "studyLibraryScreenshots"
))
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
btnOpenStudyLibrary.OnEvent("Click", OpenStudyLibrary)

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

ui.Add("Text", "x" ewLabelX " y+10 w" ewLabelW, "Overlay Opacity:")
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
udFSize_EW := ui.Add("UpDown", "Range6-200")
txtFontSizeHint_EW := ui.Add("Text", "x+8 yp w44 h23 Hidden Center Border +0x200", "A")
chkFontBold_EW := ui.Add("CheckBox", "x+8 yp+2", "Bold")
chkFontBold_EW.Value := fontBold_EW

ewControlX := ewLabelX + ewLabelW + pad
btnMoveResize_EW := ui.Add("Button", "x" ewControlX " y+28 w180", "Move / Resize")
txtMoveResize_EW := ui.Add("Text", "x" ewControlX " y+5 w590 cGray"
    , "Left stick or arrows move; right stick or Screenshot + Translate + arrows resize. Enter saves, Esc cancels.")
CPRegisterMutedControl(txtMoveResize_EW)
btnMoveResize_EW.OnEvent("Click", StartOverlayAdjustment.Bind("Explainer"))

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

chkUseTerminologyOverrides := ui.Add("CheckBox", "xm y+20", "Use terminology overrides")
chkUseTerminologyOverrides.Value := useTerminologyOverrides
chkUseTerminologyOverrides.OnEvent("Click", TerminologyOverridesChanged)

txtGlossaryHelp2 := ui.Add("Text", "xm y+20 cGray w760"
  , "TL -> TL glossary: Corrects matching words or phrases locally after the model returns its translation; these entries are never sent to the model. Add incorrect or inconsistent outputs as you encounter them while playing, or enter known and likely variants in advance. Example: Esuteru -> Estelle.")

; --- Row 1: target-language -> target-language glossary
ui.Add("Text", "xm y+12 w112", "TL -> TL profile:")
ddlENG := ui.Add("DropDownList", "x+m w210 0x210", [])
btnENG_Edit := ui.Add("Button", "x+6 w125", "Manage Entries...")
btnENG_New  := ui.Add("Button", "x+6 w105", "New Profile...")
btnENG_Del  := ui.Add("Button", "x+6 w110", "Delete Profile...")

txtGlossaryHelp3 := ui.Add("Text", "xm y+20 cGray w760"
  , "JP -> TL glossary: Sends exact Japanese-to-target-language mappings to the model as additional instructions. This can stabilize known names and terms, but results depend on the model and the complexity of the selected prompt. A model may ignore a rule or apply it unexpectedly, including to unrelated or similar-sounding names. Example: エステル -> Estelle.")

; --- Row 2: Japanese -> target-language glossary
ui.Add("Text", "xm y+12 w112", "JP -> TL profile:")
ddlJPG := ui.Add("DropDownList", "x+m w210 0x210", [])   ; filled by RefreshGlossaryProfilesList
btnJPG_Edit := ui.Add("Button", "x+6 w125", "Manage Entries...")
btnJPG_New  := ui.Add("Button", "x+6 w105", "New Profile...")
btnJPG_Del  := ui.Add("Button", "x+6 w110", "Delete Profile...")

txtGlossaryHelp4 := ui.Add("Text", "xm y+20 cGray w760"
  , 'How to use: Choose an independent profile for each glossary type. "Manage Entries..." opens its terminology table; "New Profile..." creates only that glossary type, and "Delete Profile..." removes only that type. Changes apply to the next screenshot translation.')

for cpMutedGlossaryCtrl in [txtGlossaryHelp1, txtGlossaryHelp2, txtGlossaryHelp3, txtGlossaryHelp4]
    CPRegisterMutedControl(cpMutedGlossaryCtrl)

; wire up events
btnJPG_Edit.OnEvent("Click", (*) => OpenGlossaryManager("jp"))
btnJPG_New .OnEvent("Click", (*) => NewGlossaryProfile("jp"))
btnJPG_Del .OnEvent("Click", (*) => DeleteGlossaryProfile("jp"))

btnENG_Edit.OnEvent("Click", (*) => OpenGlossaryManager("en"))
btnENG_New .OnEvent("Click", (*) => NewGlossaryProfile("en"))
btnENG_Del .OnEvent("Click", (*) => DeleteGlossaryProfile("en"))

ddlJPG.OnEvent("Change", (*) => GlossaryChanged("jp"))
ddlENG.OnEvent("Change", (*) => GlossaryChanged("en"))

; initial fill + selection
RefreshGlossaryProfilesList(jp2enGlossaryProfile, en2enGlossaryProfile)

; --- Tab 7: PROFILES
tab.UseTab(7)
ui.Add("Text", "xm y+4 w0 h0")
lblGameProfilesTitle := ui.Add("Text", "xm y+10 w760", "Profiles")
lblGameProfilesTitle.SetFont("Bold")
txtGameProfileIntro := ui.Add("Text", "xm y+8 w760 cGray"
    , "Save and restore complete setups for translation and explanation prompts, capture, terminology, overlays, controls, and the selected Study Library.")
txtGameProfileGlobal := ui.Add("Text", "xm y+6 w760 cGray"
    , "Audio provider, model, input device, and target language remain global and are not changed by a profile.")
CPRegisterMutedControl(txtGameProfileIntro)
CPRegisterMutedControl(txtGameProfileGlobal)

ui.Add("Text", "xm y+24 w120", "Profile:")
ddlGameProfile := ui.Add("DropDownList", "x+m w330 0x210", [])
btnGameProfileAdd := ui.Add("Button", "x+8 w80", "Add")
btnGameProfileSave := ui.Add("Button", "x+6 w110", "Save Current")
btnGameProfileApply := ui.Add("Button", "x+6 w80", "Apply")
btnGameProfileDelete := ui.Add("Button", "x+6 w80", "Delete")

txtGameProfileState := ui.Add("Text", "xm y+18 w760", "")
txtGameProfileDetails := ui.Add("Text", "xm y+18 w760 cGray"
    , "Saved settings include overlay appearance and placement, startup choices, capture target, both prompts, terminology profiles, and the active Study Library. Screenshot output processing follows the selected prompt automatically.")
CPRegisterMutedControl(txtGameProfileDetails)

ddlGameProfile.OnEvent("Change", GameProfileUpdateSummary)
btnGameProfileAdd.OnEvent("Click", CreateGameProfile)
btnGameProfileSave.OnEvent("Click", SaveSelectedGameProfile)
btnGameProfileApply.OnEvent("Click", ApplySelectedGameProfile)
btnGameProfileDelete.OnEvent("Click", DeleteSelectedGameProfile)
RefreshGameProfilesList()

; --- Tab 10: PATHS
tab.UseTab(10)
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

btnSavePaths := ui.Add("Button", "xm y+16 w140", "Save paths")
btnSavePaths.Enabled := false
btnSavePaths.OnEvent("Click", (*) => SaveEditedPaths())

; --- Advanced screenshot output override (hidden with the Paths tab) ---
directOpts := "xm y+18 w220"
if (directModelOutput)
    directOpts .= " Checked"
cbDirectModelOutput := ui.Add("CheckBox", directOpts, "Direct model output")
TooltipBind(cbDirectModelOutput, "Advanced: bypass screenshot translation extraction and show the model response unchanged")
cbDirectModelOutput.OnEvent("Click", CPOnDirectModelOutputToggle)

; --- Debug toggle (bottom of Paths tab) ---
opts := "xm y+12 w140"
if (debugMode)
    opts .= " Checked"
cbDebug := ui.Add("CheckBox", opts, "Debug mode")
TooltipBind(cbDebug, "Write diagnostic logs for the Control Panel, overlays, and live audio translator")
cbDebug.OnEvent("Click", CPOnDebugModeToggle)

; --- Tab 8: CONTROLS
tab.UseTab(8)
ui.Add("Text", "xm y+4 w0 h0")

; Push-like radio buttons make the two input methods feel like views of one
; page while retaining native keyboard, mouse, and controller navigation.
global rbControlsKeyboard := ui.Add("Radio", "xm y+6 w180 h32 Group +0x1000", "Keyboard inputs")
global rbControlsController := ui.Add("Radio", "x+0 yp w180 h32 +0x1000", "Controller inputs")
rbControlsKeyboard.OnEvent("Click", CPSetControlsView.Bind("keyboard", true))
rbControlsController.OnEvent("Click", CPSetControlsView.Bind("controller", true))

rbControlsKeyboard.GetPos(&controlsViewX, &controlsViewY, , &controlsViewH)
controlsContentY := controlsViewY + controlsViewH + 14
controlsActionX := controlsViewX
controlsBindingX := controlsActionX + 270
controlsButtonX := controlsBindingX + 250
controlsRowY := controlsContentY + 28
controlsRowStep := 38

; Keyboard input view. These are the original hotkey controls and continue to
; work with physical keyboards and keyboard-emulation tools such as JoyToKey.
CPAddControlsViewControl("keyboard", ui.Add("Text", "x" controlsActionX " y" controlsContentY " w260", "Action"))
CPAddControlsViewControl("keyboard", ui.Add("Text", "x" controlsBindingX " y" controlsContentY " w240", "Current keyboard input"))

for keyboardActionKey in hotkeyActions {
    label := hotkeyLabels[keyboardActionKey]
    cur := IniRead(iniPath, "hotkeys", keyboardActionKey, hotkeyDefaults[keyboardActionKey])

    lblKeyboardAction := CPAddControlsViewControl("keyboard", ui.Add("Text", "x" controlsActionX " y" controlsRowY " w260 h26 +0x200", label))
    e := CPAddControlsViewControl("keyboard", ui.Add("Edit", "x" controlsBindingX " y" controlsRowY " w240 h28 ReadOnly", cur))
    bChg := CPAddControlsViewControl("keyboard", ui.Add("Button", "x" controlsButtonX " y" controlsRowY " w90 h28", "Change..."))
    bDis := CPAddControlsViewControl("keyboard", ui.Add("Button", "x+6 yp w80 h28", "Disable"))
    bDef := CPAddControlsViewControl("keyboard", ui.Add("Button", "x+6 yp w100 h28", "Default"))

    hkEdits[keyboardActionKey] := e
    hkBtnChg[keyboardActionKey] := bChg
    hkBtnDis[keyboardActionKey] := bDis
    hkBtnDef[keyboardActionKey] := bDef
    actKey := keyboardActionKey
    bChg.OnEvent("Click", Hotkey_Row_Change.Bind(actKey))
    bDis.OnEvent("Click", Hotkey_Row_Disable.Bind(actKey))
    bDef.OnEvent("Click", Hotkey_Row_Default.Bind(actKey))
    controlsRowY += controlsRowStep
}

global hkConflictText
hkConflictText := CPAddControlsViewControl("keyboard", ui.Add("Text", "x" controlsActionX " y" controlsRowY " w800 h18 cRed", ""))
Hotkeys_ShowConflicts()

; Keep the controller binding list aligned with the keyboard list. Controller-
; only options use the otherwise empty area to the right of the Disable buttons.
controllerTopY := controlsContentY
controllerBindingX := controlsActionX + 235
controllerButtonX := controllerBindingX + 220
controllerOptionsX := controlsActionX + 630
controllerHeaderY := controllerTopY
controllerRowY := controllerHeaderY + 28

CPAddControlsViewControl("controller", ui.Add("Text", "x" controlsActionX " y" controllerHeaderY " w225", "Action"))
CPAddControlsViewControl("controller", ui.Add("Text", "x" controllerBindingX " y" controllerHeaderY " w210", "Current controller input"))
global txtControllerStatus := CPAddControlsViewControl("controller", ui.Add("Text", "x" controllerOptionsX " y" controllerHeaderY " w350 h20 cGray", "Action bindings are off."))
CPRegisterMutedControl(txtControllerStatus)

global cbControllerInputsEnabled := CPAddControlsViewControl("controller", ui.Add("CheckBox", "x" controllerOptionsX " y" controllerRowY " w350 h28 +0x2000", "Enable direct controller action bindings"))
cbControllerInputsEnabled.OnEvent("Click", CPControllerEnabledChanged)
TooltipBind(cbControllerInputsEnabled, "Enable optional controller buttons for translator actions. The game still receives the same button presses.")

global cbControllerDpadNavigationEnabled := CPAddControlsViewControl("controller", ui.Add("CheckBox", "x" controllerOptionsX " y" (controllerRowY + controlsRowStep) " w350 h28 +0x2000", "Use D-pad for control panel navigation"))
cbControllerDpadNavigationEnabled.OnEvent("Click", CPControllerDpadNavigationChanged)
TooltipBind(cbControllerDpadNavigationEnabled, "Turn this off when another app maps the D-pad to arrow keys, to prevent duplicate navigation. A / Cross, B / Circle, and keyboard controls remain available.")

global txtControllerDpadNote := CPAddControlsViewControl("controller", ui.Add("Text", "x" controllerOptionsX " y" (controllerRowY + (controlsRowStep * 2) + 2) " w350 h72 cGray", "LB / L1 and RB / R1 switch main tabs while this panel is active. Disable D-pad navigation if another app also sends arrow keys; A / Cross and B / Circle remain available."))
CPRegisterMutedControl(txtControllerDpadNote)

for controllerActionKey in hotkeyActions {
    label := hotkeyLabels[controllerActionKey]
    lblControllerAction := CPAddControlsViewControl("controller", ui.Add("Text", "x" controlsActionX " y" controllerRowY " w225 h26 +0x200", label))
    controllerEdit := CPAddControlsViewControl("controller", ui.Add("Edit", "x" controllerBindingX " y" controllerRowY " w210 h28 ReadOnly", "Disabled"))
    btnControllerAssign := CPAddControlsViewControl("controller", ui.Add("Button", "x" controllerButtonX " y" controllerRowY " w80 h28", "Assign"))
    btnControllerDisable := CPAddControlsViewControl("controller", ui.Add("Button", "x+6 yp w75 h28", "Disable"))

    CPControllerBindingEdits[controllerActionKey] := controllerEdit
    CPControllerAssignButtons[controllerActionKey] := btnControllerAssign
    CPControllerDisableButtons[controllerActionKey] := btnControllerDisable
    controllerAction := controllerActionKey
    btnControllerAssign.OnEvent("Click", CPControllerAssign.Bind(controllerAction))
    btnControllerDisable.OnEvent("Click", CPControllerDisable.Bind(controllerAction))
    controllerRowY += controlsRowStep
}

CPControllerLoadBindings()
savedControlsView := IniRead(iniPath, "controller_inputs", "view", "keyboard")
CPSetControlsView(savedControlsView, false)
CPControllerSetEnabled(Integer(IniRead(iniPath, "controller_inputs", "enabled", 0)) != 0, false)
CPControllerSetDpadNavigationEnabled(Integer(IniRead(iniPath, "controller_inputs", "dpad_navigation", 1)) != 0, false)

tab.UseTab()
FixAllEditableCombos()

; --- visual separator above global (non-tab) controls ---
; An opaque sibling behind the fixed footer prevents vertically scrolled tab
; controls from painting through the footer area.
CPFooterFill := ui.Add("Text", "x0 y0 w1 h1 Hidden BackgroundF0F0F0 +0x100 +0x04000000")
; SS_ETCHEDHORZ = 0x10 -> draws a 1â€“2 px horizontal etched line
sepAction := ui.Add("Text", "xm y+6 w1000 h2 0x10")
btnOv      := ui.Add("Button", "xm y+18 w130",  "Open Translator")
btnOvClose := ui.Add("Button", "x+6 w140",  "Close Translator")
btnAudio  := ui.Add("Button", "x+6 w180", "Audio Translation Off")
btnExplainerLaunch := ui.Add("Button", "x+6 w140", "Open Explainer")
btnExplainerClose  := ui.Add("Button", "x+6 w140", "Close Explainer")

bClose    := ui.Add("Button", "xm y+14 w120", "Close all")
chkTop    := ui.Add("CheckBox", "x+12 yp+6", "Always on top")
chkDarkMode := ui.Add("CheckBox", "x+18 yp", "Dark mode")
chkDarkMode.Value := controlDarkMode
chkDarkMode.OnEvent("Click", CPOnDarkModeToggle)
txtControlOpacity := ui.Add("Text", "x+18 yp", "Opacity:")
slControlOpacity := ui.Add("Slider", "x+8 yp-6 w110 h24 Range70-100 ToolTip")
slControlOpacity.Value := controlPanelOpacity
slControlOpacity.OnEvent("Change", CPOnControlPanelOpacityChange)
lblControlOpacityPct := ui.Add("Text", "x+8 yp+6 w44", controlPanelOpacity "%")

; The footer stays pinned to the visible bottom while page content scrolls.
for footerRegistrationCtrl in [CPFooterFill, sepAction, btnOv, btnOvClose, btnAudio, btnExplainerLaunch
    , btnExplainerClose, bClose, chkTop, chkDarkMode, txtControlOpacity
    , slControlOpacity, lblControlOpacityPct]
    CPRegisterCanvasFixedControl(footerRegistrationCtrl, false, true)
CPConfigurePreferredPanelHeight()

; Only manually typed path values require an explicit save.
for pathEditControl in [ePython, eOverlay, eImg, eAudio, eExplain]
    pathEditControl.OnEvent("Change", UpdatePathsDirtyState)

btnAudio.OnEvent("Click", ToggleAudioFromButton)
btnOv.OnEvent("Click",    LaunchOverlay)
btnOvClose.OnEvent("Click", CloseTranslatorOverlay)
btnExplainerLaunch.OnEvent("Click", LaunchExplainerOverlay)
btnExplainerClose.OnEvent("Click",  CloseExplainerOverlay)

bClose.OnEvent("Click", ClosePanel)

ClosePanel(*) {
    static closing := false
    if closing
        return
    if !ConfirmUnsavedPaths()
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
chkTop.OnEvent("Click", (*) => (
    ui.Opt(chkTop.Value ? "+AlwaysOnTop" : "-AlwaysOnTop"),
    IniWrite(chkTop.Value ? 1 : 0, iniPath, "cfg_control", "winTop")
))

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

; --- Tab 9: API KEYS
tab.UseTab(9)
ui.Add("Text", "xm y+4 w0 h0")  ; spacer

; Recommended section: keep secrets outside the app's plain-text .env file.
ui.SetFont("w600")
ui.Add("Text", "xm y+10", "Recommended: Store API keys in Windows")
ui.SetFont("w400")
txtApiHelp1 := ui.Add("Text"
    , "xm y+4 w740 cGray"
    , "Windows user environment variables keep your keys outside JRPG Translator's plain-text .env file. The app reads GEMINI_API_KEY and OPENAI_API_KEY automatically."
)
txtApiHelp2 := ui.Add("Text"
    , "xm y+6 w740 cGray"
    , "Click the button below. Under 'User variables', click 'New...' and add GEMINI_API_KEY with your Gemini key, then OPENAI_API_KEY with your OpenAI key. Confirm with OK and restart JRPG Translator."
)
txtApiHelp3 := ui.Add("Text"
    , "xm y+6 w740 cGray"
    , "(Alternatively, click Start, search for and select 'Edit the system environment variables', open the 'Advanced' tab, and click 'Environment Variables...'.)"
)

btnOpenEnvVars := ui.Add("Button", "xm y+12 w250", "Open Windows Environment Variables...")

; Alternative section: direct in-app entry.
ui.SetFont("w600")
ui.Add("Text", "xm y+24", "Alternative: Store API keys directly in JRPG Translator")
ui.SetFont("w400")
txtApiHelp4 := ui.Add("Text"
    , "xm y+4 w740 cGray"
    , "This convenient method stores the keys as plain text in Settings\.env. Leave it disabled to use the more secure Windows method above."
)

; Master toggle: default OFF (use Windows env). If .env exists, reflect that.
cbApiInApp := ui.Add("CheckBox", "xm y+12", "Enter API Keys in JRPG Translator (.env)")

for cpMutedApiCtrl in [txtApiHelp1, txtApiHelp2, txtApiHelp3, txtApiHelp4]
    CPRegisterMutedControl(cpMutedApiCtrl)

ui.Add("Text", "xm y+14 w200", "Gemini API key:")
eGemini := ui.Add("Edit", "x+m w420 Password")

ui.Add("Text", "xm y+10 w200", "OpenAI API key:")
eOpenAI := ui.Add("Edit", "x+m w420 Password")

; Buttons row
btnSaveEnv  := ui.Add("Button", "xm y+14 w120", "Save Keys")
btnDelEnv   := ui.Add("Button", "x+8 w120", "Delete .env")

; Keep About in the same visible action row, but far enough left that its full
; width remains inside the preferred magnetic-snap width at scaled DPIs.
btnAbout := ui.Add("Button", "x670 yp w120", "About...")

OpenWindowsEnvironmentVariables(*) {
    try Run('"' A_WinDir '\System32\rundll32.exe" sysdm.cpl,EditEnvironmentVariables')
    catch as ex
        MsgBox("Could not open Windows Environment Variables:`n" ex.Message,
            "API Keys", "OK Icon!")
}

btnOpenEnvVars.OnEvent("Click", OpenWindowsEnvironmentVariables)

OpenAboutUrl(url, description, *) {
    try Run(url)
    catch as ex
        MsgBox("Could not open " description ":`n" ex.Message,
            "About JRPG Translator", "OK Icon!")
}

AboutVersionInfo() {
    global APP_VERSION, PROJECT_URL
    return "JRPG Translator " APP_VERSION "`r`n"
        . "Windows: " A_OSVersion " (" (A_Is64bitOS ? "64-bit" : "32-bit") ")`r`n"
        . "Application: " (A_PtrSize = 8 ? "64-bit" : "32-bit")
        . (A_IsCompiled ? " compiled executable" : " source script") "`r`n"
        . "AutoHotkey runtime: " A_AhkVersion "`r`n"
        . "Project: " PROJECT_URL
}

CopyAboutVersionInfo(*) {
    try {
        A_Clipboard := AboutVersionInfo()
        ClipWait(0.5)
        Toast("Version information copied")
    } catch as ex {
        MsgBox("Could not copy version information:`n" ex.Message,
            "About JRPG Translator", "OK Icon!")
    }
}

CloseAboutDialog(dlg, *) {
    try dlg.Destroy()
}

SaveWelcomeGuidePreference(dontShowAgain) {
    global iniPath
    try IniWriteRetry(dontShowAgain.Value ? 1 : 0, iniPath, "cfg_control", "welcomeGuideDismissed")
}

CloseWelcomeGuide(dlg, dontShowAgain, *) {
    global CPWelcomeDialog
    SaveWelcomeGuidePreference(dontShowAgain)
    try dlg.Destroy()
    CPWelcomeDialog := 0
}

OpenWelcomeApiKeys(dlg, dontShowAgain, *) {
    global CPWelcomeDialog, ui, btnOpenEnvVars
    SaveWelcomeGuidePreference(dontShowAgain)
    try dlg.Destroy()
    CPWelcomeDialog := 0
    try ui.Show()
    try WinActivate("ahk_id " ui.Hwnd)
    CPSelectCustomTab(9)
    try btnOpenEnvVars.Focus()
}

ShowWelcomeDialog(manual := false, *) {
    global ui, iniPath, CPWelcomeDialog, BEGINNER_VIDEO_URL, WRITTEN_GUIDE_URL

    if (IsObject(CPWelcomeDialog) && CPWelcomeDialog.Hwnd) {
        try CPWelcomeDialog.Show()
        try WinActivate("ahk_id " CPWelcomeDialog.Hwnd)
        return
    }

    dismissed := Integer(IniRead(iniPath, "cfg_control", "welcomeGuideDismissed", 0)) != 0
    dlg := Gui("+Owner" ui.Hwnd " +AlwaysOnTop +OwnDialogs", "Welcome to JRPG Translator")
    CPWelcomeDialog := dlg
    dlg.MarginX := 20
    dlg.MarginY := 18
    dlg.SetFont("s10", "Segoe UI")

    dlg.SetFont("s17 Bold")
    dlg.Add("Text", "xm w640", "Welcome to JRPG Translator")
    dlg.SetFont("s10 Norm")
    dlg.Add("Text", "xm y+10 w640",
        "Complete these basics before starting your first game:")
    dlg.Add("Text", "xm y+12 w640",
        "1. Add at least one API key.`n"
        . "2. Choose your translation and explanation models and prompts.`n"
        . "3. Select a screenshot capture region or game window.`n"
        . "4. Configure keyboard shortcuts or direct controller inputs.`n"
        . "5. Save the finished setup as a Profile for the LaunchBox plugin.")

    dlg.Add("Text", "xm y+14 w640 cGray",
        "The LaunchBox plugin normally starts JRPG Translator with its control panel hidden. "
        . "Show it again at any time with your configured controller button or keyboard hotkey.")

    dontShowAgain := dlg.Add("CheckBox", "xm y+16 Checked", "Don't show this welcome guide again")
    if manual
        dontShowAgain.Value := dismissed ? 1 : 0

    btnWelcomeApiKeys := dlg.Add("Button", "xm y+20 w135", "Open API Keys")
    btnWelcomeVideo := dlg.Add("Button", "x+10 yp w170", "Watch Beginner Guide")
    btnWelcomeDocs := dlg.Add("Button", "x+10 yp w155", "Open Written Guide")
    btnWelcomeContinue := dlg.Add("Button", "x+10 yp w110 Default", "Continue")

    btnWelcomeApiKeys.OnEvent("Click", OpenWelcomeApiKeys.Bind(dlg, dontShowAgain))
    btnWelcomeVideo.OnEvent("Click", OpenAboutUrl.Bind(BEGINNER_VIDEO_URL, "the beginner guide"))
    btnWelcomeDocs.OnEvent("Click", OpenAboutUrl.Bind(WRITTEN_GUIDE_URL, "the written guide"))
    closeCallback := CloseWelcomeGuide.Bind(dlg, dontShowAgain)
    btnWelcomeContinue.OnEvent("Click", closeCallback)
    dlg.OnEvent("Escape", closeCallback)
    dlg.OnEvent("Close", closeCallback)

    dlg.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(dlg)
    try btnWelcomeApiKeys.Focus()
}

OpenWelcomeFromAbout(aboutDlg, *) {
    try aboutDlg.Destroy()
    SetTimer(ShowWelcomeDialog.Bind(true), -1)
}

ShowAboutDialog(*) {
    global ui, APP_VERSION, PROJECT_URL, BUG_REPORT_URL

    dlg := Gui("+Owner" ui.Hwnd " +AlwaysOnTop +OwnDialogs", "About JRPG Translator")
    dlg.MarginX := 20
    dlg.MarginY := 18
    dlg.SetFont("s10", "Segoe UI")

    dlg.SetFont("s16 Bold")
    dlg.Add("Text", "xm w540", "JRPG Translator")
    dlg.SetFont("s10 Norm")
    dlg.Add("Text", "xm y+8 w540", "Version " APP_VERSION)
    dlg.SetFont("s10 Bold")
    dlg.Add("Text", "xm y+14 w540", "Created by retrogamer0815")
    dlg.SetFont("s10 Norm")
    dlg.Add("Text", "xm y+6 w540", "MIT License")
    dlg.Add("Text", "xm y+16 w540",
        "Bug reports and source code are available on GitHub.")
    dlg.SetFont("s10 Bold")
    dlg.Add("Text", "xm y+6 w540", "Contact/update news: @retr0gamer42 on X")
    dlg.SetFont("s10 Norm")

    btnWelcomeGuide := dlg.Add("Button", "xm y+20 w180", "Open Welcome Guide...")
    btnReportBug := dlg.Add("Button", "xm y+10 w130", "Report a Bug...")
    btnGitHub := dlg.Add("Button", "x+10 yp w110", "Open GitHub")
    btnCopyVersion := dlg.Add("Button", "x+10 yp w150", "Copy Version Info")
    btnCloseAbout := dlg.Add("Button", "x+10 yp w100 Default", "Close")

    btnWelcomeGuide.OnEvent("Click", OpenWelcomeFromAbout.Bind(dlg))
    btnReportBug.OnEvent("Click", OpenAboutUrl.Bind(BUG_REPORT_URL, "the bug-report page"))
    btnGitHub.OnEvent("Click", OpenAboutUrl.Bind(PROJECT_URL, "GitHub"))
    btnCopyVersion.OnEvent("Click", CopyAboutVersionInfo)
    closeCallback := CloseAboutDialog.Bind(dlg)
    btnCloseAbout.OnEvent("Click", closeCallback)
    dlg.OnEvent("Escape", closeCallback)
    dlg.OnEvent("Close", closeCallback)

    dlg.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(dlg)
    try btnWelcomeGuide.Focus()
}

btnAbout.OnEvent("Click", ShowAboutDialog)

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

ui.OnEvent("Close",  ClosePanel)
ui.OnEvent("Escape", ExitControlPanel)

; wire buttons
btnA_GM_Add   .OnEvent("Click", (*) => AddModelInteractive(model_gemini_audio, "gemini_audio", ddlA_GM, "gemini", "audio"))
btnA_GM_Del   .OnEvent("Click", (*) => DeleteModel(model_gemini_audio, "gemini_audio", ddlA_GM))

btnTR_Add     .OnEvent("Click", (*) => AddModelInteractive(model_openai_audio, "openai_audio", ddlTR, "openai", "audio"))
btnTR_Del     .OnEvent("Click", (*) => DeleteModel(model_openai_audio, "openai_audio", ddlTR))

btnIMG_Add    .OnEvent("Click", (*) => AddModelInteractive(model_openai_img, "openai_img", ddlIMG, "openai", "screenshot"))
btnIMG_Del    .OnEvent("Click", (*) => DeleteModel(model_openai_img,"openai_img",   ddlIMG))

btnIMG_GM_Add .OnEvent("Click", (*) => AddModelInteractive(model_gemini_img, "gemini_img", ddlIMG_GM, "gemini", "screenshot"))
btnIMG_GM_Del .OnEvent("Click", (*) => DeleteModel(model_gemini_img,"gemini_img",   ddlIMG_GM))

; Initial status gets one WMI recovery scan. The repeating timer uses only the
; tracked PID and therefore performs no background process enumeration.
LoadFontsIntoCombo()
LoadFontsIntoCombo_EW()
Repaint()
UpdateStatus(true)
SetTimer(_UpdateStatus, 1000)

; Background windows already received ToolWindow/WS_EX_NOACTIVATE when the GUI
; was created, before any native child controls or hidden measurement frames.

; A dedicated Study process never exposes the control panel. Do not call Show,
; perform frame measurements, lay it out, theme it, or force a paint even with
; the Hide option: those native passes can still be briefly presented by DWM on
; some systems before the standalone Library/Reader window is ready.
if !CP_STUDY_ONLY_PROCESS {
; Create the native window once, hidden, so client/outer frame measurements are
; available before restored bounds are fitted to the current monitor work area.
ui.Show("Hide")
ui.GetPos(,, &cpMeasureOuterW, &cpMeasureOuterH)
ui.GetClientPos(,, &cpMeasureClientW, &cpMeasureClientH)
ui.Hide()
cpNonClientW := cpMeasureOuterW - cpMeasureClientW
cpNonClientH := cpMeasureOuterH - cpMeasureClientH

; Restore saved bounds when possible; otherwise start with the design defaults.
cpUsingSavedBounds := IsValidBounds(guiX_saved, guiY_saved, guiW_saved, guiH_saved)
if cpUsingSavedBounds {
    ; Older INI files stored outer sizes. Convert them once before fitting.
    if (bounds_mode != "client") {
        guiW_saved := Max(0, guiW_saved - cpNonClientW)
        guiH_saved := Max(0, guiH_saved - cpNonClientH)
        IniWrite("client", iniPath, "gui_bounds", "bounds_mode")
        DbgCP("Converted saved bounds from OUTER to CLIENT using ncW=" cpNonClientW " ncH=" cpNonClientH)
    }
    cpPanelX := guiX_saved, cpPanelY := guiY_saved
    cpPanelW := guiW_saved, cpPanelH := guiH_saved
} else {
    cpPanelX := defGuiX, cpPanelY := defGuiY
    cpPanelW := defGuiW, cpPanelH := defGuiH
}

cpBoundsAdjusted := CPConstrainPanelBoundsToWorkArea(
    &cpPanelX, &cpPanelY, &cpPanelW, &cpPanelH, cpNonClientW, cpNonClientH)
DbgCP((cpUsingSavedBounds ? "Restore" : "Use default")
    " panel bounds (client): x=" cpPanelX " y=" cpPanelY " w=" cpPanelW " h=" cpPanelH
    (cpBoundsAdjusted ? " [fitted to current work area]" : ""))
cpShowOptions := "w" cpPanelW " h" cpPanelH " x" cpPanelX " y" cpPanelY
; Assemble the complete first frame while hidden. Showing before the explicit
; layout, theme, and opacity passes produces several visible intermediate
; frames (and is particularly noticeable on a layered window).
ui.Show("Hide " cpShowOptions)
if !cpUsingSavedBounds
    SavePanelBounds()

; Do not depend on Windows delivering a final WM_SIZE after an off-screen or
; DPI/resolution-adjusted restore. Lay out the full current client area now.
ui.GetClientPos(,, &cpInitialClientW, &cpInitialClientH)
ResizeUI(ui, 0, cpInitialClientW, cpInitialClientH)
; Register live Size handling only after the hidden measurement/final layout
; passes. Otherwise a stale measurement-size event can repaint the just-opened
; panel with an intermediate layout at non-100% DPI.
ui.OnEvent("Size", ResizeUI)
; Ensure first paint draws all children cleanly (fixes clipped checkbox text/box)
DllCall("RedrawWindow"
    , "ptr", ui.Hwnd
    , "ptr", 0
    , "ptr", 0
    , "uint", 0x0001 | 0x0080 | 0x0100) ; RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW
CPCreateComboArrowOverlays()
CPApplyControlPanelTheme()
CPApplyControlPanelOpacity(false)
}

Rebind_LaunchExplainerRequest()
Rebind_ExplainLastTranslation()
Rebind_StartStopAudio()
Rebind_HideShowControlPanel()
RegisterControlPanelArrowNavigation()
SetTimer(UpdateCPFocusRing, 60)
UpdateCPFocusRing()
SetTimer(UpdateCPActiveTabHighlight, 80)
UpdateCPActiveTabHighlight()

; LaunchBox can select a unified Profile for a game. A first process applies
; the command-line selection here; a helper process sends the same request to
; this hidden window through WM_COPYDATA when the control panel already exists.
OnMessage(0x004A, CPControlPanelCopyData)
OnMessage(0x0231, StudyWindowSizeMoveMessage) ; WM_ENTERSIZEMOVE
OnMessage(0x0232, StudyWindowSizeMoveMessage) ; WM_EXITSIZEMOVE
if (CP_START_PROFILE != "")
    CPApplyExternalProfile(CP_START_PROFILE)

if (CP_STUDY_START_MODE = "library")
    SetTimer(OpenStandaloneStudyLibrary, -100)
else if (CP_STUDY_START_MODE = "reader")
    SetTimer(OpenStandaloneStudyReader, -100)

; A front end can explicitly request the Translator on a cold background start.
; This keeps the control panel hidden while still giving the game its overlay.
if (CP_STUDY_START_MODE = ""
    && (CP_START_TRANSLATOR
        || Integer(IniRead(iniPath, "cfg", "openTranslatorOnLaunch", 0)))) {
    SetTimer(LaunchOverlay, -100)
}
if (CP_STUDY_START_MODE = ""
    && Integer(IniRead(iniPath, "cfg", "openExplainerOnLaunch", 0))) {
    SetTimer(LaunchExplainerOverlay, -200)
}

; Some custom controls repaint themselves during initialization. Enforce the
; hidden state once more, then restore normal window styles for the first
; intentional hotkey/double-launch reveal.
if (CP_BACKGROUND_START) {
    ui.Hide()
    ; A normal LaunchBox background process may later reveal its control panel,
    ; so restore the ordinary styles after initialization. A dedicated Study
    ; process never exposes that panel and keeps the non-activating tool style
    ; for its full lifetime.
    if !CP_STUDY_ONLY_PROCESS
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
ForcePaint(ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt)
; Do one more pass after creation so no dropdowns look â€œselectedâ€ on first open
ClearAllComboSelections(*) {
    global ddlAProv, ddlA_GM, ddlTR, ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt
    for cmb in [ddlAProv, ddlA_GM, ddlTR, ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt]
        ComboUnselectText(cmb)
}

SetGuiAndTrayIcon(ui, A_ScriptDir "\icon.ico")

; A normal launch becomes visible only after layout, theme, opacity, custom
; controls, and the title-bar icon are ready. Paint the completed frame once.
if !CP_BACKGROUND_START {
    ; Remove the temporary startup-only styles while the prepared frame is still
    ; hidden/cloaked, so the taskbar and ordinary activation behavior are normal.
    CPSetWindowCloaked(ui.Hwnd, true)
    ui.Opt("-ToolWindow -E0x08000000")
    CPShowControlPanelReady(true)
}

; A normal first launch presents the compact setup guide. Background launches
; from LaunchBox/Big Box remain completely silent and hidden.
if (!CP_BACKGROUND_START
 && !Integer(IniRead(iniPath, "cfg_control", "welcomeGuideDismissed", 0)))
    SetTimer(ShowWelcomeDialog.Bind(false), -350)

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
    global ddlPrompt, promptProfile, imgPostproc
    global directModelOutput, cbDirectModelOutput
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

    RefreshPromptProfilesList(promptProfile)
    imgPostproc := SyncPromptPostproc(promptProfile)
    try cbDirectModelOutput.Value := directModelOutput

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
    global explainProvider, explainOpenAIModel, explainGeminiModel
    global model_openai_explain, model_gemini_explain

    prov := StrLower(Trim(IniRead(iniPath, "cfg_explainer", "explainProvider", "")))
    gm   := Trim(IniRead(iniPath, "cfg_explainer", "explainGeminiModel", ""))
    om   := Trim(IniRead(iniPath, "cfg_explainer", "explainOpenAIModel", ""))

    if (prov = "gemini" || prov = "openai") {
        explainProvider := prov
        ddlEProv.Choose(prov = "gemini" ? 1 : 2)
    }
    if (gm != "") {
        explainGeminiModel := gm
        SetComboToExistingItem(ddlEGem, model_gemini_explain, gm)
    }
    if (om != "") {
        explainOpenAIModel := om
        SetComboToExistingItem(ddlEOpenAI, model_openai_explain, om)
    }
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

SetAudioTestStatus(message) {
    global txtAudioTestStatus
    try txtAudioTestStatus.Text := "Test status: " message
}

TestAudioInput(*) {
    global pythonExe, audioScript, ddlSpeaker, btnAudioTest
    global gAudioTestPid, gAudioTestResultPath

    if (gAudioTestPid && ProcessExist(gAudioTestPid))
        return

    if AudioIsRunning() {
        SetAudioTestStatus(
            "Stop Audio Translation before testing."
        )
        return
    }

    px := ResolvePath(pythonExe)
    ap := ResolvePath(audioScript)
    if !(FileExist(px) && FileExist(ap)) {
        SetAudioTestStatus(
            "Could not open the selected device."
        )
        return
    }

    spick := Trim(ddlSpeaker.Text)
    if (spick = "" || spick = "[Windows Default]")
        EnvSet("SPEAKER_NAME", "")
    else
        EnvSet("SPEAKER_NAME", spick)

    resultPath := A_Temp "\jrpg_audio_test_"
        . A_TickCount "_" Random(1000, 9999) ".txt"
    try FileDelete(resultPath)
    EnvSet("AUDIO_TEST_RESULT_FILE", resultPath)

    SetAudioTestStatus("Listening for audio...")
    try btnAudioTest.Enabled := false
    try {
        testPid := 0
        Run(
            '"' px '" "' ap '" --test-audio',
            A_ScriptDir,
            "Hide",
            &testPid
        )
        gAudioTestPid := testPid
    } catch as ex {
        EnvSet("AUDIO_TEST_RESULT_FILE", "")
        gAudioTestPid := 0
        gAudioTestResultPath := ""
        try btnAudioTest.Enabled := true
        SetAudioTestStatus(
            "Could not open the selected device."
        )
        DbgCP("Audio test launch failed: " ex.Message)
        return
    }

    EnvSet("AUDIO_TEST_RESULT_FILE", "")
    gAudioTestResultPath := resultPath
    SetTimer(AudioTestPoll, 100)
}

AudioTestPoll() {
    global gAudioTestPid, gAudioTestResultPath, btnAudioTest

    if (gAudioTestPid && ProcessExist(gAudioTestPid))
        return

    SetTimer(AudioTestPoll, 0)
    gAudioTestPid := 0
    output := ""
    try output := FileRead(gAudioTestResultPath, "UTF-8")
    try FileDelete(gAudioTestResultPath)
    gAudioTestResultPath := ""
    try btnAudioTest.Enabled := true

    if InStr(output, "JRPG_AUDIO_TEST:DETECTED") {
        SetAudioTestStatus(
            "Audio detected. This device is ready."
        )
        DbgCP("Audio test detected signal: " Trim(output))
    } else if InStr(output, "JRPG_AUDIO_TEST:SILENT") {
        SetAudioTestStatus(
            "No audio detected. Check the device or application audio driver."
        )
        DbgCP("Audio test found silence: " Trim(output))
    } else if InStr(output, "JRPG_AUDIO_TEST:ERROR") {
        detail := RegExReplace(
            Trim(output),
            "^JRPG_AUDIO_TEST:ERROR:\s*"
        )
        SetAudioTestStatus(
            "Device error: " (detail != "" ? detail : "Could not open the selected device.")
        )
        DbgCP("Audio test device error: " Trim(output))
    } else {
        SetAudioTestStatus(
            "Audio test ended without returning a result."
        )
        DbgCP("Audio test failed: " Trim(output))
    }
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
    global explainScript
	global explainProvider, explainOpenAIModel, explainGeminiModel
    global ddlEProv, ddlEOpenAI, ddlEGem
    global ddlTR,ddlAProv,ddlA_GM,ddlAudioTarget,ddlProv,ddlIMG,ddlIMG_GM,slTrans
	global ddlFont, edFSize, chkFontBold, fontName, fontSize, fontBold
    global ddlPrompt, promptProfile
    global imgPostproc, directModelOutput, cbDirectModelOutput
	global debugMode, cbDebug
    overlayTrans     := slTrans.Value
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
    directModelOutput := cbDirectModelOutput.Value ? 1 : 0
    imgPostproc      := SyncPromptPostproc(promptProfile)
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

CPConfigurePreferredPanelWidth() {
    global tab, CPTabNaturalWidths, CPPreferredViewportW, defGuiW
    tab.GetPos(&tabX)
    naturalTabWidth := 0
    for tabWidth in CPTabNaturalWidths
        naturalTabWidth += tabWidth
    ; This is the exact client width where the custom titles still use their
    ; natural centered layout; one pixel less activates compact alignment.
    CPPreferredViewportW := Round(naturalTabWidth + tabX * 2)
    defGuiW := CPPreferredViewportW
}

CPConfigurePreferredPanelHeight() {
    global chkTop_TW, pad, CP_VIEWPORT_MIN_H, CP_CANVAS_MIN_H
    global CPPreferredViewportH, defGuiH

    chkTop_TW.GetPos(, &toggleY, , &toggleH)
    footerSpan := pad + (32 * 2 + 10 + 12 + pad) - 4
    CPPreferredViewportH := Max(CP_VIEWPORT_MIN_H, Round(toggleY + toggleH + 24 + footerSpan))
    ; The preferred snap height is also the no-scroll design height. A vertical
    ; scrollbar appears only after the visible client area becomes shorter.
    CP_CANVAS_MIN_H := CPPreferredViewportH
    defGuiH := CPPreferredViewportH
}

CPLayoutControllerOptions(rightEdge) {
    global controllerOptionsX, controllerTopY
    global txtControllerStatus, cbControllerInputsEnabled
    global cbControllerDpadNavigationEnabled, txtControllerDpadNote

    if !(IsSet(txtControllerStatus) && IsSet(cbControllerInputsEnabled)
        && IsSet(cbControllerDpadNavigationEnabled) && IsSet(txtControllerDpadNote))
        return

    ; Shrink the right-hand column before allowing the canvas to clip it.
    ; Multiline checkbox labels and a taller note use the free space below.
    optionsW := Min(350, Max(160, rightEdge - controllerOptionsX))
    statusH := (optionsW < 220) ? 38 : 20
    checkboxH := (optionsW < 300) ? 42 : 28
    optionGap := (optionsW < 300) ? 4 : 10

    txtControllerStatus.Move(controllerOptionsX, controllerTopY, optionsW, statusH)
    directY := controllerTopY + statusH + 8
    cbControllerInputsEnabled.Move(controllerOptionsX, directY, optionsW, checkboxH)
    dpadY := directY + checkboxH + optionGap
    cbControllerDpadNavigationEnabled.Move(controllerOptionsX, dpadY, optionsW, checkboxH)

    noteY := dpadY + checkboxH + 8
    noteH := CPMeasureWrappedTextHeight(txtControllerDpadNote, optionsW) + 6
    txtControllerDpadNote.Move(controllerOptionsX, noteY, optionsW, noteH)
}

CPMeasureWrappedTextHeight(ctrl, width) {
    fallbackH := 120
    if !(ctrl && ctrl.Hwnd)
        return fallbackH

    dc := DllCall("user32\GetDC", "ptr", ctrl.Hwnd, "ptr")
    if !dc
        return fallbackH

    oldFont := 0
    font := SendMessage(0x0031, 0, 0, ctrl.Hwnd) ; WM_GETFONT
    if font
        oldFont := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", font, "ptr")

    rect := Buffer(16, 0)
    NumPut("int", Max(1, Round(width)), rect, 8)
    ; DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX
    DllCall("user32\DrawTextW", "ptr", dc, "wstr", ctrl.Text, "int", -1
        , "ptr", rect.Ptr, "uint", 0x00000410 | 0x00000800)
    measuredH := Max(1, NumGet(rect, 12, "int") - NumGet(rect, 4, "int"))

    if oldFont
        DllCall("gdi32\SelectObject", "ptr", dc, "ptr", oldFont, "ptr")
    DllCall("user32\ReleaseDC", "ptr", ctrl.Hwnd, "ptr", dc)
    return measuredH
}

ResizeUI(gui, minMax, w, h){
    global CPPanelInteractiveResize, CPPanelLastLayoutW, CPPanelLastLayoutH
    if (minMax = -1 || w <= 0 || h <= 0)
        return

    layoutW := Round(w)
    layoutH := Round(h)
    ; Hidden startup measurement and DPI correction can queue duplicate Size
    ; events. Avoid repainting the identical layout a second time.
    if (layoutW = CPPanelLastLayoutW && layoutH = CPPanelLastLayoutH)
        return

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
    global tab, CPFooterFill, sepAction
    global tPython,ePython,bPy,tAud,eAudio,bAud,tOv,eOverlay,bOvSel,tImg,eImg,bImgSel,tExplain,eExplain,bExplainSel
    global ddlSpeaker,ddlAProv,ddlA_GM,ddlTR,ddlAudioTarget,ddlProv,ddlIMG,ddlIMG_GM
    global txtAudioHelp,txtAudioTestStatus
    global btnStart,btnStop,btnAudio,btnOv,btnOvClose,btnExplainerLaunch,btnExplainerClose,btnExplainNow,bClose,chkTop,chkDarkMode,chkGuess
    global txtControlOpacity,slControlOpacity,lblControlOpacityPct
    global controllerOptionsX, controllerTopY
    global txtControllerStatus, cbControllerInputsEnabled
    global cbControllerDpadNavigationEnabled, txtControllerDpadNote
    global btnA_GM_Add, btnA_GM_Del, btnTR_Add, btnTR_Del
    global btnIMG_Add, btnIMG_Del, btnIMG_GM_Add, btnIMG_GM_Del
    ; NEW prompt widgets
        ; NEW prompt widgets
    global ddlPrompt, btnPrEdit, btnPrNew, btnPrDel
    ; AUDIO target language widget
    global ddlAudioTarget

    browseW := 80
    btnH    := 32

    gap1 := 10
    gap2 := 12
    bottomBlockH := btnH*2 + gap1 + gap2 + pad

    ; Header and footer follow the visible client area. Only page content keeps
    ; the larger virtual-canvas height and scrolls between these fixed regions.
    tabH := Max(260, viewportH - pad*2 - bottomBlockH)
    tab.Move(pad, pad, w - pad*2, tabH)
    CPLayoutCustomTabBar(viewportW)

    tab.GetPos(&tx,&ty,&tw,&th)
    rightEdge := tx + tw - (pad + 28)
    CPLayoutControllerOptions(rightEdge)

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
    for audioText in [txtAudioHelp, txtAudioTestStatus] {
        audioText.GetPos(&audioTextX)
        audioText.Move(, , Max(300, rightEdge - audioTextX))
    }
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

    ; Place the sticky footer directly below the visible tab viewport.
    sepY := pad + tabH + 4
    CPFooterFill.Move(0, sepY, w, Max(1, viewportH - sepY))
    CPFooterFill.Visible := true
    try sepAction.Move(pad, sepY, w - pad*2, 2)

    ; action row always sits just below the separator
    yAction := sepY + 8
    ; Order (left -> right): Open, Close, Audio Translation, Open Explainer, Close Explainer
    btnOv.Move(pad, yAction, 130, btnH)                                ; Open Translator
    btnOvClose.Move(pad + 130 + gap, yAction, 140, btnH)               ; Close Translator
    btnAudio.Move(pad + 130 + 140 + gap*2, yAction, 180, btnH)          ; Audio Translation On/Off
    btnExplainerLaunch.Move(pad + 130 + 140 + 180 + gap*3, yAction, 140, btnH)  ; Open Explainer
    btnExplainerClose.Move(pad + 130 + 140 + 180 + 140 + gap*4, yAction, 140, btnH) ; Close Explainer

    ySave := viewportH - btnH - pad

    bClose.Move(pad, ySave, 120, btnH)
    bClose.GetPos(&cx,&cy,&btnW,)
    chkTop.Move(cx + btnW + 12, ySave + 6)
    chkTop.GetPos(&cpTopX, &cpTopY, &cpTopW)
    chkDarkMode.Move(cpTopX + cpTopW + 18, ySave + 6)
    chkDarkMode.GetPos(&cpDarkX, &cpDarkY, &cpDarkW)
    txtControlOpacity.Move(cpDarkX + cpDarkW + 18, ySave + 6)
    txtControlOpacity.GetPos(&cpOpacityLabelX, &cpOpacityLabelY, &cpOpacityLabelW)
    slControlOpacity.Move(cpOpacityLabelX + cpOpacityLabelW + 8, ySave + 1, 110, 24)
    slControlOpacity.GetPos(&cpOpacitySliderX, &cpOpacitySliderY, &cpOpacitySliderW)
    lblControlOpacityPct.Move(cpOpacitySliderX + cpOpacitySliderW + 8, ySave + 6, 44)
    CPUpdateComboArrowOverlays()
    CPCanvasFinishLayout(viewportW, viewportH, canvasRestore, false)
    CPClipScrollableControlsToViewport(false)
    CPMaintainFooterZOrder()

    ; --- Re-enable redraw and force repaint of the whole window and all children ---
    if (hwnd) {
        ; WM_SETREDRAW â†’ 1 = resume redraw (WinAPI with HWND)
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x000B, "ptr", 1, "ptr", 0)
        if CPPanelInteractiveResize {
            ; Let Windows combine successive mouse-move paints. A synchronous
            ; erase/update for every pixel is what makes a layered panel flash.
            DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0
                , "uint", 0x0001 | 0x0020 | 0x0080) ; INVALIDATE|NOERASE|ALLCHILDREN
        } else {
            ; Programmatic/finished layouts still receive one complete repaint.
            DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0
                , "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100)
        }
    }
    CPPanelLastLayoutW := layoutW
    CPPanelLastLayoutH := layoutH
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

CPControllerColorKeyboardNavigate(direction, hueSlider, saturationSlider, brightnessSlider
    , applyButton, cancelButton, *) {
    if CPControllerKeyboardMirrorActive(direction)
        return
    cpColorKeyboardStep := CPControllerColorNavigate.Bind(direction, hueSlider
        , saturationSlider, brightnessSlider, applyButton, cancelButton)
    if CPKeyboardSliderRepeatStart(direction, cpColorKeyboardStep)
        return
    cpColorKeyboardStep.Call()
}

CPControllerColorActivate(hueSlider, saturationSlider, brightnessSlider, applyButton, cancelButton, *) {
    focusHwnd := DllCall("user32\GetFocus", "ptr")
    if (focusHwnd = applyButton.Hwnd)
        SendMessage(0x00F5, 0, 0, applyButton.Hwnd) ; BM_CLICK
    else if (focusHwnd = cancelButton.Hwnd)
        SendMessage(0x00F5, 0, 0, cancelButton.Hwnd)
}

CPControllerColorSliderRepeatActive(targetHwnd, direction) {
    global CPControllerColorDialogState
    if (direction != "Left" && direction != "Right")
        return false
    state := CPControllerColorDialogState
    if !(state.Has("active") && state["active"]
     && state.Has("hwnd") && state["hwnd"] = targetHwnd)
        return false

    focusHwnd := DllCall("user32\GetFocus", "ptr")
    for sliderKey in ["hue", "saturation", "brightness"] {
        try {
            if (focusHwnd = state[sliderKey].Hwnd)
                return true
        }
    }
    return false
}

CPControllerAcceleratedSliderRepeatActive(targetHwnd, direction) {
    global ui, slTrans, slTrans_EW, slControlOpacity, eCapMax
    cpAcceleratedFocusHwnd := DllCall("user32\GetFocus", "ptr")
    ; Font size always advances in single-point steps, even while the D-pad is held.
    if CPFontSizeAdjustActive()
        return false
    try {
        if (targetHwnd = ui.Hwnd && CPMaxPngAdjustActive()
         && cpAcceleratedFocusHwnd = eCapMax.Hwnd)
            return true
    }
    if (direction != "Left" && direction != "Right")
        return false

    try {
        if (targetHwnd = ui.Hwnd) {
            for cpAcceleratedSliderCtrl in [slTrans, slTrans_EW, slControlOpacity] {
                if (cpAcceleratedFocusHwnd = cpAcceleratedSliderCtrl.Hwnd)
                    return true
            }
            return false
        }
    }
    return CPControllerColorSliderRepeatActive(targetHwnd, direction)
}

CPShowDialogFocusCues(dialogHwnd) {
    if !dialogHwnd
        return
    ; Direct controller focus changes do not make Windows reveal its native
    ; slider/button focus rectangles the way a keyboard navigation message does.
    try SendMessage(0x0127, 0x00030002, 0, dialogHwnd) ; WM_CHANGEUISTATE, UIS_CLEAR, HIDEFOCUS|HIDEACCEL
    try DllCall("user32\RedrawWindow", "ptr", dialogHwnd, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0080 | 0x0100) ; INVALIDATE | ALLCHILDREN | UPDATENOW
}

CPControllerColorDispatch(command, targetHwnd) {
    global CPControllerColorDialogState
    state := CPControllerColorDialogState
    if !(state.Has("active") && state["active"]
     && state.Has("hwnd") && state["hwnd"] = targetHwnd)
        return false

    CPShowDialogFocusCues(targetHwnd)
    try {
        switch command {
            case "Up", "Down", "Left", "Right":
                CPControllerColorNavigate(command, state["hue"], state["saturation"]
                    , state["brightness"], state["apply"], state["cancel"])
            case "Activate":
                CPControllerColorActivate(state["hue"], state["saturation"]
                    , state["brightness"], state["apply"], state["cancel"])
            case "Cancel":
                SendMessage(0x00F5, 0, 0, state["cancel"].Hwnd) ; BM_CLICK
        }
    }
    return true
}

CPControllerColorDialog(initHex, dialogTitle := "Adjust color") {
    global ui, CPControllerColorDialogState
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
    CPControllerColorDialogState := Map(
        "active", true,
        "hwnd", dlg.Hwnd,
        "hue", hueSlider,
        "saturation", saturationSlider,
        "brightness", brightnessSlider,
        "apply", applyButton,
        "cancel", cancelButton)

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
        try Hotkey(keyName, CPControllerColorKeyboardNavigate.Bind(direction, hueSlider, saturationSlider
            , brightnessSlider, applyButton, cancelButton), "On")
    try Hotkey("$Left Up", CPKeyboardSliderRepeatRelease.Bind("Left"), "On")
    try Hotkey("$Right Up", CPKeyboardSliderRepeatRelease.Bind("Right"), "On")
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
        CPShowDialogFocusCues(dlg.Hwnd)
        while !closed
            Sleep(30)
    } finally {
        CPControllerColorDialogState := Map("active", false)
        CPKeyboardSliderRepeatStop()
        CPControllerResetNavigation()
        CPUnregisterColorSwatch(previewHwnd)
        CPUnregisterControllerColorGradients(gradientSliderHwnds*)
        HotIfWinActive(dialogHotIf)
        for keyName, direction in dialogArrowHotkeys
            try Hotkey(keyName, "Off")
        try Hotkey("$Left Up", "Off")
        try Hotkey("$Right Up", "Off")
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
    global fontSize, CPFontSizeAdjustSyncing
    if CPFontSizeAdjustSyncing
        return
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
    global fontSize_EW, CPFontSizeAdjustSyncing
    if CPFontSizeAdjustSyncing
        return
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
    ib := CPThemedInputBox("Enter a name for the new EXPLANATION prompt:", "New EXPLAIN prompt")
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
    ib := CPThemedInputBox(
        "Enter a name for the new prompt profile:",
        "New prompt",
        "To show a transcript in the Translation overlay, include`n"
            . "with_transcript or with_kanji_reading in the name.",
        "",
        520
    )
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

GlossaryPath(kind, name) {
    return (kind = "jp") ? GlossaryJP2ENPath(name) : GlossaryEN2ENPath(name)
}

GlossaryKindLabel(kind) {
    return (kind = "jp") ? "JP -> TL" : "TL -> TL"
}

GlossaryHeader(kind) {
    return "# Managed by JRPG Translator - " GlossaryKindLabel(kind)
        . " terminology overrides`r`n"
}

ListGlossaryProfiles(kind := "jp") {
    global glossariesDir
    out := []
    fileName := (kind = "jp") ? "jp2en.txt" : "en2en.txt"

    ; Each glossary type has its own profile list. Existing folders that contain
    ; both files remain valid and simply appear in both lists.
    if DirExist(glossariesDir) {
        Loop Files glossariesDir "\*", "D" {
            prof := A_LoopFileName
            if FileExist(glossariesDir "\" prof "\" fileName)
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

    ddlJPG.Add(ListGlossaryProfiles("jp"))
    ddlENG.Add(ListGlossaryProfiles("en"))

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

GlossaryReadDocument(path) {
    doc := Map("entries", [], "metadata", [], "malformed", [])
    if !FileExist(path)
        return doc

    text := ""
    try text := FileRead(path, "UTF-8")
    catch {
        try text := FileRead(path)
    }

    separators := ["->", Chr(0x2192), "â†’", "`t", ":", "="]
    for rawLine in StrSplit(text, "`n", "`r") {
        line := Trim(StrReplace(rawLine, Chr(0xFEFF), ""))
        if (line = "") {
            doc["metadata"].Push("")
            continue
        }
        if (SubStr(line, 1, 1) = "#") {
            doc["metadata"].Push(rawLine)
            continue
        }

        source := ""
        target := ""
        for separator in separators {
            separatorPos := InStr(line, separator)
            if !separatorPos
                continue
            source := Trim(SubStr(line, 1, separatorPos - 1))
            target := Trim(SubStr(line, separatorPos + StrLen(separator)))
            break
        }

        if (source != "" && target != "")
            doc["entries"].Push(Map("source", source, "target", target))
        else
            doc["malformed"].Push(rawLine)
    }
    return doc
}

GlossaryCloneEntries(entries) {
    cloned := []
    for entry in entries
        cloned.Push(Map("source", entry["source"], "target", entry["target"]))
    return cloned
}

GlossaryNormalizeSource(source) {
    return StrLower(Trim(source))
}

GlossaryDuplicateSummary(entries) {
    seen := Map()
    duplicates := []
    for rowIndex, entry in entries {
        key := GlossaryNormalizeSource(entry["source"])
        if seen.Has(key)
            duplicates.Push("'" entry["source"] "' (rows " seen[key] " and " rowIndex ")")
        else
            seen[key] := rowIndex
    }
    return duplicates
}

GlossaryValidateEntries(entries, allowDuplicates := false) {
    seen := Map()
    for rowIndex, entry in entries {
        source := Trim(entry["source"])
        target := Trim(entry["target"])
        if (source = "" || target = "")
            return "Both fields are required (row " rowIndex ")."
        if InStr(source, "`r") || InStr(source, "`n")
         || InStr(target, "`r") || InStr(target, "`n")
            return "Entries must stay on one line (row " rowIndex ")."
        if InStr(source, "->") || InStr(source, Chr(0x2192))
         || InStr(source, "â†’") || InStr(source, "`t")
            return "The source field cannot contain an arrow or tab separator (row " rowIndex ")."

        key := GlossaryNormalizeSource(source)
        if seen.Has(key) && !allowDuplicates
            return "Duplicate source '" source "' in rows " seen[key] " and " rowIndex "."
        seen[key] := rowIndex
    }
    return ""
}

GlossaryBuildText(doc, entries, kind) {
    metadata := []
    hasComment := false
    for line in doc["metadata"] {
        cleanLine := RTrim(line, "`r`n")
        metadata.Push(cleanLine)
        if (SubStr(Trim(cleanLine), 1, 1) = "#")
            hasComment := true
    }
    while (metadata.Length && Trim(metadata[-1]) = "")
        metadata.Pop()
    if !hasComment
        metadata.InsertAt(1, RTrim(GlossaryHeader(kind), "`r`n"))

    output := ""
    for line in metadata
        output .= line "`r`n"
    if (metadata.Length && entries.Length)
        output .= "`r`n"
    for entry in entries
        output .= Trim(entry["source"]) " -> " Trim(entry["target"]) "`r`n"
    return output
}

GlossaryEnsureFile(kind, profile) {
    path := GlossaryPath(kind, profile)
    if FileExist(path)
        return path
    if !DirExist(GlossaryProfileDir(profile))
        DirCreate(GlossaryProfileDir(profile))
    SaveTextAtomic(path, GlossaryHeader(kind), false)
    return path
}

OpenRawGlossaryEditor(kind := "jp", profile := "") {
    global ddlJPG, ddlENG
    prof := Trim(profile)
    if (prof = "")
        prof := (kind = "jp") ? Trim(ddlJPG.Text) : Trim(ddlENG.Text)
    if (prof = "")
        prof := "default"

    title := "Repair raw " GlossaryKindLabel(kind) " glossary - " prof
    path := GlossaryEnsureFile(kind, prof)

    txt := ""
    try txt := FileRead(path, "UTF-8")

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

GlossaryManagerRegistry() {
    static registry := Map()
    return registry
}

GlossaryManagerRegister(state) {
    static messageRegistered := false
    if !messageRegistered {
        OnMessage(0x0100, GlossaryManagerOnKeyDown) ; WM_KEYDOWN
        messageRegistered := true
    }
    GlossaryManagerRegistry()[state["gui"].Hwnd] := state
}

GlossaryManagerRefresh(state, preferredRow := 0) {
    lv := state["list"]
    lv.Delete()
    for entry in state["entries"]
        lv.Add("", entry["source"], entry["target"])

    lv.ModifyCol(1, state["kind"] = "jp" ? 280 : 300)
    lv.ModifyCol(2, "AutoHdr")
    hasEntries := state["entries"].Length > 0
    state["editButton"].Enabled := hasEntries
    state["deleteButton"].Enabled := hasEntries

    duplicates := GlossaryDuplicateSummary(state["entries"])
    status := state["entries"].Length " entr" (state["entries"].Length = 1 ? "y" : "ies")
    if duplicates.Length
        status .= "  |  Duplicate sources: " duplicates.Length
    state["status"].Text := status

    if hasEntries {
        row := preferredRow ? Min(preferredRow, state["entries"].Length) : 1
        lv.Modify(row, "Select Focus Vis")
    }
}

GlossaryManagerResize(state, guiObj, minMax, width, height) {
    if (minMax = -1)
        return
    listHeight := Max(170, height - 150)
    state["list"].Move(, , Max(480, width - 32), listHeight)
    buttonY := height - 52
    state["addButton"].Move(16, buttonY, 100, 32)
    state["editButton"].Move(126, buttonY, 100, 32)
    state["deleteButton"].Move(236, buttonY, 100, 32)
    state["closeButton"].Move(Max(346, width - 116), buttonY, 100, 32)
}

GlossaryManagerClose(state, *) {
    registry := GlossaryManagerRegistry()
    guiHwnd := state["gui"].Hwnd
    if registry.Has(guiHwnd)
        registry.Delete(guiHwnd)
    try state["gui"].Destroy()
}

GlossaryManagerOnKeyDown(wParam, lParam, msg, hwnd) {
    registry := GlossaryManagerRegistry()
    focusedHwnd := DllCall("user32\GetFocus", "ptr")
    if !focusedHwnd
        return
    rootHwnd := DllCall("user32\GetAncestor", "ptr", focusedHwnd, "uint", 2, "ptr") ; GA_ROOT
    if !registry.Has(rootHwnd)
        return

    state := registry[rootHwnd]
    lv := state["list"]
    if (focusedHwnd = lv.Hwnd) {
        selectedRow := lv.GetNext()
        if (wParam = 0x0D || wParam = 0x71) { ; Enter or F2
            if selectedRow
                GlossaryManagerEdit(state)
            else
                GlossaryManagerAdd(state)
            return 0
        }
        if (wParam = 0x2D) { ; Insert
            GlossaryManagerAdd(state)
            return 0
        }
        if (wParam = 0x2E && selectedRow) { ; Delete
            GlossaryManagerDelete(state)
            return 0
        }
        if (wParam = 0x28
            && (!state["entries"].Length || selectedRow = state["entries"].Length)) {
            state["addButton"].Focus()
            return 0
        }
        return
    }

    buttons := [state["addButton"], state["editButton"], state["deleteButton"], state["closeButton"]]
    buttonIndex := 0
    for index, buttonCtrl in buttons {
        if (buttonCtrl.Hwnd = focusedHwnd) {
            buttonIndex := index
            break
        }
    }
    if !buttonIndex
        return
    if (wParam = 0x26) { ; Up returns to the table
        lv.Focus()
        return 0
    }
    if (wParam = 0x25 || wParam = 0x27) {
        direction := (wParam = 0x25) ? -1 : 1
        nextIndex := buttonIndex + direction
        while (nextIndex >= 1 && nextIndex <= buttons.Length
            && !buttons[nextIndex].Enabled)
            nextIndex += direction
        if (nextIndex >= 1 && nextIndex <= buttons.Length)
            buttons[nextIndex].Focus()
        return 0
    }
}

GlossaryOwnedMessage(ownerHwnd, message, title := "Terminology overrides"
    , buttons := "ok", icon := "warning") {
    ; MsgBox ownership can be lost when a GUI event passes through a shared
    ; helper. Use the native owner handle explicitly and make this short-lived
    ; message topmost so it cannot fall behind the control panel or entry dialog.
    if !ownerHwnd || !DllCall("user32\IsWindow", "ptr", ownerHwnd, "int")
        ownerHwnd := DllCall("user32\GetForegroundWindow", "ptr")

    flags := 0x00010000 | 0x00040000 ; MB_SETFOREGROUND | MB_TOPMOST
    if (buttons = "yesno")
        flags |= 0x00000004 ; MB_YESNO
    if (icon = "error")
        flags |= 0x00000010 ; MB_ICONERROR
    else if (icon = "info")
        flags |= 0x00000040 ; MB_ICONINFORMATION
    else
        flags |= 0x00000030 ; MB_ICONWARNING

    return DllCall("user32\MessageBoxW"
        , "ptr", ownerHwnd
        , "wstr", message
        , "wstr", title
        , "uint", flags
        , "int")
}

GlossaryManagerWriteCandidate(state, candidateEntries, preferredRow := 0
    , allowRemainingDuplicates := false, ownerHwnd := 0) {
    validationError := GlossaryValidateEntries(candidateEntries, allowRemainingDuplicates)
    if (validationError != "") {
        GlossaryOwnedMessage(ownerHwnd, validationError)
        return false
    }

    try SaveTextAtomic(
        state["path"],
        GlossaryBuildText(state["doc"], candidateEntries, state["kind"])
    )
    catch as ex {
        GlossaryOwnedMessage(ownerHwnd,
            "Could not save the glossary:`n`n" ex.Message,
            "Terminology overrides", "ok", "error")
        return false
    }

    state["entries"] := candidateEntries
    GlossaryManagerRefresh(state, preferredRow)
    Toast("Saved " GlossaryKindLabel(state["kind"]) " entries for '" state["profile"] "'")
    return true
}

GlossaryEntryDialogClose(dialogState, managerGui, dialogGui, *) {
    if dialogState["closed"]
        return
    dialogState["closed"] := true
    try managerGui.Opt("-Disabled")
    try dialogGui.Destroy()
    try managerGui.Show()
}

GlossaryEntryDialogAccept(state, rowIndex, sourceEdit, targetEdit
    , dialogState, managerGui, dialogGui, *) {
    source := Trim(sourceEdit.Value)
    target := Trim(targetEdit.Value)
    candidateEntries := GlossaryCloneEntries(state["entries"])
    newEntry := Map("source", source, "target", target)
    if rowIndex
        candidateEntries[rowIndex] := newEntry
    else
        candidateEntries.Push(newEntry)

    preferredRow := rowIndex ? rowIndex : candidateEntries.Length
    originalDuplicateCount := GlossaryDuplicateSummary(state["entries"]).Length
    candidateDuplicateCount := GlossaryDuplicateSummary(candidateEntries).Length
    allowRemainingDuplicates := rowIndex && originalDuplicateCount
        && candidateDuplicateCount < originalDuplicateCount
    if !GlossaryManagerWriteCandidate(
        state, candidateEntries, preferredRow, allowRemainingDuplicates, dialogGui.Hwnd)
        return
    GlossaryEntryDialogClose(dialogState, managerGui, dialogGui)
}

GlossaryEntryDialog(state, rowIndex := 0) {
    global ui
    managerGui := state["gui"]
    isEdit := rowIndex > 0
    sourceValue := isEdit ? state["entries"][rowIndex]["source"] : ""
    targetValue := isEdit ? state["entries"][rowIndex]["target"] : ""
    kindLabel := GlossaryKindLabel(state["kind"])
    title := (isEdit ? "Edit " : "Add ") kindLabel " entry"

    try managerGui.Opt("+Disabled")
    ; Own the editor directly to the control panel so its dark-mode brushes and
    ; controller-navigation owner-chain handling are identical to other dialogs.
    dlg := Gui("+Owner" ui.Hwnd " +AlwaysOnTop +OwnDialogs", title)
    dlg.MarginX := 18
    dlg.MarginY := 16
    dlg.SetFont("s10", "Segoe UI")
    sourceLabel := (state["kind"] = "jp") ? "Japanese source term:" : "Translation output to replace:"
    dlg.Add("Text", "xm w500", sourceLabel)
    sourceEdit := dlg.Add("Edit", "xm y+6 w500", sourceValue)
    dlg.Add("Text", "xm y+14 w500", "Target-language replacement:")
    targetEdit := dlg.Add("Edit", "xm y+6 w500", targetValue)
    hint := dlg.Add("Text", "xm y+12 w500 cGray",
        "Both fields are required. Matching uses the complete source text you enter.")
    CPRegisterMutedControl(hint)
    btnSave := dlg.Add("Button", "xm y+16 w120 Default", isEdit ? "Save" : "Add")
    btnCancel := dlg.Add("Button", "x+10 yp w120", "Cancel")
    dialogState := Map("closed", false)

    btnSave.OnEvent("Click", GlossaryEntryDialogAccept.Bind(
        state, rowIndex, sourceEdit, targetEdit, dialogState, managerGui, dlg))
    closeCallback := GlossaryEntryDialogClose.Bind(dialogState, managerGui, dlg)
    btnCancel.OnEvent("Click", closeCallback)
    dlg.OnEvent("Escape", closeCallback)
    dlg.OnEvent("Close", closeCallback)
    dlg.Show("AutoSize Center")
    CPApplyOwnedDialogTheme(dlg)
    try sourceEdit.Focus()
}

GlossaryManagerAdd(state, *) {
    GlossaryEntryDialog(state)
}

GlossaryManagerEdit(state, listCtrl := 0, eventRow := 0, *) {
    row := eventRow
    if !row
        row := state["list"].GetNext()
    if !row {
        if !state["entries"].Length
            GlossaryManagerAdd(state)
        else
            GlossaryOwnedMessage(state["gui"].Hwnd, "Select an entry to edit.")
        return
    }
    GlossaryEntryDialog(state, row)
}

GlossaryManagerDelete(state, *) {
    row := state["list"].GetNext()
    if !row {
        GlossaryOwnedMessage(state["gui"].Hwnd, "Select an entry to delete.")
        return
    }
    entry := state["entries"][row]
    if (GlossaryOwnedMessage(state["gui"].Hwnd,
        "Delete this entry?`n`n" entry["source"] "  ->  " entry["target"],
        "Terminology overrides", "yesno") != 6) ; IDYES
        return

    candidateEntries := GlossaryCloneEntries(state["entries"])
    candidateEntries.RemoveAt(row)
    ; Deleting is always allowed to reduce or remove duplicate legacy rows.
    GlossaryManagerWriteCandidate(state, candidateEntries, row, true, state["gui"].Hwnd)
}

OpenGlossaryManager(kind := "jp") {
    global ui, ddlJPG, ddlENG
    profile := (kind = "jp") ? Trim(ddlJPG.Text) : Trim(ddlENG.Text)
    if (profile = "")
        profile := "default"
    path := GlossaryEnsureFile(kind, profile)
    doc := GlossaryReadDocument(path)

    if doc["malformed"].Length {
        preview := ""
        for index, line in doc["malformed"] {
            if (index > 4) {
                preview .= "`n..."
                break
            }
            preview .= "`n" line
        }
        answer := GlossaryOwnedMessage(ui.Hwnd,
            doc["malformed"].Length " line(s) could not be read as terminology pairs."
            . "`n`nTable editing is disabled to avoid losing those lines."
            . "`nOpen the raw repair editor now?`n" preview,
            "Terminology overrides", "yesno")
        if (answer = 6) ; IDYES
            OpenRawGlossaryEditor(kind, profile)
        return
    }

    title := GlossaryKindLabel(kind) " terminology - " profile
    g := Gui("+Resize +Owner" ui.Hwnd " +OwnDialogs", title)
    g.MarginX := 16
    g.MarginY := 14
    g.SetFont("s10", "Segoe UI")
    description := (kind = "jp")
        ? "Exact Japanese terms sent to the translation model and their requested output."
        : "Translation outputs corrected locally after the model responds."
    g.Add("Text", "xm w748", description)
    columns := (kind = "jp")
        ? ["Japanese source", "Target-language replacement"]
        : ["Translation output", "Local replacement"]
    lv := g.Add("ListView", "xm y+10 w748 h350 Grid -Multi", columns)
    status := g.Add("Text", "xm y+7 w500 cGray", "")
    CPRegisterMutedControl(status)
    btnAdd := g.Add("Button", "xm y+10 w100", "Add...")
    btnEdit := g.Add("Button", "x+10 yp w100", "Edit...")
    btnDelete := g.Add("Button", "x+10 yp w100", "Delete")
    btnClose := g.Add("Button", "x+312 yp w100", "Close")

    state := Map(
        "kind", kind, "profile", profile, "path", path, "doc", doc,
        "entries", doc["entries"], "gui", g, "list", lv, "status", status,
        "addButton", btnAdd, "editButton", btnEdit,
        "deleteButton", btnDelete, "closeButton", btnClose
    )

    btnAdd.OnEvent("Click", GlossaryManagerAdd.Bind(state))
    btnEdit.OnEvent("Click", GlossaryManagerEdit.Bind(state))
    btnDelete.OnEvent("Click", GlossaryManagerDelete.Bind(state))
    btnClose.OnEvent("Click", GlossaryManagerClose.Bind(state))
    lv.OnEvent("DoubleClick", GlossaryManagerEdit.Bind(state))
    lv.OnEvent("ItemFocus", (*) => (
        state["editButton"].Enabled := state["list"].GetNext() > 0,
        state["deleteButton"].Enabled := state["list"].GetNext() > 0
    ))
    g.OnEvent("Escape", GlossaryManagerClose.Bind(state))
    g.OnEvent("Close", GlossaryManagerClose.Bind(state))
    g.OnEvent("Size", GlossaryManagerResize.Bind(state))

    GlossaryManagerRegister(state)
    GlossaryManagerRefresh(state)
    g.Opt("+MinSize620x340")
    g.Show("w780 h500 Center")
    CPApplyOwnedDialogTheme(g)
    try lv.Focus()
}

NewGlossaryProfile(kind := "jp", *) {
    global ddlJPG, ddlENG, iniPath, jp2enGlossaryProfile, en2enGlossaryProfile
    kindLabel := GlossaryKindLabel(kind)
    ib := CPThemedInputBox(
        "Enter a name for the new " kindLabel " profile:",
        "New " kindLabel " profile")
    if (ib.Result = "Cancel")
        return
    name := Trim(ib.Value)
    if (name = "") {
        MsgBox("Please enter a non-empty name.", "New " kindLabel " profile", "OK Icon!")
        return
    }
    name := RegExReplace(name, '[\\/:*?"<>|]+', "_")
    if RegExMatch(name, 'i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$')
        name := "_" name

    path := GlossaryPath(kind, name)
    if FileExist(path) {
        if (MsgBox(kindLabel " profile '" name "' already exists.`nOpen its entries?",
            "New " kindLabel " profile", "YesNo Icon!") = "Yes") {
            if (kind = "jp")
                jp2enGlossaryProfile := name
            else
                en2enGlossaryProfile := name
            IniWrite(name, iniPath, "cfg",
                kind = "jp" ? "jp2enGlossaryProfile" : "en2enGlossaryProfile")
            RefreshGlossaryProfilesList(jp2enGlossaryProfile, en2enGlossaryProfile)
            OpenGlossaryManager(kind)
        }
        return
    }

    GlossaryEnsureFile(kind, name)
    if (kind = "jp") {
        jp2enGlossaryProfile := name
        IniWrite(jp2enGlossaryProfile, iniPath, "cfg", "jp2enGlossaryProfile")
    } else {
        en2enGlossaryProfile := name
        IniWrite(en2enGlossaryProfile, iniPath, "cfg", "en2enGlossaryProfile")
    }
    RefreshGlossaryProfilesList(jp2enGlossaryProfile, en2enGlossaryProfile)
    OpenGlossaryManager(kind)
}

DeleteGlossaryProfile(kind := "jp", *) {
    global ddlJPG, ddlENG, iniPath, jp2enGlossaryProfile, en2enGlossaryProfile
    name := (kind = "jp") ? Trim(ddlJPG.Text) : Trim(ddlENG.Text)
    if (name = "") {
        MsgBox("No profile selected.",, "OK Icon!")
        return
    }
    if (name = "default") {
        MsgBox("The 'default' profile cannot be deleted.",, "OK Icon!")
        return
    }
    kindLabel := GlossaryKindLabel(kind)
    if (MsgBox("Delete the " kindLabel " profile '" name "'?"
        . "`n`nThe other glossary type will not be changed.",
        "Delete " kindLabel " profile", "YesNo Icon!") != "Yes")
        return

    dir := GlossaryProfileDir(name)
    path := GlossaryPath(kind, name)
    try FileDelete(path)
    catch as ex {
        MsgBox("Could not delete the profile:`n`n" ex.Message,
            "Delete " kindLabel " profile", "OK Icon!")
        return
    }
    partnerPath := GlossaryPath(kind = "jp" ? "en" : "jp", name)
    if !FileExist(partnerPath) {
        ; No active glossary remains in this profile folder.
        try DirDelete(dir, true)
    }

    if (kind = "jp") {
        jp2enGlossaryProfile := "default"
        IniWrite(jp2enGlossaryProfile, iniPath, "cfg", "jp2enGlossaryProfile")
    } else {
        en2enGlossaryProfile := "default"
        IniWrite(en2enGlossaryProfile, iniPath, "cfg", "en2enGlossaryProfile")
    }

    RefreshGlossaryProfilesList(jp2enGlossaryProfile, en2enGlossaryProfile)
    Toast("Deleted " kindLabel " profile '" name "'")
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

TerminologyOverridesChanged(*) {
    global chkUseTerminologyOverrides, useTerminologyOverrides, iniPath
    useTerminologyOverrides := chkUseTerminologyOverrides.Value ? 1 : 0
    IniWrite(useTerminologyOverrides, iniPath, "cfg", "useTerminologyOverrides")
    EnvSet("USE_TERMINOLOGY_OVERRIDES", useTerminologyOverrides ? "1" : "0")
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

ArrayIndexOf(arr, val) {
    for i, v in arr
        if (v = val)
            return i
    return 0
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
