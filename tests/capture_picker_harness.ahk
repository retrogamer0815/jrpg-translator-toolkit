#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn All, StdOut
; @RUNTIME_GLOBALS@
global ControlIni := A_ScriptDir "\fixture.ini"
global __CapPickFlagPath := A_ScriptDir "\controller_adjust.active"
global __HidSelf := false, __HidOther := false, __EXPLAIN_MODE := false
global Cap_MaxKB := 1400, Cap_Mode := "region", Cap_RectStr := "-29980,-29970,100,80"
global Cap_Rect := Map("x", -29980, "y", -29970, "w", 100, "h", 80), Cap_WinTit := ""
global fixtureFailHud := false, fixtureFailTick := false, fixtureProbeStartup := true
global fixtureChecks := 0, fixtureHandles := [], fixtureButtons := Map("confirm", false, "cancel", false)
global iniPath := ControlIni, CPBigBoxCaptureWatch := Map("active", false), fixtureCaptureNotice := ""
global Overlay := Gui("-Caption +ToolWindow -DPIScale", "Picker fixture translator")
global otherOverlay := Gui("-Caption +ToolWindow -DPIScale", "Picker fixture other overlay")
global foreignGui := Gui("-Caption +ToolWindow -DPIScale", "Unrelated fixture window")
OnMessage(0x0201, Region_LButtonDown)
OnMessage(0x0202, Region_LButtonUp)
OnExit(FixtureCleanup)
try {
    IniWrite(Cap_RectStr, ControlIni, "capture", "rect")
    Overlay.Show("NA x-31000 y-31000 w50 h50")
    otherOverlay.Show("NA x-31000 y-31100 w50 h50")
    foreignGui.Show("NA x-31000 y-31200 w50 h50")
    Check(Region_LButtonUp() = "", "inactive release passes through")
    Check(Region_LButtonDown() = "", "inactive press passes through")
    Loop 100 {
        StartPickRegion(1400, Mod(A_Index, 2))
        Check(__CapPickState["active"] && __Sel_Active && __CapPickBusy, "picker ready")
        TrackHandles()
        canvas := __Sel_Gui.Hwnd
        sequence := IniRead(ControlIni, "capture", "pickSeq", "")
        ; The activation click's release must never finish an unstarted drag.
        MouseMessage(0x0202, canvas, 100, 80)
        Check(__CapPickState["active"] && !__SelDragging, "orphan release ignored")
        Check(sequence = IniRead(ControlIni, "capture", "pickSeq", ""), "orphan does not persist")
        Check(Region_LButtonDown(0, 0, 0x0201, foreignGui.Hwnd) = "", "foreign press passes through")
        Check(!__SelDragging, "foreign press cannot start drag")
        StartPickWindow()
        StartPickRegion()
        Check(__Sel_Gui.Hwnd = canvas && __HidSelf && __HidOther, "duplicate requests keep original session")
        MouseMessage(0x0201, canvas, 240, 180)
        Check(__SelDragging, "native press starts drag")
        Check(Region_LButtonUp(0, 0, 0x0202, foreignGui.Hwnd) = "", "foreign release passes through")
        Check(__SelDragging, "foreign release cannot finish drag")
        MouseMessage(0x0202, canvas, 40, 30)
        Check(IniRead(ControlIni, "capture", "rect") = "-29960,-29970,200,150", "reverse drag uses message coordinates on negative monitor")
        Check(IniRead(ControlIni, "capture", "pickStatus") = "selected", "success result")
        AssertClean()
        Check(Region_LButtonUp(0, 0, 0x0202, canvas) = "", "late release is inert")
    }
    StartPickRegion()
    TrackHandles()
    MouseMessage(0x0201, __Sel_Gui.Hwnd, 40, 30)
    MouseMessage(0x0202, __Sel_Gui.Hwnd, 41, 31)
    Check(__CapPickState["active"] && !__SelDragging, "tiny drag remains retryable")
    CapPickCancel()
    AssertClean()
    Check(IniRead(ControlIni, "capture", "pickStatus") = "canceled", "cancel result")

    StartPickRegion()
    TrackHandles()
    DllCall("PostMessage", "ptr", __Sel_Gui.Hwnd, "uint", 0x0010, "ptr", 0, "ptr", 0)
    Sleep(100)
    AssertClean()
    Check(IniRead(ControlIni, "capture", "pickStatus") = "canceled", "native Close/Alt-F4 cancels instead of hiding an active picker")

    StartPickRegion()
    TrackHandles()
    CapPickConfirm()
    Check(__CapPickState["active"], "activation confirm debounce")
    __CapPickState["acceptAfter"] := 0
    CapPickArrow(18, 0)
    expectedX := __CapPickState["x"]
    CapPickConfirm()
    Check(Cap_Rect["x"] = expectedX, "keyboard adjustment/confirm")
    AssertClean()

    StartPickRegion()
    TrackHandles()
    __CapPickState["acceptAfter"] := 0
    fixtureButtons := Map("confirm", true, "cancel", false)
    CapPickTick()
    AssertClean()
    Check(IniRead(ControlIni, "capture", "pickStatus") = "selected", "controller confirm without stick movement")
    fixtureButtons := Map("confirm", false, "cancel", false)

    StartPickRegion()
    TrackHandles()
    fixtureButtons := Map("confirm", false, "cancel", true)
    CapPickTick()
    AssertClean()
    fixtureButtons := Map("confirm", false, "cancel", false)

    fixtureFailHud := true
    StartPickRegion()
    AssertClean()
    Check(IniRead(ControlIni, "capture", "pickStatus") = "failed", "partial startup reports failure")
    fixtureFailHud := false

    StartPickRegion()
    TrackHandles()
    fixtureFailTick := true
    CapPickTick()
    AssertClean()
    fixtureFailTick := false

    StartPickRegion()
    TrackHandles()
    originalIni := ControlIni, originalRect := Cap_RectStr
    ControlIni := A_ScriptDir ; a directory cannot be overwritten as an INI
    __CapPickState["acceptAfter"] := 0
    CapPickConfirm()
    AssertClean()
    Check(Cap_RectStr = originalRect, "failed write does not publish cached capture")
    ControlIni := originalIni

    StartPickRegion()
    TrackHandles()
    ControlIni := A_ScriptDir
    CapPickCancel() ; completion write also fails, but restoration must succeed
    AssertClean()
    ControlIni := originalIni

    Loop 5 {
        StartPickWindow()
        TrackHandles()
        Check(__CapPickState["active"], "window picker ready after region picker")
        __CapPickState["acceptAfter"] := 0
        CapPickConfirm()
        Check(Cap_Mode = "window", "window selection commits")
        AssertClean()
        StartPickRegion()
        TrackHandles()
        CapPickFailSafe()
        AssertClean()
    }
    for fixtureStatus, expectedNotice in Map("selected", "updated", "canceled", "cancelled", "failed", "failed") {
        CPBigBoxCaptureWatch := Map("active", true, "path", iniPath, "seq", "old"
            , "started", DllCall("kernel32\GetTickCount64", "uint64"))
        IniWrite(fixtureStatus, iniPath, "capture", "pickStatus")
        IniWrite("new-" fixtureStatus, iniPath, "capture", "pickSeq")
        CPBigBoxWatchCapture()
        Check(!CPBigBoxCaptureWatch["active"] && InStr(fixtureCaptureNotice, expectedNotice)
            , "Big Box receives " fixtureStatus " result without waiting for timeout")
    }
    Overlay.Hide()
    otherOverlay.Hide()
    StartPickRegion()
    TrackHandles()
    CapPickCancel()
    Check(!DllCall("IsWindowVisible", "ptr", Overlay.Hwnd), "already hidden translator stays hidden")
    Check(!DllCall("IsWindowVisible", "ptr", otherOverlay.Hwnd), "already hidden explainer stays hidden")
    FileAppend("PASS: " fixtureChecks " capture-picker checks; 100 repeated native-message drags, cancel/retry, controller/keyboard, window switching and failure cleanup.`n", "*")
    ExitApp(0)
} catch as fixtureFailure {
    FileAppend("FAIL: " fixtureFailure.Message " at " fixtureFailure.Line "`n" fixtureFailure.Stack "`n", "*")
    ExitApp(1)
}

Check(ok, message) {
    global fixtureChecks
    fixtureChecks++
    if !ok
        throw Error(message)
}
TrackHandles() {
    global fixtureHandles, __Sel_Gui, __SelBand, __CapPickState
    for guiObject in [__Sel_Gui, __SelBand, __CapPickState.Get("hud", 0)]
        if IsObject(guiObject)
            fixtureHandles.Push(guiObject.Hwnd)
}
AssertClean() {
    global fixtureHandles, __CapPickState, __CapPickBusy, __Sel_Active, __SelDragging
    global __CapPickFlagPath, __HidSelf, __HidOther, Overlay, otherOverlay
    Check(!__CapPickBusy && !__CapPickState["active"] && !__Sel_Active && !__SelDragging, "all input/session guards reset")
    Check(!CapPickHotIf(), "picker hotkeys disabled")
    Check(!FileExist(__CapPickFlagPath), "controller lock removed")
    Check(!__HidSelf && !__HidOther, "restore flags reset")
    Check(DllCall("IsWindowVisible", "ptr", Overlay.Hwnd), "translator restored")
    Check(DllCall("IsWindowVisible", "ptr", otherOverlay.Hwnd), "explainer restored")
    for hwnd in fixtureHandles
        Check(!DllCall("IsWindow", "ptr", hwnd), "picker GUI destroyed")
    fixtureHandles := []
}
MouseMessage(msg, hwnd, x, y) {
    DllCall("SendMessage", "ptr", hwnd, "uint", msg, "uptr", msg = 0x0201 ? 1 : 0
        , "ptr", (x & 0xFFFF) | ((y & 0xFFFF) << 16), "ptr")
}
CapPickVirtualBounds(&x, &y, &w, &h) {
    x := -30000, y := -30000, w := 640, h := 480
}
FixtureWorkArea(index, &left, &top, &right, &bottom) {
    left := -30000, top := -30000, right := -29360, bottom := -29520
}
FixtureCreateHud() {
    global fixtureFailHud, fixtureProbeStartup, __Sel_Gui, __SelDragging, __CapPickState, __CapPickBusy
    if fixtureProbeStartup && IsObject(__Sel_Gui) {
        Check(!__CapPickState["active"], "startup not accepting input")
        Check(Region_LButtonDown(0, 0, 0x0201, __Sel_Gui.Hwnd) = "", "early press ignored")
        Check(Region_LButtonUp(0, 0, 0x0202, __Sel_Gui.Hwnd) = "", "early release ignored")
        Check(!__SelDragging && __CapPickBusy, "early events leave startup intact")
    }
    CapPickCreateHud()
    if fixtureFailHud {
        TrackHandles()
        throw Error("Simulated failure after creating the picker HUD")
    }
}
CapPickScanControllers() {
    return [Map("type", "fixture", "id", 1, "name", "Fixture controller")]
}
CapPickReadControllerButtons(controller) {
    global fixtureButtons
    return fixtureButtons.Clone()
}
CapPickDetectController() {
    global fixtureFailTick
    if fixtureFailTick
        throw Error("Simulated controller polling failure")
}
CapPickReadAxes(*) => false
CapPickReadConfirmCancel(*) => false
CapPickResizeModifierHeld() => false
CapPickRegisterHotkeys() {
    ; Never install global input hooks in a test runner.
}
FixtureToolTip(*) {
}
CPBigBoxFinishCapture(message) {
    global CPBigBoxCaptureWatch, fixtureCaptureNotice
    CPBigBoxCaptureWatch["active"] := false
    fixtureCaptureNotice := message
}
FixtureMouseGetPos(&x := 0, &y := 0, &hwnd := 0) {
    x := -29900, y := -29900, hwnd := 0
}
CapPickEnumerateWindows() {
    return [Map("kind", "window", "title", "Fixture game", "label", "Fixture game", "exe", "fixture.exe", "class", "FixtureClass", "hwnd", 1)]
}
CapPickWindowCandidate(*) => false
CapPickRaiseWindowPreview(*) {
}
CapPickShowWindowCandidate(*) {
}
ShowWindowNoActivate(hwnd) {
    DllCall("ShowWindow", "ptr", hwnd, "int", 4)
}
FixtureCleanup(*) {
    global Overlay, otherOverlay, foreignGui
    try CapPickEndSession()
    for guiObject in [Overlay, otherOverlay, foreignGui]
        try guiObject.Destroy()
}
