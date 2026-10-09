#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn All, StdOut
global checks := 0, CPToastGui := 0, CPToastText := 0, CPToastTimer := 0
global controlDarkMode := true, Overlay := 0, __LastActiveHwnd := 123, simulateNoForeground := false
try {
    beforeForeground := DllCall("user32\GetForegroundWindow", "ptr")
    for theme in [true, false] {
        controlDarkMode := theme
        Toast("Explanation timed out. Try again or choose another model.`nDetails are in the Explainer overlay.", 15000, 2)
        first := CPToastGui, timer := CPToastTimer
        Check(IsObject(first) && DllCall("user32\IsWindow", "ptr", first.Hwnd), "Native failure toast exists")
        Check(IsWindowTopmost(first.Hwnd), "Toast is above normal windows")
        Check((DllCall("user32\GetWindowLongW", "ptr", first.Hwnd, "int", -20, "uint") & 0x08080020) = 0x08080020,
            "Toast remains non-activating, layered, and click-through")
        Check(ToastState()["expires"] - A_TickCount > 14000, "Failure gets fifteen seconds")
        Check(InStr(CPToastText.Text, "timed out"), "Toast has actionable failure text")
        Toast("Saved")
        Check(CPToastGui = first && CPToastTimer = timer, "Routine notification cannot mask the failure")
        Toast("Generating explanation…", 1600, 2)
        replacement := CPToastGui, replacementTimer := CPToastTimer
        Check(replacement != first && CPToastText.Text = "Generating explanation…", "Retry replaces stale failure")
        ToastDestroy(first)
        Check(CPToastGui = replacement && CPToastTimer = replacementTimer, "Queued old expiry cannot destroy new toast or timer")
        ToastDestroy()
        Check(CPToastGui = 0 && CPToastTimer = 0 && ToastState()["priority"] = 0, "Cleanup resets notification state")
        Toast("Synthetic expiry", 40)
        Sleep(100)
        Check(CPToastGui = 0, "Native expiry timer dismisses toast")
    }
    for simulated in [false, true] {
        simulateNoForeground := simulated
        pos := ProductionToastPosition(500, 70)
        Check(pos.Length = 2 && IsNumber(pos[1]) && IsNumber(pos[2]), "Monitor placement handles absent foreground")
        Overlay := Gui("-Caption +ToolWindow", "Synthetic explanation feedback overlay")
        Overlay.Show("Hide x-20000 y-20000 w200 h100")
        oldHidden := A_DetectHiddenWindows
        DetectHiddenWindows(false)
        Loop 50 {
            ToggleTop()
            Check(IsWindowTopmost(Overlay.Hwnd), "Hidden overlay becomes topmost without visible-window lookup")
            if simulated
                Check(__LastActiveHwnd = 0, "No active window is accepted without an AHK error")
            ToggleTop()
            Check(!IsWindowTopmost(Overlay.Hwnd), "Hidden overlay safely leaves topmost band")
            Check(A_DetectHiddenWindows = false, "Toggle restores hidden-window mode")
            Check(!DllCall("user32\IsWindowVisible", "ptr", Overlay.Hwnd), "Topmost toggle never unhides the overlay")
        }
        stale := Overlay.Hwnd
        Overlay.Destroy()
        Check(!IsWindowTopmost(stale), "Destroyed HWND is safe")
        ToggleTop()
        Check(A_DetectHiddenWindows = false, "Destroyed-overlay fallback restores window mode")
        DetectHiddenWindows(oldHidden)
    }
    Check(DllCall("user32\GetForegroundWindow", "ptr") = beforeForeground, "Feedback and hotkeys do not steal foreground")
    FileAppend("PASS: " checks " native explanation-feedback assertions (" (A_PtrSize * 8) "-bit).`n", "*")
    ExitApp(0)
} catch as testFeedbackError {
    ToastDestroy()
    try Overlay.Destroy()
    FileAppend("FAIL: " testFeedbackError.Message " at " testFeedbackError.Line "`n" testFeedbackError.Stack "`n", "*")
    ExitApp(1)
}
Check(condition, description) {
    global checks
    checks += 1
    if !condition
        throw Error(description)
}
TestForegroundWindow() => simulateNoForeground ? 0 : DllCall("user32\GetForegroundWindow", "ptr")
ToastPosition(*) => [-20000, -20000] ; Only test windows, never cover the user's desktop.
DbgCP(*) => 0
