; Included by the desktop harness. Real themed footer, Toast and audio lifecycle;
; the child is an isolated AHK fixture, never Python, a microphone or an AI API.
TestAudioFeedbackBegin() {
    global ui, pythonExe, audioScript, CPDesktop, TestFeedbackState, CPOverlayAdjustState
    CPOverlayAdjustState := Map("active", false)
    pythonExe := A_AhkPath
    audioScript := A_ScriptDir "\audio_runtime_fixture.ahk"
    EnvSet("JRPG_TEST_AUDIO_RUNTIME_LOG", A_ScriptDir "\feedback-child.log")
    EnvSet("JRPG_TEST_AUDIO_RUNTIME_MODE", "wait")
    TestFeedbackState := Map("step", 0, "footerPaints", 0, "toastPaints", 0)
    OnMessage(0x2B, TestAudioFeedbackDraw, -1)
    OnMessage(0xF, TestAudioFeedbackPaint, -1)
    OnExit(AudioShutdown)
    OnExit(ToastDestroy)
    ui.Show("NA x0 y0 w1120 h760")
    CPDesktopNavigate(8)
    SetTimer(_UpdateStatus, 1000)
    SetTimer(TestAudioFeedbackStep, -500)
    FileAppend("READY`n", A_ScriptDir "\feedback-events.txt")
    ; Return to AHK's idle message loop. Do not drive this with Sleep, forced
    ; redraws, PrintWindow or synthetic mouse movement, which can hide stalls.
    return true
}

TestAudioFeedbackDraw(wParam, lParam, *) {
    global CPDesktop, TestFeedbackState
    if NumGet(lParam, A_PtrSize = 8 ? 24 : 20, "ptr") = CPDesktop["chrome"]["audio"].Hwnd {
        TestFeedbackState["footerPaints"] += 1
        TestFeedbackState["paintedFooter"] := CPDesktop["chrome"]["audio"].Text
    }
}

TestAudioFeedbackPaint(wParam, lParam, msg, hwnd) {
    global CPToastText, TestFeedbackState
    if IsObject(CPToastText) && hwnd = CPToastText.Hwnd
        TestFeedbackState["toastPaints"] += 1
}

TestAudioFeedbackStep(*) {
    global TestFeedbackState, TestAssertions, CPToastGui, CPToastText, CPDesktop, ui, controlDarkMode
    try {
        step := TestFeedbackState["step"]
        FileAppend("STEP " step " critical=" A_IsCritical "`n", A_ScriptDir "\feedback-events.txt")
        if Mod(step, 3) = 0 {
            if step = 6 || step = 12 {
                controlDarkMode := step = 12
                CPDesktopTheme()
                if step = 12
                    WinSetTransparent(220, ui.Hwnd)
            }
            if step = 18
                ui.Hide()
            TestFeedbackState["footerPaints"] := 0
            TestFeedbackState["toastPaints"] := 0
            FileAppend("before toggle`n", A_ScriptDir "\feedback-events.txt")
            if Mod(step // 3, 2) = 0
                SafeCall(StartStopAudio) ; Production keyboard/JoyToKey wrapper.
            else
                CPControllerDispatchAction("start_stop_audio")
            DesktopAssert(AudioIsRunning() = (Mod(step // 3, 2) = 0), "Shortcut changes the fixture process state")
            FileAppend("after toggle critical=" A_IsCritical "`n", A_ScriptDir "\feedback-events.txt")
            DesktopAssert(A_IsCritical = 0, "Audio lifecycle restores interruptibility")
            TestFeedbackState["toastHwnd"] := CPToastGui.Hwnd
            DesktopAssert((WinGetExStyle(CPToastGui.Hwnd) & 0x08080028) = 0x08080028,
                "Toast is layered, non-activating, topmost and click-through after showing")
            DesktopAssert(DllCall("user32\GetForegroundWindow", "ptr") != CPToastGui.Hwnd, "Toast never takes foreground focus")
            TestFeedbackState["step"] += 1
            SetTimer(TestAudioFeedbackStep, -350)
        } else if Mod(step, 3) = 1 {
            expected := Mod(step // 3, 2) = 0 ? "On" : "Off"
            DesktopAssert(IsObject(CPToastText) && CPToastText.Text = "Audio Translation " expected, "Toast text is assigned")
            DesktopAssert(TestFeedbackState["toastPaints"] > 0, "Toast text actually paints without mouse movement")
            TestAudioFeedbackPixels(CPToastText, controlDarkMode)
            if step < 18 {
                DesktopAssert(TestFeedbackState["footerPaints"] > 0 && TestFeedbackState.Get("paintedFooter", "") = "Audio: " expected,
                    "Footer actually paints the new audio state without mouse movement")
                TestAudioFeedbackPixels(CPDesktop["chrome"]["audio"], controlDarkMode)
            }
            TestFeedbackState["step"] += 1
            SetTimer(TestAudioFeedbackStep, -1800)
        } else {
            DesktopAssert(!DllCall("user32\IsWindow", "ptr", TestFeedbackState["toastHwnd"]), "Toast dismisses while idle")
            if step < 23 {
                TestFeedbackState["step"] += 1
                SetTimer(TestAudioFeedbackStep, -100)
            } else {
                ; A newer notification must replace a still-visible one and
                ; receive its own dismissal, including other users of Toast.
                Toast("Generating explanation…")
                previousToast := CPToastGui.Hwnd
                TestFeedbackState["toastPaints"] := 0
                Toast("Audio Translation Off")
                DesktopAssert(!DllCall("user32\IsWindow", "ptr", previousToast), "New toast removes the preceding notification")
                TestFeedbackState["toastHwnd"] := CPToastGui.Hwnd
                SetTimer(TestAudioFeedbackRapidCheck, -350)
            }
        }
    } catch as ex {
        FileAppend("FAIL: " ex.Message "`n" ex.Stack "`n", "*")
        ExitApp(1)
    }
}

TestAudioFeedbackRapidCheck(*) {
    global TestFeedbackState, CPToastText, controlDarkMode
    DesktopAssert(TestFeedbackState["toastPaints"] > 0 && CPToastText.Text = "Audio Translation Off", "Replacement toast paints its own text")
    TestAudioFeedbackPixels(CPToastText, controlDarkMode)
    SetTimer(TestAudioFeedbackFinish, -1800)
}

TestAudioFeedbackFinish(*) {
    global TestFeedbackState, TestAssertions, ui
    DesktopAssert(!DllCall("user32\IsWindow", "ptr", TestFeedbackState["toastHwnd"]), "Replacement toast dismisses while idle")
    FileAppend("PASS: " TestAssertions " real desktop audio-feedback assertions.`n", "*")
    ui.Destroy()
    ExitApp(0)
}

TestAudioFeedbackPixels(ctrl, dark) {
    ; Read the existing client pixels, never request WM_PRINT/PrintWindow or a
    ; repaint: those would conceal exactly the regression under test.
    rect := Buffer(16, 0)
    DllCall("user32\GetClientRect", "ptr", ctrl.Hwnd, "ptr", rect)
    width := NumGet(rect, 8, "int"), height := NumGet(rect, 12, "int")
    dc := DllCall("user32\GetDC", "ptr", ctrl.Hwnd, "ptr")
    found := 0
    try {
        Loop Max(0, height - 12) {
            y := A_Index + 5
            Loop Max(0, width - 20) {
                rgb := DllCall("gdi32\GetPixel", "ptr", dc, "int", A_Index + 9, "int", y, "uint")
                r := rgb & 255, g := (rgb >> 8) & 255, b := (rgb >> 16) & 255
                if rgb != 0xFFFFFFFF && (dark ? Min(r, g, b) > 180 : Max(r, g, b) < 100)
                    found += 1
                if found > 12
                    break
            }
            if found > 12
                break
        }
    } finally DllCall("user32\ReleaseDC", "ptr", ctrl.Hwnd, "ptr", dc)
    DesktopAssert(found > 12, "Text pixels are present in " ctrl.Text)
}
