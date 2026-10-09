; End-to-end switching: real INI files, real unsaved-change detection, real
; modal dialog and real application. All files/windows are synthetic fixtures.
TestProfileSwitch() {
    global ui, tab, CPDesktop, iniPath, ddlGameProfile, btnGameProfileApply, boxBgHex
    global ddlPrompt, ddlEPr, TestSwitchState
    global capMode := "region", capRect := "10,10,100,100", capWinInfo := ""
    global captureDir := A_ScriptDir "\captures", overlayDir := A_ScriptDir "\overlay"
    global appDir := A_ScriptDir "\Settings"
    global studyLibraryDefaultDir := A_ScriptDir "\study-fixture"
    global studyLibrariesRoot := A_ScriptDir "\study-libraries", studyLibraryDir := studyLibraryDefaultDir
    global ewX := "", ewY := "", ewW := "", ewH := ""
    global ew_lastX := "", ew_lastY := "", ew_lastW := "", ew_lastH := "", ew_bounds_watch_running := false
    DirCreate(captureDir), DirCreate(overlayDir)
    IniWrite(capMode, iniPath, "capture", "mode")
    IniWrite(capRect, iniPath, "capture", "rect")
    ui.Show("NA x-9000 y-9000 w1120 h760")
    CPDesktopNavigate(7)
    OnExit(ToastDestroy)
    ; A watchdog prevents an unhandled modal from leaving the harness hanging.
    SetTimer(TestSwitchTimeout, -45000)
    for entry in ["header", "page"] {
        for route in ["keyboard", "controller", "mouse"] {
            for answer in ["Yes", "No", "Cancel"] {
                boxBgHex := "FF0000"
                ddlPrompt.Choose("literal"), ddlEPr.Choose("detailed_grammar")
                DesktopAssert(GameProfileSave("Alternate profile", false), "Save target fixture")
                boxBgHex := "202020"
                ddlPrompt.Choose("default_with_kanji_reading_en"), ddlEPr.Choose("default_en")
                DesktopAssert(GameProfileSave("Demo profile", false), "Save active fixture")
                DesktopAssert(!GameProfileHasUnsavedChanges("Demo profile"), "Clean saved baseline")
                activePath := GameProfilePath("Demo profile"), targetPath := GameProfilePath("Alternate profile")
                activeBefore := FileRead(activePath), targetBefore := FileRead(targetPath)
                boxBgHex := "112233"
                DesktopAssert(GameProfileHasUnsavedChanges("Demo profile"), "Live edit is detected")
                RefreshGameProfilesList("Alternate profile")
                CPDesktopRefreshProfileSelector(true)
                TestSwitchState := Map("answer", answer, "route", route, "dialogs", 0, "responded", false)
                SetTimer(TestSwitchRespond, 20)
                try {
                    if entry = "header" {
                        ctrl := CPDesktop["chrome"]["profile"]
                        ctrl.Focus()
                        CPControllerDispatchNavigation("Activate", ui.Hwnd)
                        Sleep(25)
                        ctrl.Choose("Alternate profile")
                        CPControllerDispatchNavigation("Activate", ui.Hwnd)
                    } else
                        SendMessage(0xF5, 0, 0, btnGameProfileApply.Hwnd)
                    Sleep(250)
                } finally {
                    SetTimer(TestSwitchRespond, 0)
                }
                DesktopAssert(TestSwitchState["dialogs"] = 1 && TestSwitchState["responded"],
                    "One real dirty confirmation: " entry " / " route " / " answer)
                switched := answer != "Cancel"
                expected := switched ? "Alternate profile" : "Demo profile"
                DesktopAssert(IniRead(iniPath, "game_profiles", "active") = expected, "Actual active profile after " answer)
                DesktopAssert(CPDesktop["chrome"]["profile"].Text = expected, "Header reflects actual active profile")
                DesktopAssert(boxBgHex = (switched ? "FF0000" : "112233"), "Applied target appearance or preserved unsaved edit")
                DesktopAssert(ddlPrompt.Text = (switched ? "literal" : "default_with_kanji_reading_en"), "Translation prompt switched correctly")
                DesktopAssert(ddlEPr.Text = (switched ? "detailed_grammar" : "default_en"), "Explanation prompt switched correctly")
                DesktopAssert(FileRead(targetPath) = targetBefore, "Target profile is never overwritten by save-and-switch")
                if answer = "Yes" {
                    DesktopAssert(IniRead(activePath, "translator", "boxBg") = "112233", "Live edit saved to old active profile")
                    DesktopAssert(FileRead(activePath ".bak") = activeBefore, "Previous saved profile remains recoverable")
                } else
                    DesktopAssert(FileRead(activePath) = activeBefore, "No/Cancel do not overwrite saved profile")
                DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", ui.Hwnd), "Main window is enabled after modal")
                DesktopAssert(!IsObject(CPDesktop.Get("profileSelection", 0)), "No pending dropdown callback remains")
                DesktopAssert(!TestSwitchHasMessage(), "Confirmation window is destroyed")
                CPDesktopNavigate(1), CPDesktopNavigate(7)
                DesktopAssert(tab.Value = 7, "Main navigation still works after switch")
                ToastDestroy()
            }
        }
    }
    ; A clean profile and Current settings must not need a dirty prompt.
    for active in ["Demo profile", ""] {
        DesktopAssert(GameProfileApply("Demo profile", false), "Restore clean saved profile")
        IniWrite(active, iniPath, "game_profiles", "active")
        RefreshGameProfilesList("Alternate profile")
        CPDesktopRefreshProfileSelector(true)
        TestSwitchState := Map("answer", "Cancel", "route", "keyboard", "dialogs", 0, "responded", false)
        SetTimer(TestSwitchRespond, 20)
        try ApplySelectedGameProfile()
        finally SetTimer(TestSwitchRespond, 0)
        DesktopAssert(TestSwitchState["dialogs"] = 0, "Clean/current-settings switch skips confirmation")
        DesktopAssert(IniRead(iniPath, "game_profiles", "active") = "Alternate profile", "Clean/current-settings switch applies")
    }
    SetTimer(TestSwitchTimeout, 0)
}

TestSwitchHasMessage() {
    for hwnd, d in StudyDesktopRegistry()
        if d["kind"] = "message"
            return true
    return false
}

TestSwitchRespond(*) {
    global ui, TestSwitchState, TestMessageState
    for hwnd, d in StudyDesktopRegistry() {
        if d["kind"] != "message" || d["width"] <= 0
            continue
        s := d["state"]
        if s["closed"] || TestSwitchState["responded"]
            return
        DesktopAssert(s["gui"].Title = "Switch profile", "Expected dirty-profile confirmation, not an error")
        DesktopAssert(!DllCall("user32\IsWindowEnabled", "ptr", ui.Hwnd), "Owner disabled during actual modal")
        ; The real modal show helper initially focuses Cancel; wait for that.
        if StudyControllerFocusedHwnd(hwnd) != s["cancel"].Hwnd
            return
        TestSwitchState["dialogs"] += 1
        TestSwitchState["responded"] := true
        TestMessageState := s
        answer := TestSwitchState["answer"], route := TestSwitchState["route"]
        steps := answer = "Yes" ? 2 : answer = "No" ? 1 : 0
        Loop steps
            TestMessageNativeKey("Left")
        if route = "mouse"
            SendMessage(0xF5, 0, 0, s[answer = "Yes" ? "addButton" : answer = "No" ? "noButton" : "cancel"].Hwnd)
        else if route = "controller"
            TestMessageControllerDispatch("Activate", hwnd)
        else
            TestMessageNativeKey("Enter")
        return
    }
}

TestSwitchTimeout(*) {
    FileAppend("FAIL: Profile switch/modal did not complete within the watchdog limit.`n", "*")
    ExitApp(1)
}
