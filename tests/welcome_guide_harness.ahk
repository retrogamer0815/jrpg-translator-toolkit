; Synthetic/offscreen welcome checks: no browser, real settings, or API calls.
TestWelcomeGuide() {
    global ui, tab, iniPath, CPWelcomeDialog, controlDarkMode
    global BEGINNER_VIDEO_URL, WRITTEN_GUIDE_URL, TestWelcomeUrls, TestWelcomeActivation
    TestWelcomeUrls := [], TestWelcomeActivation := ""
    ui.Show("NA x-12000 y-12000 w1120 h760")
    TestHelpDialogs(Map("gui", ui))
    ; Exercise the real Settings link callbacks without opening a browser or
    ; reading the user's credentials. This harness uses synthetic settings.
    CPDesktopNavigate(9)
    apiPage := CPDesktop["organizePages"][9]
    for spec in [["getGeminiKey", "https://aistudio.google.com/apikey"],
        ["getOpenAIKey", "https://platform.openai.com/api-keys"],
        ["geminiPricing", "https://ai.google.dev/gemini-api/docs/pricing"]] {
        urlsBefore := TestWelcomeUrls.Length
        TestWelcomeClick(apiPage[spec[1]])
        DesktopAssert(TestWelcomeUrls.Length = urlsBefore + 1 && TestWelcomeUrls[-1] = spec[2],
            "API setup link opens its official provider page: " spec[1])
    }

    for dark in [1, 0] {
        controlDarkMode := dark
        IniWrite(0, iniPath, "cfg_control", "welcomeGuideDismissed")
        CPSelectCustomTab(1)
        welcome := ShowWelcomeDialog(), g := welcome["gui"], c := welcome["controls"]
        hwnd := g.Hwnd
        for attempt in [1, 2] {
            TestWelcomeClick(c["apiKeys"])
            TestWelcomeAssertOpen(g)
            DesktopAssert(!welcome["closed"] && ObjPtr(CPWelcomeDialog) = ObjPtr(welcome),
                "Repeated API Keys clicks preserve the same welcome state")
            DesktopAssert(IniRead(iniPath, "cfg_control", "welcomeGuideDismissed", -1) = 0,
                "Opening API Keys does not persist the welcome dismissal preference")
        }
        reopened := ShowWelcomeDialog(true)
        DesktopAssert(reopened["gui"].Hwnd = hwnd,
            "Reopening the guide reuses the still-open welcome window")
        urlsBefore := TestWelcomeUrls.Length
        TestWelcomeClick(c["video"]), TestWelcomeClick(c["guide"])
        DesktopAssert(TestWelcomeUrls.Length = urlsBefore + 2
            && TestWelcomeUrls[urlsBefore + 1] = BEGINNER_VIDEO_URL
            && TestWelcomeUrls[urlsBefore + 2] = WRITTEN_GUIDE_URL,
            "Both guide buttons still open the intended URLs after API Keys")
        TestWelcomeClick(c["continue"])
        DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd) && !IsObject(CPWelcomeDialog)
            && IniRead(iniPath, "cfg_control", "welcomeGuideDismissed", -1) = 1,
            "Continue closes the guide and saves its checked preference")
        DesktopAssert(tab.Value = 9 && DllCall("user32\IsWindowEnabled", "ptr", ui.Hwnd),
            "API Keys stays selected and usable after Continue")
        CPSelectCustomTab(1)
        CPDesktopWelcomeOpenApiKeys(welcome)
        DesktopAssert(tab.Value = 1, "A late API Keys callback on a closed guide is ignored")

        welcome := ShowWelcomeDialog(true), g := welcome["gui"]
        welcome["controls"]["dontShow"].Value := 0
        TestWelcomeClick(welcome["controls"]["apiKeys"])
        DesktopAssert(IniRead(iniPath, "cfg_control", "welcomeGuideDismissed", -1) = 1,
            "Changing the checkbox then opening API Keys still waits for dismissal")
        hwnd := g.Hwnd
        PostMessage(0x10, 0, 0, hwnd) ; WM_CLOSE
        Sleep(50)
        DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd)
            && IniRead(iniPath, "cfg_control", "welcomeGuideDismissed", -1) = 0,
            "Closing the welcome window saves the unchecked preference")
    }

    ; The fallback presentation shares the same non-closing API Keys action.
    ui.CPDialogPresentation := "classic"
    try {
        CPSelectCustomTab(1)
        IniWrite(0, iniPath, "cfg_control", "welcomeGuideDismissed")
        ShowWelcomeDialog()
        g := CPWelcomeDialog, hwnd := g.Hwnd
        steps := ""
        for ctrl in g {
            if ctrl.Type = "Text" && InStr(ctrl.Text, "5. Save the finished setup")
                steps := ctrl.Text
        }
        DesktopAssert(InStr(steps, "LaunchBox plugin (optional)."),
            "Fallback welcome also identifies LaunchBox as optional")
        for attempt in [1, 2] {
            TestWelcomeClick(TestWelcomeControl(g, "Open API Keys"))
            TestWelcomeAssertOpen(g)
            DesktopAssert(CPWelcomeDialog.Hwnd = hwnd
                && IniRead(iniPath, "cfg_control", "welcomeGuideDismissed", -1) = 0,
                "Fallback API Keys keeps the guide and preference unchanged")
        }
        urlsBefore := TestWelcomeUrls.Length
        TestWelcomeClick(TestWelcomeControl(g, "Watch Beginner Guide"))
        TestWelcomeClick(TestWelcomeControl(g, "Open Written Guide"))
        DesktopAssert(TestWelcomeUrls.Length = urlsBefore + 2
            && TestWelcomeUrls[urlsBefore + 1] = BEGINNER_VIDEO_URL
            && TestWelcomeUrls[urlsBefore + 2] = WRITTEN_GUIDE_URL,
            "Fallback guide links remain available after API Keys")
        TestWelcomeClick(TestWelcomeControl(g, "Continue"))
        DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd) && !IsObject(CPWelcomeDialog)
            && IniRead(iniPath, "cfg_control", "welcomeGuideDismissed", -1) = 1,
            "Fallback Continue still saves and dismisses")
    } finally {
        ui.DeleteProp("CPDialogPresentation")
    }
    TestStudyWelcomeGuide()
}

TestStudyWelcomeGuide() {
    global ui, iniPath, CPStudyLibraryWelcomeDialog, controlDarkMode
    global STUDY_VIDEO_URL, STUDY_WRITTEN_GUIDE_URL, PROJECT_URL, TestWelcomeUrls
    global CPDesktopNativePaintHwnds
    STUDY_VIDEO_URL := "https://example.invalid/study-video"
    STUDY_WRITTEN_GUIDE_URL := PROJECT_URL "/blob/main/docs/manual/11-study-library.md"
    owner := Map("gui", ui)
    for dark in [1, 0] {
        controlDarkMode := dark
        IniWrite(0, iniPath, "study_library", "welcomeDismissed")
        ShowStudyLibraryWelcome(owner)
        welcome := CPStudyLibraryWelcomeDialog, g := welcome["gui"], hwnd := g.Hwnd
        expected := ["Watch Video Guide", "Open Written Guide", "Open AnkiConnect Page", "Continue"]
        buttons := []
        for ctrl in g
            if ctrl.Type = "Button"
                buttons.Push(ctrl)
        DesktopAssert(buttons.Length = 4, "Study welcome has four actions without a folder button")
        previousRight := 0
        for i, ctrl in buttons {
            ctrl.GetPos(&x, &y, &w, &h)
            DesktopAssert(ctrl.Text = expected[i] && x >= previousRight,
                "Study welcome actions retain their left-to-right navigation order")
            DesktopAssert(DesktopTextWidth(ctrl) < w - 15,
                "Study welcome button label fits: " ctrl.Text)
            previousRight := x + w
        }
        DesktopAssert(StudyControllerFocusedHwnd(hwnd) = buttons[4].Hwnd,
            "Study welcome initially focuses Continue")
        continueHwnd := buttons[4].Hwnd
        DesktopAssert(CPDesktopNativePaintHwnds.Has(continueHwnd)
            && CPDesktopNativePaintHwnds[continueHwnd]["paintState"]["paint"][continueHwnd]["kind"] = "primary",
            "Study Continue uses the shared blue primary-button renderer")
        for ctrl in [buttons[1], buttons[2], buttons[3], buttons[4]] {
            ctrl.Focus()
            CPApplyOwnedDialogTheme(g)
            DesktopAssert((DllCall("user32\GetWindowLongPtr", "ptr", continueHwnd, "int", -16, "ptr") & 0xF) = 0xB,
                "Continue retains its blue owner-drawn style after focus and theme changes")
        }
        DesktopAssert((SendMessage(0x87, 0, 0, continueHwnd) & 0x10) != 0,
            "Painted Continue retains default-button Enter/controller A semantics")
        urlsBefore := TestWelcomeUrls.Length
        TestWelcomeClick(buttons[1]), TestWelcomeClick(buttons[2]), TestWelcomeClick(buttons[3])
        DesktopAssert(TestWelcomeUrls.Length = urlsBefore + 3
            && TestWelcomeUrls[urlsBefore + 1] = STUDY_VIDEO_URL
            && TestWelcomeUrls[urlsBefore + 2] = STUDY_WRITTEN_GUIDE_URL
            && TestWelcomeUrls[urlsBefore + 3] = "https://ankiweb.net/shared/info/2055492159",
            "Reordered Study guide buttons open the video, written guide, and AnkiConnect page")
        DesktopAssert(DllCall("user32\IsWindowVisible", "ptr", hwnd)
            && IniRead(iniPath, "study_library", "welcomeDismissed", -1) = 0,
            "Opening Study guides keeps the welcome window and preference unchanged")
        g.GetClientPos(,, &pixelW, &pixelH)
        dpi := GetWindowDPI(hwnd) / 96
        TestDesktopStudyCapture(g, dark ? "help-study-welcome.png" : "help-study-welcome-light.png",
            Round(pixelW * dpi), Round(pixelH * dpi))
        welcome["dontShowAgain"].Value := 0
        TestWelcomeClick(buttons[4])
        DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd)
            && !IsObject(CPStudyLibraryWelcomeDialog)
            && IniRead(iniPath, "study_library", "welcomeDismissed", -1) = 0,
            "Continue closes Study welcome without suppressing an unchecked introduction")
        DesktopAssert(!CPDesktopNativePaintHwnds.Has(continueHwnd),
            "Closing Study welcome releases its button painting state")
        ShowStudyLibraryWelcome(owner)
        g := CPStudyLibraryWelcomeDialog["gui"], hwnd := g.Hwnd
        CPStudyLibraryWelcomeDialog["dontShowAgain"].Value := 1
        TestWelcomeClick(TestWelcomeControl(g, "Continue"))
        ShowStudyLibraryWelcome(owner)
        DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd)
            && !IsObject(CPStudyLibraryWelcomeDialog)
            && IniRead(iniPath, "study_library", "welcomeDismissed", -1) = 1,
            "Continue saves a checked Study preference and suppresses the next introduction")
    }
}

TestWelcomeAssertOpen(g) {
    global ui, tab, TestWelcomeActivation
    DesktopAssert(DllCall("user32\IsWindowVisible", "ptr", g.Hwnd)
        && DllCall("user32\IsWindowEnabled", "ptr", g.Hwnd),
        "Welcome remains visible and interactive after API Keys")
    DesktopAssert(tab.Value = 9 && DllCall("user32\IsWindowVisible", "ptr", ui.Hwnd)
        && DllCall("user32\IsWindowEnabled", "ptr", ui.Hwnd),
        "API Keys is open in an enabled main window behind the modeless guide")
    DesktopAssert(TestWelcomeActivation = "ahk_id " g.Hwnd,
        "API Keys returns foreground activation to the welcome guide")
}

TestWelcomeClick(ctrl) {
    SendMessage(0xF5, 0, 0, ctrl.Hwnd) ; BM_CLICK, delivered only to test controls
    Sleep(50)
}

TestWelcomeControl(g, text) {
    for ctrl in g
        if ctrl.Type = "Button" && ctrl.Text = text
            return ctrl
    throw Error("Missing welcome button: " text)
}

TestWelcomeOpenUrl(url, description) {
    global TestWelcomeUrls
    TestWelcomeUrls.Push(url)
}

TestWelcomeActivate(target) {
    global TestWelcomeActivation
    TestWelcomeActivation := target
}
