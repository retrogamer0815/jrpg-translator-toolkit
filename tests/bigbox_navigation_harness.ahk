#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn All, StdOut
#NoTrayIcon
; @BIGBOX_GLOBALS@
; @DESKTOP_TABS@
; @AUDIO_LANGUAGES@
global CP_GAME_TITLE := "The Legend of Xanadu: Kaze no Densetsu Xanadu — an exceptionally long test title"
global CP_GAME_PLATFORM := "NEC PC Engine-CD"
global CP_GAME_BOX_ART := A_Args.Length ? A_Args[1] : ""
global CP_GAME_CLEAR_LOGO := A_Args.Length > 1 ? A_Args[2] : ""
global CP_PLATFORM_CLEAR_LOGO := "", CP_PLATFORM_DEVICE_IMAGE := "", CP_PLATFORM_DEFAULT_ART := ""
global controlDarkMode := 1, iniPath := A_ScriptDir "\test-settings.ini"
global envPath := A_ScriptDir "\Settings\.env"
global APP_VERSION := "v-test", PROJECT_URL := "https://example.invalid/project"
global BUG_REPORT_URL := "https://example.invalid/bugs", BEGINNER_VIDEO_URL := "https://example.invalid/video"
global WRITTEN_GUIDE_URL := "https://example.invalid/guide"
global showPathsTab := false
global imgProvider := "gemini", geminiImgModel := "gemini-test", imgModel := "gpt-test"
global explainProvider := "openai", explainOpenAIModel := "gpt-test", explainGeminiModel := "gemini-test"
global audioProvider := "gemini", trModel := "audio-test", geminiAudioModel := "gemini-audio-test"
global CPControllerInputsEnabled := false, CPControllerDpadNavigationEnabled := true
global CPControllerCaptureActive := false, CPControllerLastStatusText := "XInput controller 1"
global CPControllerNavPreviousState := Map(), CPControllerNavTargetHwnd := 0
global CPControllerNavHeldDirection := "", CPControllerNavNextRepeatAt := 0, CPControllerNavHeldSince := 0
global CPControllerSurfaceTransitionState := ""
global CPControllerLastNativeNavigationAt := Map()
global CPPreviousForegroundHwnd := 0, ui := Gui()
global CPStudyLibraryState := 0, TestStudyLibraryOpens := 0
global TestStudyLibraryBigBox := false
global TestStudyLibraryOpenedOverDashboard := false
; Synthetic desktop controls and configuration. The production UpdateVars,
; SaveAll and ApplyShotSettings routines operate on these, not personal data.
global model_openai_img := ["gpt-test", "gpt-alternate"]
global model_gemini_img := ["gemini-test", "gemini-alternate", "model-three", "model-four",
    "model-five", "model-six", "model-seven", "model-eight", "a-very-long-model-name-for-testing-layout-and-pagination"]
global model_openai_explain := ["gpt-test", "gpt-explain-alternate"]
global model_gemini_explain := ["gemini-test", "gemini-explain-alternate"]
global model_openai_audio := ["audio-test", "audio-alternate"]
global model_gemini_audio := ["gemini-audio-test", "gemini-audio-alternate"]
global ddlProv := TestCombo(["Gemini", "OpenAI"]), ddlIMG_GM := TestCombo(model_gemini_img)
global ddlIMG := TestCombo(model_openai_img), ddlPrompt := TestCombo(["default", "literal"])
global ddlEProv := TestCombo(["Gemini", "OpenAI"], 2), ddlEGem := TestCombo(model_gemini_explain)
global ddlEOpenAI := TestCombo(model_openai_explain), ddlEPr := TestCombo(["default", "detailed"])
global ddlAProv := TestCombo(["Gemini", "OpenAI"]), ddlTR := TestCombo(model_openai_audio)
global ddlA_GM := TestCombo(model_gemini_audio), ddlAudioTarget := TestCombo(["English", "German", "French"])
global btnIMG_Add := ui.AddButton(), btnIMG_Del := ui.AddButton()
global btnIMG_GM_Add := ui.AddButton(), btnIMG_GM_Del := ui.AddButton()
global btnTR_Add := ui.AddButton(), btnTR_Del := ui.AddButton()
global btnA_GM_Add := ui.AddButton(), btnA_GM_Del := ui.AddButton()
global slTrans := ui.AddSlider("Range0-255", 255), ddlFont := TestCombo(["Segoe UI"])
global edFSize := ui.AddEdit(, "14"), chkFontBold := ui.AddCheckbox()
global udFSize := ui.AddUpDown("Range6-128", 14), lblTransPct := ui.AddText()
global rectBg := ui.AddText(), rectTxt := ui.AddText(), rectName := ui.AddText()
global slTrans_EW := ui.AddSlider("Range0-255", 255), lblTransPct_EW := ui.AddText()
global rectBg_EW := ui.AddText(), rectTxt_EW := ui.AddText()
global ddlFont_EW := TestCombo(["Segoe UI", "Arial", "Consolas", "Georgia", "Verdana", "Tahoma"])
global edFSize_EW := ui.AddEdit(, "14"), udFSize_EW := ui.AddUpDown("Range6-200", 14), chkFontBold_EW := ui.AddCheckbox()
global CPFontSizeAdjustSyncing := false
global TestOverlayThemes := [], TestOverlayThemeOK := true, TestOverlayGradientKeys := Map()
global TestOverlayWindows := Map()
global CPControllerColorGradientSliders := TestOverlayGradientKeys, CPControllerColorGradientMessageRegistered := false
global TestOverlayGui := Gui("-Caption +ToolWindow"), CPOverlayAdjustState := Map("active", false)
global CP_PRESENTATION_MODE := "bigbox", TestOverlayReady := true, TestOverlayReturns := 0, TestOverlaySaves := 0
global cbDirectModelOutput := ui.AddCheckbox(), cbDebug := ui.AddCheckbox()
global cbApiInApp := ui.AddCheckbox(), eOpenAI := ui.AddEdit(, ""), eGemini := ui.AddEdit(, "")
global btnSaveEnv := ui.AddButton(), btnDelEnv := ui.AddButton()
global envSavedOpenAI := "", envSavedGemini := ""
global chkDel := ui.AddCheckbox(), chkTop := ui.AddCheckbox(), chkDarkMode := ui.AddCheckbox()
global chkGuess := ui.AddCheckbox(), chkName := ui.AddCheckbox()
global saveLibraryChk := ui.AddCheckbox("Checked"), saveLibraryScreenshotsChk := ui.AddCheckbox("Checked")
global saveExplChk := ui.AddCheckbox(), chkOpenEW := ui.AddCheckbox(), chkTop_EW := ui.AddCheckbox()
global chkOpenTW := ui.AddCheckbox(), chkTop_TW := ui.AddCheckbox()
global chkUseTerminologyOverrides := ui.AddCheckbox()
global jp2enGlossaryProfile := "default", en2enGlossaryProfile := "default"
global ddlJPG := TestCombo(["default"]), ddlENG := TestCombo(["default"])
global gameProfilesDir := A_ScriptDir "\profiles", glossariesDir := A_ScriptDir "\glossaries"
global promptsDir := A_ScriptDir "\prompts", explainPromptsDir := A_ScriptDir "\prompts_explain"
global ddlGameProfile := TestCombo(["Demo"]), txtGameProfileState := ui.AddText()
global ddlStartupOverlays := TestCombo(CPStartupOverlayOptions())
global GameProfileLastError := "", TestProfileSaves := 0, TestProfileApplies := 0
global ddlSpeaker := TestCombo(["[Windows Default]", "Game speakers", "Headphones", "USB DAC", "TV audio", "日本語 ＆ 音声"])
global btnSpRef := ui.AddButton(), btnAudioTest := ui.AddButton(), txtAudioTestStatus := ui.AddText()
global eCapMax := ui.AddEdit(, "1400"), CPMaxPngAdjustSyncing := false
global TestAudioStarts := 0, TestAudioStops := 0, TestCaptureSends := [], TestCaptureResumes := 0
global TestCaptureSendOK := true, TestTranslatorReady := true, TestAudioStartOK := true
global TestAudioApiConfigured := true, TestAudioRecoveryScans := 0
global TestAudioBusyProbe := Map("active", false, "samples", [])
global TestToasts := []
global gPidAudio := 0, gJustStoppedUntil := 0, gLastAction := ""
global gAudioSessionFile := A_ScriptDir "\audio-session.pid"
global speakerName := "[Windows Default]", gAudioInputJob := Map("active", false), gAudioInputStatus := "Not tested."
global TestLiveAudioRunning := false
global pythonExe := "python.exe", audioScript := "audio.py", overlayAhk := "overlay.ahk", imgScript := "image.py"
global explainScript := "explain.py", captureDir := "captures", overlayTrans := 255
global ePython := ui.AddEdit(, pythonExe), eOverlay := ui.AddEdit(, overlayAhk)
global eImg := ui.AddEdit(, imgScript), eAudio := ui.AddEdit(, audioScript), eExplain := ui.AddEdit(, explainScript)
global btnSavePaths := ui.AddButton(), pathsDirty := false
global audioTargetLang := "English", promptProfile := "default", explainPromptProfile := "default", imgPostproc := "test"
global capMaxKB := 1400, capMode := "region", capRect := "0,0,100,100"
global debugMode := 0, directModelOutput := 0, controlPanelOpacity := 1, useTerminologyOverrides := 0
global boxBgHex := "222222", bdrOutHex := "222222", bdrInHex := "222222", txtHex := "FFFFFF", nameHex := "55AAFF"
global fontName := "Segoe UI", fontSize := 14, fontBold := 0, bdrOutW := 0, bdrInW := 0
global fontName_EW := "Segoe UI", fontSize_EW := 14, fontBold_EW := 0, overlayTrans_EW := 255
global boxBgHex_EW := "222222", txtHex_EW := "FFFFFF", ewX := "", ewY := "", ewW := "", ewH := ""
global TestCount := 0, TestExternalCalls := 0
global TestModelWrites := [], TestModelWriteFails := false
global TestCatalogRequests := [], TestCatalogDeferred := false, TestCatalogBusy := false
global TestCatalogCallback := 0, TestCatalogCancels := 0
global TestCatalogResult := Map("ok", true, "source", "online", "models", [], "warnings", [], "error", "")
global TestPaintCounts := Map()
global TestPickerFocusProbe := Map("armed", false, "count", 0, "error", "")
global hotkeyActions := [
    "screenshot_translate", "explain_last_translation", "hide_show_translator",
    "hide_show_explainer", "hide_show_control_panel", "take_screenshot",
    "screenshot_translation", "launch_explainer_request", "recapture_region", "start_stop_audio"
]
global hotkeyLabels := Map(
    "screenshot_translate", "Capture + Translate",
    "explain_last_translation", "Explain last translation",
    "hide_show_translator", "Show/Hide Translator",
    "hide_show_explainer", "Show/Hide Explainer",
    "hide_show_control_panel", "Show/Hide Control Panel",
    "take_screenshot", "Make Capture",
    "screenshot_translation", "Translate Captures",
    "launch_explainer_request", "Launch Explainer + Req.",
    "recapture_region", "Recapture Region",
    "start_stop_audio", "Audio Translation On/Off"
)
global hotkeyDefaults := Map(
    "screenshot_translate", "^+t", "explain_last_translation", "^+e",
    "hide_show_translator", "^+h", "hide_show_explainer", "^+x",
    "hide_show_control_panel", "^+c", "take_screenshot", "^+s",
    "screenshot_translation", "^+d", "launch_explainer_request", "^+a",
    "recapture_region", "^+r", "start_stop_audio", "^+l"
)
global hkEdits := Map(), CPControllerBindings := Map(), CPControllerBindingEdits := Map()
global cbControllerInputsEnabled := ui.AddCheckbox()
global cbControllerDpadNavigationEnabled := ui.AddCheckbox("Checked")
global overlayDir := A_ScriptDir "\overlay-temp"
global TestControllerSnapshots := [], TestRebinds := 0
for testAction in hotkeyActions {
    hkEdits[testAction] := ui.AddEdit(, hotkeyDefaults[testAction])
    CPControllerBindings[testAction] := ""
    CPControllerBindingEdits[testAction] := ui.AddEdit(, "Disabled")
    IniWrite(hotkeyDefaults[testAction], iniPath, "hotkeys", testAction)
    IniWrite("", iniPath, "controller_inputs", testAction)
}
IniWrite(0, iniPath, "controller_inputs", "enabled")
IniWrite(1, iniPath, "controller_inputs", "dpad_navigation")
IniWrite("Demo", iniPath, "game_profiles", "active")
DirCreate(gameProfilesDir)
DirCreate(glossariesDir "\default")
DirCreate(promptsDir)
DirCreate(explainPromptsDir)
FileAppend("Translate {jp}", promptsDir "\default.txt", "UTF-8")
FileAppend("Literal {jp}", promptsDir "\literal.txt", "UTF-8")
FileAppend("Explain {jp}", explainPromptsDir "\default.txt", "UTF-8")
FileAppend("Detailed {jp}", explainPromptsDir "\detailed.txt", "UTF-8")
FileAppend("[profile]`r`nschemaVersion=1`r`nname=Demo`r`n", gameProfilesDir "\Demo.ini", "UTF-8")
FileAppend("# JP`r`n王 -> king`r`n", glossariesDir "\default\jp2en.txt", "UTF-8")
FileAppend("# TL`r`nking -> sovereign`r`n", glossariesDir "\default\en2en.txt", "UTF-8")

ToggleApiKeyControls(*) {
}

OpenWindowsEnvironmentVariables(*) {
    global TestExternalCalls
    TestExternalCalls += 1
}

OpenStudyLibraryWindow(standalone := false, bigBoxPresentation := false, *) {
    global CPStudyLibraryState, TestStudyLibraryOpens, TestStudyLibraryBigBox
    global TestStudyLibraryOpenedOverDashboard
    TestStudyLibraryOpens += 1
    TestStudyLibraryBigBox := bigBoxPresentation
    TestStudyLibraryOpenedOverDashboard := CPBigBoxDashboardVisible()
    CPStudyLibraryState := Map(
        "alive", true,
        "standaloneWindow", standalone,
        "bigBoxPresentation", bigBoxPresentation
    )
}

StudyLibraryStateAlive(studyState) {
    return IsObject(studyState) && studyState.Get("alive", false)
}

StudyControllerDispatchNavigation(*) => false
StudyControllerSurfaceIsRoot(*) => false

try {
    CPBigBoxDashboardCreate()
    TestAssert(CPBigBoxHomeTiles().Length = 8, "Eight Home tiles")
    TestAssert(CPBigBoxPageOrder().Length = 10, "Home plus nine normally visible desktop tabs")
    TestAssert(CPBigBoxNavigationControls.Length = 12, "Home has two arrows, eight tiles and two bottom actions")
    TestAssert((DllCall("user32\GetWindowLongW", "ptr", CPBigBoxControls["gameTitle"].Hwnd,
        "int", -16, "uint") & 0x4000) != 0, "Game title retains end ellipsis")
    TestHomeHeaderContent()

    for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
        CPBigBoxGui.Show("Hide w" size[1] " h" size[2])
        CPBigBoxDashboardResize(CPBigBoxGui, 0, size[1], size[2])
        CPBigBoxDashboardApplyTheme()
        TestLayout(size[1], size[2])
        TestCapture("home-" size[1] ".png", size[1], size[2])
        TestHomeHeaderLongNames(size[1], size[2])
        CPBigBoxSetPage("screenshot", false)
        TestLayout(size[1], size[2])
        TestScreenshotHintDetails(size[1])
        TestCapture("screenshot-settings-" size[1] ".png", size[1], size[2])
        CPBigBoxSetPage("quickTranslation", false)
        TestLayout(size[1], size[2])
        TestCapture("translation-quick-" size[1] ".png", size[1], size[2])
        CPBigBoxSetPage("explanation", false)
        TestLayout(size[1], size[2])
        TestCapture("explanation-settings-" size[1] ".png", size[1], size[2])
        CPBigBoxSetPage("audio", false)
        TestLayout(size[1], size[2])
        TestAudioHelpAndStatus(size[1], size[2])
        TestCapture("audio-input-" size[1] ".png", size[1], size[2])
        CPBigBoxSetPage("home", false)
    }
    ; Full pages must match every original desktop tab, not the Home tile list.
    showPathsTab := true
    TestAssert(CPBigBoxPageOrder().Length = TestDesktopTabNames.Length + 1, "All desktop tabs represented")
    for testPageIndex, testPageKey in CPBigBoxPageOrder() {
        if (testPageIndex = 1)
            continue
        TestAssert(CPBigBoxPageDefinitions()[testPageKey][1] = TestDesktopTabNames[testPageIndex - 1],
            "Main page matches desktop tab " TestDesktopTabNames[testPageIndex - 1])
        CPBigBoxSwitchPage(1)
        TestAssert(CPBigBoxCurrentPage = testPageKey, "Shoulders reach " testPageKey)
        TestAssert(CPBigBoxControls["pageNext"].Enabled, "Full settings page retains arrows")
        if testPageKey != "controls"
            TestAssert(InStr(CPBigBoxControls["previewTitle"].Text, "Full settings page"), "Full page preview is labeled")
        TestLayout(3840, 2160)
    }
    TestAssert(CPBigBoxCurrentPage = "paths", "Paths is last when enabled")
    CPBigBoxSwitchPage(1)
    TestAssert(CPBigBoxCurrentPage = "home", "Paths wraps to Home")
    CPBigBoxSetPage("paths")
    showPathsTab := false
    CPBigBoxDashboardUpdateContent()
    TestAssert(CPBigBoxCurrentPage = "home", "Hiding the active Paths page returns safely to Home")
    TestAssert(!CPBigBoxSetPage("paths"), "Hidden Paths page cannot be opened")
    TestAssert(CPBigBoxPageOrder().Length = 10, "Paths follows desktop optional visibility")
    TestAssert(CPBigBoxPageIndicator["w"] = Floor(CPBigBoxPageIndicator["width"] / 10)
        - CPBigBoxPageIndicator["gap"], "Indicator updates after optional tab removal")
    CPBigBoxSwitchPage(-1)
    TestAssert(CPBigBoxCurrentPage = "apiKeys", "Previous wraps Home to API Keys")
    CPBigBoxSwitchPage(1)
    TestAssert(CPBigBoxCurrentPage = "home", "Next wraps API Keys to Home")
    for testTile in CPBigBoxHomeTiles() {
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls[testTile[1]]))
        if testTile[1] = "audioToggle" {
            CPBigBoxOpenHomeTile(testTile[3])
            TestAssert(CPBigBoxCurrentPage = "home" && TestLiveAudioRunning, "Home audio tile starts directly without opening a preview")
            CPBigBoxOpenHomeTile(testTile[3])
            TestAssert(!TestLiveAudioRunning, "Home audio tile stops directly")
            continue
        }
        if testTile[1] = "study" {
            continue
        }
        CPBigBoxOpenHomeTile(testTile[3])
        TestAssert(CPBigBoxCurrentPage = testTile[3], "Tile opens " testTile[3])
        TestAssert(!CPBigBoxMainPageIndex(testTile[3]), "Home shortcut is not a full settings page")
        TestAssert(!CPBigBoxControls["pageNext"].Enabled, "Quick view has no active page-switching arrow")
        TestAssert(!TestControlShown(CPBigBoxControls["pageNext"]), "Quick view hides page arrows")
        TestAssert(CPBigBoxNavigationControls.Length = (CPBigBoxAIDomain() != "" || testTile[3] = "quickCapture"
            || testTile[3] = "quickControls" ? 6 : testTile[3] = "quickOverlays" ? 6 : 3),
            "Quick view focus excludes hidden page arrows")
        TestAssert(InStr(CPBigBoxControls["pageHint"].Text, "Home ›"), "Quick view shows Home breadcrumb")
        CPBigBoxSwitchPage(1)
        CPBigBoxSwitchPage(-1)
        TestAssert(CPBigBoxCurrentPage = testTile[3], "Shoulders leave the separate quick view unchanged")
        TestAssert(!CPBigBoxControls["translation"].Enabled, "Home actions inactive away from Home")
        TestAssert(CPBigBoxControls["backHome"].Enabled, "Preview has Back to Home")
        CPBigBoxBack()
        TestAssert(CPBigBoxCurrentPage = "home", "Back returns from " testTile[3])
        TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd
            = CPBigBoxControls[testTile[1]].Hwnd, "Home restores tile focus for " testTile[3])
    }
    CPBigBoxSetPage("screenshot")
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["advanced"]))
    CPBigBoxSetPage("explanation")
    CPBigBoxSetPage("screenshot")
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd
        = CPBigBoxControls["advanced"].Hwnd, "Each page remembers its own focused control")
    CPBigBoxSetPage("quickTranslation")
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["return"]))
    CPBigBoxSetPage("screenshot")
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["advanced"].Hwnd,
        "Full screenshot page focus is independent of Translation AI quick view")
    CPBigBoxSetPage("quickTranslation")
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["return"].Hwnd,
        "Translation AI quick view remembers its own focus")
    CPBigBoxSetPage("screenshot")
    CPBigBoxModalDepth := 1
    CPBigBoxSwitchPage(1)
    CPBigBoxBack()
    TestAssert(CPBigBoxCurrentPage = "screenshot", "Modal guard blocks page switching and Back")
    CPBigBoxModalDepth := 0
    CPBigBoxGui.Opt("+Disabled")
    CPBigBoxSwitchPage(1)
    TestAssert(CPBigBoxCurrentPage = "screenshot", "Disabled owner cannot switch pages")
    CPBigBoxGui.Opt("-Disabled")
    CPBigBoxSetPage("home")
    ; Test production edge detection with XInput snapshots, without a device.
    hwnd := CPBigBoxGui.Hwnd
    CPControllerResetNavigation()
    CPControllerHandleNavigation(TestSnapshot(), hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:RB"), hwnd)
    TestAssert(CPBigBoxCurrentPage = "screenshot", "RB opens full Game Text Translation, not quick AI settings")
    Loop 12
        CPControllerHandleNavigation(TestSnapshot("X:RB"), hwnd)
    TestAssert(CPBigBoxCurrentPage = "screenshot", "Held RB does not skip pages")
    CPControllerHandleNavigation(TestSnapshot(), hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:B"), hwnd)
    TestAssert(CPBigBoxCurrentPage = "home", "B returns to Home")
    Loop 12
        CPControllerHandleNavigation(TestSnapshot("X:B"), hwnd)
    TestAssert(TestExternalCalls = 0, "Held B does not also return to game")
    CPControllerHandleNavigation(TestSnapshot(), hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:LB"), hwnd)
    TestAssert(CPBigBoxCurrentPage = "apiKeys", "LB wraps one page")
    CPBigBoxSetPage("home")
    ; A fullscreen surface transition blocks all commands while busy and until
    ; the initiating button has really been released.
    CPControllerBeginSurfaceTransition()
    CPControllerHandleNavigation(TestSnapshot("X:RB"), hwnd)
    TestAssert(CPBigBoxCurrentPage = "home"
        && CPControllerSurfaceTransitionState = "busy",
        "Controller commands are blocked while a fullscreen handoff is busy")
    CPControllerFinishSurfaceTransition()
    CPControllerHandleNavigation(TestSnapshot("X:RB"), hwnd)
    TestAssert(CPBigBoxCurrentPage = "home"
        && CPControllerSurfaceTransitionState = "release",
        "Held controller input remains blocked on the destination surface")
    CPControllerHandleNavigation(TestSnapshot(), hwnd)
    TestAssert(CPControllerSurfaceTransitionState = "",
        "A physical release arms the destination surface")
    CPControllerHandleNavigation(TestSnapshot("X:RB"), hwnd)
    TestAssert(CPBigBoxCurrentPage = "screenshot",
        "A fresh controller press works after the transition guard")
    CPBigBoxSetPage("home")
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["translation"]))
    CPControllerHandleNavigation(TestSnapshot(), hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), hwnd)
    Sleep(30)
    TestAssert(CPBigBoxCurrentPage = "quickTranslation", "Native A opens Translation AI quick view, not full settings")
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["backHome"]))
    Loop 12
        CPControllerHandleNavigation(TestSnapshot("X:A"), hwnd)
    Sleep(30)
    TestAssert(CPBigBoxCurrentPage = "quickTranslation", "Held A does not activate the quick view's Back button")
    CPControllerHandleNavigation(TestSnapshot(), hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), hwnd)
    Sleep(30)
    TestAssert(CPBigBoxCurrentPage = "home", "Released and pressed A activates Back to Home")
    CPControllerHandleNavigation(TestSnapshot(), hwnd)
    CPControllerHandleNavigation(Map("name", "Legacy USB gamepad", "tokens", Map("J:Button 6", true)), hwnd)
    TestAssert(CPBigBoxCurrentPage = "screenshot", "Legacy right shoulder switches full pages")
    ; Owned dialogs must never forward shoulder presses to the dashboard.
    CPControllerResetNavigation()
    CPControllerHandleNavigation(TestSnapshot(), 12345)
    CPControllerHandleNavigation(TestSnapshot("X:RB"), 12345)
    TestAssert(CPBigBoxCurrentPage = "screenshot", "Owned dialog suppresses shoulders")
    CPBigBoxSetPage("home")
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["translation"]))
    CPBigBoxDashboardMoveFocus("Up")
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["pagePrevious"].Hwnd,
        "D-pad can reach previous-page button")
    CPBigBoxDashboardMoveFocus("Right")
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["pageNext"].Hwnd,
        "D-pad can reach next-page button")
    CPBigBoxDashboardActivate()
    Sleep(30)
    TestAssert(CPBigBoxCurrentPage = "screenshot", "Clickable next arrow opens full settings page")
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["pageNext"].Hwnd,
        "Chevron activation keeps chevron focused")
    CPBigBoxSetPage("home")
    controlDarkMode := 0
    CPBigBoxDashboardApplyTheme()
    TestAssert(CPBigBoxGui.BackColor = "F0F0F0", "Light theme remains available")
    controlDarkMode := 1
    CPBigBoxDashboardApplyTheme()

    TestAISelectors()
    TestLongChoiceLists()
    TestExplanationPreferences()
    TestScreenshotPreferences()
    TestAudioInputs()
    TestBigBoxStage3D()
    TestBigBoxAudioFeedback()
    TestBigBoxAudioStartup()
    TestBigBoxOverlays()
    TestBigBoxBackgroundOpacity()
    TestBigBoxAlwaysOnTop()
    TestNativePickerOwnership()
    TestBigBoxControls()
    TestBigBoxTerminologyProfiles()
    TestBigBoxModelManagement()
    TestBigBoxSetupManagement()
    TestPickerReturnFocus()
    TestGroupedFocusAndPainting()

    CPBigBoxSetPage("home", false)
    CPBigBoxDashboardShowReady()
    CPBigBoxOpenHomeTile("study")
    TestAssert(TestStudyLibraryOpens = 1,
        "Home Study tile opens the real Study Library")
    TestAssert(CPStudyLibraryState.Get("returnToBigBox", false),
        "Study Library records its Big Box return route")
    TestAssert(TestStudyLibraryBigBox
        && CPStudyLibraryState.Get("bigBoxPresentation", false),
        "Home Study tile requests the fullscreen Big Box Library presentation")
    TestAssert(TestStudyLibraryOpenedOverDashboard,
        "Dashboard remains visible until the fullscreen Study Library is ready")
    TestAssert(!CPBigBoxDashboardVisible(),
        "Dashboard hides while the Study Library is active")
    CPStudyLibraryState := 0
    CPBigBoxReturnFromStudy()
    TestAssert(CPBigBoxCurrentPage = "home",
        "Study Library returns to the Home page")
    CPBigBoxDashboardHide(false)

    ; Exercise animation completion math and shutdown without a visible window.
    CPBigBoxPageAnimation := Map("active", true, "from", 0, "to", 20, "start", A_TickCount)
    CPBigBoxPageAnimationTick()
    TestAssert(!CPBigBoxPageAnimation["active"], "Hidden window stops animation")
    CPBigBoxGui.Destroy()
    CPBigBoxPageAnimationTick()
    TestAssert(!CPBigBoxDashboardAlive(), "Destroyed GUI is safe for timer cleanup")
    TestAssert(TestExternalCalls = 0, "Preview navigation invoked no application action")
    FileAppend("PASS: " TestCount " Big Box assertions.`n", "*", "UTF-8")
    ExitApp(0)
} catch as harnessFailure {
    FileAppend("FAIL: " harnessFailure.Message "`n" harnessFailure.Stack "`n", "*", "UTF-8")
    ExitApp(1)
}

TestCombo(items, selected := 1) {
    global ui
    control := ui.AddDropDownList("w300", items)
    control.Choose(selected)
    return control
}

TestAISelectors() {
    global
    AutoPersist()
    for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
        CPBigBoxGui.Show("Hide w" size[1] " h" size[2])
        CPBigBoxDashboardResize(CPBigBoxGui, 0, size[1], size[2])
        CPBigBoxSetPage("quickTranslation", false)
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ai_model"]))
        CPBigBoxOpenAIChoice("model")
        TestAssert(CPBigBoxAIListActive(), "Long model selector opens as a scrolling list")
        TestAssert(CPBigBoxAIChoice["index"] = 1, "List starts on saved model")
        TestAssert(CPBigBoxNavigationControls.Length = 3, "Model list, Manage models and Cancel are bounded focus stops")
        TestAssert(!TestControlShown(CPBigBoxControls["choice1"]), "Long list does not display choice tiles")
        TestLayout(size[1], size[2])
        TestCapture("model-choices-" size[1] ".png", size[1], size[2])
        CPBigBoxAIListScroll(4)
        TestAssert(CPBigBoxAIChoice["index"] = 5, "Scrolling reaches middle of list without paging buttons")
        TestLayout(size[1], size[2])
        TestCapture("model-middle-" size[1] ".png", size[1], size[2])
        CPBigBoxAIListBoundary(true)
        TestAssert(CPBigBoxAIChoice["index"] = 9, "End reaches last item")
        TestAssert(!CPBigBoxControls["listChoice4"].Enabled && !CPBigBoxControls["listChoice5"].Enabled,
            "Rows past end are hidden and disabled")
        CPBigBoxDashboardMoveFocus("Down")
        TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["choiceManage"].Hwnd,
            "Down from the final model reaches Manage models")
        CPBigBoxDashboardMoveFocus("Up")
        TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["listChoice3"].Hwnd
            && CPBigBoxAIChoice["index"] = 9,
            "Up from Manage models returns to the final model without changing it")
        CPBigBoxAIListScroll(1)
        TestAssert(CPBigBoxAIChoice["index"] = 9, "Cannot scroll beyond last choice")
        TestCapture("model-last-" size[1] ".png", size[1], size[2])
        CPBigBoxSwitchPage(1)
        TestAssert(CPBigBoxCurrentPage = "quickTranslation", "Shoulders do not leave picker")
        CPBigBoxBack()
        TestAssert(!CPBigBoxAIChoiceActive() && CPBigBoxCurrentPage = "quickTranslation", "B cancels one layer only")
        TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["ai_model"].Hwnd,
            "Cancel restores originating setting focus")
        TestAssert(ddlIMG_GM.Text = "gemini-test" && IniRead(iniPath, "cfg", "geminiImgModel") = "gemini-test",
            "Browsing and cancelling never saves a selection")
    }

    for testPage in ["quickTranslation", "screenshot", "quickExplanation", "explanation", "quickAudio", "audio"] {
        CPBigBoxSetPage(testPage)
        TestAssert(TestControlShown(CPBigBoxControls["ai_provider"]), testPage " has working selectors")
        CPBigBoxOpenAIChoice("provider")
        TestAssert(!CPBigBoxSetPage("home") && !CPBigBoxControls["advanced"].Enabled,
            "Picker blocks hidden navigation and Advanced Settings")
        CPBigBoxCommitAIChoice(CPBigBoxAIChoice["value"] = "Gemini" ? 1 : 2)
        TestAssert(InStr(CPBigBoxAINotice, "Already selected"), "Selecting current provider makes no changes")
    }

    CPBigBoxSetPage("quickTranslation")
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ai_provider"]))
    CPBigBoxOpenAIChoice("provider")
    CPBigBoxFocusAIChoice(2)
    testHwnd := CPBigBoxGui.Hwnd
    CPControllerResetNavigation()
    CPControllerHandleNavigation(TestSnapshot(), testHwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), testHwnd)
    Sleep(30)
    TestAssert(!CPBigBoxAIChoiceActive() && imgProvider = "OpenAI", "Controller A confirms provider")
    TestAssert(ddlProv.Text = "OpenAI" && IniRead(iniPath, "cfg", "imgProvider") = "OpenAI",
        "Provider updates native desktop control and real INI persistence")
    TestAssert(ddlIMG.Enabled && !ddlIMG_GM.Enabled, "Provider applies desktop model enablement")
    Loop 12
        CPControllerHandleNavigation(TestSnapshot("X:A"), testHwnd)
    Sleep(30)
    TestAssert(!CPBigBoxAIChoiceActive(), "Held confirm cannot reopen settings after applying")
    CPBigBoxOpenAIChoice("model")
    CPBigBoxCommitAIChoice(2)
    TestAssert(imgModel = "gpt-alternate" && EnvGet("MODEL_NAME") = "gpt-alternate", "Screenshot model updates runtime environment")
    CPBigBoxOpenAIChoice("provider")
    CPBigBoxCommitAIChoice(1)
    TestAssert(geminiImgModel = "gemini-test" && imgModel = "gpt-alternate", "Each provider keeps its own model")
    CPBigBoxOpenAIChoice("model")
    CPBigBoxAIListScroll(4)
    CPBigBoxAIListClick(3)
    TestAssert(geminiImgModel = "model-five" && IniRead(iniPath, "cfg", "geminiImgModel") = "model-five"
        && EnvGet("GEMINI_MODEL_NAME") = "models/model-five", "Scrolled model choice saves correct underlying value")
    CPBigBoxOpenAIChoice("detail")
    CPBigBoxCommitAIChoice(2)
    TestAssert(promptProfile = "literal" && EnvGet("PROMPT_PROFILE") = "literal"
        && IniRead(iniPath, "cfg", "promptProfile") = "literal" && EnvGet("POSTPROC_MODE") = "literal",
        "Screenshot prompt uses existing postprocessing and persistence")

    CPBigBoxSetPage("quickExplanation")
    CPBigBoxOpenAIChoice("model")
    CPBigBoxCommitAIChoice(2)
    TestAssert(explainOpenAIModel = "gpt-explain-alternate"
        && IniRead(iniPath, "cfg_explainer", "explainOpenAIModel") = "gpt-explain-alternate",
        "Explanation model saves independently")
    CPBigBoxOpenAIChoice("detail")
    CPBigBoxCommitAIChoice(2)
    TestAssert(explainPromptProfile = "detailed" && IniRead(iniPath, "cfg", "explainPromptProfile") = "detailed",
        "Explanation prompt calls its independent desktop handler")
    TestAssert(promptProfile = "literal" && imgModel = "gpt-alternate", "Explanation changes do not alter translation selections")
    CPBigBoxSetPage("explanation")
    TestAssert(InStr(CPBigBoxControls["ai_detail"].Text, "detailed"), "Full page reflects Home quick choice")
    ddlEPr.Choose(1)
    ExplainPromptChanged()
    CPBigBoxDashboardUpdateContent()
    TestAssert(InStr(CPBigBoxControls["ai_detail"].Text, "default"), "Dashboard reflects desktop prompt changes")

    CPBigBoxSetPage("audio")
    CPBigBoxOpenAIChoice("provider")
    CPBigBoxCommitAIChoice(2)
    CPBigBoxOpenAIChoice("model")
    CPBigBoxCommitAIChoice(2)
    CPBigBoxOpenAIChoice("detail")
    CPBigBoxCommitAIChoice(3)
    TestAssert(audioProvider = "OpenAI" && trModel = "audio-alternate" && audioTargetLang = "French",
        "Audio provider model and language all update")
    TestAssert(IniRead(iniPath, "cfg", "audioTargetLanguage") = "French" && ddlTR.Enabled && !ddlA_GM.Enabled,
        "Audio changes persist and synchronize desktop controls")
    TestAssert(TestExternalCalls = 0 && InStr(CPBigBoxControls["aiNote"].Text, "next audio start"),
        "Audio changes never start or stop external processes and explain when applied")

    CPBigBoxSetPage("screenshot")
    CPBigBoxOpenAIChoice("provider")
    IniWrite("Another profile", iniPath, "game_profiles", "active")
    CPBigBoxCommitAIChoice(2)
    TestAssert(imgProvider = "Gemini" && InStr(CPBigBoxAINotice, "Settings changed"), "Profile change rejects stale picker")
    IniWrite("Demo", iniPath, "game_profiles", "active")
    CPBigBoxOpenAIChoice("model")
    ddlProv.Choose(2)
    AutoPersist()
    CPBigBoxAIListClick(3)
    TestAssert(imgModel = "gpt-alternate" && InStr(CPBigBoxAINotice, "Settings changed"), "Provider change rejects stale model picker")
    CPBigBoxOpenAIChoice("model")
    ddlIMG.Delete(1)
    CPBigBoxCommitAIChoice(1)
    TestAssert(InStr(CPBigBoxAINotice, "no longer available"), "Removed model cannot be selected from old list")
    ddlIMG.Delete()
    ddlIMG.Add(model_openai_img)
    ddlIMG.Choose(2)

    CPBigBoxSetPage("quickExplanation")
    ddlEPr.Delete()
    CPBigBoxOpenAIChoice("detail")
    TestAssert(!CPBigBoxAIChoiceActive() && InStr(CPBigBoxAINotice, "No choices"), "Empty list gives useful message without blank picker")
    ddlEPr.Add(["default", "detailed"])
    ddlEPr.Choose(1)

    testRealIni := iniPath
    iniPath := A_ScriptDir "\missing-parent\cannot-save.ini"
    CPBigBoxOpenAIChoice("detail")
    CPBigBoxCommitAIChoice(2)
    TestAssert(InStr(CPBigBoxAINotice, "Could not save"), "Persistence failure is reported instead of claiming success")
    iniPath := testRealIni
    ExplainPromptChanged()
    CPBigBoxOpenAIChoice("provider")
    CPBigBoxDashboardHide(false)
    TestAssert(!CPBigBoxAIChoiceActive(), "Hiding dashboard cancels outstanding choice")
    CPBigBoxSetPage("home")
}

TestLongChoiceLists() {
    global
    CPBigBoxSetPage("quickAudio")
    ; Use the production language definitions, not the synthetic short lists
    ; below, to check both presentations and saved selection after reordering.
    ddlAudioTarget.Delete()
    ddlAudioTarget.Add(TestAudioTargetLangs)
    ddlAudioTarget.Text := "English (en)"
    AutoPersist()
    desktopLanguages := ControlGetItems(ddlAudioTarget.Hwnd)
    TestAssert(desktopLanguages.Length = 14, "All audio output languages remain available")
    CPBigBoxOpenAIChoice("detail")
    for languageIndex, language in desktopLanguages {
        if languageIndex > 1
            TestAssert(StrCompare(desktopLanguages[languageIndex-1], language, false) < 0,
                "Desktop audio languages are alphabetical: " language)
        TestAssert(CPBigBoxAIChoice["options"][languageIndex] = language,
            "Fullscreen audio languages match the desktop order: " language)
    }
    TestAssert(CPBigBoxAIChoice["index"] = 4 && ddlAudioTarget.Text = "English (en)"
        && audioTargetLang = "English (en)", "Reordering preserves the saved language instead of its old index")
    CPBigBoxAIListBoundary(false)
    CPBigBoxBack()
    TestAssert(ddlAudioTarget.Text = "English (en)", "Cancelling language browsing preserves English")
    CPBigBoxOpenAIChoice("detail")
    CPBigBoxCommitAIChoiceIndex(6)
    TestAssert(ddlAudioTarget.Text = "German (de)" && audioTargetLang = "German (de)"
        && IniRead(iniPath, "cfg", "audioTargetLanguage") = "German (de)",
        "Selecting from the sorted list persists the matching language label and code")
    ddlAudioTarget.Delete()
    ddlAudioTarget.Add(["English", "German", "French", "Spanish"])
    ddlAudioTarget.Choose(3)
    AutoPersist()
    CPBigBoxOpenAIChoice("detail")
    TestAssert(!CPBigBoxAIListActive() && TestControlShown(CPBigBoxControls["choice4"]),
        "Exactly four choices remain tiles")
    TestAssert(!TestControlShown(CPBigBoxControls["listChoice3"]), "Short choices hide the vertical list")
    CPBigBoxBack()
    ddlAudioTarget.Add(["Italian"])
    CPBigBoxOpenAIChoice("detail")
    TestAssert(CPBigBoxAIListActive() && CPBigBoxAIChoice["index"] = 3,
        "Five language choices use list centered on current selection")
    testListConfigBefore := FileRead(iniPath)
    testListHwnd := CPBigBoxGui.Hwnd
    CPControllerResetNavigation()
    CPControllerHandleNavigation(TestSnapshot(), testListHwnd)
    CPControllerHandleNavigation(TestSnapshot("X:DPAD_DOWN"), testListHwnd)
    TestAssert(CPBigBoxAIChoice["index"] = 4, "Native controller Down advances one list item")
    CPControllerHandleNavigation(TestSnapshot(), testListHwnd)
    CPControllerHandleNavigation(TestSnapshot("X:DPAD_UP"), testListHwnd)
    TestAssert(CPBigBoxAIChoice["index"] = 3, "Native controller Up moves back one list item")
    CPBigBoxAIListWheel(1)
    TestAssert(CPBigBoxAIChoice["index"] = 4, "Wheel scrolls one item directly")
    CPBigBoxAIListWheel(-1)
    TestAssert(CPBigBoxAIChoice["index"] = 3, "Reverse wheel scrolls back")
    CPBigBoxDashboardMoveFocus("Right")
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["choiceBack"].Hwnd,
        "Cancel remains reachable by controller without scrolling to list end")
    CPBigBoxDashboardMoveFocus("Up")
    TestAssert(CPBigBoxAIChoice["index"] = 3 && CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd
        = CPBigBoxControls["listChoice3"].Hwnd, "Up from Cancel restores pending item unchanged")
    CPBigBoxKeyboardTraverse(1)
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["choiceBack"].Hwnd,
        "Tab reaches Cancel")
    CPBigBoxKeyboardTraverse(-1)
    TestAssert(CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd = CPBigBoxControls["listChoice3"].Hwnd,
        "Shift Tab returns to list")
    TestAssert(FileRead(iniPath) = testListConfigBefore, "Browsing with controller or wheel never writes settings")
    TestAssert(!CPBigBoxAIListWheelActive(), "Hidden dashboard cannot capture mouse wheel")
    CPBigBoxAIListBoundary(false)
    CPBigBoxDashboardMoveFocus("Up")
    TestAssert(CPBigBoxAIChoice["index"] = 1 && !CPBigBoxControls["listChoice2"].Enabled,
        "Top boundary clamps without wrapping or phantom rows")
    CPBigBoxAIListBoundary(true)
    CPControllerHandleNavigation(TestSnapshot(), testListHwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), testListHwnd)
    Sleep(30)
    TestAssert(!CPBigBoxAIChoiceActive() && audioTargetLang = "Italian", "A confirms centered item and updates language")
    Loop 12
        CPControllerHandleNavigation(TestSnapshot("X:A"), testListHwnd)
    Sleep(30)
    TestAssert(!CPBigBoxAIChoiceActive(), "Held A does not reopen the selector after list confirmation")
    CPBigBoxOpenAIChoice("detail")
    CPBigBoxAIListScroll(-2)
    CPControllerHandleNavigation(TestSnapshot(), testListHwnd)
    CPControllerHandleNavigation(TestSnapshot("X:B"), testListHwnd)
    Loop 12
        CPControllerHandleNavigation(TestSnapshot("X:B"), testListHwnd)
    TestAssert(!CPBigBoxAIChoiceActive() && CPBigBoxCurrentPage = "quickAudio" && audioTargetLang = "Italian",
        "Held B only cancels list and keeps saved language")
    CPBigBoxOpenAIChoice("detail")
    CPBigBoxAIListBoundary(false)
    CPBigBoxAIListClick(4)
    TestAssert(audioTargetLang = "German", "Clicking neighboring row applies that exact item")
    CPBigBoxOpenAIChoice("detail")
    CPBigBoxAIListBoundary(true)
    CPControllerResetNavigation()
    CPControllerHandleNavigation(Map("name", "Legacy USB gamepad", "tokens", Map()), testListHwnd)
    CPControllerHandleNavigation(Map("name", "Legacy USB gamepad", "tokens", Map("J:Button 1", true)), testListHwnd)
    Sleep(30)
    TestAssert(!CPBigBoxAIChoiceActive() && audioTargetLang = "Italian", "Legacy controller confirms list item")

    CPBigBoxSetPage("quickExplanation")
    ddlEPr.Delete()
    testManyPrompts := []
    Loop 100
        testManyPrompts.Push("Prompt " A_Index)
    testManyPrompts[50] := "English & Japanese [study]"
    ddlEPr.Add(testManyPrompts)
    ddlEPr.Choose(50)
    ExplainPromptChanged()
    CPBigBoxOpenAIChoice("detail")
    TestAssert(CPBigBoxAIChoice["index"] = 50 && InStr(CPBigBoxControls["listChoice3"].Text, "English && Japanese"),
        "Long prompt list opens at saved item and preserves literal ampersands")
    TestAssert(CPBigBoxNavigationControls.Length = 3,
        "Hundred-item prompt list keeps one centered item plus Manage and Cancel")
    CPBigBoxAIListScroll(5)
    TestAssert(CPBigBoxAIChoice["index"] = 55, "Larger list scroll step works")
    CPBigBoxAIListBoundary(true)
    TestAssert(InStr(CPBigBoxControls["pageHint"].Text, "Item 100 of 100"), "Counter shows exact position in long list")
    CPBigBoxAIListClick(3)
    TestAssert(explainPromptProfile = "Prompt 100" && IniRead(iniPath, "cfg", "explainPromptProfile") = "Prompt 100",
        "Long prompt list retains existing independent save handler")
    ddlEPr.Delete()
    ddlEPr.Add(["default", "detailed"])
    ddlEPr.Choose(1)
    ExplainPromptChanged()
    CPBigBoxSetPage("home")
}

TestExplanationPreferences() {
    global
    CPSetExplanationPreference("library", 1)
    CPSetExplanationPreference("screenshots", 1)
    CPSetExplanationPreference("plainText", 0)
    CPSetExplanationPreference("openOnStartup", 0)
    CPSetExplanationPreference("alwaysOnTop", 0)
    IniWrite("keep", iniPath, "test_unrelated", "sentinel")
    testPrefConfigBefore := FileRead(iniPath)
    CPBigBoxSetPage("explanation")
    TestAssert(CPBigBoxNavigationControls.Length = 13, "Full Explanation page includes eight settings and navigation")
    TestAssert(FileRead(iniPath) = testPrefConfigBefore, "Opening settings page does not change configuration")
    for testOption in CPBigBoxExplanationOptions() {
        testPref := CPExplanationPreference(testOption)
        testValueBefore := testPref["control"].Value
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["exp_" testOption]))
        TestAssert(InStr(CPBigBoxControls["modeBody"].Text, testPref["help"]), "Focused setting explains its effect: " testOption)
        CPBigBoxToggleExplanationOption(testOption)
        TestAssert(testPref["control"].Value = !testValueBefore, "Toggle updates desktop checkbox: " testOption)
        TestAssert(IniRead(iniPath, testPref["section"], testPref["name"]) = !testValueBefore,
            "Toggle persists correct key: " testOption)
        TestAssert(InStr(CPBigBoxAINotice, "Saved"), "Toggle confirms successful save: " testOption
            . " (notice=" CPBigBoxAINotice ", focus=" CPBigBoxPageFocus.Get("explanation", "") ")")
        if testOption = "library" {
            TestAssert(!saveLibraryScreenshotsChk.Enabled && !CPBigBoxControls["exp_screenshots"].Enabled,
                "Library off disables screenshot option in both interfaces")
            TestAssert(saveLibraryScreenshotsChk.Value = 1 && IniRead(iniPath, "cfg", "studyLibraryScreenshots") = 1,
                "Library off preserves screenshot preference")
            TestAssert(!CPBigBoxDashboardControlIndex(CPBigBoxControls["exp_screenshots"]),
                "Disabled screenshot option is excluded from controller navigation")
            CPBigBoxToggleExplanationOption("screenshots")
            TestAssert(saveLibraryScreenshotsChk.Value = 1, "Disabled screenshot toggle cannot be activated indirectly")
            CPBigBoxGui.Show("Hide w1280 h720")
            CPBigBoxDashboardResize(CPBigBoxGui, 0, 1280, 720)
            TestLayout(1280, 720)
            TestCapture("explanation-library-off-1280.png", 1280, 720)
        }
        CPBigBoxToggleExplanationOption(testOption)
        TestAssert(testPref["control"].Value = testValueBefore, "Toggle restores original value: " testOption)
    }
    TestAssert(saveLibraryScreenshotsChk.Enabled && CPBigBoxControls["exp_screenshots"].Enabled,
        "Library on restores enabled screenshot setting and its remembered value")
    TestAssert(IniRead(iniPath, "test_unrelated", "sentinel") = "keep", "Unrelated configuration is untouched")
    TestAssert(TestExternalCalls = 0, "Saving/startup toggles never launch overlays or perform study actions")

    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["exp_library"]))
    testPrefHwnd := CPBigBoxGui.Hwnd
    CPControllerResetNavigation()
    CPControllerHandleNavigation(TestSnapshot(), testPrefHwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), testPrefHwnd)
    Sleep(30)
    Loop 12
        CPControllerHandleNavigation(TestSnapshot("X:A"), testPrefHwnd)
    TestAssert(saveLibraryChk.Value = 0, "Held controller A changes toggle once")
    CPControllerHandleNavigation(TestSnapshot(), testPrefHwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), testPrefHwnd)
    Sleep(30)
    TestAssert(saveLibraryChk.Value = 1, "New A press changes toggle again")

    ; Same callback as the desktop Library checkbox, and a profile-change-like
    ; update while its dependent Big Box setting currently owns focus.
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["exp_screenshots"]))
    saveLibraryChk.Value := 0
    StudyLibrarySaveToggleChanged()
    TestAssert(!CPBigBoxControls["exp_screenshots"].Enabled && CPBigBoxNavigationControls[CPBigBoxFocusIndex].Hwnd
        = CPBigBoxControls["exp_library"].Hwnd, "Disabling focused dependency restores focus to Library saving")
    saveLibraryChk.Value := 1
    StudyLibrarySaveToggleChanged()
    saveExplChk.Value := 1
    CPExplanationPreferenceChanged("plainText")
    TestAssert(InStr(CPBigBoxControls["exp_plainText"].Text, "On") && IniRead(iniPath, "cfg", "saveExplains") = 1,
        "Desktop preference callback immediately synchronizes Big Box")

    CPBigBoxSetPage("quickExplanation")
    testPrefConfigBefore := FileRead(iniPath)
    for testOption in CPBigBoxExplanationOptions() {
        TestAssert(!TestControlShown(CPBigBoxControls["exp_" testOption]) && !CPBigBoxControls["exp_" testOption].Enabled,
            "Home AI shortcut excludes full-page preference: " testOption)
        CPBigBoxToggleExplanationOption(testOption)
    }
    TestAssert(FileRead(iniPath) = testPrefConfigBefore, "Hidden preference callbacks cannot mutate settings from Home shortcut")
    CPBigBoxSetPage("explanation")
    CPBigBoxOpenAIChoice("provider")
    CPBigBoxToggleExplanationOption("plainText")
    TestAssert(FileRead(iniPath) = testPrefConfigBefore && !CPBigBoxControls["exp_plainText"].Enabled,
        "Open AI selector isolates underlying preference controls")
    CPBigBoxBack()
    CPBigBoxModalDepth := 1
    CPBigBoxToggleExplanationOption("plainText")
    CPBigBoxModalDepth := 0
    TestAssert(FileRead(iniPath) = testPrefConfigBefore, "Owned modal guard blocks preference changes")

    testPrefRealIni := iniPath
    testPrefValueBefore := saveExplChk.Value
    iniPath := A_ScriptDir "\missing-parent\cannot-save-preference.ini"
    CPBigBoxToggleExplanationOption("plainText")
    TestAssert(saveExplChk.Value = testPrefValueBefore && InStr(CPBigBoxAINotice, "Could not save"),
        "Failed preference write keeps original checkbox value and reports failure")
    iniPath := testPrefRealIni
    CPBigBoxDashboardUpdateContent()
    CPBigBoxSetPage("home")
}

TestScreenshotHintDetails(width) {
    global CPBigBoxControls, CPBigBoxAINotice
    originalNotice := CPBigBoxAINotice
    CPBigBoxAINotice := ""
    seenDetails := Map()
    try {
        for hintKey in CPBigBoxSettingsKeys("screenshot") {
            CPBigBoxUpdateSettingsHint(hintKey)
            hintText := CPBigBoxControls["modeBody"].Text
            hintLines := StrSplit(hintText, "`n", "`r")
            TestAssert(hintLines.Length = 2 && hintLines[2] != ""
                && hintLines[2] = CPBigBoxScreenshotHintDetail(hintKey)
                && !InStr(hintText, "remain in Advanced Settings"),
                "Screenshot tile has current, specific second-line help: " hintKey " at " width)
            TestAssert(!seenDetails.Has(hintLines[2]),
                "Screenshot help is not repeated between tiles: " hintKey)
            seenDetails[hintLines[2]] := true
            CPBigBoxControls["modeBody"].GetPos(,,, &hintHeight)
            TestAssert(TestTextHeight(CPBigBoxControls["modeBody"]) <= hintHeight,
                "Both screenshot help lines fit: " hintKey " at " width)
        }
        CPBigBoxAINotice := "Saved · Clear captures on startup: On"
        CPBigBoxUpdateSettingsHint("shot_clearOnStartup")
        TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "`n" CPBigBoxAINotice),
            "Screenshot save feedback still takes priority over the second help line")
        CPBigBoxAINotice := "Could not save this option."
        CPBigBoxUpdateSettingsHint("shot_clearOnStartup")
        TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "`n" CPBigBoxAINotice),
            "Screenshot failure feedback still takes priority over the second help line")
    } finally {
        CPBigBoxAINotice := originalNotice
        CPBigBoxUpdateSettingsHint()
    }
}

TestScreenshotPreferences() {
    global CPBigBoxControls, CPBigBoxNavigationControls, CPBigBoxNavigationRows, CPBigBoxAINotice
    global CPBigBoxGui, CPBigBoxAICommitting, CPBigBoxModalDepth, CPBigBoxPageFocus, CPBigBoxFocusFrame
    global iniPath, chkGuess, chkName, chkDel, chkOpenTW, chkTop_TW, TestExternalCalls
    local testOption, testPref, testValueBefore, testRealIni
    for testOption in CPBigBoxScreenshotOptions()
        CPSetScreenshotPreference(testOption, 0)
    IniWrite("keep", iniPath, "test_unrelated", "keep")
    testOriginalIni := FileRead(iniPath)
    testExplainerTop := IniRead(iniPath, "cfg_explainer", "winTop")
    testLibrarySaving := IniRead(iniPath, "cfg", "saveStudyLibrary")
    testExternalBefore := TestExternalCalls
    CPBigBoxSetPage("screenshot", false)
    TestAssert(CPBigBoxNavigationControls.Length = 14, "Screenshot page has nine settings/actions plus navigation")
    TestAssert(CPBigBoxNavigationRows.Length = 5, "Screenshot page uses three semantic setting rows")
    TestAssert(FileRead(iniPath) = testOriginalIni, "Opening Screenshot page does not change settings")
    for groupIndex, group in CPBigBoxSettingsGroups("screenshot") {
        TestAssert(CPBigBoxControls["settingsGroup" groupIndex].Text = group[1], "Screenshot caption matches group")
        for index, key in group[2]
            TestAssert(CPBigBoxNavigationRows[groupIndex + 1][index].Hwnd = CPBigBoxControls[key].Hwnd,
                "Screenshot D-pad row matches group: " key)
    }
    for testOption in CPBigBoxScreenshotOptions() {
        testPref := CPScreenshotPreference(testOption)
        testValueBefore := testPref["control"].Value
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["shot_" testOption]))
        TestAssert(InStr(CPBigBoxControls["modeBody"].Text, testPref["help"]), "Screenshot focus displays help: " testOption)
        CPBigBoxToggleScreenshotOption(testOption)
        TestAssert(testPref["control"].Value = !testValueBefore, "Screenshot toggle synchronizes desktop checkbox: " testOption)
        TestAssert(IniRead(iniPath, testPref["section"], testPref["name"]) = !testValueBefore,
            "Screenshot toggle saves correct single key: " testOption)
        TestAssert(InStr(CPBigBoxAINotice, "Saved"), "Screenshot toggle confirms save: " testOption)
        if testOption = "highlight"
            TestAssert(EnvGet("SHOT_ITALICIZE_GUESSED") = "1", "Guessed-subject runtime flag updates immediately")
        if testOption = "speakerColor"
            TestAssert(EnvGet("SHOT_COLOR_SPEAKER") = "1", "Speaker-color runtime flag updates immediately")
        CPBigBoxToggleScreenshotOption(testOption)
        TestAssert(testPref["control"].Value = testValueBefore, "Second screenshot toggle restores value: " testOption)
    }
    TestAssert(EnvGet("SHOT_ITALICIZE_GUESSED") = "0" && EnvGet("SHOT_COLOR_SPEAKER") = "0",
        "Disabling formatting synchronizes both runtime flags")
    TestAssert(IniRead(iniPath, "test_unrelated", "keep") = "keep"
        && IniRead(iniPath, "cfg_explainer", "winTop") = testExplainerTop
        && IniRead(iniPath, "cfg", "saveStudyLibrary") = testLibrarySaving,
        "Screenshot writes preserve unrelated Library and Explainer settings")
    TestAssert(TestExternalCalls = testExternalBefore, "Startup/cleanup toggles launch nothing and delete no files")

    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["shot_highlight"]))
    CPControllerResetNavigation()
    CPControllerHandleNavigation(TestSnapshot(), CPBigBoxGui.Hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
    Loop 8
        CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
    Sleep(30)
    TestAssert(chkGuess.Value = 1, "Held controller A toggles Screenshot setting only once")
    CPControllerHandleNavigation(TestSnapshot(), CPBigBoxGui.Hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
    Sleep(30)
    TestAssert(chkGuess.Value = 0, "New controller A press toggles Screenshot setting again")
    chkName.Value := 1
    CPScreenshotPreferenceChanged("speakerColor")
    TestAssert(EnvGet("SHOT_COLOR_SPEAKER") = "1" && InStr(CPBigBoxControls["shot_speakerColor"].Text, "On"),
        "Desktop formatting callback synchronizes runtime and Big Box")
    chkTop_TW.Value := 1
    CPBigBoxDashboardUpdateContent()
    TestAssert(InStr(CPBigBoxControls["shot_alwaysOnTop"].Text, "On"),
        "Dashboard reflects native controls changed by profile application")

    CPBigBoxSetPage("quickTranslation", false)
    testProtectedIni := FileRead(iniPath)
    for testOption in CPBigBoxScreenshotOptions() {
        TestAssert(!TestControlShown(CPBigBoxControls["shot_" testOption]) && !CPBigBoxControls["shot_" testOption].Enabled,
            "Home Translation AI hides full-page preference: " testOption)
        CPBigBoxToggleScreenshotOption(testOption)
    }
    TestAssert(FileRead(iniPath) = testProtectedIni, "Quick view cannot toggle hidden Screenshot preferences")
    CPBigBoxSetPage("explanation", false)
    TestAssert(CPBigBoxControls["settingsGroup2"].Text = "STUDY LIBRARY", "Explanation group is named STUDY LIBRARY")
    CPBigBoxToggleScreenshotOption("clearOnStartup")
    TestAssert(FileRead(iniPath) = testProtectedIni, "Explanation page cannot change Screenshot cleanup preference")
    for testOption in CPBigBoxScreenshotOptions()
        TestAssert(!TestControlShown(CPBigBoxControls["shot_" testOption]), "Explanation hides Screenshot-only control")
    CPBigBoxSetPage("screenshot", false)
    CPBigBoxOpenAIChoice("provider")
    CPBigBoxToggleScreenshotOption("highlight")
    TestAssert(FileRead(iniPath) = testProtectedIni && !CPBigBoxControls["shot_highlight"].Enabled,
        "AI picker isolates underlying Screenshot preferences")
    CPBigBoxBack()
    CPBigBoxModalDepth := 1
    CPBigBoxToggleScreenshotOption("highlight")
    CPBigBoxModalDepth := 0
    CPBigBoxAICommitting := true
    CPBigBoxToggleScreenshotOption("highlight")
    CPBigBoxAICommitting := false
    TestAssert(FileRead(iniPath) = testProtectedIni, "Modal and in-progress write guards block Screenshot toggles")
    testRealIni := iniPath
    iniPath := A_ScriptDir "\missing-parent\cannot-save-screenshot-setting.ini"
    for testOption in CPBigBoxScreenshotOptions() {
        testPref := CPScreenshotPreference(testOption)
        testValueBefore := testPref["control"].Value
        testGuessBefore := EnvGet("SHOT_ITALICIZE_GUESSED")
        testSpeakerBefore := EnvGet("SHOT_COLOR_SPEAKER")
        CPBigBoxToggleScreenshotOption(testOption)
        TestAssert(testPref["control"].Value = testValueBefore && InStr(CPBigBoxAINotice, "Could not save"),
            "Failed Screenshot save preserves control value and reports error: " testOption)
        TestAssert(EnvGet("SHOT_ITALICIZE_GUESSED") = testGuessBefore && EnvGet("SHOT_COLOR_SPEAKER") = testSpeakerBefore,
            "Failed Screenshot save does not change runtime flags: " testOption)
    }
    iniPath := testRealIni
    CPBigBoxSetPage("home", false)
}

TestAudioHelpAndStatus(width, height) {
    global CPBigBoxControls, CPBigBoxAINotice, CPBigBoxFocusFrame, gAudioInputStatus, txtAudioTestStatus
    originalNotice := CPBigBoxAINotice, originalStatus := gAudioInputStatus
    originalDesktopStatus := txtAudioTestStatus.Text
    CPBigBoxAINotice := ""
    seenDetails := Map()
    try {
        for hintKey in CPBigBoxSettingsKeys("audio") {
            CPBigBoxUpdateSettingsHint(hintKey)
            hintLines := StrSplit(CPBigBoxControls["modeBody"].Text, "`n", "`r")
            TestAssert(hintLines.Length = 2 && hintLines[2] != ""
                && hintLines[2] = CPBigBoxAudioHintDetail(hintKey),
                "Audio tile has current, specific second-line help: " hintKey " at " width)
            TestAssert(!seenDetails.Has(hintLines[2]), "Audio help is not repeated between tiles: " hintKey)
            seenDetails[hintLines[2]] := true
            CPBigBoxControls["modeBody"].GetPos(,,, &hintHeight)
            TestAssert(TestTextHeight(CPBigBoxControls["modeBody"]) <= hintHeight,
                "Both audio help lines fit: " hintKey " at " width)
        }
        for notice in ["Saved · Audio input: Game speakers", "Could not save this option."] {
            CPBigBoxAINotice := notice
            CPBigBoxUpdateSettingsHint("audio_device")
            TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "`n" notice),
                "Audio save/error feedback takes priority over help: " notice)
        }
        CPBigBoxAINotice := ""
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["audio_test"]))
        helpBefore := CPBigBoxControls["modeBody"].Text
        CPBigBoxControls["audio_status_label"].GetPos(&titleX, &titleY, &titleW, &titleH)
        CPBigBoxControls["audio_status"].GetPos(&statusX, &statusY, &statusW, &statusH)
        CPBigBoxControls["audio_refresh"].GetPos(&refreshX, &refreshY,, &refreshH)
        CPBigBoxControls["audio_test"].GetPos(&testX,, &testW)
        CPBigBoxControls["audio_power"].GetPos(&powerX, &powerY, &powerW, &powerH)
        CPBigBoxControls["backHome"].GetPos(, &footerY)
        TestAssert(titleX = refreshX && statusX = titleX && statusW = titleW
            && Abs(statusX + statusW - testX - testW) <= 1,
            "Input check readout aligns beneath Refresh/Test at " width)
        TestAssert(titleY > refreshY + refreshH && titleY = powerY
            && statusY >= titleY + titleH && statusY + statusH <= footerY
            && statusX > powerX + powerW,
            "Input check title/body do not overlap input buttons, session switch, or footer at " width)
        TestAssert(TestTextHeight(CPBigBoxControls["audio_status_label"]) <= titleH,
            "Input check heading fits at " width)
        for message in ["Not tested.", "Listening for audio...", "Refreshing audio devices...",
            "Audio detected. This device is ready.",
            "No audio detected. Check the selected device or the game's audio driver.",
            "Could not open the selected audio device. Refresh devices, check your audio driver and try again.",
            "Audio test ended without a result. Check the helper paths and try again.",
            "Devices refreshed. The selected device is unavailable; reconnect it or choose another input.",
            "Could not refresh audio devices. The previous list was kept. Check your audio driver and helper version."] {
            SetAudioTestStatus(message)
            expectedText := message = "Not tested."
                ? "Not tested yet. Play game audio, then select Test audio." : message
            TestAssert(CPBigBoxControls["audio_status"].Text = expectedText
                && TestTextHeight(CPBigBoxControls["audio_status"]) <= statusH,
                "Labeled audio input message fits at " width ": " message)
            TestAssert(gAudioInputStatus = message && txtAudioTestStatus.Text = "Test status: " message,
                "Fullscreen presentation does not change desktop diagnostic wording")
            TestAssert(CPBigBoxFocusFrame["key"] = "audio_test"
                && CPBigBoxControls["modeBody"].Text = helpBefore,
                "Input status updates preserve focus and tile help")
        }
        TestCapture("audio-input-error-" width ".png", width, height)
        for key in ["audio_status_label", "audio_status"]
            TestAssert(TestControlShown(CPBigBoxControls[key])
                && !CPBigBoxDashboardControlIndex(CPBigBoxControls[key]),
                "Input check text is visible without adding a controller stop: " key)
        CPBigBoxOpenAIChoice("device")
        for key in ["audio_status_label", "audio_status"]
            TestAssert(!TestControlShown(CPBigBoxControls[key]), "Input check hides inside device picker: " key)
        CPBigBoxBack()
        for page in ["home", "quickAudio", "screenshot"] {
            CPBigBoxSetPage(page, false)
            for key in ["audio_status_label", "audio_status"]
                TestAssert(!TestControlShown(CPBigBoxControls[key]), "Input check hides on " page ": " key)
        }
    } finally {
        CPBigBoxAINotice := originalNotice
        SetAudioTestStatus(originalStatus)
        txtAudioTestStatus.Text := originalDesktopStatus
        CPBigBoxSetPage("audio", false)
    }
}

TestAudioInputs() {
    global CPBigBoxControls, CPBigBoxNavigationControls, CPBigBoxAIChoice, CPBigBoxAINotice, CPBigBoxGui
    global ddlSpeaker, speakerName, iniPath, gAudioInputJob, gAudioInputStatus, btnSpRef, btnAudioTest
    global pythonExe, audioScript, TestLiveAudioRunning, CPBigBoxFocusFrame, TestExternalCalls
    audOriginalIni := iniPath, audOriginalPython := pythonExe, audOriginalScript := audioScript
    audEnv := Map()
    for name in ["SPEAKER_NAME", "AUDIO_DEVICE_LIST_RESULT_FILE", "AUDIO_TEST_RESULT_FILE", "JRPG_TEST_AUDIO_MODE"]
        audEnv[name] := EnvGet(name)
    try {
        CPBigBoxSetPage("audio", false)
        TestAssert(CPBigBoxNavigationControls.Length = 12, "Audio page has seven buttons plus navigation; status is not focusable")
        TestAssert(!CPBigBoxDashboardControlIndex(CPBigBoxControls["audio_status"]), "Audio status never takes controller focus")
        for audSize in [[1280, 720], [1920, 1080], [3840, 2160]] {
            CPBigBoxDashboardResize(CPBigBoxGui, 0, audSize[1], audSize[2])
            for result in ["JRPG_AUDIO_TEST:DETECTED", "JRPG_AUDIO_TEST:SILENT", "JRPG_AUDIO_TEST:ERROR", ""] {
                AudioInputApplyTestResult(result)
                CPBigBoxControls["audio_status"].GetPos(,, &statusWidth, &statusHeight)
                TestAssert(TestTextHeight(CPBigBoxControls["audio_status"]) <= statusHeight,
                    "Inline audio result fits at " audSize[1] "x" audSize[2] ": " result)
            }
            SetAudioTestStatus("Devices refreshed. The selected device is unavailable; reconnect it or choose another input.")
            TestAssert(TestTextHeight(CPBigBoxControls["audio_status"]) <= statusHeight,
                "Unavailable device notice fits at " audSize[1] "x" audSize[2])
        }
        CPBigBoxDashboardResize(CPBigBoxGui, 0, 1280, 720)
        EnvSet("SPEAKER_NAME", "running-session-input")
        CPSetAudioDeviceSelection("Game speakers")
        audBefore := FileRead(iniPath)
        CPBigBoxAudioInputAction("device")
        TestAssert(CPBigBoxAIListActive() && CPBigBoxAIChoice["field"] = "device", "Many audio devices use the scrolling list")
        CPBigBoxAIListScroll(2)
        CPBigBoxBack()
        TestAssert(FileRead(iniPath) = audBefore && ddlSpeaker.Text = "Game speakers", "Cancelling device choice does not save")
        CPBigBoxAudioInputAction("device")
        CPBigBoxCommitAIChoiceIndex(6)
        TestAssert(ddlSpeaker.Text = "日本語 ＆ 音声" && speakerName = ddlSpeaker.Text,
            "Device picker preserves Unicode device names exactly")
        TestAssert(IniRead(iniPath, "cfg", "speakerName") = ddlSpeaker.Text, "Device choice persists to the shared INI")
        TestAssert(EnvGet("SPEAKER_NAME") = "running-session-input", "Selecting an input does not alter the running process environment")

        CPBigBoxAudioInputAction("device")
        iniPath := A_ScriptDir "\missing-parent\cannot-save-device.ini"
        CPBigBoxAIChoice["profile"] := "" ; Reach the write-failure path, not the separate stale-profile guard.
        CPBigBoxCommitAIChoiceIndex(1)
        TestAssert(ddlSpeaker.Text = "日本語 ＆ 音声" && InStr(CPBigBoxAINotice, "Could not save"),
            "Failed device save retains the original input")
        iniPath := audOriginalIni
        ddlSpeaker.Choose(2)
        SpeakerChanged()
        TestAssert(speakerName = "Game speakers" && IniRead(iniPath, "cfg", "speakerName") = speakerName,
            "Desktop device callback uses the same persistence")
        audBefore := FileRead(iniPath)
        AudioInputApplyDeviceResult("JRPG_AUDIO_DEVICES:OK`nHeadphones`nHeadphones`n日本語 ＆ 音声`n")
        TestAssert(ddlSpeaker.Text = "Game speakers" && InStr(gAudioInputStatus, "unavailable"),
            "Refresh retains a disconnected selection and reports it")
        TestAssert(ControlGetItems(ddlSpeaker.Hwnd).Length = 4 && FileRead(iniPath) = audBefore,
            "Device refresh deduplicates results without changing saved input")
        AudioInputApplyDeviceResult("Traceback: simulated failure")
        TestAssert(ControlGetItems(ddlSpeaker.Hwnd).Length = 4 && ddlSpeaker.Text = "Game speakers"
            && InStr(gAudioInputStatus, "previous list was kept"), "Bad refresh output preserves the old list")
        AudioInputApplyDeviceResult("JRPG_AUDIO_DEVICES:OK`n")
        TestAssert(ControlGetItems(ddlSpeaker.Hwnd).Length = 2 && ddlSpeaker.Text = "Game speakers",
            "Empty device list retains default and unavailable current input")

        CPBigBoxSetPage("quickAudio", false)
        for action in ["device", "refresh", "test"] {
            CPBigBoxAudioInputAction(action)
            TestAssert(!TestControlShown(CPBigBoxControls["audio_" action]) && !AudioInputJobBusy(),
                "Home Audio AI cannot activate full-page input controls: " action)
        }
        CPBigBoxOpenAIChoice("device")
        TestAssert(!CPBigBoxAIChoiceActive(), "Device chooser cannot be opened from Home Audio AI")
        CPBigBoxSetPage("audio", false)
        TestLiveAudioRunning := true
        TestAudioInput()
        TestAssert(!AudioInputJobBusy() && InStr(gAudioInputStatus, "Stop Audio Translation"),
            "Audio test refuses to interrupt a live session")
        TestLiveAudioRunning := false
        pythonExe := A_ScriptDir "\missing-python.exe"
        TestAudioInput()
        TestAssert(!AudioInputJobBusy() && InStr(gAudioInputStatus, "helper not found"), "Missing helper reports an inline error")

        ; Run the real hidden-process/timer/handle code with an AHK fixture,
        ; never Python, a real audio driver, personal settings, or an AI API.
        pythonExe := A_AhkPath
        audioScript := A_ScriptDir "\audio_input_fixture.ahk"
        EnvSet("AUDIO_DEVICE_LIST_RESULT_FILE", "previous-list-target")
        EnvSet("AUDIO_TEST_RESULT_FILE", "previous-test-target")
        EnvSet("JRPG_TEST_AUDIO_MODE", "ok")
        audBefore := FileRead(iniPath)
        TestAssert(AudioInputJobStart("devices"), "Device refresh starts a hidden asynchronous helper")
        audJob := gAudioInputJob
        TestAssert(audJob["handle"] != 0 && !ddlSpeaker.Enabled && !btnSpRef.Enabled && !btnAudioTest.Enabled,
            "Running diagnostic retains its process handle and disables competing controls")
        TestAssert(!AudioInputJobStart("test") && gAudioInputJob["pid"] = audJob["pid"], "Busy guard prevents a second diagnostic")
        TestAssert(InStr(CPBigBoxControls["audio_refresh"].Text, "Refreshing"), "Refresh has a visible busy label")
        TestAssert(EnvGet("AUDIO_DEVICE_LIST_RESULT_FILE") = "previous-list-target"
            && EnvGet("SPEAKER_NAME") = "running-session-input", "Helper launch restores inherited environment variables")
        CPBigBoxSetPage("explanation", false)
        TestWaitAudioJob()
        TestAssert(FileRead(iniPath) = audBefore && ddlSpeaker.Text = "Game speakers",
            "Refresh completes safely after navigating away, without saving a different input")
        TestAssert(!FileExist(audJob["path"]) && audJob["handle"] = 0 && ddlSpeaker.Enabled,
            "Completed job closes its handle, removes only its result and re-enables controls")
        CPBigBoxSetPage("audio", false)
        for mode in ["ok", "silent", "error", "missing"] {
            EnvSet("JRPG_TEST_AUDIO_MODE", mode)
            TestAssert(AudioInputJobStart("test"), "Fixture audio test starts: " mode)
            audJob := gAudioInputJob
            TestAssert(InStr(CPBigBoxControls["audio_test"].Text, "Listening"), "Audio test has a visible listening label")
            TestWaitAudioJob()
            expected := mode = "ok" ? "Audio detected" : mode = "silent" ? "No audio detected"
                : mode = "error" ? "Could not open" : "without a result"
            TestAssert(InStr(gAudioInputStatus, expected) && InStr(CPBigBoxControls["audio_status"].Text, expected),
                "Audio result is displayed inline: " mode)
            TestAssert(!FileExist(audJob["path"]) && audJob["handle"] = 0, "Test result and process handle are cleaned: " mode)
        }
        EnvSet("JRPG_TEST_AUDIO_MODE", "hang")
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["audio_test"]))
        CPControllerResetNavigation()
        CPControllerHandleNavigation(TestSnapshot(), CPBigBoxGui.Hwnd)
        CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
        Sleep(30)
        audJob := gAudioInputJob
        Loop 8
            CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
        TestAssert(AudioInputJobBusy() && gAudioInputJob["pid"] = audJob["pid"], "Held A cannot launch repeated audio tests")
        TestAssert(CPBigBoxFocusFrame["key"] = "ai_provider", "Busy input button hands focus to an enabled control")
        audJob["timeout"] := 1
        TestWaitAudioJob()
        TestAssert(InStr(gAudioInputStatus, "timed out") && !ProcessExist(audJob["pid"]),
            "A hung diagnostic times out and terminates only its retained child process")
        TestAssert(btnAudioTest.Enabled && !FileExist(audJob["path"]), "Timeout restores controls and clears the temporary result")
        TestAssert(AudioInputJobStart("devices"), "Another diagnostic can start after a timeout")
        audJob := gAudioInputJob
        AudioInputJobCancel()
        AudioInputJobControls()
        TestAssert(!AudioInputJobBusy() && !ProcessExist(audJob["pid"]) && audJob["handle"] = 0,
            "Shutdown cleanup stops and releases only the owned diagnostic")
        TestAssert(TestExternalCalls = 0, "Input settings/tests never invoke translation or overlay actions")
    } finally {
        AudioInputJobCancel()
        AudioInputJobControls()
        iniPath := audOriginalIni, pythonExe := audOriginalPython, audioScript := audOriginalScript
        TestLiveAudioRunning := false
        for name, value in audEnv
            EnvSet(name, value)
        CPBigBoxSetPage("home", false)
    }
}

TestWaitAudioJob() {
    started := A_TickCount
    while AudioInputJobBusy() && A_TickCount - started < 4000 {
        AudioInputJobPoll()
        Sleep(20)
    }
    TestAssert(!AudioInputJobBusy(), "Diagnostic fixture completes within test deadline")
}

TestAssert(condition, message) {
    global TestCount
    if !condition
        throw Error(message)
    TestCount += 1
    FileAppend("OK: " message "`n", "*", "UTF-8")
}

TestControlActionIndex(action) {
    global hotkeyActions
    for index, knownAction in hotkeyActions
        if knownAction = action
            return index
    throw Error("Unknown test action: " action)
}

TestControlCaptureSnapshot(token := "", connected := true) {
    return Map("connected", connected, "name", connected ? "XInput controller 1" : "",
        "tokens", token = "" ? Map() : Map(token, true))
}

TestCaptureControllerBinding(token) {
    global CPBigBoxControlCapture
    SetTimer(CPBigBoxControllerCaptureTick, 0)
    captureStart := CPBigBoxControlCapture["started"]
    CPBigBoxControllerCaptureApplySnapshot(TestControlCaptureSnapshot(), captureStart + 1)
    CPBigBoxControllerCaptureApplySnapshot(TestControlCaptureSnapshot(token), captureStart + 2)
    CPBigBoxControllerCaptureApplySnapshot(TestControlCaptureSnapshot(), captureStart + 3)
}

TestOpenControlAction(kind, action, parent := "controls") {
    global CPBigBoxAIChoice
    CPBigBoxSetPage(parent, false)
    CPBigBoxOpenControlActions(kind)
    index := TestControlActionIndex(action)
    CPBigBoxAIChoice["index"] := index
    CPBigBoxUpdateAIContent()
    CPBigBoxCommitAIChoiceIndex(index)
}

TestBigBoxControls() {
    global
    local dims, before, oldValue, realIni, actionOne, actionTwo, index, captureStart, rebindsBefore
    actionOne := "screenshot_translate", actionTwo := "explain_last_translation"
    for dims in [[1280, 720], [1920, 1080], [3840, 2160]] {
        CPBigBoxGui.Show("Hide w" dims[1] " h" dims[2])
        CPBigBoxDashboardResize(CPBigBoxGui, 0, dims[1], dims[2])
        before := FileRead(iniPath)
        CPBigBoxSetPage("controls", false)
        TestAssert(FileRead(iniPath) = before && CPBigBoxGroupedSettingsActive(),
            "Full Controls page opens without changing bindings")
        TestAssert(TestControlShown(CPBigBoxControls["ctrl_keyboard"])
            && TestControlShown(CPBigBoxControls["ctrl_controller"])
            && !TestControlShown(CPBigBoxControls["previewTitle"]),
            "Full Controls page replaces the migration preview")
        TestLayout(dims[1], dims[2])
        TestCapture("controls-settings-" dims[1] ".png", dims[1], dims[2])
        CPBigBoxSetPage("quickControls", false)
        TestAssert(!TestControlShown(CPBigBoxControls["ctrl_keyboard"])
            && TestControlShown(CPBigBoxControls["ctrl_controller"]),
            "Home Button Configuration remains controller-focused")
        TestLayout(dims[1], dims[2])
        TestCapture("controls-quick-" dims[1] ".png", dims[1], dims[2])
    }
    CPBigBoxGui.Show("Hide w1920 h1080")
    CPBigBoxDashboardResize(CPBigBoxGui, 0, 1920, 1080)

    CPBigBoxSetPage("controls", false)
    oldValue := CPControllerInputsEnabled
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ctrl_enabled"]))
    CPBigBoxToggleControllerOption("enabled")
    TestAssert(CPControllerInputsEnabled = !oldValue && cbControllerInputsEnabled.Value = !oldValue,
        "Direct action toggle synchronizes runtime and desktop checkbox")
    TestAssert(IniRead(iniPath, "controller_inputs", "enabled") = !oldValue
        && InStr(CPBigBoxControlNotice, "Saved"), "Direct action toggle persists exact existing INI key")
    CPBigBoxToggleControllerOption("enabled")
    TestAssert(CPControllerInputsEnabled = oldValue, "Direct action toggle restores its original state")
    oldValue := CPControllerDpadNavigationEnabled
    CPBigBoxToggleControllerOption("dpad")
    TestAssert(CPControllerDpadNavigationEnabled = !oldValue
        && cbControllerDpadNavigationEnabled.Value = !oldValue,
        "D-pad toggle synchronizes runtime and desktop checkbox")
    CPBigBoxToggleControllerOption("dpad")
    TestAssert(CPControllerDpadNavigationEnabled = oldValue
        && IniRead(iniPath, "controller_inputs", "dpad_navigation") = oldValue,
        "D-pad toggle restores and persists its original state")
    IniWrite(!oldValue, iniPath, "controller_inputs", "dpad_navigation")
    CPBigBoxToggleControllerOption("dpad")
    TestAssert(CPControllerDpadNavigationEnabled = oldValue && InStr(CPBigBoxControlNotice, "changed elsewhere"),
        "Controller option rejects an external settings change")
    IniWrite(oldValue, iniPath, "controller_inputs", "dpad_navigation")

    realIni := iniPath
    oldValue := CPControllerInputsEnabled
    iniPath := A_ScriptDir "\missing-controls-folder\settings.ini"
    CPBigBoxToggleControllerOption("enabled")
    TestAssert(CPControllerInputsEnabled = oldValue && InStr(CPBigBoxControlNotice, "Could not save"),
        "Failed controller option save leaves runtime state unchanged")
    iniPath := realIni
    CPBigBoxSetPage("home", false)

    CPBigBoxSetPage("controls", false)
    before := FileRead(iniPath)
    CPBigBoxOpenControlActions("controller")
    TestAssert(CPBigBoxAIListActive() && CPBigBoxAIChoice["options"].Length = 10,
        "Ten controller actions use the bounded scrolling list")
    TestAssert(InStr(CPBigBoxControls["modeTitle"].Text, "Choose an action")
        && InStr(CPBigBoxControls["listChoice3"].Text, "Capture + Translate"),
        "Controller list shows friendly action and binding labels")
    TestAssert(FileRead(iniPath) = before, "Browsing controller actions never writes settings")
    TestLayout(1920, 1080)
    CPBigBoxCommitAIChoiceIndex(TestControlActionIndex(actionOne))
    TestAssert(CPBigBoxCurrentPage = "controlDetail" && CPBigBoxControlState["action"] = actionOne,
        "Selecting a controller action opens its modern detail page")
    TestAssert(!TestControlShown(CPBigBoxControls["ctrl_default"])
        && CPBigBoxNavigationControls.Length = 3, "Controller detail exposes Assign, Disable and Back only")
    TestLayout(1920, 1080)
    TestCapture("controls-controller-detail-1920.png", 1920, 1080)
    CPBigBoxControlDetailAction("primary")
    SetTimer(CPBigBoxControllerCaptureTick, 0)
    TestAssert(CPBigBoxCurrentPage = "controlCapture" && CPBigBoxControlCapture["active"]
        && CPControllerCaptureActive, "Controller assignment enters an isolated fullscreen capture state")
    TestLayout(1920, 1080)
    TestCapture("controls-controller-capture-1920.png", 1920, 1080)
    CPBigBoxControllerCaptureApplySnapshot(TestControlCaptureSnapshot("", false))
    TestAssert(InStr(CPBigBoxControls["ctrl_status"].Text, "No compatible controller"),
        "Disconnected controller reports an actionable capture status")
    TestCaptureControllerBinding("X:X")
    TestAssert(CPBigBoxCurrentPage = "controlDetail" && !CPBigBoxControlCapture["active"]
        && !CPControllerCaptureActive, "Release saves a captured button and exits capture")
    TestAssert(CPControllerBindings[actionOne] = "X:X"
        && CPControllerBindingEdits[actionOne].Value = "X / Square"
        && IniRead(iniPath, "controller_inputs", actionOne) = "X:X",
        "Captured controller button synchronizes runtime, desktop UI and INI")
    CPBigBoxControlDetailAction("disable")
    TestAssert(CPControllerBindings[actionOne] = "" && IniRead(iniPath, "controller_inputs", actionOne) = "",
        "Controller binding can be disabled from its detail page")

    CPBigBoxControlDetailAction("primary")
    TestCaptureControllerBinding("X:X")
    CPBigBoxReturnToControlActions()
    TestAssert(CPBigBoxAIListActive() && CPBigBoxAIChoice["index"] = TestControlActionIndex(actionOne),
        "Returning from detail restores the selected action in the list")
    CPBigBoxCommitAIChoiceIndex(TestControlActionIndex(actionTwo))
    CPBigBoxControlDetailAction("primary")
    TestCaptureControllerBinding("X:X")
    TestAssert(CPBigBoxCurrentPage = "controlConflict"
        && CPBigBoxControlState["conflict"] = actionOne,
        "Duplicate controller button opens a controller-friendly conflict page")
    TestLayout(1920, 1080)
    TestCapture("controls-conflict-1920.png", 1920, 1080)
    TestAssert(CPControllerBindings[actionOne] = "X:X" && CPControllerBindings[actionTwo] = "",
        "Conflict prompt makes no binding change before confirmation")
    CPBigBoxResolveControlConflict(false)
    TestAssert(CPBigBoxCurrentPage = "controlDetail" && CPControllerBindings[actionOne] = "X:X"
        && CPControllerBindings[actionTwo] = "", "Keeping current controller bindings is non-destructive")
    CPBigBoxControlDetailAction("primary")
    TestCaptureControllerBinding("X:X")
    CPBigBoxResolveControlConflict(true)
    TestAssert(CPControllerBindings[actionOne] = "" && CPControllerBindings[actionTwo] = "X:X"
        && IniRead(iniPath, "controller_inputs", actionOne) = ""
        && IniRead(iniPath, "controller_inputs", actionTwo) = "X:X",
        "Move binding atomically clears the old action and saves the new action")

    before := CPControllerBindings[actionTwo]
    CPBigBoxControlDetailAction("primary")
    SetTimer(CPBigBoxControllerCaptureTick, 0)
    TestAssert(CPBigBoxControls["ctrl_cancel"].Text = "Cancel",
        "Capture restores its Cancel label after visiting the conflict page")
    captureStart := CPBigBoxControlCapture["started"]
    CPBigBoxControllerCaptureApplySnapshot(TestControlCaptureSnapshot(), captureStart + 1)
    CPBigBoxControllerCaptureApplySnapshot(TestControlCaptureSnapshot("X:B"), captureStart + 2)
    CPBigBoxControllerCaptureApplySnapshot(TestControlCaptureSnapshot("X:B"), captureStart + 1003)
    TestAssert(CPBigBoxCurrentPage = "controlDetail" && CPControllerBindings[actionTwo] = before,
        "Holding B for one second cancels capture without changing the binding")
    CPBigBoxControlDetailAction("primary")
    SetTimer(CPBigBoxControllerCaptureTick, 0)
    captureStart := CPBigBoxControlCapture["started"]
    CPBigBoxControllerCaptureApplySnapshot(TestControlCaptureSnapshot(), captureStart + 60000)
    TestAssert(CPBigBoxCurrentPage = "controlDetail" && CPControllerBindings[actionTwo] = before
        && InStr(CPBigBoxControlNotice, "timed out"), "Controller capture times out safely after one minute")
    IniWrite("X:Y", iniPath, "controller_inputs", actionTwo)
    CPBigBoxControlDetailAction("disable")
    TestAssert(CPControllerBindings[actionTwo] = before && IniRead(iniPath, "controller_inputs", actionTwo) = "X:Y"
        && InStr(CPBigBoxControlNotice, "changed elsewhere"), "Binding save rejects stale external state")
    IniWrite(before, iniPath, "controller_inputs", actionTwo)
    CPBigBoxBack()
    TestAssert(CPBigBoxAIListActive(), "B from binding detail returns one layer to the action list")
    CPBigBoxBack()
    TestAssert(CPBigBoxCurrentPage = "controls" && CPBigBoxFocusFrame["key"] = "ctrl_controller",
        "B from action list returns to the originating Controls tile")

    TestOpenControlAction("keyboard", "hide_show_translator")
    TestAssert(CPBigBoxCurrentPage = "controlDetail" && TestControlShown(CPBigBoxControls["ctrl_default"])
        && CPBigBoxNavigationControls.Length = 4, "Keyboard detail adds Restore default")
    CPBigBoxControlDetailAction("primary")
    TestAssert(CPBigBoxCurrentPage = "controlHotkey" && TestControlShown(CPBigBoxControls["ctrl_hotkey"]),
        "Keyboard shortcut opens an in-dashboard Hotkey editor")
    TestAssert(CPBigBoxControls["ctrl_cancel"].Text = "Cancel",
        "Hotkey editor restores its Cancel label after conflict workflows")
    TestLayout(1920, 1080)
    TestCapture("controls-hotkey-1920.png", 1920, 1080)
    CPBigBoxControls["ctrl_hotkey"].Value := "^!q"
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ctrl_hotkey"]))
    CPBigBoxDashboardActivate()
    TestAssert(CPBigBoxFocusFrame["key"] = "ctrl_save", "A/Enter from the Hotkey field advances to Save")
    rebindsBefore := TestRebinds
    CPBigBoxSaveControlHotkey()
    TestAssert(hkEdits["hide_show_translator"].Value = "^!q"
        && IniRead(iniPath, "hotkeys", "hide_show_translator") = "^!q",
        "Keyboard shortcut saves to desktop control and existing INI key")
    TestAssert(TestRebinds = rebindsBefore + 4 && FileExist(overlayDir "\hotkeys.reload"),
        "Keyboard change triggers the same live-reload path as Advanced Settings")
    CPBigBoxControlDetailAction("default")
    TestAssert(hkEdits["hide_show_translator"].Value = hotkeyDefaults["hide_show_translator"],
        "Restore default uses the existing shortcut default")
    CPBigBoxControlDetailAction("primary")
    CPBigBoxControls["ctrl_hotkey"].Value := hotkeyDefaults["explain_last_translation"]
    CPBigBoxSaveControlHotkey()
    TestAssert(CPBigBoxCurrentPage = "controlConflict"
        && hkEdits["explain_last_translation"].Value = hotkeyDefaults["explain_last_translation"],
        "Duplicate keyboard shortcut also waits for explicit conflict confirmation")
    CPBigBoxResolveControlConflict(true)
    TestAssert(CPBigBoxControlComparableBinding("keyboard", hkEdits["hide_show_translator"].Value)
        = CPBigBoxControlComparableBinding("keyboard", hotkeyDefaults["explain_last_translation"])
        && hkEdits["explain_last_translation"].Value = "",
        "Keyboard conflict can move the shortcut without duplicate bindings")
    CPBigBoxControlDetailAction("disable")
    TestAssert(hkEdits["hide_show_translator"].Value = "", "Keyboard binding can be disabled")
    CPBigBoxReturnToControlActions()
    CPBigBoxBack()
    TestAssert(CPBigBoxCurrentPage = "controls", "Keyboard action flow returns cleanly to full Controls")

    TestOpenControlAction("controller", actionTwo, "quickControls")
    CPBigBoxControlDetailAction("primary")
    SetTimer(CPBigBoxControllerCaptureTick, 0)
    before := FileRead(iniPath)
    CPBigBoxDashboardHide(false)
    TestAssert(!CPBigBoxControlCapture["active"] && !CPControllerCaptureActive
        && FileRead(iniPath) = before && CPBigBoxCurrentPage = "quickControls",
        "Hiding the dashboard cancels pending controller capture without writes")
    CPBigBoxGui.Show("Hide w1920 h1080")
    CPBigBoxDashboardRelayoutAndUpdate()
    CPBigBoxApplyPageVisibility()
    TestAssert(TestControlShown(CPBigBoxControls["ctrl_controller"])
        && !TestControlShown(CPBigBoxControls["ctrl_cancel"]),
        "Reopening after capture cancellation restores the parent controls, not stale capture UI")
    CPBigBoxSetPage("home", false)
}

TestBigBoxStage3D() {
    global
    local before, startCount, stopCount, stageSize, stageIni, stageOverlay, sentBefore, resumed, captureKind, targetBefore
    CPBigBoxSetPage("home", false)
    TestLiveAudioRunning := false
    startCount := TestAudioStarts, stopCount := TestAudioStops
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["audioToggle"]))
    CPControllerResetNavigation()
    CPControllerHandleNavigation(TestSnapshot(), CPBigBoxGui.Hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
    Sleep(30)
    Loop 12
        CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
    TestAssert(TestAudioStarts = startCount + 1 && TestAudioStops = stopCount && TestLiveAudioRunning,
        "Held Home audio action starts exactly once")
    TestAssert(CPBigBoxCurrentPage = "home" && CPBigBoxFocusFrame["key"] = "audioToggle"
        && InStr(CPBigBoxControls["audioToggle"].Text, "On"), "Home retains focus and shows actual audio state")
    CPControllerHandleNavigation(TestSnapshot(), CPBigBoxGui.Hwnd)
    CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
    Sleep(30)
    TestAssert(TestAudioStops = stopCount + 1 && !TestLiveAudioRunning, "New A press stops audio")
    TestAudioStartOK := false
    CPBigBoxToggleAudio()
    TestAssert(!TestLiveAudioRunning && InStr(CPBigBoxActionNotice, "startup error")
        && CPBigBoxControls["audioToggle"].Enabled, "Audio startup failure is inline and restores the action")
    TestAudioStartOK := true
    startCount := TestAudioStarts
    gAudioInputJob := Map("active", true, "mode", "test")
    CPBigBoxToggleAudio()
    gAudioInputJob := Map("active", false)
    TestAssert(TestAudioStarts = startCount && InStr(CPBigBoxActionNotice, "input check"), "Audio cannot start during an input test")
    CPBigBoxModalDepth := 1
    CPBigBoxToggleAudio()
    CPBigBoxModalDepth := 0
    TestAssert(TestAudioStarts = startCount, "Modal guard also blocks the Home audio action")
    CPBigBoxSetPage("audio", false)
    CPBigBoxToggleAudio()
    TestAssert(TestLiveAudioRunning && CPBigBoxFocusFrame["key"] = "audio_power", "Full Audio page uses the same action")
    CPBigBoxToggleAudio()
    TestLiveAudioRunning := true
    CPBigBoxUpdateAudioPower()
    TestAssert(InStr(CPBigBoxActionNotice, "is on"), "External audio start refreshes the status notice")
    TestLiveAudioRunning := false
    CPBigBoxUpdateAudioPower()
    TestAssert(InStr(CPBigBoxActionNotice, "is off"), "Unexpected audio exit cannot leave a stale On notice")

    CPSetCaptureMaxKB(1400)
    IniWrite("region", iniPath, "capture", "mode")
    IniWrite("10,20,640,480", iniPath, "capture", "rect")
    IniWrite("Synthetic game", iniPath, "capture", "winTitle")
    for stageSize in [[1280, 720], [1920, 1080], [3840, 2160]] {
        CPBigBoxGui.Show("Hide w" stageSize[1] " h" stageSize[2])
        CPBigBoxDashboardResize(CPBigBoxGui, 0, stageSize[1], stageSize[2])
        CPBigBoxSetPage("home", false)
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["capture"]))
        CPBigBoxOpenCaptureSettings()
        TestAssert(CPBigBoxCaptureParent = "home" && CPBigBoxNavigationControls.Length = 6,
            "Capture opens as a modern Home quick view")
        TestLayout(stageSize[1], stageSize[2])
        TestCapture("capture-" stageSize[1] ".png", stageSize[1], stageSize[2])
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["cap_max"]))
        before := FileRead(iniPath)
        CPBigBoxOpenCaptureLimit()
        TestAssert(CPBigBoxNavigationControls.Length = 6 && !CPBigBoxControls["advanced"].Enabled,
            "PNG editor has four adjustments plus explicit Save/Cancel")
        CPBigBoxAdjustCaptureLimit(500)
        TestLayout(stageSize[1], stageSize[2])
        TestCapture("capture-limit-" stageSize[1] ".png", stageSize[1], stageSize[2])
        CPBigBoxSwitchPage(1)
        TestAssert(CPBigBoxCurrentPage = "captureLimit", "Shoulders do not leave a pending PNG adjustment")
        CPBigBoxBack()
        TestAssert(CPBigBoxCurrentPage = "quickCapture" && CPBigBoxFocusFrame["key"] = "cap_max"
            && FileRead(iniPath) = before && capMaxKB = 1400, "Cancel PNG adjustment keeps settings and restores its tile")
        CPBigBoxBack()
        TestAssert(CPBigBoxCurrentPage = "home" && CPBigBoxFocusFrame["key"] = "capture", "Capture returns to its Home tile")
    }
    CPBigBoxSetPage("screenshot", false)
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["capture_settings"]))
    CPBigBoxOpenCaptureSettings()
    TestAssert(CPBigBoxCaptureParent = "screenshot" && InStr(CPBigBoxControls["pageHint"].Text, "Game Text Translation"),
        "Full Screenshot page shares Capture with its own return breadcrumb")
    CPBigBoxOpenCaptureLimit()
    CPBigBoxAdjustCaptureLimit(-20000)
    TestAssert(CPBigBoxCaptureLimit["value"] = 100, "PNG value clamps at 100 KB")
    CPBigBoxAdjustCaptureLimit(20000)
    TestAssert(CPBigBoxCaptureLimit["value"] = 10000, "PNG value clamps at 10000 KB")
    CPBigBoxAdjustCaptureLimit(-8500)
    CPBigBoxSaveCaptureLimit()
    TestAssert(capMaxKB = 1500 && eCapMax.Value = "1500" && IniRead(iniPath, "capture", "maxKB") = 1500,
        "PNG Save updates the shared key, runtime value and desktop field")
    before := FileRead(iniPath)
    CPBigBoxOpenCaptureLimit()
    CPBigBoxSaveCaptureLimit()
    TestAssert(FileRead(iniPath) = before, "Saving unchanged PNG value does not rewrite settings")
    CPBigBoxOpenCaptureLimit()
    CPBigBoxAdjustCaptureLimit(100)
    IniWrite(1700, iniPath, "capture", "maxKB")
    CPBigBoxSaveCaptureLimit()
    TestAssert(InStr(CPBigBoxCaptureNotice, "Settings changed") && IniRead(iniPath, "capture", "maxKB") = 1700,
        "Stale PNG adjustment cannot overwrite an external change")
    CPSetCaptureMaxKB(1500)
    CPBigBoxOpenCaptureLimit()
    CPBigBoxAdjustCaptureLimit(100)
    stageIni := iniPath
    iniPath := A_ScriptDir "\missing-parent\cannot-save-png.ini"
    CPBigBoxCaptureLimit["profile"] := ""
    CPBigBoxSaveCaptureLimit()
    iniPath := stageIni
    TestAssert(capMaxKB = 1500 && eCapMax.Value = "1500" && InStr(CPBigBoxCaptureNotice, "Could not save"),
        "Failed PNG write retains previous runtime and desktop values")
    CPBigBoxBack()
    TestAssert(CPBigBoxCurrentPage = "screenshot" && CPBigBoxFocusFrame["key"] = "capture_settings",
        "Capture opened from full settings returns to that tile, not Home")
    CPBigBoxOpenCaptureSettings()
    CPBigBoxOpenCaptureLimit()
    CPBigBoxAdjustCaptureLimit(100)
    CPBigBoxDashboardHide(false)
    TestAssert(!CPBigBoxCaptureLimit["active"] && capMaxKB = 1500, "Hiding the dashboard cancels pending PNG changes")

    stageOverlay := overlayAhk
    try {
        overlayAhk := A_ScriptDir "\missing-overlay.ahk"
        sentBefore := TestCaptureSends.Length
        CPBigBoxBeginCapture("region")
        TestAssert(!CPBigBoxCaptureWatch["active"] && TestCaptureSends.Length = sentBefore
            && InStr(CPBigBoxCaptureNotice, "not found"), "Missing overlay reports inline without hiding or starting capture")
        overlayAhk := A_AhkPath ; An existing test-only file; the launch is stubbed.
        TestTranslatorReady := false
        CPBigBoxBeginCapture("region")
        TestAssert(!CPBigBoxCaptureWatch["active"] && InStr(CPBigBoxCaptureNotice, "Could not open"), "Overlay launch failure is inline")
        TestTranslatorReady := true
        TestCaptureSendOK := false
        resumed := TestCaptureResumes
        CPBigBoxBeginCapture("window")
        TestAssert(!CPBigBoxCaptureWatch["active"] && TestCaptureResumes = resumed + 1
            && InStr(CPBigBoxCaptureNotice, "did not respond"), "Capture send failure restores dashboard and stops its watcher")
        TestCaptureSendOK := true
        for captureKind in ["region", "window"] {
            CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["cap_" captureKind]))
            CPControllerResetNavigation()
            CPControllerHandleNavigation(TestSnapshot(), CPBigBoxGui.Hwnd)
            CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
            Sleep(30)
            sentBefore := TestCaptureSends.Length
            TestAssert(CPBigBoxCaptureWatch["active"] && InStr(TestCaptureSends[-1], "kind=" captureKind)
                && InStr(TestCaptureSends[-1], "regionpreset=1"), "Controller capture uses the existing picker protocol: " captureKind)
            CPBigBoxBeginCapture(captureKind)
            TestAssert(TestCaptureSends.Length = sentBefore, "Only one capture request can be outstanding")
            IniWrite("unrelated", iniPath, "test", "duringCapture")
            CPBigBoxWatchCapture()
            TestAssert(CPBigBoxCaptureWatch["active"], "Unrelated INI writes do not end capture")
            targetBefore := IniRead(iniPath, "capture", "rect")
            IniWrite("canceled", iniPath, "capture", "pickStatus")
            IniWrite("test-" captureKind, iniPath, "capture", "pickSeq")
            CPBigBoxWatchCapture()
            Loop 12
                CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
            TestAssert(!CPBigBoxCaptureWatch["active"] && InStr(CPBigBoxCaptureNotice, "cancelled")
                && IniRead(iniPath, "capture", "rect") = targetBefore && TestCaptureSends.Length = sentBefore,
                "Cancel returns without changing target or relaunching from held A")
        }
        CPBigBoxBeginCapture("region")
        IniWrite("selected", iniPath, "capture", "pickStatus")
        IniWrite("test-selected", iniPath, "capture", "pickSeq")
        CPBigBoxWatchCapture()
        TestAssert(!CPBigBoxCaptureWatch["active"] && InStr(CPBigBoxCaptureNotice, "updated"),
            "Reselecting the same rectangle still completes via sequence marker")
        CPBigBoxBeginCapture("region")
        CPBigBoxCaptureWatch["started"] := DllCall("kernel32\GetTickCount64", "uint64") - 125001
        CPBigBoxWatchCapture()
        TestAssert(!CPBigBoxCaptureWatch["active"] && InStr(CPBigBoxCaptureNotice, "No completion"), "Lost completion has a bounded return path")
        CPBigBoxBeginCapture("window")
        resumed := TestCaptureResumes
        CPBigBoxCancelCaptureWatch()
        CPBigBoxWatchCapture()
        TestAssert(!CPBigBoxCaptureWatch["active"] && TestCaptureResumes = resumed, "Shutdown cancels watcher without reopening a GUI")
    } finally {
        CPBigBoxCancelCaptureWatch()
        overlayAhk := stageOverlay
        TestTranslatorReady := true, TestCaptureSendOK := true
    }
    CPBigBoxSetPage("home", false)
}

TestBigBoxAudioFeedback() {
    global CPBigBoxGui, CPBigBoxControls, CPBigBoxCurrentPage, CPBigBoxFocusFrame, controlDarkMode
    global TestAudioBusyProbe, TestAudioStarts, TestAudioStops, TestLiveAudioRunning, TestAudioStartOK
    savedTheme := controlDarkMode
    try {
        for theme in [0, 1] {
            controlDarkMode := theme
            CPBigBoxDashboardApplyTheme()
            for page in ["home", "audio"] {
                CPBigBoxSetPage(page, false)
                key := page = "home" ? "audioToggle" : "audio_power"
                CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls[key]))
                TestLiveAudioRunning := false, TestAudioStartOK := true
                for action in ["start", "stop", "failed start"] {
                    TestAudioStartOK := action != "failed start"
                    startsBefore := TestAudioStarts, stopsBefore := TestAudioStops
                    TestAudioBusyProbe := Map("active", true, "samples", [])
                    CPControllerResetNavigation()
                    CPControllerHandleNavigation(TestSnapshot(), CPBigBoxGui.Hwnd)
                    CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
                    Sleep(40)
                    Loop 8
                        CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
                    TestAudioBusyProbe["active"] := false
                    prefix := page " " action " theme=" theme ": "
                    TestAssert(TestAudioBusyProbe["samples"].Length = 1, prefix "one in-flight observation")
                    sample := TestAudioBusyProbe["samples"][1]
                    TestAssert(sample["focus"] = CPBigBoxControls[key].Hwnd && sample["frame"] = key,
                        prefix "native focus and blue frame stay on audio during the action")
                    TestAssert(sample["enabled"] && InStr(sample["text"], "Please wait") && !sample["navigation"],
                        prefix "busy tile keeps focus without accepting another action")
                    TestAssert(CPBigBoxCurrentPage = page && CPBigBoxFocusFrame["key"] = key,
                        prefix "completion preserves the originating tile")
                    TestAssert(TestAudioStarts = startsBefore + (action != "stop")
                        && TestAudioStops = stopsBefore + (action = "stop"), prefix "held A does not repeat")
                }
            }
        }
    } finally {
        TestAudioBusyProbe["active"] := false
        TestLiveAudioRunning := false, TestAudioStartOK := true
        controlDarkMode := savedTheme
        CPBigBoxDashboardApplyTheme()
        CPBigBoxSetPage("home", false)
    }
}

TestObserveAudioBusy() {
    global TestAudioBusyProbe, CPBigBoxGui, CPBigBoxControls, CPBigBoxCurrentPage, CPBigBoxFocusFrame
    if !TestAudioBusyProbe["active"]
        return
    ; Pump queued native focus notifications while the start/stop is in flight.
    Sleep(20)
    key := CPBigBoxCurrentPage = "home" ? "audioToggle" : "audio_power"
    focused := CPBigBoxGui.FocusedCtrl
    TestAudioBusyProbe["samples"].Push(Map(
        "focus", IsObject(focused) ? focused.Hwnd : 0, "frame", CPBigBoxFocusFrame["key"],
        "enabled", CPBigBoxControls[key].Enabled, "text", CPBigBoxControls[key].Text,
        "navigation", CPBigBoxPageNavigationAllowed()))
}

TestBigBoxAudioStartup() {
    global pythonExe, audioScript, gPidAudio, TestAudioApiConfigured, TestAudioRecoveryScans, CPBigBoxActionNotice
    global TestToasts
    savedPython := pythonExe, savedScript := audioScript
    savedEnvironment := Map()
    for key in ["AUDIO_PROVIDER", "TEXT_PROVIDER", "TRANSLATE_MODEL", "GEMINI_AUDIO_MODEL", "TARGET_LANGUAGE_CODE",
        "TARGET_LANGUAGE_NAME", "SETTINGS_DIR", "JRPG_DEBUG", "PYTHONIOENCODING", "SPEAKER_NAME",
        "JRPG_TEST_AUDIO_RUNTIME_LOG", "JRPG_TEST_AUDIO_RUNTIME_MODE"]
        savedEnvironment[key] := EnvGet(key)
    runtimeLog := A_ScriptDir "\audio-runtime-starts.txt"
    try FileDelete(runtimeLog)
    childHandle := 0
    toastsBefore := TestToasts.Length
    try {
        pythonExe := A_ScriptDir "\missing-python.exe"
        TestAssert(!StartAudioCore(true) && InStr(CPBigBoxActionNotice, "not found"), "Production audio startup validates paths inline")
        pythonExe := A_AhkPath, audioScript := A_ScriptDir "\audio_runtime_fixture.ahk"
        TestAudioApiConfigured := false
        TestAssert(!StartAudioCore(true) && InStr(CPBigBoxActionNotice, "API key") && !gPidAudio,
            "Missing key prevents a process start and is reported inside Big Box")
        TestAssert(TestToasts.Length = toastsBefore, "Invalid audio configuration does not announce a successful start")
        TestAudioApiConfigured := true
        EnvSet("JRPG_TEST_AUDIO_RUNTIME_LOG", runtimeLog)
        EnvSet("JRPG_TEST_AUDIO_RUNTIME_MODE", "wait")
        TestAssert(StartAudioCore(true) && gPidAudio && ProcessExist(gPidAudio), "Production audio start receives the exact child PID")
        childHandle := DllCall("kernel32\OpenProcess", "uint", 0x100001, "int", false, "uint", gPidAudio, "ptr")
        TestAssert(TestToasts.Length = toastsBefore && InStr(CPBigBoxActionNotice, "is on"),
            "Big Box audio start reports success inside the dashboard without a corner toast")
        TestAssert(StartAudioCore(true) && TestToasts.Length = toastsBefore,
            "Already-running audio keeps feedback inside the dashboard")
        TestAssert(childHandle && TestAudioRecoveryScans = 0 && FileRead(runtimeLog, "UTF-8") = "started`n",
            "Successful Big Box start does not scan or launch a second process")
        DllCall("kernel32\TerminateProcess", "ptr", childHandle, "uint", 0)
        DllCall("kernel32\WaitForSingleObject", "ptr", childHandle, "uint", 1000)
        DllCall("kernel32\CloseHandle", "ptr", childHandle)
        childHandle := 0, gPidAudio := 0
        EnvSet("JRPG_TEST_AUDIO_RUNTIME_MODE", "exit")
        earlyStartResult := StartAudioCore(true)
        TestAssert(!earlyStartResult && InStr(CPBigBoxActionNotice, "exited during startup"),
            "An early audio-helper exit produces an inline failure (result=" earlyStartResult
                ", pid=" gPidAudio ", notice=" CPBigBoxActionNotice ")")
        TestAssert(TestToasts.Length = toastsBefore, "Early audio-helper failure does not show an On notification")
        TestAssert(FileRead(runtimeLog, "UTF-8") = "started`nstarted`n" && TestAudioRecoveryScans = 0,
            "Failed Big Box start is not retried synchronously or through process scans")
        TestAssert(!StartAudioCore(false), "Desktop audio startup reports an early helper exit")
        TestAssert(FileRead(runtimeLog, "UTF-8") = "started`nstarted`nstarted`n" && TestAudioRecoveryScans = 1,
            "Failed desktop startup performs one recovery scan but never launches a blocking retry")
    } finally {
        if childHandle {
            DllCall("kernel32\TerminateProcess", "ptr", childHandle, "uint", 0)
            DllCall("kernel32\CloseHandle", "ptr", childHandle)
        }
        gPidAudio := 0, pythonExe := savedPython, audioScript := savedScript, TestAudioApiConfigured := true
        for key, value in savedEnvironment
            EnvSet(key, value)
    }
}

TestBigBoxOverlays() {
    global
    local dims, page, target, fields, field, pref, before, sendsBefore, value, expected, oldProfile
    local realIni, original, failed, returnsBefore, oldHidden, x, y, w, h, theme, busyBefore
    ddlFont.Add(["Arial", "Consolas", "Georgia", "Verdana", "Tahoma"])
    TestOverlayStatusStates()
    for dims in [[1280, 720], [1920, 1080], [3840, 2160]] {
        CPBigBoxGui.Show("Hide w" dims[1] " h" dims[2])
        CPBigBoxDashboardResize(CPBigBoxGui, 0, dims[1], dims[2])
        for page in ["translationWindow", "explanationWindow", "quickOverlays", "quickTranslatorWindow", "quickExplainerWindow"] {
            before := FileRead(iniPath)
            CPBigBoxSetPage(page, false)
            TestAssert(FileRead(iniPath) = before && !TestControlShown(CPBigBoxControls["previewTitle"]),
                "Overlay page replaces preview without writes: " page)
            TestLayout(dims[1], dims[2])
            TestCapture(page "-" dims[1] ".png", dims[1], dims[2])
            if CPBigBoxOverlayTarget() = ""
                continue
            for field in ["bg", "opacity"] {
                pref := CPOverlayPreference(CPBigBoxOverlayTarget(), field)
                CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_" field]))
                sendsBefore := TestOverlayThemes.Length
                CPBigBoxOverlayAction(field)
                TestAssert(CPBigBoxCurrentPage = "overlayEdit" && !CPBigBoxControls["pageNext"].Enabled,
                    "Editor is outside shoulder navigation")
                TestLayout(dims[1], dims[2])
                if page = "translationWindow"
                    TestCapture("overlay-" field "-" dims[1] ".png", dims[1], dims[2])
                TestAssert(CPBigBoxOverlayEdit["value"] == pref["value"], "Opening editor retains exact original value")
                CPBigBoxDashboardMoveFocus("Left")
                CPBigBoxSwitchPage(1)
                TestAssert(CPBigBoxCurrentPage = "overlayEdit" && FileRead(iniPath) = before
                    && TestOverlayThemes.Length = sendsBefore, "Pending adjustments never write or send a live theme")
                CPBigBoxBack()
                TestAssert(CPBigBoxCurrentPage = page && CPBigBoxFocusFrame["key"] = "ov_" field
                    && CPOverlayPreference(CPBigBoxOverlayTarget(), field)["value"] == pref["value"],
                    "Cancel restores exact parent/tile and original setting")
                TestAssert(TestOverlayGradientKeys.Count = 0, "Color editor releases gradient registrations on close")
            }
        }
    }
    for page in ["translationWindow", "explanationWindow"] {
        CPBigBoxSetPage(page, false)
        target := CPBigBoxOverlayTarget()
        fields := target = "Translator" ? ["bg", "txt", "name", "opacity", "size", "bold", "font"]
            : ["bg", "txt", "opacity", "size", "bold", "font"]
        for field in fields {
            pref := CPOverlayPreference(target, field)
            original := pref["value"]
            value := pref["type"] = "color" ? "123ABC" : field = "opacity" ? 128
                : field = "size" ? 24 : field = "bold" ? !original : "Consolas"
            sendsBefore := TestOverlayThemes.Length
            CPSetOverlayPreference(target, field, value)
            TestAssert(CPOverlayPreference(target, field)["value"] == value
                && IniRead(iniPath, pref["section"], pref["key"]) == value, "Single overlay setting persists to correct section: " target " " field)
            if pref["type"] != "color"
                TestAssert((field = "font" ? pref["control"].Text : pref["control"].Value) == value,
                    "Overlay setting synchronizes desktop control: " field)
            TestAssert(TestOverlayThemes.Length = sendsBefore + 1 && TestOverlayThemes[-1][1] = target
                && TestOverlayThemes[-1][2] = 2500, "Only selected overlay receives a bounded theme update")
            CPSetOverlayPreference(target, field, original)
        }
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_font"]))
        before := FileRead(iniPath)
        CPBigBoxOverlayAction("font")
        TestAssert(CPBigBoxAIListActive() && CPBigBoxAIChoice["field"] = "overlayFont", "Fonts reuse the scrollable choice list")
        CPBigBoxAIListBoundary(true)
        CPBigBoxBack()
        TestAssert(FileRead(iniPath) = before && CPBigBoxFocusFrame["key"] = "ov_font", "Cancel font selection returns to font without saving")
        CPBigBoxOverlayAction("font")
        CPBigBoxCommitAIChoiceIndex(3)
        TestAssert(CPOverlayPreference(target, "font")["value"] = "Consolas" && CPBigBoxFocusFrame["key"] = "ov_font", "Font confirmation saves and restores its focus")
        CPBigBoxOverlayAction("font")
        oldProfile := IniRead(iniPath, "game_profiles", "active")
        IniWrite("Changed profile", iniPath, "game_profiles", "active")
        CPBigBoxCommitAIChoiceIndex(2)
        TestAssert(CPOverlayPreference(target, "font")["value"] = "Consolas" && InStr(CPBigBoxOverlayNotice, "changed"), "Font rejects a stale profile")
        IniWrite(oldProfile, iniPath, "game_profiles", "active")
        for field in ["bg", "opacity", "size"] {
            pref := CPOverlayPreference(target, field)
            CPBigBoxOverlayAction(field)
            CPBigBoxDashboardMoveFocus("Right")
            expected := CPBigBoxOverlayEdit["value"]
            CPBigBoxSaveOverlayEditor()
            TestAssert(CPBigBoxCurrentPage = page && CPOverlayPreference(target, field)["value"] == expected,
                "Editor Save applies its pending " target " " field)
            CPSetOverlayPreference(target, field, pref["value"])
        }
        CPBigBoxOverlayAction("size")
        CPBigBoxControls["ov_slider1"].Value := 200
        CPBigBoxOverlaySliderChanged()
        TestAssert(CPBigBoxOverlayEdit["value"] = (target = "Translator" ? 128 : 200), "Font size slider enforces overlay-specific maximum")
        CPBigBoxBack()
        original := CPOverlayPreference(target, "bold")["value"]
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_bold"]))
        CPControllerResetNavigation()
        CPControllerHandleNavigation(TestSnapshot(), CPBigBoxGui.Hwnd)
        CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
        Sleep(30)
        Loop 12
            CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
        TestAssert(CPOverlayPreference(target, "bold")["value"] = !original, "Held A toggles bold once: " target)
        CPSetOverlayPreference(target, "bold", original)
    }
    CPBigBoxSetPage("translationWindow", false)
    CPBigBoxOverlayAction("bg")
    CPBigBoxDashboardMoveFocus("Right")
    original := boxBgHex
    IniWrite("ABCDEF", iniPath, "cfg", "boxBg")
    CPBigBoxSaveOverlayEditor()
    TestAssert(CPBigBoxOverlayEdit["active"] && boxBgHex = original && IniRead(iniPath, "cfg", "boxBg") = "ABCDEF",
        "Editor rejects external changes without overwriting them")
    CPBigBoxBack()
    IniWrite(original, iniPath, "cfg", "boxBg")
    realIni := iniPath
    try {
        iniPath := A_ScriptDir "\missing-folder\settings.ini"
        CPBigBoxOverlayAction("size")
        original := fontSize
        CPBigBoxDashboardMoveFocus("Right")
        CPBigBoxSaveOverlayEditor()
        TestAssert(CPBigBoxOverlayEdit["active"] && fontSize = original && edFSize.Value = original,
            "Failed save preserves native and runtime font size and leaves Cancel available")
        CPBigBoxBack()
    } finally {
        iniPath := realIni
    }
    CPBigBoxOverlayAction("opacity")
    CPBigBoxDashboardMoveFocus("Left")
    before := FileRead(iniPath)
    CPBigBoxDashboardHide(false)
    TestAssert(!CPBigBoxOverlayEdit["active"] && FileRead(iniPath) = before, "Hiding dashboard cancels a pending overlay edit")
    for theme in [0, 1] {
        controlDarkMode := theme
        CPBigBoxDashboardApplyTheme()
        CPBigBoxOverlayAction("txt")
        TestAssert(CPBigBoxOverlaySliderFocused() = "ov_slider1", "Color editor begins on first slider")
        CPBigBoxDashboardMoveFocus("Down")
        TestAssert(CPBigBoxFocusFrame["key"] = "ov_slider2", "D-pad reaches saturation")
        CPBigBoxControls["ov_slider3"].Focus()
        CPBigBoxOverlaySliderMouse(0, 0, 0x201, CPBigBoxControls["ov_slider3"].Hwnd)
        Sleep(30)
        TestAssert(CPBigBoxFocusFrame["key"] = "ov_slider3", "Mouse-style slider focus updates the same blue outline")
        CPBigBoxDashboardMoveFocus("Up")
        CPBigBoxDashboardActivate()
        TestAssert(CPBigBoxFocusFrame["key"] = "ov_slider3", "A advances to brightness without saving")
        CPBigBoxDashboardMoveFocus("Down")
        TestAssert(CPBigBoxFocusFrame["key"] = "ov_save", "D-pad reaches explicit Save")
        CPBigBoxBack()
    }
    TestOverlayReady := false
    CPBigBoxOverlayAction("position")
    TestAssert(!CPBigBoxOverlayPosition["active"] && InStr(CPBigBoxOverlayNotice, "could not"), "Missing overlay reports inline without stranding dashboard")
    TestOverlayReady := true
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
    TestOverlayGui.Show("Hide x100 y100 w400 h200")
    WinGetPos(&x, &y, &w, &h, "ahk_id " TestOverlayGui.Hwnd)
    original := Map("x", x, "y", y, "w", w, "h", h)
        for page in ["translationWindow", "explanationWindow", "quickTranslatorWindow", "quickExplainerWindow"] {
            CPBigBoxSetPage(page, false)
            CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_position"]))
            returnsBefore := TestOverlayReturns
            CPBigBoxOverlayAction("position")
            TestAssert(CPOverlayAdjustState["active"] && CPOverlayAdjustState["bigBox"]
                && CPBigBoxOverlayPosition["active"], "Move/resize uses production workflow with Big Box return context")
            CPBigBoxSwitchPage(1)
            TestAssert(CPBigBoxCurrentPage = page, "Underlying dashboard cannot navigate while moving overlay")
            WinMove(130, 140, 500, 250, "ahk_id " TestOverlayGui.Hwnd)
            CPFinishOverlayAdjustment(false)
            Sleep(150)
            WinGetPos(&x, &y, &w, &h, "ahk_id " TestOverlayGui.Hwnd)
            TestAssert(x = original["x"] && y = original["y"] && w = original["w"] && h = original["h"],
                "Cancelled native move restores original bounds")
            TestAssert(!CPBigBoxOverlayPosition["active"] && TestOverlayReturns = returnsBefore + 1
                && CPBigBoxFocusFrame["key"] = "ov_position", "Adjustment returns to its original modern tile, not desktop")
            CPControllerHandleNavigation(TestSnapshot("X:A"), CPBigBoxGui.Hwnd)
            TestAssert(!CPOverlayAdjustState["active"], "Held confirmation on return cannot restart move/resize")
        }
        CPBigBoxOverlayAction("position")
        WinMove(160, 170, 420, 230, "ahk_id " TestOverlayGui.Hwnd)
        returnsBefore := TestOverlayReturns
        CPFinishOverlayAdjustment(true)
        Sleep(150)
        WinGetPos(&x, &y, &w, &h, "ahk_id " TestOverlayGui.Hwnd)
        TestAssert(x = 160 && y = 170 && w = 420 && h = 230 && TestOverlayReturns = returnsBefore + 1,
            "Confirmed native move keeps its new bounds and returns once")
        CPBigBoxOverlayAction("position")
        returnsBefore := TestOverlayReturns
        CPFinishOverlayAdjustment(false, true)
        Sleep(150)
        TestAssert(TestOverlayReturns = returnsBefore, "Quiet shutdown cleanup does not reopen a dashboard")
        CPBigBoxOverlayPosition := Map("active", false)
    } finally {
        DetectHiddenWindows oldHidden
        if CPOverlayAdjustState["active"]
            CPFinishOverlayAdjustment(false, true)
        TestOverlayGui.Destroy()
    }
    CPBigBoxSetPage("quickOverlays", false)
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_explainer"]))
    CPBigBoxSetPage("quickExplainerWindow", false)
    CPBigBoxBack()
    TestAssert(CPBigBoxCurrentPage = "quickOverlays" && CPBigBoxFocusFrame["key"] = "ov_explainer", "Quick overlay Back returns to the selected overlay")
    CPBigBoxBack()
    controlDarkMode := 1
}

TestBigBoxBackgroundOpacity() {
    global
    local before := FileRead(iniPath), sends := TestOverlayThemes.Length
    local alpha := Buffer(1), flags := Buffer(4), color := Buffer(4)
    local backdropHwnd, dims, bx, by, bw, bh, fx, fy, fw, fh, newPage
    ; Like the rest of this harness, navigate while hidden: CI processes may
    ; not own the interactive desktop. Show without activation for native
    ; layer/geometry checks, never steal foreground from the user's apps.
    CPBigBoxGui.Show("Hide w1280 h720")
    TestAssert(CPBigBoxBackgroundOpacity = 100 && CPBigBoxEffectiveBackgroundOpacity() = 100,
        "Fullscreen opacity defaults to solid and remains independent of desktop opacity")
    CPBigBoxSetPage("quickOverlays", false)
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_main"]))
    CPBigBoxSetPage("quickMainWindow", false)
    TestAssert(CPBigBoxNavigationControls.Length = 5 && CPBigBoxControls["ov_opacity"].Enabled
        && CPBigBoxControls["ov_topmost"].Enabled
        && !CPBigBoxControls["ov_bg"].Enabled && !CPBigBoxControls["ov_position"].Enabled,
        "Main window exposes background opacity and Always on top with standard return actions")
    CPBigBoxOverlayAction("opacity")
    CPBigBoxOverlaySliderMove("Left")
    CPBigBoxOverlaySliderMove("Left")
    CPBigBoxOverlaySliderMove("Left")
    TestAssert(CPBigBoxBackgroundPreview = 85 && CPBigBoxBackgroundOpacity = 100
        && FileRead(iniPath) = before && TestOverlayThemes.Length = sends,
        "Three controller steps preview 85 percent without writes or changing translation overlays")
    TestAssert(!CPBigBoxBackdrop.Get("failed", false) && CPBigBoxBackdrop.Has("gui"),
        "Native background-only composition succeeds")
    backdropHwnd := CPBigBoxBackdrop["gui"].Hwnd
    TestAssert(DllCall("user32\GetLayeredWindowAttributes", "ptr", CPBigBoxGui.Hwnd,
        "ptr", color, "ptr", alpha, "ptr", flags) && NumGet(flags, 0, "uint") = 1,
        "Foreground uses only a color key, never uniform alpha that would fade text")
    DllCall("user32\GetLayeredWindowAttributes", "ptr", backdropHwnd,
        "ptr", color, "ptr", alpha, "ptr", flags)
    TestAssert(NumGet(flags, 0, "uint") = 2 && NumGet(alpha, 0, "uchar") = Round(85 * 255 / 100),
        "Only the background receives 85 percent alpha")
    TestAssert((WinGetExStyle(backdropHwnd) & 0x08000000) && !(WinGetExStyle(backdropHwnd) & 0x20),
        "Backdrop is non-activating and does not pass clicks through to the game")
    TestAssert(CPBigBoxOverlaySliderFocused() = "ov_slider1",
        "Live preview retains the dashboard's slider focus")
    for dims in [[1280, 720], [1920, 1080], [3840, 2160]] {
        CPBigBoxGui.Show("NA x0 y0 w" dims[1] " h" dims[2])
        CPBigBoxDashboardResize(CPBigBoxGui, 0, dims[1], dims[2])
        CPBigBoxBackgroundSync(true)
        CPBigBoxGui.GetClientPos(&fx, &fy, &fw, &fh)
        CPBigBoxBackdrop["gui"].GetClientPos(&bx, &by, &bw, &bh)
        TestAssert(fx = bx && fy = by && fw = bw && fh = bh,
            "Background matches dashboard position and size at " dims[1])
        TestAssert(DllCall("user32\GetWindow", "ptr", CPBigBoxGui.Hwnd, "uint", 2, "ptr") = backdropHwnd,
            "Dimmer stays directly behind the menu at " dims[1])
        CPBigBoxControls["modeBody"].GetPos(,,, &bh)
        TestAssert(TestTextHeight(CPBigBoxControls["modeBody"]) <= bh,
            "Live preview instructions fit at " dims[1])
        TestCapture("background-opacity-85-" dims[1] ".png", dims[1], dims[2])
    }
    CPBigBoxGui.Hide()
    CPBigBoxBack()
    TestAssert(CPBigBoxEffectiveBackgroundOpacity() = 100 && FileRead(iniPath) = before
        && !DllCall("user32\IsWindowVisible", "ptr", backdropHwnd)
        && CPBigBoxFocusFrame["key"] = "ov_opacity", "B cancels the preview and restores solid rendering and originating focus")
    CPBigBoxOverlayAction("opacity")
    CPBigBoxControls["ov_slider1"].Value := 85
    CPBigBoxOverlaySliderChanged()
    CPBigBoxSaveOverlayEditor()
    TestAssert(CPBigBoxBackgroundOpacity = 85 && CPBigBoxBackgroundPreview = -1
        && IniRead(iniPath, "cfg_control", "bigBoxBackgroundOpacity") = 85,
        "Save persists the fullscreen setting and retains its appearance")
    for newPage in ["home", "quickOverlays", "quickTranslation", "controls", "translationWindow", "apiKeys"] {
        CPBigBoxSetPage(newPage, false)
        TestAssert(CPBigBoxEffectiveBackgroundOpacity() = 85 && CPBigBoxGui.BackColor = "010203"
            && !CPBigBoxBackdrop.Get("failed", false), "Submenu retains background-only opacity: " newPage)
    }
    CPBigBoxSetPage("quickTranslation", false)
    CPBigBoxOpenAIChoice("model")
    TestAssert(CPBigBoxAIChoiceActive() && CPBigBoxEffectiveBackgroundOpacity() = 85, "Choice lists inherit menu opacity")
    CPBigBoxBack()
    CPBigBoxSetPage("quickMainWindow", false)
    CPBigBoxOverlayAction("opacity")
    Loop 15
        CPBigBoxOverlaySliderMove("Left")
    TestAssert(CPBigBoxBackgroundPreview = 50, "Opacity is clamped at a readable 50 percent minimum")
    Loop 15
        CPBigBoxOverlaySliderMove("Right")
    TestAssert(CPBigBoxBackgroundPreview = 100 && !DllCall("user32\IsWindowVisible", "ptr", backdropHwnd),
        "100 percent preview returns to normal solid rendering")
    CPBigBoxCloseOverlayEditor()
    TestAssert(CPBigBoxEffectiveBackgroundOpacity() = 85, "Cancel restores the saved non-default opacity too")
    for darkTheme in [false, true] {
        controlDarkMode := darkTheme
        CPBigBoxDashboardApplyTheme()
        TestAssert(CPBigBoxGui.BackColor = "010203"
            && CPBigBoxBackdrop["gui"].BackColor = CPPalette(darkTheme)["window"],
            "Light/dark mode recolors the background without fading foreground text: " darkTheme)
        CPBigBoxBackgroundPreview := 100
        CPBigBoxDashboardApplyTheme()
        TestAssert(CPBigBoxGui.BackColor = CPPalette(darkTheme)["window"],
            "100 percent restores the original theme background: " darkTheme)
        CPBigBoxBackgroundPreview := -1
    }
    CPBigBoxBackdrop["failed"] := true
    CPBigBoxDashboardApplyTheme()
    TestAssert(CPBigBoxEffectiveBackgroundOpacity() = 100 && CPBigBoxBackgroundOpacity = 85
        && CPBigBoxGui.BackColor = CPPalette(true)["window"],
        "Composition failure uses a solid fallback without discarding the saved preference")
    CPBigBoxBackdrop["failed"] := false
    CPBigBoxDashboardApplyTheme()
    CPBigBoxGui.Show("NA")
    CPBigBoxBackgroundSync(true)
    TestAssert(DllCall("user32\IsWindowVisible", "ptr", backdropHwnd), "Saved translucent background is shown with the menu")
    CPBigBoxGui.Hide()
    Sleep(20)
    TestAssert(!DllCall("user32\IsWindowVisible", "ptr", backdropHwnd), "Direct capture/Study-style hide also removes the backdrop")
    CPBigBoxDashboardShowReady()
    TestAssert(CPBigBoxEffectiveBackgroundOpacity() = 85 && DllCall("user32\IsWindowVisible", "ptr", backdropHwnd),
        "Returning from a hidden dashboard restores the saved background")
    CPSetOverlayPreference("Main window", "opacity", 100)
    CPBigBoxDashboardApplyTheme()
    CPBigBoxBackgroundDispose()
    TestAssert(!DllCall("user32\IsWindow", "ptr", backdropHwnd), "Background cleanup destroys its native window")
    IniDelete(iniPath, "cfg_control", "bigBoxBackgroundOpacity")
    CPBigBoxGui.Hide()
    CPBigBoxSetPage("home", false)
}

TestBigBoxAlwaysOnTop() {
    global
    local originalIni := iniPath, desktopTop := IniRead(iniPath, "cfg_control", "winTop", "missing")
    local overlaySends := TestOverlayThemes.Length, other := Gui("+ToolWindow", "Synthetic other app")
    local dims, opacity, foregroundBefore, backdropHwnd, x, y, w, h, key, page
    try {
        CPBigBoxGui.Hide()
        CPBigBoxSetPage("quickMainWindow", false)
        TestAssert(CPBigBoxAlwaysOnTop && (WinGetExStyle(CPBigBoxGui.Hwnd) & 0x8),
            "Fullscreen Always on top defaults to the existing topmost behavior")
        for dims in [[1280, 720], [1920, 1080], [3840, 2160]] {
            CPBigBoxGui.Show("Hide w" dims[1] " h" dims[2])
            CPBigBoxDashboardResize(CPBigBoxGui, 0, dims[1], dims[2])
            TestLayout(dims[1], dims[2])
            for key in ["ov_opacity", "ov_topmost", "ov_note"] {
                CPBigBoxControls[key].GetPos(&x, &y, &w, &h)
                TestAssert(TestTextHeight(CPBigBoxControls[key]) <= h,
                    "Main window setting text fits at " dims[1] ": " key)
            }
            TestCapture("main-window-settings-" dims[1] ".png", dims[1], dims[2])
        }
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_topmost"]))
        CPBigBoxOverlayAction("topmost")
        TestAssert(!CPBigBoxAlwaysOnTop && !(WinGetExStyle(CPBigBoxGui.Hwnd) & 0x8)
            && IniRead(iniPath, "cfg_control", "bigBoxAlwaysOnTop") = 0
            && CPBigBoxControls["ov_topmost"].Text = "Always on top`nOff"
            && CPBigBoxFocusFrame["key"] = "ov_topmost",
            "A toggles, saves, and applies Always on top without moving controller focus")
        for opacity in [85, 100] {
            CPSetOverlayPreference("Main window", "opacity", opacity)
            CPBigBoxDashboardApplyTheme()
            CPBigBoxDashboardShowReady()
            backdropHwnd := CPBigBoxBackdrop["gui"].Hwnd
            TestAssert(!(WinGetExStyle(CPBigBoxGui.Hwnd) & 0x8)
                && !(WinGetExStyle(backdropHwnd) & 0x8),
                "Reopening honors Off for both foreground and backing at opacity " opacity)
            other.Show("NA x20 y20 w300 h160")
            DllCall("user32\SetWindowPos", "ptr", other.Hwnd, "ptr", 0,
                "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x0013)
            foregroundBefore := DllCall("user32\GetForegroundWindow", "ptr")
            TestAssert(TestWindowIsAbove(other.Hwnd, CPBigBoxGui.Hwnd),
                "A normal external window can cover the relaxed fullscreen menu")
            CPBigBoxDashboardApplyTheme()
            CPBigBoxDashboardUpdateContent()
            CPBigBoxBackgroundSync()
            TestAssert(TestWindowIsAbove(other.Hwnd, CPBigBoxGui.Hwnd)
                && TestWindowIsAbove(other.Hwnd, backdropHwnd)
                && DllCall("user32\GetForegroundWindow", "ptr") = foregroundBefore,
                "Menu refresh never reclaims z-order or foreground from the other window")
            ; Simulate the same temporary demotion used around native file pickers.
            CPBigBoxModalDepth += 1
            CPBigBoxBackgroundSync()
            CPBigBoxModalDepth -= 1
            CPBigBoxBackgroundSync()
            TestAssert(!(WinGetExStyle(CPBigBoxGui.Hwnd) & 0x8)
                && !(WinGetExStyle(backdropHwnd) & 0x8),
                "Returning from a native dialog does not force Always on top back on")
            CPBigBoxGui.Hide()
        }
        other.Hide()
        CPBigBoxOverlayAction("topmost")
        TestAssert(CPBigBoxAlwaysOnTop && (WinGetExStyle(CPBigBoxGui.Hwnd) & 0x8)
            && (WinGetExStyle(backdropHwnd) & 0x8)
            && IniRead(iniPath, "cfg_control", "bigBoxAlwaysOnTop") = 1,
            "Turning On restores topmost behavior on both layers and saves it")
        CPBigBoxModalDepth += 1
        CPBigBoxBackgroundSync()
        TestAssert(!(WinGetExStyle(CPBigBoxGui.Hwnd) & 0x8)
            && !(WinGetExStyle(backdropHwnd) & 0x8), "Native file selection temporarily relaxes both topmost layers")
        CPBigBoxModalDepth -= 1
        CPBigBoxBackgroundSync()
        TestAssert((WinGetExStyle(CPBigBoxGui.Hwnd) & 0x8)
            && (WinGetExStyle(backdropHwnd) & 0x8), "Native dialog return restores On only when it is the saved preference")
        iniPath := A_ScriptDir "\missing-topmost-folder\settings.ini"
        CPBigBoxOverlayAction("topmost")
        TestAssert(CPBigBoxAlwaysOnTop && (WinGetExStyle(CPBigBoxGui.Hwnd) & 0x8)
            && InStr(CPBigBoxOverlayNotice, "Could not save"), "Failed toggle persistence leaves the current behavior unchanged")
        iniPath := originalIni
        for page in ["quickTranslatorWindow", "quickExplainerWindow", "translationWindow", "explanationWindow"] {
            CPBigBoxSetPage(page, false)
            TestAssert(!CPBigBoxControls["ov_topmost"].Enabled && !TestControlShown(CPBigBoxControls["ov_topmost"]),
                "Main-window-only toggle is not added to translation overlays: " page)
        }
        TestAssert(IniRead(iniPath, "cfg_control", "winTop", "missing") = desktopTop
            && TestOverlayThemes.Length = overlaySends, "Fullscreen toggle never changes desktop or translation-overlay settings")
    } finally {
        iniPath := originalIni
        CPBigBoxModalDepth := 0
        CPBigBoxAlwaysOnTop := true
        CPBigBoxBackgroundOpacity := 100
        CPBigBoxBackgroundPreview := -1
        CPBigBoxDashboardApplyTheme()
        CPBigBoxBackgroundDispose()
        try IniDelete(iniPath, "cfg_control", "bigBoxAlwaysOnTop")
        try IniDelete(iniPath, "cfg_control", "bigBoxBackgroundOpacity")
        other.Destroy()
        CPBigBoxGui.Hide()
        CPBigBoxSetPage("home", false)
    }
}

TestBigBoxNativePickerBackend(spec) {
    global CPBigBoxModalDepth, CPBigBoxGui, TestBigBoxPickerSpec
    TestBigBoxPickerSpec := spec
    TestAssert(CPBigBoxModalDepth = 1 && !CPBigBoxEffectiveAlwaysOnTop(),
        "Owned native picker temporarily yields fullscreen topmost status")
    DllCall("user32\EnableWindow", "ptr", CPBigBoxGui.Hwnd, "int", 0)
    return "C:\picked\python.exe"
}

TestNativePickerOwnership() {
    global CPBigBoxGui, CPBigBoxModalDepth, TestBigBoxPickerSpec := 0
    CPBigBoxGui.Show("Hide w1920 h1080")
    CPBigBoxModalDepth := 0
    selected := CPNativeFileSelect(
        CPBigBoxGui.Hwnd, 3, "C:\configured\python.exe",
        "Select Python interpreter", "Programs (*.exe)",
        TestBigBoxNativePickerBackend
    )
    TestAssert(selected = "C:\picked\python.exe"
        && IsObject(TestBigBoxPickerSpec)
        && TestBigBoxPickerSpec["owner"] = CPBigBoxGui.Hwnd
        && TestBigBoxPickerSpec["presentation"] = "fullscreen"
        && TestBigBoxPickerSpec["options"] = 3
        && TestBigBoxPickerSpec["root"] = "C:\configured\python.exe"
        && TestBigBoxPickerSpec["prompt"] = "Select Python interpreter"
        && TestBigBoxPickerSpec["filter"] = "Programs (*.exe)",
        "Fullscreen native picker preserves the complete selection request")
    TestAssert(CPBigBoxModalDepth = 0
        && DllCall("user32\IsWindowEnabled", "ptr", CPBigBoxGui.Hwnd),
        "Fullscreen native picker restores modal depth and owner state")
    CPBigBoxGui.Hide()
}

TestWindowIsAbove(upper, lower) {
    local hwnd := lower
    Loop 500 {
        hwnd := DllCall("user32\GetWindow", "ptr", hwnd, "uint", 3, "ptr") ; GW_HWNDPREV
        if !hwnd
            return false
        if hwnd = upper
            return true
    }
    return false
}

TestOverlayStatusStates() {
    global TestOverlayWindows
    local statusGui := Gui("+ToolWindow +AlwaysOnTop", "Synthetic overlay state")
    TestAssert(CPBigBoxOverlayState("Translator") = "Closed", "Missing overlay reports Closed")
    TestOverlayWindows["Translator"] := statusGui.Hwnd
    TestAssert(CPBigBoxOverlayState("Translator") = "Hidden", "Actually hidden overlay reports Hidden")
    statusGui.Show("NA x-32000 y-32000 w80 h40")
    TestAssert(CPBigBoxOverlayState("Translator") = "Visible", "Visible topmost overlay reports Visible")
    statusGui.Opt("-AlwaysOnTop")
    TestAssert(CPBigBoxOverlayState("Translator") = "Hidden", "Safe sent-behind overlay reports Hidden")
    statusGui.Hide()
    TestAssert(CPBigBoxOverlayState("Translator") = "Hidden", "Hidden non-topmost overlay remains Hidden")
    statusGui.Destroy()
    TestOverlayWindows.Delete("Translator")
    TestAssert(CPBigBoxOverlayState("Translator") = "Closed", "Destroyed overlay reports Closed")
}

TestBigBoxModelManagement() {
    global
    originalGemini := model_gemini_img.Clone()
    originalOpenAIAudio := model_openai_audio.Clone()
    newModel := "gemini-stage3h-new"
    try {
        parsedCatalog := ModelCatalogParseOutput(
            "JRPG_MODEL_CATALOG_V1`nSTATUS`tOK`nPROVIDER`tgemini`nPURPOSE`tscreenshot`n"
                . "SOURCE`tcache`nFETCHED_AT`t2026-09-03T12:00:00Z`n"
                . "MODEL`tgemini-parser-test`tParser Test`nWARNING`tSaved response`n",
            "gemini", "screenshot")
        TestAssert(parsedCatalog["ok"] && parsedCatalog["source"] = "cache"
            && parsedCatalog["models"][1] = "gemini-parser-test"
            && parsedCatalog["displayNames"]["gemini-parser-test"] = "Parser Test"
            && parsedCatalog["warnings"][1] = "Saved response",
            "Shared catalogue parser preserves structured model metadata")
        malformedCatalog := ModelCatalogParseOutput("not a catalogue", "openai", "audio", "Synthetic unreadable")
        TestAssert(!malformedCatalog["ok"] && malformedCatalog["error"] = "Synthetic unreadable",
            "Shared catalogue parser gives a bounded fallback error")
        for testSpec in [
            ["translation", "OpenAI", "openai_img", "screenshot"],
            ["translation", "Gemini", "gemini_img", "screenshot"],
            ["explanation", "OpenAI", "openai_explain", "explanation"],
            ["explanation", "Gemini", "gemini_explain", "explanation"],
            ["audio", "OpenAI", "openai_audio", "audio"],
            ["audio", "Gemini", "gemini_audio", "audio"]
        ] {
            testModelMapping := CPBigBoxModelSpec(testSpec[1], testSpec[2])
            TestAssert(testModelMapping["key"] = testSpec[3] && testModelMapping["purpose"] = testSpec[4],
                "Model manager maps " testSpec[1] "/" testSpec[2] " to its independent list")
        }

        ddlProv.Choose(1)
        AutoPersist()
        CPBigBoxGui.Show("Hide w1920 h1080")
        CPBigBoxDashboardResize(CPBigBoxGui, 0, 1920, 1080)
        CPBigBoxSetPage("quickTranslation", false)
        CPBigBoxOpenAIChoice("model")
        TestAssert(TestControlShown(CPBigBoxControls["choiceManage"])
            && CPBigBoxNavigationControls.Length = 3,
            "Model picker exposes controller-friendly management without unbounded controls")
        CPBigBoxOpenModelManager()
        TestAssert(CPBigBoxModelManageActive() && CPBigBoxAIChoice["field"] = "modelManage"
            && CPBigBoxAIChoice["options"].Length = 4,
            "Manage models opens four explicit controller actions")
        TestLayout(1920, 1080)
        TestCapture("model-manager-actions-1920.png", 1920, 1080)
        TestAssert(TestExternalCalls = 0, "Opening model management does not launch a desktop dialog")

        manualModel := "gemini-manual-controller-test"
        CPBigBoxCommitAIChoiceIndex(2)
        TestAssert(CPBigBoxCurrentPage = "setupEdit" && CPBigBoxSetupState["flow"] = "modelManual",
            "Manual model entry uses the fullscreen editor")
        CPBigBoxControls["setup_edit"].Value := manualModel
        CPBigBoxSaveSetup()
        TestAssert(ModelAlreadyAdded(model_gemini_img, manualModel) && ddlIMG_GM.Text = manualModel
            && CPBigBoxCurrentPage = "quickTranslation",
            "Manual model ID is persisted, selected and returned to its Model tile")
        CPBigBoxOpenAIChoice("model")
        CPBigBoxOpenModelManager()

        existing := model_gemini_img[1]
        TestCatalogResult := Map("ok", true, "source", "cache",
            "models", [existing, newModel], "warnings", ["Using a saved catalogue."], "error", "")
        requestCount := TestCatalogRequests.Length
        CPBigBoxCommitAIChoiceIndex(1)
        TestAssert(TestCatalogRequests.Length = requestCount + 1
            && TestCatalogRequests[-1][1] = "gemini"
            && TestCatalogRequests[-1][2] = "screenshot" && !TestCatalogRequests[-1][3],
            "Add from catalogue requests the active provider and purpose")
        TestAssert(CPBigBoxAIChoice["field"] = "modelCatalog"
            && CPBigBoxAIChoice["options"].Length = 1 && CPBigBoxAIChoice["options"][1] = newModel,
            "Catalogue filters models already present in the local list")
        TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "saved catalogue cache"),
            "Catalogue source and warning are visible in the modern view")
        TestLayout(1920, 1080)
        TestCapture("model-manager-catalogue-1920.png", 1920, 1080)
        writesBefore := TestModelWrites.Length
        CPBigBoxCommitAIChoiceIndex(1)
        TestAssert(!CPBigBoxModelManageActive() && !CPBigBoxAIChoiceActive()
            && ModelAlreadyAdded(model_gemini_img, newModel) && ddlIMG_GM.Text = newModel,
            "Adding a catalogue model selects it and returns to the originating page")
        testGeminiWriteFound := false
        Loop TestModelWrites.Length - writesBefore
            if TestModelWrites[writesBefore + A_Index][1] = "gemini_img"
                testGeminiWriteFound := true
        TestAssert(TestModelWrites.Length > writesBefore
            && testGeminiWriteFound && geminiImgModel = newModel,
            "Added model persists through the independent screenshot model setting")
        TestAssert(CPBigBoxFocusFrame["key"] = "ai_model", "Successful add restores Model tile focus")

        CPBigBoxOpenAIChoice("model")
        CPBigBoxOpenModelManager()
        CPBigBoxCommitAIChoiceIndex(4)
        TestAssert(CPBigBoxAIChoice["field"] = "modelRemove", "Remove action opens the local model list")
        removeIndex := ArrayIndexOf(CPBigBoxAIChoice["options"], newModel)
        CPBigBoxCommitAIChoiceIndex(removeIndex)
        TestAssert(CPBigBoxAIChoice["field"] = "modelRemoveConfirm"
            && InStr(CPBigBoxControls["modeBody"].Text, newModel),
            "Removal requires an explicit confirmation naming the model")
        CPBigBoxBack()
        TestAssert(CPBigBoxAIChoice["field"] = "modelRemove", "B from confirmation returns to the removal list")
        removeIndex := ArrayIndexOf(CPBigBoxAIChoice["options"], newModel)
        CPBigBoxCommitAIChoiceIndex(removeIndex)
        CPBigBoxCommitAIChoiceIndex(1)
        TestAssert(!ModelAlreadyAdded(model_gemini_img, newModel) && ddlIMG_GM.Text != newModel,
            "Removing the active model selects a safe remaining model")

        CPBigBoxOpenAIChoice("model")
        CPBigBoxOpenModelManager()
        TestCatalogResult := Map("ok", false, "source", "none", "models", [], "warnings", [],
            "error", "Synthetic provider error")
        CPBigBoxCommitAIChoiceIndex(1)
        TestAssert(CPBigBoxAIChoice["field"] = "modelLoadError"
            && InStr(CPBigBoxControls["modeBody"].Text, "Synthetic provider error"),
            "Catalogue failure is visible and offers a retry instead of waiting silently")
        CPBigBoxBack()
        TestAssert(CPBigBoxAIChoice["field"] = "modelManage", "B from catalogue error returns to model actions")

        TestCatalogResult := Map("ok", true, "source", "stale_cache",
            "models", model_gemini_img.Clone(), "warnings", ["The online request failed."], "error", "")
        CPBigBoxCommitAIChoiceIndex(1)
        TestAssert(CPBigBoxAIChoice["field"] = "modelCatalogEmpty"
            && InStr(CPBigBoxControls["modeBody"].Text, "older saved catalogue cache"),
            "Empty stale catalogue explains both the cache state and lack of new models")
        CPBigBoxCommitAIChoiceIndex(1)
        TestAssert(TestCatalogRequests[-1][3], "Refresh from an empty catalogue bypasses the fresh cache")

        ; A deferred synthetic request exercises the real loading/cancellation UI contract.
        CPBigBoxBack()
        TestCatalogDeferred := true
        cancelsBefore := TestCatalogCancels
        CPBigBoxCommitAIChoiceIndex(3)
        TestAssert(CPBigBoxAIChoice["field"] = "modelLoading"
            && CPBigBoxNavigationControls.Length = 1
            && CPBigBoxNavigationControls[1].Hwnd = CPBigBoxControls["choiceBack"].Hwnd,
            "Loading state has one safe Cancel focus stop")
        TestLayout(1920, 1080)
        TestCapture("model-manager-loading-1920.png", 1920, 1080)
        CPBigBoxBack()
        TestAssert(TestCatalogCancels = cancelsBefore + 1 && CPBigBoxAIChoice["field"] = "modelManage",
            "Cancelling loading stops only the exact catalogue request and returns to actions")
        TestCatalogDeferred := false

        ; A model-list change outside the open manager invalidates its snapshot.
        model_gemini_img.Push("synthetic-external-change")
        CPBigBoxCommitAIChoiceIndex(1)
        TestAssert(CPBigBoxAIChoice["field"] = "modelManage"
            && InStr(CPBigBoxControls["modeBody"].Text, "changed elsewhere"),
            "Stale model manager refuses to overwrite an externally changed list")
        model_gemini_img.Pop()
        CPBigBoxBack() ; manager actions -> original model picker
        CPBigBoxBack() ; model picker -> page

        ; Never allow an empty list: empty persisted lists would reload defaults on startup.
        CPBigBoxReplaceModelArray(model_openai_audio, [originalOpenAIAudio[1]])
        RefreshModelCombos("openai_audio", ddlTR, originalOpenAIAudio[1])
        ddlAProv.Choose(2)
        AutoPersist()
        CPBigBoxSetPage("quickAudio", false)
        CPBigBoxOpenAIChoice("model")
        CPBigBoxOpenModelManager()
        CPBigBoxCommitAIChoiceIndex(4)
        TestAssert(CPBigBoxAIChoice["field"] = "modelManage"
            && InStr(CPBigBoxControls["modeBody"].Text, "At least one model"),
            "Last local model cannot be removed")

        ; Hiding the dashboard cancels an in-flight catalogue child and clears nested state.
        TestCatalogDeferred := true
        CPBigBoxCommitAIChoiceIndex(3)
        cancelsBefore := TestCatalogCancels
        CPBigBoxDashboardHide(false)
        TestAssert(!CPBigBoxModelManageActive() && !CPBigBoxAIChoiceActive()
            && TestCatalogCancels = cancelsBefore + 1,
            "Dashboard shutdown cancels catalogue work and clears model-management state")
    } finally {
        TestCatalogDeferred := false
        TestCatalogBusy := false
        TestCatalogCallback := 0
        CPBigBoxResetModelManage()
        CPBigBoxAIChoice := Map("active", false)
        CPBigBoxReplaceModelArray(model_gemini_img, originalGemini)
        CPBigBoxReplaceModelArray(model_openai_audio, originalOpenAIAudio)
        RefreshModelCombos("gemini_img", ddlIMG_GM, originalGemini[1])
        RefreshModelCombos("openai_audio", ddlTR, originalOpenAIAudio[1])
        ddlProv.Choose(1)
        ddlAProv.Choose(1)
        AutoPersist()
        CPBigBoxSetPage("home", false)
    }
}

TestBigBoxSetupManagement() {
    global
    secret := "synthetic-test-key-never-display"
    promptName := "controller_test_prompt"
    promptPath := promptsDir "\" promptName ".txt"
    originalPython := pythonExe
    originalDirect := directModelOutput
    originalDebug := debugMode
    originalPrompt := ddlPrompt.Text
    try {
        try if FileExist(envPath)
            FileDelete(envPath)
        CPBigBoxSyncApiDesktop()
        showPathsTab := true
        CPBigBoxSetPage("apiKeys", false)
        TestAssert(CPBigBoxSettingsGroups("apiKeys").Length = 3
            && CPBigBoxNavigationControls.Length = 8,
            "API Keys is a complete grouped fullscreen page")
        TestLayout(1920, 1080)
        TestCapture("api-keys-1920.png", 1920, 1080)
        TestAssert(!CPBigBoxControls["api_delete"].Enabled,
            "Delete in-app keys is disabled when no key file exists")
        CPBigBoxOpenApiEditor("gemini")
        TestAssert(CPBigBoxCurrentPage = "setupEdit" && CPBigBoxControls["setup_secret"].Visible,
            "Provider key opens a masked fullscreen editor")
        TestLayout(1920, 1080)
        TestCapture("api-key-editor-1920.png", 1920, 1080)
        CPBigBoxControls["setup_secret"].Value := secret
        CPBigBoxSaveSetup()
        TestAssert(CPBigBoxCurrentPage = "apiKeys" && CPBigBoxApiLocalValue("gemini") = secret,
            "Masked API editor saves the Gemini aliases atomically")
        TestAssert(!InStr(CPBigBoxControls["modeTitle"].Text, secret)
            && !InStr(CPBigBoxControls["modeBody"].Text, secret),
            "Saved API secret is never copied into dashboard labels")
        CPBigBoxOpenApiEditor("gemini")
        CPBigBoxDeleteApiRequest("gemini")
        TestAssert(CPBigBoxCurrentPage = "setupConfirm", "Removing one in-app key requires confirmation")
        CPBigBoxCancelSetup()
        TestAssert(CPBigBoxCurrentPage = "setupEdit" && CPBigBoxApiLocalValue("gemini") = secret,
            "Cancelling API-key removal preserves the key")
        CPBigBoxDeleteApiRequest("gemini")
        CPBigBoxConfirmSetup()
        TestAssert(CPBigBoxCurrentPage = "apiKeys" && CPBigBoxApiLocalValue("gemini") = "",
            "Confirmed API-key removal clears only the in-app provider key")

        CPBigBoxSetPage("paths", false)
        TestAssert(CPBigBoxSettingsGroups("paths").Length = 3,
            "Paths is a complete grouped fullscreen page")
        TestLayout(1920, 1080)
        TestCapture("paths-1920.png", 1920, 1080)
        CPBigBoxOpenPathEditor("python")
        CPBigBoxControls["setup_edit"].Value := A_ScriptFullPath
        CPBigBoxSaveSetup()
        TestAssert(pythonExe = A_ScriptFullPath && IniRead(iniPath, "cfg", "pythonExe", "") = A_ScriptFullPath,
            "Existing path saves to the shared runtime setting")
        CPBigBoxOpenPathEditor("python")
        missingPath := A_ScriptDir "\missing-python.exe"
        CPBigBoxControls["setup_edit"].Value := missingPath
        CPBigBoxSaveSetup()
        TestAssert(CPBigBoxCurrentPage = "setupConfirm" && pythonExe = A_ScriptFullPath,
            "Missing path waits for an explicit save-anyway confirmation")
        CPBigBoxCancelSetup()
        TestAssert(CPBigBoxCurrentPage = "setupEdit" && CPBigBoxControls["setup_edit"].Value = missingPath,
            "Cancelling missing-path confirmation returns to the pending editor")
        CPBigBoxSaveSetup()
        CPBigBoxConfirmSetup()
        TestAssert(pythonExe = missingPath, "Confirmed missing path uses the desktop tab's save-anyway behavior")
        CPBigBoxSetPage("paths", false)
        CPBigBoxTogglePathOption("direct")
        CPBigBoxTogglePathOption("debug")
        TestAssert(directModelOutput != originalDirect && debugMode != originalDebug,
            "Advanced path-page switches save immediately")

        CPBigBoxSetPage("quickTranslation", false)
        CPBigBoxOpenAIChoice("detail")
        TestAssert(CPBigBoxPromptManageButtonVisible() && TestControlShown(CPBigBoxControls["choiceManage"]),
            "Translation prompt picker exposes fullscreen management")
        CPBigBoxOpenChoiceManager()
        TestAssert(CPBigBoxCurrentPage = "setupTools" && CPBigBoxSetupState["flow"] = "promptTools",
            "Prompt management opens without a desktop dialog")
        TestLayout(1920, 1080)
        TestCapture("prompt-tools-1920.png", 1920, 1080)
        CPBigBoxPromptAction("new")
        CPBigBoxControls["setup_edit"].Value := promptName
        CPBigBoxSaveSetup()
        TestAssert(CPBigBoxSetupState["flow"] = "promptText" && !FileExist(promptPath),
            "New prompt is not created before its text is confirmed")
        TestLayout(1920, 1080)
        TestCapture("prompt-editor-1920.png", 1920, 1080)
        CPBigBoxControls["setup_raw"].Value := "Test prompt {jp}"
        CPBigBoxSaveSetup()
        TestAssert(FileExist(promptPath) && ddlPrompt.Text = promptName
            && FileRead(promptPath, "UTF-8") = "Test prompt {jp}",
            "New prompt saves atomically and becomes the active selection")
        CPBigBoxOpenAIChoice("detail")
        CPBigBoxOpenChoiceManager()
        CPBigBoxPromptAction("edit")
        promptSelectionStart := Buffer(4, 0)
        promptSelectionEnd := Buffer(4, 0)
        SendMessage(0x00B0, promptSelectionStart.Ptr, promptSelectionEnd.Ptr,
            CPBigBoxControls["setup_raw"].Hwnd) ; EM_GETSEL
        TestAssert(CPBigBoxPromptEditorActive()
            && NumGet(promptSelectionStart, 0, "uint") = 0
            && NumGet(promptSelectionEnd, 0, "uint") = 0,
            "Prompt editor opens keyboard-ready with an unselected caret")
        promptEditStyle := DllCall("user32\GetWindowLongPtr", "ptr",
            CPBigBoxControls["setup_raw"].Hwnd, "int", -16, "ptr")
        TestAssert(!(promptEditStyle & 0x0800),
            "Fullscreen prompt text remains writable rather than read-only")
        CPBigBoxControls["setup_raw"].Value := "Edited prompt {jp}"
        CPBigBoxSaveSetup()
        TestAssert(FileRead(promptPath, "UTF-8") = "Edited prompt {jp}" && FileExist(promptPath ".bak"),
            "Prompt editing creates a backup and returns to its setting tile")
        CPBigBoxOpenAIChoice("detail")
        CPBigBoxOpenChoiceManager()
        CPBigBoxPromptAction("delete")
        CPBigBoxCancelSetup()
        TestAssert(FileExist(promptPath) && CPBigBoxCurrentPage = "setupTools",
            "Cancelling prompt deletion preserves the file and returns to tools")
        CPBigBoxPromptAction("delete")
        CPBigBoxConfirmSetup()
        TestAssert(!FileExist(promptPath) && CPBigBoxCurrentPage = "quickTranslation",
            "Confirmed prompt deletion selects a remaining prompt safely")
    } finally {
        try if FileExist(envPath)
            FileDelete(envPath)
        try if FileExist(promptPath)
            FileDelete(promptPath)
        try if FileExist(promptPath ".bak")
            FileDelete(promptPath ".bak")
        CPBigBoxSyncApiDesktop()
        CPBigBoxWritePath("python", originalPython)
        if directModelOutput != originalDirect
            CPBigBoxTogglePathOption("direct")
        if debugMode != originalDebug
            CPBigBoxTogglePathOption("debug")
        RefreshPromptProfilesList(originalPrompt)
        CPBigBoxApplyAISelection("translation", "detail")
        showPathsTab := false
        CPBigBoxResetSetup()
        CPBigBoxAIChoice := Map("active", false)
        CPBigBoxSetPage("home", false)
    }
}

TestPickerReturnFocus() {
    global CPBigBoxGui, CPBigBoxControls, CPBigBoxCurrentPage, CPBigBoxPageFocus
    global CPBigBoxAIChoice, CPBigBoxFocusFrame, CPBigBoxPageChanging, TestPickerFocusProbe, iniPath
    callback := CallbackCreate(TestPickerTransitionFocusProbe, , 6)
    probed := []
    try {
        for key in ["choice1", "choice2", "choice3", "choice4", "choiceManage", "choiceBack", "listChoice3"] {
            handle := CPBigBoxControls[key].Hwnd
            if !DllCall("comctl32\SetWindowSubclass", "ptr", handle, "ptr", callback, "uptr", 2, "uptr", 0)
                throw Error("Could not install the picker focus probe.")
            probed.Push(handle)
        }
        entries := []
        for page in ["quickTranslation", "quickExplanation", "quickAudio", "screenshot", "explanation", "audio"]
            for field in ["provider", "model", "detail"]
                entries.Push([page, field])
        entries.Push(["audio", "device"])
        for entry in entries {
            for action in ["cancel", "legacyCancel", "escape", "same", "changed"] {
                CPBigBoxSetPage(entry[1], false)
                returnKey := entry[2] = "device" ? "audio_device" : "ai_" entry[2]
                CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls[returnKey]))
                CPBigBoxOpenAIChoice(entry[2])
                TestAssert(CPBigBoxAIChoiceActive(), "Return-focus picker opens: " entry[1] "/" entry[2])
                state := CPBigBoxAIChoice
                before := FileRead(iniPath)
                probeCount := TestPickerFocusProbe["count"]
                TestPickerFocusProbe["armed"] := true
                if action = "cancel" || action = "legacyCancel" {
                    empty := action = "cancel" ? TestSnapshot()
                        : Map("name", "Legacy USB gamepad", "tokens", Map())
                    pressed := action = "cancel" ? TestSnapshot("X:B")
                        : Map("name", "Legacy USB gamepad", "tokens", Map("J:Button 2", true))
                    CPControllerResetNavigation()
                    CPControllerHandleNavigation(empty, CPBigBoxGui.Hwnd)
                    CPControllerHandleNavigation(pressed, CPBigBoxGui.Hwnd)
                    Loop 8
                        CPControllerHandleNavigation(pressed, CPBigBoxGui.Hwnd)
                } else if action = "escape" {
                    CPBigBoxKeyboardCancel()
                } else {
                    chosen := action = "same" ? state["index"] : Mod(state["index"], state["options"].Length) + 1
                    CPBigBoxCommitAIChoiceIndex(chosen)
                }
                Sleep(20) ; Deliver queued native Focus callbacks as well.
                TestAssert(TestPickerFocusProbe["error"] = "" && TestPickerFocusProbe["count"] = probeCount + 1,
                    "Native hide/disable probe reproduced a temporary focus change")
                TestAssert(!CPBigBoxAIChoiceActive() && CPBigBoxCurrentPage = entry[1] && !CPBigBoxPageChanging,
                    "Picker closes just one layer: " entry[1] "/" entry[2] "/" action)
                TestAssert(CPBigBoxFocusFrame["key"] = returnKey && CPBigBoxPageFocus[entry[1]] = returnKey
                    && CPBigBoxGui.FocusedCtrl.Hwnd = CPBigBoxControls[returnKey].Hwnd,
                    "Picker restores its exact opening tile despite transient focus: " entry[1] "/" entry[2] "/" action)
                if action != "changed"
                    TestAssert(FileRead(iniPath) = before, "Cancel/current choice does not write settings: " action)
                CPBigBoxDashboardButtonFocused("backHome")
                TestAssert(CPBigBoxFocusFrame["key"] = returnKey && CPBigBoxPageFocus[entry[1]] = returnKey,
                    "Stale Back to Home focus cannot overwrite restored tile")
            }
        }
    } finally {
        TestPickerFocusProbe["armed"] := false
        for handle in probed
            DllCall("comctl32\RemoveWindowSubclass", "ptr", handle, "ptr", callback, "uptr", 2)
        CallbackFree(callback)
        CPBigBoxCloseAIChoice()
        CPBigBoxSetPage("home", false)
    }
}

TestPickerTransitionFocusProbe(hwnd, msg, wParam, lParam, subclassId, refData) {
    global TestPickerFocusProbe, CPBigBoxControls
    result := DllCall("comctl32\DefSubclassProc", "ptr", hwnd, "uint", msg,
        "ptr", wParam, "ptr", lParam, "ptr")
    ; Simulate Windows moving focus as the active picker button is hidden or
    ; disabled. Only synthetic controls are touched, never the foreground app.
    if TestPickerFocusProbe["armed"] && (msg = 0x18 || msg = 0xA) && !wParam && !CPBigBoxAIChoiceActive() {
        TestPickerFocusProbe["armed"] := false
        TestPickerFocusProbe["count"] += 1
        try {
            CPBigBoxControls["backHome"].Focus()
            CPBigBoxDashboardButtonFocused("backHome")
            CPBigBoxRememberPageFocus()
        } catch as ex {
            TestPickerFocusProbe["error"] := ex.Message
        }
    }
    return result
}

TestGroupedFocusAndPainting() {
    global CPBigBoxGui, CPBigBoxControls, CPBigBoxNavigationRows, CPBigBoxFocusFrame
    global CPBigBoxFocusIndex, CPBigBoxNavigationControls, controlDarkMode, TestPaintCounts
    CPBigBoxSetPage("explanation", false)
    TestAssert(CPBigBoxNavigationRows.Length = 5, "Explanation has arrows, three semantic rows and bottom actions")
    for groupIndex, group in CPBigBoxExplanationGroups() {
        for index, key in group[2]
            TestAssert(CPBigBoxNavigationRows[groupIndex + 1][index].Hwnd = CPBigBoxControls[key].Hwnd,
                "Navigation follows semantic group: " key)
    }
    for key in ["ai_provider", "exp_library", "exp_openOnStartup", "advanced", "pageNext"] {
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls[key]))
        TestAssert(CPBigBoxFocusFrame["key"] = key, "Accent frame follows focus: " key)
        TestFocusFrameBounds()
    }
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ai_provider"]))
    CPBigBoxDashboardMoveFocus("Down")
    TestAssert(CPBigBoxFocusFrame["key"] = "exp_library", "Down reaches Study Library from AI settings")
    CPBigBoxDashboardMoveFocus("Down")
    TestAssert(CPBigBoxFocusFrame["key"] = "exp_openOnStartup", "Down reaches Explainer startup from Study Library")
    CPBigBoxDashboardMoveFocus("Right")
    TestAssert(CPBigBoxFocusFrame["key"] = "exp_alwaysOnTop", "Right reaches the other startup option")
    CPBigBoxControls["exp_plainText"].Focus()
    Sleep(30)
    TestAssert(CPBigBoxFocusFrame["key"] = "exp_plainText", "Native mouse-style focus updates the accent frame")
    CPBigBoxDashboardButtonFocused("ai_provider")
    TestAssert(CPBigBoxFocusFrame["key"] = "exp_plainText", "Stale Focus callback cannot move frame back")
    CPBigBoxOpenAIChoice("provider")
    TestAssert(SubStr(CPBigBoxFocusFrame["key"], 1, 6) = "choice", "Picker frame moves to a choice")
    Loop 3
        TestAssert(!TestControlShown(CPBigBoxControls["settingsGroup" A_Index]), "Picker hides group caption " A_Index)
    CPBigBoxBack()
    CPBigBoxSetPage("quickExplanation", false)
    Loop 3
        TestAssert(!TestControlShown(CPBigBoxControls["settingsGroup" A_Index]), "Quick view hides group caption " A_Index)
    TestAssert(TestFontHeight(CPBigBoxControls["ai_provider"]) = TestFontHeight(CPBigBoxControls["advanced"]),
        "Leaving grouped page restores standard button font size")
    controlDarkMode := 0
    CPBigBoxDashboardApplyTheme()
    TestAssert(CPBigBoxFocusColor() = "005A9E", "Light mode uses a dark blue focus frame")
    controlDarkMode := 1
    CPBigBoxDashboardApplyTheme()
    TestAssert(CPBigBoxFocusColor() = "62C7FF", "Dark mode uses a bright blue focus frame")

    ; Exercise real native child paint messages, without covering or activating
    ; anything on the user's desktop. The test GUI stays cloaked and off-screen.
    CPBigBoxSetPage("home", false)
    CPBigBoxGui.Show("NA x-30000 y-30000 w1280 h720")
    CPBigBoxDashboardResize(CPBigBoxGui, 0, 1280, 720)
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["translation"]))
    callback := CallbackCreate(TestPaintProbe, , 6)
    probed := []
    try {
        probeKeys := CPBigBoxButtonKeys()
        probeKeys.Push("modeBody")
        for key in probeKeys {
            handle := CPBigBoxControls[key].Hwnd
            TestPaintCounts[handle] := Map("paint", 0, "text", 0)
            if !DllCall("comctl32\SetWindowSubclass", "ptr", handle, "ptr", callback, "uptr", 1, "uptr", 0)
                throw Error("Could not install the native paint probe.")
            probed.Push(handle)
        }
        DllCall("user32\RedrawWindow", "ptr", CPBigBoxGui.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x181)
        Sleep(70)
        TestAssert(TestPaintCounts[CPBigBoxControls["audioToggle"].Hwnd]["paint"] > 0,
            "Paint probe detects a deliberately forced whole-window repaint")
        for handle, counts in TestPaintCounts
            counts["paint"] := 0
        ; Invoke the final focus operation directly: the foreground guard must
        ; continue rejecting D-pad routing to this non-active off-screen GUI.
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["explanation"]))
        Sleep(70)
        for tile in CPBigBoxHomeTiles() {
            if tile[1] = "translation" || tile[1] = "explanation"
                continue
            TestAssert(TestPaintCounts[CPBigBoxControls[tile[1]].Hwnd]["paint"] = 0,
                "Focus navigation does not repaint unrelated Home tile: " tile[1])
        }
        TestAssert(CPBigBoxFocusFrame["key"] = "explanation", "Off-screen native navigation retains correct frame")
        CPBigBoxGui.Hide()
        CPBigBoxSetPage("explanation", false)
        CPBigBoxGui.Show("NA x-30000 y-30000 w1280 h720")
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ai_provider"]))
        Sleep(70)
        for handle, counts in TestPaintCounts
            counts["paint"] := 0
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ai_model"]))
        Sleep(70)
        for option in CPBigBoxExplanationOptions()
            TestAssert(TestPaintCounts[CPBigBoxControls["exp_" option].Hwnd]["paint"] = 0,
                "Focus navigation does not repaint unrelated Explanation option: " option)
        helpHandle := CPBigBoxControls["modeBody"].Hwnd
        TestPaintCounts[helpHandle]["text"] := 0
        CPBigBoxDashboardButtonFocused("ai_model")
        CPBigBoxDashboardButtonFocused("ai_model")
        TestAssert(TestPaintCounts[helpHandle]["text"] = 0,
            "Repeated focus notifications do not rewrite unchanged help text")
        CPBigBoxGui.Hide()
        CPBigBoxSetPage("screenshot", false)
        CPBigBoxGui.Show("NA x-30000 y-30000 w1280 h720")
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["shot_highlight"]))
        Sleep(70)
        for handle, counts in TestPaintCounts
            counts["paint"] := 0
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["shot_speakerColor"]))
        Sleep(70)
        for key in ["ai_provider", "ai_model", "ai_detail", "shot_openOnStartup", "shot_alwaysOnTop", "shot_clearOnStartup"]
            TestAssert(TestPaintCounts[CPBigBoxControls[key].Hwnd]["paint"] = 0,
                "Screenshot focus change does not repaint unrelated setting: " key)
        CPBigBoxGui.Hide()
        CPBigBoxSetPage("audio", false)
        CPBigBoxGui.Show("NA x-30000 y-30000 w1280 h720")
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["audio_device"]))
        Sleep(70)
        for handle, counts in TestPaintCounts
            counts["paint"] := 0
        CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["audio_refresh"]))
        Sleep(70)
        for key in ["ai_provider", "ai_model", "ai_detail", "audio_test", "advanced"]
            TestAssert(TestPaintCounts[CPBigBoxControls[key].Hwnd]["paint"] = 0,
                "Audio focus change does not repaint unrelated setting: " key)
        TestPaintCounts[CPBigBoxControls["audio_power"].Hwnd]["paint"] := 0
        CPBigBoxUpdateAudioPower()
        CPBigBoxUpdateAudioPower()
        Sleep(70)
        TestAssert(TestPaintCounts[CPBigBoxControls["audio_power"].Hwnd]["paint"] = 0,
            "Unchanged periodic audio status does not repaint its button")
        for overlayPage in ["translationWindow", "explanationWindow"] {
            CPBigBoxGui.Hide()
            CPBigBoxSetPage(overlayPage, false)
            CPBigBoxGui.Show("NA x-30000 y-30000 w1280 h720")
            CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_bg"]))
            Sleep(60)
            for handle, counts in TestPaintCounts
                counts["paint"] := 0
            CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["ov_txt"]))
            Sleep(60)
            for key in ["ov_font", "ov_size", "ov_bold", "ov_opacity", "ov_position"]
                TestAssert(TestPaintCounts[CPBigBoxControls[key].Hwnd]["paint"] = 0,
                    overlayPage " focus does not repaint unrelated tile: " key)
        }
    } finally {
        for handle in probed
            DllCall("comctl32\RemoveWindowSubclass", "ptr", handle, "ptr", callback, "uptr", 1)
        CallbackFree(callback)
        CPBigBoxGui.Hide()
    }
    CPBigBoxDashboardHide(false)
    for key in CPBigBoxFocusFrameKeys()
        TestAssert(!TestControlShown(CPBigBoxControls[key]), "Hiding dashboard clears focus decoration: " key)
    CPBigBoxRestorePageFocus()
    TestAssert(CPBigBoxFocusFrame["key"] != "", "Focus frame restores without creating a new GUI")
    CPBigBoxSetPage("home", false)
}

TestBigBoxTerminologyProfiles() {
    global CPBigBoxCurrentPage, CPBigBoxControls, CPBigBoxAIChoice, CPBigBoxGui
    global useTerminologyOverrides, chkUseTerminologyOverrides, iniPath
    global jp2enGlossaryProfile, en2enGlossaryProfile, glossariesDir
    global gameProfilesDir, CPBigBoxProfileState, TestProfileApplies, TestProfileSaves

    CPBigBoxGui.Show("Hide w1280 h720")
    CPBigBoxDashboardResize(CPBigBoxGui, 0, 1280, 720)
    CPBigBoxSetPage("terminology", false)
    TestLayout(1280, 720)
    CPBigBoxGui.Show("Hide w1920 h1080")
    CPBigBoxDashboardResize(CPBigBoxGui, 0, 1920, 1080)
    TestLayout(1920, 1080)
    TestCapture("terminology-settings-1920.png", 1920, 1080)
    TestAssert(CPBigBoxGroupedSettingsActive(), "Terminology is a complete grouped Big Box page")
    TestAssert(!TestControlShown(CPBigBoxControls["backHome"]),
        "Terminology page hides the unused Back to Home control")
    TestAssert(CPBigBoxControls["term_enabled"].Enabled && CPBigBoxControls["term_jp_manage"].Enabled,
        "Terminology controls are controller reachable")
    CPBigBoxUpdateSettingsHint("term_en_profile")
    TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "locally after translation")
        && InStr(CPBigBoxControls["modeBody"].Text, "never sent to the AI")
        && InStr(CPBigBoxControls["modeBody"].Text, "Esuteru → Estelle"),
        "TL to TL profile focus explains local replacement behavior and gives an example")
    CPBigBoxUpdateSettingsHint("term_jp_profile")
    TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "sent to the AI as instructions")
        && InStr(CPBigBoxControls["modeBody"].Text, "ignore or overapply")
        && InStr(CPBigBoxControls["modeBody"].Text, "エステル → Estelle"),
        "JP to TL profile focus explains model behavior, limitations and an example")
    CPBigBoxUpdateSettingsHint("term_en_tools")
    TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "independent")
        && InStr(CPBigBoxControls["modeBody"].Text, "leaves the other unchanged"),
        "Terminology profile tools explain that glossary types are independent")
    CPBigBoxUpdateSettingsHint("term_enabled")
    TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "next capture translation"),
        "Terminology toggle focus explains when changes take effect")
    CPBigBoxToggleTerminology()
    TestAssert(useTerminologyOverrides && chkUseTerminologyOverrides.Value
        && IniRead(iniPath, "cfg", "useTerminologyOverrides", 0) = "1",
        "Terminology toggle updates runtime, desktop control and INI")

    DirCreate(glossariesDir "\custom")
    FileAppend("# custom`r`nhero -> champion`r`n", glossariesDir "\custom\en2en.txt", "UTF-8")
    CPBigBoxOpenGlossaryProfileChoice("en")
    TestAssert(!TestControlShown(CPBigBoxControls["previewTitle"])
        && !TestControlShown(CPBigBoxControls["previewBody"]),
        "Terminology picker hides underlying preview guidance")
    customIndex := ArrayIndexOf(CPBigBoxAIChoice["options"], "custom")
    CPBigBoxCommitAIChoiceIndex(customIndex)
    TestAssert(en2enGlossaryProfile = "custom" && IniRead(iniPath, "cfg", "en2enGlossaryProfile", "") = "custom",
        "Terminology profile picker uses shared selection and persistence")
    TestAssert(CPBigBoxCurrentPage = "terminology" && !CPBigBoxAIChoiceActive(),
        "Terminology picker returns to its opening tile")
    TestAssert(!TestControlShown(CPBigBoxControls["backHome"]),
        "Returning from terminology picker does not reveal Back to Home")

    CPBigBoxOpenGlossaryEntries("jp")
    TestAssert(CPBigBoxAIChoiceActive() && CPBigBoxAIChoice["field"] = "glossaryEntry",
        "Terminology manager opens the modern entry list")
    TestAssert(!TestControlShown(CPBigBoxControls["previewTitle"])
        && !TestControlShown(CPBigBoxControls["previewBody"]),
        "Terminology entry list hides underlying preview guidance")
    CPBigBoxCommitAIChoiceIndex(1)
    CPBigBoxGui.Show("Hide w1280 h720")
    CPBigBoxDashboardResize(CPBigBoxGui, 0, 1280, 720)
    TestLayout(1280, 720)
    CPBigBoxGui.Show("Hide w1920 h1080")
    CPBigBoxDashboardResize(CPBigBoxGui, 0, 1920, 1080)
    TestLayout(1920, 1080)
    TestCapture("terminology-entry-1920.png", 1920, 1080)
    TestAssert(CPBigBoxCurrentPage = "manageEntry" && !CPBigBoxControls["mg_delete"].Enabled,
        "Add entry opens a modern non-destructive editor")
    CPBigBoxControls["mg_source"].Value := "魔王"
    CPBigBoxControls["mg_target"].Value := "demon king"
    CPBigBoxSaveGlossaryEntry()
    TestAssert(CPBigBoxAIChoiceActive() && InStr(FileRead(GlossaryPath("jp", "default"), "UTF-8"), "魔王 -> demon king"),
        "New terminology entry is atomically written and returns to the list")
    editIndex := 0
    for index, value in CPBigBoxAIChoice["options"]
        if InStr(value, "魔王") {
            editIndex := index
            break
        }
    CPBigBoxCommitAIChoiceIndex(editIndex)
    TestAssert(CPBigBoxControls["mg_delete"].Enabled, "Existing terminology entry enables Delete")
    CPBigBoxRequestGlossaryEntryDelete()
    TestAssert(CPBigBoxCurrentPage = "manageConfirm", "Entry deletion requires modern confirmation")
    CPBigBoxCancelManage()
    TestAssert(CPBigBoxCurrentPage = "manageEntry", "Cancelling entry deletion returns to the editor")
    CPBigBoxRequestGlossaryEntryDelete()
    CPBigBoxConfirmManage()
    TestAssert(CPBigBoxCurrentPage = "terminology"
        && !InStr(FileRead(GlossaryPath("jp", "default"), "UTF-8"), "魔王 -> demon king"),
        "Confirmed entry deletion removes only the selected pair")

    CPBigBoxOpenGlossaryTools("jp")
    TestLayout(1920, 1080)
    TestAssert(CPBigBoxCurrentPage = "manageTools" && !CPBigBoxControls["mg_delete"].Enabled,
        "Default glossary profile is protected from deletion")
    CPBigBoxStartGlossaryName()
    CPBigBoxControls["mg_name"].Value := "Boss/Names"
    CPBigBoxSaveManageName()
    TestAssert(jp2enGlossaryProfile = "Boss_Names" && FileExist(GlossaryPath("jp", "Boss_Names")),
        "New terminology profile sanitizes its name and becomes active")
    CPBigBoxOpenGlossaryTools("jp")
    TestAssert(CPBigBoxControls["mg_delete"].Enabled, "Non-default glossary profile can be deleted")
    CPBigBoxOpenGlossaryDelete("jp")
    CPBigBoxConfirmManage()
    TestAssert(jp2enGlossaryProfile = "default" && !FileExist(GlossaryPath("jp", "Boss_Names")),
        "Glossary deletion resets only that type to default")

    FileAppend("malformed line`r`n", GlossaryPath("jp", "default"), "UTF-8")
    CPBigBoxOpenGlossaryEntries("jp")
    TestAssert(CPBigBoxCurrentPage = "manageRaw" && InStr(CPBigBoxControls["mg_raw"].Value, "malformed line"),
        "Malformed terminology opens the modern raw repair editor without discarding text")
    CPBigBoxCancelManage()
    TestAssert(InStr(FileRead(GlossaryPath("jp", "default"), "UTF-8"), "malformed line"),
        "Cancelling raw repair leaves the source file unchanged")

    FileAppend("[profile]`r`nschemaVersion=1`r`nname=Other`r`n", gameProfilesDir "\Other.ini", "UTF-8")
    CPBigBoxSetPage("profiles", false)
    TestStartupOverlays()
    TestLayout(1920, 1080)
    TestCapture("profiles-settings-1920.png", 1920, 1080)
    TestAssert(CPBigBoxGroupedSettingsActive(), "Profiles is a complete grouped Big Box page")
    TestAssert(!TestControlShown(CPBigBoxControls["backHome"]),
        "Profiles page hides the unused Back to Home control")
    CPBigBoxOpenProfileChoice()
    TestAssert(!TestControlShown(CPBigBoxControls["previewTitle"])
        && !TestControlShown(CPBigBoxControls["previewBody"]),
        "Profile picker hides underlying preview guidance")
    otherIndex := ArrayIndexOf(CPBigBoxAIChoice["options"], "Other")
    CPBigBoxCommitAIChoiceIndex(otherIndex)
    TestAssert(CPBigBoxProfileState["selected"] = "Other"
        && IniRead(iniPath, "game_profiles", "active", "") = "Demo",
        "Selecting a profile does not apply it")
    TestAssert(!TestControlShown(CPBigBoxControls["backHome"]),
        "Returning from profile picker does not reveal Back to Home")
    CPBigBoxProfileAction("apply")
    TestAssert(TestProfileApplies = 1 && IniRead(iniPath, "game_profiles", "active", "") = "Other",
        "Apply invokes the shared profile workflow")
    CPBigBoxProfileAction("save")
    TestAssert(TestProfileSaves = 1 && FileExist(GameProfilePath("Other")),
        "Save Current invokes the shared profile workflow")
    CPBigBoxProfileAction("new")
    TestLayout(1920, 1080)
    TestCapture("profile-name-1920.png", 1920, 1080)
    CPBigBoxDashboardSetFocus(CPBigBoxDashboardControlIndex(CPBigBoxControls["mg_name"]))
    CPBigBoxDashboardActivate()
    TestAssert(CPBigBoxGui.FocusedCtrl.Hwnd = CPBigBoxControls["mg_save"].Hwnd,
        "Controller A on a text field advances to its confirmation action")
    CPBigBoxControls["mg_name"].Value := "New/Test"
    CPBigBoxSaveManageName()
    TestAssert(CPBigBoxProfileState["selected"] = "New_Test" && FileExist(GameProfilePath("New_Test")),
        "New game profile uses a safe name and current settings")
    CPBigBoxProfileAction("delete")
    TestLayout(1920, 1080)
    TestAssert(CPBigBoxCurrentPage = "manageConfirm" && FileExist(GameProfilePath("New_Test")),
        "Profile deletion waits for confirmation")
    CPBigBoxConfirmManage()
    TestAssert(CPBigBoxCurrentPage = "profiles" && !FileExist(GameProfilePath("New_Test")),
        "Confirmed profile deletion removes only the selected profile")
    CPBigBoxSetPage("home", false)
}

TestStartupOverlays() {
    global iniPath, chkOpenTW, chkOpenEW, ddlStartupOverlays, CPBigBoxControls
    global CPBigBoxManageNotice, TestExternalCalls, CPBigBoxGui
    original := CPStartupOverlayIndex()
    profile := GameProfilePath("Startup-test")
    calls := TestExternalCalls
    Loop 4 {
        index := A_Index
        CPSetStartupOverlays(index)
        GameProfileSaveStartup(profile)
        CPSetStartupOverlays(index = 4 ? 1 : 4)
        GameProfileApplyStartup(profile)
        TestAssert(CPStartupOverlayIndex() = index && ddlStartupOverlays.Value = index,
            "Profile startup combination round-trips and synchronizes the desktop choice: " index)
        TestAssert(IniRead(profile, "translator", "openOnLaunch") = chkOpenTW.Value
            && IniRead(profile, "explainer", "openOnLaunch") = chkOpenEW.Value,
            "Startup choices keep the existing schema-1 profile keys")
        CPSetStartupOverlays(index)
        plan := CPStartupOverlayPlan()
        TestAssert(!!plan["translator"] = !!chkOpenTW.Value
            && !!plan["explainer"] = !!chkOpenEW.Value,
            "Cold launch uses exactly the two profile startup settings")
        for studyMode in ["library", "reader"] {
            plan := CPStartupOverlayPlan(studyMode, true)
            TestAssert(!plan["translator"] && !plan["explainer"],
                "Standalone Study launches do not open overlays")
        }
        CPBigBoxOpenStartupOverlayChoice()
        CPBigBoxCommitAIChoiceIndex(index = 4 ? 1 : index + 1)
        TestAssert(CPStartupOverlayIndex() = (index = 4 ? 1 : index + 1),
            "Controller startup selector commits every combination")
        TestAssert(InStr(CPBigBoxManageNotice, "Save current settings"),
            "Startup selector explains how to save the choice per profile")
    }
    CPSetStartupOverlays(1)
    TestAssert(CPStartupOverlayPlan("", true)["translator"],
        "An explicit custom-launcher command-line override remains supported")
    CPBigBoxOpenStartupOverlayChoice()
    CPBigBoxCloseAIChoice()
    TestAssert(CPStartupOverlayIndex() = 1, "Cancel leaves the startup combination unchanged")
    CPBigBoxOpenStartupOverlayChoice()
    CPSetStartupOverlays(3)
    CPBigBoxCommitAIChoiceIndex(2)
    TestAssert(CPStartupOverlayIndex() = 3 && InStr(CPBigBoxManageNotice, "Settings changed"),
        "An outdated controller choice cannot replace newer startup settings")
    CPSetScreenshotPreference("openOnStartup", true)
    TestAssert(ddlStartupOverlays.Value = 4, "Original Translator checkbox synchronizes the combined choice")
    CPSetExplanationPreference("openOnStartup", false)
    TestAssert(ddlStartupOverlays.Value = 2, "Original Explainer checkbox synchronizes the combined choice")
    ddlStartupOverlays.Choose(3)
    CPStartupOverlaysChanged(ddlStartupOverlays)
    TestAssert(!chkOpenTW.Value && chkOpenEW.Value, "Desktop combined choice updates both original checkboxes")
    FileDelete(profile)
    GameProfileApplyStartup(profile)
    TestAssert(CPStartupOverlayIndex() = 3, "Missing legacy startup keys preserve the current choices")
    IniWrite(1, profile, "translator", "openOnLaunch")
    GameProfileApplyStartup(profile)
    TestAssert(CPStartupOverlayIndex() = 4, "A partial old profile preserves the missing Explainer choice")
    savedIni := iniPath
    failed := false
    try {
        iniPath := A_ScriptDir ; A directory cannot be used as an INI file.
        CPSetStartupOverlays(1)
    } catch {
        failed := true
    } finally {
        iniPath := savedIni
    }
    TestAssert(failed && CPStartupOverlayIndex() = 4 && ddlStartupOverlays.Value = 4,
        "A failed settings write preserves the visible startup choices")
    for dimension in [[1280, 720], [1920, 1080], [3840, 2160]] {
        CPBigBoxGui.Show("Hide w" dimension[1] " h" dimension[2])
        CPBigBoxDashboardResize(CPBigBoxGui, 0, dimension[1], dimension[2])
        TestLayout(dimension[1], dimension[2])
        TestCapture("profile-startup-" dimension[1] ".png", dimension[1], dimension[2])
        TestAssert(TestControlShown(CPBigBoxControls["prof_startup"]),
            "Startup overlays tile is visible at " dimension[1])
    }
    TestAssert(TestExternalCalls = calls, "Editing startup choices does not launch or stop overlays")
    CPSetStartupOverlays(original)
    FileDelete(profile)
    CPBigBoxManageNotice := ""
    CPBigBoxGui.Show("Hide w1920 h1080")
    CPBigBoxDashboardResize(CPBigBoxGui, 0, 1920, 1080)
}

TestPaintProbe(hwnd, msg, wParam, lParam, subclassId, refData) {
    global TestPaintCounts
    if TestPaintCounts.Has(hwnd) {
        if msg = 0x000F
            TestPaintCounts[hwnd]["paint"] += 1
        if msg = 0x000C
            TestPaintCounts[hwnd]["text"] += 1
    }
    return DllCall("comctl32\DefSubclassProc", "ptr", hwnd, "uint", msg,
        "ptr", wParam, "ptr", lParam, "ptr")
}

TestFocusFrameBounds() {
    global CPBigBoxControls, CPBigBoxFocusFrame
    CPBigBoxControls[CPBigBoxFocusFrame["key"]].GetPos(&x, &y, &w, &h)
    for key in CPBigBoxFocusFrameKeys() {
        control := CPBigBoxControls[key]
        control.GetPos(&frameX, &frameY, &frameW, &frameH)
        TestAssert(TestControlShown(control) && !control.Enabled, "Focus frame is visible but non-interactive: " key)
        TestAssert(frameX + frameW <= x || frameX >= x + w || frameY + frameH <= y || frameY >= y + h,
            "Focus frame never covers the button: " key)
        TestAssert(Min(frameW, frameH) >= 3, "Focus border remains at least three pixels thick: " key)
    }
}
TestSnapshot(token := "") {
    return Map("name", "XInput controller 1", "tokens", token = "" ? Map() : Map(token, true))
}
TestControlShown(control) {
    ; Read the child's own visibility flag even when its test GUI is hidden.
    return (DllCall("user32\GetWindowLongW", "ptr", control.Hwnd, "int", -16, "uint") & 0x10000000) != 0
}
TestHomeHeaderContent() {
    global CPBigBoxControls, CPBigBoxPageFocus, CPBigBoxActionNotice
    CPBigBoxDashboardUpdateContent()
    TestAssert(CPBigBoxControls["status_translation_model"].Text = "Gemini · gemini-test",
        "Header shows the effective screenshot model")
    TestAssert(CPBigBoxControls["status_explanation_model"].Text = "OpenAI · gpt-test",
        "Header shows the effective explanation model")
    TestAssert(CPBigBoxControls["status_audio_model"].Text = "Gemini · gemini-audio-test",
        "Header shows the effective audio model")
    TestAssert(CPBigBoxControls["status_translation_detail"].Text = "Prompt: default"
        && CPBigBoxControls["status_explanation_detail"].Text = "Prompt: default"
        && CPBigBoxControls["status_audio_detail"].Text = "Language: English",
        "Header includes both selected prompts and the audio language")
    for key, label in Map("translation", "Translation AI", "explanation", "Explanation AI",
        "audioAI", "Audio AI", "overlays", "Overlay Windows", "study", "Study Library",
        "controls", "Controller Settings")
        TestAssert(CPBigBoxControls[key].Text = label, "Home label is concise: " key)
    TestAssert(CPBigBoxControls["audioToggle"].Text = "Audio Translation`nOff"
        && CPBigBoxControls["capture"].Text = "Capture…`nRegion",
        "Home retains useful short audio and capture states")
    TestAssert(FileExist(CPBigBoxBrandLogoPath()), "Big Box brand asset is bundled with source")
    CPBigBoxActionNotice := "Test action feedback"
    CPBigBoxUpdateHomeHint("audioToggle")
    TestAssert(CPBigBoxControls["modeBody"].Text = CPBigBoxActionNotice,
        "Contextual Home hints retain audio action feedback on its tile")
    CPBigBoxUpdateHomeHint("translation")
    TestAssert(InStr(CPBigBoxControls["modeBody"].Text, "Prompt: default"),
        "Moving to an AI tile reveals its details even after audio feedback")
    CPBigBoxActionNotice := ""
    for key, expected in Map("translation", "Prompt: default", "explanation", "OpenAI · gpt-test",
        "audioAI", "Language: English", "audioToggle", "off. Select to start.",
        "study", "Active library:", "controls", "keyboard shortcuts") {
        CPBigBoxUpdateHomeHint(key)
        TestAssert(InStr(CPBigBoxControls["modeBody"].Text, expected), "Home contextual help: " key)
    }
    CPBigBoxPageFocus["home"] := "translation"
    CPBigBoxUpdateHomeHint()
}

TestHomeHeaderLongNames(width, height) {
    global CPBigBoxControls, imgProvider, imgModel, geminiImgModel, promptProfile
    global explainProvider, explainPromptProfile, audioProvider, audioTargetLang
    saved := [imgProvider, imgModel, geminiImgModel, promptProfile,
        explainProvider, explainPromptProfile, audioProvider, audioTargetLang]
    try {
        imgProvider := "openai"
        imgModel := "long-model-name-to-check-header-ellipsis-without-resizing-tiles"
        promptProfile := "default_with_kanji_reading_en_and_additional_custom_guidance"
        explainProvider := "gemini"
        explainPromptProfile := "detailed_explanation_with_vocabulary_and_readings"
        audioProvider := "openai"
        audioTargetLang := "German"
        CPBigBoxDashboardUpdateContent()
        CPBigBoxUpdateHomeHint("translation")
        TestAssert(CPBigBoxControls["status_translation_model"].Text = "OpenAI · " imgModel
            && CPBigBoxControls["status_translation_detail"].Text = "Prompt: " promptProfile,
            "Header refresh follows changed model and prompt without truncating stored labels")
        TestAssert(CPBigBoxControls["status_explanation_model"].Text = "Gemini · gemini-test"
            && CPBigBoxControls["status_explanation_detail"].Text = "Prompt: " explainPromptProfile
            && CPBigBoxControls["status_audio_model"].Text = "OpenAI · audio-test"
            && CPBigBoxControls["status_audio_detail"].Text = "Language: German",
            "Header refresh follows all three providers, explanation prompt and audio language")
        TestAssert(InStr(CPBigBoxControls["modeBody"].Text, imgModel)
            && InStr(CPBigBoxControls["modeBody"].Text, promptProfile),
            "Focused Home action exposes full long model and prompt names")
        TestLayout(width, height)
        TestCapture("home-long-names-" width ".png", width, height)
    } finally {
        imgProvider := saved[1], imgModel := saved[2], geminiImgModel := saved[3], promptProfile := saved[4]
        explainProvider := saved[5], explainPromptProfile := saved[6], audioProvider := saved[7], audioTargetLang := saved[8]
        CPBigBoxDashboardUpdateContent()
    }
}

TestLayout(width, height) {
    global CPBigBoxControls, CPBigBoxNavigationControls, CPBigBoxCurrentPage
    for control in CPBigBoxNavigationControls {
        control.GetPos(&x, &y, &w, &h)
        TestAssert(x >= 0 && y >= 0 && x + w <= width && y + h <= height,
            "Navigation control within " width "x" height ": " control.Text)
    }
    CPBigBoxControls["gameTitle"].GetPos(, &titleY,, &titleH)
    CPBigBoxControls["gamePlatform"].GetPos(, &platformY)
    TestAssert(titleY + titleH <= platformY, "Game title does not overlap platform")
    CPBigBoxControls["artwork"].GetPos(,, &artW, &artH)
    CPBigBoxControls["gameLogo"].GetPos(,, &logoW, &logoH)
    TestAssert(Abs(artW / artH - 180 / 300) < 0.02, "Tall box art keeps its aspect ratio")
    TestAssert(Abs(logoW / logoH - 400 / 90) < 0.1, "Wide logo keeps its aspect ratio")
    CPBigBoxControls["brandLogo"].GetPos(&brandX, &brandY, &brandW, &brandH)
    CPBigBoxControls["title"].GetPos(&headingX, &headingY,, &headingH)
    CPBigBoxControls["panel"].GetPos(, &panelY)
    CPBigBoxControls["artFrame"].GetPos(&artFrameX)
    TestAssert(TestControlShown(CPBigBoxControls["brandLogo"])
        && Abs(brandW / brandH - 434 / 500) < 0.02,
        "Big Box brand logo is visible and keeps its aspect ratio")
    TestAssert(brandX + brandW < headingX && brandY + brandH <= headingY + headingH,
        "Brand logo stays beside the heading without overlap")
    previousRowBottom := headingY + headingH
    for key in ["translation", "explanation", "audio"] {
        previousRight := 0
        rowY := 0
        for field in ["label", "model", "detail"] {
            control := CPBigBoxControls["status_" key "_" field]
            control.GetPos(&sx, &sy, &sw, &sh)
            TestAssert(sx >= previousRight && sx + sw < artFrameX && sy >= previousRowBottom
                && sy + sh < panelY && (!rowY || sy = rowY),
                "Aligned header status remains above tiles and left of game card: " key " " field)
            TestAssert(TestControlShown(control) && TestFontHeight(control) <= sh,
                "Header status text is visible and fits its row: " key " " field)
            style := DllCall("user32\GetWindowLongW", "ptr", control.Hwnd, "int", -16, "uint")
            TestAssert((style & 0x4200) = 0x4200 && !(style & 0x10000),
                "Header status ellipsizes long text, centers it vertically and adds no tab stop")
            previousRight := sx + sw
            rowY := sy
        }
        previousRowBottom := rowY + sh
    }
    CPBigBoxControls["modeBody"].GetPos(, &bodyY,, &bodyH)
    TestAssert(TestTextHeight(CPBigBoxControls["modeBody"]) <= bodyH,
        "Status/help text fits vertically at " width "x" height)
    CPBigBoxControls["advanced"].GetPos(, &bottomY)
    if (CPBigBoxCurrentPage = "home") {
        CPBigBoxControls["translation"].GetPos(, &row1Y,, &row1H)
        CPBigBoxControls["capture"].GetPos(, &row2Y,, &row2H)
        TestAssert(bodyY + bodyH <= row1Y && row1Y + row1H <= row2Y
            && row2Y + row2H <= bottomY, "Home rows do not overlap")
    } else if CPBigBoxCurrentPage = "quickCapture" {
        CPBigBoxControls["cap_region"].GetPos(, &captureY,, &captureH)
        CPBigBoxControls["cap_note"].GetPos(, &captureNoteY,, &captureNoteH)
        TestAssert(bodyY + bodyH <= captureY && captureY + captureH <= captureNoteY
            && captureNoteY + captureNoteH <= bottomY, "Capture controls and summary do not overlap")
        TestAssert(TestTextHeight(CPBigBoxControls["cap_note"]) <= captureNoteH, "Capture summary fits vertically")
    } else if CPBigBoxCurrentPage = "captureLimit" {
        CPBigBoxControls["cap_value"].GetPos(, &valueY,, &valueH)
        CPBigBoxControls["cap_plus100"].GetPos(, &adjustY,, &adjustH)
        CPBigBoxControls["cap_save"].GetPos(, &saveY)
        TestAssert(bodyY + bodyH <= valueY && valueY + valueH <= adjustY && adjustY + adjustH <= saveY,
            "PNG adjustment rows do not overlap")
        TestAssert(TestTextHeight(CPBigBoxControls["cap_value"]) <= valueH, "PNG value font fits vertically")
    } else if CPBigBoxCurrentPage = "overlayEdit" {
        CPBigBoxControls["ov_save"].GetPos(, &saveY)
        CPBigBoxControls["ov_preview"].GetPos(&previewX, &previewY, &previewW, &previewH)
        TestAssert(previewY >= bodyY + bodyH && previewY + previewH <= saveY, "Overlay preview fits between help and Save")
        previousBottom := bodyY + bodyH
        Loop 3 {
            if !CPBigBoxControls["ov_slider" A_Index].Enabled
                continue
            CPBigBoxControls["ov_label" A_Index].GetPos(&labelX, &labelY,, &labelH)
            CPBigBoxControls["ov_slider" A_Index].GetPos(&sliderX, &sliderY,, &sliderH)
            TestAssert(labelY >= previousBottom && labelY + labelH <= sliderY && sliderY + sliderH <= saveY,
                "Overlay editor rows do not overlap")
            TestAssert(labelX > previewX + previewW && TestTextHeight(CPBigBoxControls["ov_label" A_Index]) <= labelH,
                "Overlay slider label fits beside preview")
            previousBottom := sliderY + sliderH
        }
    } else if CPBigBoxCurrentPage = "quickOverlays" || CPBigBoxOverlayQuick() {
        key := CPBigBoxOverlayQuick() ? "ov_bg" : "ov_translator"
        CPBigBoxControls[key].GetPos(, &quickY,, &quickH)
        TestAssert(quickY >= bodyY + bodyH && quickY + quickH <= bottomY, "Overlay quick actions fit")
        if CPBigBoxOverlayQuick() {
            CPBigBoxControls["ov_note"].GetPos(, &noteY,, &noteH)
            TestAssert(quickY + quickH <= noteY && noteY + noteH <= bottomY && TestTextHeight(CPBigBoxControls["ov_note"]) <= noteH,
                "Overlay quick-view help fits below its tiles")
        }
    } else if CPBigBoxCurrentPage = "quickControls" && !CPBigBoxAIChoiceActive() {
        CPBigBoxControls["ctrl_controller"].GetPos(, &controlY,, &controlH)
        CPBigBoxControls["ctrl_status"].GetPos(, &statusY,, &statusH)
        TestAssert(controlY >= bodyY + bodyH && controlY + controlH <= statusY
            && statusY + statusH <= bottomY, "Quick Button Configuration rows do not overlap")
        TestAssert(TestTextHeight(CPBigBoxControls["ctrl_status"]) <= statusH,
            "Quick controller status fits vertically")
    } else if CPBigBoxCurrentPage = "controlDetail" {
        CPBigBoxControls["ctrl_primary"].GetPos(, &detailY,, &detailH)
        CPBigBoxControls["ctrl_back"].GetPos(, &detailBackY,, &detailBackH)
        TestAssert(detailY >= bodyY + bodyH && detailY + detailH <= detailBackY
            && detailBackY + detailBackH <= bottomY + detailBackH,
            "Binding detail rows do not overlap")
    } else if CPBigBoxCurrentPage = "controlHotkey" {
        CPBigBoxControls["ctrl_hotkey"].GetPos(, &hotkeyY,, &hotkeyH)
        CPBigBoxControls["ctrl_save"].GetPos(, &hotkeySaveY)
        TestAssert(hotkeyY >= bodyY + bodyH && hotkeyY + hotkeyH <= hotkeySaveY,
            "Hotkey field fits above Save and Cancel")
    } else if CPBigBoxCurrentPage = "controlCapture" {
        CPBigBoxControls["ctrl_status"].GetPos(, &captureStatusY,, &captureStatusH)
        CPBigBoxControls["ctrl_cancel"].GetPos(, &captureCancelY)
        TestAssert(captureStatusY >= bodyY + bodyH && captureStatusY + captureStatusH <= captureCancelY,
            "Controller capture instructions fit above Cancel")
    } else if CPBigBoxCurrentPage = "controlConflict" {
        CPBigBoxControls["ctrl_move"].GetPos(, &conflictY,, &conflictH)
        TestAssert(conflictY >= bodyY + bodyH && conflictY + conflictH <= bottomY + conflictH,
            "Conflict choices fit below their explanation")
    } else if CPBigBoxCurrentPage = "manageTools" {
        CPBigBoxControls["mg_new"].GetPos(, &manageY,, &manageH)
        CPBigBoxControls["mg_cancel"].GetPos(, &cancelY)
        TestAssert(manageY >= bodyY + bodyH && manageY + manageH <= cancelY,
            "Terminology profile tools fit above Back")
    } else if CPBigBoxCurrentPage = "manageName" {
        CPBigBoxControls["mg_label1"].GetPos(, &labelY,, &labelH)
        CPBigBoxControls["mg_name"].GetPos(, &editY,, &editH)
        CPBigBoxControls["mg_save"].GetPos(, &saveY)
        TestAssert(labelY >= bodyY + bodyH && labelY + labelH <= editY
            && editY + editH <= saveY, "Profile name editor fits above its actions")
    } else if CPBigBoxCurrentPage = "manageEntry" {
        CPBigBoxControls["mg_source"].GetPos(, &sourceY,, &sourceH)
        CPBigBoxControls["mg_label2"].GetPos(, &targetLabelY,, &targetLabelH)
        CPBigBoxControls["mg_target"].GetPos(, &targetY,, &targetH)
        CPBigBoxControls["mg_save"].GetPos(, &saveY)
        TestAssert(sourceY >= bodyY + bodyH && sourceY + sourceH <= targetLabelY
            && targetLabelY + targetLabelH <= targetY && targetY + targetH <= saveY,
            "Terminology entry fields fit above their actions")
    } else if CPBigBoxCurrentPage = "manageRaw" {
        CPBigBoxControls["mg_raw"].GetPos(, &rawY,, &rawH)
        CPBigBoxControls["mg_save"].GetPos(, &saveY)
        TestAssert(rawY >= bodyY + bodyH && rawY + rawH <= saveY,
            "Raw terminology editor fits above Save and Cancel")
    } else if CPBigBoxCurrentPage = "manageConfirm" {
        CPBigBoxControls["mg_confirm"].GetPos(, &confirmY,, &confirmH)
        TestAssert(confirmY >= bodyY + bodyH && confirmY + confirmH <= bottomY + confirmH,
            "Modern confirmation actions fit below their explanation")
    } else if CPBigBoxAIListActive() {
        listBottom := bodyY + bodyH
        for row, key in CPBigBoxAIListKeys() {
            CPBigBoxControls[key].GetPos(&listX, &listY, &listW, &listH)
            TestAssert(listY >= listBottom && listX >= 0 && listX + listW <= width,
                "Virtual list row fits without overlap: " row)
            TestAssert(TestFontHeight(CPBigBoxControls[key]) <= listH, "List font fits row " row)
            listBottom := listY + listH
        }
        CPBigBoxControls["choiceBack"].GetPos(, &cancelY)
        TestAssert(listBottom <= cancelY, "List ends before Cancel")
        TestAssert(TestFontHeight(CPBigBoxControls["listChoice3"]) > TestFontHeight(CPBigBoxControls["listChoice2"])
            && TestFontHeight(CPBigBoxControls["listChoice2"]) > TestFontHeight(CPBigBoxControls["listChoice1"]),
            "Focus lens uses progressively smaller fonts: " TestFontHeight(CPBigBoxControls["listChoice3"]) "/"
                . TestFontHeight(CPBigBoxControls["listChoice2"]) "/" TestFontHeight(CPBigBoxControls["listChoice1"]))
    } else if CPBigBoxAIChoiceActive() {
        CPBigBoxControls["choice1"].GetPos(, &firstY,, &firstH)
        CPBigBoxControls["choice3"].GetPos(, &secondY,, &secondH)
        CPBigBoxControls["choiceBack"].GetPos(, &cancelY)
        TestAssert(bodyY + bodyH <= firstY && firstY + firstH <= secondY
            && secondY + secondH <= cancelY, "Choice rows do not overlap")
    } else if CPBigBoxGroupedSettingsActive() {
        previousBottom := bodyY + bodyH
        for groupIndex, group in CPBigBoxSettingsGroups() {
            CPBigBoxControls["settingsGroup" groupIndex].GetPos(&labelX, &labelY, &labelW, &labelH)
            TestAssert(TestTextHeight(CPBigBoxControls["settingsGroup" groupIndex]) <= labelH,
                "Group caption text is fully visible")
            for key in group[2] {
                CPBigBoxControls[key].GetPos(&settingX, &settingY, &settingW, &settingH)
                TestAssert(settingY >= previousBottom && settingY + settingH <= bottomY,
                    "Grouped settings rows do not overlap: " key)
                TestAssert(labelX + labelW < settingX && labelY >= settingY && labelY + labelH <= settingY + settingH,
                    "Group caption fits beside its row: " key)
            }
            previousBottom := settingY + settingH
            if groupIndex < 3 {
                CPBigBoxControls["settingsDivider" groupIndex].GetPos(, &dividerY,, &dividerH)
                TestAssert(dividerY >= previousBottom && dividerY + dividerH < bottomY,
                    "Divider is between groups, not over a button")
                previousBottom := dividerY + dividerH
            }
        }
        TestAssert(!TestControlShown(CPBigBoxControls["aiNote"]), "Grouped settings page hides the quick-view note")
        for key in CPBigBoxSettingsKeys() {
            CPBigBoxControls[key].GetPos(,, &settingW, &settingH)
            TestAssert(TestFontHeight(CPBigBoxControls[key]) * 2 + 8 <= settingH,
                "Settings tile has room for label and state: " key)
        }
    } else if CPBigBoxAIDomain() != "" {
        CPBigBoxControls["ai_provider"].GetPos(, &firstY,, &firstH)
        CPBigBoxControls["aiNote"].GetPos(, &noteY,, &noteH)
        TestAssert(bodyY + bodyH <= firstY && firstY + firstH <= noteY
            && noteY + noteH <= bottomY, "AI settings rows do not overlap")
    } else {
        CPBigBoxControls["previewTitle"].GetPos(, &previewY,, &previewH)
        CPBigBoxControls["previewBody"].GetPos(, &textY,, &textH)
        CPBigBoxControls["backHome"].GetPos(, &backY,, &backH)
        TestAssert(bodyY + bodyH <= previewY && previewY + previewH <= textY
            && textY + textH <= backY && backY + backH <= bottomY, "Preview rows do not overlap")
    }
}
TestTextHeight(control) {
    control.GetPos(,, &width)
    dc := DllCall("user32\GetDC", "ptr", control.Hwnd, "ptr")
    oldFont := DllCall("gdi32\SelectObject", "ptr", dc,
        "ptr", SendMessage(0x0031, 0, 0, control.Hwnd), "ptr")
    rect := Buffer(16, 0)
    NumPut("int", width, rect, 8)
    try {
        DllCall("user32\DrawTextW", "ptr", dc, "wstr", control.Text, "int", -1,
            "ptr", rect, "uint", 0x400 | 0x10 | 0x800) ; CALCRECT | WORDBREAK | NOPREFIX
        return NumGet(rect, 12, "int")
    } finally {
        DllCall("gdi32\SelectObject", "ptr", dc, "ptr", oldFont)
        DllCall("user32\ReleaseDC", "ptr", control.Hwnd, "ptr", dc)
    }
}
TestFontHeight(control) {
    fontHandle := SendMessage(0x0031, 0, 0, control.Hwnd)
    fontInfo := Buffer(92, 0)
    DllCall("gdi32\GetObjectW", "ptr", fontHandle, "int", fontInfo.Size, "ptr", fontInfo)
    return Abs(NumGet(fontInfo, 0, "int"))
}
TestCapture(name, width, height) {
    global CPBigBoxGui
    ; WM_PRINT renders only this synthetic test GUI, never the user's desktop.
    screenDC := DllCall("user32\GetDC", "ptr", CPBigBoxGui.Hwnd, "ptr")
    dc := DllCall("gdi32\CreateCompatibleDC", "ptr", screenDC, "ptr")
    bitmap := DllCall("gdi32\CreateCompatibleBitmap", "ptr", screenDC, "int", width, "int", height, "ptr")
    previous := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", bitmap, "ptr")
    SendMessage(0x0317, dc, 0x1C, CPBigBoxGui.Hwnd) ; WM_PRINT + ERASEBKGND + CLIENT + CHILDREN
    gdiplus := DllCall("kernel32\LoadLibraryW", "wstr", "gdiplus.dll", "ptr")
    startup := Buffer(24, 0), token := 0, image := 0, encoder := Buffer(16, 0)
    NumPut("uint", 1, startup)
    DllCall("gdiplus\GdiplusStartup", "ptr*", &token, "ptr", startup, "ptr", 0)
    DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "ptr", bitmap, "ptr", 0, "ptr*", &image)
    DllCall("ole32\CLSIDFromString", "wstr", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "ptr", encoder)
    DllCall("gdiplus\GdipSaveImageToFile", "ptr", image, "wstr", A_ScriptDir "\" name, "ptr", encoder, "ptr", 0)
    DllCall("gdiplus\GdipDisposeImage", "ptr", image)
    DllCall("gdiplus\GdiplusShutdown", "ptr", token)
    DllCall("kernel32\FreeLibrary", "ptr", gdiplus)
    DllCall("gdi32\SelectObject", "ptr", dc, "ptr", previous)
    DllCall("gdi32\DeleteObject", "ptr", bitmap)
    DllCall("gdi32\DeleteDC", "ptr", dc)
    DllCall("user32\ReleaseDC", "ptr", CPBigBoxGui.Hwnd, "ptr", screenDC)
}

; Read-only / no-op external services. The production navigation code below
; is copied verbatim by the test runner, including controller edge handling.
CPDesktopSyncModel(*) => 0 ; desktop shell is exercised by test_desktop_layout.ps1
CPDesktopSyncAudioModel(*) => 0
CPDesktopSyncExplanationModel(*) => 0
CPDesktopRefreshAudio(*) => 0
CPDialogDefaultOwner(*) => 0
CPDialogPresentation(*) => "fullscreen"
CPAdaptiveOwnedMessage(*) => "OK"
SyncUnifiedWindowAppearance(*) => 0
SyncPromptPostproc(name) => name = "literal" ? "literal" : "test"
SetDebugMode(*) => 0
ModelListWrite(key, arr) {
    global TestModelWrites, TestModelWriteFails
    if TestModelWriteFails
        throw Error("Synthetic model-list write failure")
    ModelListSort(arr)
    TestModelWrites.Push([key, arr.Clone()])
}
CPModelCatalogJobStart(provider, purpose, forceRefresh, callback) {
    global TestCatalogRequests, TestCatalogDeferred, TestCatalogBusy
    global TestCatalogCallback, TestCatalogResult
    TestCatalogRequests.Push([provider, purpose, forceRefresh])
    if TestCatalogDeferred {
        TestCatalogBusy := true
        TestCatalogCallback := callback
    } else
        callback.Call(TestCatalogResult)
    return true
}
CPModelCatalogJobCancel(*) {
    global TestCatalogBusy, TestCatalogCancels, TestCatalogCallback
    if TestCatalogBusy
        TestCatalogCancels += 1
    TestCatalogBusy := false
    TestCatalogCallback := 0
}
AudioIsRunning(*) {
    global TestLiveAudioRunning, gPidAudio
    if gPidAudio && ProcessExist(gPidAudio)
        return true
    return TestLiveAudioRunning
}
CPStartBigBoxAudio() {
    global TestAudioStarts, TestLiveAudioRunning, TestAudioStartOK
    TestAudioStarts += 1
    TestObserveAudioBusy()
    TestLiveAudioRunning := TestAudioStartOK
    CPBigBoxSetActionNotice(TestAudioStartOK ? "Audio Translation is on." : "Synthetic startup error; nothing started.")
    return TestAudioStartOK
}
CPApiKeyConfigured(provider) {
    global TestAudioApiConfigured
    return TestAudioApiConfigured
}
CPShowMissingApiKey(*) {
    throw Error("Big Box startup must not launch a desktop error workflow.")
}
AudioTargetCode(*) => "en"
AudioTargetName(name) => name
UpdateStatus(*) => CPBigBoxUpdateAudioPower()
Toast(message) {
    global TestToasts
    TestToasts.Push(message)
}
FilterPythonStderr(text) => text
AudioPidsByScript() {
    global TestAudioRecoveryScans
    TestAudioRecoveryScans += 1
    return []
}
CPStopBigBoxAudio() {
    global TestAudioStops, TestLiveAudioRunning
    TestAudioStops += 1
    TestObserveAudioBusy()
    TestLiveAudioRunning := false
    return true
}
EnsureTranslatorForCapture() {
    global TestTranslatorReady
    return TestTranslatorReady
}
CPSendBigBoxCaptureCommand(payload) {
    global TestCaptureSends, TestCaptureSendOK
    TestCaptureSends.Push(payload)
    return TestCaptureSendOK
}
CPResumeBigBoxAfterCapture(returnHwnd) {
    global TestCaptureResumes
    TestCaptureResumes += 1
    CPBigBoxRestorePageFocus()
    CPControllerResetNavigation()
}
ResolvePath(path) => path
CPRegisterColorSwatch(control, *) => control
SendOverlayTheme(title := "", timeoutMs := 0) {
    global TestOverlayThemes, TestOverlayThemeOK
    TestOverlayThemes.Push([title, timeoutMs])
    return TestOverlayThemeOK
}
CPEnsureOverlayForAdjustment(*) {
    global TestOverlayGui, TestOverlayReady
    return TestOverlayReady ? TestOverlayGui.Hwnd : 0
}
CPScanOverlayAdjustControllers() => []
CPInitializeOverlayAdjustButtonStates() => 0
CPOverlayAdjustModifierKey(key) => key
CPOverlayAdjustFlag(*) => 0
CPRegisterOverlayAdjustHotkeys() => 0
CPCreateOverlayAdjustHud() => 0
CPOverlayAdjustTick(*) => 0
SendOverlayCmdTo(title, payload) {
    global TestOverlaySaves
    if payload = "action=save_bounds"
        TestOverlaySaves += 1
}
SaveExplainerBoundsIfChanged(*) => 0
CPRestoreControlPanelAfterAdjustment(*) {
    throw Error("Big Box adjustment must not reopen the desktop control panel.")
}
CPResumeBigBoxOverlayDashboard(returnHwnd) {
    global TestOverlayReturns
    TestOverlayReturns += 1
    CPBigBoxRestorePageFocus()
    CPControllerResetNavigation()
}
StudyLibraryConfiguredName(*) => "Demo"
RefreshGlossaryProfilesList(selJP := "", selEN := "") {
    global ddlJPG, ddlENG, jp2enGlossaryProfile, en2enGlossaryProfile
    jp2enGlossaryProfile := selJP != "" ? selJP : jp2enGlossaryProfile
    en2enGlossaryProfile := selEN != "" ? selEN : en2enGlossaryProfile
    ddlJPG.Delete(), ddlJPG.Add(ListGlossaryProfiles("jp")), ddlJPG.Text := jp2enGlossaryProfile
    ddlENG.Delete(), ddlENG.Add(ListGlossaryProfiles("en")), ddlENG.Text := en2enGlossaryProfile
}
RefreshGameProfilesList(select := "") {
    global ddlGameProfile, iniPath
    options := ListGameProfiles()
    ddlGameProfile.Delete()
    if options.Length {
        ddlGameProfile.Add(options)
        ddlGameProfile.Text := select != "" ? select : options[1]
    } else
        ddlGameProfile.Text := ""
}
GameProfileUpdateSummary(*) => 0
GameProfileSave(name, announce := true) {
    global TestProfileSaves, GameProfileLastError
    name := GameProfileSafeName(name)
    if name = "" {
        GameProfileLastError := "Please enter a non-empty profile name."
        return false
    }
    TestProfileSaves += 1
    if FileExist(GameProfilePath(name))
        FileDelete(GameProfilePath(name))
    FileAppend("[profile]`r`nschemaVersion=1`r`nname=" name "`r`n", GameProfilePath(name), "UTF-8")
    return true
}
GameProfileApply(name, announce := true) {
    global TestProfileApplies, GameProfileLastError, iniPath
    if !FileExist(GameProfilePath(name)) {
        GameProfileLastError := "Profile not found."
        return false
    }
    TestProfileApplies += 1
    IniWrite(name, iniPath, "game_profiles", "active")
    return true
}
IniWriteRetry(value, path, section, key) => IniWrite(value, path, section, key)
CPOverlayWindowHwnd(title) {
    global TestOverlayWindows
    if !TestOverlayWindows.Has(title)
        return 0
    overlayHwnd := TestOverlayWindows[title]
    return DllCall("user32\IsWindow", "ptr", overlayHwnd, "int") ? overlayHwnd : 0
}
CPControllerSetEnabled(enabled, persist := true) {
    global CPControllerInputsEnabled, cbControllerInputsEnabled, iniPath
    CPControllerInputsEnabled := enabled ? true : false
    cbControllerInputsEnabled.Value := CPControllerInputsEnabled ? 1 : 0
    if persist
        IniWrite(CPControllerInputsEnabled ? 1 : 0, iniPath, "controller_inputs", "enabled")
}
CPControllerSetDpadNavigationEnabled(enabled, persist := true) {
    global CPControllerDpadNavigationEnabled, cbControllerDpadNavigationEnabled, iniPath
    CPControllerDpadNavigationEnabled := enabled ? true : false
    cbControllerDpadNavigationEnabled.Value := CPControllerDpadNavigationEnabled ? 1 : 0
    if persist
        IniWrite(CPControllerDpadNavigationEnabled ? 1 : 0, iniPath, "controller_inputs", "dpad_navigation")
    CPControllerResetNavigation()
}
CPControllerReadSnapshot() {
    global TestControllerSnapshots
    return TestControllerSnapshots.Length
        ? TestControllerSnapshots.RemoveAt(1)
        : Map("connected", false, "name", "", "tokens", Map())
}
Rebind_LaunchExplainerRequest(*) => TestCountRebind()
Rebind_ExplainLastTranslation(*) => TestCountRebind()
Rebind_StartStopAudio(*) => TestCountRebind()
Rebind_HideShowControlPanel(*) => TestCountRebind()
Hotkeys_ShowConflicts(*) => 0
TestCountRebind() {
    global TestRebinds
    TestRebinds += 1
}
CPSetPreferredAppDarkMode(*) => 0
CPAllowDarkModeForWindow(*) => 0
CPApplyThemeToControl(*) => 0
CPControllerKeyboardMirrorActive(*) => false
CPControllerAcceleratedSliderRepeatActive(*) => false
CPFocusedHwnd(*) => 0
CPHwndIsTab(*) => false
CPControllerColorDispatch(*) => false
CPControllerSendDialogKey(*) => 0
CPNavMove(*) => 0
CPNavSwitchTab(*) => 0
CPNavActivate(*) => 0
CPNavCancel(*) => 0
CPNavCancelCurrent(*) => 0
SavePanelBounds(*) => 0
DbgCP(*) => 0
CaptureControlPanelReturnWindow(*) => 0
RestoreControlPanelReturnWindow(*) {
    global TestExternalCalls
    TestExternalCalls += 1
}
CPShowControlPanelReady(*) {
    global TestExternalCalls
    TestExternalCalls += 1
}

; @BIGBOX_SOURCE@
