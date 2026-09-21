#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
; @UI_GLOBALS@
global TestAssertions := 0
global DesktopNavCancelCount := 0
global controlDarkMode := 1, iniPath := A_ScriptDir "\test-settings.ini", controlPanelOpacity := 100
global envPath := A_ScriptDir "\Settings\.env"
DirCreate(A_ScriptDir "\Settings")
if FileExist(envPath)
    FileDelete(envPath)
FileAppend("OPENAI_API_KEY=synthetic-openai`nGEMINI_API_KEY=synthetic-gemini`n", envPath, "UTF-8")
global APP_VERSION := "0.9.9.0", PROJECT_URL := "https://example.invalid/jrpg-translator"
global BUG_REPORT_URL := PROJECT_URL "/issues/new", WRITTEN_GUIDE_URL := PROJECT_URL "#quick-start"
global BEGINNER_VIDEO_URL := "https://example.invalid/beginner"
global gPidAudio := 0, gJustStoppedUntil := 0, gLastAction := ""
global gAudioProcess := 0
global gAudioSessionFile := A_ScriptDir "\audio-session.pid"
global CPToastGui := 0, CPToastText := 0, CPToastTimer := 0
global gAudioInputJob := Map("active", false), gAudioInputStatus := "Not tested.", speakerName := "[Windows Default]"
global __DBG_ENABLED_CP := false
global overlayTrans := 220, boxBgHex := "202020", txtHex := "FFFFFF", nameHex := "56C4F5"
global fontName := "Segoe UI", fontSize := 18, fontBold := 0
global overlayTrans_EW := 192, boxBgHex_EW := "101824", txtHex_EW := "E8E8E8"
global fontName_EW := "Consolas", fontSize_EW := 22, fontBold_EW := 1
global model_gemini_img := ["gemini-3.5-flash", "gemini-long-alternative"]
global model_openai_img := ["gpt-4o", "gpt-alternative"]
global imgProvider := "Gemini", geminiImgModel := "gemini-3.5-flash", imgModel := "gpt-4o"
global promptProfile := "default_with_kanji_reading_en", clearScreenshotsOnStartup := 0, capMaxKB := 1400
global model_gemini_audio := ["gemini-3.5-live-translate-preview", "gemini-alternative"]
global model_openai_audio := ["gpt-realtime", "gpt-alternative"]
global audioProvider := "Gemini", geminiAudioModel := "gemini-3.5-live-translate-preview", trModel := "gpt-realtime"
global audioTargetLangs := ["English (en)", "German (de)", "Japanese (ja)"], audioTargetLang := "English (en)"
global pythonExe := A_ScriptDir "\synthetic-python.exe"
global overlayAhk := A_ScriptDir "\synthetic-overlay.exe"
global imgScript := A_ScriptDir "\synthetic-image.py"
global audioScript := A_ScriptDir "\synthetic-audio.py"
global explainScript := A_ScriptDir "\synthetic-explainer.py"
global directModelOutput := 0, debugMode := 0, imgPostproc := "translation"
for syntheticPath in [pythonExe, overlayAhk, imgScript, audioScript, explainScript]
    if !FileExist(syntheticPath)
        FileAppend("synthetic", syntheticPath, "UTF-8")
global model_gemini_explain := ["gemini-explanation", "gemini-explanation-alternative"]
global model_openai_explain := ["gpt-explanation", "gpt-explanation-alternative"]
global explainProvider := "Gemini", explainGeminiModel := "gemini-explanation"
global explainOpenAIModel := "gpt-explanation", explainPromptProfile := "default_en"
global glossariesDir := A_ScriptDir "\glossaries"
global jp2enGlossaryProfile := "Japanese terms", en2enGlossaryProfile := "Local terms"
global useTerminologyOverrides := 1
for glossaryFixture in [["Japanese terms", "jp2en.txt"], ["Japanese alternate", "jp2en.txt"],
    ["Local terms", "en2en.txt"], ["Local alternate", "en2en.txt"]] {
    glossaryFixtureDir := glossariesDir "\" glossaryFixture[1]
    DirCreate(glossaryFixtureDir)
    FileAppend("Synthetic terminology", glossaryFixtureDir "\" glossaryFixture[2], "UTF-8")
}
global promptsDir := A_ScriptDir "\translation-prompts"
DirCreate(promptsDir)
for testPrompt in ["default_with_kanji_reading_en", "literal"] {
    testPromptPath := promptsDir "\" testPrompt ".txt"
    if !FileExist(testPromptPath)
        FileAppend("Synthetic prompt", testPromptPath, "UTF-8")
}
global explainPromptsDir := A_ScriptDir "\explanation-prompts"
DirCreate(explainPromptsDir)
for testPrompt in ["default_en", "detailed_grammar"] {
    testPromptPath := explainPromptsDir "\" testPrompt ".txt"
    if !FileExist(testPromptPath)
        FileAppend("Synthetic explanation prompt", testPromptPath, "UTF-8")
}
global gameProfilesDir := A_ScriptDir "\game-profiles"
DirCreate(gameProfilesDir)
for testProfile in ["Demo profile", "Alternate profile"] {
    testProfilePath := gameProfilesDir "\" testProfile ".ini"
    if !FileExist(testProfilePath)
        FileAppend("[profile]`nname=" testProfile "`n", testProfilePath, "UTF-8")
}
IniWrite("", iniPath, "game_profiles", "active")
global defGuiW := 1120, defGuiH := 760, pad := 12, gap := 8
global tabNames := ["Game Text Translation", "Audio Translation", "Translation Window", "Explanation",
    "Explanation Window", "Terminology Overrides", "Profiles", "Controls", "API Keys"]
global CPTabVisiblePages := [1, 2, 3, 4, 5, 6, 7, 8, 9]
; Begin with a real native caption, just like production. The desktop shell
; removes it; PrintWindow(PW_CLIENTONLY) excludes the retained resize frame.
global ui := Gui("+Resize +0x300000", "Synthetic desktop layout")
ui.SetFont("s10", "Segoe UI")
CPRegisterThemeMessages()
CPRefreshThemeBrushes()
global tab := ui.Add("Tab", "x12 y12 w1000 h560 Buttons -Wrap", tabNames)
CPRegisterCanvasFixedControl(tab, false, true)
DesktopBuildOrganizeFixtures()
global DesktopExplanationClicks := 0
global CPFooterFill := ui.AddText("x0 y570 w1120 h190"), sepAction := ui.AddText("x12 y570 w1000 h2 0x10")
global btnOv := ui.AddButton(, "Open Translator"), btnOvClose := ui.AddButton(, "Close Translator")
global btnAudio := ui.AddButton(, "Audio Translation Off"), btnExplainerLaunch := ui.AddButton(, "Open Explainer")
global btnExplainerClose := ui.AddButton(, "Close Explainer"), bClose := ui.AddButton(, "Close all")
global chkTop := ui.AddCheckbox(, "Always on top"), chkDarkMode := ui.AddCheckbox(, "Dark mode")
global txtControlOpacity := ui.AddText(, "Opacity:"), slControlOpacity := ui.AddSlider("Range70-100", 100)
global lblControlOpacityPct := ui.AddText(, "100%")
chkDarkMode.Value := 1
for testCtrl in CPDesktopFooterControls()
    CPRegisterCanvasFixedControl(testCtrl, false, true)
CPRegisterCanvasFixedControl(CPFooterFill, false, true)
CPRegisterCanvasFixedControl(sepAction, false, true)
CPCreateCustomTabBar()
IniWrite(0, iniPath, "cfg_control", "modernLayout") ; Exercise migration from the retired setting.
; Match production's cold-start order: startup synchronizers run before the
; modern Screenshot, Audio, Explanation, Overlay, Terminology, Profiles,
; Controls, and API Keys pages assign their shared aliases.
FixAllEditableCombos()
CPStartupOverlaysSync()
global TestColdStartProfilesGuard := !IsSet(ddlStartupOverlays)
CPDesktopCreate()
global TestApiKeysLoaded := cbApiInApp.Value = 1 && eGemini.Value = "synthetic-gemini"
    && eOpenAI.Value = "synthetic-openai" && eGemini.Enabled && eOpenAI.Enabled
cbApiInApp.Value := 0
ToggleApiKeyControls()
LoadFontsIntoCombo()
LoadFontsIntoCombo_EW()
; Source preflight validates production wiring. Replace the only action which
; would issue an AI request with a local callback for this synthetic harness.
btnExplainNow.OnEvent("Click", ExplainNow, 0)
btnExplainNow.OnEvent("Click", DesktopExplanationClicked)
FixAllEditableCombos()
CPStartupOverlaysSync()
ddlSpeaker.Add(["Game speakers", "Headphones"])
chkGuess.Value := 1, chkName.Value := 1
; Production wiring is validated by source preflight. Remove persistence handlers
; so synthetic native ComboBox messages cannot write incomplete fixture state.
for testCtrl in [ddlProv, ddlIMG, ddlIMG_GM, ddlPrompt]
    testCtrl.OnEvent("Change", CPScreenshotAISelectionChanged, 0)
for testCtrl in [ddlAProv, ddlA_GM, ddlTR, ddlAudioTarget]
    testCtrl.OnEvent("Change", CPAudioAISelectionChanged, 0)
CPRegisterThemeMessages()
CPRegisterCanvasMessages()
CPRefreshThemeBrushes()
CPCreateComboArrowOverlays()
OnError(DesktopTestUnhandledError)
try {
    ; @STUDY_ONLY@
    TestDesktopResizing()
    DesktopAssert(IniRead(iniPath, "cfg_control", "modernLayout", 0) = 1,
        "Legacy classic-layout preference migrates to modern")
    DesktopAssert(TestColdStartProfilesGuard && ddlStartupOverlays.Value = 1,
        "Cold-start synchronizers tolerate Profiles controls that are not constructed yet")
    DesktopAssert(CPDesktop["pages"].Count = 9, "Desktop registry contains every main page")
    modernCheckboxCount := 0, allModernCheckboxesPainted := true
    for desktopControlHwnd, desktopSurface in CPDesktop["surfaces"] {
        desktopControl := GuiCtrlFromHwnd(desktopControlHwnd)
        if desktopControl.Type != "CheckBox"
            continue
        modernCheckboxCount += 1
        if !CPDesktopNativePaintHwnds.Has(desktopControlHwnd)
            || CPDesktopNativePaintHwnds[desktopControlHwnd]["kind"] != "checkbox"
            allModernCheckboxesPainted := false
    }
    DesktopAssert(modernCheckboxCount >= 10 && allModernCheckboxesPainted,
        "Every modern desktop checkbox uses the shared subdued painter")
    expectedPages := Map(
        1, ["screenshot", "Game Text Translation", "screenshot"],
        2, ["audio", "Audio Translation", "audioPage"],
        3, ["translatorOverlay", "Overlay windows", "overlays"],
        4, ["explanation", "Explanation", "explanation"],
        5, ["explainerOverlay", "Overlay windows", "overlays"],
        6, ["terminology", "Settings", "settings"],
        7, ["profiles", "Profiles", "profiles"],
        8, ["controls", "Settings", "settings"],
        9, ["apiKeys", "Settings", "settings"])
    expectedControls := Map(1, CPDesktop["shot"], 2, CPDesktop["audioPage"],
        3, CPDesktop["overlayPages"][3], 4, CPDesktop["explanationPage"],
        5, CPDesktop["overlayPages"][5], 6, CPDesktop["organizePages"][6],
        7, CPDesktop["organizePages"][7], 8, CPDesktop["organizePages"][8],
        9, CPDesktop["organizePages"][9])
    ownedMinimums := Map(1, 29, 2, 24, 3, 28, 4, 28, 5, 26, 6, 22, 7, 18, 8, 96, 9, 18)
    registryPageKeys := Map()
    for registryPageId, registryExpected in expectedPages {
        registryDescriptor := CPDesktopPage(registryPageId)
        DesktopAssert(IsObject(registryDescriptor) && registryDescriptor["id"] = registryPageId
            && registryDescriptor["key"] = registryExpected[1] && registryDescriptor["title"] = registryExpected[2]
            && registryDescriptor["navKey"] = registryExpected[3] && registryDescriptor["subtitle"] != "",
            "Desktop page registry preserves identity and presentation metadata")
        DesktopAssert(!registryPageKeys.Has(registryDescriptor["key"]), "Desktop page registry keys are unique")
        registryPageKeys[registryDescriptor["key"]] := true
        pageControlsOwned := ownedMinimums.Has(registryPageId)
            ? registryDescriptor["adaptedControls"].Length = 0
                && registryDescriptor["controls"].Count >= ownedMinimums[registryPageId]
            : registryDescriptor["adaptedControls"].Length > 0
        DesktopAssert(ObjPtr(registryDescriptor["controls"]) = ObjPtr(expectedControls[registryPageId])
            && pageControlsOwned && IsObject(registryDescriptor["layout"]),
            "Desktop page registry owns control groups and layout dispatch")
        tab.Value := registryPageId
        CPDesktopLayout(ui, 0, 1120, 760)
        DesktopAssert(CPDesktop["chrome"]["title"].Text = registryDescriptor["title"]
            && CPDesktop["chrome"]["subtitle"].Text = registryDescriptor["subtitle"],
            "Registry metadata drives the visible desktop heading")
    }
    tab.Value := 1
    CPDesktopLayout(ui, 0, 1120, 760)
    for testSize in [[1120, 760], [900, 640], [1400, 820]] {
        testW := testSize[1], testH := testSize[2]
        ui.Show("NA x-9000 y-9000 w" testW " h" testH)
        CPDesktopLayout(ui, 0, testW, testH)
        CPApplyControlPanelTheme()
        CPDesktopRefreshStatus()
        DesktopAssert(CPDesktop["logo"].Get("image", 0), "Header logo loads from the bundled asset")
        DesktopAssert(CPDesktop["logo"]["bounds"][3] < 1254 && CPDesktop["logo"]["bounds"][4] < 1254, "Logo excludes empty margins without changing the asset")
        CPDesktop["chrome"]["brand"].GetPos(&desktopBrandX, &desktopBrandY, &desktopBrandW, &desktopBrandH)
        DesktopAssert(desktopBrandX = 92, "App name leaves a 20-pixel gap after the 52-pixel logo")
        DesktopAssert(DesktopBrandPointSize() = 14, "Header title uses the larger 14-point font")
        DesktopAssert(desktopBrandY + desktopBrandH / 2 = CPDesktop["headerH"] / 2, "Logo and app name share the header centerline")
        DesktopAssert(DesktopHitAt(40, 36) = 2 && DesktopHitAt(300, 64) = 2, "Logo and extended header remain draggable")
        desktopStyle := DllCall("user32\GetWindowLongPtr", "ptr", ui.Hwnd, "int", -16, "ptr")
        DesktopAssert(!(desktopStyle & 0xC00000), "Modern header removes the native caption")
        DesktopAssert((desktopStyle & 0xF0000) = 0xF0000, "Native resize, system menu, minimize and maximize remain enabled")
        DesktopAssert(DesktopHitAt(300, 28) = 2, "Blank header uses native caption dragging")
        DesktopAssert(DesktopHitAt(40, 28) = 2, "App-name region also drags the window")
        DesktopAssert(DesktopHitAt(testW - 240, 28) = 1, "Profile shortcut is not a drag target")
        desktopClientRectBuffer := Buffer(16, 0)
        DllCall("user32\GetClientRect", "ptr", ui.Hwnd, "ptr", desktopClientRectBuffer)
        desktopClosePointPx := Round((testW - 34) * GetWindowDPI(ui.Hwnd) / 96)
        DesktopAssert(desktopClosePointPx >= NumGet(desktopClientRectBuffer, 8, "int")
            || DesktopHitAt(testW - 34, 28) = 1,
            "Close button is not a drag target")
        DesktopAssert(DesktopHitAt(300, 100) = 1, "Page content is not a drag target")
        DesktopAssert(CPDesktopIsCombo(ddlProv.Hwnd) && CPDesktopIsCombo(ddlSpeaker.Hwnd) && CPDesktopIsCombo(ddlEPr.Hwnd) && CPDesktopIsCombo(ddlFont.Hwnd) && CPDesktopIsCombo(ddlFont_EW.Hwnd) && CPDesktopIsCombo(ddlGameProfile.Hwnd) && CPDesktopIsCombo(ddlJPG.Hwnd) && CPDesktopIsCombo(CPDesktop["chrome"]["profile"].Hwnd), "Rounded dropdown rendering covers the refreshed pages and header profile selector")
        DesktopAssert((DllCall("user32\GetWindowLongPtr", "ptr", ui.Hwnd, "int", -20, "ptr") & 0x02000000) != 0, "Modern desktop buffers parent and child painting together")
        DesktopAssert((DllCall("user32\GetClassLongPtr", "ptr", ui.Hwnd, "int", -26, "ptr") & 0xE0) = 0, "Desktop window class supports native composited painting")
        for desktopArrow in CPComboArrowOverlays {
            if CPDesktopIsCombo(desktopArrow["combo"].Hwnd)
                DesktopAssert(!DesktopShown(desktopArrow["arrow"]), "Modern dropdown has no separate native arrow overlay")
        }
        CPDesktop["chrome"]["profile"].GetPos(&desktopProfileX,, &desktopProfileW)
        CPDesktop["chrome"]["minimize"].GetPos(&desktopMinX)
        DesktopAssert(desktopProfileX + desktopProfileW + 12 <= desktopMinX, "Profile and caption controls have separate hit targets")
        DesktopAssert(!DesktopShown(CPTabButtons[1]), "Classic tab strip is hidden in the modern sidebar")
        DesktopAssert(!DesktopShown(CPDesktop["chrome"]["appearance"]) && !CPDesktop["chrome"].Has("classic"),
            "Window options has one footer entry and no classic-layout action")
        DesktopAssert(CPDesktop["paint"][CPDesktop["chrome"]["screenshot"].Hwnd]["selected"], "Screenshot sidebar item is selected")
        DesktopAssert(CPDesktop["contentW"] <= 1040, "Content width stays bounded on wide displays")
        ddlIMG.GetPos(, &modelRowY)
        ddlPrompt.GetPos(, &promptRowY)
        DesktopAssert(testW >= 1120 ? modelRowY = promptRowY : promptRowY > modelRowY, "AI fields reflow on narrow windows")
        profileSelector := CPDesktop["chrome"]["profile"]
        profileItemCount := SendMessage(0x146, 0, 0, profileSelector.Hwnd)
        profileManageLength := SendMessage(0x149, profileItemCount - 1, 0, profileSelector.Hwnd)
        profileManageText := Buffer((profileManageLength + 1) * 2, 0)
        SendMessage(0x148, profileItemCount - 1, profileManageText.Ptr, profileSelector.Hwnd)
        DesktopAssert(profileSelector.Text = "Current settings"
            && StrGet(profileManageText, "UTF-16") = "Manage profiles…"
            && CPDesktop["combos"][profileSelector.Hwnd]["separatorBefore"] = profileItemCount - 1,
            "Header profile selector shows current settings with a separated Manage profiles action")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlProv.Hwnd, "int", 0, "ptr") = ddlIMG_GM.Hwnd, "Keyboard Tab reaches the active model")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlIMG_GM.Hwnd, "int", 0, "ptr") = CPDesktop["shot"]["models"].Hwnd, "Keyboard Tab reaches model management next")
        DesktopAssert(DesktopShown(ddlIMG_GM) && !DesktopShown(ddlIMG), "Only active Gemini model is shown")
        DesktopAssert(!DesktopShown(chkDel) && !DesktopShown(eCapMax), "Advanced options start collapsed")
        DesktopAssert(ddlPrompt.Text = "default_with_kanji_reading_en", "Prompt choice preserved")
        DesktopAssert(!CPDesktop["shot"].Has("editPrompt") && CPDesktop["shot"]["prompts"].Text = "Manage…",
            "Game Text exposes prompt editing only through Manage")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlPrompt.Hwnd, "int", 0, "ptr") = CPDesktop["shot"]["prompts"].Hwnd,
            "Game Text prompt selector tabs directly to Manage")
        DesktopAssert(!DesktopShown(btnOv) && DesktopShown(CPDesktop["chrome"]["translator"]), "Compact footer replaces paired buttons")
        btnST.GetPos(&firstX, &firstY, &firstW)
        btnTS.GetPos(&secondX, &secondY)
        DesktopAssert(secondX >= firstX + firstW + 10 && secondY = firstY, "Capture actions do not overlap")
        before := DesktopRect(ddlPrompt)
        Loop 4
            CPDesktopLayout(ui, 0, testW, testH)
        DesktopAssert(before = DesktopRect(ddlPrompt), "Resizing does not accumulate coordinate drift")
        ddlProv.Choose(2)
        ToggleModelControls()
        DesktopAssert(DesktopShown(ddlIMG) && !DesktopShown(ddlIMG_GM), "Provider switch shows only OpenAI")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlProv.Hwnd, "int", 0, "ptr") = ddlIMG.Hwnd, "Provider switch keeps keyboard order correct")
        DesktopAssert(ddlIMG_GM.Text = "gemini-3.5-flash", "Inactive model choice is preserved")
        testScale := GetWindowDPI(ui.Hwnd) / 96
        ui.GetPos(,, &captureW, &captureH)
        TestCapture("desktop-" testW "-dark.png", Round(captureW * testScale), Round(captureH * testScale))
        CPDesktopToggleMulti()
        DesktopAssert(DesktopShown(CPDesktop["shot"]["multiHelp"]), "Multi-screenshot guidance expands")
        CPDesktopToggleMulti()
        DesktopAssert(!DesktopShown(CPDesktop["shot"]["multiHelp"]), "Multi-screenshot guidance collapses")
        CPDesktopToggleAdvanced()
        DesktopAssert(DesktopShown(chkDel) && DesktopShown(eCapMax), "Advanced options expand")
        DesktopAssert(CPCanvasScrollMaxY > 0, "Canvas exposes expanded content")
        DesktopAssert(CPDesktop["paint"][btnST.Hwnd]["kind"] = "primary", "Capture is the primary action")
        panelBefore := DesktopRect(CPDesktop["shot"]["formatPanel"])
        footerBefore := DesktopRect(CPDesktop["chrome"]["options"])
        CPCanvasScrollTo(0, CPCanvasScrollMaxY)
        DesktopAssert(DesktopRect(CPDesktop["shot"]["formatPanel"]) != panelBefore, "Panel geometry follows scrolling")
        DesktopAssert(DesktopRect(CPDesktop["chrome"]["options"]) = footerBefore, "Footer remains fixed while content scrolls")
        eCapMax.GetPos(, &fieldY,, &fieldH)
        DesktopAssert(fieldY >= CPDesktop["headerH"] + 88 && fieldY + fieldH <= testH - 58, "Advanced input is reachable above the footer")
        TestCapture("desktop-" testW "-expanded.png", Round(captureW * testScale), Round(captureH * testScale))
        CPCanvasScrollTo(0, 0)
        CPDesktopToggleAdvanced()
        tab.Value := 6
        CPDesktopLayout(ui, 0, testW, testH)
        DesktopAssert(CPDesktopPage(6)["adaptedControls"].Length = 0 && DesktopShown(ddlJPG) && !DesktopShown(btnST), "Settings owns the Terminology controls directly")
        DesktopAssert(CPDesktop["paint"][CPDesktop["chrome"]["settings"].Hwnd]["selected"], "Sidebar selection follows the settings page")
        DesktopAssert(!DesktopShown(CPDesktop["shot"]["models"]), "Screenshot tools do not leak to other pages")
        tab.Value := 1
        CPDesktopLayout(ui, 0, testW, testH)
        DesktopAssert(before = DesktopRect(ddlPrompt), "Round-trip page navigation restores geometry")
        ddlProv.Choose(1)
        ToggleModelControls()
    }
    DesktopTestAudioPage()
    DesktopTestExplanationPage()
    DesktopTestOwnerDrawFocusCues()
    DesktopTestOwnerDrawNavigationStability()
    DesktopTestSectionNavigationOrder()
    DesktopTestControllerPageScrollNavigation()
    DesktopTestActionMenus()
    DesktopTestOverlayPages()
    DesktopTestOrganizePages()
    DesktopTestNativeWheelAndFocus()
    DesktopTestScrollTransitionClipping()
    DesktopTestScrollPainting()
    DesktopTestSidebarRefresh()
    controlDarkMode := 0
    CPApplyControlPanelTheme()
    DesktopTestSidebarRefresh()
    controlDarkMode := 1
    CPApplyControlPanelTheme()
    for desktopLogoScale in [1, 1.5, 2, 2.5]
        DesktopCaptureLogoHeader(desktopLogoScale)
    CPDesktopLayout(ui, 0, 640, 560)
    CPDesktop["chrome"]["brand"].GetPos(&desktopBrandX,, &desktopBrandW)
    CPDesktop["chrome"]["profileLabel"].GetPos(&desktopProfileLabelX)
    DesktopAssert(desktopBrandX + desktopBrandW + 16 <= desktopProfileLabelX, "Brand and profile remain separated at minimum window width")
    DesktopAssert(DesktopBrandTextWidth() <= desktopBrandW, "Larger title fits without clipping at minimum window width")
    CPDesktopLayout(ui, 0, 1400, 820)
    ; Exercise the real ComboBox keyboard procedure while offscreen. No global
    ; Send/Click is used, and no application settings callback is attached.
    ddlPrompt.Choose(1)
    SendMessage(0x100, 0x28, 0, ddlPrompt.Hwnd) ; WM_KEYDOWN / Down
    SendMessage(0x101, 0x28, 0, ddlPrompt.Hwnd)
    DesktopAssert(ddlPrompt.Value = 2, "Rounded dropdown retains native arrow-key selection")
    CPShowCombo(ddlPrompt.Hwnd, true)
    DesktopAssert(CPComboDropped(ddlPrompt.Hwnd), "Rounded dropdown still opens its native list")
    SendMessage(0x100, 0x1B, 0, ddlPrompt.Hwnd)
    SendMessage(0x101, 0x1B, 0, ddlPrompt.Hwnd)
    DesktopAssert(!CPComboDropped(ddlPrompt.Hwnd), "Escape closes the dropdown list")
    ddlPrompt.Choose(1)
    ; Cloak this synthetic window during window-state tests so maximize never
    ; paints a test window over the user's desktop or takes foreground focus.
    ui.Opt("+E0x08000000")
    CPSetWindowCloaked(ui.Hwnd, true)
    ui.Show("Hide x20 y20 w651 h505")
    CPDesktopLayout(ui, 0, 651, 505)
    DesktopAssert(CPPrepareDesktopAfterBigBox(), "Fullscreen return prepares the desktop before reveal")
    ui.GetClientPos(,, &desktopRestoredW, &desktopRestoredH)
    DesktopAssert(desktopRestoredW > 651 && desktopRestoredH > 505,
        "Fullscreen return replaces a compact hidden measurement with normal desktop bounds")
    DesktopAssert(CPDesktop["width"] = desktopRestoredW && CPDesktop["height"] = desktopRestoredH,
        "Fullscreen return immediately relayouts the modern desktop at its restored size: layout "
            CPDesktop["width"] "x" CPDesktop["height"] ", client " desktopRestoredW "x" desktopRestoredH)
    CPDesktopWindowAction("maximize")
    DesktopAssert(DllCall("user32\IsZoomed", "ptr", ui.Hwnd), "Custom maximize invokes native maximization")
    DesktopAssert(CPDesktop["chrome"]["maximize"].Text = "Restore", "Maximized header exposes Restore")
    CPDesktopWindowAction("maximize")
    DesktopAssert(!DllCall("user32\IsZoomed", "ptr", ui.Hwnd), "Custom restore returns to windowed mode")
    CPDesktopWindowAction("minimize")
    DesktopAssert(DllCall("user32\IsIconic", "ptr", ui.Hwnd), "Custom minimize invokes native minimization")
    ui.Hide()
    CPSetWindowCloaked(ui.Hwnd, false)
    ui.Show("NA x-9000 y-9000 w1400 h820")
    CPDesktopLayout(ui, 0, 1400, 820)
    DesktopAssert(ObjPtr(CPDesktop["shot"]["captureTranslate"]) = ObjPtr(btnST)
        && ObjPtr(CPDesktop["shot"]["providerChoice"]) = ObjPtr(ddlProv),
        "Screenshot page registry owns the shared behavior aliases")
    for desktopTestPage in [1, 2, 3, 4, 5, 6, 7, 8, 9, 1] {
        CPDesktopNavigate(desktopTestPage)
        DesktopAssert(tab.Value = desktopTestPage, "Grouped navigation reaches existing page " desktopTestPage)
    }
    controlDarkMode := 0
    CPApplyControlPanelTheme()
    TestCapture("desktop-light.png", Round(captureW * testScale), Round(captureH * testScale))
    DesktopAssert(chkGuess.Value = 1 && chkName.Value = 1, "Formatting values survive layout and theme changes")
    ddlProv.Choose(2), ddlIMG.Choose(2), ddlPrompt.Choose(2), eCapMax.Value := "1800"
    ResizeUI(ui, 0, 1400, 820)
    DesktopAssert(CPDesktopActive() && ddlProv.Text = "OpenAI" && ddlIMG.Text = "gpt-alternative"
        && ddlPrompt.Text = "literal" && eCapMax.Value = "1800",
        "Modern-only relayout preserves live settings")
    DesktopAssert((DllCall("user32\GetWindowLongPtr", "ptr", ui.Hwnd, "int", -16, "ptr") & 0xC00000) = 0
        && CPDesktopIsCombo(ddlPrompt.Hwnd), "Modern-only relayout cannot return to classic chrome")
    global DesktopTestCloseCount := 0
    ui.OnEvent("Close", DesktopTestClosed)
    CPDesktopWindowAction("close")
    Sleep(30)
    DesktopAssert(DesktopTestCloseCount = 1, "Custom close routes through the existing GUI Close event")
    TestDesktopStudyWindows()
    DesktopTestWheelLifetime()
    SetTimer(CPControllerPoll, 0)
    ui.Destroy()
    DesktopAssert(CPCanvasWheelWindows.Count = 0, "Normal GUI destruction removes all wheel tracking")
    FileAppend("PASS: " TestAssertions " desktop layout assertions.`n", "*")
    ExitApp(0)
} catch as testError {
    FileAppend("FAIL: " testError.Message "`n" testError.Extra "`n" testError.File ":" testError.Line "`n" testError.Stack "`n", "*")
    ExitApp(1)
}

TestDesktopResizing() {
    global ui, tab, CPDesktop, CPPanelInteractiveResize, CPPanelLiveResizeTick
    global ddlProv, ddlPrompt, chkGuess, chkName, CPCanvasScrollX, CPCanvasScrollY
    global CPCanvasScrollMaxX, CPCanvasScrollMaxY
    global controlPanelOpacity, CPPanelResizeOpacitySuspended
    ; Deliberately omit the GUI Size callback: release must recover even when
    ; a Size event was dropped while another layout was still running.
    ui.Show("Restore NA x-9000 y-9000 w900 h640")
    CPDesktopRelayout()
    ui.Show("NA x-9000 y-9000 w1120 h900")
    CPFinalizeInteractiveResize()
    ui.GetClientPos(,, &w, &h)
    DesktopAssert(CPDesktop["width"] = w && CPDesktop["height"] = h,
        "Resize release lays out current dimensions even if the Size event was lost (layout "
            CPDesktop["width"] "x" CPDesktop["height"] ", client " w "x" h ")")
    selection := [ddlProv.Text, ddlPrompt.Text, chkGuess.Value, chkName.Value]
    ; Real native messages against the fixture only, without a GUI Size event
    ; handler or explicit layout call: check BEFORE release to prove live reflow.
    for page in [1, 2, 3, 4, 5, 6, 7, 8, 9] {
        CPDesktopNavigate(page)
        SendMessage(0x0231, 0, 0, ui.Hwnd)
        DesktopAssert(CPPanelInteractiveResize, "Native drag starts on page " page)
        for size in [[900, 640], [900, 900], [1400, 900], [820, 540], [1120, 640]] {
            Sleep(40)
            ui.Show("NA x-9000 y-9000 w" size[1] " h" size[2])
            DesktopAssertResizeBounds("Live page " page " at " size[1] "x" size[2])
            DesktopAssert(tab.Value = page, "Live resize preserves current page")
        }
        ; Force the final change to fall inside the live throttle window.
        CPPanelLiveResizeTick := A_TickCount + 1000
        ui.Show("NA x-9000 y-9000 w950 h700")
        SendMessage(0x0232, 0, 0, ui.Hwnd)
        Sleep(80)
        DesktopAssert(!CPPanelInteractiveResize, "Native release ends drag")
        DesktopAssertResizeBounds("Released page " page)
        DesktopAssert(CPCanvasScrollX >= 0 && CPCanvasScrollX <= CPCanvasScrollMaxX
            && CPCanvasScrollY >= 0 && CPCanvasScrollY <= CPCanvasScrollMaxY,
            "Released resize keeps scroll positions within the current page")
    }
    ; Normal programmatic sizes and stale queued GUI events use fresh bounds.
    ui.OnEvent("Size", ResizeUI)
    try {
        for size in [[1400, 900], [900, 540], [1120, 760]] {
            ui.Show("NA x-9000 y-9000 w" size[1] " h" size[2])
            Sleep(80)
            DesktopAssertResizeBounds("GUI Size event")
            ResizeUI(ui, 0, 820, 380)
            DesktopAssertResizeBounds("Delayed stale GUI Size event")
        }
        CPOnWindowEnterSizeMove(0, 0, 0x0231, ui.Hwnd)
        CPOnWindowExitSizeMove(0, 0, 0x0232, ui.Hwnd)
        CPOnWindowEnterSizeMove(0, 0, 0x0231, ui.Hwnd)
        CPFinalizeInteractiveResize()
        DesktopAssert(CPPanelInteractiveResize, "Previous finalizer cannot finish a newer drag")
        CPOnWindowExitSizeMove(0, 0, 0x0232, ui.Hwnd)
        Sleep(80)
        DesktopAssertResizeBounds("Rapid consecutive drags")
    } finally {
        ui.OnEvent("Size", ResizeUI, 0)
    }
    savedOpacity := controlPanelOpacity
    try {
        controlPanelOpacity := 85
        CPApplyControlPanelOpacity(false)
        SendMessage(0x0231, 0, 0, ui.Hwnd)
        DesktopAssert(CPPanelResizeOpacitySuspended, "Transparent window suspends alpha during drag")
        ui.Show("NA x-9000 y-9000 w1120 h760")
        CPRestoreOpacityAfterResize()
        DesktopAssert(CPPanelResizeOpacitySuspended, "Opacity timer cannot interrupt an active drag")
        SendMessage(0x0232, 0, 0, ui.Hwnd)
        Sleep(120)
        DesktopAssert(!CPPanelResizeOpacitySuspended
            && WinGetTransparent("ahk_id " ui.Hwnd) = Round(85 * 255 / 100),
            "Release restores exactly the configured opacity")
        DesktopAssertResizeBounds("Transparent window release")
    } finally {
        controlPanelOpacity := savedOpacity
        CPApplyControlPanelOpacity(false)
    }
    DesktopAssert(ddlProv.Text = selection[1] && ddlPrompt.Text = selection[2]
        && chkGuess.Value = selection[3] && chkName.Value = selection[4],
        "Resizing preserves provider, prompt and checkbox values")
    CPDesktopNavigate(1)
    ui.GetClientPos(,, &w, &h)
    scale := DllCall("user32\GetDpiForWindow", "ptr", ui.Hwnd, "uint") / 96
    TestCapture("desktop-resize-final.png", Round(w * scale), Round(h * scale))
}

DesktopAssertResizeBounds(label) {
    global ui, CPDesktop
    ui.GetClientPos(,, &w, &h)
    DesktopAssert(CPDesktop["width"] = w && CPDesktop["height"] = h,
        label " uses actual client dimensions (layout " CPDesktop["width"] "x" CPDesktop["height"]
            ", client " w "x" h ")")
    CPDesktop["chrome"]["footerPanel"].GetPos(&fx, &fy, &fw, &fh)
    DesktopAssert(fx = 0 && Abs(fy + fh - h) <= 1 && Abs(fw - w) <= 1,
        label " anchors the footer to the bottom and both sides")
    CPDesktop["chrome"]["headerPanel"].GetPos(&hx, &hy, &hw)
    DesktopAssert(hx = 0 && hy = 0 && Abs(hw - w) <= 1,
        label " fills the header width")
}

DesktopAssert(condition, message) {
    global TestAssertions
    TestAssertions += 1
    if !condition
        throw Error(message)
}

DesktopCountColor(dc, color, left, top, right, bottom) {
    count := 0
    Loop Max(0, bottom - top) {
        y := top + A_Index - 1
        Loop Max(0, right - left) {
            x := left + A_Index - 1
            if DllCall("gdi32\GetPixel", "ptr", dc, "int", x, "int", y, "uint") = color
                count += 1
        }
    }
    return count
}

DesktopTestOwnerDrawFocusCues() {
    global ui, CPDesktop, controlDarkMode
    originalDark := controlDarkMode
    controlDarkMode := 1
    CPDesktopNavigate(1)
    CPDesktopLayout(ui, 0, 1120, 760)
    CPApplyControlPanelTheme()
    ctrl := CPDesktop["shot"]["models"]
    client := Buffer(16, 0)
    DllCall("user32\GetClientRect", "ptr", ctrl.Hwnd, "ptr", client.Ptr)
    width := NumGet(client, 8, "int"), height := NumGet(client, 12, "int")
    windowDc := DllCall("user32\GetDC", "ptr", ctrl.Hwnd, "ptr")
    dc := DllCall("gdi32\CreateCompatibleDC", "ptr", windowDc, "ptr")
    bitmap := DllCall("gdi32\CreateCompatibleBitmap", "ptr", windowDc,
        "int", width, "int", height, "ptr")
    previous := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", bitmap, "ptr")
    data := Buffer(A_PtrSize = 8 ? 64 : 48, 0)
    NumPut("uint", 4, data, 0) ; ODT_BUTTON
    NumPut("ptr", ctrl.Hwnd, data, A_PtrSize = 8 ? 24 : 20)
    NumPut("ptr", dc, data, A_PtrSize = 8 ? 32 : 24)
    rectOffset := A_PtrSize = 8 ? 40 : 28
    NumPut("int", width, data, rectOffset + 8)
    NumPut("int", height, data, rectOffset + 12)
    try {
        linkColor := CPColorRef(CPDesktopPalette()["link"])
        strip := Max(4, Round(4 * GetWindowDPI(ctrl.Hwnd) / 96))
        midLeft := Round(width * 0.25), midRight := Round(width * 0.75)
        midTop := Round(height * 0.25), midBottom := Round(height * 0.75)

        NumPut("uint", 0x10, data, 16) ; ODS_FOCUS: keyboard/controller cue visible
        DesktopAssert(CPDesktopDrawItem(data.Ptr), "Focused link renders through the desktop owner-draw path")
        DesktopAssert(DesktopCountColor(dc, linkColor, midLeft, 0, midRight, strip) > 0,
            "Keyboard focus outline has a visible top edge")
        DesktopAssert(DesktopCountColor(dc, linkColor, midLeft, height - strip, midRight, height) > 0,
            "Keyboard focus outline has a visible bottom edge")
        DesktopAssert(DesktopCountColor(dc, linkColor, 0, midTop, strip, midBottom) > 0,
            "Keyboard focus outline has a visible left edge")
        DesktopAssert(DesktopCountColor(dc, linkColor, width - strip, midTop, width, midBottom) > 0,
            "Keyboard focus outline has a visible right edge")

        NumPut("uint", 0x210, data, 16) ; ODS_FOCUS | ODS_NOFOCUSRECT: pointer activation
        DesktopAssert(CPDesktopDrawItem(data.Ptr), "Pointer-focused link renders through the desktop owner-draw path")
        DesktopAssert(DesktopCountColor(dc, linkColor, midLeft, 0, midRight, strip) = 0
            && DesktopCountColor(dc, linkColor, midLeft, height - strip, midRight, height) = 0,
            "Pointer activation does not leave a sticky blue focus outline")
    } finally {
        DllCall("gdi32\SelectObject", "ptr", dc, "ptr", previous)
        DllCall("gdi32\DeleteObject", "ptr", bitmap)
        DllCall("gdi32\DeleteDC", "ptr", dc)
        DllCall("user32\ReleaseDC", "ptr", ctrl.Hwnd, "ptr", windowDc)
        controlDarkMode := originalDark
        CPApplyControlPanelTheme()
    }
}

DesktopTestOwnerDrawNavigationStability() {
    global ui, tab, CPDesktop, btnSpRef, btnAudioTest, CPFocusVisualNavHwnd
    global DesktopNavCancelCount
    CPDesktopNavigate(2)
    ui.Show("NA x-9000 y-9000 w1120 h760")
    CPDesktopLayout(ui, 0, 1120, 760)
    CPApplyControlPanelTheme()

    paintedButtons := []
    originalFonts := Map()
    for hwnd, paint in CPDesktop["paint"] {
        if paint["ctrl"].Type != "Button"
            continue
        paintedButtons.Push(hwnd)
        originalFonts[hwnd] := SendMessage(0x31, 0, 0, hwnd) ; WM_GETFONT
        DesktopAssert((DllCall("user32\GetWindowLongPtr", "ptr", hwnd, "int", -16, "ptr") & 0xF) = 0xB,
            "Desktop button begins keyboard navigation as owner-drawn")
    }

    focusSequence := [
        CPDesktop["chrome"]["screenshot"],
        CPDesktop["chrome"]["audioPage"],
        CPDesktop["chrome"]["explanation"],
        btnSpRef,
        btnAudioTest,
        CPDesktop["audioPage"]["power"],
        CPDesktop["chrome"]["translator"],
        CPDesktop["chrome"]["audio"]
    ]
    for ctrl in focusSequence {
        CPSetFocusHwnd(ctrl.Hwnd)
        UpdateCPFocusRing()
        DesktopAssert(DllCall("user32\GetFocus", "ptr") = ctrl.Hwnd,
            "Keyboard/controller focus reaches the requested desktop button")
        for hwnd in paintedButtons {
            DesktopAssert((DllCall("user32\GetWindowLongPtr", "ptr", hwnd, "int", -16, "ptr") & 0xF) = 0xB,
                "Focus traversal preserves owner-draw rendering for every desktop button")
            DesktopAssert(SendMessage(0x31, 0, 0, hwnd) = originalFonts[hwnd],
                "Focus traversal does not mutate desktop button fonts")
        }
    }

    CPDesktopNavigate(1)
    CPSetFocusHwnd(CPDesktop["chrome"]["explanation"].Hwnd)
    UpdateCPFocusRing()
    CPNavActivate("Enter")
    Sleep(20)
    DesktopAssert(tab.Value = 4,
        "Keyboard Enter invokes a focused owner-drawn desktop navigation button")

    CPDesktopNavigate(1)
    CPSetFocusHwnd(CPDesktop["chrome"]["audioPage"].Hwnd)
    UpdateCPFocusRing()
    CPControllerDispatchNavigation("Activate", ui.Hwnd)
    Sleep(20)
    DesktopAssert(tab.Value = 2,
        "Controller A invokes a focused owner-drawn desktop navigation button")

    DesktopNavCancelCount := 0
    CPControllerDispatchNavigation("Cancel", ui.Hwnd, DesktopRecordRootCancel)
    DesktopAssert(DesktopNavCancelCount = 1,
        "Controller B reaches the desktop root cancel action when no editor consumes it")

    ui.GetPos(,, &navCaptureW, &navCaptureH)
    navDpi := GetWindowDPI(ui.Hwnd) / 96
    TestCapture("desktop-navigation-focus-stable.png", Round(navCaptureW * navDpi), Round(navCaptureH * navDpi))
    CPRestoreFocusVisual()
    CPFocusVisualNavHwnd := 0
}

DesktopRecordRootCancel(*) {
    global DesktopNavCancelCount
    DesktopNavCancelCount += 1
}

DesktopTestSectionNavigationOrder() {
    global ui, tab, CPTabVisiblePages, CPDesktop, CPFocusVisualHwnd

    expected := [1, 2, 4, 3, 5, "study", 7, 8, 6, 9]
    CPDesktopNavigate(expected[1])
    displayedPage := expected[1]
    Loop expected.Length - 1 {
        expectedStop := expected[A_Index + 1]
        CPNavSwitchTab(1)
        if expectedStop = "study" {
            UpdateCPFocusRing()
            DesktopAssert(tab.Value = displayedPage
                && DllCall("user32\GetFocus", "ptr") = CPDesktop["chrome"]["study"].Hwnd,
                "Forward section navigation focuses Study Library without opening or changing the page")
            DesktopAssert(CPFocusVisualHwnd = CPDesktop["chrome"]["study"].Hwnd
                && !CPDesktop["paint"][CPDesktop["chrome"]["study"].Hwnd]["selected"],
                "Study Library receives only the blue navigation border")
        } else {
            displayedPage := expectedStop
            DesktopAssert(tab.Value = expectedStop,
                "Forward section navigation follows the modern sidebar order")
        }
    }
    CPNavSwitchTab(1)
    DesktopAssert(tab.Value = expected[1],
        "Forward section navigation wraps from API Keys to Screenshot")

    CPNavSwitchTab(-1)
    DesktopAssert(tab.Value = expected[-1],
        "Backward section navigation wraps from Screenshot to API Keys")
    displayedPage := expected[-1]
    Loop expected.Length - 1 {
        expectedStop := expected[expected.Length - A_Index]
        CPNavSwitchTab(-1)
        if expectedStop = "study" {
            UpdateCPFocusRing()
            DesktopAssert(tab.Value = displayedPage
                && DllCall("user32\GetFocus", "ptr") = CPDesktop["chrome"]["study"].Hwnd,
                "Backward section navigation focuses Study Library without opening or changing the page")
        } else {
            displayedPage := expectedStop
            DesktopAssert(tab.Value = expectedStop,
                "Backward section navigation follows the modern sidebar order")
        }
    }

    CPDesktopNavigate(2)
    CPControllerDispatchNavigation("NextTab", ui.Hwnd)
    DesktopAssert(tab.Value = 4,
        "Controller next-section navigation reaches Explanation after Audio")
    CPControllerDispatchNavigation("PreviousTab", ui.Hwnd)
    DesktopAssert(tab.Value = 2,
        "Controller previous-section navigation returns from Explanation to Audio")

    CPDesktopNavigate(7)
    for spec in [[8, "controlsTab"], [6, "termsTab"], [9, "apiTab"]] {
        CPControllerDispatchNavigation("NextTab", ui.Hwnd)
        DesktopAssert(tab.Value = spec[1],
            "Controller R follows the visible Settings order to page " spec[1])
        DesktopAssert(DllCall("user32\GetFocus", "ptr")
            = CPDesktop["organizePages"][spec[1]][spec[2]].Hwnd,
            "Controller R focuses the selected Settings button: " spec[2])
    }
    for spec in [[6, "termsTab"], [8, "controlsTab"]] {
        CPControllerDispatchNavigation("PreviousTab", ui.Hwnd)
        DesktopAssert(tab.Value = spec[1],
            "Controller L reverses the visible Settings order to page " spec[1])
        DesktopAssert(DllCall("user32\GetFocus", "ptr")
            = CPDesktop["organizePages"][spec[1]][spec[2]].Hwnd,
            "Controller L focuses the selected Settings button: " spec[2])
    }
    CPControllerDispatchNavigation("PreviousTab", ui.Hwnd)
    DesktopAssert(tab.Value = 7
        && DllCall("user32\GetFocus", "ptr") = CPDesktop["chrome"]["profiles"].Hwnd,
        "Controller L returns from Controls to Profiles in visible sidebar order")

    CPDesktopNavigate(5)
    CPControllerDispatchNavigation("NextTab", ui.Hwnd)
    DesktopAssert(tab.Value = 5
        && DllCall("user32\GetFocus", "ptr") = CPDesktop["chrome"]["study"].Hwnd,
        "Controller R pauses on Study Library after Explanation Window")
    CPControllerDispatchNavigation("NextTab", ui.Hwnd)
    DesktopAssert(tab.Value = 7
        && DllCall("user32\GetFocus", "ptr") = CPDesktop["chrome"]["profiles"].Hwnd,
        "Controller R continues from Study Library to Profiles")
    CPControllerDispatchNavigation("PreviousTab", ui.Hwnd)
    DesktopAssert(tab.Value = 7
        && DllCall("user32\GetFocus", "ptr") = CPDesktop["chrome"]["study"].Hwnd,
        "Controller L pauses on Study Library before Explanation Window")
    CPControllerDispatchNavigation("PreviousTab", ui.Hwnd)
    DesktopAssert(tab.Value = 5
        && DllCall("user32\GetFocus", "ptr") = CPDesktop["overlayPages"][5]["explainer"].Hwnd,
        "Controller L returns from Study Library to the already displayed Explanation Window page")

    CPDesktopNavigate(1)
}

DesktopTestControllerPageScrollNavigation() {
    global ui, tab, CPDesktop, CPCanvasScrollY, CPCanvasScrollMaxY
    global chkGuess, chkName

    ui.Show("NA x-9000 y-9000 w1120 h720")
    CPDesktop["advanced"] := false
    CPDesktop["audioHelpOpen"] := false
    CPDesktop["explanationStartupOpen"] := false

    ; These are the two reported short-window transitions. In both cases the
    ; fixed footer is physically nearer than the next page row at scroll zero.
    for testCase in [[1, chkGuess, chkName, "Game Text"],
        [2, CPDesktop["audioPage"]["troubleshoot"], CPDesktop["audioPage"]["power"], "Audio"]] {
        CPDesktopNavigate(testCase[1])
        CPDesktopLayout(ui, 0, 1120, 720)
        CPCanvasScrollTo(0, 0)
        CPSetFocusHwnd(testCase[2].Hwnd)
        CPNavMove("Down")
        Sleep(30)
        DesktopAssert(DllCall("user32\GetFocus", "ptr") = testCase[3].Hwnd,
            testCase[4] " D-pad Down reaches the next offscreen page control before the footer")
        DesktopAssert(CPCanvasScrollY > 0,
            testCase[4] " D-pad focus automatically scrolls the lower control into view")
    }

    ; Exercise the same rule on every scrollable desktop page. Pick the lowest
    ; currently visible page control that still has a lower logical row; Down
    ; must stay in page content, and Up from the footer must return to the final
    ; page row instead of whichever control happens to be nearest onscreen.
    for page in [1, 2, 3, 4, 5, 6, 7, 8, 9] {
        CPDesktopNavigate(page)
        CPDesktopLayout(ui, 0, 1120, 720)
        CPCanvasScrollTo(0, 0)
        if CPCanvasScrollMaxY <= 0
            continue
        pageHwnds := CPDesktopPageNavigationHwnds()
        footerHwnds := CPDesktopFooterNavigationHwnds()
        if !pageHwnds.Length || !footerHwnds.Length
            continue
        footerTop := CPGetHwndRect(footerHwnds[1])["t"]
        source := 0, sourceBottom := 0
        for hwnd in pageHwnds {
            rect := CPGetHwndRect(hwnd)
            if rect["b"] > footerTop || rect["b"] <= sourceBottom
                continue
            hasLower := false
            for candidate in pageHwnds {
                if candidate != hwnd && CPGetHwndRect(candidate)["t"] >= rect["b"] - 4 {
                    hasLower := true
                    break
                }
            }
            if hasLower
                source := hwnd, sourceBottom := rect["b"]
        }
        if source {
            CPSetFocusHwnd(source)
            CPNavMove("Down")
            Sleep(30)
            focused := DllCall("user32\GetFocus", "ptr")
            DesktopAssert(CPHwndArrayIndex(focused, CPDesktopPageNavigationHwnds()) > 0,
                "Scrollable page " page " keeps D-pad Down in page content before its footer")
        }

        CPCanvasScrollTo(0, 0)
        bottom := -2147483648
        for hwnd in CPDesktopPageNavigationHwnds()
            bottom := Max(bottom, CPGetHwndRect(hwnd)["b"])
        finalRow := Map()
        for hwnd in CPDesktopPageNavigationHwnds() {
            if CPGetHwndRect(hwnd)["b"] >= bottom - 8
                finalRow[hwnd] := true
        }
        CPSetFocusHwnd(footerHwnds[1])
        CPNavMove("Up")
        Sleep(30)
        focused := DllCall("user32\GetFocus", "ptr")
        DesktopAssert(finalRow.Has(focused),
            "Scrollable page " page " returns from the footer to its final content row")
    }

    CPDesktopNavigate(1)
    ui.Show("NA x-9000 y-9000 w1400 h820")
    CPDesktopLayout(ui, 0, 1400, 820)
}

DesktopActionMenuRecord(action, *) {
    global DesktopActionMenuSelections
    DesktopActionMenuSelections.Push(action)
}

DesktopActionMenuFinish(selection, captureName, *) {
    global DesktopActionMenuTimer
    for hwnd, s in CPThemedChoicePopupRegistry() {
        if s["rows"].Length != 2
            continue
        if Trim(s["rows"][1].Text) != "Add model…"
            || Trim(s["rows"][2].Text) != "Remove selected model…"
            continue
        SetTimer(DesktopActionMenuTimer, 0)
        if captureName != "" {
            s["gui"].GetClientPos(,, &width, &height)
            scale := GetWindowDPI(s["gui"].Hwnd) / 96
            TestDesktopStudyCapture(s["gui"], captureName,
                Round(width * scale), Round(height * scale))
        }
        CPThemedChoicePopupFinish(s, s["gui"], selection)
        return
    }
}

DesktopTestActionMenus() {
    global ui, CPDesktop, controlDarkMode
    global DesktopActionMenuTimer := 0, DesktopActionMenuSelections := []
    originalDark := controlDarkMode
    choices := ["Add model…", "Remove selected model…"]
    actions := [DesktopActionMenuRecord.Bind("add"), DesktopActionMenuRecord.Bind("remove")]
    try {
        ui.Show("NA x-9000 y-9000 w1120 h760")
        CPDesktopNavigate(1)
        CPDesktopLayout(ui, 0, 1120, 760)
        anchor := CPDesktop["shot"]["models"]
        for spec in [[1, 1, "dark"], [0, 2, "light"]] {
            controlDarkMode := spec[1]
            CPApplyControlPanelTheme()
            WinActivate("ahk_id " ui.Hwnd)
            captureName := "desktop-action-menu-" spec[3] ".png"
            DesktopActionMenuTimer := DesktopActionMenuFinish.Bind(spec[2], captureName)
            SetTimer(DesktopActionMenuTimer, 20)
            result := CPDesktopActionMenu(anchor, choices, actions)
            SetTimer(DesktopActionMenuTimer, 0)
            DesktopAssert(result = spec[2], "Modern desktop action menu returns its selected row")
        }
        DesktopAssert(DesktopActionMenuSelections.Length = 2
            && DesktopActionMenuSelections[1] = "add"
            && DesktopActionMenuSelections[2] = "remove",
            "Modern desktop action menu dispatches only the chosen callback")

        DesktopActionMenuTimer := DesktopActionMenuFinish.Bind(0, "")
        SetTimer(DesktopActionMenuTimer, 20)
        result := CPDesktopActionMenu(anchor, choices, actions)
        SetTimer(DesktopActionMenuTimer, 0)
        DesktopAssert(result = 0 && DesktopActionMenuSelections.Length = 2,
            "Dismissing a modern desktop action menu is inert")
    } finally {
        if IsObject(DesktopActionMenuTimer)
            SetTimer(DesktopActionMenuTimer, 0)
        controlDarkMode := originalDark
        CPApplyControlPanelTheme()
    }
}

DesktopTestExitWithLiveWindow() {
    global ui, CPCanvasWheelWindows, DesktopExitWheelHandles
    global DesktopExitStudyWindows, DesktopExitHeaderProcs, CPStudyThemedHeaderHwnds
    DesktopTestWheelLifetime()
    ui.Show("NA x-9000 y-9000 w900 h640")
    CPDesktopNavigate(8)
    CPDesktopLayout(ui, 0, 900, 640)
    DesktopExitWheelHandles := []
    for hwnd in CPCanvasWheelWindows
        DesktopExitWheelHandles.Push(hwnd)
    library := TestDesktopStudyLibrary(), reader := TestDesktopStudyReader()
    library["suspend"] := true
    DesktopExitStudyWindows := [library["desktop"], reader["desktop"]]
    for state in [library, reader] {
        state["gui"].Show("NA x-12000 y-12000 w1200 h800")
        StudyDesktopResize(state, 1200, 800)
    }
    dialogs := [TestDesktopAnkiControls(reader)]
    dialogs.Push(TestDesktopCandidateControls(library))
    dialogs.Push(StudyDesktopMessageCreate(reader["gui"].Hwnd, "Synthetic exit check", "Add to Anki", "yesno", "Confirmation", "Add to Anki", "Back to review"))
    dialogs.Push(TestDesktopChapterControls(library), TestDesktopColumnsControls(library),
        TestDesktopFilterControls(library), TestDesktopDetailsControls(library))
    filterDialog := dialogs[6], filterControls := filterDialog["desktopControls"]
    dialogs.Push(StudyDesktopDatePickerCreate(filterDialog, filterControls["date"],
        filterDialog["dateFrom"], filterControls["from"], "From"))
    manager := TestDesktopManagerControls(library)
    dialogs.Push(manager, TestDesktopArchivesControls(manager),
        StudyDesktopLibraryNameCreate(manager, "new"), TestDesktopStorageControls(library))
    recommendation := TestDesktopRecommendationConfirmControls(library)
    preferences := TestDesktopRecommendationPreferencesControls(recommendation)
    dialogs.Push(recommendation, preferences, TestDesktopRecommendationPromptControls(preferences))
    newVersion := TestDesktopNewVersionControls(reader)
    dialogs.Push(TestDesktopConnectionControls(library), TestDesktopBulkControls(library), newVersion,
        StudyReaderPromptDialogCreate(newVersion, "edit", "Synthetic prompt", "Fixture"),
        StudyReaderPromptDialogCreate(newVersion, "name"))
    dialogs.Push(CPInputDialogCreate(reader["gui"].Hwnd, "Profile name", "Create profile", "Synthetic shutdown fixture"))
    dialogs.Push(CPModelDialogCreate(reader["gui"].Hwnd, "gemini", "screenshot", "source"),
        CPModelDialogCreate(reader["gui"].Hwnd, "openai", "audio", "browser", TestModelCatalogFixture(), []))
    dialogs.Push(CPTextEditorDialogCreate(reader["gui"].Hwnd, "Prompt editor",
        "Synthetic shutdown fixture", "Prompt instructions", "日本語 {jp}",
        "Unsaved text remains a local draft."))
    dialogs.Push(CPHotkeyDialogCreate(reader["gui"].Hwnd, "^!t", "take_screenshot"))
    dialogs.Push(CPControllerCaptureDialogCreate(reader["gui"].Hwnd, "take_screenshot"))
    dialogs.Push(CPDesktopAppearanceCreate(reader["gui"].Hwnd))
    dialogs.Push(CPControllerColorDialogCreate(reader["gui"].Hwnd, "2563EB",
        "Adjust Translator window color"))
    dialogs.Push(CPDesktopWelcomeDialogCreate(reader["gui"].Hwnd, false))
    dialogs.Push(CPDesktopAboutDialogCreate(reader["gui"].Hwnd))
    dialogs.Push(StudyDesktopMessageCreate(reader["gui"].Hwnd,
        "The current settings contain unsaved edits.", "Unsaved settings", "yesnocancel",
        "Choose what to do with these changes.",
        "Save changes", "Discard changes", "Cancel"))
    for state in dialogs {
        DesktopExitStudyWindows.Push(state["desktop"])
        DesktopDialogFixtureShow(state, 1040, 800)
    }
    livePopup := TestSharedChoicePopup(reader["gui"].Hwnd, ["Copy current section", "Copy full explanation"])
    livePopup["gui"].Show("NA x-12000 y-12000 AutoSize")
    CPThemedChoicePopupRegister(livePopup)
    liveMenu := CPFullscreenMenuCreate(reader["gui"].Hwnd, [Map("label", "Open in Reader")])
    liveMenu["gui"].Show("NA x-12000 y-12000 w1280 h720")
    StudyCandidatesRecommendationBigBoxResize(liveMenu, liveMenu["gui"], 0, 1280, 720)
    StudyBigBoxFocusFrameStart(liveMenu)
    DesktopExitHeaderProcs := CPStudyThemedHeaderHwnds.Clone()
    OnExit(DesktopTestAssertShutdown)
    Critical "On"
    CPCanvasQueueWheel(-120)
    FileAppend("Exiting with a live desktop window and a queued scroll frame.`n", "*")
    ; Deliberately do not call ui.Destroy(): automatic GUI destruction runs
    ; after OnExit and global-variable release, just as in the real app.
    ExitApp(0)
}

DesktopTestUnhandledError(err, *) {
    FileAppend("FAIL: Unhandled test callback: " err.Message "`n" err.Extra "`n" err.File ":" err.Line "`n" err.Stack "`n", "*")
    ExitApp(1)
}

DesktopTestAssertShutdown(*) {
    global CPCanvasWheelWindows, CPCanvasWheelSubclass, DesktopExitWheelHandles
    global CPCanvasPendingScrollValid, CPCanvasScrollFlushScheduled, TestAssertions
    global DesktopExitStudyWindows, DesktopExitHeaderProcs
    try {
        DesktopAssert(StudyDesktopRegistry().Count = 0, "Exit unregisters live Study paint handlers")
        DesktopAssert(StudyLibraryDatePickerRegistry().Count = 0, "Exit unregisters the live date picker before static state is released")
        for d in DesktopExitStudyWindows
            DesktopAssert(d["iconFonts"].Count = 0, "Exit releases Study caption fonts")
        DesktopAssert(DesktopExitHeaderProcs.Count >= 2, "Shutdown fixture includes both candidate table headers")
        for hwnd, nativeProc in DesktopExitHeaderProcs
            DesktopAssert(DllCall("user32\GetWindowLongPtr", "ptr", hwnd, "int", -4, "ptr") = nativeProc,
                "Exit restores every native header procedure before releasing global state")
        DesktopAssert(DesktopExitWheelHandles.Length > 1, "Exit fixture has parent and child wheel handlers")
        DesktopAssert(!CPCanvasPendingScrollValid && !CPCanvasScrollFlushScheduled, "Exit cancels queued scrolling before globals are released")
        DesktopAssert(CPCanvasWheelWindows.Count = 0, "Exit clears native wheel tracking before globals are released")
        for hwnd in DesktopExitWheelHandles {
            refData := 0
            DesktopAssert(!DllCall("comctl32\GetWindowSubclass", "ptr", hwnd, "ptr", CPCanvasWheelSubclass,
                "uptr", 12, "uptr*", &refData), "Exit detaches each native wheel handler")
        }
        CPShutdownCanvasMessages()
        DesktopAssert(CPCanvasWheelWindows.Count = 0 && !CPCanvasScrollFlushScheduled, "Repeated shutdown cleanup is harmless")
        FileAppend("PASS: " TestAssertions " live-window shutdown assertions.`n", "*")
    } catch as err {
        FileAppend("FAIL: " err.Message "`n" err.Stack "`n", "*")
    }
    return 0 ; A failing assertion must not cancel this isolated process's exit.
}

DesktopTestWheelLifetime() {
    global CPCanvasWheelWindows, CPCanvasWheelSubclass, CPCanvasShuttingDown
    for releaseGlobals in [false, true] {
        probe := Gui("+ToolWindow -Caption", "Isolated wheel lifetime probe")
        child := probe.AddText(, "Synthetic child")
        handles := [probe.Hwnd, child.Hwnd]
        savedWindows := CPCanvasWheelWindows, savedCallback := CPCanvasWheelSubclass
        savedShuttingDown := CPCanvasShuttingDown
        for hwnd in handles {
            DesktopAssert(DllCall("comctl32\SetWindowSubclass", "ptr", hwnd, "ptr", savedCallback,
                "uptr", 12, "uptr", savedCallback), "Lifetime probe installs the production wheel callback")
            CPCanvasWheelWindows[hwnd] := true
        }
        try {
            if releaseGlobals {
                ; Reproduce the reported teardown order deterministically.
                CPCanvasWheelWindows := unset
                CPCanvasWheelSubclass := unset
                CPCanvasShuttingDown := unset
                for hwnd in handles
                    SendMessage(0x20A, 120 << 16, 0, hwnd)
            }
            probe.Destroy()
            for hwnd in handles {
                DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd), "Native destruction completes even after wheel globals are released")
                if !releaseGlobals
                    DesktopAssert(!CPCanvasWheelWindows.Has(hwnd), "Normal destruction removes the child/parent tracking entry")
            }
        } finally {
            CPCanvasWheelWindows := savedWindows, CPCanvasWheelSubclass := savedCallback
            CPCanvasShuttingDown := savedShuttingDown
            for hwnd in handles {
                if CPCanvasWheelWindows.Has(hwnd)
                    CPCanvasWheelWindows.Delete(hwnd)
            }
            try probe.Destroy()
        }
    }
}
DesktopTestClosed(*) {
    global DesktopTestCloseCount
    DesktopTestCloseCount += 1
    return true
}
DesktopShown(ctrl) => (DllCall("user32\GetWindowLongPtr", "ptr", ctrl.Hwnd, "int", -16, "ptr") & 0x10000000) != 0
DesktopRect(ctrl) {
    ctrl.GetPos(&x, &y, &w, &h)
    return x "|" y "|" w "|" h
}

DesktopTopChildAtCenter(parentHwnd, ctrl) {
    rect := Buffer(16, 0), point := Buffer(8, 0)
    DllCall("user32\GetWindowRect", "ptr", ctrl.Hwnd, "ptr", rect.Ptr)
    NumPut("int", Floor((NumGet(rect, 0, "int") + NumGet(rect, 8, "int")) / 2), point, 0)
    NumPut("int", Floor((NumGet(rect, 4, "int") + NumGet(rect, 12, "int")) / 2), point, 4)
    DllCall("user32\ScreenToClient", "ptr", parentHwnd, "ptr", point.Ptr)
    return DllCall("user32\ChildWindowFromPointEx", "ptr", parentHwnd,
        "int64", NumGet(point, 0, "int64"), "uint", 7, "ptr")
}
DesktopWindowAtCenter(ctrl) {
    rect := Buffer(16, 0), point := Buffer(8, 0)
    DllCall("user32\GetWindowRect", "ptr", ctrl.Hwnd, "ptr", rect.Ptr)
    NumPut("int", Floor((NumGet(rect, 0, "int") + NumGet(rect, 8, "int")) / 2), point, 0)
    NumPut("int", Floor((NumGet(rect, 4, "int") + NumGet(rect, 12, "int")) / 2), point, 4)
    return DllCall("user32\WindowFromPoint", "int64", NumGet(point, 0, "int64"), "ptr")
}
DesktopHitAt(x, y) {
    global ui
    scale := GetWindowDPI(ui.Hwnd) / 96
    pt := Buffer(8)
    NumPut("int", Round(x * scale), "int", Round(y * scale), pt)
    DllCall("user32\ClientToScreen", "ptr", ui.Hwnd, "ptr", pt)
    packed := (NumGet(pt, 0, "int") & 0xFFFF) | ((NumGet(pt, 4, "int") & 0xFFFF) << 16)
    return SendMessage(0x84, 0, packed, ui.Hwnd)
}
DesktopTestAudioPage() {
    global ui, tab, CPDesktop, ddlAProv, ddlA_GM, ddlTR, ddlAudioTarget, ddlSpeaker
    global btnSpRef, btnAudioTest
    global gAudioInputJob, gAudioInputStatus, gPidAudio, speakerName, iniPath
    a := CPDesktop["audioPage"]
    for size in [[1120, 760], [900, 640], [1400, 820], [820, 560]] {
        audioW := size[1], audioH := size[2]
        ui.Show("NA x-9000 y-9000 w" audioW " h" audioH)
        tab.Value := 2
        CPDesktopLayout(ui, 0, audioW, audioH)
        ToggleAudioControls()
        DesktopAssert(DesktopShown(ddlAProv) && CPDesktopPage(2)["adaptedControls"].Length = 0,
            "Audio uses only its directly owned control tree")
        DesktopAssert(DesktopShown(ddlA_GM) && !DesktopShown(ddlTR), "Audio shows only the selected provider's model")
        DesktopAssert(DesktopShown(a["models"]) && CPDesktop["paint"].Has(a["models"].Hwnd),
            "Audio model Add/Delete rows are replaced by the owned Manage models action")
        DesktopAssert(!DesktopShown(a["troubleshootHelp"]), "Audio troubleshooting starts collapsed")
        DesktopAssert(InStr(a["result"].Text, "Not tested yet") && InStr(a["result"].Text, "Test audio"), "Untested input has useful inline guidance")
        ddlA_GM.GetPos(, &modelY), ddlAudioTarget.GetPos(, &languageY)
        DesktopAssert(audioW >= 1120 ? modelY = languageY : languageY > modelY, "Audio AI fields reflow on narrow windows")
        ddlSpeaker.GetPos(&deviceX, &deviceY, &deviceW)
        btnSpRef.GetPos(&refreshX), btnAudioTest.GetPos(&testX)
        DesktopAssert(deviceX + deviceW + 12 <= refreshX && refreshX + 124 + 12 <= testX, "Device and test actions fit without overlap")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlAProv.Hwnd, "int", 0, "ptr") = ddlA_GM.Hwnd, "Audio provider tabs to the active model")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlSpeaker.Hwnd, "int", 0, "ptr") = btnSpRef.Hwnd, "Audio device tabs to Refresh devices")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", btnAudioTest.Hwnd, "int", 0, "ptr") = a["troubleshoot"].Hwnd, "Test audio tabs to troubleshooting")
        original := DesktopRect(ddlAudioTarget)
        Loop 3
            CPDesktopLayout(ui, 0, audioW, audioH)
        DesktopAssert(DesktopRect(ddlAudioTarget) = original, "Audio layout has no coordinate drift")
        scale := GetWindowDPI(ui.Hwnd) / 96
        ui.GetPos(,, &imageW, &imageH)
        TestCapture("desktop-audio-" audioW ".png", Round(imageW * scale), Round(imageH * scale))
        CPDesktopToggleAudioHelp()
        DesktopAssert(DesktopShown(a["troubleshootHelp"]) && CPCanvasScrollMaxY > 0, "Expanded troubleshooting remains scrollable")
        a["troubleshootHelp"].GetPos(,, &helpW, &helpH)
        DesktopAssert(helpH * scale >= CPMeasureWrappedTextHeight(a["troubleshootHelp"], helpW * scale), "Expanded audio help fits without clipping")
        a["sessionHelp"].GetPos(,, &sessionW, &sessionH)
        DesktopAssert(sessionH * scale >= CPMeasureWrappedTextHeight(a["sessionHelp"], sessionW * scale), "Audio session guidance fits without clipping")
        CPCanvasScrollTo(0, CPCanvasScrollMaxY)
        a["power"].GetPos(, &powerY,, &powerH)
        DesktopAssert(powerY >= CPDesktop["headerH"] + 88 && powerY + powerH <= audioH - 58, "Audio power action stays reachable above footer")
        TestCapture("desktop-audio-" audioW "-help.png", Round(imageW * scale), Round(imageH * scale))
        CPCanvasScrollTo(0, 0)
        CPDesktopToggleAudioHelp()
        ddlAProv.Choose(2), ToggleAudioControls()
        DesktopAssert(!DesktopShown(ddlA_GM) && DesktopShown(ddlTR) && ddlTR.Enabled, "Audio provider switch reveals the correct model")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlAProv.Hwnd, "int", 0, "ptr") = ddlTR.Hwnd, "Audio provider switch updates tab order")
        DesktopAssert(ddlA_GM.Text = "gemini-3.5-live-translate-preview", "Inactive audio model choice is preserved")
        tab.Value := 1
        CPDesktopLayout(ui, 0, audioW, audioH)
        DesktopAssert(!DesktopShown(a["power"]) && !DesktopShown(ddlSpeaker), "Audio controls do not leak onto Game Text Translation")
        tab.Value := 2
        CPDesktopLayout(ui, 0, audioW, audioH)
        DesktopAssert(ddlAProv.Text = "OpenAI" && DesktopShown(ddlTR), "Audio selection survives a page round-trip")
        ddlAProv.Choose(1), ToggleAudioControls()
    }
    ui.Show("NA x-9000 y-9000 w820 h560")
    CPDesktopLayout(ui, 0, 820, 560)
    for result in ["Audio detected. This device is ready.",
        "No audio detected. Check the selected device or the game's audio driver.",
        "Devices refreshed. The selected device is unavailable; reconnect it or choose another input.",
        "Could not open the selected audio device. Refresh devices, check your audio driver and try again.",
        "Could not refresh audio devices. The previous list was kept. Check your audio driver and helper version."] {
        SetAudioTestStatus(result)
        DesktopAssert(a["result"].Text = result, "Audio results are shown in full beside the input controls")
        a["result"].GetPos(,, &resultW, &resultH)
        DesktopAssert(resultH * scale >= CPMeasureWrappedTextHeight(a["result"], resultW * scale), "Audio diagnostic text fits at the narrowest layout")
    }
    SendMessage(0x100, 0x28, 0, ddlAudioTarget.Hwnd)
    SendMessage(0x101, 0x28, 0, ddlAudioTarget.Hwnd)
    DesktopAssert(ddlAudioTarget.Value = 2, "Modern audio dropdown retains native keyboard selection")
    CPShowCombo(ddlSpeaker.Hwnd, true)
    DesktopAssert(CPComboDropped(ddlSpeaker.Hwnd), "Modern device dropdown opens its native list")
    SendMessage(0x100, 0x1B, 0, ddlSpeaker.Hwnd)
    SendMessage(0x101, 0x1B, 0, ddlSpeaker.Hwnd)
    DesktopAssert(!CPComboDropped(ddlSpeaker.Hwnd), "Escape closes the audio device list")
    gAudioInputJob := Map("active", true, "mode", "test")
    AudioInputJobControls()
    SetAudioTestStatus("Listening for audio… Play sound in your game. No AI request is made.")
    DesktopAssert(!ddlSpeaker.Enabled && !btnSpRef.Enabled && !btnAudioTest.Enabled && a["check"].Text = "Checking input…", "Busy diagnostic state disables duplicate input checks")
    gAudioInputJob := Map("active", false)
    AudioInputJobControls()
    AudioInputApplyTestResult("JRPG_AUDIO_TEST:DETECTED")
    DesktopAssert(ddlSpeaker.Enabled && btnAudioTest.Enabled && a["result"].Text = "Audio detected. This device is ready.", "Finished audio test restores controls and displays its result")
    ; Status-only borrowed handle; release it before any Stop test.
    AudioSetProcess(AudioProcessRecord(DllCall("kernel32\OpenProcess", "uint", 0x101000,
        "int", false, "uint", DllCall("kernel32\GetCurrentProcessId"), "ptr"), DllCall("kernel32\GetCurrentProcessId")))
    CPDesktopRefreshAudio()
    DesktopAssert(a["power"].Text = "Stop audio translation" && InStr(a["sessionHelp"].Text, "restart"), "Running session presents Stop and explains pending settings")
    CPDesktop["chrome"]["audio"].Text := "Audio: Off"
    UpdateStatus()
    DesktopAssert(CPDesktop["chrome"]["audio"].Text = "Audio: On", "Immediate audio status refresh also updates the desktop footer")
    try FileDelete(gAudioSessionFile)
    FileAppend(String(gPidAudio), gAudioSessionFile, "UTF-8")
    AudioSetProcess()
    DesktopAssert(!AudioIsRunning(), "A PID-only marker cannot adopt an unrelated live process")
    try FileDelete(gAudioSessionFile)
    gPidAudio := 0
    FileAppend("2147483647", gAudioSessionFile, "UTF-8")
    DesktopAssert(!AudioIsRunning() && !FileExist(gAudioSessionFile), "Stale audio session markers are ignored and removed")
    audioFixturePid := 0
    audioFixtureLog := A_ScriptDir "\desktop-audio-runtime.txt"
    try {
        try FileDelete(audioFixtureLog)
        EnvSet("JRPG_TEST_AUDIO_RUNTIME_LOG", audioFixtureLog)
        EnvSet("JRPG_TEST_AUDIO_RUNTIME_MODE", "wait")
        AudioSetProcess(AudioLaunchProcess(A_AhkPath, A_ScriptDir "\audio_runtime_fixture.ahk"))
        audioFixturePid := gPidAudio
        Sleep(120)
        DesktopAssert(AudioIsRunning() && StopAudioCore(true, false) && !ProcessExist(audioFixturePid),
            "Owned process handle lets the desktop stop the exact worker")
    } finally {
        if audioFixturePid && ProcessExist(audioFixturePid)
            try ProcessClose(audioFixturePid)
        try FileDelete(gAudioSessionFile)
        try FileDelete(audioFixtureLog)
        EnvSet("JRPG_TEST_AUDIO_RUNTIME_LOG", "")
        EnvSet("JRPG_TEST_AUDIO_RUNTIME_MODE", "")
    }
    AudioSetProcess()
    CPDesktopRefreshAudio()
    DesktopAssert(a["power"].Text = "Start audio translation", "Stopped session presents Start")
    ddlAudioTarget.Choose(2), ddlTR.Choose(2), ddlAProv.Choose(2), ToggleAudioControls()
    CPSetAudioDeviceSelection("Game speakers")
    CPDesktopLayout(ui, 0, 1400, 820)
    DesktopAssert(CPDesktopIsCombo(ddlSpeaker.Hwnd) && DesktopShown(ddlTR) && !DesktopShown(ddlA_GM),
        "Modern relayout restores the selected audio model")
    DesktopAssert(ddlTR.Text = "gpt-alternative" && ddlSpeaker.Text = "Game speakers" && ddlAudioTarget.Text = "German (de)",
        "Audio choices survive a modern relayout")
    DesktopAssert(IniRead(iniPath, "cfg", "speakerName") = "Game speakers", "Modern relayout preserves the saved audio device")
    SetAudioTestStatus("Not tested.")
    ddlAProv.Choose(1), ddlTR.Choose(1), ddlAudioTarget.Choose(1), ddlSpeaker.Choose(1)
    speakerName := "[Windows Default]"
    ToggleAudioControls()
    tab.Value := 1
    ui.Show("NA x-9000 y-9000 w1400 h820")
    CPDesktopLayout(ui, 0, 1400, 820)
}
DesktopTestOverlayPages() {
    global ui, tab, CPDesktop, CPDesktopNativePaintHwnds
    global boxBgHex, boxBgHex_EW, nameHex, ddlFont, ddlFont_EW
    for page in [3, 5] {
        overlayPage := CPDesktop["overlayPages"][page], bindings := CPDesktopOverlayBindings(page)
        otherBindings := CPDesktopOverlayBindings(page = 3 ? 5 : 3)
        DesktopAssert(CPDesktopPage(page)["adaptedControls"].Length = 0,
            bindings["title"] " overlay has no adapted legacy controls")
        DesktopAssert(overlayPage["opacitySlider"] = bindings["opacity"]
            && overlayPage["fontChoice"] = bindings["font"]
            && overlayPage["moveResize"] = bindings["position"],
            bindings["title"] " overlay registry owns its shared control aliases")
        originalFont := bindings["font"].Text, originalSize := bindings["size"].Value
        originalOpacity := bindings["opacity"].Value, originalBold := bindings["bold"].Value
        for size in [[1120, 760], [900, 640], [1400, 820], [820, 560]] {
            overlayW := size[1], overlayH := size[2]
            ui.Show("NA x-9000 y-9000 w" overlayW " h" overlayH)
            tab.Value := page
            CPDesktopLayout(ui, 0, overlayW, overlayH)
            DesktopAssert(CPDesktop["chrome"]["title"].Text = "Overlay windows" && InStr(CPDesktop["chrome"]["subtitle"].Text, bindings["title"]), "Shared overlay heading identifies the selected window")
            DesktopAssert(DesktopShown(bindings["opacity"]) && DesktopShown(overlayPage["bg"]), "Modern overlay controls are page-owned and visible")
            DesktopAssert(DesktopShown(bindings["font"]) && !DesktopShown(otherBindings["font"]), "Only the selected overlay's controls are shown")
            DesktopAssert(CPDesktop["inputFrames"].Has(bindings["size"].Hwnd),
                bindings["title"] " font size uses the shared app-colored input frame")
            sizeStyle := DllCall("user32\GetWindowLongPtr", "ptr", bindings["size"].Hwnd, "int", -16, "ptr")
            sizeExStyle := DllCall("user32\GetWindowLongPtr", "ptr", bindings["size"].Hwnd, "int", -20, "ptr")
            DesktopAssert(!(sizeStyle & 0x00800000) && !(sizeExStyle & 0x00000200),
                bindings["title"] " font size removes the bright native edit frame")
            DesktopAssert(CPDesktopNativePaintHwnds.Has(bindings["spinner"].Hwnd)
                && CPDesktopNativePaintHwnds[bindings["spinner"].Hwnd]["kind"] = "spinner",
                bindings["title"] " font-size arrows use the integrated desktop painter")
            DesktopAssert(CPDesktopNativePaintHwnds.Has(bindings["bold"].Hwnd)
                && CPDesktopNativePaintHwnds[bindings["bold"].Hwnd]["kind"] = "checkbox",
                bindings["title"] " Bold option uses the restrained desktop checkbox painter")
            DesktopAssert(CPDesktop["paint"][overlayPage[page = 3 ? "translator" : "explainer"].Hwnd]["selected"], "Overlay selector marks the active window")
            DesktopAssert(CPDesktop["paint"][CPDesktop["chrome"]["overlays"].Hwnd]["selected"], "Overlay sidebar remains selected for both windows")
            DesktopAssert(page = 3 ? overlayPage.Has("name") : !overlayPage.Has("name"), "Speaker-name color belongs only to the Translator")
            for field in (page = 3 ? ["bg", "txt", "name"] : ["bg", "txt"]) {
                DesktopAssert(CPColorSwatchTarget(overlayPage[field].Hwnd) = StrLower(bindings["title"]) ":" field, "Modern color button retains its controller target")
                DesktopAssert(CPDesktop["paint"][overlayPage[field].Hwnd]["color"] = CPOverlayPreference(bindings["title"], field)["value"], "Modern color button displays the configured color")
            }
            DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", overlayPage["explainer"].Hwnd, "int", 0, "ptr") = bindings["opacity"].Hwnd, "Overlay keyboard order reaches opacity after the selectors")
            DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", bindings["opacity"].Hwnd, "int", 0, "ptr") = overlayPage["bg"].Hwnd, "Overlay keyboard order reaches the color buttons")
            DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", bindings["font"].Hwnd, "int", 0, "ptr") = bindings["size"].Hwnd, "Overlay keyboard order reaches font size")
            DesktopAssert(!DesktopShown(bindings["sizeHint"]), "Controller size hint does not clutter the inactive form")
            bindings["font"].GetPos(&fontX,, &fontW)
            bindings["size"].GetPos(&sizeX)
            DesktopAssert(fontX + fontW + 16 <= sizeX, "Overlay font and size fields do not overlap at width " overlayW)
            scale := GetWindowDPI(ui.Hwnd) / 96
            bindings["positionHelp"].GetPos(,, &helpW, &helpH)
            DesktopAssert(helpH * scale >= CPMeasureWrappedTextHeight(bindings["positionHelp"], helpW * scale), "Overlay move/resize instructions fit at every width")
            rect := DesktopRect(bindings["font"])
            Loop 3
                CPDesktopLayout(ui, 0, overlayW, overlayH)
            DesktopAssert(DesktopRect(bindings["font"]) = rect, "Overlay resize does not drift")
            DesktopAssert(bindings["font"].Text = originalFont && bindings["size"].Value = originalSize && bindings["opacity"].Value = originalOpacity && bindings["bold"].Value = originalBold, "Layout preserves each overlay's independent settings")
            ui.GetPos(,, &overlayCaptureW, &overlayCaptureH)
            TestCapture("desktop-overlay-" page "-" overlayW ".png", Round(overlayCaptureW * scale), Round(overlayCaptureH * scale))
            CPCanvasScrollTo(0, CPCanvasScrollMaxY)
            bindings["position"].GetPos(, &moveY,, &moveH)
            DesktopAssert(moveY >= CPDesktop["headerH"] + 88 && moveY + moveH <= overlayH - 58, "Move / resize is reachable above the footer")
            TestCapture("desktop-overlay-" page "-" overlayW "-scrolled.png", Round(overlayCaptureW * scale), Round(overlayCaptureH * scale))
            CPCanvasScrollTo(0, 0)
            ; Exercise the real navigation callback without opening a popup or an overlay.
            ; It queries the client width, which excludes the native scrollbar.
            CPDesktopRelayout()
            navigationRect := DesktopRect(bindings["font"])
            ui.GetClientPos(,, &overlayNavigationClientW)
            tab.Value := 1
            CPDesktopLayout(ui, 0, overlayW, overlayH)
            DesktopAssert(!DesktopShown(overlayPage["translator"]) && !DesktopShown(bindings["position"]), "Overlay controls do not leak onto Game Text Translation")
            CPDesktopOverlayMenu()
            DesktopAssert(tab.Value = page, "Overlay sidebar returns directly to the last selected window")
            overlayReturnedRect := DesktopRect(bindings["font"])
            ui.GetClientPos(,, &overlayReturnedClientW)
            DesktopAssert(overlayReturnedRect = navigationRect,
                "Overlay navigation restores its layout at the actual client width (before "
                navigationRect " / " overlayNavigationClientW ", after " overlayReturnedRect
                " / " overlayReturnedClientW ")")
        }
        CPShowCombo(bindings["font"].Hwnd, true)
        DesktopAssert(CPComboDropped(bindings["font"].Hwnd), "Overlay font dropdown opens its native list")
        SendMessage(0x100, 0x1B, 0, bindings["font"].Hwnd)
        SendMessage(0x101, 0x1B, 0, bindings["font"].Hwnd)
        DesktopAssert(!CPComboDropped(bindings["font"].Hwnd), "Escape closes the overlay font list")
        lower := Buffer(4), upper := Buffer(4)
        SendMessage(0x0470, lower.Ptr, upper.Ptr, bindings["spinner"].Hwnd) ; UDM_GETRANGE32
        DesktopAssert(NumGet(lower, 0, "int") = 6 && NumGet(upper, 0, "int") = (page = 3 ? 128 : 200), "Overlay font-size ranges are unchanged")
        otherFont := otherBindings["font"].Text, otherSize := otherBindings["size"].Value
        SetComboToExistingItem(bindings["font"], GetInstalledFonts(), "Arial")
        bindings["size"].Value := 27, bindings["spinner"].Value := 27
        bindings["opacity"].Value := 150, bindings["percent"].Text := "59%"
        bindings["bold"].Value := !originalBold
        CPDesktopLayout(ui, 0, 1400, 820)
        DesktopAssert(bindings["font"].Text = "Arial" && bindings["size"].Value = 27 && bindings["opacity"].Value = 150 && bindings["percent"].Text = "59%" && bindings["bold"].Value = !originalBold,
            "Overlay choices and opacity readout survive a modern relayout")
        DesktopAssert(otherBindings["font"].Text = otherFont && otherBindings["size"].Value = otherSize, "Editing one overlay never changes the other's typography")
        bindings["font"].Choose(originalFont), bindings["size"].Value := originalSize, bindings["spinner"].Value := originalSize
        bindings["opacity"].Value := originalOpacity, bindings["bold"].Value := originalBold
    }
    priorColor := boxBgHex
    boxBgHex := "2468AC"
    CPDesktopRefreshOverlayColors()
    DesktopAssert(CPDesktop["overlayPages"][3]["bg"].Text = "Window color: #2468AC" && CPDesktop["overlayPages"][5]["bg"].Text = "Window color: #" boxBgHex_EW, "Color display updates independently and includes an accessible setting name")
    boxBgHex := priorColor
    CPDesktopRefreshOverlayColors()
    ; Check the two selector buttons themselves, not only direct page assignment.
    SendMessage(0xF5, 0, 0, CPDesktop["overlayPages"][5]["translator"].Hwnd)
    Sleep(20)
    DesktopAssert(tab.Value = 3, "Translator selector opens the Translator settings")
    SendMessage(0xF5, 0, 0, CPDesktop["overlayPages"][3]["explainer"].Hwnd)
    Sleep(20)
    DesktopAssert(tab.Value = 5, "Explainer selector opens the Explainer settings")
    tab.Value := 1
    ui.Show("NA x-9000 y-9000 w1400 h820")
    CPDesktopLayout(ui, 0, 1400, 820)
}

DesktopExplanationClicked(*) {
    global DesktopExplanationClicks
    DesktopExplanationClicks += 1
}

DesktopTestExplanationPage() {
    global ui, tab, CPDesktop, ddlEProv, ddlEGem, ddlEOpenAI, ddlEPr
    global btnExplainNow, btnOpenStudyLibrary, DesktopExplanationClicks
    global saveLibraryChk, saveLibraryScreenshotsChk, saveExplChk, chkOpenEW, chkTop_EW, iniPath
    global ddlProv, ddlPrompt, ddlAProv, ddlAudioTarget
    expPage := CPDesktop["explanationPage"]
    DesktopAssert(CPDesktopPage(4)["adaptedControls"].Length = 0, "Explanation page has no adapted legacy controls")
    DesktopAssert(expPage["providerChoice"] = ddlEProv && expPage["promptChoice"] = ddlEPr
        && expPage["createExplanation"] = btnExplainNow && expPage["saveLibrary"] = saveLibraryChk,
        "Explanation registry owns its shared control aliases")
    otherChoices := ddlProv.Text "|" ddlPrompt.Text "|" ddlAProv.Text "|" ddlAudioTarget.Text
    for size in [[1120, 760], [900, 640], [1400, 820], [820, 560]] {
        w := size[1], h := size[2]
        ui.Show("NA x-9000 y-9000 w" w " h" h)
        tab.Value := 4
        CPDesktopLayout(ui, 0, w, h)
        ToggleExplanationControls()
        DesktopAssert(DesktopShown(ddlEProv), "Explanation uses its page-owned grouped layout")
        DesktopAssert(DesktopShown(ddlEGem) && !DesktopShown(ddlEOpenAI), "Explanation shows only the active provider's model")
        DesktopAssert(DesktopShown(expPage["models"]) && DesktopShown(expPage["prompts"]), "Explanation management is consolidated into page-owned links")
        DesktopAssert(!expPage.Has("editPrompt") && expPage["prompts"].Text = "Manage…",
            "Explanation exposes prompt editing only through Manage")
        DesktopAssert(CPDesktop["paint"][CPDesktop["chrome"]["explanation"].Hwnd]["selected"], "Explanation sidebar row is selected")
        DesktopAssert(!DesktopShown(chkOpenEW) && !DesktopShown(chkTop_EW), "Explanation startup choices begin collapsed")
        DesktopAssert(CPDesktop["paint"][btnExplainNow.Hwnd]["kind"] = "primary" && btnExplainNow.Text = "Explain latest text", "Explanation has one clear primary action")
        ddlEGem.GetPos(, &modelY), ddlEPr.GetPos(, &promptY)
        DesktopAssert(w >= 1120 ? modelY = promptY : promptY > modelY, "Explanation AI fields reflow at narrow widths")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlEProv.Hwnd, "int", 0, "ptr") = ddlEGem.Hwnd, "Explanation keyboard order reaches the active model")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlEGem.Hwnd, "int", 0, "ptr") = expPage["models"].Hwnd, "Explanation keyboard order reaches model management")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlEPr.Hwnd, "int", 0, "ptr") = expPage["prompts"].Hwnd,
            "Explanation prompt selector tabs directly to Manage")
        btnExplainNow.GetPos(&actionX, &actionY, &actionW)
        btnOpenStudyLibrary.GetPos(&libraryX, &libraryY)
        DesktopAssert(libraryY = actionY && libraryX >= actionX + actionW + 12, "Explanation and Library actions do not overlap")
        rect := DesktopRect(ddlEPr)
        Loop 3
            CPDesktopLayout(ui, 0, w, h)
        DesktopAssert(DesktopRect(ddlEPr) = rect, "Explanation resize has no coordinate drift")
        scale := GetWindowDPI(ui.Hwnd) / 96
        for key in ["actionHelp", "libraryHelp", "screenshotsHelp", "plainTextHelp"] {
            expPage[key].GetPos(,, &helpW, &helpH)
            DesktopAssert(helpH * scale >= CPMeasureWrappedTextHeight(expPage[key], helpW * scale), "Explanation " key " fits at width " w)
        }
        ddlEProv.Choose(2), ToggleExplanationControls()
        DesktopAssert(DesktopShown(ddlEOpenAI) && !DesktopShown(ddlEGem) && ddlEGem.Text = "gemini-explanation", "Explanation provider change retains the inactive model")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlEProv.Hwnd, "int", 0, "ptr") = ddlEOpenAI.Hwnd, "Explanation keyboard order follows provider changes")
        ui.GetPos(,, &explanationCaptureW, &explanationCaptureH)
        TestCapture("desktop-explanation-" w ".png", Round(explanationCaptureW * scale), Round(explanationCaptureH * scale))
        CPDesktopToggleExplanationStartup()
        DesktopAssert(DesktopShown(chkOpenEW) && DesktopShown(chkTop_EW), "Explanation startup choices expand")
        expPage["startupHelp"].GetPos(,, &helpW, &helpH)
        DesktopAssert(helpH * scale >= CPMeasureWrappedTextHeight(expPage["startupHelp"], helpW * scale), "Explanation startup guidance fits")
        CPCanvasScrollTo(0, CPCanvasScrollMaxY)
        chkTop_EW.GetPos(, &startupY,, &startupH)
        DesktopAssert(startupY >= CPDesktop["headerH"] + 88 && startupY + startupH <= h - 58, "Explanation startup choices are reachable above the fixed footer")
        TestCapture("desktop-explanation-" w "-expanded.png", Round(explanationCaptureW * scale), Round(explanationCaptureH * scale))
        CPCanvasScrollTo(0, 0)
        CPDesktopToggleExplanationStartup()
        tab.Value := 1
        CPDesktopLayout(ui, 0, w, h)
        DesktopAssert(!DesktopShown(expPage["models"]) && !DesktopShown(btnExplainNow), "Explanation actions do not leak onto Game Text Translation")
        tab.Value := 4
        CPDesktopLayout(ui, 0, w, h)
        DesktopAssert(DesktopRect(ddlEPr) = rect && !DesktopShown(CPDesktop["audioPage"]["power"]), "Explanation navigation restores geometry without other-page controls")
        ddlEProv.Choose(1), ToggleExplanationControls()
    }
    SendMessage(0x100, 0x28, 0, ddlEPr.Hwnd)
    SendMessage(0x101, 0x28, 0, ddlEPr.Hwnd)
    DesktopAssert(ddlEPr.Text = "detailed_grammar", "Explanation prompt retains native keyboard selection")
    CPShowCombo(ddlEPr.Hwnd, true)
    DesktopAssert(CPComboDropped(ddlEPr.Hwnd), "Explanation prompt opens its native list")
    SendMessage(0x100, 0x1B, 0, ddlEPr.Hwnd)
    SendMessage(0x101, 0x1B, 0, ddlEPr.Hwnd)
    DesktopAssert(!CPComboDropped(ddlEPr.Hwnd), "Escape closes the explanation prompt list")
    ; Real preference setter, isolated to the temporary INI. No AI or library writes.
    CPSetExplanationPreference("library", false)
    DesktopAssert(!saveLibraryScreenshotsChk.Enabled && saveLibraryScreenshotsChk.Value = 1, "Disabling Library saves preserves but disables the screenshot preference")
    CPDesktopRelayout()
    DesktopAssert(!saveLibraryScreenshotsChk.Enabled, "Explanation relayout keeps dependent screenshot control disabled")
    CPSetExplanationPreference("plainText", true)
    DesktopAssert(saveExplChk.Value && IniRead(iniPath, "cfg", "saveExplains") = 1, "Plain-text saving remains independent of the Library")
    CPSetExplanationPreference("library", true)
    CPSetExplanationPreference("screenshots", false)
    DesktopAssert(saveLibraryScreenshotsChk.Enabled && IniRead(iniPath, "cfg", "studyLibraryScreenshots") = 0, "Library and screenshot preferences use their existing settings")
    CPSetExplanationPreference("alwaysOnTop", false)
    DesktopAssert(IniRead(iniPath, "cfg_explainer", "winTop") = 0, "Explanation topmost preference keeps its original settings scope")
    ; The page-owned action retains its production event path; the harness swaps
    ; only the final AI callback for this local counter.
    SendMessage(0xF5, 0, 0, btnExplainNow.Hwnd)
    Sleep(20)
    DesktopAssert(DesktopExplanationClicks = 1, "Page-owned explanation action dispatches its Click handler")
    ddlEProv.Choose(2), ddlEOpenAI.Choose(2), ToggleExplanationControls()
    CPDesktopLayout(ui, 0, 1400, 820)
    DesktopAssert(ddlEPr.Text = "detailed_grammar" && ddlEOpenAI.Text = "gpt-explanation-alternative" && DesktopShown(ddlEOpenAI) && !DesktopShown(ddlEGem),
        "Explanation selections survive a modern relayout")
    DesktopAssert(saveLibraryChk.Value && !saveLibraryScreenshotsChk.Value && saveExplChk.Value,
        "Explanation saving preferences survive a modern relayout")
    DesktopAssert(otherChoices = ddlProv.Text "|" ddlPrompt.Text "|" ddlAProv.Text "|" ddlAudioTarget.Text, "Explanation changes never alter translation or audio choices")
    ddlEProv.Choose(1), ddlEOpenAI.Choose(1), ddlEPr.Choose(1), ToggleExplanationControls()
    tab.Value := 1
    ui.Show("NA x-9000 y-9000 w1400 h820")
    CPDesktopLayout(ui, 0, 1400, 820)
}

DesktopBuildOrganizeFixtures() {
    global hotkeyActions, hotkeyLabels, hotkeyDefaults
    hotkeyActions := ["screenshot_translate", "explain_last_translation", "hide_show_translator", "hide_show_explainer", "hide_show_control_panel", "take_screenshot", "screenshot_translation", "launch_explainer_request", "recapture_region", "start_stop_audio"]
    fixtureLabels := ["Capture + Translate", "Explain last translation", "Show/Hide Translator", "Show/Hide Explainer", "Show/Hide Control Panel", "Make Capture", "Translate Captures", "Launch Explainer + Req.", "Recapture Region", "Audio Translation On/Off"]
    hotkeyLabels := Map(), hotkeyDefaults := Map()
    for index, action in hotkeyActions {
        hotkeyLabels[action] := fixtureLabels[index]
        hotkeyDefaults[action] := "^+" index
    }
}

DesktopTestOrganizePages() {
    global ui, tab, tabNames, CPDesktop, CPTabVisiblePages, CPCanvasScrollMaxY, CPControlsCurrentView
    global ddlGameProfile, ddlStartupOverlays, txtGameProfileState, btnGameProfileApply, btnGameProfileSave, btnGameProfileAdd, btnGameProfileDelete
    global ddlENG, ddlJPG, btnENG_Edit, btnJPG_Edit, chkUseTerminologyOverrides
    global eGemini, eOpenAI, cbApiInApp, btnSaveEnv, btnDelEnv, btnOpenEnvVars
    global hkEdits, hkBtnChg, hotkeyActions, hotkeyLabels, hkConflictText, CPControllerBindingEdits, CPControllerAssignButtons, cbControllerInputsEnabled, cbControllerDpadNavigationEnabled
    global rbControlsKeyboard, rbControlsController, CPHotkeyNotice, CPApiKeysNotice
    global CPToastGui
    global TestApiKeysLoaded, envSavedOpenAI, envSavedGemini
    global iniPath, pythonExe, directModelOutput, debugMode
    global chkGuess
    termPage := CPDesktop["organizePages"][6]
    DesktopAssert(CPDesktopPage(6)["adaptedControls"].Length = 0,
        "Terminology page has no adapted legacy controls")
    DesktopAssert(termPage["enabled"] = chkUseTerminologyOverrides
        && termPage["localChoice"] = ddlENG && termPage["modelChoice"] = ddlJPG
        && termPage["localManage"] = btnENG_Edit && termPage["modelManage"] = btnJPG_Edit,
        "Terminology registry owns its shared control aliases")
    profilesPage := CPDesktop["organizePages"][7]
    DesktopAssert(CPDesktopPage(7)["adaptedControls"].Length = 0,
        "Profiles page has no adapted legacy controls")
    DesktopAssert(profilesPage["profileChoice"] = ddlGameProfile
        && profilesPage["startupChoice"] = ddlStartupOverlays
        && profilesPage["profileState"] = txtGameProfileState
        && profilesPage["newProfile"] = btnGameProfileAdd
        && profilesPage["saveProfile"] = btnGameProfileSave
        && profilesPage["applyProfile"] = btnGameProfileApply
        && profilesPage["deleteProfile"] = btnGameProfileDelete,
        "Profiles registry owns its shared control aliases")
    controlsPage := CPDesktop["organizePages"][8]
    firstAction := hotkeyActions[1]
    DesktopAssert(CPDesktopPage(8)["adaptedControls"].Length = 0,
        "Controls page has no adapted legacy controls")
    DesktopAssert(controlsPage["keyboard"] = rbControlsKeyboard
        && controlsPage["gamepad"] = rbControlsController
        && controlsPage["keyboardBinding_" firstAction] = hkEdits[firstAction]
        && controlsPage["controllerBinding_" firstAction] = CPControllerBindingEdits[firstAction]
        && controlsPage["controllerEnabled"] = cbControllerInputsEnabled
        && controlsPage["dpadNavigation"] = cbControllerDpadNavigationEnabled,
        "Controls registry owns its shared control aliases")
    for ctrl in [hkEdits[firstAction], CPControllerBindingEdits[firstAction]] {
        style := WinGetStyle(ctrl.Hwnd), exStyle := WinGetExStyle(ctrl.Hwnd)
        DesktopAssert(!(style & 0x00800000) && !(style & 0x00200000)
            && !(exStyle & 0x00000200),
            "Binding display has no bright native edge or scrollbar")
        DesktopAssert(CPDesktop["inputFrames"].Has(ctrl.Hwnd)
            && CPDesktop["paint"][CPDesktop["inputFrames"][ctrl.Hwnd].Hwnd]["kind"] = "fieldFrame",
            "Binding display uses the app-painted field frame")
    }
    apiPage := CPDesktop["organizePages"][9]
    DesktopAssert(CPDesktopPage(9)["adaptedControls"].Length = 0,
        "API Keys page has no adapted legacy controls")
    DesktopAssert(apiPage["inAppEntry"] = cbApiInApp
        && apiPage["geminiKey"] = eGemini && apiPage["openAIKey"] = eOpenAI
        && apiPage["saveKeys"] = btnSaveEnv && apiPage["deleteEnv"] = btnDelEnv
        && apiPage["openEnvironmentVariables"] = btnOpenEnvVars,
        "API Keys registry owns its shared control aliases")
    DesktopAssert(TestApiKeysLoaded,
        "Page-owned API key controls load and enable an existing .env")
    DesktopAssert(!CPDesktopPage(10) && !CPDesktop["organizePages"].Has(10),
        "The retired Paths page is absent from the desktop registry")
    for page in [6, 7, 8, 9] {
        tab.Value := page
        for dimensions in [[1120, 760], [820, 560], [1400, 820]] {
            testWidth := dimensions[1], testHeight := dimensions[2]
            ui.Show("NA x-9000 y-9000 w" testWidth " h" testHeight)
            CPDesktopLayout(ui, 0, testWidth, testHeight)
            p := CPDesktop["organizePages"][page]
            DesktopAssert(CPDesktop["chrome"]["title"].Text = (page = 7 ? "Profiles" : "Settings"), "Organize pages have a clear heading")
            DesktopAssert(CPDesktop["paint"][CPDesktop["chrome"][page = 7 ? "profiles" : "settings"].Hwnd]["selected"], "Organize sidebar selection stays correct")
            for otherPage, controls in CPDesktop["organizePages"] {
                if otherPage != page
                    for key, ctrl in controls
                        DesktopAssert(!DesktopShown(ctrl), "Inactive organize controls stay hidden")
            }
            if page != 7 {
                DesktopAssert(DesktopShown(p["controlsTab"]) && DesktopShown(p["termsTab"])
                    && DesktopShown(p["apiTab"]), "All Settings sections are directly reachable")
                DesktopAssert(!p.Has("classic"), "Settings exposes no classic-layout action")
            }
            for panelKey in CPDesktop["organizePanels"][page] {
                p[panelKey].GetPos(&cardX, &cardY, &cardW, &cardH)
                DesktopAssert(cardX >= 208 && cardX + cardW <= testWidth - 24 && cardH >= 0, "Cards fit the viewport width")
            }
            for ctrl in CPDesktop["organizeStops"] {
                if !DesktopShown(ctrl)
                    continue
                ctrl.GetPos(&itemX,, &itemW)
                DesktopAssert(itemX >= 208 && itemX + itemW <= testWidth - 24, "Actions and inputs remain inside the content width")
            }
            for key, ctrl in p {
                if InStr(key, "Help") && DesktopShown(ctrl) {
                    ctrl.GetPos(,, &helpWidth, &helpHeight)
                    dpi := GetWindowDPI(ui.Hwnd) / 96
                    DesktopAssert(helpHeight * dpi >= CPMeasureWrappedTextHeight(ctrl, helpWidth * dpi), "Help text has enough wrapped height")
                }
            }
            sample := page = 7 ? ddlGameProfile : p["controlsTab"]
            originalRect := DesktopRect(sample)
            Loop 3
                CPDesktopLayout(ui, 0, testWidth, testHeight)
            DesktopAssert(originalRect = DesktopRect(sample), "Organize relayout does not accumulate drift")
            ui.GetPos(,, &shotW, &shotH)
            dpi := GetWindowDPI(ui.Hwnd) / 96
            TestCapture("desktop-organize-" page "-" testWidth ".png", Round(shotW * dpi), Round(shotH * dpi))
            CPCanvasScrollTo(0, CPCanvasScrollMaxY)
            lastStop := CPDesktop["organizeStops"][-1]
            lastStop.GetPos(, &lastY,, &lastH)
            DesktopAssert(lastY >= CPDesktop["headerH"] + 88 && lastY + lastH <= testHeight - 58, "Last action can be reached above the footer")
            TestCapture("desktop-organize-" page "-" testWidth "-scrolled.png", Round(shotW * dpi), Round(shotH * dpi))
            CPCanvasScrollTo(0, 0)
        }
        CPDesktopLayout(ui, 0, 1400, 820)
        DesktopAssert(CPDesktopActive() && tab.Value = page && !CPDesktop["organizePages"][page].Has("classic"),
            "Modern-only relayout retains the selected organize page")
    }
    CPDesktopNavigate(7)
    activeProfileBeforeSelection := IniRead(iniPath, "game_profiles", "active", "")
    SetComboToExistingItem(ddlGameProfile, ListGameProfiles(), "Alternate profile"), GameProfileUpdateSummary()
    DesktopAssert(InStr(txtGameProfileState.Text, "Alternate profile")
        && IniRead(iniPath, "game_profiles", "active", "") = activeProfileBeforeSelection,
        "Selecting a profile does not apply it")
    dirtyFixtureName := "Header selector dirty fixture"
    dirtyFixturePath := GameProfilePath(dirtyFixtureName)
    originalGuessValue := chkGuess.Value
    IniWriteRetry(originalGuessValue ? 1 : 0, dirtyFixturePath, "screenshot", "highlightGuessed")
    DesktopAssert(!GameProfileHasUnsavedChanges(dirtyFixtureName),
        "Saved profile values match the current header-switch state")
    chkGuess.Value := originalGuessValue ? 0 : 1
    DesktopAssert(GameProfileHasUnsavedChanges(dirtyFixtureName),
        "Header profile switching detects a changed profile-owned setting")
    chkGuess.Value := originalGuessValue
    try FileDelete(dirtyFixturePath)
    try FileDelete(dirtyFixturePath ".bak")
    IniWriteRetry(activeProfileBeforeSelection, iniPath, "game_profiles", "active")
    RefreshGameProfilesList("Alternate profile")
    ddlStartupOverlays.Choose(4)
    ddlENG.Choose(2), ddlJPG.Choose(1)
    CPDesktopNavigate(8)
    SendMessage(0xF5, 0, 0, CPDesktop["organizePages"][8]["gamepad"].Hwnd)
    Sleep(10)
    DesktopAssert(CPControlsCurrentView = "controller"
        && CPDesktop["paint"][rbControlsController.Hwnd]["selected"],
        "Modern input selector updates the shared controller view")
    for action in hotkeyActions
        DesktopAssert(DesktopShown(CPControllerBindingEdits[action]) && !DesktopShown(hkEdits[action]), "Only controller bindings are shown")
    DesktopAssert(DesktopShown(cbControllerInputsEnabled) && DesktopShown(cbControllerDpadNavigationEnabled), "Controller options remain reachable")
    ui.Show("NA x-9000 y-9000 w820 h560"), CPDesktopLayout(ui, 0, 820, 560)
    ui.GetPos(,, &shotW, &shotH)
    TestCapture("desktop-controller-820.png", Round(shotW * dpi), Round(shotH * dpi))
    CPCanvasScrollTo(0, CPCanvasScrollMaxY)
    TestCapture("desktop-controller-820-scrolled.png", Round(shotW * dpi), Round(shotH * dpi))
    CPDesktopNavigate(1), CPDesktopSettingsMenu()
    DesktopAssert(tab.Value = 8 && CPControlsCurrentView = "controller", "Settings returns directly to its last section and input view")
    CPSetControlsView("keyboard", false)
    previousBinding := hkEdits[hotkeyActions[1]].Value
    hkEdits[hotkeyActions[1]].Value := hkEdits[hotkeyActions[2]].Value
    DesktopAssert(Hotkeys_ShowConflicts() = 1, "Original hotkey conflict detection remains active")
    DesktopAssert(CPDesktop["bindingConflicts"].Has(hkEdits[hotkeyActions[1]].Hwnd) && CPDesktop["bindingConflicts"].Has(hkEdits[hotkeyActions[2]].Hwnd), "Both conflicting fields retain visual emphasis")
    hkConflictText.GetPos(,, &bannerW, &bannerH)
    dpi := GetWindowDPI(ui.Hwnd) / 96
    DesktopAssert(bannerH * dpi >= CPMeasureWrappedTextHeight(hkConflictText, bannerW * dpi), "Conflict banner expands without clipping")
    hkEdits[hotkeyActions[1]].Value := previousBinding
    DesktopAssert(Hotkeys_ShowConflicts() = 0 && CPDesktop["bindingConflicts"].Count = 0, "Resolving a duplicate clears its modern emphasis")
    DesktopAssert(CPHotkeyNoticeText(hotkeyActions[1], "saved") = "Keyboard shortcut saved · " hotkeyLabels[hotkeyActions[1]]
        && CPHotkeyNoticeText(hotkeyActions[1], "removed") = "Keyboard shortcut removed · " hotkeyLabels[hotkeyActions[1]]
        && CPHotkeyNoticeText(hotkeyActions[1], "default") = "Default keyboard shortcut restored · " hotkeyLabels[hotkeyActions[1]],
        "Keyboard shortcut feedback distinguishes save, removal and default restoration")
    hotkeyNotice := CPHotkeyNoticeText(hotkeyActions[1], "removed")
    CPHotkeySetNotice(hotkeyNotice)
    DesktopAssert(CPHotkeyNotice = hotkeyNotice && CPDesktop["chrome"]["subtitle"].Text = hotkeyNotice,
        "Keyboard shortcut feedback appears inside the visible desktop UI")
    CPHotkeyClearNotice(hotkeyNotice)
    DesktopAssert(CPHotkeyNotice = "" && CPDesktop["chrome"]["subtitle"].Text = CPDesktopPageSubtitle(8),
        "Keyboard shortcut feedback restores the normal page subtitle cleanly")
    ui.Show("NA x-9000 y-9000 w1120 h760")
    CPDesktopNavigate(9)
    CPDesktopLayout(ui, 0, 1120, 760)
    CPCanvasScrollTo(0, 0)
    for ctrl in [eGemini, eOpenAI] {
        DesktopAssert(DesktopWindowAtCenter(ctrl) = ctrl.Hwnd,
            "API key edit receives real pointer input through its decorative field frame")
        passwordChar := SendMessage(0xD2, 0, 0, ctrl.Hwnd)
        DesktopAssert(passwordChar != 0, "API key field keeps its password mask (character: " passwordChar ")")
        DesktopAssert(ctrl.Enabled, "Opening API Keys refreshes and enables a saved in-app key")
        style := WinGetStyle(ctrl.Hwnd), exStyle := WinGetExStyle(ctrl.Hwnd)
        DesktopAssert(!(style & 0x00800000) && !(style & 0x00200000)
            && !(exStyle & 0x00000200),
            "API key field has no bright native edge or vertical arrows")
        DesktopAssert(CPDesktop["inputFrames"].Has(ctrl.Hwnd),
            "API key field uses the app-painted field frame")
    }
    DesktopAssert(cbApiInApp.Value = 1 && !btnSaveEnv.Enabled && btnDelEnv.Enabled
        && eGemini.Value = "synthetic-gemini" && eOpenAI.Value = "synthetic-openai",
        "API page navigation reloads both masked values and their enabled state from .env")
    cbApiInApp.Value := 0
    ToggleApiKeyControls()
    DesktopAssert(!eGemini.Enabled && !eOpenAI.Enabled && !btnDelEnv.Enabled,
        "Explicitly turning off in-app entry still disables its editors")
    cbApiInApp.Value := 1
    ToggleApiKeyControls()
    DesktopAssert(eGemini.Enabled && eOpenAI.Enabled && btnDelEnv.Enabled && !btnSaveEnv.Enabled,
        "In-app key enablement restores both masked fields without creating a false dirty state")
    savedGeminiBeforeEdit := eGemini.Value
    DllCall("user32\SetFocus", "ptr", eGemini.Hwnd)
    SendMessage(0x00B1, StrLen(savedGeminiBeforeEdit), StrLen(savedGeminiBeforeEdit), eGemini.Hwnd) ; EM_SETSEL
    SendMessage(0x0102, Ord("x"), 0, eGemini.Hwnd) ; WM_CHAR through the native editor
    Sleep(10)
    DesktopAssert(eGemini.Value = savedGeminiBeforeEdit "x" && btnSaveEnv.Enabled,
        "A pointer-reachable API key edit accepts keyboard input and enables Save keys")
    CPDesktopNavigate(9)
    DesktopAssert(eGemini.Value = savedGeminiBeforeEdit "x" && btnSaveEnv.Enabled,
        "Re-entering API Keys does not overwrite a deliberate unsaved desktop edit")
    SaveApiEnv()
    savedBody := FileRead(envPath, "UTF-8")
    DesktopAssert(ParseEnvLine(savedBody, "GEMINI_API_KEY") = savedGeminiBeforeEdit "x"
        && ParseEnvLine(savedBody, "GOOGLE_API_KEY") = savedGeminiBeforeEdit "x"
        && !btnSaveEnv.Enabled,
        "Saving from the desktop editor persists both Gemini aliases and clears dirty state")
    DesktopAssert(CPApiKeysNotice = "In-app API keys saved."
        && CPDesktop["chrome"]["subtitle"].Text = CPApiKeysNotice
        && !IsObject(CPToastGui),
        "API key save feedback stays inside the Settings page without a floating toast")
    ui.GetPos(,, &noticeW, &noticeH)
    noticeDpi := GetWindowDPI(ui.Hwnd) / 96
    TestCapture("desktop-api-save-notice-1120.png",
        Round(noticeW * noticeDpi), Round(noticeH * noticeDpi))
    CPApiKeysClearNotice(CPApiKeysNotice)
    DesktopAssert(CPApiKeysNotice = ""
        && CPDesktop["chrome"]["subtitle"].Text = CPDesktopPageSubtitle(9),
        "API key feedback restores the normal Settings subtitle")
    CPBigBoxWriteApiKey("gemini", "synthetic-fullscreen-gemini")
    eGemini.Value := "", eOpenAI.Value := "", cbApiInApp.Value := 0
    ToggleApiKeyControls()
    CPDesktopNavigate(9)
    DesktopAssert(eGemini.Value = "synthetic-fullscreen-gemini"
        && eOpenAI.Value = "synthetic-openai" && cbApiInApp.Value = 1
        && eGemini.Enabled && eOpenAI.Enabled && !btnSaveEnv.Enabled,
        "Desktop API page recovers both providers after a fullscreen save and stale hidden state")
    for ctrl in [eGemini, eOpenAI]
        DesktopAssert(SendMessage(0xD2, 0, 0, ctrl.Hwnd) != 0,
            "Fullscreen-to-desktop synchronization preserves the password mask")
    cbApiInApp.Value := 0
    ToggleApiKeyControls()
    DesktopAssert(CPTabVisiblePages.Length = 9 && tabNames.Length = 9,
        "Desktop navigation contains no Paths tab")
    DesktopAssert(IniRead(iniPath, "cfg", "pythonExe", pythonExe) = pythonExe
        && Integer(IniRead(iniPath, "cfg", "directModelOutput", directModelOutput)) = directModelOutput
        && Integer(IniRead(iniPath, "cfg", "debugMode", debugMode)) = debugMode,
        "Advanced runtime settings remain available through control.ini")
    DesktopAssert(ddlGameProfile.Text = "Alternate profile" && ddlStartupOverlays.Value = 4 && ddlENG.Value = 2 && ddlJPG.Value = 1, "Independent profile, startup and glossary selections are preserved")
    tab.Value := 1
    ui.Show("NA x-9000 y-9000 w1400 h820"), CPDesktopLayout(ui, 0, 1400, 820)
}

DesktopWheelPoint(ctrl) {
    rect := Buffer(16)
    DllCall("user32\GetWindowRect", "ptr", ctrl.Hwnd, "ptr", rect)
    x := NumGet(rect, 0, "int") + 10, y := NumGet(rect, 4, "int") + 10
    return (x & 0xFFFF) | ((y & 0xFFFF) << 16)
}

DesktopTestNativeWheelAndFocus() {
    global ui, tab, CPDesktop, ddlENG, ddlJPG, slTrans
    global CPCanvasScrollY, CPCanvasScrollMaxY, CPCanvasPendingScrollY, CPCanvasPendingScrollValid
    global CPCanvasWheelWindows
    ui.Show("NA x-9000 y-9000 w900 h640")
    CPDesktopNavigate(8), CPDesktopLayout(ui, 0, 900, 640)
    for spec in [[6, "termsTab"], [9, "apiTab"], [8, "controlsTab"]] {
        beforePage := tab.Value
        button := CPDesktop["organizePages"][beforePage][spec[2]]
        DllCall("user32\SetFocus", "ptr", button.Hwnd)
        SendMessage(0xF5, 0, 0, button.Hwnd) ; Same Click callback as mouse/keyboard activation.
        Sleep(20)
        target := CPDesktop["organizePages"][spec[1]][spec[2]]
        DesktopAssert(tab.Value = spec[1] && DllCall("user32\GetFocus", "ptr") = target.Hwnd, "Settings focus follows the selected section, not its first button")
        DesktopAssert(CPDesktop["paint"][target.Hwnd]["selected"], "Only the selected section receives initial focus")
        if spec[1] != 8
            DesktopAssert(DllCall("user32\GetFocus", "ptr") != CPDesktop["organizePages"][spec[1]]["controlsTab"].Hwnd, "Controls has no extra focus outline on another page")
    }
    CPDesktopNavigate(6)
    CPDesktopLayout(ui, 0, 900, 640)
    point := DesktopWheelPoint(ddlENG)
    selections := ddlENG.Text "|" ddlJPG.Text
    DesktopAssert(CPCanvasWheelWindows.Has(ui.Hwnd) && CPCanvasWheelWindows.Has(ddlENG.Hwnd), "Native wheel routing covers parent and child windows")
    for destination in [ui.Hwnd, ddlENG.Hwnd] {
        CPCanvasCancelQueuedScroll(), CPCanvasScrollTo(0, 0)
        Critical "On" ; Deliver a burst before the 16-ms frame timer fires.
        try {
            Loop 240
                SendMessage(0x20A, (-120 & 0xFFFF) << 16, point, destination)
            DesktopAssert(CPCanvasPendingScrollValid && CPCanvasPendingScrollY = CPCanvasScrollMaxY, "Native wheel burst accumulates and clamps at the bottom")
            DesktopAssert(CPCanvasScrollY = 0, "A burst does not repaint/move content for every wheel message")
            Loop 240
                SendMessage(0x20A, 120 << 16, point, destination)
            DesktopAssert(CPCanvasPendingScrollY = 0, "Reverse wheel burst immediately returns toward the top")
            Loop 5
                SendMessage(0x20A, (-24 & 0xFFFF) << 16, point, destination)
            DesktopAssert(Abs(CPCanvasPendingScrollY - 48) < 0.01, "High-resolution deltas accumulate into one normal wheel step")
            CPCanvasFlushQueuedScroll()
            DesktopAssert(CPCanvasScrollY = 48, "Native wheel burst flushes to the expected scroll position")
        } finally {
            Critical "Off"
            CPCanvasCancelQueuedScroll()
        }
    }
    DesktopAssert(ddlENG.Text "|" ddlJPG.Text = selections, "Wheel over a closed dropdown scrolls the page without changing its selection")
    CPCanvasScrollTo(0, 0)
    DllCall("user32\SetFocus", "ptr", ddlENG.Hwnd)
    SendMessage(0x14F, 1, 0, ddlENG.Hwnd) ; CB_SHOWDROPDOWN
    DesktopAssert(!CPCanvasAcceptsWheel(ddlENG.Hwnd, DesktopWheelPoint(ddlENG)), "An open dropdown retains native wheel input")
    SendMessage(0x14F, 0, 0, ddlENG.Hwnd)
    CPCanvasQueueWheel(-120)
    CPDesktopNavigate(9)
    Sleep(30)
    DesktopAssert(!CPCanvasPendingScrollValid && CPCanvasScrollY = 0, "A queued wheel frame never scrolls the next page")
    CPDesktopNavigate(3)
    DesktopAssert(!CPCanvasAcceptsWheel(slTrans.Hwnd, DesktopWheelPoint(slTrans)), "Opacity slider retains its own wheel behavior")
    CPDesktopNavigate(6)
    CPCanvasScrollTo(0, 0)
    Critical "On"
    try {
        SendMessage(0x20A, (-120 & 0xFFFF) << 16, DesktopWheelPoint(ddlENG), ddlENG.Hwnd)
        DesktopAssert(CPCanvasPendingScrollValid, "Modern layout uses non-hotkey wheel routing")
    } finally {
        Critical "Off"
        CPCanvasCancelQueuedScroll()
    }
    CPDesktopNavigate(1)
}

DesktopTestScrollTransitionClipping() {
    global ui, tab, CPDesktop, CPCanvasScrollMaxY, CPCanvasFixedYHwnds
    global DesktopTransitionViolations, DesktopTransitionMoves, DesktopTransitionProbeActive
    DesktopTransitionViolations := 0, DesktopTransitionMoves := 0, DesktopTransitionProbeActive := false
    callback := CallbackCreate(DesktopTransitionProbe, , 6), handles := []
    for ctrl in CPCanvasDirectControls() {
        if !CPCanvasFixedYHwnds.Has(ctrl.Hwnd) {
            handles.Push(ctrl.Hwnd)
            DllCall("comctl32\SetWindowSubclass", "ptr", ctrl.Hwnd, "ptr", callback, "uptr", 15, "uptr", 0)
        }
    }
    try {
        for page in [6, 8, 9, 3] {
            CPDesktopNavigate(page), CPDesktopLayout(ui, 0, 900, 640)
            for transparent in [false, true] {
                WinSetTransparent(transparent ? 230 : "Off", ui.Hwnd)
                DesktopTransitionProbeActive := true
                for target in [48, 96, 17, CPCanvasScrollMaxY, 72, 0]
                    CPCanvasScrollTo(0, target)
                DesktopTransitionProbeActive := false
            }
        }
        DesktopAssert(DesktopTransitionMoves > 200, "Transition probe observes real child-window move callbacks")
        DesktopAssert(DesktopTransitionViolations = 0, "Child regions stay inside the viewport even during WM_WINDOWPOSCHANGED (violations: " DesktopTransitionViolations ")")
        ; Positive control: moving an unclipped control across the boundary must
        ; be detected. End-frame screenshots alone cannot detect this brief gap.
        sample := CPDesktop["overlayPages"][3]["appearance"]
        DllCall("user32\SetWindowRgn", "ptr", sample.Hwnd, "ptr", 0, "int", 0)
        DesktopTransitionProbeActive := true
        sample.GetPos(&x)
        sample.Move(x, CPDesktop["headerH"] + 80)
        DesktopTransitionProbeActive := false
        DesktopAssert(DesktopTransitionViolations > 0, "Transition observer detects a deliberately unprotected move into the header")
    } finally {
        DesktopTransitionProbeActive := false
        for hwnd in handles
            DllCall("comctl32\RemoveWindowSubclass", "ptr", hwnd, "ptr", callback, "uptr", 15)
        CallbackFree(callback)
        WinSetTransparent("Off", ui.Hwnd)
        CPDesktopNavigate(1)
    }
}

DesktopTransitionProbe(hwnd, msg, wParam, lParam, subclassId, refData) {
    global ui, CPDesktop, DesktopTransitionViolations, DesktopTransitionMoves, DesktopTransitionProbeActive
    if DesktopTransitionProbeActive && msg = 0x47 && DllCall("user32\IsWindowVisible", "ptr", hwnd) {
        DesktopTransitionMoves += 1
        rect := Buffer(16), bounds := Buffer(16), region := DllCall("gdi32\CreateRectRgn", "int", 0, "int", 0, "int", 0, "int", 0, "ptr")
        kind := DllCall("user32\GetWindowRgn", "ptr", hwnd, "ptr", region)
        if kind != 1 { ; NULLREGION cannot paint anywhere.
            if kind = 0
                DllCall("user32\GetWindowRect", "ptr", hwnd, "ptr", bounds)
            else {
                DllCall("gdi32\GetRgnBox", "ptr", region, "ptr", bounds)
                ; Window regions include the non-client border; MapWindowPoints
                ; on hwnd would incorrectly add its Edit/Combo client inset.
                DllCall("user32\GetWindowRect", "ptr", hwnd, "ptr", rect)
                DllCall("user32\OffsetRect", "ptr", bounds, "int", NumGet(rect, 0, "int"), "int", NumGet(rect, 4, "int"))
            }
            DllCall("user32\MapWindowPoints", "ptr", 0, "ptr", ui.Hwnd, "ptr", bounds, "uint", 2)
            scale := GetWindowDPI(ui.Hwnd) / 96
            if NumGet(bounds, 0, "int") < Round(CPDesktop["side"] * scale)
                || NumGet(bounds, 4, "int") < Round((CPDesktop["headerH"] + 88) * scale)
                || NumGet(bounds, 12, "int") > Round((CPDesktop["height"] - 58) * scale)
                DesktopTransitionViolations += 1
        }
        DllCall("gdi32\DeleteObject", "ptr", region)
    }
    return DllCall("comctl32\DefSubclassProc", "ptr", hwnd, "uint", msg, "ptr", wParam, "ptr", lParam, "ptr")
}

DesktopTestScrollPainting() {
    global ui, tab, CPDesktop, DesktopScrollPaints, DesktopScrollRedrawSuspensions
    global CPCanvasScrollMaxY, ddlPrompt, ddlAudioTarget
    DesktopScrollPaints := Map(), DesktopScrollRedrawSuspensions := 0
    probe := CallbackCreate(DesktopScrollPaintProbe, , 6)
    handles := [ui.Hwnd]
    for key, ctrl in CPDesktop["chrome"] {
        if DesktopShown(ctrl) && !InStr(key, "Panel") {
            handles.Push(ctrl.Hwnd)
            DesktopScrollPaints[ctrl.Hwnd] := 0
        }
    }
    for handle in handles {
        if !DllCall("comctl32\SetWindowSubclass", "ptr", handle, "ptr", probe, "uptr", 4, "uptr", 0)
            throw Error("Could not install scroll paint probe")
    }
    try {
        for page in [1, 2, 3, 4, 5, 6, 7, 8, 9] {
            tab.Value := page
            CPDesktop["advanced"] := true, CPDesktop["audioHelpOpen"] := true
            CPDesktop["explanationStartupOpen"] := true
            ui.Show("NA x-9000 y-9000 w900 h640")
            CPDesktopLayout(ui, 0, 900, 640)
            DesktopAssert(CPCanvasScrollMaxY > 96, "Scroll regression has enough content on page " page)
            for layered in [false, true] {
                WinSetTransparent(layered ? 230 : "Off", ui.Hwnd)
                CPCanvasScrollTo(0, 0)
                ui.GetPos(,, &scrollCaptureW, &scrollCaptureH)
                scrollScale := GetWindowDPI(ui.Hwnd) / 96
                scrollName := "desktop-scroll-page" page "-" (layered ? "translucent" : "opaque")
                TestCapture(scrollName "-before.png", Round(scrollCaptureW * scrollScale), Round(scrollCaptureH * scrollScale))
                FileAppend(scrollName "|" Round(CPDesktop["side"] * scrollScale) "|" Round((CPDesktop["headerH"] + 88) * scrollScale) "|" Round((CPDesktop["height"] - 58) * scrollScale) "`n", A_ScriptDir "\scroll-bounds.txt")
                for handle in handles
                    DllCall("user32\UpdateWindow", "ptr", handle)
                for handle in DesktopScrollPaints
                    DesktopScrollPaints[handle] := 0
                DesktopScrollRedrawSuspensions := 0
                headerBefore := DesktopRect(CPDesktop["chrome"]["brand"])
                scrollFooterBefore := DesktopRect(CPDesktop["chrome"]["options"])
                promptBefore := ddlPrompt.Text, languageBefore := ddlAudioTarget.Text
                ; Both immediate wheel steps and coalesced thumb-drag steps.
                for target in [48, 96, 160, 80, 0] {
                    CPCanvasScrollTo(0, target)
                    for handle in handles
                        DllCall("user32\UpdateWindow", "ptr", handle)
                }
                for target in [120, 180, 48, 0] {
                    CPCanvasQueueScrollTo(0, target)
                    CPCanvasFlushQueuedScroll()
                    for handle in handles
                        DllCall("user32\UpdateWindow", "ptr", handle)
                }
                DesktopAssert(DesktopScrollRedrawSuspensions = 0, "Scrolling never suspends/hides the whole desktop surface")
                fixedPaints := 0
                fixedDetails := ""
                for handle, count in DesktopScrollPaints
                    fixedPaints += count
                for key, ctrl in CPDesktop["chrome"] {
                    if DesktopScrollPaints.Has(ctrl.Hwnd) && DesktopScrollPaints[ctrl.Hwnd]
                        fixedDetails .= " " key "=" DesktopScrollPaints[ctrl.Hwnd] "@" DesktopRect(ctrl)
                }
                DesktopAssert(fixedPaints = 0, "Scrolling does not repaint fixed chrome (observed " fixedPaints ":" fixedDetails ")")
                DesktopAssert(DesktopRect(CPDesktop["chrome"]["brand"]) = headerBefore && DesktopRect(CPDesktop["chrome"]["options"]) = scrollFooterBefore, "Scroll preserves sticky header and footer positions")
                DesktopAssert(ddlPrompt.Text = promptBefore && ddlAudioTarget.Text = languageBefore, "Scroll preserves native selections")
                exStyle := DllCall("user32\GetWindowLongPtr", "ptr", ui.Hwnd, "int", -20, "ptr")
                DesktopAssert((exStyle & 0x02000000) && !!(exStyle & 0x80000) = layered, "Scrolling preserves buffering and opacity styles")
                CPCanvasScrollTo(0, CPCanvasScrollMaxY)
                TestCapture(scrollName ".png", Round(scrollCaptureW * scrollScale), Round(scrollCaptureH * scrollScale))
                CPCanvasScrollTo(0, 0)
                ; Positive control: the observers must still detect a repaint.
                CPDesktop["chrome"]["brand"].Redraw()
                DllCall("user32\UpdateWindow", "ptr", CPDesktop["chrome"]["brand"].Hwnd)
                DesktopAssert(DesktopScrollPaints[CPDesktop["chrome"]["brand"].Hwnd] > 0, "Scroll chrome-paint observer is active")
            }
        }
    } finally {
        CPCanvasCancelQueuedScroll()
        for handle in handles
            DllCall("comctl32\RemoveWindowSubclass", "ptr", handle, "ptr", probe, "uptr", 4)
        CallbackFree(probe)
        WinSetTransparent("Off", ui.Hwnd)
        CPDesktop["advanced"] := false, CPDesktop["audioHelpOpen"] := false
        CPDesktop["explanationStartupOpen"] := false
        tab.Value := 1
        ui.Show("NA x-9000 y-9000 w1400 h820")
        CPDesktopLayout(ui, 0, 1400, 820)
    }
}
DesktopScrollPaintProbe(hwnd, msg, wParam, lParam, subclassId, refData) {
    global ui, DesktopScrollPaints, DesktopScrollRedrawSuspensions
    if hwnd = ui.Hwnd && msg = 0xB && !wParam
        DesktopScrollRedrawSuspensions += 1
    if msg = 0xF && DesktopScrollPaints.Has(hwnd) && DllCall("user32\GetUpdateRect", "ptr", hwnd, "ptr", 0, "int", 0) {
        DesktopScrollPaints[hwnd] += 1
    }
    return DllCall("comctl32\DefSubclassProc", "ptr", hwnd, "uint", msg, "ptr", wParam, "ptr", lParam, "ptr")
}
DesktopTestSidebarRefresh() {
    global CPDesktop, DesktopSidebarPaints
    DesktopSidebarPaints := Map()
    callback := CallbackCreate(DesktopSidebarPaintProbe, , 6)
    try {
        for key in ["screenshot", "audioPage", "explanation", "overlays", "study", "profiles", "settings"] {
            handle := CPDesktop["chrome"][key].Hwnd
            if !DllCall("comctl32\SetWindowSubclass", "ptr", handle, "ptr", callback, "uptr", 3, "uptr", 0)
                throw Error("Could not install sidebar paint probe")
            DesktopSidebarPaints[handle] := 0
        }
        ; Flush initial layout paints before reproducing the production 80-ms
        ; active-tab timer. This must not erase/repaint unchanged sidebar rows.
        for handle in DesktopSidebarPaints
            DllCall("user32\UpdateWindow", "ptr", handle)
        for handle in DesktopSidebarPaints
            DesktopSidebarPaints[handle] := 0
        Loop 12 {
            UpdateCPActiveTabHighlight()
            Sleep(80)
            for handle in DesktopSidebarPaints
                DllCall("user32\UpdateWindow", "ptr", handle)
        }
        paintCount := 0
        for handle, count in DesktopSidebarPaints
            paintCount += count
        DesktopAssert(paintCount = 0, "Idle sidebar timer produces no paints (observed " paintCount ")")
        ; The probe must still detect a genuine redraw, so a hidden window or
        ; a broken observer cannot make the no-flicker assertion pass silently.
        CPRenderCustomTabBar(true)
        for handle in DesktopSidebarPaints
            DllCall("user32\UpdateWindow", "ptr", handle)
        paintCount := 0
        for handle, count in DesktopSidebarPaints
            paintCount += count
        DesktopAssert(paintCount > 0, "Explicit sidebar refresh still repaints native controls")
        ; Simulate the old/new selected row state independently of the native
        ; Tab's page-visibility redraws, isolating this selection-update path.
        selectedHandle := CPDesktop["chrome"]["screenshot"].Hwnd
        previousHandle := CPDesktop["chrome"]["audioPage"].Hwnd
        CPDesktop["paint"][selectedHandle]["selected"] := false
        CPDesktop["paint"][previousHandle]["selected"] := true
        for handle in DesktopSidebarPaints
            DesktopSidebarPaints[handle] := 0
        CPDesktopUpdateNavigation()
        for handle in DesktopSidebarPaints
            DllCall("user32\UpdateWindow", "ptr", handle)
        for handle, count in DesktopSidebarPaints
            DesktopAssert(count = (handle = selectedHandle || handle = previousHandle ? 1 : 0), "Selection change repaints only changed rows")
        DesktopAssert(CPDesktop["paint"][selectedHandle]["selected"] && !CPDesktop["paint"][previousHandle]["selected"], "Selection cache follows the current page")
    } finally {
        for handle in DesktopSidebarPaints
            DllCall("comctl32\RemoveWindowSubclass", "ptr", handle, "ptr", callback, "uptr", 3)
        CallbackFree(callback)
    }
}
DesktopSidebarPaintProbe(hwnd, msg, wParam, lParam, subclassId, refData) {
    global DesktopSidebarPaints
    if msg = 0xF && DesktopSidebarPaints.Has(hwnd)
        DesktopSidebarPaints[hwnd] += 1
    return DllCall("comctl32\DefSubclassProc", "ptr", hwnd, "uint", msg, "ptr", wParam, "ptr", lParam, "ptr")
}
DesktopCaptureLogoHeader(scale) {
    global ui, CPDesktop
    width := Round(320 * scale), height := Round(CPDesktop["headerH"] * scale)
    screenDC := DllCall("user32\GetDC", "ptr", ui.Hwnd, "ptr")
    dc := DllCall("gdi32\CreateCompatibleDC", "ptr", screenDC, "ptr")
    bitmap := DllCall("gdi32\CreateCompatibleBitmap", "ptr", screenDC, "int", width, "int", height, "ptr")
    previous := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", bitmap, "ptr")
    colors := CPDesktopPalette(), rect := Buffer(16, 0), image := 0
    NumPut("int", width, "int", height, rect, 8)
    DllCall("user32\FillRect", "ptr", dc, "ptr", rect, "ptr", CPDesktopBrush(colors["bar"]))
    DesktopAssert(CPDesktopPaintLogo(dc, scale), "Production logo renderer succeeds at " Round(scale * 100) "%")
    DesktopAssert(DllCall("gdi32\GetPixel", "ptr", dc, "int", 0, "int", 0, "uint") = CPColorRef(colors["bar"]), "Logo preserves transparent background at " Round(scale * 100) "%")
    font := DllCall("gdi32\CreateFontW", "int", -Round(DesktopBrandPointSize() * 96 / 72 * scale), "int", 0,
        "int", 0, "int", 0, "int", 700, "uint", 0, "uint", 0, "uint", 0,
        "uint", 1, "uint", 0, "uint", 0, "uint", 5, "uint", 0, "wstr", "Segoe UI", "ptr")
    oldFont := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", font, "ptr")
    CPDesktop["chrome"]["brand"].GetPos(&brandX)
    NumPut("int", Round(brandX * scale), rect, 0)
    DllCall("gdi32\SetBkMode", "ptr", dc, "int", 1)
    DllCall("gdi32\SetTextColor", "ptr", dc, "uint", CPColorRef(colors["text"]))
    DllCall("user32\DrawTextW", "ptr", dc, "wstr", "JRPG Translator", "int", -1, "ptr", rect, "uint", 0x8824)
    DllCall("gdi32\SelectObject", "ptr", dc, "ptr", oldFont)
    DllCall("gdi32\DeleteObject", "ptr", font)
    encoder := Buffer(16, 0)
    DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "ptr", bitmap, "ptr", 0, "ptr*", &image)
    DllCall("ole32\CLSIDFromString", "wstr", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "ptr", encoder)
    DllCall("gdiplus\GdipSaveImageToFile", "ptr", image, "wstr", A_ScriptDir "\desktop-logo-" Round(scale * 100) ".png", "ptr", encoder, "ptr", 0)
    DllCall("gdiplus\GdipDisposeImage", "ptr", image)
    DllCall("gdi32\SelectObject", "ptr", dc, "ptr", previous)
    DllCall("gdi32\DeleteObject", "ptr", bitmap)
    DllCall("gdi32\DeleteDC", "ptr", dc)
    DllCall("user32\ReleaseDC", "ptr", ui.Hwnd, "ptr", screenDC)
}
DesktopBrandPointSize() {
    global CPDesktop
    handle := CPDesktop["chrome"]["brand"].Hwnd
    fontInfo := Buffer(92, 0) ; LOGFONTW: read the actual production title font.
    DllCall("gdi32\GetObjectW", "ptr", SendMessage(0x31, 0, 0, handle), "int", 92, "ptr", fontInfo)
    return Round(Abs(NumGet(fontInfo, 0, "int")) * 72 / GetWindowDPI(handle))
}
DesktopBrandTextWidth() {
    global CPDesktop
    return DesktopTextWidth(CPDesktop["chrome"]["brand"])
}
DesktopTextWidth(ctrl) {
    dc := DllCall("user32\GetDC", "ptr", ctrl.Hwnd, "ptr")
    oldFont := DllCall("gdi32\SelectObject", "ptr", dc, "ptr", SendMessage(0x31, 0, 0, ctrl.Hwnd), "ptr")
    size := Buffer(8, 0)
    DllCall("gdi32\GetTextExtentPoint32W", "ptr", dc, "wstr", ctrl.Text, "int", StrLen(ctrl.Text), "ptr", size)
    DllCall("gdi32\SelectObject", "ptr", dc, "ptr", oldFont)
    DllCall("user32\ReleaseDC", "ptr", ctrl.Hwnd, "ptr", dc)
    return NumGet(size, 0, "int") * 96 / GetWindowDPI(ctrl.Hwnd)
}
TestDesktopStudyLibrary() {
    global controlDarkMode, CPStudyLibraryState, ui, iniPath
    global studyLibraryDir := A_ScriptDir "\study-fixture"
    slBigBoxPresentation := false, slStandalone := true
    slActiveLibrary := "Test library", slOutputDir := A_ScriptDir
    ; @STUDY_LIBRARY_CONTROLS@
    return slState
}

TestDesktopStudyReader() {
    global controlDarkMode, CPStudyReaderState, studyLibraryDir, iniPath
    srWantBigBox := false, srBigBoxLibrary := 0, srGroupId := 1
    srGroups := [], srOutputDir := A_ScriptDir
    ; @STUDY_READER_CONTROLS@
    return srState
}

TestDesktopStudyWindows() {
    global controlDarkMode, CPStudyLibraryState, CPStudyReaderState
    global CPComboSeparatorBefore
    controlDarkMode := 1
    TestDesktopStudyDialogs()
    DesktopTestStudyIdle()
    DesktopTestStudyReaderOwnership()
    for kind in ["library", "reader"] {
        s := kind = "library" ? TestDesktopStudyLibrary() : TestDesktopStudyReader()
        dlg := s["gui"], hwnd := dlg.Hwnd
        s["suspend"] := true ; No bridge requests from native selection events.
        DesktopAssert(s.Has("desktop"), kind " opts into desktop shell")
        DesktopAssert(IsObject(StudyDesktopContext(hwnd)), kind " has independent paint state")
        DesktopAssert(!(WinGetStyle(hwnd) & 0xC00000), kind " integrated header replaces native caption")
        DesktopAssert((WinGetStyle(hwnd) & 0x40000), kind " retains native resizing")
        s["source"].Value := "お城には行けました？`r`nポルダ村なんかに負けるか！"
        s["metadata"].Value := "Profile: Kabuki Den  ·  Chapter: 2`r`nAnki: Not checked`r`nSpeaker: Kate of Sanju`r`nTags: story, dialogue`r`nModel: Gemini / example-model-name`r`nPrompt: default_en"
        s["versionView"].Value := "v02  ·  2026-09-08  (latest)"
        s["detailTitle"].Text := "Saved explanation"
        s["openImage"].Enabled := true
        if kind = "library" {
            s["libraryDdl"].Add(["Test library", "Another library", "Manage Study Libraries…"])
            s["libraryDdl"].Choose(1)
            CPComboSeparatorBefore[s["libraryDdl"].Hwnd] := 2
            s["desktop"]["combos"][s["libraryDdl"].Hwnd]["separatorBefore"] := 2
            s["status"].Text := "2 sources · 3 explanations · 1 not linked to Anki"
            s["storageButton"].Text := "Storage: 1.4 MB"
            s["currentChapterStatus"].Text := "New entries: Chapter 2"
            s["list"].Add("", "2026-09-08", "Kabuki Den", "2", "Kate of Sanju", "story", "お城には行けました？", "可能形", "2", "Not checked", "1")
            s["list"].Add("", "2026-09-07", "Kabuki Den", "2", "", "story", "ポルダ村なんかに負けるか！", "なんか", "1", "Found in Anki", "2")
            for i, colWidth in [122, 110, 0, 0, 0, 220, 116, 64, 116]
                s["list"].ModifyCol(i, colWidth)
            s["list"].ModifyCol(s["columns"].Length + 1, 0)
            for key in ["studyButton", "editDetailsButton", "storageButton"]
                s[key].Enabled := true
        } else {
            s["entryStatus"].Text := "1 of 2", s["sectionStatus"].Text := "Full explanation"
            s["ankiCheck"].Text := "Added to Anki (manual)  (Alt+A)"
            s["explanation"].Value := "Original Japanese`r`n`r`nお城（おしろ）には行けました？`r`n`r`nNatural English translation`r`n`r`nWere you able to go to the castle?`r`n`r`nDetailed analysis`r`n`r`nお城 is the polite form of castle. The particle に marks the destination.`r`n`r`nKey vocabulary`r`n`r`nお城（おしろ） — castle.`r`n行く（いく） — to go."
            for key in ["fullSectionButton", "newVersionButton", "addAnkiButton", "copyButton", "editExplanationButton", "ankiCheck"]
                s[key].Enabled := true
        }
        for bounds in [[960, 700], [1200, 800], [1500, 960]] {
            w := bounds[1], h := bounds[2]
            dlg.Show("Hide w" w " h" h)
            if kind = "library"
                StudyLibraryResize(s, dlg, 0, w, h)
            else
                StudyReaderResize(s, dlg, 0, w, h)
            CPApplyOwnedDialogTheme(dlg)
            ; Visible but off-screen: render only this synthetic window.
            dlg.Show("NA x-12000 y-12000 w" w " h" h)
            Sleep(50)
            for key, item in s["desktop"]["paint"] {
                ctrl := item["ctrl"]
                if !DesktopShown(ctrl)
                    continue
                ctrl.GetPos(&x, &y, &cw, &ch)
                DesktopAssert(x >= 0 && y >= 0 && x + cw <= w + 1 && y + ch <= h + 1, kind " button stays inside " w ": " ctrl.Text)
                DesktopAssert((WinGetStyle(ctrl.Hwnd) & 0xF) = 0xB, kind " keeps modern button after re-theme")
                if item["kind"] != "caption"
                    DesktopAssert(DesktopTextWidth(ctrl) < cw - 15, kind " label fits: " ctrl.Text)
            }
            s["source"].GetPos(&sx, &sy, &sourceWidth, &sh)
            DesktopAssert(sy + sh < h - 62, kind " source stays above footer")
            main := s[kind = "reader" ? "explanation" : "list"]
            main.GetPos(&mx, &my, &mw, &mh)
            DesktopAssert(mh >= 240 && mw >= 490, kind " has a useful reading/table area at " w)
            DesktopAssert(!(WinGetExStyle(s["source"].Hwnd) & 0x200), kind " read-only source has no desktop frame")
            if kind = "library"
                DesktopAssert(CPDesktopIsCombo(s["libraryDdl"].Hwnd), "Library picker uses modern dropdown renderer")
            if kind = "library" {
                DesktopAssert(!s.Has("newLibraryButton"), "Library manager has no standalone toolbar button")
                DesktopAssert(SendMessage(0x146, 0, 0, s["libraryDdl"].Hwnd) = 3,
                    "Library picker includes one integrated management action")
                s["libraryDdl"].GetPos(, , &libraryPickerW)
                DesktopAssert(libraryPickerW >= 300, "Library picker uses the space freed by the removed button")
                DesktopAssert(CPComboSeparatorBefore.Get(s["libraryDdl"].Hwnd, -1) = 2,
                    "Library management action is visually separated from libraries")
            }
            if kind = "library"
                DesktopAssert(!(SendMessage(0x1037, 0, 0, s["list"].Hwnd) & 1), "Modern table removes native grid lines")
            DesktopAssert(!(WinGetExStyle(hwnd) & 0x02000000), kind " avoids whole-window compositing during native movement")
            DesktopAssert(WinGetStyle(hwnd) & 0x02000000, kind " clips background painting around child controls")
            if kind = "library"
                DesktopAssert(SendMessage(0x1037, 0, 0, s["list"].Hwnd) & 0x10000, "Table buffers its own painting")
            dpi := GetWindowDPI(hwnd) / 96
            TestDesktopStudyCapture(dlg, "desktop-study-" kind "-" w ".png", Round(w * dpi), Round(h * dpi))
            dlg.Hide()
        }
        ; Exercise media sizing with an actual bundled bitmap, not personal images.
        s["media"] := [Map("path", A_ScriptDir "\assets\bigbox-logo.png", "width", 0, "height", 0)]
        s["mediaIndex"] := 1
        StudyLibraryShowImage(s)
        DesktopAssert(DesktopShown(s["picture"]) && s["openImage"].Enabled, kind " retains screenshot rendering/open action")
        s["picture"].GetPos(&px, &py, &pw, &ph)
        area := s["imageArea"]
        DesktopAssert(px >= area["x"] && py >= area["y"] && pw <= area["w"] && ph <= area["h"], kind " image remains inside context panel")
        DesktopAssert(s["imageInfo"].Text = "1 / 1", kind " compact screenshot counter fits")
        for mediaBounds in [[1200, 800], [960, 700]] {
            dlg.Show("NA x-12000 y-12000 w" mediaBounds[1] " h" mediaBounds[2])
            StudyDesktopResize(s, mediaBounds[1], mediaBounds[2])
            Sleep(80)
            DesktopAssert(!s["imageLayoutQueued"], kind " releases queued image flag after each resize")
            s["picture"].GetPos(&px, &py, &pw, &ph)
            area := s["imageArea"]
            DesktopAssert(px >= area["x"] && py >= area["y"] && px + pw <= area["x"] + area["w"] && py + ph <= area["y"] + area["h"], kind " repeated visible resizes refit the screenshot")
            DesktopAssertStudyPicture(s, kind " after resize")
            dpi := GetWindowDPI(hwnd) / 96
            TestDesktopStudyCapture(dlg, "desktop-study-" kind "-image-" mediaBounds[1] ".png", Round(mediaBounds[1] * dpi), Round(mediaBounds[2] * dpi))
        }
        for iteration in [1, 2] {
            for shape in ["landscape", "wide", "portrait"] {
                ; Replace media exactly as selecting another explanation does.
                s["media"] := [Map("path", A_ScriptDir "\study-preview-" shape ".bmp", "width", 0, "height", 0)]
                s["mediaIndex"] := 1
                StudyLibraryShowImage(s)
                StudyLibraryRedraw(s)
                Sleep(30)
                DesktopAssertStudyPicture(s, kind " switching to " shape)
                DesktopAssert(s["imageInfo"].Text = "1 / 1" && s["openImage"].Enabled, kind " switched image keeps its counter/open action")
                TestDesktopStudyCapture(dlg, "desktop-study-" kind "-" shape ".png", Round(960 * dpi), Round(700 * dpi))
                if iteration = 1 {
                    s["media"] := [], s["mediaIndex"] := 0
                    StudyLibraryShowImage(s)
                    DesktopAssert(!DesktopShown(s["picture"]) && !s["openImage"].Enabled, kind " entry without media clears preview")
                }
            }
        }
        s["media"] := [Map("path", A_ScriptDir "\study-preview-landscape.bmp", "width", 0, "height", 0),
            Map("path", A_ScriptDir "\study-preview-portrait.bmp", "width", 0, "height", 0)]
        s["mediaIndex"] := 1
        StudyLibraryShowImage(s)
        StudyLibraryNextImage(s)
        DesktopAssertStudyPicture(s, kind " next screenshot")
        DesktopAssert(s["imageInfo"].Text = "2 / 2", kind " next screenshot updates counter")
        StudyLibraryPreviousImage(s)
        DesktopAssertStudyPicture(s, kind " previous screenshot")
        dlg.Hide()
        if kind = "reader" {
            originalExplanation := s["explanation"].Value
            StudyReaderSetEditing(s, true)
            DesktopAssert(DesktopShown(s["saveEditButton"]) && !DesktopShown(s["addAnkiButton"]), "Reader editing keeps Save/Cancel workflow")
            DesktopAssert(s["explanation"].Value = originalExplanation, "Presentation preserves full Japanese/explanation text")
        }
        SetTimer(s["imageLayoutCallback"], 0)
        SetTimer(s["redrawCallback"], 0)
        if kind = "library"
            OnMessage(0x4E, s["headerNotifyCallback"], 0)
        dlg.Destroy()
        DesktopAssert(!IsObject(StudyDesktopContext(hwnd)), kind " cleans up desktop registry")
    }
    CPStudyLibraryState := 0, CPStudyReaderState := 0
}

DesktopAssertStudyPicture(s, label) {
    hwnd := s["gui"].Hwnd, scale := GetWindowDPI(hwnd) / 96
    s["picture"].GetPos(&x, &y, &w, &h)
    area := s["imageArea"], item := s["media"][s["mediaIndex"]]
    DesktopAssert(DesktopShown(s["picture"]), label " is visible")
    DesktopAssert(x >= area["x"] && y >= area["y"] && x + w <= area["x"] + area["w"] && y + h <= area["y"] + area["h"], label " fits the preview")
    fit := Min(area["w"] / item["width"], area["h"] / item["height"])
    DesktopAssert(Abs(w - item["width"] * fit) <= 1 && Abs(h - item["height"] * fit) <= 1, label " preserves the whole image and aspect ratio")
    for fraction in [[.1, .1], [.9, .1], [.5, .5], [.1, .9], [.9, .9]] {
        point := Buffer(8)
        NumPut("int", Round((x + w * fraction[1]) * scale), "int", Round((y + h * fraction[2]) * scale), point)
        imageHit := DllCall("user32\ChildWindowFromPointEx", "ptr", hwnd, "int64", NumGet(point, 0, "int64"), "uint", 1, "ptr")
        DesktopAssert(imageHit = s["picture"].Hwnd, label " is not covered by a sibling control")
    }
}

DesktopTestStudyIdle() {
    global CPStudyLibraryState, CPStudyReaderState, DesktopStudyPaints, DesktopStudyHeartbeats
    s := TestDesktopStudyLibrary(), dlg := s["gui"]
    s["suspend"] := true
    s["list"].Add("", "2026-09-08", "Test profile", "2", "Speaker", "story", "お城", "Grammar", "1", "Not checked", "1")
    StudyLibraryApplyColumns(s)
    dlg.Show("Hide w1200 h800")
    StudyLibraryResize(s, dlg, 0, 1200, 800)
    CPApplyOwnedDialogTheme(dlg)
    StudyLibraryApplyHeaderIndicators(s)
    DesktopStudyPaints := Map(dlg.Hwnd, 0, s["list"].Hwnd, 0, s["headerHwnd"], 0)
    DesktopStudyHeartbeats := 0
    OnMessage(0xF, DesktopStudyPaintProbe)
    SetTimer(DesktopStudyHeartbeat, 20)
    try {
        dlg.Show("NA x-12000 y-12000")
        Sleep(300)
        for hwnd in DesktopStudyPaints
            DesktopStudyPaints[hwnd] := 0
        started := A_TickCount
        Sleep(500)
        FileAppend("Study idle: " (A_TickCount - started) "ms, parent/list/header paints "
            DesktopStudyPaints[dlg.Hwnd] "/" DesktopStudyPaints[s["list"].Hwnd] "/" DesktopStudyPaints[s["headerHwnd"]]
            ", heartbeats " DesktopStudyHeartbeats "`n", "*")
        DesktopAssert(DesktopStudyHeartbeats >= 10, "Library leaves the message loop responsive")
        DesktopAssert(DesktopStudyPaints[dlg.Hwnd] < 5 && DesktopStudyPaints[s["list"].Hwnd] < 5,
            "Idle Library does not continuously repaint the table")
        ; Windows must exclude child rectangles from the parent paint DC:
        ; otherwise our opaque card fills can erase the table between frames.
        dc := DllCall("user32\GetDC", "ptr", dlg.Hwnd, "ptr")
        try {
            s["list"].GetPos(&x, &y, &w, &h)
            scale := GetWindowDPI(dlg.Hwnd) / 96
            DesktopAssert(!DllCall("gdi32\PtVisible", "ptr", dc,
                "int", Round((x + w / 2) * scale), "int", Round((y + h / 2) * scale)),
                "Parent background cannot paint over the native table")
        } finally DllCall("user32\ReleaseDC", "ptr", dlg.Hwnd, "ptr", dc)
        for iteration in [1, 2, 3, 4, 5] {
            DllCall("user32\InvalidateRect", "ptr", s["headerHwnd"], "ptr", 0, "int", true)
            DllCall("user32\UpdateWindow", "ptr", s["headerHwnd"])
            DesktopAssert(!DllCall("user32\GetUpdateRect", "ptr", s["headerHwnd"], "ptr", 0, "int", false),
                "Header completes its paint transaction")
        }
    } finally {
        SetTimer(DesktopStudyHeartbeat, 0)
        OnMessage(0xF, DesktopStudyPaintProbe, 0)
        OnMessage(0x4E, s["headerNotifyCallback"], 0)
        SetTimer(s["imageLayoutCallback"], 0)
        SetTimer(s["redrawCallback"], 0)
        SetTimer(s["listRedrawCallback"], 0)
        dlg.Destroy()
        CPStudyLibraryState := 0, CPStudyReaderState := 0
    }
}

DesktopTestStudyReaderOwnership() {
    global CPStudyLibraryState, CPStudyReaderState
    for topmost in [false, true] {
        library := TestDesktopStudyLibrary(), lg := library["gui"]
        library["suspend"] := true
        lg.Opt(topmost ? "+AlwaysOnTop" : "-AlwaysOnTop")
        lg.Show("NA x-12000 y-12000 w1200 h800")
        reader := TestDesktopStudyReader(), rg := reader["gui"], rh := rg.Hwnd
        rg.Show("NA x-12000 y-12000 w1200 h800")
        ; Production activates Reader after showing it. Exercise the same
        ; z-order promotion without taking focus from the user's desktop.
        DllCall("user32\SetWindowPos", "ptr", rh, "ptr", 0,
            "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x13)
        DesktopAssert(DllCall("user32\GetWindow", "ptr", rh, "uint", 4, "ptr") = lg.Hwnd,
            "Reader is owned by its Library, including topmost Library")
        DesktopAssert(DllCall("user32\IsWindowVisible", "ptr", rh)
            && DllCall("user32\IsWindowVisible", "ptr", lg.Hwnd), "Reader opens while Library remains open")
        DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", lg.Hwnd), "Reader does not disable its Library")
        ; Reopening an already attached Reader must not replace its original
        ; owner/topmost snapshot with the temporary Library relationship.
        previousTopmost := reader["desktopPreviousTopmost"]
        StudyReaderSyncDesktopOwner(reader)
        DesktopAssert(reader["desktopPreviousTopmost"] = previousTopmost,
            "Reusing Reader preserves original topmost preference")
        z := DllCall("user32\GetWindow", "ptr", lg.Hwnd, "uint", 3, "ptr"), found := false
        while z {
            if z = rh {
                found := true
                break
            }
            z := DllCall("user32\GetWindow", "ptr", z, "uint", 3, "ptr")
        }
        DesktopAssert(found, "Reader is above Library in native window order (topmost Library: " topmost ")")
        reader["explanation"].Value := "Unsaved test edit: お城"
        reader["editing"] := true
        StudyLibraryClose(library)
        DesktopAssert(DllCall("user32\IsWindow", "ptr", rh), "Closing Library preserves open Reader")
        DesktopAssert(!DllCall("user32\GetWindow", "ptr", rh, "uint", 4, "ptr"), "Reader detaches from closed Library")
        DesktopAssert(!!(WinGetExStyle(rh) & 8) = previousTopmost, "Reader restores original topmost state")
        DesktopAssert(reader["explanation"].Value = "Unsaved test edit: お城" && reader["editing"],
            "Closing Library preserves Reader's unsaved edit")
        DesktopAssert(reader["desktop"]["chrome"]["back"].Text = "Close Reader", "Detached Reader offers Close Reader")
        SetTimer(reader["imageLayoutCallback"], 0)
        SetTimer(reader["redrawCallback"], 0)
        rg.Destroy()
        CPStudyLibraryState := 0, CPStudyReaderState := 0
    }
}

DesktopStudyPaintProbe(wParam, lParam, msg, hwnd) {
    global DesktopStudyPaints
    if DesktopStudyPaints.Has(hwnd)
        DesktopStudyPaints[hwnd] += 1
}

DesktopStudyHeartbeat() {
    global DesktopStudyHeartbeats
    DesktopStudyHeartbeats += 1
}

TestDesktopAnkiControls(reader, vocabulary := true, screenshot := true) {
    srState := reader
    srState["media"] := screenshot ? [Map("path", A_ScriptDir "\study-preview-landscape.bmp", "width", 720, "height", 540)] : []
    srState["mediaIndex"] := screenshot ? 1 : 0
    saBigBox := false, saIsVocabulary := vocabulary, saProfile := "Kabuki Den"
    saCardKind := vocabulary ? "vocabulary" : "explanation"
    saMapping := Map("model", "Basic", "japaneseField", "Front", "explanationField", "Back", "addDeck", "Japanese", "deck", "Japanese")
    saDecks := ["Japanese", "Vocabulary", "Sentence practice"], saCancelReturn := 0, saHandoffSource := 0
    saFront := vocabulary ? "お城" : "お城には行けました？"
    saBack := vocabulary ? "お城（おしろ） — castle. The polite prefix お is used with 城."
        : "Natural English translation`r`nWere you able to go to the castle?`r`n`r`nDetailed analysis`r`nお城 is the polite form of castle. The particle に marks the destination."
    saInferenceConfident := true
    ; @STUDY_ANKI_CONTROLS@
    return saAddState
}

TestDesktopCandidateControls(library) {
    global controlDarkMode, CPStudyCandidateState
    slState := library, scWantBigBox := false, scOutputDir := A_ScriptDir
    ; @STUDY_CANDIDATES_CONTROLS@
    return scState
}

TestDesktopStudyDialogs() {
    global CPStudyLibraryState, CPStudyReaderState, CPStudyCandidateState
    library := TestDesktopStudyLibrary(), reader := TestDesktopStudyReader()
    TestDesktopMetadataDialogs(library)
    TestDesktopManagementDialogs(library)
    TestDesktopRecommendationDialogs(library)
    TestDesktopSecondaryStudyDialogs(library, reader)
    TestSharedDialogs(library, reader)
    TestNamingDialogs(reader)
    TestModelDialogs(reader)
    TestAuthoringDialogs(reader)
    TestControlDialogs(reader)
    TestAppearanceDialogs(reader)
    TestHelpDialogs(reader)
    TestPickerPolish()
    for vocabulary in [true, false] {
        s := TestDesktopAnkiControls(reader, vocabulary, vocabulary), g := s["gui"], hwnd := g.Hwnd
        DesktopAssert(s.Has("desktop"), "Anki review uses modern desktop shell")
        DesktopAssert(CPDesktopIsCombo(s["deckDdl"].Hwnd), "Anki deck has a modern dropdown")
        DesktopAssert(!DesktopShown(s["desktopControls"]["imageFrame"]), "Anki preview frame cannot cover image")
        for size in [[960, 700], [1040, 760], [1400, 900]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(s, w, h)
            DesktopDialogAssertChrome(s, w, h)
            s["front"].GetPos(&fx, &fy, &fw, &fh)
            s["back"].GetPos(&bx, &by, &bw, &bh)
            DesktopAssert(fw > 500 && fh >= 80 && bh >= 220, "Anki editors remain usable at " w)
            DesktopAssert(!(WinGetExStyle(s["back"].Hwnd) & 0x200), "Anki editor has no classic white client edge")
            DesktopAssert(s["front"].Value = (vocabulary ? "お城" : "お城には行けました？"), "Anki layout preserves Japanese content")
            DesktopAssert(s["includeScreenshot"].Enabled = vocabulary, "Screenshot toggle reflects available media")
            if vocabulary {
                image := s["desktopControls"]["image"]
                image.GetPos(&ix, &iy, &iw, &ih)
                DesktopAssert(DesktopShown(image) && iw > 100 && ih > 100, "Anki screenshot is visible")
                DesktopAssert(Abs(iw / ih - 4 / 3) < 0.02, "Anki screenshot preserves aspect ratio")
                scale := GetWindowDPI(hwnd) / 96
                point := Buffer(8)
                NumPut("int", Round((ix + iw / 2) * scale), "int", Round((iy + ih / 2) * scale), point)
                hit := DllCall("user32\ChildWindowFromPointEx", "ptr", hwnd,
                    "int64", NumGet(point, 0, "int64"), "uint", 1, "ptr")
                DesktopAssert(hit = image.Hwnd, "Anki screenshot is not covered by another control")
                DesktopAssert(DesktopShown(s["exampleButton"]) && (WinGetStyle(s["exampleButton"].Hwnd) & 0x10000), "Generate example remains keyboard reachable")
            }
            dpi := GetWindowDPI(hwnd) / 96
            TestDesktopStudyCapture(g, "desktop-anki-" (vocabulary ? "vocabulary" : "explanation") "-" w ".png", Round(w * dpi), Round(h * dpi))
        }
        s["front"].Value := "手動編集", s["back"].Value := "Edited by the user."
        StudyDesktopDialogResize(s, g, 0, 1040, 760)
        DesktopAssert(s["front"].Value = "手動編集" && s["back"].Value = "Edited by the user.", "Anki resizing preserves manual edits")
        g.Destroy()
        DesktopAssert(!IsObject(StudyDesktopContext(hwnd)), "Anki review cleans paint registry")
    }
    s := TestDesktopCandidateControls(library), g := s["gui"], hwnd := g.Hwnd
    sample := Map("recommendation", 1, "score", 4, "reason", "A useful everyday expression with a clear example.")
    alternate := Map("recommendation", 0, "score", 2, "reason", "A deliberately different second-row assessment.")
    s["sentences"] := [sample, alternate], s["vocabulary"] := [sample, alternate]
    s["allSentences"] := s["sentences"], s["allVocabulary"] := s["vocabulary"]
    s["baseStatus"] := "1 sentence · 1 vocabulary entry", s["snapshotThrough"] := "2026-09-08"
    s["sentenceList"].Add("Select Focus", "Recommended", "2026-09-08", "Kabuki Den", "お城には行けました？", "1")
    s["sentenceList"].Add("", "Not recommended", "2026-09-07", "Kabuki Den", "二つ目の文", "1")
    s["vocabularyList"].Add("Select Focus", "Recommended", "お城（おしろ）", "castle; polite form", "2", "Kabuki Den", "2026-09-08")
    s["vocabularyList"].Add("", "Not recommended", "言葉（ことば）", "second-row word", "1", "Kabuki Den", "2026-09-07")
    DesktopAssert(s.Has("desktop") && s["tabs"].Type = "Tab2", "Candidate manager uses desktop tab container")
    for size in [[960, 620], [1040, 640], [960, 760], [1040, 800], [1400, 900]] {
        w := size[1], h := size[2]
        DesktopDialogFixtureShow(s, w, h)
        DesktopDialogAssertChrome(s, w, h)
        DesktopAssert(CPDesktopIsCombo(s["scopeDdl"].Hwnd) && CPDesktopIsCombo(s["aiFilterDdl"].Hwnd), "Candidate filters use modern dropdowns")
        for index in [2, 1, 2] {
            StudyDesktopCandidatesChoose(s, index)
            DesktopAssert(DesktopShown(s[index = 2 ? "vocabularyList" : "sentenceList"]), "Selected candidate table is visible")
            DesktopAssert(!DesktopShown(s[index = 2 ? "sentenceList" : "vocabularyList"]), "Inactive table is hidden")
            DesktopAssert(s["desktop"]["paint"][s["desktopTabs"][index].Hwnd]["selected"], "Active candidate tab has selected styling")
            DesktopAssert(!s["desktop"]["paint"][s["desktopTabs"][3 - index].Hwnd]["selected"], "Previous candidate tab loses selected styling")
            DesktopAssert(s["addButton"].Enabled && s["openButton"].Enabled, "Selected candidate keeps existing row actions")
            DesktopAssert(s["triageButton"].Enabled = (index = 2), "Ignore action is enabled only for vocabulary")
            DesktopAssert(InStr(s["aiStatus"].Value, "Recommended (4/5)"), "AI assessment remains tied to selected candidate")
            DesktopAssert(DesktopShown(s["desktopAiSummary"]) && !DesktopShown(s["aiStatus"]), "Short assessment is plain text without a scrolling box")
            DesktopAssert(s["desktop"]["paint"][s["desktopTabs"][index].Hwnd]["kind"] = "segment", "Selected tab uses the selection-aware renderer")
            activeList := s[index = 2 ? "vocabularyList" : "sentenceList"]
            activeList.GetPos(, &listY,, &listH)
            s["desktopAiSummary"].GetPos(, &assessmentY,, &assessmentH)
            s["addButton"].GetPos(, &actionY,, &actionH)
            DesktopAssert(listY + listH <= assessmentY - 8,
                "Candidate table stays above assessment reasoning at " h "px")
            DesktopAssert(assessmentY + assessmentH <= actionY - 8,
                "Assessment reasoning stays above actions at " h "px")
            DesktopAssert(actionY + actionH <= h - 82,
                "Candidate actions stay above the footer at " h "px")
        }
        dpi := GetWindowDPI(hwnd) / 96
        TestDesktopStudyCapture(g, "desktop-anki-candidates-" w "x" h ".png", Round(w * dpi), Round(h * dpi))
    }
    for candidateTab in [1, 2] {
        StudyDesktopCandidatesChoose(s, candidateTab)
        candidateList := s[candidateTab = 1 ? "sentenceList" : "vocabularyList"]
        candidateList.Modify(1, "Select Focus")
        StudyCandidatesUpdateActions(s)
        DesktopAssert(InStr(s["aiStatus"].Value, sample["reason"]),
            "Initial candidate assessment belongs to the first row")
        ; Reproduce the native event order: the selected-row notification can
        ; arrive while the ListView still reports the preceding focused row.
        StudyCandidatesItemSelected(s, candidateList, 1, false)
        StudyCandidatesItemSelected(s, candidateList, 2, true)
        Sleep(45)
        DesktopAssert(candidateList.GetNext(0, "F") = 1
            && InStr(s["aiStatus"].Value, alternate["reason"]),
            "A stale deselection cannot overwrite the single-click " (candidateTab = 1 ? "sentence" : "vocabulary") " assessment")
        ; A pointer click has its own row-bearing event. Keep the ListView's
        ; reported focus deliberately on row 1 to prove the clicked row wins.
        StudyCandidatesItemSelected(s, candidateList, 2, false)
        StudyCandidatesRowClicked(s, candidateList, 2)
        Sleep(45)
        DesktopAssert(candidateList.GetNext(0, "F") = 1
            && InStr(s["aiStatus"].Value, alternate["reason"]),
            "The exact single-click " (candidateTab = 1 ? "sentence" : "vocabulary") " row refreshes its assessment")
    }
    DesktopDialogFixtureShow(s, 1040, 640)
    for activeTab in [1, 2, 1] {
        SendMessage(0x00F5, 0, 0, s["desktopTabs"][activeTab].Hwnd) ; BM_CLICK
        Sleep(35)
        inactiveTab := 3 - activeTab
        DesktopAssert(s["desktop"]["paint"][s["desktopTabs"][activeTab].Hwnd]["selected"]
            && !s["desktop"]["paint"][s["desktopTabs"][inactiveTab].Hwnd]["selected"],
            "Only the active candidate segment retains selected state")
        if activeTab = 2
            TestDesktopStudyCapture(g, "desktop-anki-candidate-tab-vocabulary.png", 1040, 640)
    }
    TestDesktopStudyCapture(g, "desktop-anki-candidate-tab-sentences.png", 1040, 640)
    StudyCandidatesSetRecommendationActivity(s, true, "Test provider", Map("sentences", 1, "vocabulary", 1))
    DesktopAssert(DesktopShown(s["progressText"]) && DesktopShown(s["progressBar"]) && !s["recommendButton"].Enabled, "Candidate assessment progress retains busy state")
    StudyCandidatesSetRecommendationActivity(s, false)
    DesktopAssert(!DesktopShown(s["progressText"]) && s["recommendButton"].Enabled, "Candidate assessment progress resets")
    s["vocabularyList"].Modify(0, "-Select")
    s["vocabularyList"].Modify(1, "Select Focus")
    sample["reason"] := ""
    Loop 60
        sample["reason"] .= "Long assessment detail remains readable in the scrollable area. "
    StudyCandidatesUpdateActions(s)
    DesktopAssert(!DesktopShown(s["desktopAiSummary"]) && DesktopShown(s["aiStatus"]), "Long assessment uses the scrollable detail area")
    DesktopDialogFixtureShow(s, 1040, 640)
    s["aiStatus"].GetPos(, &longAssessmentY,, &longAssessmentH)
    s["addButton"].GetPos(, &compactActionY)
    DesktopAssert(longAssessmentH = 64 && longAssessmentY + longAssessmentH <= compactActionY - 8,
        "Scrollable assessment reasoning remains fully reachable at 720p")

    CPSetWindowCloaked(g.Hwnd, true)
    StudyDesktopWindowAction(s, "maximize")
    Sleep(80)
    DesktopAssert(DllCall("user32\IsZoomed", "ptr", g.Hwnd),
        "Candidate window uses native taskbar-aware maximization")
    monitor := DllCall("user32\MonitorFromWindow", "ptr", g.Hwnd, "uint", 2, "ptr")
    monitorInfo := Buffer(40, 0), NumPut("uint", 40, monitorInfo)
    clientRect := Buffer(16, 0), clientBottomRight := Buffer(8, 0)
    DllCall("user32\GetMonitorInfoW", "ptr", monitor, "ptr", monitorInfo)
    DllCall("user32\GetClientRect", "ptr", g.Hwnd, "ptr", clientRect)
    NumPut("int", NumGet(clientRect, 8, "int"), "int", NumGet(clientRect, 12, "int"), clientBottomRight)
    DllCall("user32\ClientToScreen", "ptr", g.Hwnd, "ptr", clientBottomRight)
    clientBottom := NumGet(clientBottomRight, 4, "int")
    workBottom := NumGet(monitorInfo, 32, "int")
    DesktopAssert(clientBottom <= workBottom,
        "Maximized candidate window keeps its client area above the taskbar (" clientBottom " <= " workBottom ")")
    StudyDesktopWindowAction(s, "maximize")
    CPSetWindowCloaked(g.Hwnd, false)
    g.Destroy()
    DesktopAssert(!IsObject(StudyDesktopContext(hwnd)), "Candidate dialog cleans paint registry")
    for longMessage in [false, true] {
        message := longMessage ? "Long detail: " : "Add this vocabulary entry to Anki?`n`nDeck: Japanese`nNote type: Basic`nFields: Front / Back`nScreenshot: Include on the card back"
        if longMessage
            Loop 80
                message .= "A long diagnostic message should remain readable without overlapping the actions. "
        s := StudyDesktopMessageCreate(reader["gui"].Hwnd, message, "Add to Anki", "yesno", "Check the destination before sending this card.", "Add to Anki", "Back to review")
        g := s["gui"]
        for size in [[640, 400], [760, 520], [1040, 700]] {
            DesktopDialogFixtureShow(s, size[1], size[2])
            DesktopDialogAssertChrome(s, size[1], size[2])
            DesktopAssert(s["result"] = "No", "Confirmation defaults to not adding a card")
            if longMessage
                DesktopAssert(DesktopShown(s["overflow"]) && !DesktopShown(s["body"]), "Long messages scroll without overflowing")
            else if size[1] = 760
                DesktopAssert(DesktopShown(s["body"]) && !DesktopShown(s["overflow"]), "Ordinary confirmation has plain text, not a text box")
            dpi := GetWindowDPI(g.Hwnd) / 96
            TestDesktopStudyCapture(g, "desktop-anki-confirm-" longMessage "-" size[1] ".png", Round(size[1] * dpi), Round(size[2] * dpi))
        }
        hwnd := g.Hwnd
        PostMessage(0x10, 0, 0, hwnd)
        Sleep(25)
        DesktopAssert(s["closed"] && s["result"] = "No", "Confirmation cancel is non-destructive")
        DesktopAssert(!IsObject(StudyDesktopContext(hwnd)), "Confirmation cleans paint registry")
    }
    s := StudyDesktopMessageCreate(reader["gui"].Hwnd, "Added the vocabulary entry to Japanese.", "Added to Anki", "ok", "Anki card workflow", "Yes", "Cancel")
    DesktopDialogFixtureShow(s, 760, 460)
    DesktopAssert(!s.Has("addButton") && s["cancel"].Text = "OK", "Success notice has one acknowledgement action")
    DesktopAssert(DesktopShown(s["body"]), "Success notice uses plain text")
    CPThemedDialogFinish(s, s["gui"], "OK")
    library["gui"].Destroy(), reader["gui"].Destroy()
    CPStudyLibraryState := 0, CPStudyReaderState := 0, CPStudyCandidateState := 0
}

TestDesktopChapterControls(library) {
    slState := library, slProfileName := "Kabuki Den", slProfileLabel := slProfileName
    slDirectory := A_ScriptDir "\chapter-fixture", slCurrentChapter := "Chapter 2"
    slChapterHistory := ["Prologue", "Chapter 1", "Chapter 2"]
    ; @STUDY_CHAPTER_CONTROLS@
    return slDesktop
}

TestDesktopColumnsControls(library) {
    slState := library
    ; @STUDY_COLUMNS_CONTROLS@
    return slDesktop
}

TestDesktopFilterControls(library) {
    slState := library
    for key in ["profile", "chapter", "speaker", "tag"] {
        slState[key "Choices"] := [Map("mode", "all", "value", ""), Map("mode", "value", "value", "Example")]
        slState[key "Labels"] := ["Any " key, "Example"]
        slState[key "Mode"] := "all", slState[key "Filter"] := ""
    }
    slState["ankiMode"] := "all", slState["dateMode"] := "all"
    slState["dateFrom"] := "20260901000000", slState["dateTo"] := "20260908182359"
    ; @STUDY_FILTERS_CONTROLS@
    return slForm
}

TestDesktopDetailsControls(library) {
    slState := library
    slState["currentChapter"] := "Chapter 2", slState["currentSpeaker"] := "マリアン"
    slState["currentTags"] := "story, dialogue", slState["currentAddedToAnkiAt"] := ""
    ; @STUDY_DETAILS_CONTROLS@
    return slDesktop
}

TestDesktopMetadataDialogs(library) {
    constructors := Map("chapter", TestDesktopChapterControls, "columns", TestDesktopColumnsControls,
        "filters", TestDesktopFilterControls, "details", TestDesktopDetailsControls)
    minimums := Map("chapter", [760, 600], "columns", [620, 630], "filters", [780, 720], "details", [720, 600])
    for kind, constructor in constructors {
        s := constructor(library), g := s["gui"], hwnd := g.Hwnd, c := s["desktopControls"]
        DesktopAssert(s.Has("desktop"), kind " has a modern desktop shell")
        DesktopAssert(s["desktop"]["paint"][c["save"].Hwnd]["kind"] = "primary", kind " distinguishes the main action")
        for size in [minimums[kind], [1000, 800]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(s, w, h)
            DesktopDialogAssertChrome(s, w, h)
            for ctrl in g {
                if !DesktopShown(ctrl)
                    continue
                ctrl.GetPos(&x, &y, &cw, &ch)
                ; Combo dropdown height includes the popup, not just the field.
                DesktopAssert(x >= 0 && y >= 0 && x + cw <= w + 1 && y < h,
                    kind " control fits: " ctrl.Type)
                if ctrl.Type = "DDL" || ctrl.Type = "ComboBox"
                    DesktopAssert(CPDesktopIsCombo(ctrl.Hwnd), kind " uses modern combo painting")
                else
                    DesktopAssert(y + ch <= h + 1, kind " control does not overflow the footer")
                if ctrl.Type = "Edit"
                    DesktopAssert(!(WinGetExStyle(ctrl.Hwnd) & 0x200), kind " field has no bright native edge")
            }
            dpi := GetWindowDPI(hwnd) / 96
            TestDesktopStudyCapture(g, "desktop-study-dialog-" kind "-" w ".png", Round(w * dpi), Round(h * dpi))
        }
        if kind = "chapter" {
            c["chapter"].Text := "手入力の章"
            StudyDesktopDialogResize(s, g, 0, 760, 600)
            DesktopAssert(c["chapter"].Text = "手入力の章", "Editable chapter preserves typed Japanese after resize")
            c["chapter"].Choose(2)
            DesktopAssert(c["chapter"].Text = "Chapter 1", "Chapter history selection remains native")
        } else if kind = "columns" {
            for i, check in c["checks"] {
                DesktopAssert(check.Value = (library["columns"][i]["visible"] ? 1 : 0), "Column visibility draft matches saved value: " i " = " check.Value "/" library["columns"][i]["visible"])
                DesktopAssert(check.Enabled = !library["columns"][i]["required"], "Required Japanese source stays protected")
            }
            c["checks"][1].Value := !c["checks"][1].Value
            DesktopAssert(c["checks"][1].Value != (library["columns"][1]["visible"] ? 1 : 0), "Column changes are a local draft")
        } else if kind = "details" {
            c["chapter"].Value := "Edited chapter", c["tags"].Value := "edited"
            c["anki"].Value := 1
            StudyDesktopDialogResize(s, g, 0, 720, 600)
            DesktopAssert(c["chapter"].Value = "Edited chapter" && c["tags"].Value = "edited", "Details survive resizing")
            DesktopAssert(library["currentChapter"] = "Chapter 2" && library["currentAddedToAnkiAt"] = "", "Details and manual Anki state remain uncommitted")
        } else if kind = "filters" {
            for save in [false, true] {
                picker := StudyDesktopDatePickerCreate(s, c["date"], s["dateFrom"], c["from"], "From")
                for size in [[720, 550], [1000, 700]] {
                    DesktopDialogFixtureShow(picker, size[1], size[2])
                    DesktopDialogAssertChrome(picker, size[1], size[2])
                    dpi := GetWindowDPI(picker["gui"].Hwnd) / 96
                    TestDesktopStudyCapture(picker["gui"], "desktop-study-date-picker-" size[1] ".png", Round(size[1] * dpi), Round(size[2] * dpi))
                }
                StudyLibraryDatePickerAdjust(picker, 3, 1)
                DesktopAssert(picker["stamp"] = "20260902000000", "Date picker changes the requested component")
                DesktopAssert(s["dateFrom"].Value = "20260901000000", "Date picker edits stay local until confirmed")
                pickerHwnd := picker["gui"].Hwnd
                StudyLibraryDatePickerClose(picker, save)
                DesktopAssert(s["dateFrom"].Value = (save ? "20260902000000" : "20260901000000"), "Date picker honors Use and Cancel")
                DesktopAssert(c["date"].Value = (save ? 6 : 1), "Confirmed date selects custom range only")
                DesktopAssert(library["dateFrom"] = "20260901000000", "Date picker never applies filters itself")
                DesktopAssert(!StudyLibraryDatePickerRegistry().Has(pickerHwnd) && !IsObject(StudyDesktopContext(pickerHwnd)), "Date picker cleans both registries")
            }
            c["profile"].Choose(2)
            DesktopAssert(library["profileMode"] = "all", "Filter selection is a local draft")
        }
        ; Exercise the real native Close event, never an action that writes data.
        PostMessage(0x10, 0, 0, hwnd)
        Sleep(35)
        DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd) && !IsObject(StudyDesktopContext(hwnd)), kind " closes and unregisters cleanly")
    }
}

TestDesktopConnectionControls(library) {
    global iniPath
    slState := library, saProfiles := ["Kabuki Den", "Another study profile"]
    ; @STUDY_ANKI_CONNECTION_CONTROLS@
    return saState
}

TestDesktopBulkControls(library) {
    slState := library, slIds := [1, 2, 3]
    ; @STUDY_BULK_DETAILS_CONTROLS@
    return slDesktop
}

TestDesktopNewVersionControls(reader) {
    global model_openai_explain, model_gemini_explain, explainProvider, explainOpenAIModel, explainGeminiModel
    global explainPromptProfile, explainPromptsDir
    model_openai_explain := ["gpt-4o"], model_gemini_explain := ["gemini-3.5-flash", "gemini-other"]
    explainProvider := "gemini", explainOpenAIModel := "gpt-4o", explainGeminiModel := "gemini-3.5-flash"
    explainPromptProfile := "fixture", explainPromptsDir := A_ScriptDir "\fixture-prompts"
    DirCreate(explainPromptsDir)
    if !FileExist(explainPromptsDir "\fixture.txt")
        FileAppend("Explain this Japanese text: {jp}", explainPromptsDir "\fixture.txt", "UTF-8")
    srState := reader
    srState["currentProvider"] := "gemini", srState["currentModel"] := "gemini-3.5-flash", srState["currentPrompt"] := "fixture"
    ; @STUDY_NEW_VERSION_CONTROLS@
    return srNewState
}

TestDesktopSecondaryStudyDialogs(library, reader) {
    connection := TestDesktopConnectionControls(library)
    c := connection["controls"]
    c["status"].Text := "AnkiConnect is available. Choose a Study Profile and map its fields."
    for key in ["deck", "model", "japanese", "explanation"] {
        c[key].Add([key = "deck" ? "Japanese::Game dialogue" : key = "model" ? "Basic" : key = "japanese" ? "Front" : "Back"])
        c[key].Choose(1)
    }
    DesktopTestStudyToolSizes(connection, [[820, 760], [900, 760], [1200, 900]])
    DesktopAssert(c["profile"].Text = "Kabuki Den" && c["deck"].Text = "Japanese::Game dialogue", "Connection mapping selections survive restyling")
    DesktopDialogTestClose(connection)
    bulk := TestDesktopBulkControls(library), c := bulk["controls"]
    DesktopTestStudyToolSizes(bulk, [[820, 720], [900, 760], [1200, 900]])
    for key in ["chapter", "speaker", "tags"] {
        DesktopAssert(c[key "Mode"].Value = 1 && !c[key].Enabled, "Bulk starts at Keep existing: " key)
        c[key "Mode"].Choose(2)
        StudyLibraryBulkValueModeChanged(c[key "Mode"], c[key], key = "tags" ? "2,3,4" : "2")
        DesktopAssert(c[key].Enabled, "Bulk Set/Add enables value: " key)
        c[key].Value := "Unsaved synthetic value"
    }
    DesktopDialogTestClose(bulk)
    version := TestDesktopNewVersionControls(reader), c := version["controls"]
    DesktopTestStudyToolSizes(version, [[820, 620], [820, 680], [920, 740], [1200, 900]])
    for provider in ["openai", "gemini"] {
        c["provider"].Choose(provider = "gemini" ? 1 : 2)
        StudyReaderNewVersionProviderChanged(version)
        for key in ["gemini", "openai"]
            DesktopAssert(c[key].Visible = (key = provider) && c[key].Enabled = (key = provider)
                && c[key "Label"].Visible = (key = provider), "Only selected provider model is visible: " key)
    }
    for mode in ["name", "edit"] {
        prompt := StudyReaderPromptDialogCreate(version, mode, mode = "edit" ? "Explain Japanese. 日本語 {jp}" : "", "fixture")
        DesktopTestStudyToolSizes(prompt, mode = "edit" ? [[820, 620], [1000, 680], [1000, 780]] : [[680, 450], [760, 480]])
        if mode = "edit" {
            editor := prompt["controls"]["editor"]
            editor.Focus()
            SendMessage(0x00B1, 0, -1, editor.Hwnd) ; Select all, then verify the open-page caret policy.
            StudyReaderPromptPlaceInitialCaret(prompt)
            selection := Buffer(8, 0)
            SendMessage(0x00B0, selection.Ptr, selection.Ptr + 4, editor.Hwnd)
            DesktopAssert(NumGet(selection, 0, "uint") = 0 && NumGet(selection, 4, "uint") = 0,
                "Explanation prompt opens with a caret instead of selecting all text")
        }
        prompt["controls"]["editor"].Value := "Unsaved changes"
        DesktopDialogTestClose(prompt)
        DesktopAssert(prompt["result"].Result = "Cancel", "Prompt close discards unconfirmed result: " mode)
    }
    prompt := StudyReaderPromptDialogCreate(version, "name")
    prompt["controls"]["editor"].Value := "New fixture prompt"
    StudyReaderPromptDialogClose(prompt, true)
    DesktopAssert(prompt["result"].Result = "OK" && prompt["result"].Value = "New fixture prompt", "Prompt name accepts only after explicit Create")
    DesktopAssert(!FileExist(ExplainProfilePath("New fixture prompt")), "Name dialog does not write a prompt file or generate on its own")
    TestReaderPromptFullscreen(version)
    TestReaderPromptModal(version)
    DesktopDialogTestClose(version)
    DesktopAssert(!reader["newVersionDialog"], "Closing new version releases Reader state")
}

DesktopTestStudyToolSizes(s, sizes) {
    for size in sizes {
        w := size[1], h := size[2]
        DesktopDialogFixtureShow(s, w, h)
        DesktopDialogAssertChrome(s, w, h)
        DesktopAssert(s["desktop"]["paint"][s["addButton"].Hwnd]["kind"] = "primary", "Study tool has a primary action")
        for key, ctrl in s["controls"] {
            if !ctrl.Visible
                continue
            ctrl.GetPos(&x, &y, &cw, &ch)
            DesktopAssert(x >= 24 && x + cw <= w - 24 && y >= 174, "Study tool control fits: " key)
            if ctrl.Type != "DDL"
                DesktopAssert(y + ch <= (s["desktop"]["surfaces"][ctrl.Hwnd] = "bar" ? h - 16 : h - 86), "Study tool control does not bleed into footer: " key)
            if ctrl.Type = "Edit"
                DesktopAssert(!(WinGetExStyle(ctrl.Hwnd) & 0x200) && !(WinGetStyle(ctrl.Hwnd) & 0x800000), "Study tool editor has no native white edge")
        }
        dpi := GetWindowDPI(s["gui"].Hwnd) / 96
        TestDesktopStudyCapture(s["gui"], "desktop-tool-" s["desktop"]["kind"] "-" w ".png", Round(w * dpi), Round(h * dpi))
    }
}

TestReaderPromptFullscreen(version) {
    parent := Map("gui", version["gui"], "bigBoxPresentation", true)
    for mode in ["name", "edit"] {
        s := StudyReaderPromptDialogCreate(parent, mode, mode = "edit" ? "Explain this Japanese dialogue. 日本語 {jp}" : "", "fixture")
        f := s["bigBoxForm"], g := s["gui"]
        for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
            w := size[1], h := size[2]
            g.Show("Hide w" w " h" h)
            StudyCandidatesRecommendationBigBoxResize(f, g, 0, w, h)
            CPApplyOwnedDialogTheme(g)
            g.Show("NA x-12000 y-12000 w" w " h" h)
            Sleep(35)
            DesktopAssert(!(WinGetStyle(g.Hwnd) & 0xC00000), "Fullscreen prompt has no native caption")
            f["shell"]["footer"].GetPos(, &footerY)
            for key, ctrl in s["controls"] {
                ctrl.GetPos(&x, &y, &cw, &ch)
                DesktopAssert(x >= 0 && y >= 0 && x + cw <= w && y + ch < footerY, "Fullscreen prompt fits above navigation footer: " key)
                if ctrl.Type = "Button"
                    DesktopAssert(DesktopTextWidth(ctrl) < cw - 12, "Fullscreen prompt button label fits: " key)
            }
            editor := s["controls"]["editor"], save := s["controls"]["save"], cancel := s["controls"]["cancel"]
            StudyControllerSetFocus(g.Hwnd, editor.Hwnd)
            SendMessage(0xB1, 3, 8, editor.Hwnd)
            StudyControllerMoveFocus(g.Hwnd, "Down")
            DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = save.Hwnd, "Fullscreen Down reaches prompt Save/Create")
            StudyControllerMoveFocus(g.Hwnd, "Right")
            DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = cancel.Hwnd, "Fullscreen Right reaches prompt Close/Cancel")
            StudyControllerMoveFocus(g.Hwnd, "Up")
            DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = editor.Hwnd, "Fullscreen Up returns to prompt editor")
            if mode = "edit" {
                selection := Buffer(8, 0)
                SendMessage(0xB0, selection.Ptr, selection.Ptr + 4, editor.Hwnd)
                DesktopAssert(NumGet(selection, 0, "uint") = 3 && NumGet(selection, 4, "uint") = 8, "Fullscreen navigation preserves prompt text selection")
            }
            TestDesktopStudyCapture(g, "fullscreen-reader-prompt-" mode "-" w ".png", w, h)
        }
        StudyReaderPromptDialogClose(s)
        DesktopAssert(s["result"].Result = "Cancel", "Fullscreen prompt closes without accepting")
    }
}

TestReaderPromptModal(version) {
    for initiallyEnabled in [true, false] {
        owner := version["gui"].Hwnd
        DllCall("user32\EnableWindow", "ptr", owner, "int", initiallyEnabled)
        s := StudyReaderPromptDialogCreate(version, "name")
        closeTimer := TestReaderPromptModalClose.Bind(s, owner)
        SetTimer(closeTimer, -100)
        result := StudyReaderPromptDialogRun(s)
        DesktopAssert(result.Result = "Cancel", "Modal prompt cancellation returns Cancel")
        DesktopAssert(!!DllCall("user32\IsWindowEnabled", "ptr", owner) = initiallyEnabled, "Modal prompt restores parent's previous enabled state")
    }
    DllCall("user32\EnableWindow", "ptr", version["gui"].Hwnd, "int", 1)
}

TestReaderPromptModalClose(s, owner) {
    DesktopAssert(!DllCall("user32\IsWindowEnabled", "ptr", owner), "Parent is disabled while prompt is modal")
    StudyReaderPromptDialogClose(s)
}

TestSharedChoicePopup(ownerHwnd, choices, cpResultState := 0) {
    global controlDarkMode
    ; @SHARED_CHOICE_POPUP_CONTROLS@
    return cpPopupState
}

TestSharedContextPopup(ownerHwnd, items) {
    global controlDarkMode
    ; @SHARED_CONTEXT_POPUP_CONTROLS@
    return cpPopupState
}

TestNativePickerBackend(result, mode, spec) {
    global TestNativePickerSpecs, TestNativePickerDepths, CPBigBoxModalDepth
    TestNativePickerSpecs.Push(spec)
    TestNativePickerDepths.Push(CPBigBoxModalDepth)
    if spec["owner"]
        DllCall("user32\EnableWindow", "ptr", spec["owner"], "int", 0)
    if mode = "throw"
        throw Error("Synthetic native picker failure")
    return result
}

TestCapturePopupFinish(action, *) {
    global TestCapturePopupSeen, TestCapturePopupTimer
    for hwnd, s in CPThemedChoicePopupRegistry() {
        if s["rows"].Length != 3
            continue
        labels := []
        for row in s["rows"]
            labels.Push(Trim(row.Text))
        if labels[1] != "Capture region" || labels[2] != "Capture window"
            || labels[3] != "Cancel"
            continue
        SetTimer(TestCapturePopupTimer, 0)
        TestCapturePopupSeen := true
        s["gui"].GetClientPos(,, &w, &h)
        dpi := GetWindowDPI(s["gui"].Hwnd) / 96
        for row in s["rows"] {
            row.GetPos(,, &rowW, &rowH)
            DesktopAssert(rowH = 40 && DesktopTextWidth(row) < rowW - 8,
                "Capture mode uses readable modern action rows")
        }
        if action = "cancel" {
            TestDesktopStudyCapture(
                s["gui"], "capture-mode-desktop.png",
                Round(w * dpi), Round(h * dpi)
            )
            CPThemedChoicePopupFinish(s, s["gui"], 3)
        } else {
            CPThemedChoicePopupFocus(s, 2)
            CPThemedChoicePopupOnKeyDown(
                0x0D, 0, 0x0100, s["rows"][2].Hwnd
            )
        }
        return
    }
}

TestPickerPolish() {
    global ui, CPDesktop, CPBigBoxModalDepth, controlDarkMode
    global TestNativePickerSpecs := [], TestNativePickerDepths := []
    global TestCapturePopupSeen := false, TestCapturePopupTimer := 0
    originalDark := controlDarkMode
    controlDarkMode := 1

    owner := Gui("+Owner" ui.Hwnd, "Native picker owner fixture")
    owner.CPDialogPresentation := "desktop"
    first := owner.AddButton("x20 y20 w160 h36", "Browse files")
    second := owner.AddButton("x20 y70 w160 h36", "Other action")
    owner.Show("x-9000 y-9000 w220 h130")
    WinActivate("ahk_id " owner.Hwnd)
    first.Focus()

    selected := CPNativeFileSelect(
        owner.Hwnd, 3, "C:\picker-start", "Select helper",
        "Programs (*.exe)", TestNativePickerBackend.Bind(
            "C:\picked\helper.exe", "return"
        )
    )
    spec := TestNativePickerSpecs[1]
    DesktopAssert(selected = "C:\picked\helper.exe"
        && spec["kind"] = "file" && spec["owner"] = owner.Hwnd
        && spec["options"] = 3 && spec["root"] = "C:\picker-start"
        && spec["prompt"] = "Select helper"
        && spec["filter"] = "Programs (*.exe)"
        && spec["presentation"] = "desktop",
        "File picker wrapper preserves owner, defaults, title and filter")
    DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", owner.Hwnd)
        && DllCall("user32\GetFocus", "ptr") = first.Hwnd,
        "File picker restores its enabled owner and prior control focus")

    DllCall("user32\EnableWindow", "ptr", owner.Hwnd, "int", 0)
    selected := CPNativeDirSelect(
        owner.Hwnd, "C:\captures", 1, "Select screenshot folder",
        TestNativePickerBackend.Bind("C:\picked\captures", "return")
    )
    spec := TestNativePickerSpecs[2]
    DesktopAssert(selected = "C:\picked\captures"
        && spec["kind"] = "folder" && spec["root"] = "C:\captures"
        && spec["options"] = 1 && spec["prompt"] = "Select screenshot folder"
        && spec["filter"] = "",
        "Folder picker wrapper preserves its native arguments")
    DesktopAssert(!DllCall("user32\IsWindowEnabled", "ptr", owner.Hwnd),
        "Picker restores a parent that was already disabled")
    DllCall("user32\EnableWindow", "ptr", owner.Hwnd, "int", 1)

    fullscreen := Gui("-Caption -DPIScale", "Fullscreen picker owner fixture")
    fullscreen.CPDialogPresentation := "fullscreen"
    fullButton := fullscreen.AddButton("x20 y20 w180 h40", "Browse fullscreen")
    fullscreen.Show("x-9000 y-9000 w240 h90")
    WinActivate("ahk_id " fullscreen.Hwnd)
    fullButton.Focus()
    CPBigBoxModalDepth := 0
    failed := false
    try CPNativeFileSelect(
        fullscreen.Hwnd, "S16", "C:\export.xlsx", "Export Study Library",
        "Excel Workbook (*.xlsx)", TestNativePickerBackend.Bind("", "throw")
    )
    catch as ex
        failed := ex.Message = "Synthetic native picker failure"
    DesktopAssert(failed && TestNativePickerDepths[3] = 1
        && CPBigBoxModalDepth = 0,
        "Fullscreen picker suspends topmost mode and restores it after errors")
    DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", fullscreen.Hwnd)
        && DllCall("user32\GetFocus", "ptr") = fullButton.Hwnd,
        "Fullscreen picker restores its owner and controller focus after errors")
    fullscreen.Destroy()

    ui.Show("x-9000 y-9000 w900 h640")
    WinActivate("ahk_id " ui.Hwnd)
    TestCapturePopupSeen := false
    TestCapturePopupTimer := TestCapturePopupFinish.Bind("cancel")
    SetTimer(TestCapturePopupTimer, 20)
    selection := OpenCapturePicker()
    SetTimer(TestCapturePopupTimer, 0)
    DesktopAssert(TestCapturePopupSeen && selection = 3,
        "Desktop Capture opens the shared modern chooser and Cancel is inert (seen="
            TestCapturePopupSeen ", selection=" selection ", presentation="
            CPDialogPresentation(ui.Hwnd) ")")

    TestCapturePopupSeen := false
    activation := Map()
    TestCapturePopupTimer := TestCapturePopupFinish.Bind("keyboard")
    SetTimer(TestCapturePopupTimer, 20)
    selection := CPThemedChoicePopup(
        ui.Hwnd, 0, ["Capture region", "Capture window", "Cancel"],
        activation
    )
    SetTimer(TestCapturePopupTimer, 0)
    DesktopAssert(TestCapturePopupSeen && selection = 2
        && activation["result"] = 2
        && activation["activationSource"] = "controller",
        "Capture chooser preserves keyboard/controller activation semantics")

    owner.Destroy()
    ui.Hide()
    controlDarkMode := originalDark
}

TestSharedDialogs(library, reader) {
    global ui, CPDesktop, controlDarkMode
    DesktopAssert(CPDialogPresentation(ui.Hwnd) = "desktop", "Main desktop is recognized for shared dialogs")
    child := Gui("+Owner" reader["gui"].Hwnd)
    grandchild := Gui("+Owner" child.Hwnd)
    DesktopAssert(CPDialogPresentation(grandchild.Hwnd) = "desktop", "Nested desktop dialog inherits presentation")
    fullscreen := Gui("+Owner" ui.Hwnd " -Caption -DPIScale")
    fullscreen.CPDialogPresentation := "fullscreen"
    nestedFullscreen := Gui("+Owner" fullscreen.Hwnd)
    DesktopAssert(CPDialogPresentation(nestedFullscreen.Hwnd) = "fullscreen", "Explicit fullscreen owner wins over desktop ancestor")
    DesktopAssert(CPDialogPresentation(0) = "classic", "Invalid owner never attaches to unrelated foreground window")
    DesktopAssert(CPDialogPresentation(ui.Hwnd) = "desktop" && !CPDesktop.Has("modern"),
        "Main desktop presentation remains modern-only")
    choices := ["Copy current section", "Copy full explanation", "Copy Japanese without readings"]
    items := [Map("label", "Open in Reader…"), Map("separator", true),
        Map("label", "Unavailable action", "enabled", false), Map("label", "Edit details…"), Map("label", "Remove explanation…")]
    for dark in [1, 0] {
        controlDarkMode := dark
        for s in [TestSharedChoicePopup(reader["gui"].Hwnd, choices), TestSharedContextPopup(reader["gui"].Hwnd, items)] {
            g := s["gui"], g.Show("Hide AutoSize")
            CPApplyOwnedDialogTheme(g), CPDialogPopupFinishLayout(s)
            CPThemedChoicePopupRegister(s)
            g.Show("NA x-12000 y-12000")
            CPThemedChoicePopupResetRows(s), CPThemedChoicePopupFocus(s, 1)
            for row in s["rows"] {
                row.GetPos(,, &rowW, &rowH)
                DesktopAssert(rowH = 40 && !(WinGetStyle(row.Hwnd) & 0x800000), "Modern popup has roomy rows without native borders")
                DesktopAssert(DesktopTextWidth(row) < rowW - 8, "Modern popup text fits")
            }
            CPThemedChoicePopupFocus(s, 2)
            DesktopAssert(s["focusIndex"] = 2, "Popup keyboard selection changes without executing")
            CPThemedChoicePopupMouseMove(0, 0, 0x200, s["rows"][1].Hwnd)
            DesktopAssert(s["focusIndex"] = 1 && !s["closed"], "Popup hover highlights without executing")
            CPThemedChoicePopupFocus(s, 2)
            if s.Has("resultValues")
                DesktopAssert(s["resultValues"][2] = 4, "Context menu preserves original indices while skipping separator/disabled item")
            g.GetClientPos(,, &pw, &ph)
            dpi := GetWindowDPI(g.Hwnd) / 96
            TestDesktopStudyCapture(g, "shared-popup-" (s.Has("resultValues") ? "context" : "choice") "-" dark ".png", Round(pw * dpi), Round(ph * dpi))
            hwnd := g.Hwnd
            CPThemedChoicePopupFinish(s, g, 0), CPThemedChoicePopupUnregister(s)
            DesktopAssert(!CPThemedChoicePopupRegistry().Has(hwnd) && s["result"] = 0, "Popup cancel cleans registry and has no action result")
        }
    }
    controlDarkMode := 1
    fullItems := items.Clone()
    Loop 11
        fullItems.Push(Map("label", "Action " A_Index))
    s := CPFullscreenMenuCreate(fullscreen.Hwnd, fullItems, "Explanation actions")
    for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
        w := size[1], h := size[2], g := s["gui"]
        g.Show("Hide w" w " h" h)
        StudyCandidatesRecommendationBigBoxResize(s, g, 0, w, h)
        StudyCandidatesRecommendationBigBoxApplyTheme(s)
        g.Show("NA x-12000 y-12000")
        for page in [1, 2, 3] {
            CPFullscreenMenuPage(s, page - s["page"])
            s["shell"]["footer"].GetPos(, &footerY)
            for key, ctrl in s["controls"] {
                if !ctrl.Visible
                    continue
                ctrl.GetPos(&x, &y, &cw, &ch)
                DesktopAssert(x >= 0 && y >= 0 && x + cw <= w && y + ch < footerY, "Fullscreen menu control fits: " key)
                if ctrl.Type = "Button"
                    DesktopAssert(DesktopTextWidth(ctrl) < cw - 12, "Fullscreen menu button label fits: " key)
            }
            TestDesktopStudyCapture(g, "shared-fullscreen-menu-" w "-" page ".png", w, h)
        }
    }
    CPFullscreenMenuPage(s, 1 - s["page"])
    StudyControllerMoveFocus(s["gui"].Hwnd, "Down")
    DesktopAssert(StudyControllerFocusedHwnd(s["gui"].Hwnd) = s["slots"][3].Hwnd, "Fullscreen navigation skips disabled action")
    CPFullscreenMenuChoose(s, 2)
    DesktopAssert(!s["closed"], "Disabled fullscreen option cannot execute")
    CPFullscreenMenuChoose(s, 3)
    DesktopAssert(s["closed"] && s["result"] = 4, "Fullscreen action returns original item index")
    s := CPFullscreenMenuCreate(fullscreen.Hwnd, items)
    CPFullscreenMenuClose(s)
    DesktopAssert(s["result"] = 0, "Fullscreen Back makes no selection")
    threeWay := StudyDesktopMessageCreate(child.Hwnd,
        "The current settings contain unsaved edits.", "Unsaved settings", "yesnocancel",
        "Choose what to do with these changes.",
        "Save changes", "Discard changes", "Cancel")
    for size in [[640, 400], [760, 520], [1040, 700]] {
        DesktopDialogFixtureShow(threeWay, size[1], size[2])
        DesktopDialogAssertChrome(threeWay, size[1], size[2])
        threeWay["addButton"].GetPos(&yesX, &yesY, &yesW, &yesH)
        threeWay["noButton"].GetPos(&noX, &noY, &noW, &noH)
        threeWay["cancel"].GetPos(&cancelX, &cancelY, &cancelW, &cancelH)
        DesktopAssert(yesX >= 24 && cancelX + cancelW <= size[1] - 24,
            "Three-way message actions fit the desktop footer")
        DesktopAssert(yesY = noY && noY = cancelY && yesX + yesW < noX
            && noX + noW < cancelX, "Three-way message actions stay aligned and separate")
    }
    threeWayHwnd := threeWay["gui"].Hwnd
    threeWay["gui"].Destroy()
    DesktopAssert(!IsObject(StudyDesktopContext(threeWayHwnd)), "Three-way message clears paint registry")
    TestSharedMessageRoute(child.Hwnd, "desktop", "yesno", "warning")
    TestSharedMessageRoute(nestedFullscreen.Hwnd, "fullscreen", "yesno", "error")
    TestSharedMessageRoute(child.Hwnd, "desktop", "yesnocancel", "warning", "yes")
    TestSharedMessageRoute(child.Hwnd, "desktop", "yesnocancel", "warning", "no")
    TestSharedMessageRoute(child.Hwnd, "desktop", "yesnocancel", "warning", "cancel")
    TestSharedMessageRoute(nestedFullscreen.Hwnd, "fullscreen", "yesnocancel", "warning")
    TestSharedMessageRoute(child.Hwnd, "desktop", "ok", "info")
    TestSharedMessageRoute(nestedFullscreen.Hwnd, "fullscreen", "ok", "info")
    controlDarkMode := 0
    TestSharedMessageRoute(child.Hwnd, "desktop", "ok", "info")
    DllCall("user32\EnableWindow", "ptr", child.Hwnd, "int", 0)
    TestSharedMessageRoute(child.Hwnd, "desktop", "yesno", "warning")
    DllCall("user32\EnableWindow", "ptr", child.Hwnd, "int", 1)
    controlDarkMode := 1
    grandchild.Destroy(), child.Destroy(), nestedFullscreen.Destroy(), fullscreen.Destroy()
}

TestModelCatalogFixture(source := "online", models := 0, warnings := 0) {
    return Map("ok", true, "source", source, "models", IsObject(models) ? models : ["gemini-existing", "model-alpha", "model-beta", "models/very-long-provider-model-id-for-translation-with-a-dated-snapshot-2026-09-08-preview-and-regional-variant"],
        "warnings", IsObject(warnings) ? warnings : [])
}

TestAuthoringDialogs(reader) {
    global controlDarkMode, promptsDir, explainPromptsDir, glossariesDir
    global ddlPrompt, ddlEPr, ddlJPG, ddlENG
    global promptProfile, explainPromptProfile, jp2enGlossaryProfile, en2enGlossaryProfile
    global TestAuthoringMessageCaptured
    root := A_ScriptDir "\authoring-fixtures"
    promptsDir := root "\translation-prompts", explainPromptsDir := root "\explanation-prompts"
    glossariesDir := root "\glossaries"
    for dir in [promptsDir, explainPromptsDir, glossariesDir]
        DirCreate(dir)
    FileAppend("Translate the Japanese text. 日本語", PromptFilePath("fixture"), "UTF-8")
    FileAppend("Explain the Japanese text. 日本語 {jp}", ExplainProfilePath("fixture"), "UTF-8")
    FileAppend("Legacy explanation prompt. 日本語 {jp}", ExplainPromptFilePath(), "UTF-8")
    promptProfile := "fixture", explainPromptProfile := "fixture"
    ddlPrompt.Delete(), ddlPrompt.Add(["fixture"]), ddlPrompt.Choose(1)
    ddlEPr.Delete(), ddlEPr.Add(["fixture"]), ddlEPr.Choose(1)

    editorOpeners := [OpenPromptEditor, OpenExplainPromptEditor_Multi, OpenExplainPromptEditor]
    for index, openEditor in editorOpeners {
        s := openEditor.Call()
        DesktopAssert(IsObject(s) && s.Has("desktop"), "Prompt editor uses modern desktop shell: " index)
        c := s["controls"], g := s["gui"], initial := c["editor"].Value
        for size in [[760, 600], [900, 720], [1200, 860]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(s, w, h)
            DesktopDialogAssertChrome(s, w, h)
            c["editor"].GetPos(&x, &y, &cw, &ch)
            DesktopAssert(c["editor"].Value = initial && ch >= 200, "Prompt text survives responsive layout: " index)
            DesktopAssert(!(WinGetExStyle(c["editor"].Hwnd) & 0x200)
                && !(WinGetStyle(c["editor"].Hwnd) & 0x800000), "Prompt editor has no classic white frame")
            DesktopAssert(y + ch < h - 140, "Prompt editor leaves guidance and actions visible")
            if index = 1 && w = 900 {
                dpi := GetWindowDPI(g.Hwnd) / 96
                TestDesktopStudyCapture(g, "authoring-prompt-desktop.png", Round(w * dpi), Round(h * dpi))
            }
        }
        if index = 1 {
            c["editor"].Value := "Saved prompt draft. 日本語"
            SendMessage(0xF5, 0, 0, c["save"].Hwnd)
            Sleep(25)
            DesktopAssert(FileRead(s["path"], "UTF-8") = "Saved prompt draft. 日本語", "Prompt Save writes the edited text")
            DesktopAssert(FileExist(s["path"] ".bak"), "Prompt Save keeps the previous file as a backup")
        }
        hwnd := g.Hwnd
        CPTextEditorDialogClose(s)
        DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd)
            && !IsObject(StudyDesktopContext(hwnd)), "Prompt editor closes and unregisters cleanly")
    }

    jp2enGlossaryProfile := "default", en2enGlossaryProfile := "default"
    ddlJPG.Delete(), ddlJPG.Add(["default"]), ddlJPG.Choose(1)
    ddlENG.Delete(), ddlENG.Add(["default"]), ddlENG.Choose(1)
    path := GlossaryEnsureFile("jp", "default")
    SaveTextAtomic(path, GlossaryHeader("jp") "`r`n魔王 -> demon king`r`n王国 -> kingdom`r`n")
    manager := OpenGlossaryManager("jp")
    DesktopAssert(IsObject(manager) && manager.Has("desktop"), "Terminology table uses modern desktop shell")
    DesktopAssert(manager["list"].GetCount() = 2, "Terminology table preserves parsed entries")
    DesktopAssert(manager["controls"]["intro"].Text != manager["desktop"]["chrome"]["subtitle"].Text,
        "Terminology table uses complementary header and row guidance")
    for size in [[780, 600], [900, 700], [1200, 860]] {
        w := size[1], h := size[2]
        DesktopDialogFixtureShow(manager, w, h)
        DesktopDialogAssertChrome(manager, w, h)
        manager["list"].GetPos(&x, &y, &cw, &ch)
        DesktopAssert(ch >= 160 && y + ch < h - 130, "Terminology rows stay clear of status and footer")
        DesktopAssert(!(SendMessage(0x1037, 0, 0, manager["list"].Hwnd) & 1), "Modern terminology table removes native grid lines")
        DesktopAssert(manager["list"].GetText(1, 1) = "魔王", "Terminology text survives resize")
        if w = 900 {
            dpi := GetWindowDPI(manager["gui"].Hwnd) / 96
            TestDesktopStudyCapture(manager["gui"], "authoring-terminology-manager.png", Round(w * dpi), Round(h * dpi))
        }
    }

    manager["list"].Modify(1, "Select Focus Vis")
    entry := GlossaryEntryDialog(manager, 1)
    DesktopAssert(IsObject(entry) && entry.Has("desktop"), "Terminology entry uses modern desktop shell")
    DesktopAssert(entry["gui"].Title = "Edit model instruction"
        && !InStr(entry["desktop"]["chrome"]["subtitle"].Text, "->"), "Terminology entry uses plain-language modern labels")
    DesktopAssert(!DllCall("user32\IsWindowEnabled", "ptr", manager["gui"].Hwnd), "Entry editor blocks its manager")
    for size in [[620, 540], [720, 600], [960, 720]] {
        w := size[1], h := size[2]
        DesktopDialogFixtureShow(entry, w, h)
        DesktopDialogAssertChrome(entry, w, h)
        DesktopAssert(entry["controls"]["source"].Value = "魔王"
            && entry["controls"]["target"].Value = "demon king", "Entry fields preserve exact source and replacement")
        if w = 720 {
            dpi := GetWindowDPI(entry["gui"].Hwnd) / 96
            TestDesktopStudyCapture(entry["gui"], "authoring-terminology-entry.png", Round(w * dpi), Round(h * dpi))
        }
    }
    entry["controls"]["source"].Value := "勇者"
    entry["controls"]["target"].Value := "hero"
    SendMessage(0xF5, 0, 0, entry["controls"]["save"].Hwnd)
    Sleep(35)
    DesktopAssert(entry["closed"] && DllCall("user32\IsWindowEnabled", "ptr", manager["gui"].Hwnd), "Saving closes entry editor and restores manager")
    DesktopAssert(InStr(FileRead(path, "UTF-8"), "勇者 -> hero"), "Entry Save keeps the existing atomic glossary workflow")

    TestAuthoringMessageCaptured := false
    timer := TestAuthoringMessageClose
    SetTimer(timer, 80)
    try result := GlossaryOwnedMessage(manager["gui"].Hwnd,
        "Delete this synthetic terminology entry?", "Terminology overrides", "yesno")
    finally SetTimer(timer, 0)
    DesktopAssert(TestAuthoringMessageCaptured && result = 7, "Terminology confirmation uses modern dialog with No as safe default")

    SaveTextAtomic(path, GlossaryHeader("jp") "`r`nmalformed line`r`n")
    raw := OpenRawGlossaryEditor("jp", "default")
    DesktopAssert(IsObject(raw) && raw.Has("desktop"), "Malformed terminology uses modern raw repair editor")
    DesktopDialogFixtureShow(raw, 900, 720)
    DesktopAssert(InStr(raw["controls"]["editor"].Value, "malformed line"), "Raw editor never discards malformed text")
    raw["controls"]["editor"].Value := "unsaved replacement"
    CPTextEditorDialogClose(raw)
    DesktopAssert(InStr(FileRead(path, "UTF-8"), "malformed line"), "Closing raw editor discards its unsaved draft")
    GlossaryManagerClose(manager)

    fullscreen := Gui("+Owner" reader["gui"].Hwnd " -Caption -DPIScale")
    fullscreen.CPDialogPresentation := "fullscreen"
    DesktopAssert(!CPTextEditorDialogCreate(fullscreen.Hwnd, "Synthetic", "Fullscreen", "Text", "Draft", "Hint"),
        "Fullscreen continues to use its established in-page prompt and terminology editors")
    fullscreen.Destroy()
}

TestAuthoringMessageClose() {
    global TestAuthoringMessageCaptured
    for hwnd, d in StudyDesktopRegistry() {
        if d["kind"] != "message" || d["state"]["gui"].Title != "Terminology overrides" || d["width"] < 1
            continue
        TestAuthoringMessageCaptured := true
        s := d["state"]
        CPThemedDialogFinish(s, s["gui"], "No")
        return
    }
}

TestControlDialogs(reader) {
    global controlDarkMode, TestControlMessageCaptured, TestControlMessageTitle
    DesktopAssert(HotkeyPretty("^+{F10}") = "Ctrl + Shift + F10"
        && HotkeyPretty("!PgUp") = "Alt + Page Up",
        "Shortcut display keeps function and named keys readable")
    for dark in [1, 0] {
        controlDarkMode := dark
        hotkeyState := CPHotkeyDialogCreate(reader["gui"].Hwnd, "^!t", "take_screenshot")
        DesktopAssert(IsObject(hotkeyState) && hotkeyState.Has("desktop"), "Keyboard shortcut capture uses modern desktop shell")
        hc := hotkeyState["controls"]
        DesktopAssert(hc["capture"].Value = "^!t" && hc["editor"].Text = "Ctrl + Alt + T"
            && hc["heading"].Text != "Selected action",
            "Keyboard shortcut capture preserves the binding and action label")
        for size in [[680, 500], [760, 560], [1040, 700]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(hotkeyState, w, h)
            DesktopDialogAssertChrome(hotkeyState, w, h)
            hc["editor"].GetPos(&x, &y, &cw, &ch)
            DesktopAssert(cw = w - 80 && ch = 40 && y + ch < h - 120,
                "Keyboard shortcut field remains responsive and clear of footer actions")
            DesktopAssert(!(WinGetExStyle(hc["editor"].Hwnd) & 0x200)
                && !(WinGetStyle(hc["editor"].Hwnd) & 0x800000),
                "Keyboard shortcut field has no classic white frame")
            DesktopAssert(DesktopTopChildAtCenter(hotkeyState["gui"].Hwnd, hc["editor"]) = hc["editor"].Hwnd,
                "The themed shortcut display stays above the native capture field before typing")
            if w = 680 {
                SendMessage(0xF5, 0, 0, hc["editor"].Hwnd) ; BM_CLICK
                Sleep(10)
                DesktopAssert(StudyControllerFocusedHwnd(hotkeyState["gui"].Hwnd) = hc["capture"].Hwnd,
                    "Clicking the modern shortcut display transfers focus to native key capture")
            }
            if dark && w = 760 {
                dpi := GetWindowDPI(hotkeyState["gui"].Hwnd) / 96
                TestDesktopStudyCapture(hotkeyState["gui"], "controls-keyboard-capture.png",
                    Round(w * dpi), Round(h * dpi))
            }
        }
        hc["capture"].Value := "^+x"
        CPHotkeyDialogSync(hotkeyState)
        savedValue := hc["capture"].Value
        CPHotkeyDialogClose(hotkeyState, savedValue)
        DesktopAssert(hotkeyState["closed"] && hotkeyState["result"] = savedValue,
            "Keyboard shortcut Save returns the exact captured binding")

        emptyHotkey := CPHotkeyDialogCreate(reader["gui"].Hwnd, "", "take_screenshot")
        DesktopDialogFixtureShow(emptyHotkey, 760, 560)
        emptyControls := emptyHotkey["controls"]
        DesktopAssert(emptyControls["capture"].Value = "" && emptyControls["editor"].Text = "None",
            "An empty shortcut has a clear themed initial value")
        emptyControls["capture"].Focus(), CPHotkeyDialogPresentEditor(emptyHotkey), Sleep(10)
        DesktopAssert(StudyControllerFocusedHwnd(emptyHotkey["gui"].Hwnd) = emptyControls["capture"].Hwnd
            && DesktopTopChildAtCenter(emptyHotkey["gui"].Hwnd, emptyControls["editor"]) = emptyControls["editor"].Hwnd,
            "The native capture retains focus while its dark display remains visible")
        if dark {
            dpi := GetWindowDPI(emptyHotkey["gui"].Hwnd) / 96
            TestDesktopStudyCapture(emptyHotkey["gui"], "controls-keyboard-capture-empty.png",
                Round(760 * dpi), Round(560 * dpi))
        }
        CPHotkeyDialogClose(emptyHotkey)

        controller := CPControllerCaptureDialogCreate(reader["gui"].Hwnd, "take_screenshot")
        DesktopAssert(IsObject(controller) && controller.Has("desktop"),
            "Controller assignment uses modern desktop shell")
        cc := controller["controls"]
        for size in [[680, 500], [760, 560], [1040, 700]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(controller, w, h)
            DesktopDialogAssertChrome(controller, w, h)
            cc["status"].GetPos(&x, &y, &cw, &ch)
            DesktopAssert(cw = w - 80 && ch >= 32 && y + ch < h - 90,
                "Controller capture status remains responsive and clear of the footer")
            if dark && w = 760 {
                cc["status"].Value := "Ready. Press one controller button."
                dpi := GetWindowDPI(controller["gui"].Hwnd) / 96
                TestDesktopStudyCapture(controller["gui"], "controls-controller-capture.png",
                    Round(w * dpi), Round(h * dpi))
            }
        }
        CPControllerCaptureDialogClose(controller)
        controller["gui"].Destroy()
        DesktopAssert(controller["closed"], "Controller assignment Cancel keeps the existing binding")
    }

    controlDarkMode := 1
    for title in ["Shortcut conflicts", "Controller assignment"] {
        TestControlMessageCaptured := false, TestControlMessageTitle := title
        SetTimer(TestControlMessageClose, 80)
        try accepted := title = "Shortcut conflicts"
            ? CPConfirmDuplicateHotkeys(reader["gui"].Hwnd)
            : CPConfirmControllerBindingMove(reader["gui"].Hwnd,
                "X:A", "take_screenshot", "screenshot_translate")
        finally SetTimer(TestControlMessageClose, 0)
        DesktopAssert(TestControlMessageCaptured && !accepted,
            title " uses the modern confirmation with No as the safe default")
    }

    fullscreen := Gui("+Owner" reader["gui"].Hwnd " -Caption -DPIScale")
    fullscreen.CPDialogPresentation := "fullscreen"
    DesktopAssert(!CPHotkeyDialogCreate(fullscreen.Hwnd, "^!t", "take_screenshot")
        && !CPControllerCaptureDialogCreate(fullscreen.Hwnd, "take_screenshot"),
        "Fullscreen retains its established in-page keyboard and controller binding workflow")
    fullscreen.Destroy()
}

TestControlMessageClose() {
    global TestControlMessageCaptured, TestControlMessageTitle
    for hwnd, d in StudyDesktopRegistry() {
        if d["kind"] != "message" || d["state"]["gui"].Title != TestControlMessageTitle || d["width"] < 1
            continue
        TestControlMessageCaptured := true
        s := d["state"]
        CPThemedDialogFinish(s, s["gui"], "No")
        return
    }
}

TestAppearanceDialogs(reader) {
    global controlDarkMode, chkDarkMode, chkTop, controlPanelOpacity
    global CPControllerColorGradientSliders
    global TestAppearanceColorCaptured, CPControllerColorDialogState, ui
    originalDark := controlDarkMode, originalDarkValue := chkDarkMode.Value
    originalTop := chkTop.Value, originalOpacity := controlPanelOpacity
    try {
        chkTop.Value := 0, controlPanelOpacity := 92
        for dark in [1, 0] {
            controlDarkMode := dark, chkDarkMode.Value := dark
            appearance := CPDesktopAppearanceCreate(reader["gui"].Hwnd)
            DesktopAssert(IsObject(appearance) && appearance.Has("desktop")
                && appearance["desktop"]["kind"] = "appearance",
                "Desktop window preferences use the modern dialog shell")
            ac := appearance["controls"]
            DesktopAssert(ac["dark"].Value = dark && !ac["top"].Value
                && ac["opacity"].Value = 92 && InStr(ac["opacityLabel"].Text, "92%"),
                "Window preferences preserve current theme, topmost and opacity values")
            for size in [[680, 600], [760, 640], [1040, 760]] {
                w := size[1], h := size[2]
                DesktopDialogFixtureShow(appearance, w, h)
                DesktopDialogAssertChrome(appearance, w, h)
                ac["opacity"].GetPos(&x, &y, &cw, &ch)
                ac["about"].GetPos(&aboutX, &aboutY, &aboutW, &aboutH)
                ac["done"].GetPos(&doneX, &doneY, &doneW, &doneH)
                DesktopAssert(x = 40 && cw = w - 80 && y + ch < h - 66,
                    "Window opacity slider remains responsive and above the footer")
                DesktopAssert(ac["about"].Text = "About JRPG Translator…"
                    && aboutX = 24 && aboutY = doneY && aboutX + aboutW < doneX,
                    "About action remains left-aligned and separate from Done")
                DesktopAssert(doneX + doneW = w - 24 && doneY + doneH <= h - 16,
                    "Window preference action remains aligned in the footer")
                if dark && w = 760 {
                    dpi := GetWindowDPI(appearance["gui"].Hwnd) / 96
                    TestDesktopStudyCapture(appearance["gui"], "appearance-window-options.png",
                        Round(w * dpi), Round(h * dpi))
                }
            }
            hwnd := appearance["gui"].Hwnd
            CPDesktopAppearanceClose(appearance)
            DesktopAssert(appearance["closed"] && !DllCall("user32\IsWindow", "ptr", hwnd)
                && !IsObject(StudyDesktopContext(hwnd)),
                "Window preferences close and unregister cleanly")

            color := CPControllerColorDialogCreate(reader["gui"].Hwnd, "2563EB",
                "Adjust Translator window color")
            DesktopAssert(IsObject(color) && color.Has("desktop")
                && color["desktop"]["kind"] = "colorAdjust",
                "Overlay color adjustment uses the modern dialog shell")
            cc := color["controls"]
            DesktopAssert(CPIsColorSwatchControl(cc["preview"].Hwnd)
                && CPControllerColorGradientSliders.Has(cc["hue"].Hwnd)
                && CPControllerColorGradientSliders.Has(cc["saturation"].Hwnd)
                && CPControllerColorGradientSliders.Has(cc["brightness"].Hwnd),
                "Color preview and all three channel gradients remain registered")
            for size in [[720, 680], [800, 700], [1100, 820]] {
                w := size[1], h := size[2]
                DesktopDialogFixtureShow(color, w, h)
                DesktopDialogAssertChrome(color, w, h)
                cc["preview"].GetPos(&px, &py, &pw, &ph)
                cc["brightness"].GetPos(&sx, &sy, &channelW, &sh)
                DesktopAssert(px = 40 && pw = w - 80 && ph = 62,
                    "Color preview expands with the dialog")
                DesktopAssert(sx = 166 && channelW >= 220 && sy + sh < h - 120,
                    "Color sliders stay usable and clear of guidance and footer actions")
                if dark && w = 800 {
                    dpi := GetWindowDPI(color["gui"].Hwnd) / 96
                    TestDesktopStudyCapture(color["gui"], "appearance-color-adjustment.png",
                        Round(w * dpi), Round(h * dpi))
                }
            }
            cc["hue"].Value := 180, cc["saturation"].Value := 100
            cc["brightness"].Value := 50
            DesktopAssert(color["updatePreview"].Call() = "008080"
                && cc["hueValue"].Text = "180" && cc["saturationValue"].Text = "100%"
                && cc["brightnessValue"].Text = "50%",
                "HSV edits update the exact preview color and readable values")
            cc["hue"].Focus(), CPControllerColorNavigate("Down", cc["hue"],
                cc["saturation"], cc["brightness"], cc["apply"], cc["cancel"])
            DesktopAssert(StudyControllerFocusedHwnd(color["gui"].Hwnd) = cc["saturation"].Hwnd,
                "Color controller navigation moves through channel sliders")
            previewHwnd := cc["preview"].Hwnd, gradientHwnds := color["gradientHwnds"].Clone()
            CPControllerColorDialogClose(color, "008080")
            DesktopAssert(color["closed"] && color["result"] = "008080"
                && !CPIsColorSwatchControl(previewHwnd),
                "Applying a color returns its exact RGB value and unregisters the preview")
            for sliderHwnd in gradientHwnds
                DesktopAssert(!CPControllerColorGradientSliders.Has(sliderHwnd),
                    "Closing color adjustment unregisters a channel gradient")
        }

        controlDarkMode := 1, chkDarkMode.Value := 1
        TestAppearanceColorCaptured := false
        expectedColor := CPColorHSVToHex(32, 75, 80)
        SetTimer(TestAppearanceColorApply, 80)
        try selectedColor := CPControllerColorDialog("2563EB", "Adjust Explainer text color")
        finally SetTimer(TestAppearanceColorApply, 0)
        DesktopAssert(TestAppearanceColorCaptured && selectedColor = expectedColor
            && !CPControllerColorDialogState.Get("active", false)
            && DllCall("user32\IsWindowEnabled", "ptr", ui.Hwnd),
            "Color dialog run returns Apply result, clears controller state, and restores its owner")

        fullscreen := Gui("+Owner" reader["gui"].Hwnd " -Caption -DPIScale")
        fullscreen.CPDialogPresentation := "fullscreen"
        DesktopAssert(!CPDesktopAppearanceCreate(fullscreen.Hwnd)
            && !CPControllerColorDialogCreate(fullscreen.Hwnd, "2563EB"),
            "Fullscreen retains its established in-page appearance and color controls")
        fullscreen.Destroy()
    } finally {
        controlDarkMode := originalDark, chkDarkMode.Value := originalDarkValue
        chkTop.Value := originalTop, controlPanelOpacity := originalOpacity
    }
}

TestAppearanceColorApply() {
    global TestAppearanceColorCaptured
    for hwnd, d in StudyDesktopRegistry() {
        if d["kind"] != "colorAdjust" || d["width"] < 1
            || !d["state"].Get("ready", false)
            continue
        TestAppearanceColorCaptured := true
        c := d["state"]["controls"]
        c["hue"].Value := 32, c["saturation"].Value := 75, c["brightness"].Value := 80
        d["state"]["updatePreview"].Call()
        SendMessage(0xF5, 0, 0, c["apply"].Hwnd)
        return
    }
}

TestHelpDialogs(reader) {
    global controlDarkMode, APP_VERSION, PROJECT_URL, BUG_REPORT_URL
    global BEGINNER_VIDEO_URL, WRITTEN_GUIDE_URL
    APP_VERSION := "0.9.9.0"
    PROJECT_URL := "https://example.invalid/jrpg-translator"
    BUG_REPORT_URL := PROJECT_URL "/issues/new"
    BEGINNER_VIDEO_URL := "https://example.invalid/beginner"
    WRITTEN_GUIDE_URL := PROJECT_URL "#quick-start"

    for dark in [1, 0] {
        controlDarkMode := dark
        welcome := CPDesktopWelcomeDialogCreate(reader["gui"].Hwnd, false)
        DesktopAssert(IsObject(welcome) && welcome.Has("desktop")
            && welcome["desktop"]["kind"] = "welcome",
            "Welcome guide uses the modern desktop dialog shell")
        wc := welcome["controls"]
        DesktopAssert(InStr(wc["steps"].Text, "1. Add at least one API key.")
            && InStr(wc["steps"].Text, "5. Save the finished setup")
            && wc["dontShow"].Value = 1,
            "Welcome guide preserves the complete setup checklist and default preference")
        for size in [[780, 700], [880, 720], [1200, 860]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(welcome, w, h)
            DesktopDialogAssertChrome(welcome, w, h)
            wc["steps"].GetPos(&x, &y, &cw, &ch)
            wc["continue"].GetPos(&continueX, &continueY, &continueW, &continueH)
            dpi := GetWindowDPI(welcome["gui"].Hwnd) / 96
            DesktopAssert(x = 40 && cw = w - 80
                && CPMeasureWrappedTextHeight(wc["steps"], cw * dpi) <= ch * dpi,
                "Welcome checklist remains readable at each responsive size")
            DesktopAssert(continueX + continueW = w - 24
                && continueY + continueH <= h - 16,
                "Welcome Continue action remains aligned in the footer")
            if dark && w = 880
                TestDesktopStudyCapture(welcome["gui"], "help-welcome-guide.png",
                    Round(w * dpi), Round(h * dpi))
        }
        hwnd := welcome["gui"].Hwnd
        CPDesktopWelcomeClose(welcome, false)
        DesktopAssert(welcome["closed"] && !DllCall("user32\IsWindow", "ptr", hwnd)
            && !IsObject(StudyDesktopContext(hwnd)),
            "Welcome guide closes and unregisters cleanly without changing test preferences")

        about := CPDesktopAboutDialogCreate(reader["gui"].Hwnd)
        DesktopAssert(IsObject(about) && about.Has("desktop")
            && about["desktop"]["kind"] = "about",
            "About window uses the modern desktop dialog shell")
        ac := about["controls"]
        DesktopAssert(ac["version"].Text = "Version 0.9.9.0"
            && InStr(AboutVersionInfo(), PROJECT_URL)
            && ac["creator"].Text = "Created by retrogamer0815",
            "About window and copied diagnostics retain version and project information")
        for size in [[720, 640], [800, 660], [1100, 800]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(about, w, h)
            DesktopDialogAssertChrome(about, w, h)
            ac["report"].GetPos(&reportX, &reportY, &reportW, &reportH)
            ac["github"].GetPos(&githubX, &githubY, &githubW, &githubH)
            DesktopAssert(reportW = githubW && reportX = 40
                && githubX > reportX + reportW && reportY = githubY,
                "About actions keep a balanced responsive two-column layout")
            if dark && w = 800 {
                dpi := GetWindowDPI(about["gui"].Hwnd) / 96
                TestDesktopStudyCapture(about["gui"], "help-about.png",
                    Round(w * dpi), Round(h * dpi))
            }
        }
        hwnd := about["gui"].Hwnd
        CPDesktopAboutClose(about)
        DesktopAssert(about["closed"] && !DllCall("user32\IsWindow", "ptr", hwnd)
            && !IsObject(StudyDesktopContext(hwnd)),
            "About window closes and unregisters cleanly")
    }

    fullscreen := Gui("+Owner" reader["gui"].Hwnd " -Caption -DPIScale")
    fullscreen.CPDialogPresentation := "fullscreen"
    DesktopAssert(!CPDesktopWelcomeDialogCreate(fullscreen.Hwnd)
        && !CPDesktopAboutDialogCreate(fullscreen.Hwnd),
        "Fullscreen keeps its established in-page About and setup guidance")
    fullscreen.Destroy()
}

TestModelDialogs(reader) {
    global controlDarkMode, ui, TestModelRefreshState, TestModelQueryMode, TestModelNotice
    catalog := TestModelCatalogFixture(), existing := ["GEMINI-EXISTING"]
    for dark in [1, 0] {
        controlDarkMode := dark
        for mode in ["source", "browser"] {
            s := CPModelDialogCreate(reader["gui"].Hwnd, "gemini", "screenshot", mode, catalog, existing)
            c := s["controls"], g := s["gui"]
            for size in (mode = "source" ? [[760, 560], [800, 580], [1080, 680]] : [[760, 720], [900, 820], [1200, 900]]) {
                w := size[1], h := size[2]
                DesktopDialogFixtureShow(s, w, h)
                DesktopDialogAssertChrome(s, w, h)
                if mode = "browser"
                    CPModelDialogListMetrics(s)
                dpi := GetWindowDPI(g.Hwnd) / 96
                for key, ctrl in c {
                    ctrl.GetPos(&x, &y, &cw, &ch)
                    DesktopAssert(x >= 24 && x + cw <= w - 24 && y >= 174, "Model dialog control fits: " key)
                    DesktopAssert(y + ch <= (s["desktop"]["surfaces"][ctrl.Hwnd] = "bar" ? h - 16 : h - 86), "Model content stays above footer: " key)
                    if ctrl.Type = "Text"
                        DesktopAssert(CPMeasureWrappedTextHeight(ctrl, cw * dpi) <= ch * dpi, "Model dialog help fits: " key)
                    if ctrl.Type = "Edit" || ctrl.Type = "ListBox"
                        DesktopAssert(!(WinGetExStyle(ctrl.Hwnd) & 0x200) && !(WinGetStyle(ctrl.Hwnd) & 0x800000), "Model browser has no native white edge")
                }
                if dark
                    TestDesktopStudyCapture(g, "model-" mode "-desktop-" w ".png", Round(w * dpi), Round(h * dpi))
            }
            if mode = "source" {
                DesktopAssert(s["source"] = "online", "Model source starts with online selected")
                DesktopAssert(s["desktop"]["paint"][c["online"].Hwnd]["kind"] = "segment", "Source options use a renderer that paints the selected state")
                CPModelDialogSelectSource(s, "manual")
                DesktopAssert(s["source"] = "manual" && !s["closed"], "Source selection alone performs no action")
                DesktopAssert(s["desktop"]["paint"][c["manual"].Hwnd]["selected"] && !s["desktop"]["paint"][c["online"].Hwnd]["selected"], "Only selected model source stays highlighted")
                StudyControllerSetFocus(g.Hwnd, c["online"].Hwnd)
                CPModelDialogNavigate(s, "Down")
                DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["manual"].Hwnd, "Source navigation reaches manual option")
                CPModelDialogNavigate(s, "Down")
                DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["save"].Hwnd, "Source navigation reaches Continue")
                CPModelDialogActivate(s)
                Sleep(25)
                DesktopAssert(s["closed"] && s["result"] = "manual", "Continue returns selected source")
            } else {
                DesktopAssert(SendMessage(0x18B, 0, 0, c["list"].Hwnd) = 3 && c["list"].Text = "model-alpha", "Browser excludes already-added models case-insensitively")
                DesktopAssert(SendMessage(0x193, 0, 0, c["list"].Hwnd) >= StudyLibraryMeasureListText(c["list"].Hwnd, catalog["models"][4]), "Long model ID can be scrolled horizontally")
                c["list"].Choose(3), StudyControllerSetFocus(g.Hwnd, c["list"].Hwnd)
                CPModelDialogNavigate(s, "Down")
                DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["save"].Hwnd, "End of model list navigates to Add model")
                CPModelDialogNavigate(s, "Left")
                DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["refresh"].Hwnd, "Browser footer Left follows visual order")
                CPModelDialogNavigate(s, "Right")
                DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["save"].Hwnd, "Browser footer Right follows visual order")
                CPModelDialogActivate(s)
                DesktopAssert(s["closed"] && s["result"] = catalog["models"][4], "Adding returns exact selected model ID")
            }
        }
    }
    controlDarkMode := 1
    purposeLabels := Map("screenshot", "Game Text Translation", "audio", "Audio Translation", "explanation", "Explanation")
    for source in ["online", "manual"] {
        s := CPModelDialogCreate(reader["gui"].Hwnd, "openai", "audio", "source")
        DesktopDialogFixtureShow(s, 800, 580)
        SendMessage(0xF5, 0, 0, s["controls"][source].Hwnd)
        Sleep(25)
        DesktopAssert(s["closed"] && s["result"] = source,
            "Clicking the " source " source card proceeds immediately")
    }
    for provider in ["gemini", "openai"] {
        for purpose in ["screenshot", "audio", "explanation"] {
            s := CPModelDialogCreate(reader["gui"].Hwnd, provider, purpose, "source")
            DesktopAssert(InStr(StrLower(s["gui"].Title), provider), "Model source identifies provider")
            DesktopAssert(InStr(StrLower(s["desktop"]["chrome"]["subtitle"].Text), StrLower(purposeLabels[purpose])), "Model source identifies purpose")
            CPModelDialogClose(s, "cancel")
        }
    }
    s := CPModelDialogCreate(reader["gui"].Hwnd, "openai", "audio", "browser", catalog, existing)
    c := s["controls"], g := s["gui"]
    DesktopDialogFixtureShow(s, 900, 820)
    for source in ["cache", "stale_cache"] {
        cached := TestModelCatalogFixture(source, ["cached-model"], ["Synthetic warning. " . "A readable cached list remains available."])
        ModelPickerPopulate(c["list"], c["status"], c["save"], cached, [])
        DesktopAssert(InStr(c["status"].Text, "cache") && InStr(c["status"].Text, "Synthetic warning"), "Cache source and warning remain visible")
    }
    ModelPickerPopulate(c["list"], c["status"], c["save"], catalog, catalog["models"])
    DesktopAssert(!c["list"].Enabled && !c["save"].Enabled && InStr(c["status"].Text, "already"), "All-added catalog cannot add duplicates")
    ModelPickerPopulate(c["list"], c["status"], c["save"], TestModelCatalogFixture("online", []), [])
    DesktopAssert(!c["list"].Enabled && InStr(c["status"].Text, "No compatible"), "Empty catalog has accurate empty state")
    StudyControllerSetFocus(g.Hwnd, c["refresh"].Hwnd)
    CPModelDialogNavigate(s, "Left")
    DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["cancel"].Hwnd, "Empty catalog navigation skips disabled Add")
    CPModelDialogNavigate(s, "Right")
    DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["refresh"].Hwnd, "Empty catalog navigation returns to Refresh")
    TestModelRefreshState := s, TestModelQueryMode := "failure", TestModelNotice := ""
    TestModelPickerRefresh("openai", "audio", existing, g, c["list"], c["status"], c["save"], c["refresh"], s)
    DesktopAssert(!c["list"].Enabled && !c["save"].Enabled && c["refresh"].Enabled && !s["refreshing"], "Failed refresh preserves empty disabled state")
    DesktopAssert(InStr(TestModelNotice, "Synthetic failure"), "Refresh failure is reported through shared notice")
    TestModelQueryMode := "success"
    TestModelPickerRefresh("openai", "audio", existing, g, c["list"], c["status"], c["save"], c["refresh"], s)
    DesktopAssert(c["list"].Enabled && c["list"].Text = "model-alpha" && !s["refreshing"], "Refresh replaces list using the existing duplicate filter")
    TestModelQueryMode := "close"
    TestModelPickerRefresh("openai", "audio", existing, g, c["list"], c["status"], c["save"], c["refresh"], s)
    DesktopAssert(s["closed"] && s["result"] = "" && !s["refreshing"], "Closing during refresh ignores late result safely")
    owner := Gui("+Owner" ui.Hwnd " -Caption -DPIScale"), owner.CPDialogPresentation := "fullscreen"
    for mode in ["source", "browser"] {
        for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
            s := CPModelDialogCreate(owner.Hwnd, "openai", "audio", mode, catalog, existing)
            g := s["gui"], f := s["bigBoxForm"], w := size[1], h := size[2]
            g.Show("Hide w" w " h" h)
            StudyCandidatesRecommendationBigBoxResize(f, g, 0, w, h)
            StudyCandidatesRecommendationBigBoxApplyTheme(f)
            if mode = "browser" {
                CPModelDialogListMetrics(s)
                modelList := s["controls"]["list"]
                expectedExtent := StudyLibraryMeasureListText(modelList.Hwnd, catalog["models"][4]) + 24
                DesktopAssert(SendMessage(0x193, 0, 0, modelList.Hwnd) = expectedExtent, "Fullscreen horizontal range uses final scaled font")
            }
            g.Show("NA x-12000 y-12000")
            Sleep(35)
            f["shell"]["footer"].GetPos(, &footerY)
            for key, ctrl in s["controls"] {
                ctrl.GetPos(&x, &y, &cw, &ch)
                DesktopAssert(x >= 0 && y > 0 && x + cw <= w && y + ch < footerY, "Fullscreen model dialog fits: " key)
                if ctrl.Type = "Text" || ctrl.Type = "Button"
                    DesktopAssert(CPMeasureWrappedTextHeight(ctrl, cw - 16) <= ch, "Fullscreen model label fits: " key)
            }
            TestDesktopStudyCapture(g, "model-" mode "-fullscreen-" w ".png", w, h)
            if mode = "browser" && w = 1280 {
                longId := catalog["models"][4] . catalog["models"][4]
                ModelPickerPopulate(modelList, s["controls"]["status"], s["controls"]["save"], TestModelCatalogFixture("online", [longId]), [])
                DesktopAssert(SendMessage(0x193, 0, 0, modelList.Hwnd) = StudyLibraryMeasureListText(modelList.Hwnd, longId) + 24, "Refreshed long IDs use the current fullscreen font")
                StudyControllerSetFocus(g.Hwnd, modelList.Hwnd)
                CPModelDialogNavigate(s, "Right")
                DesktopAssert(DllCall("user32\GetScrollPos", "ptr", modelList.Hwnd, "int", 0) > 0, "Browser arrows scroll long IDs without changing selection")
                TestDesktopStudyCapture(g, "model-browser-fullscreen-long-id.png", w, h)
            }
            CPModelDialogClose(s, mode = "source" ? "cancel" : "")
        }
    }
    for hwnd in [reader["gui"].Hwnd, owner.Hwnd] {
        for mode in ["source", "browser"] {
            actions := mode = "source" ? ["accept", "direct-online", "direct-manual", "cancel", "escape", "close"]
                : ["accept", "cancel", "escape", "close"]
            for action in actions
                TestModelDialogModal(hwnd, mode, action)
        }
    }
    owner.Destroy()
}

TestModelCatalogQuery(provider, purpose, forceRefresh) {
    global TestModelRefreshState, TestModelQueryMode
    c := TestModelRefreshState["controls"]
    DesktopAssert(!c["list"].Enabled && !c["save"].Enabled && !c["refresh"].Enabled, "Refresh disables actions during query")
    if TestModelQueryMode = "close"
        CPModelDialogClose(TestModelRefreshState)
    if TestModelQueryMode = "failure"
        return Map("ok", false, "error", "Synthetic failure")
    return TestModelCatalogFixture()
}

TestModelRefreshNotice(owner, message, title) {
    global TestModelNotice
    TestModelNotice := message
}

TestModelDialogModal(owner, mode, action) {
    s := CPModelDialogCreate(owner, "gemini", "screenshot", mode, TestModelCatalogFixture(), [])
    wasEnabled := DllCall("user32\IsWindowEnabled", "ptr", owner)
    priorFocus := StudyControllerFocusedHwnd(owner)
    timer := TestModelDialogModalFinish.Bind(s, action)
    SetTimer(timer, 60)
    try result := CPModelDialogRun(s)
    finally SetTimer(timer, 0)
    expected := action = "direct-online" ? "online" : action = "direct-manual" ? "manual"
        : action = "accept" ? mode = "source" ? "manual" : "model-alpha"
        : mode = "source" ? "cancel" : ""
    DesktopAssert(result = expected, "Model modal result matches action")
    DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", owner) = wasEnabled, "Model modal restores parent enabled state")
    if wasEnabled && priorFocus && CPHwndIsFocusable(priorFocus)
        DesktopAssert(StudyControllerFocusedHwnd(owner) = priorFocus, "Model modal restores parent focus")
}

TestModelDialogModalFinish(s, action) {
    if s["closed"] || !s.Get("ready", false)
        return
    DesktopAssert(!DllCall("user32\IsWindowEnabled", "ptr", s["owner"]), "Model modal disables owner while open")
    c := s["controls"]
    switch action {
        case "accept":
            if s["mode"] = "source"
                CPModelDialogSelectSource(s, "manual")
            else
                c["list"].Choose(2)
            SendMessage(0xF5, 0, 0, c["save"].Hwnd)
        case "direct-online": SendMessage(0xF5, 0, 0, c["online"].Hwnd)
        case "direct-manual": SendMessage(0xF5, 0, 0, c["manual"].Hwnd)
        case "cancel": SendMessage(0xF5, 0, 0, c["cancel"].Hwnd)
        case "escape": PostMessage(0x100, 0x1B, 0, s["gui"].Hwnd)
        case "close": PostMessage(0x10, 0, 0, s["gui"].Hwnd)
    }
}

TestNamingDialogs(reader) {
    global controlDarkMode, CPDesktop, ui
    prompts := [
        ["Create Profile", "Enter a name for the new profile:", ""],
        ["New EXPLAIN prompt", "Enter a name for the new EXPLANATION prompt:", ""],
        ["New prompt", "Enter a name for the new prompt profile:", "To show a transcript in the Translation overlay, include`nwith_transcript or with_kanji_reading in the name."],
        ["New local corrections profile", "Enter a name for the new local corrections profile:", ""],
        ["New model instructions profile", "Enter a name for the new model instructions profile:", ""],
        ["Add model", "Model ID", "Enter the exact model ID supplied by the provider."]]
    for dark in [1, 0] {
        controlDarkMode := dark
        for index, spec in prompts {
            s := CPInputDialogCreate(reader["gui"].Hwnd, spec[2], spec[1], spec[3], "日本語 & name")
            for size in [[680, 500], [760, 520], [1000, 650]] {
                w := size[1], h := size[2]
                DesktopDialogFixtureShow(s, w, h)
                DesktopDialogAssertChrome(s, w, h)
                c := s["controls"], g := s["gui"], dpi := GetWindowDPI(g.Hwnd) / 96
                for key in ["prompt", "info"] {
                    c[key].GetPos(,, &cw, &ch)
                    DesktopAssert(CPMeasureWrappedTextHeight(c[key], cw * dpi) <= ch * dpi, "Name prompt/help fits: " spec[1])
                }
                DesktopAssert(c["editor"].Value = "日本語 & name", "Naming layout preserves literal text")
                DesktopAssert(!(WinGetExStyle(c["editor"].Hwnd) & 0x200) && !(WinGetStyle(c["editor"].Hwnd) & 0x800000), "Name field has no classic white frame")
                if dark && w = 760
                    TestDesktopStudyCapture(g, "naming-desktop-" index ".png", Round(w * dpi), Round(h * dpi))
            }
            hwnd := s["gui"].Hwnd
            s["controls"]["editor"].Value := "  Edited 日本語 & name  "
            CPInputDialogClose(s, true)
            DesktopAssert(s["result"].Result = "OK" && s["result"].Value = "  Edited 日本語 & name  ", "Name entry returns unmodified value for existing caller validation")
            DesktopAssert(!IsObject(StudyDesktopContext(hwnd)), "Name dialog clears paint registry")
        }
    }
    controlDarkMode := 1
    owner := Gui("+Owner" ui.Hwnd " -Caption -DPIScale")
    owner.CPDialogPresentation := "fullscreen"
    for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
        s := CPInputDialogCreate(owner.Hwnd, prompts[3][2], prompts[3][1], prompts[3][3], "with_kanji_reading_日本語")
        g := s["gui"], f := s["bigBoxForm"], c := s["controls"], w := size[1], h := size[2]
        g.Show("Hide w" w " h" h)
        StudyCandidatesRecommendationBigBoxResize(f, g, 0, w, h)
        StudyCandidatesRecommendationBigBoxApplyTheme(f)
        g.Show("NA x-12000 y-12000")
        Sleep(35)
        TestDesktopStudyCapture(g, "naming-fullscreen-" w ".png", w, h)
        f["shell"]["footer"].GetPos(, &footerY)
        for key, ctrl in c {
            ctrl.GetPos(&x, &y, &cw, &ch)
            DesktopAssert(x >= 0 && y > 0 && x + cw <= w && y + ch < footerY, "Fullscreen naming control fits: " key)
            if ctrl.Type = "Text"
                DesktopAssert(CPMeasureWrappedTextHeight(ctrl, cw) <= ch, "Fullscreen naming text fits: " key)
            point := Buffer(8)
            NumPut("int", x + Floor(cw / 2), "int", y + Floor(ch / 2), point)
            hit := DllCall("user32\ChildWindowFromPointEx", "ptr", g.Hwnd, "int64", NumGet(point, 0, "int64"), "uint", 3, "ptr")
            DesktopAssert(hit = ctrl.Hwnd, "Fullscreen naming control remains reachable above disabled decoration: " key)
            FileAppend("naming-fullscreen-" w ".png|" key "|" (x + 8) "|" (y + 8) "|" (cw - 16) "|" (ch - 16) "`n", A_ScriptDir "\naming-visibility.txt")
        }
        StudyControllerSetFocus(g.Hwnd, c["editor"].Hwnd)
        StudyControllerMoveFocus(g.Hwnd, "Down")
        DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["save"].Hwnd, "Fullscreen naming navigation reaches OK")
        StudyControllerMoveFocus(g.Hwnd, "Right")
        DesktopAssert(StudyControllerFocusedHwnd(g.Hwnd) = c["cancel"].Hwnd, "Fullscreen naming navigation reaches Cancel")
        TestDesktopStudyCapture(g, "naming-fullscreen-" w ".png", w, h)
        CPInputDialogClose(s)
        DesktopAssert(s["result"].Result = "Cancel" && s["result"].Value = "", "Fullscreen cancellation discards the draft")
    }
    for hwnd in [reader["gui"].Hwnd, owner.Hwnd] {
        for action in ["ok", "cancel", "escape", "close"]
            TestNamingDialogRoute(hwnd, action)
        DllCall("user32\EnableWindow", "ptr", hwnd, "int", 0)
        TestNamingDialogRoute(hwnd, "cancel")
        DllCall("user32\EnableWindow", "ptr", hwnd, "int", 1)
    }
    owner.Destroy()
}

TestNamingDialogRoute(owner, action) {
    global TestNamingCaptured
    TestNamingCaptured := false
    wasEnabled := DllCall("user32\IsWindowEnabled", "ptr", owner)
    wasFocused := StudyControllerFocusedHwnd(owner)
    timer := TestNamingDialogFinish.Bind(action)
    SetTimer(timer, 80)
    try result := CPThemedInputBox("Synthetic name", "Naming dialog test", "No files are written.", "Initial value", 420, owner)
    finally SetTimer(timer, 0)
    DesktopAssert(TestNamingCaptured, "Shared input follows modern owner")
    DesktopAssert(result.Result = (action = "ok" ? "OK" : "Cancel"), "Name result matches action: " action)
    DesktopAssert(result.Value = (action = "ok" ? "  Test 日本語 & name  " : ""), "Name cancellation never leaks edited value")
    DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", owner) = wasEnabled, "Name dialog restores prior parent enabled state")
    if wasEnabled && wasFocused && CPHwndIsFocusable(wasFocused)
        DesktopAssert(StudyControllerFocusedHwnd(owner) = wasFocused, "Name dialog restores parent focus")
}

TestNamingDialogFinish(action) {
    global TestNamingCaptured
    for hwnd in WinGetList("Naming dialog test ahk_class AutoHotkeyGUI") {
        g := GuiFromHwnd(hwnd)
        if !g.HasOwnProp("CPInputDialog") || !g.CPInputDialog.Get("ready", false)
            continue
        s := g.CPInputDialog
        DesktopAssert(!DllCall("user32\IsWindowEnabled", "ptr", s["owner"]), "Name dialog blocks parent while open")
        TestNamingCaptured := true
        s["controls"]["editor"].Value := "  Test 日本語 & name  "
        switch action {
            case "ok": SendMessage(0xF5, 0, 0, s["controls"]["save"].Hwnd)
            case "cancel": SendMessage(0xF5, 0, 0, s["controls"]["cancel"].Hwnd)
            case "escape": PostMessage(0x100, 0x1B, 0, g.Hwnd)
            case "close": PostMessage(0x10, 0, 0, g.Hwnd)
        }
        return
    }
}

TestSharedMessageRoute(owner, presentation, buttons, icon, action := "default") {
    global TestSharedMessageCaptured
    TestSharedMessageCaptured := false
    priorEnabled := DllCall("user32\IsWindowEnabled", "ptr", owner)
    timer := TestSharedMessageClose.Bind(presentation, buttons, action)
    SetTimer(timer, 80)
    try {
        if buttons = "yesnocancel"
            result := CPAdaptiveOwnedMessage(owner,
                "The current settings contain unsaved edits. 日本語 & names remain literal.",
                "Shared message test", buttons, icon, 620,
                "Save changes", "Discard changes", "Cancel")
        else
            result := CPAdaptiveOwnedMessage(owner,
                "Synthetic shared notice. 日本語 & names remain literal.",
                "Shared message test", buttons, icon)
    }
    finally SetTimer(timer, 0)
    DesktopAssert(TestSharedMessageCaptured, "Shared message used expected presentation: " presentation)
    expected := action = "yes" ? "Yes" : action = "no" ? "No"
        : action = "cancel" ? "Cancel" : CPDialogDefaultResult(buttons)
    DesktopAssert(result = expected, "Shared message returns the selected action: " buttons " / " action)
    DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", owner) = priorEnabled, "Shared message restores owner enabled state")
}

TestSharedMessageClose(presentation, buttons, action := "default") {
    global TestSharedMessageCaptured
    if presentation = "desktop" {
        for hwnd, d in StudyDesktopRegistry() {
            if d["kind"] != "message" || d["state"]["gui"].Title != "Shared message test" || d["width"] < 1
                continue
            s := d["state"]
            s["gui"].GetClientPos(,, &w, &h)
            dpi := GetWindowDPI(s["gui"].Hwnd) / 96
            if buttons = "yesnocancel" {
                DesktopAssert(s.Has("addButton") && s.Has("noButton"), "Desktop three-way message exposes all actions")
                DesktopAssert(s["addButton"].Text = "Save changes"
                    && s["noButton"].Text = "Discard changes"
                    && s["cancel"].Text = "Cancel", "Desktop three-way message uses explicit action labels")
                s["addButton"].GetPos(&yesX,, &yesW)
                s["noButton"].GetPos(&noX,, &noW)
                s["cancel"].GetPos(&cancelX,, &cancelW)
                DesktopAssert(yesX + yesW < noX && noX + noW < cancelX,
                    "Desktop three-way actions remain separate and ordered")
            }
            TestDesktopStudyCapture(s["gui"], "shared-message-desktop-" buttons ".png", Round(w * dpi), Round(h * dpi))
            TestSharedMessageCaptured := true
            if action = "yes"
                SendMessage(0xF5, 0, 0, s["addButton"].Hwnd)
            else if action = "no"
                SendMessage(0xF5, 0, 0, s.Get("noButton", s["cancel"]).Hwnd)
            else if action = "cancel"
                SendMessage(0xF5, 0, 0, s["cancel"].Hwnd)
            else
                CPThemedDialogFinish(s, s["gui"], s["result"])
            return
        }
    } else {
        for hwnd in WinGetList("Shared message test ahk_class AutoHotkeyGUI") {
            g := GuiFromHwnd(hwnd)
            if !g.HasOwnProp("StudyOwnedMessage") || !g.StudyOwnedMessage.Get("ready", false)
                continue
            s := g.StudyOwnedMessage
            g.GetClientPos(,, &w, &h)
            if buttons = "yesnocancel" {
                DesktopAssert(s["controls"].Has("yes") && s["controls"].Has("no")
                    && s["controls"].Has("cancel"), "Fullscreen three-way message exposes all actions")
                DesktopAssert(s["controls"]["yes"].Text = "Save changes"
                    && s["controls"]["no"].Text = "Discard changes"
                    && s["controls"]["cancel"].Text = "Cancel",
                    "Fullscreen three-way message uses explicit action labels")
            }
            TestDesktopStudyCapture(g, "shared-message-fullscreen-" buttons ".png", w, h)
            TestSharedMessageCaptured := true
            if action = "yes"
                SendMessage(0xF5, 0, 0, s["controls"]["yes"].Hwnd)
            else if action = "no"
                SendMessage(0xF5, 0, 0, s["controls"]["no"].Hwnd)
            else if action = "cancel"
                SendMessage(0xF5, 0, 0, s["controls"]["cancel"].Hwnd)
            else
                StudyLibraryOwnedMessageClose(s, s["result"])
            return
        }
    }
}

TestDesktopRecommendationConfirmControls(parent, regenerate := false) {
    global controlDarkMode
    scOwner := parent["gui"].Hwnd, scBigBox := false, scRegenerate := regenerate
    scSettings := StudyCandidatesRecommendationDefaults()
    scProviderLabel := "Gemini", scModel := "gemini-3.5-flash"
    scCounts := Map("total", 1234, "sentences", 1200, "vocabulary", 34)
    ; @STUDY_RECOMMENDATION_CONFIRM_CONTROLS@
    return scDialogState
}

TestDesktopRecommendationPreferencesControls(parent) {
    StudyCandidatesRecommendationSyncBasicSettings(parent)
    scSettings := parent["settings"], scOwner := parent["gui"].Hwnd, scBigBox := false
    ; @STUDY_RECOMMENDATION_PREFERENCES_CONTROLS@
    return scAdvanced
}

TestDesktopRecommendationPromptControls(preferences) {
    scAdvanced := preferences, scSettings := StudyCandidatesRecommendationAdvancedDraft(preferences)
    scOwner := preferences["gui"].Hwnd, scBigBox := false
    ; @STUDY_RECOMMENDATION_PROMPT_CONTROLS@
    return scPreview
}

TestDesktopRecommendationDialogs(library) {
    for regenerate in [false, true] {
        confirm := TestDesktopRecommendationConfirmControls(library, regenerate)
        DesktopTestRecommendationSizes(confirm, [[780, 670], [900, 700], [1200, 850]])
        DesktopAssert(!confirm["result"] && !confirm["closed"], "Recommendation setup does not generate automatically")
        DesktopAssert(CPDesktopIsCombo(confirm["levelDdl"].Hwnd) && CPDesktopIsCombo(confirm["styleDdl"].Hwnd), "Recommendation selectors use modern dropdowns")
        DesktopAssert(InStr(confirm["controls"]["question"].Text, "1234") && InStr(confirm["controls"]["breakdown"].Text, "1200 sentences"), "Recommendation counts survive restyling")
        DesktopAssert(InStr(confirm["controls"]["hint"].Text, regenerate ? "replaced" : "never been assessed"), "Regenerate and Generate retain different scope explanations")
        confirm["levelDdl"].Choose(1), confirm["styleDdl"].Choose(3)
        prefs := TestDesktopRecommendationPreferencesControls(confirm)
        DesktopAssert(prefs["settings"]["learnerLevel"] = "beginner" && prefs["settings"]["selectionStyle"] = "generous", "Preferences use current unsaved selection choices")
        DesktopTestRecommendationSizes(prefs, [[780, 700], [860, 740], [1200, 850]])
        prefs["vocabulary"].Value := 0, prefs["additional"].Value := "Favor reusable dialogue. 日本語"
        DesktopAssert(confirm["settings"]["focusVocabulary"] && confirm["settings"]["additionalCriteria"] = "", "Preference edits remain a local draft")
        for save in [false, true] {
            prompt := TestDesktopRecommendationPromptControls(prefs)
            DesktopTestRecommendationSizes(prompt, [[900, 700], [1000, 800], [1400, 950]])
            prompt["controls"]["editor"].Value := "Custom selection instructions. 日本語"
            for full in [true, false, true] {
                StudyCandidatesRecommendationPreviewMode(prompt, full)
                editor := prompt["controls"]["editor"]
                DesktopAssert(!!(WinGetStyle(editor.Hwnd) & 0x800) = full, "Full preview is read only; instructions remain editable")
                DesktopAssert(prompt["draftInstructions"] = "Custom selection instructions. 日本語", "Preview mode changes retain instruction draft")
                DesktopAssert(prompt["controls"]["view"].Text = (full ? "Back to editing" : "View full prompt"), "Preview mode action stays accurate")
                DesktopAssert(InStr(prompt["desktop"]["chrome"]["subtitle"].Text, full ? "Read-only" : "Edit selection"), "Preview mode has a clear header")
                if full {
                    for required in ["Custom selection instructions. 日本語", "beginner", "40-60 percent", "Favor reusable dialogue", "Return only one JSON object", "candidate data"]
                        DesktopAssert(InStr(editor.Value, required), "Full prompt retains automatic component: " required)
                } else
                    DesktopAssert(editor.Value = "Custom selection instructions. 日本語", "Editing never receives protected full-prompt rules")
            }
            DesktopDialogFixtureShow(prompt, 1000, 800)
            DesktopDialogAssertChrome(prompt, 1000, 800)
            dpi := GetWindowDPI(prompt["gui"].Hwnd) / 96
            TestDesktopStudyCapture(prompt["gui"], "desktop-recommendation-full-preview.png", Round(1000 * dpi), Round(800 * dpi))
            if !save {
                StudyCandidatesRecommendationPreviewRestore(prompt)
                DesktopAssert(!prompt["fullPreview"] && StudyCandidatesRecommendationNormalizeInstructions(prompt["controls"]["editor"].Value) = "", "Restore instructions returns to editable built-in defaults")
                prompt["controls"]["editor"].Value := "Canceled draft"
            }
            hwnd := prompt["gui"].Hwnd
            StudyCandidatesRecommendationPreviewClose(prompt, prompt["gui"], save)
            DesktopAssert(prefs.Get("promptInstructions", "") = (save ? "Custom selection instructions. 日本語" : ""), "Save draft and Cancel update only the preferences draft")
            DesktopAssert(confirm["settings"]["promptInstructions"] = "", "Saving prompt does not apply preferences or generate")
            DesktopAssert(!IsObject(StudyDesktopContext(hwnd)), "Prompt editor unregisters on close")
        }
        StudyCandidatesRecommendationAdvancedClose(prefs, true)
        DesktopAssert(prefs["applied"] && !confirm["settings"]["focusVocabulary"] && confirm["settings"]["additionalCriteria"] = "Favor reusable dialogue. 日本語", "Apply preferences transfers checkbox and guidance draft")
        DesktopAssert(confirm["settings"]["promptInstructions"] = "Custom selection instructions. 日本語" && !confirm["result"], "Applying preferences transfers prompt without generating")
        prefs := TestDesktopRecommendationPreferencesControls(confirm)
        prefs["additional"].Value := "Discard this guidance", prefs["grammar"].Value := 0
        DesktopDialogTestClose(prefs)
        DesktopAssert(!prefs["applied"] && confirm["settings"]["focusGrammar"] && confirm["settings"]["additionalCriteria"] = "Favor reusable dialogue. 日本語", "Closing preferences discards unconfirmed edits")
        prefs := TestDesktopRecommendationPreferencesControls(confirm)
        StudyCandidatesRecommendationAdvancedRestore(prefs)
        draft := StudyCandidatesRecommendationAdvancedDraft(prefs)
        for key in ["focusVocabulary", "focusGrammar", "focusNaturalPhrasing", "focusReading"]
            DesktopAssert(draft[key], "Restore preferences resets study focus: " key)
        DesktopAssert(draft["additionalCriteria"] = "" && draft["promptInstructions"] = "", "Restore preferences resets guidance and custom instructions")
        StudyCandidatesRecommendationAdvancedClose(prefs, true)
        confirm["levelDdl"].Choose(2), confirm["styleDdl"].Choose(2)
        if regenerate {
            StudyCandidatesRecommendationDialogClose(confirm, confirm["gui"], true)
            DesktopAssert(confirm["result"], "Generate confirmation explicitly authorizes generation")
        } else {
            DesktopDialogTestClose(confirm)
            DesktopAssert(!confirm["result"], "Closing recommendation setup never authorizes generation")
        }
        saved := StudyCandidatesRecommendationLoadSettings()
        DesktopAssert(saved["learnerLevel"] = "intermediate" && saved["selectionStyle"] = "balanced", "Setup retains selection settings in fixture INI without any API request")
    }
}

DesktopTestRecommendationSizes(s, sizes) {
    for size in sizes {
        w := size[1], h := size[2]
        DesktopDialogFixtureShow(s, w, h)
        DesktopDialogAssertChrome(s, w, h)
        DesktopAssert(s.Has("desktop") && s["desktop"]["paint"][s["addButton"].Hwnd]["kind"] = "primary", "Recommendation dialog has a distinct primary action")
        for key, ctrl in s["controls"] {
            ctrl.GetPos(&x, &y, &cw, &ch)
            DesktopAssert(x >= 24 && x + cw <= w - 24 && y >= 174 && y < h - 16, "Recommendation control fits: " key)
            if ctrl.Type != "DDL"
                DesktopAssert(y + ch <= h - 16, "Recommendation control stays inside footer: " key)
            if ctrl.Type = "Edit"
                DesktopAssert(!(WinGetExStyle(ctrl.Hwnd) & 0x200), "Recommendation editor has no bright native edge")
        }
        dpi := GetWindowDPI(s["gui"].Hwnd) / 96
        TestDesktopStudyCapture(s["gui"], "desktop-recommendation-" s["kind"] "-" w ".png", Round(w * dpi), Round(h * dpi))
    }
}

TestDesktopManagerControls(library) {
    slState := library, slOutputDir := A_ScriptDir "\management-ui-fixture"
    ; @STUDY_MANAGER_CONTROLS@
    return slManager
}

TestDesktopArchivesControls(manager) {
    slManager := manager, slState := manager["ownerState"], slOutputDir := A_ScriptDir "\archives-ui-fixture"
    ; @STUDY_ARCHIVES_CONTROLS@
    return slArchive
}

TestDesktopStorageControls(library) {
    slState := library
    slDirectory := "C:\Users\Example\Documents\JRPG Translator\Settings\Study Libraries\Long library name for a complete Japanese adventure"
    slState["storage"] := Map("sourceCount", 26, "explanationCount", 29,
        "databaseBytes", 143360, "mediaBytes", 574464, "mediaFiles", 3,
        "backupBytes", 143360, "backupFiles", 1, "trashBytes", 78848, "trashFiles", 1,
        "otherBytes", 0, "totalBytes", 940032, "freeBytes", 29 * 1024 ** 3, "volumeBytes", 2 * 1024 ** 4)
    ; @STUDY_STORAGE_CONTROLS@
    return slDesktop
}

TestDesktopManagementDialogs(library) {
    global studyLibrariesRoot, studyLibraryDefaultDir
    studyLibrariesRoot := A_ScriptDir "\name-dialog-fixture\libraries"
    studyLibraryDefaultDir := A_ScriptDir "\name-dialog-fixture\default"
    manager := TestDesktopManagerControls(library)
    manager["list"].Add("", "Default", "", 26, 29, "17 MB")
    manager["list"].Add("Select Focus", library["libraryName"], "Active", 3, 3, "918 KB")
    manager["list"].Add("", "Dragon Quest — Japanese adventure", "", 156, 240, "42 MB")
    archives := TestDesktopArchivesControls(manager)
    for s in [manager, archives] {
        kind := s["archived"] ? "archives" : "manager", g := s["gui"], list := s["list"]
        for size in [[900, 660], [1000, 700], [1400, 900]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(s, w, h)
            DesktopDialogAssertChrome(s, w, h)
            DesktopAssert(!(SendMessage(0x1037, 0, 0, list.Hwnd) & 1), kind " removes dense grid lines")
            DesktopAssert(SendMessage(0x1037, 0, 0, list.Hwnd) & 0x10000, kind " table is double buffered")
            DesktopAssert(WinGetStyle(list.Hwnd) & 0x8000, kind " disallows header sorting to preserve row mapping")
            DesktopAssert(SendMessage(0x101D, 0, 0, list.Hwnd) >= 180, kind " leaves room for library names")
            DesktopAssert(SendMessage(0x101D, 3, 0, list.Hwnd) >= 114, kind " explanation heading is readable")
            if s["archived"] {
                StudyLibraryArchiveUpdateActions(s)
                s["emptyText"].Text := "No archived Study Libraries."
                DesktopAssert(!s["restoreButton"].Enabled && !s["openButton"].Enabled, "Empty archives disable row actions")
            } else {
                for row in [1, 2, 3] {
                    list.Modify(0, "-Select"), list.Modify(row, "Select Focus")
                    StudyLibraryManagerUpdateActions(s)
                    DesktopAssert(s["switchButton"].Enabled = (row != 2), "Active library cannot switch to itself")
                    DesktopAssert(s["renameButton"].Enabled = (row != 1) && s["archiveButton"].Enabled = (row != 1), "Default cannot be renamed or archived")
                }
                list.Modify(0, "-Select"), list.Modify(2, "Select Focus")
                StudyLibraryManagerUpdateActions(s)
            }
            dpi := GetWindowDPI(g.Hwnd) / 96
            TestDesktopStudyCapture(g, "desktop-library-" kind "-" w ".png", Round(w * dpi), Round(h * dpi))
        }
    }
    entry := Map("name", "Completed adventure", "path", A_ScriptDir "\archived-fixture", "archivedAt", "2026-09-08 17:30:00")
    archives["entries"] := [entry]
    archives["list"].Add("Select Focus", entry["name"], entry["archivedAt"], 26, 29, "17 MB")
    archives["emptyText"].Text := ""
    StudyLibraryArchiveUpdateActions(archives)
    DesktopAssert(archives["restoreButton"].Enabled && archives["openButton"].Enabled, "Selected archive enables restore and open")
    DesktopAssert(StudyLibraryArchiveSelectedEntry(archives) = entry, "Selected row keeps its archive mapping")
    DesktopDialogFixtureShow(archives, 1000, 700)
    dpi := GetWindowDPI(archives["gui"].Hwnd) / 96
    TestDesktopStudyCapture(archives["gui"], "desktop-library-archives-populated.png", Round(1000 * dpi), Round(700 * dpi))
    for mode in ["new", "rename", "restore"] {
        s := StudyDesktopLibraryNameCreate(mode = "restore" ? archives : manager, mode,
            mode = "restore" ? entry : "Japanese adventure")
        initial := s["nameEdit"].Value
        DesktopAssert(initial = (mode = "new" ? "" : mode = "rename" ? "Japanese adventure" : entry["name"]), mode " starts with the correct name")
        for size in [[680, 500], [760, 500], [1000, 650]] {
            w := size[1], h := size[2]
            DesktopDialogFixtureShow(s, w, h)
            DesktopDialogAssertChrome(s, w, h)
            DesktopAssert(s["nameEdit"].Value = initial, mode " preserves name during resize")
            DesktopAssert(!(WinGetExStyle(s["nameEdit"].Hwnd) & 0x200), mode " field has no white client edge")
            DesktopAssert(s["desktop"]["paint"][s["addButton"].Hwnd]["kind"] = "primary", mode " has a clear primary action")
            dpi := GetWindowDPI(s["gui"].Hwnd) / 96
            TestDesktopStudyCapture(s["gui"], "desktop-library-" mode "-" w ".png", Round(w * dpi), Round(h * dpi))
        }
        s["nameEdit"].Value := "未保存の名前"
        DesktopDialogTestClose(s)
        DesktopAssert(manager["list"].GetText(2) = library["libraryName"], "Cancel name dialog does not change library data")
    }
    s := TestDesktopStorageControls(library), path := s["path"].Value
    for size in [[760, 770], [860, 770], [1000, 850]] {
        w := size[1], h := size[2]
        DesktopDialogFixtureShow(s, w, h)
        DesktopDialogAssertChrome(s, w, h)
        DesktopAssert(s["path"].Value = path && StrLen(path) > 100, "Storage retains the complete long folder path")
        DesktopAssert((WinGetStyle(s["path"].Hwnd) & 0x804) = 0x804, "Storage path is multiline and read only")
        for key, ctrl in s["values"] {
            ctrl.GetPos(&x, &y, &cw, &ch)
            DesktopAssert(y + ch <= h - 70 && x + cw <= w, "Storage values stay above footer: " key)
            if key != "notice"
                DesktopAssert(DesktopTextWidth(ctrl) < cw - 4, "Storage value fits: " key)
        }
        DesktopAssert(InStr(s["values"]["entries"].Value, "26 sources") && InStr(s["values"]["entries"].Value, "29 explanations"), "Storage displays source and explanation counts")
        dpi := GetWindowDPI(s["gui"].Hwnd) / 96
        TestDesktopStudyCapture(s["gui"], "desktop-library-storage-" w ".png", Round(w * dpi), Round(h * dpi))
    }
    library["storage"]["freeBytes"] := 1024 ** 3
    StudyLibraryUpdateStorageDialog(library, s["values"])
    DesktopAssert(InStr(s["values"]["notice"].Value, "running low"), "Storage retains low disk space warning")
    library["storage"]["freeBytes"] := 29 * 1024 ** 3
    library["storage"]["sourceCount"] := 5000
    StudyLibraryUpdateStorageDialog(library, s["values"])
    DesktopAssert(InStr(s["values"]["notice"].Value, "large library"), "Storage retains large library guidance")
    DesktopDialogTestClose(s)
    DesktopDialogTestClose(archives)
    DesktopDialogTestClose(manager)
}

DesktopDialogTestClose(s) {
    hwnd := s["gui"].Hwnd
    PostMessage(0x10, 0, 0, hwnd)
    Sleep(35)
    DesktopAssert(!DllCall("user32\IsWindow", "ptr", hwnd) && !IsObject(StudyDesktopContext(hwnd)), "Dialog closes and unregisters cleanly")
}

DesktopDialogFixtureShow(s, w, h) {
    g := s["gui"]
    g.Show("Hide w" w " h" h)
    StudyDesktopDialogResize(s, g, 0, w, h)
    CPApplyOwnedDialogTheme(g)
    g.Show("NA x-12000 y-12000 w" w " h" h)
    Sleep(35)
}

DesktopDialogAssertChrome(s, w, h) {
    DesktopAssert(!(WinGetStyle(s["gui"].Hwnd) & 0xC00000), "Modern dialog has no generic caption")
    for hwnd, item in s["desktop"]["paint"] {
        ctrl := item["ctrl"]
        if !DesktopShown(ctrl)
            continue
        ctrl.GetPos(&x, &y, &cw, &ch)
        DesktopAssert(x >= 0 && y >= 0 && x + cw <= w + 1 && y + ch <= h + 1, "Dialog button stays inside window: " ctrl.Text)
        DesktopAssert((WinGetStyle(hwnd) & 0xF) = 0xB, "Dialog button keeps modern rendering")
        if item["kind"] != "caption"
            DesktopAssert(DesktopTextWidth(ctrl) < cw - 15, "Dialog label fits: " ctrl.Text)
    }
}

; @CAPTURE@
; All production functions are available, but no production startup runs.
#Include %A_ScriptDir%\JRPG Translator.ahk
#Warn All, StdOut
