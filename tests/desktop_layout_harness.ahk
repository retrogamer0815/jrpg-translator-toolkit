#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
; @UI_GLOBALS@
global TestAssertions := 0, TestCaptureClicks := 0
global controlDarkMode := 1, iniPath := A_ScriptDir "\test-settings.ini", controlPanelOpacity := 100
global gPidAudio := 0, gJustStoppedUntil := 0, gLastAction := ""
global gAudioInputJob := Map("active", false), gAudioInputStatus := "Not tested.", speakerName := "[Windows Default]"
global __DBG_ENABLED_CP := false
global overlayTrans := 220, boxBgHex := "202020", txtHex := "FFFFFF", nameHex := "56C4F5"
global fontName := "Segoe UI", fontSize := 18, fontBold := 0
global overlayTrans_EW := 192, boxBgHex_EW := "101824", txtHex_EW := "E8E8E8"
global fontName_EW := "Consolas", fontSize_EW := 22, fontBold_EW := 1
global DesktopOverlayMoves := Map(3, 0, 5, 0)
global defGuiW := 1120, defGuiH := 760, pad := 12, gap := 8
global tabNames := ["Screenshot Translation", "Audio Translation", "Translation Window", "Explanation",
    "Explanation Window", "Terminology Overrides", "Profiles", "Controls", "API Keys", "Paths"]
global CPTabVisiblePages := [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
; Begin with a real native caption, just like production. The desktop shell
; removes it; PrintWindow(PW_CLIENTONLY) excludes the retained resize frame.
global ui := Gui("+Resize +0x300000", "Synthetic desktop layout")
ui.SetFont("s10", "Segoe UI")
CPRegisterThemeMessages()
CPRefreshThemeBrushes()
global tab := ui.Add("Tab", "x12 y12 w1000 h560 Buttons -Wrap", tabNames)
CPRegisterCanvasFixedControl(tab, false, true)
tab.UseTab(1)
global ddlProv := ui.Add("DropDownList", "x120 y60 w220 0x210", ["Gemini", "OpenAI"])
ddlProv.Choose(1)
global ddlIMG_GM := ui.Add("DropDownList", "x120 y100 w260 0x210", ["gemini-3.5-flash", "gemini-long-alternative"])
ddlIMG_GM.Choose(1)
global ddlIMG := ui.Add("DropDownList", "x120 y140 w260 0x210", ["gpt-4o", "gpt-alternative"])
ddlIMG.Choose(1)
global ddlPrompt := ui.Add("DropDownList", "x120 y180 w260 0x210", ["default_with_kanji_reading_en", "literal"])
ddlPrompt.Choose(1)
global btnPrEdit := ui.Add("Button", "x390 y180 w70 h32", "Edit…")
global btnPrNew := ui.Add("Button", "x470 y180 w70 h32", "Add…")
global btnPrDel := ui.Add("Button", "x550 y180 w70 h32", "Delete")
global btnIMG_Add := ui.AddButton("x400 y140 w70", "Add…"), btnIMG_Del := ui.AddButton("x480 y140 w70", "Delete")
global btnIMG_GM_Add := ui.AddButton("x400 y100 w70", "Add…"), btnIMG_GM_Del := ui.AddButton("x480 y100 w70", "Delete")
global chkGuess := ui.AddCheckbox("x460 y240", "Highlight guessed subjects")
global chkName := ui.AddCheckbox("x460 y300", "Use speaker name color")
global chkDel := ui.AddCheckbox("x12 y240", "Clear screenshots on startup")
global chkOpenTW := ui.AddCheckbox("x12 y500", "Open translation window with JRPG Translator")
global chkTop_TW := ui.AddCheckbox("x12 y540", "Open translation window always on top")
global eCapMax := ui.AddEdit("x220 y300 w80 Number", "1400")
global btnCapPick := ui.AddButton("x12 y340 w160", "Capture…")
global btnST := ui.AddButton("x12 y400 w200", "Screenshot + Translate")
btnST.OnEvent("Click", (*) => TestCaptureClicked())
global btnTS := ui.AddButton("x220 y400 w180", "Take Screenshot")
global btnSTO := ui.AddButton("x420 y400 w220", "Screenshot -> Translation")
global CPDesktopShotControls := []
for testCtrl in ui {
    if testCtrl.Hwnd != tab.Hwnd
        CPDesktopShotControls.Push(testCtrl)
}
tab.UseTab(2)
global TestAudioLegacyText := ui.AddText("x24 y60 w360 h30", "Existing audio settings")
global ddlSpeaker := ui.AddDropDownList("x120 y120 w360 0x210", ["[Windows Default]", "Game speakers", "Headphones"])
global btnSpRef := ui.AddButton("x490 y120 w80", "Refresh"), btnAudioTest := ui.AddButton("x578 y120 w90", "Test Audio")
global ddlAProv := ui.AddDropDownList("x120 y220 w220 0x210", ["Gemini", "OpenAI"])
global ddlA_GM := ui.AddDropDownList("x120 y260 w260 0x210", ["gemini-3.5-live-translate-preview", "gemini-alternative"])
global ddlTR := ui.AddDropDownList("x120 y300 w420 0x210", ["gpt-realtime", "gpt-alternative"])
global ddlAudioTarget := ui.AddDropDownList("x120 y350 w260 0x210", ["English (en)", "German (de)", "Japanese (ja)"])
for testCtrl in [ddlSpeaker, ddlAProv, ddlA_GM, ddlTR, ddlAudioTarget]
    testCtrl.Choose(1)
global btnA_GM_Add := ui.AddButton(), btnA_GM_Del := ui.AddButton(), btnTR_Add := ui.AddButton(), btnTR_Del := ui.AddButton()
global txtAudioHelp := ui.AddText(), txtAudioTestStatus := ui.AddText()
CPDesktopAudioControls := []
for testCtrl in ui {
    if testCtrl.Hwnd != tab.Hwnd && !ArrayIndexOf(CPDesktopShotControls, testCtrl)
        CPDesktopAudioControls.Push(testCtrl)
}
testTranslator := DesktopBuildOverlayFixture(3)
global slTrans := testTranslator["opacity"], lblTransPct := testTranslator["percent"]
global rectBg := testTranslator["bg"], rectTxt := testTranslator["txt"], rectName := testTranslator["name"]
global ddlFont := testTranslator["font"], edFSize := testTranslator["size"], udFSize := testTranslator["spinner"]
global chkFontBold := testTranslator["bold"], txtFontSizeHint := testTranslator["sizeHint"]
global btnMoveResize := testTranslator["position"], txtMoveResize := testTranslator["positionHelp"]
testExplainer := DesktopBuildOverlayFixture(5)
global slTrans_EW := testExplainer["opacity"], lblTransPct_EW := testExplainer["percent"]
global rectBg_EW := testExplainer["bg"], rectTxt_EW := testExplainer["txt"]
global ddlFont_EW := testExplainer["font"], edFSize_EW := testExplainer["size"], udFSize_EW := testExplainer["spinner"]
global chkFontBold_EW := testExplainer["bold"], txtFontSizeHint_EW := testExplainer["sizeHint"]
global btnMoveResize_EW := testExplainer["position"], txtMoveResize_EW := testExplainer["positionHelp"]
DesktopBuildOrganizeFixtures()
global controllerOptionsX := 620, controllerTopY := 70
testBeforeExplanation := Map()
for testCtrl in ui
    testBeforeExplanation[testCtrl.Hwnd] := true
tab.UseTab(4)
global TestExplanationLegacyText := ui.AddText("x24 y60 w360 h30", "Existing explanation settings")
global ddlEProv := ui.AddDropDownList("x120 y100 w220 0x210", ["Gemini", "OpenAI"])
global ddlEGem := ui.AddDropDownList("x120 y140 w260 0x210", ["gemini-explanation", "gemini-explanation-alternative"])
global ddlEOpenAI := ui.AddDropDownList("x120 y180 w260 0x210", ["gpt-explanation", "gpt-explanation-alternative"])
global ddlEPr := ui.AddDropDownList("x120 y220 w260 0x210", ["default_en", "detailed_grammar"])
for testCtrl in [ddlEProv, ddlEGem, ddlEOpenAI, ddlEPr]
    testCtrl.Choose(1)
global btnEGem_Add := ui.AddButton(), btnEGem_Del := ui.AddButton(), btnEOpenAI_Add := ui.AddButton(), btnEOpenAI_Del := ui.AddButton()
global btnEPrEdit := ui.AddButton(, "Edit…"), btnEPrNew := ui.AddButton(), btnEPrDel := ui.AddButton()
global btnExplainNow := ui.AddButton("x12 y280 w220", "Explain last jp. Text")
global btnOpenStudyLibrary := ui.AddButton("x240 y280 w200", "Open Study Library…")
global DesktopExplanationClicks := 0
btnExplainNow.OnEvent("Click", DesktopExplanationClicked)
global saveLibraryChk := ui.AddCheckbox("x12 y330 w400", "Save explanations to Study Library")
global saveLibraryScreenshotsChk := ui.AddCheckbox("x12 y370 w440", "Include source screenshots in Study Library")
global saveExplChk := ui.AddCheckbox("x12 y410 w260", "Save plain-text copies")
global txtExplainSaveInfo := ui.AddText("x12 y450 w600 h40", "Original saving information")
global chkOpenEW := ui.AddCheckbox("x12 y510 w440", "Open explanation window with JRPG Translator")
global chkTop_EW := ui.AddCheckbox("x12 y550 w440", "Open explanation window always on top")
saveLibraryChk.Value := 1, saveLibraryScreenshotsChk.Value := 1
CPDesktopExplanationControls := []
for testCtrl in ui {
    if !testBeforeExplanation.Has(testCtrl.Hwnd)
        CPDesktopExplanationControls.Push(testCtrl)
}
tab.UseTab()
global CPFooterFill := ui.AddText("x0 y570 w1120 h190"), sepAction := ui.AddText("x12 y570 w1000 h2 0x10")
global btnOv := ui.AddButton(, "Open Translator"), btnOvClose := ui.AddButton(, "Close Translator")
global btnAudio := ui.AddButton(, "Audio Translation Off"), btnExplainerLaunch := ui.AddButton(, "Open Explainer")
global btnExplainerClose := ui.AddButton(, "Close Explainer"), bClose := ui.AddButton(, "Close all")
global chkTop := ui.AddCheckbox(, "Always on top"), chkDarkMode := ui.AddCheckbox(, "Dark mode")
global txtControlOpacity := ui.AddText(, "Opacity:"), slControlOpacity := ui.AddSlider("Range70-100", 100)
global lblControlOpacityPct := ui.AddText(, "100%")
chkGuess.Value := 1, chkName.Value := 1, chkDarkMode.Value := 1
for testCtrl in CPDesktopFooterControls()
    CPRegisterCanvasFixedControl(testCtrl, false, true)
CPRegisterCanvasFixedControl(CPFooterFill, false, true)
CPRegisterCanvasFixedControl(sepAction, false, true)
CPCreateCustomTabBar()
CPDesktopCreate()
CPRegisterThemeMessages()
CPRegisterCanvasMessages()
CPRefreshThemeBrushes()
CPCreateComboArrowOverlays()
OnError(DesktopTestUnhandledError)
try {
    ; @STUDY_ONLY@
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
        DesktopAssert(DesktopHitAt(testW - 34, 28) = 1, "Close button is not a drag target")
        DesktopAssert(DesktopHitAt(300, 100) = 1, "Page content is not a drag target")
        DesktopAssert(CPDesktopIsCombo(ddlProv.Hwnd) && CPDesktopIsCombo(ddlSpeaker.Hwnd) && CPDesktopIsCombo(ddlEPr.Hwnd) && CPDesktopIsCombo(ddlFont.Hwnd) && CPDesktopIsCombo(ddlFont_EW.Hwnd) && CPDesktopIsCombo(ddlGameProfile.Hwnd) && CPDesktopIsCombo(ddlJPG.Hwnd) && !CPDesktopIsCombo(TestLegacyCombo.Hwnd), "Rounded dropdown rendering covers the refreshed pages")
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
        DesktopAssert(!DesktopShown(CPDesktop["chrome"]["appearance"]) && !DesktopShown(CPDesktop["chrome"]["classic"]), "Duplicate window options and classic fallback do not clutter sidebar")
        DesktopAssert(CPDesktop["paint"][CPDesktop["chrome"]["screenshot"].Hwnd]["selected"], "Screenshot sidebar item is selected")
        DesktopAssert(CPDesktop["contentW"] <= 1040, "Content width stays bounded on wide displays")
        ddlIMG.GetPos(, &modelRowY)
        ddlPrompt.GetPos(, &promptRowY)
        DesktopAssert(testW >= 1120 ? modelRowY = promptRowY : promptRowY > modelRowY, "AI fields reflow on narrow windows")
        DesktopAssert(CPDesktop["chrome"]["profile"].Text = "Current settings", "Header describes current settings without applying a profile")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlProv.Hwnd, "int", 0, "ptr") = ddlIMG_GM.Hwnd, "Keyboard Tab reaches the active model")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlIMG_GM.Hwnd, "int", 0, "ptr") = CPDesktop["shot"]["models"].Hwnd, "Keyboard Tab reaches model management next")
        DesktopAssert(DesktopShown(ddlIMG_GM) && !DesktopShown(ddlIMG), "Only active Gemini model is shown")
        DesktopAssert(!DesktopShown(chkDel) && !DesktopShown(eCapMax), "Advanced options start collapsed")
        DesktopAssert(ddlPrompt.Text = "default_with_kanji_reading_en", "Prompt choice preserved")
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
        DesktopAssert(!DesktopShown(TestLegacyText) && DesktopShown(ddlJPG) && !DesktopShown(btnST), "Settings replaces the legacy labels and retains its native fields")
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
    ; Use native button activation, without invoking any real capture/API work.
    SendMessage(0xF5, 0, 0, btnST.Hwnd)
    Sleep(20)
    DesktopAssert(TestCaptureClicks = 1, "Restyled primary button retains its original Click handler")
    for desktopTestPage in [3, 4, 5, 6, 7, 8, 9, 10, 1] {
        CPDesktopNavigate(desktopTestPage)
        DesktopAssert(tab.Value = desktopTestPage, "Grouped navigation reaches existing page " desktopTestPage)
    }
    controlDarkMode := 0
    CPApplyControlPanelTheme()
    TestCapture("desktop-light.png", Round(captureW * testScale), Round(captureH * testScale))
    DesktopAssert(chkGuess.Value = 1 && chkName.Value = 1, "Formatting values survive layout and theme changes")
    ddlProv.Choose(2), ddlIMG.Choose(2), ddlPrompt.Choose(2), eCapMax.Value := "1800"
    CPDesktopToggleLayout()
    DesktopAssert(!CPDesktopActive(), "Classic fallback is available")
    DesktopAssert((DllCall("user32\GetWindowLongPtr", "ptr", ui.Hwnd, "int", -16, "ptr") & 0xC00000) = 0xC00000, "Classic fallback restores native title bar")
    DesktopAssert(!CPDesktopIsCombo(ddlPrompt.Hwnd), "Classic fallback restores native dropdown appearance")
    DesktopAssert(SendMessage(0x154, -1, 0, ddlPrompt.Hwnd) = CPDesktop["combos"][ddlPrompt.Hwnd]["height"], "Classic fallback restores original dropdown height")
    DesktopAssert(DesktopShown(CPDesktop["chrome"]["classic"]), "Modern layout return action remains available")
    DesktopAssert(ddlProv.Text = "OpenAI" && ddlIMG.Text = "gpt-alternative" && ddlPrompt.Text = "literal" && eCapMax.Value = "1800", "Classic fallback preserves live settings")
    DesktopAssert((DllCall("user32\GetWindowLongPtr", "ptr", btnST.Hwnd, "int", -16, "ptr") & 0xF) != 0xB, "Classic fallback restores native button rendering")
    CPDesktopToggleLayout()
    DesktopAssert(CPDesktopActive() && ddlIMG.Text = "gpt-alternative" && ddlPrompt.Text = "literal", "Returning to modern keeps current values")
    global DesktopTestCloseCount := 0
    ui.OnEvent("Close", DesktopTestClosed)
    CPDesktopWindowAction("close")
    Sleep(30)
    DesktopAssert(DesktopTestCloseCount = 1, "Custom close routes through the existing GUI Close event")
    TestDesktopStudyWindows()
    DesktopTestWheelLifetime()
    ui.Destroy()
    DesktopAssert(CPCanvasWheelWindows.Count = 0, "Normal GUI destruction removes all wheel tracking")
    FileAppend("PASS: " TestAssertions " desktop layout assertions.`n", "*")
    ExitApp(0)
} catch as testError {
    FileAppend("FAIL: " testError.Message "`n" testError.Extra "`n" testError.File ":" testError.Line "`n" testError.Stack "`n", "*")
    ExitApp(1)
}

DesktopAssert(condition, message) {
    global TestAssertions
    TestAssertions += 1
    if !condition
        throw Error(message)
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
TestCaptureClicked() {
    global TestCaptureClicks
    TestCaptureClicks += 1
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
    global btnSpRef, btnAudioTest, btnA_GM_Add, btnA_GM_Del, btnTR_Add, btnTR_Del
    global TestAudioLegacyText, gAudioInputJob, gAudioInputStatus, gPidAudio, speakerName, iniPath
    a := CPDesktop["audioPage"]
    for size in [[1120, 760], [900, 640], [1400, 820], [820, 560]] {
        audioW := size[1], audioH := size[2]
        ui.Show("NA x-9000 y-9000 w" audioW " h" audioH)
        tab.Value := 2
        CPDesktopLayout(ui, 0, audioW, audioH)
        ToggleAudioControls()
        DesktopAssert(DesktopShown(ddlAProv) && !DesktopShown(TestAudioLegacyText), "Audio uses the new grouped layout")
        DesktopAssert(DesktopShown(ddlA_GM) && !DesktopShown(ddlTR), "Audio shows only the selected provider's model")
        DesktopAssert(!DesktopShown(btnA_GM_Add) && !DesktopShown(btnTR_Del), "Audio model Add/Delete rows are replaced by Manage models")
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
        DesktopAssert(!DesktopShown(a["power"]) && !DesktopShown(ddlSpeaker), "Audio controls do not leak onto Screenshot Translation")
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
    gPidAudio := DllCall("kernel32\GetCurrentProcessId") ; Status only: never start or stop an actual session.
    CPDesktopRefreshAudio()
    DesktopAssert(a["power"].Text = "Stop audio translation" && InStr(a["sessionHelp"].Text, "restart"), "Running session presents Stop and explains pending settings")
    gPidAudio := 0
    CPDesktopRefreshAudio()
    DesktopAssert(a["power"].Text = "Start audio translation", "Stopped session presents Start")
    ddlAudioTarget.Choose(2), ddlTR.Choose(2), ddlAProv.Choose(2), ToggleAudioControls()
    CPSetAudioDeviceSelection("Game speakers")
    CPDesktopToggleLayout()
    DesktopAssert(DesktopShown(TestAudioLegacyText) && DesktopShown(ddlTR) && DesktopShown(ddlA_GM), "Classic fallback restores the complete original audio form")
    DesktopAssert(!DesktopShown(a["power"]) && btnSpRef.Text = "Refresh" && btnAudioTest.Text = "Test Audio", "Classic fallback restores original audio actions")
    DesktopAssert(!CPDesktopIsCombo(ddlSpeaker.Hwnd) && ddlAudioTarget.Text = "German (de)" && ddlSpeaker.Text = "Game speakers", "Classic audio keeps selected values and native dropdowns")
    DesktopAssert((DllCall("user32\GetWindowLongPtr", "ptr", ui.Hwnd, "int", -20, "ptr") & 0x02000000) = CPDesktop["compositedStyle"], "Classic fallback restores its original buffering style")
    CPDesktopToggleLayout()
    DesktopAssert(CPDesktopIsCombo(ddlSpeaker.Hwnd) && DesktopShown(ddlTR) && !DesktopShown(ddlA_GM), "Returning to modern restores the selected audio model")
    DesktopAssert(ddlTR.Text = "gpt-alternative" && ddlSpeaker.Text = "Game speakers" && ddlAudioTarget.Text = "German (de)", "Audio choices survive both layout changes")
    DesktopAssert(IniRead(iniPath, "cfg", "speakerName") = "Game speakers", "Layout changes preserve the saved audio device")
    SetAudioTestStatus("Not tested.")
    ddlAProv.Choose(1), ddlTR.Choose(1), ddlAudioTarget.Choose(1), ddlSpeaker.Choose(1)
    speakerName := "[Windows Default]"
    ToggleAudioControls()
    tab.Value := 1
    ui.Show("NA x-9000 y-9000 w1400 h820")
    CPDesktopLayout(ui, 0, 1400, 820)
}
DesktopBuildOverlayFixture(page) {
    global ui, tab, CPDesktopOverlayLegacy
    fixtureBefore := Map(), fixture := Map(), ew := page = 5
    for ctrl in ui
        fixtureBefore[ctrl.Hwnd] := true
    tab.UseTab(page)
    fixture["legacy"] := ui.AddText("x24 y60 w360 h30", "Original overlay settings")
    fixture["opacity"] := ui.AddSlider("x120 y100 w280 h30 Range0-255 ToolTip", ew ? 192 : 220)
    fixture["percent"] := ui.AddText("x420 y100 w60", ew ? "75%" : "86%")
    colors := ew ? Map("bg", "101824", "txt", "E8E8E8") : Map("bg", "202020", "txt", "FFFFFF", "name", "56C4F5")
    for key, hex in colors {
        fixture[key] := CPRegisterColorSwatch(ui.AddText("x120 y+14 w84 h34 Border Background" hex), (ew ? "explainer:" : "translator:") key)
    }
    fixture["font"] := ui.AddDropDownList("x120 y290 w280 0x210", ["Segoe UI", "Consolas", "Arial"])
    fixture["font"].Choose(ew ? 2 : 1)
    fixture["size"] := ui.AddEdit("x416 y290 w60 Number", ew ? 22 : 18)
    fixture["spinner"] := ui.AddUpDown(ew ? "Range6-200" : "Range6-128", ew ? 22 : 18)
    fixture["sizeHint"] := ui.AddText("x488 y290 w44 h23 Hidden Center Border +0x200", "A")
    fixture["bold"] := ui.AddCheckbox("x540 y290 w90", "Bold")
    fixture["bold"].Value := ew
    fixture["position"] := ui.AddButton("x120 y350 w180 h32", "Move / Resize")
    fixture["position"].OnEvent("Click", DesktopOverlayMoved.Bind(page))
    fixture["positionHelp"] := ui.AddText("x120 y390 w590 h44", "Left stick or arrows move; right stick or Screenshot + Translate + arrows resize. Enter saves, Esc cancels.")
    CPRegisterMutedControl(fixture["positionHelp"])
    for ctrl in ui {
        if !fixtureBefore.Has(ctrl.Hwnd)
            CPDesktopOverlayLegacy[page].Push(ctrl)
    }
    return fixture
}

DesktopOverlayMoved(page, *) {
    global DesktopOverlayMoves
    DesktopOverlayMoves[page] += 1
}

DesktopTestOverlayPages() {
    global ui, tab, CPDesktop, CPDesktopOverlayLegacy, DesktopOverlayMoves, testTranslator, testExplainer
    global boxBgHex, boxBgHex_EW, nameHex, ddlFont, ddlFont_EW
    for page in [3, 5] {
        overlayPage := CPDesktop["overlayPages"][page], bindings := CPDesktopOverlayBindings(page)
        fixture := page = 3 ? testTranslator : testExplainer
        otherBindings := CPDesktopOverlayBindings(page = 3 ? 5 : 3)
        originalFont := bindings["font"].Text, originalSize := bindings["size"].Value
        originalOpacity := bindings["opacity"].Value, originalBold := bindings["bold"].Value
        for size in [[1120, 760], [900, 640], [1400, 820], [820, 560]] {
            overlayW := size[1], overlayH := size[2]
            ui.Show("NA x-9000 y-9000 w" overlayW " h" overlayH)
            tab.Value := page
            CPDesktopLayout(ui, 0, overlayW, overlayH)
            DesktopAssert(CPDesktop["chrome"]["title"].Text = "Overlay windows" && InStr(CPDesktop["chrome"]["subtitle"].Text, bindings["title"]), "Shared overlay heading identifies the selected window")
            DesktopAssert(!DesktopShown(fixture["legacy"]) && !DesktopShown(fixture["bg"]), "Modern overlays hide the classic form and swatches")
            DesktopAssert(DesktopShown(bindings["font"]) && !DesktopShown(otherBindings["font"]), "Only the selected overlay's controls are shown")
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
            tab.Value := 1
            CPDesktopLayout(ui, 0, overlayW, overlayH)
            DesktopAssert(!DesktopShown(overlayPage["translator"]) && !DesktopShown(bindings["position"]), "Overlay controls do not leak onto Screenshot Translation")
            CPDesktopOverlayMenu()
            DesktopAssert(tab.Value = page, "Overlay sidebar returns directly to the last selected window")
            DesktopAssert(DesktopRect(bindings["font"]) = navigationRect, "Overlay navigation restores its layout at the actual client width")
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
        bindings["font"].Choose(3), bindings["size"].Value := 27, bindings["spinner"].Value := 27
        bindings["opacity"].Value := 150, bindings["percent"].Text := "59%"
        bindings["bold"].Value := !originalBold
        SendMessage(0xF5, 0, 0, bindings["position"].Hwnd)
        Sleep(20)
        DesktopAssert(DesktopOverlayMoves[page] = 1, "Restyled move/resize retains the original Click handler")
        CPDesktopToggleLayout()
        DesktopAssert(DesktopShown(fixture["legacy"]) && DesktopShown(fixture["bg"]), "Classic mode restores the overlay form and original swatches")
        DesktopAssert(!DesktopShown(overlayPage["bg"]) && !CPDesktopIsCombo(bindings["font"].Hwnd), "Classic overlay removes modern color buttons and font styling")
        DesktopAssert(bindings["position"].Text = "Move / Resize", "Classic move/resize label is restored")
        DesktopAssert(!DesktopShown(bindings["sizeHint"]), "Classic overlay does not expose an inactive controller hint")
        CPDesktopToggleLayout()
        DesktopAssert(bindings["font"].Text = "Arial" && bindings["size"].Value = 27 && bindings["opacity"].Value = 150 && bindings["percent"].Text = "59%" && bindings["bold"].Value = !originalBold, "Overlay choices and opacity readout survive both layout changes")
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
    global ui, tab, CPDesktop, ddlEProv, ddlEGem, ddlEOpenAI, ddlEPr, btnEPrEdit, btnEPrNew, btnEPrDel
    global btnEGem_Add, btnEOpenAI_Del, btnExplainNow, btnOpenStudyLibrary, DesktopExplanationClicks
    global saveLibraryChk, saveLibraryScreenshotsChk, saveExplChk, chkOpenEW, chkTop_EW, iniPath
    global TestExplanationLegacyText, txtExplainSaveInfo, ddlProv, ddlPrompt, ddlAProv, ddlAudioTarget
    expPage := CPDesktop["explanationPage"]
    otherChoices := ddlProv.Text "|" ddlPrompt.Text "|" ddlAProv.Text "|" ddlAudioTarget.Text
    for size in [[1120, 760], [900, 640], [1400, 820], [820, 560]] {
        w := size[1], h := size[2]
        ui.Show("NA x-9000 y-9000 w" w " h" h)
        tab.Value := 4
        CPDesktopLayout(ui, 0, w, h)
        ToggleExplanationControls()
        DesktopAssert(DesktopShown(ddlEProv) && !DesktopShown(TestExplanationLegacyText) && !DesktopShown(txtExplainSaveInfo), "Explanation uses the grouped layout without legacy labels")
        DesktopAssert(DesktopShown(ddlEGem) && !DesktopShown(ddlEOpenAI), "Explanation shows only the active provider's model")
        DesktopAssert(!DesktopShown(btnEGem_Add) && !DesktopShown(btnEOpenAI_Del) && !DesktopShown(btnEPrNew) && !DesktopShown(btnEPrDel), "Explanation management is consolidated into links")
        DesktopAssert(CPDesktop["paint"][CPDesktop["chrome"]["explanation"].Hwnd]["selected"], "Explanation sidebar row is selected")
        DesktopAssert(!DesktopShown(chkOpenEW) && !DesktopShown(chkTop_EW), "Explanation startup choices begin collapsed")
        DesktopAssert(CPDesktop["paint"][btnExplainNow.Hwnd]["kind"] = "primary" && btnExplainNow.Text = "Explain latest text", "Explanation has one clear primary action")
        ddlEGem.GetPos(, &modelY), ddlEPr.GetPos(, &promptY)
        DesktopAssert(w >= 1120 ? modelY = promptY : promptY > modelY, "Explanation AI fields reflow at narrow widths")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlEProv.Hwnd, "int", 0, "ptr") = ddlEGem.Hwnd, "Explanation keyboard order reaches the active model")
        DesktopAssert(DllCall("user32\GetNextDlgTabItem", "ptr", ui.Hwnd, "ptr", ddlEGem.Hwnd, "int", 0, "ptr") = expPage["models"].Hwnd, "Explanation keyboard order reaches model management")
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
        DesktopAssert(!DesktopShown(expPage["models"]) && !DesktopShown(btnExplainNow), "Explanation actions do not leak onto Screenshot Translation")
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
    ; The original button HWND and callback still perform the action, not a duplicate proxy.
    SendMessage(0xF5, 0, 0, btnExplainNow.Hwnd)
    Sleep(20)
    DesktopAssert(DesktopExplanationClicks = 1, "Restyled explanation action retains its original Click handler")
    ddlEProv.Choose(2), ddlEOpenAI.Choose(2), ToggleExplanationControls()
    CPDesktopToggleLayout()
    DesktopAssert(DesktopShown(TestExplanationLegacyText) && DesktopShown(ddlEGem) && DesktopShown(ddlEOpenAI), "Classic mode restores the complete original Explanation form")
    DesktopAssert(btnExplainNow.Text = "Explain last jp. Text" && saveLibraryChk.Text = "Save explanations to Study Library", "Classic mode restores original Explanation labels")
    DesktopAssert(!DesktopShown(expPage["models"]) && !CPDesktopIsCombo(ddlEPr.Hwnd), "Classic Explanation removes modern controls and dropdown styling")
    CPDesktopToggleLayout()
    DesktopAssert(ddlEPr.Text = "detailed_grammar" && ddlEOpenAI.Text = "gpt-explanation-alternative" && DesktopShown(ddlEOpenAI) && !DesktopShown(ddlEGem), "Explanation selections survive both layout changes")
    DesktopAssert(saveLibraryChk.Value && !saveLibraryScreenshotsChk.Value && saveExplChk.Value, "Explanation saving preferences survive layout changes")
    DesktopAssert(otherChoices = ddlProv.Text "|" ddlPrompt.Text "|" ddlAProv.Text "|" ddlAudioTarget.Text, "Explanation changes never alter translation or audio choices")
    ddlEProv.Choose(1), ddlEOpenAI.Choose(1), ddlEPr.Choose(1), ToggleExplanationControls()
    tab.Value := 1
    ui.Show("NA x-9000 y-9000 w1400 h820")
    CPDesktopLayout(ui, 0, 1400, 820)
}

DesktopBuildOrganizeFixtures() {
    global ui, tab, iniPath, TestLegacyText, TestLegacyCombo, DesktopSettingsActions
    global ddlGameProfile, ddlStartupOverlays, txtGameProfileState
    global btnGameProfileAdd, btnGameProfileSave, btnGameProfileApply, btnGameProfileDelete
    global chkUseTerminologyOverrides, ddlENG, ddlJPG, btnENG_Edit, btnENG_New, btnENG_Del, btnJPG_Edit, btnJPG_New, btnJPG_Del
    global hotkeyActions, hotkeyLabels, hkEdits, hkBtnChg, hkBtnDis, hkBtnDef, hkConflictText, rbControlsKeyboard, rbControlsController
    global CPControllerBindingEdits, CPControllerAssignButtons, CPControllerDisableButtons
    global cbControllerInputsEnabled, cbControllerDpadNavigationEnabled, txtControllerStatus, txtControllerDpadNote
    global cbApiInApp, eGemini, eOpenAI, btnSaveEnv, btnDelEnv, btnOpenEnvVars, btnAbout
    global ePython, eOverlay, eImg, eAudio, eExplain, bPy, bOvSel, bImgSel, bAud, bExplainSel, btnSavePaths, cbDirectModelOutput, cbDebug
    DesktopSettingsActions := Map()
    tab.UseTab(6)
    fixtureBefore := CPDesktopExistingControls()
    TestLegacyText := ui.AddText("x24 y60 w360 h30", "Classic terminology help")
    TestLegacyCombo := ui.AddDropDownList("x120 y120 w360", ["Default"])
    chkUseTerminologyOverrides := ui.AddCheckbox(, "Use terminology overrides")
    chkUseTerminologyOverrides.Value := 1
    ddlENG := ui.AddDropDownList("w280", ["Local terms", "Local alternate"]), ddlENG.Choose(1)
    ddlJPG := ui.AddDropDownList("w280", ["Japanese terms", "Japanese alternate"]), ddlJPG.Choose(1)
    btnENG_Edit := ui.AddButton(, "Manage Entries..."), btnENG_New := ui.AddButton(, "New Profile..."), btnENG_Del := ui.AddButton(, "Delete Profile...")
    btnJPG_Edit := ui.AddButton(, "Manage Entries..."), btnJPG_New := ui.AddButton(, "New Profile..."), btnJPG_Del := ui.AddButton(, "Delete Profile...")
    CPDesktopCaptureOrganizeControls(6, fixtureBefore)
    tab.UseTab(7)
    fixtureBefore := CPDesktopExistingControls()
    ui.AddText(, "Classic profiles introduction")
    ddlGameProfile := ui.AddDropDownList("w330", ["Demo profile", "Alternate profile"]), ddlGameProfile.Choose(1)
    ddlStartupOverlays := ui.AddDropDownList("w330", ["None", "Translator only", "Explainer only", "Translator and Explainer"]), ddlStartupOverlays.Choose(2)
    btnGameProfileAdd := ui.AddButton(, "Add..."), btnGameProfileSave := ui.AddButton(, "Save Current")
    btnGameProfileApply := ui.AddButton(, "Apply"), btnGameProfileDelete := ui.AddButton(, "Delete")
    txtGameProfileState := ui.AddText("w760", "Selected: Demo profile")
    ddlGameProfile.OnEvent("Change", GameProfileUpdateSummary)
    CPDesktopCaptureOrganizeControls(7, fixtureBefore)
    tab.UseTab(8)
    fixtureBefore := CPDesktopExistingControls()
    rbControlsKeyboard := ui.AddRadio("Group +0x1000", "Keyboard inputs"), rbControlsController := ui.AddRadio("+0x1000", "Controller inputs")
    hotkeyActions := ["screenshot_translate", "explain_last_translation", "hide_show_translator", "hide_show_explainer", "hide_show_control_panel", "take_screenshot", "screenshot_translation", "launch_explainer_request", "recapture_region", "start_stop_audio"]
    hotkeyLabels := Map(), hkEdits := Map(), hkBtnChg := Map(), hkBtnDis := Map(), hkBtnDef := Map()
    fixtureLabels := ["Screenshot + Translate", "Explain last translation", "Show/Hide Translator", "Show/Hide Explainer", "Show/Hide Control Panel", "Take Screenshot", "Translate Screenshots", "Launch Explainer + Req.", "Recapture Region", "Audio Translation On/Off"]
    for index, action in hotkeyActions {
        hotkeyLabels[action] := fixtureLabels[index]
        CPAddControlsViewControl("keyboard", ui.AddText(, fixtureLabels[index]))
        hkEdits[action] := CPAddControlsViewControl("keyboard", ui.AddEdit("w240 ReadOnly", "^+" index))
        hkBtnChg[action] := CPAddControlsViewControl("keyboard", ui.AddButton(, "Change..."))
        hkBtnDis[action] := CPAddControlsViewControl("keyboard", ui.AddButton(, "Disable"))
        hkBtnDef[action] := CPAddControlsViewControl("keyboard", ui.AddButton(, "Default"))
        CPAddControlsViewControl("controller", ui.AddText(, fixtureLabels[index]))
        CPControllerBindingEdits[action] := CPAddControlsViewControl("controller", ui.AddEdit("w210 ReadOnly", "Disabled"))
        CPControllerAssignButtons[action] := CPAddControlsViewControl("controller", ui.AddButton(, "Assign"))
        CPControllerDisableButtons[action] := CPAddControlsViewControl("controller", ui.AddButton(, "Disable"))
        hkBtnChg[action].OnEvent("Click", DesktopSettingsClicked.Bind("keyboard." action))
        CPControllerAssignButtons[action].OnEvent("Click", DesktopSettingsClicked.Bind("controller." action))
    }
    hkConflictText := CPAddControlsViewControl("keyboard", ui.AddText("w800", ""))
    txtControllerStatus := CPAddControlsViewControl("controller", ui.AddText("w350", "Bindings off; navigation remains active."))
    cbControllerInputsEnabled := CPAddControlsViewControl("controller", ui.AddCheckbox(, "Enable direct controller action bindings"))
    cbControllerDpadNavigationEnabled := CPAddControlsViewControl("controller", ui.AddCheckbox(, "Use D-pad for control panel navigation"))
    cbControllerDpadNavigationEnabled.Value := 1
    txtControllerDpadNote := CPAddControlsViewControl("controller", ui.AddText("w350", "Classic controller help"))
    CPSetControlsView("keyboard", false)
    CPDesktopCaptureOrganizeControls(8, fixtureBefore)
    tab.UseTab(9)
    fixtureBefore := CPDesktopExistingControls()
    ui.AddText(, "Classic API help")
    cbApiInApp := ui.AddCheckbox(, "Enter API Keys in JRPG Translator (.env)")
    eGemini := ui.AddEdit("w420 Password", "synthetic-gemini"), eOpenAI := ui.AddEdit("w420 Password", "synthetic-openai")
    btnSaveEnv := ui.AddButton(, "Save Keys"), btnDelEnv := ui.AddButton(, "Delete .env")
    btnOpenEnvVars := ui.AddButton(, "Open Windows Environment Variables..."), btnAbout := ui.AddButton(, "About...")
    eGemini.Enabled := false, eOpenAI.Enabled := false, btnSaveEnv.Enabled := false, btnDelEnv.Enabled := false
    CPDesktopCaptureOrganizeControls(9, fixtureBefore)
    tab.UseTab(10)
    fixtureBefore := CPDesktopExistingControls()
    ui.AddText(, "Classic Paths help")
    ePython := ui.AddEdit("w560", "C:\Synthetic\python.exe"), eAudio := ui.AddEdit("w560", "C:\Synthetic\audio.py")
    eOverlay := ui.AddEdit("w560", "C:\Synthetic\overlay.exe"), eImg := ui.AddEdit("w560", "C:\Synthetic\image.py"), eExplain := ui.AddEdit("w560", "C:\Synthetic\explain.py")
    bPy := ui.AddButton(, "Browse"), bAud := ui.AddButton(, "Browse"), bOvSel := ui.AddButton(, "Browse"), bImgSel := ui.AddButton(, "Browse"), bExplainSel := ui.AddButton(, "Browse")
    btnSavePaths := ui.AddButton(, "Save paths"), btnSavePaths.Enabled := false
    cbDirectModelOutput := ui.AddCheckbox(, "Direct model output"), cbDebug := ui.AddCheckbox(, "Debug mode")
    CPDesktopCaptureOrganizeControls(10, fixtureBefore)
    for spec in [[btnGameProfileApply, "profile.apply"], [btnGameProfileSave, "profile.save"], [btnGameProfileAdd, "profile.add"], [btnGameProfileDelete, "profile.delete"],
        [btnENG_Edit, "terms.local"], [btnJPG_Edit, "terms.jp"], [bPy, "paths.python"], [btnSaveEnv, "keys.save"]]
        spec[1].OnEvent("Click", DesktopSettingsClicked.Bind(spec[2]))
}

DesktopSettingsClicked(key, *) {
    global DesktopSettingsActions
    DesktopSettingsActions[key] := DesktopSettingsActions.Get(key, 0) + 1
}

DesktopTestOrganizePages() {
    global ui, tab, CPDesktop, CPDesktopOrganizeLegacy, CPTabVisiblePages, CPCanvasScrollMaxY, CPControlsCurrentView
    global ddlGameProfile, ddlStartupOverlays, txtGameProfileState, btnGameProfileApply, btnGameProfileSave, btnGameProfileAdd, btnGameProfileDelete
    global ddlENG, ddlJPG, btnENG_Edit, btnJPG_Edit, chkUseTerminologyOverrides
    global eGemini, eOpenAI, cbApiInApp, btnSaveEnv, btnDelEnv, btnOpenEnvVars
    global hkEdits, hkBtnChg, hotkeyActions, hkConflictText, CPControllerBindingEdits, CPControllerAssignButtons, cbControllerInputsEnabled, cbControllerDpadNavigationEnabled
    global rbControlsKeyboard, rbControlsController, txtControllerDpadNote, DesktopSettingsActions
    global ePython, bPy, btnSavePaths, cbDebug, iniPath
    for page in [6, 7, 8, 9, 10] {
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
                DesktopAssert(DesktopShown(p["controlsTab"]) && DesktopShown(p["pathsTab"]), "Settings sections are directly reachable")
                DesktopAssert(DesktopShown(p["classic"]), "Settings retains the reversible classic option")
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
        CPDesktopToggleLayout()
        DesktopAssert(!CPDesktopActive(), "Classic layout is available on every organize page")
        for key, ctrl in CPDesktop["organizePages"][page]
            DesktopAssert(!DesktopShown(ctrl), "Modern organize tools are hidden in classic layout")
        if page = 8
            DesktopAssert(DesktopShown(rbControlsKeyboard) && !DesktopShown(txtControllerDpadNote), "Classic controls restores only its active input view")
        CPDesktopToggleLayout()
        DesktopAssert(CPDesktopActive() && tab.Value = page, "Modern round trip retains the selected page")
    }
    CPDesktopNavigate(7)
    ddlGameProfile.Choose(2), GameProfileUpdateSummary()
    DesktopAssert(InStr(txtGameProfileState.Text, "Alternate profile") && !DesktopSettingsActions.Has("profile.apply"), "Selecting a profile does not apply it")
    ddlStartupOverlays.Choose(4)
    for spec in [[btnGameProfileApply, "profile.apply"], [btnGameProfileSave, "profile.save"], [btnENG_Edit, "terms.local"], [btnJPG_Edit, "terms.jp"]] {
        SendMessage(0xF5, 0, 0, spec[1].Hwnd)
        Sleep(10)
        DesktopAssert(DesktopSettingsActions.Get(spec[2], 0) = 1, "Existing action callback remains wired: " spec[2])
    }
    ddlENG.Choose(2), ddlJPG.Choose(1)
    CPDesktopNavigate(8)
    SendMessage(0xF5, 0, 0, CPDesktop["organizePages"][8]["gamepad"].Hwnd)
    Sleep(10)
    DesktopAssert(CPControlsCurrentView = "controller" && rbControlsController.Value, "Modern input selector updates the original controller view")
    for action in hotkeyActions
        DesktopAssert(DesktopShown(CPControllerBindingEdits[action]) && !DesktopShown(hkEdits[action]), "Only controller bindings are shown")
    DesktopAssert(DesktopShown(cbControllerInputsEnabled) && DesktopShown(cbControllerDpadNavigationEnabled), "Controller options remain reachable")
    SendMessage(0xF5, 0, 0, CPControllerAssignButtons[hotkeyActions[1]].Hwnd)
    Sleep(10)
    DesktopAssert(DesktopSettingsActions.Get("controller." hotkeyActions[1], 0) = 1, "Controller Assign keeps its callback")
    ui.Show("NA x-9000 y-9000 w820 h560"), CPDesktopLayout(ui, 0, 820, 560)
    ui.GetPos(,, &shotW, &shotH)
    TestCapture("desktop-controller-820.png", Round(shotW * dpi), Round(shotH * dpi))
    CPCanvasScrollTo(0, CPCanvasScrollMaxY)
    TestCapture("desktop-controller-820-scrolled.png", Round(shotW * dpi), Round(shotH * dpi))
    CPDesktopNavigate(1), CPDesktopSettingsMenu()
    DesktopAssert(tab.Value = 8 && CPControlsCurrentView = "controller", "Settings returns directly to its last section and input view")
    CPSetControlsView("keyboard", false)
    SendMessage(0xF5, 0, 0, hkBtnChg[hotkeyActions[1]].Hwnd)
    Sleep(10)
    DesktopAssert(DesktopSettingsActions.Get("keyboard." hotkeyActions[1], 0) = 1, "Keyboard Change keeps its callback")
    previousBinding := hkEdits[hotkeyActions[1]].Value
    hkEdits[hotkeyActions[1]].Value := hkEdits[hotkeyActions[2]].Value
    DesktopAssert(Hotkeys_ShowConflicts() = 1, "Original hotkey conflict detection remains active")
    DesktopAssert(CPDesktop["bindingConflicts"].Has(hkEdits[hotkeyActions[1]].Hwnd) && CPDesktop["bindingConflicts"].Has(hkEdits[hotkeyActions[2]].Hwnd), "Both conflicting fields retain visual emphasis")
    hkConflictText.GetPos(,, &bannerW, &bannerH)
    dpi := GetWindowDPI(ui.Hwnd) / 96
    DesktopAssert(bannerH * dpi >= CPMeasureWrappedTextHeight(hkConflictText, bannerW * dpi), "Conflict banner expands without clipping")
    hkEdits[hotkeyActions[1]].Value := previousBinding
    DesktopAssert(Hotkeys_ShowConflicts() = 0 && CPDesktop["bindingConflicts"].Count = 0, "Resolving a duplicate clears its modern emphasis")
    CPDesktopNavigate(9)
    for ctrl in [eGemini, eOpenAI]
        DesktopAssert(SendMessage(0xD2, 0, 0, ctrl.Hwnd) != 0 && !ctrl.Enabled, "API keys remain masked and disabled when in-app entry is off")
    DesktopAssert(!btnSaveEnv.Enabled && !btnDelEnv.Enabled && eGemini.Value = "synthetic-gemini", "API layout preserves values and dirty/enabled state")
    CPDesktopNavigate(10)
    ePython.Value := "C:\Synthetic\unsaved-python.exe"
    cbDebug.Value := 1
    CPDesktopToggleLayout(), CPDesktopToggleLayout()
    DesktopAssert(ePython.Value = "C:\Synthetic\unsaved-python.exe" && !btnSavePaths.Enabled && cbDebug.Value, "Path edits and diagnostic preferences survive layout switches without saving")
    DesktopAssert(ddlGameProfile.Text = "Alternate profile" && ddlStartupOverlays.Value = 4 && ddlENG.Value = 2 && ddlJPG.Value = 1, "Independent profile, startup and glossary selections are preserved")
    CPTabVisiblePages.Pop()
    CPDesktopNavigate(9)
    DesktopAssert(!DesktopShown(CPDesktop["organizePages"][9]["pathsTab"]), "Hidden Paths stays hidden in modern settings")
    CPDesktop["lastSettingsPage"] := 10
    CPDesktopSettingsMenu()
    DesktopAssert(tab.Value = 8, "Settings falls back safely when Paths is unavailable")
    CPTabVisiblePages.Push(10)
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
    for spec in [[6, "termsTab"], [9, "apiTab"], [10, "pathsTab"], [8, "controlsTab"]] {
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
        if spec[1] = 10 {
            ui.GetPos(,, &focusCaptureW, &focusCaptureH)
            dpi := GetWindowDPI(ui.Hwnd) / 96
            TestCapture("desktop-settings-paths-focus.png", Round(focusCaptureW * dpi), Round(focusCaptureH * dpi))
        }
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
    CPDesktopNavigate(10)
    Sleep(30)
    DesktopAssert(!CPCanvasPendingScrollValid && CPCanvasScrollY = 0, "A queued wheel frame never scrolls the next page")
    CPDesktopNavigate(3)
    DesktopAssert(!CPCanvasAcceptsWheel(slTrans.Hwnd, DesktopWheelPoint(slTrans)), "Opacity slider retains its own wheel behavior")
    CPDesktopNavigate(6)
    CPDesktopToggleLayout()
    CPCanvasScrollTo(0, 0)
    Critical "On"
    try {
        SendMessage(0x20A, (-120 & 0xFFFF) << 16, DesktopWheelPoint(ddlENG), ddlENG.Hwnd)
        DesktopAssert(CPCanvasPendingScrollValid, "Classic layout uses the same non-hotkey wheel routing")
    } finally {
        Critical "Off"
        CPCanvasCancelQueuedScroll()
        CPDesktopToggleLayout()
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
        for page in [6, 8, 10, 3] {
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
        for page in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10] {
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
            s["libraryDdl"].Add(["Test library", "Another library"]), s["libraryDdl"].Choose(1)
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
    s["sentences"] := [sample], s["vocabulary"] := [sample], s["allSentences"] := [sample], s["allVocabulary"] := [sample]
    s["baseStatus"] := "1 sentence · 1 vocabulary entry", s["snapshotThrough"] := "2026-09-08"
    s["sentenceList"].Add("Select Focus", "Recommended", "2026-09-08", "Kabuki Den", "お城には行けました？", "1")
    s["vocabularyList"].Add("Select Focus", "Recommended", "お城（おしろ）", "castle; polite form", "2", "Kabuki Den", "2026-09-08")
    DesktopAssert(s.Has("desktop") && s["tabs"].Type = "Tab2", "Candidate manager uses desktop tab container")
    for size in [[960, 760], [1040, 800], [1400, 900]] {
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
        }
        dpi := GetWindowDPI(hwnd) / 96
        TestDesktopStudyCapture(g, "desktop-anki-candidates-" w ".png", Round(w * dpi), Round(h * dpi))
    }
    StudyCandidatesSetRecommendationActivity(s, true, "Test provider", Map("sentences", 1, "vocabulary", 1))
    DesktopAssert(DesktopShown(s["progressText"]) && DesktopShown(s["progressBar"]) && !s["recommendButton"].Enabled, "Candidate assessment progress retains busy state")
    StudyCandidatesSetRecommendationActivity(s, false)
    DesktopAssert(!DesktopShown(s["progressText"]) && s["recommendButton"].Enabled, "Candidate assessment progress resets")
    sample["reason"] := ""
    Loop 60
        sample["reason"] .= "Long assessment detail remains readable in the scrollable area. "
    StudyCandidatesUpdateActions(s)
    DesktopAssert(!DesktopShown(s["desktopAiSummary"]) && DesktopShown(s["aiStatus"]), "Long assessment uses the scrollable detail area")
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
    DesktopTestStudyToolSizes(version, [[820, 700], [920, 740], [1200, 900]])
    for provider in ["openai", "gemini"] {
        c["provider"].Choose(provider = "gemini" ? 1 : 2)
        StudyReaderNewVersionProviderChanged(version)
        for key in ["gemini", "openai"]
            DesktopAssert(c[key].Visible = (key = provider) && c[key].Enabled = (key = provider)
                && c[key "Label"].Visible = (key = provider), "Only selected provider model is visible: " key)
    }
    for mode in ["name", "edit"] {
        prompt := StudyReaderPromptDialogCreate(version, mode, mode = "edit" ? "Explain Japanese. 日本語 {jp}" : "", "fixture")
        DesktopTestStudyToolSizes(prompt, mode = "edit" ? [[820, 650], [1000, 780]] : [[680, 450], [760, 480]])
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

TestSharedChoicePopup(ownerHwnd, choices) {
    global controlDarkMode
    ; @SHARED_CHOICE_POPUP_CONTROLS@
    return cpPopupState
}

TestSharedContextPopup(ownerHwnd, items) {
    global controlDarkMode
    ; @SHARED_CONTEXT_POPUP_CONTROLS@
    return cpPopupState
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
    CPDesktop["modern"] := false
    DesktopAssert(CPDialogPresentation(ui.Hwnd) = "classic", "Classic layout keeps legacy message presentation")
    CPDesktop["modern"] := true
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
    TestSharedMessageRoute(child.Hwnd, "desktop", "yesno", "warning")
    TestSharedMessageRoute(nestedFullscreen.Hwnd, "fullscreen", "yesno", "error")
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
    for provider in ["gemini", "openai"] {
        for purpose in ["screenshot", "audio", "explanation"] {
            s := CPModelDialogCreate(reader["gui"].Hwnd, provider, purpose, "source")
            DesktopAssert(InStr(StrLower(s["gui"].Title), provider), "Model source identifies provider")
            DesktopAssert(InStr(StrLower(s["desktop"]["chrome"]["subtitle"].Text), purpose), "Model source identifies purpose")
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
            for action in ["accept", "cancel", "escape", "close"]
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
    DesktopAssert(result = (action = "accept" ? mode = "source" ? "manual" : "model-alpha" : mode = "source" ? "cancel" : ""), "Model modal result matches action")
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

TestSharedMessageRoute(owner, presentation, buttons, icon) {
    global TestSharedMessageCaptured
    TestSharedMessageCaptured := false
    priorEnabled := DllCall("user32\IsWindowEnabled", "ptr", owner)
    timer := TestSharedMessageClose.Bind(presentation, buttons)
    SetTimer(timer, 80)
    try result := CPAdaptiveOwnedMessage(owner, "Synthetic shared notice. 日本語 & names remain literal.", "Shared message test", buttons, icon)
    finally SetTimer(timer, 0)
    DesktopAssert(TestSharedMessageCaptured, "Shared message used expected presentation: " presentation)
    DesktopAssert(result = (buttons = "yesno" ? "No" : "OK"), "Shared message keeps safe default result")
    DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", owner) = priorEnabled, "Shared message restores owner enabled state")
}

TestSharedMessageClose(presentation, buttons) {
    global TestSharedMessageCaptured
    if presentation = "desktop" {
        for hwnd, d in StudyDesktopRegistry() {
            if d["kind"] != "message" || d["state"]["gui"].Title != "Shared message test" || d["width"] < 1
                continue
            s := d["state"]
            s["gui"].GetClientPos(,, &w, &h)
            dpi := GetWindowDPI(s["gui"].Hwnd) / 96
            TestDesktopStudyCapture(s["gui"], "shared-message-desktop-" buttons ".png", Round(w * dpi), Round(h * dpi))
            TestSharedMessageCaptured := true
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
            TestDesktopStudyCapture(g, "shared-message-fullscreen-" buttons ".png", w, h)
            TestSharedMessageCaptured := true
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
