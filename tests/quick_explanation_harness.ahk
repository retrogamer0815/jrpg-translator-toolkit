; Real controls and persistence, isolated offscreen fixture. No model calls.
TestQuickExplanationUI() {
    global ui, tab, CPDesktop, iniPath, ddlEProv, ddlEGem, ddlEPr
    global saveLibraryChk, saveExplChk, btnExplainNow, btnOpenStudyLibrary, DesktopExplanationClicks
    global model_gemini_explain, TestQuickClicks
    TestQuickClicks := 0
    page := CPDesktop["explanationPage"]
    page["quickGenerate"].OnEvent("Click", QuickExplainNow, 0)
    page["quickGenerate"].OnEvent("Click", TestQuickClicked)
    CPDesktopNavigate(4)
    settings := CPQuickExplanationSettings()
    DesktopAssert(settings["useDetailedModel"] = 1, "Quick defaults to detailed model")
    DesktopAssert(page["quickPrompt"].Text = "default_en", "Quick has a separate bundled prompt")
    DesktopAssert(InStr(FileRead(CPQuickExplanationPromptPath("default_en")), "{jp}"), "Quick default accepts Japanese source")
    detailedBefore := ddlEProv.Text "|" ddlEGem.Text "|" ddlEPr.Text "|" saveLibraryChk.Value "|" saveExplChk.Value
    for width in [820, 900, 1120, 1400] {
        ui.Show("NA x-9000 y-9000 w" width " h900")
        CPDesktopLayout(ui, 0, width, 900)
        for expanded in [true, false] {
            CPDesktop["quickExplanationOptionsOpen"] := expanded
            CPDesktopRelayout()
            DesktopAssert(DesktopShown(page["quickUseDetailed"]) = expanded, "Quick model settings expand/collapse")
            page["quickHelp"].GetPos(,, &helpW, &helpH)
            scale := GetWindowDPI(ui.Hwnd) / 96
            DesktopAssert(helpH * scale >= CPMeasureWrappedTextHeight(page["quickHelp"], helpW * scale), "Quick description wraps at width " width)
            page["quickPrompt"].GetPos(&px, &py, &pw, &ph)
            page["quickPrompts"].GetPos(&mx, &my, &mw, &mh)
            DesktopAssert(px + pw < mx, "Quick prompt does not cover Manage")
            btnExplainNow.GetPos(&dx, &dy, &dw, &dh)
            page["quickGenerate"].GetPos(&qx, &qy, &qw, &qh)
            btnOpenStudyLibrary.GetPos(&lx, &ly, &lw, &lh)
            DesktopAssert(dx + dw < qx && dy = qy, "Detailed and Quick actions are adjacent")
            DesktopAssert(qx + qw < lx || qy + qh < ly, "Library action wraps without overlap")
            DesktopAssert(ArrayIndexOf(CPDesktop["pageStops"], page["quickGenerate"]) > 0, "Quick action is reachable in page navigation")
        }
    }
    CPQuickExplanationToggleOptions()
    DesktopAssert(!page["quickProvider"].Enabled, "Inherited provider cannot be edited accidentally")
    page["quickUseDetailed"].Value := 0
    page["quickProvider"].Choose(2)
    page["quickOpenAI"].Choose(2)
    CPQuickExplanationChanged()
    saved := CPQuickExplanationSettings()
    DesktopAssert(!saved["useDetailedModel"] && saved["provider"] = "openai" && saved["openAIModel"] = page["quickOpenAI"].Text, "Independent model persists")
    DesktopAssert(DesktopShown(page["quickOpenAI"]) && !DesktopShown(page["quickGemini"]) && page["quickProvider"].Enabled, "Quick provider selects matching model control")
    DesktopAssert(detailedBefore = ddlEProv.Text "|" ddlEGem.Text "|" ddlEPr.Text "|" saveLibraryChk.Value "|" saveExplChk.Value, "Quick settings do not change Detailed or saving")
    model_gemini_explain.Push("new-quick-model")
    CPQuickExplanationToggleOptions()
    page["quickGemini"].Choose("new-quick-model")
    DesktopAssert(page["quickGemini"].Text = "new-quick-model", "Quick refreshes newly managed model list")
    FileAppend("Quick custom: {jp}", CPQuickExplanationPromptPath("custom"), "UTF-8")
    CPQuickExplanationRefreshPrompts("custom")
    DesktopAssert(CPQuickExplanationSettings()["prompt"] = "custom" && ddlEPr.Text = StrSplit(detailedBefore, "|")[3], "Quick prompt selection persists independently")
    DesktopAssert(CPQuickExplanationPromptPath("../invalid") = "", "Prompt name cannot escape prompt directory")
    priorDetailedClicks := DesktopExplanationClicks
    SendMessage(0xF5, 0, 0, page["quickGenerate"].Hwnd)
    Sleep(20)
    DesktopAssert(TestQuickClicks = 1 && DesktopExplanationClicks = priorDetailedClicks, "Quick click invokes only Quick action")
    SendMessage(0xF5, 0, 0, btnExplainNow.Hwnd)
    Sleep(20)
    DesktopAssert(DesktopExplanationClicks = priorDetailedClicks + 1 && TestQuickClicks = 1, "Detailed click remains separate")
    ui.Show("NA x-9000 y-9000 w1120 h1100")
    CPDesktop["quickExplanationOptionsOpen"] := true
    CPDesktopLayout(ui, 0, 1120, 1100)
    CPCanvasScrollTo(0, 0)
    ui.GetPos(,, &quickCaptureW, &quickCaptureH)
    scale := GetWindowDPI(ui.Hwnd) / 96
    TestCapture("quick-explanation-page.png", Round(quickCaptureW * scale), Round(quickCaptureH * scale))
    CPDesktopNavigate(1)
    DesktopAssert(!DesktopShown(page["quickGenerate"]) && !DesktopShown(page["quickPrompt"]), "Quick controls stay on Explanation page")
}

TestQuickClicked(*) {
    global TestQuickClicks
    TestQuickClicks += 1
}
