#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
global TestAssertions := 0, TestRenders := 0, TestStatusHides := 0, TestScrolls := 0, TestTimerTicks := 0
global AudioTxt := "", OcrTxt := "", OcrDoneTxt := "", ExplainerTxt := "", ExplainerDoneTxt := ""
global __LastAudioRaw := "", __LastOcrRaw := "", __LastOcrDoneRaw := ""
global __LastExplainRaw := "", __LastExplainDoneRaw := "", __OcrText := "", __AudioText := ""
global __TranslationRequestId := "", __TranslationPid := 0, __TranslationStartedAt := 0, __TranslationExitedAt := 0
global iniPath := A_ScriptDir "\synthetic-settings.ini", ControlIni := iniPath, __EXPLAIN_MODE := false
global BoundsIni := A_ScriptDir "\synthetic-bounds.ini", MinW := 200, MinH := 120, Overlay := 0
OnError(TestUnhandled)
try {
    TestClearing()
    TestPolling()
    TestLivePollingTimers()
    TestSettings()
    TestClosedWindows()
    FileAppend("PASS: " TestAssertions " lifecycle regression assertions (real file locks and synthetic settings).`n", "*")
    ExitApp(0)
} catch as testError {
    TestUnhandled(testError)
}
TestUnhandled(err, *) {
    FileAppend("FAIL: " err.Message " | " err.Extra " | line " err.Line "`n", "*")
    ExitApp(1)
}
TestAssert(ok, message) {
    global TestAssertions
    TestAssertions += 1
    if !ok
        throw Error(message)
}
TestWrite(path, text) {
    stream := FileOpen(path, "w", "UTF-8")
    stream.Write(text)
    stream.Close()
}
TestLock(path) {
    handle := DllCall("CreateFileW", "wstr", path, "uint", 0xC0000000,
        "uint", 0, "ptr", 0, "uint", 3, "uint", 0x80, "ptr", 0, "ptr")
    if handle = -1
        throw OSError()
    return handle
}
TestSetPaths(name) {
    global AudioTxt, OcrTxt, ExplainerTxt, OcrDoneTxt, ExplainerDoneTxt
    path := A_ScriptDir "\" name ".txt"
    AudioTxt := path, OcrTxt := path, ExplainerTxt := path
    OcrDoneTxt := path ".done", ExplainerDoneTxt := OcrDoneTxt
    return path
}
TestNoClearScratch(path) {
    count := 0
    Loop Files path ".clear-*.tmp", "F"
        count += 1
    TestAssert(count = 0, "Private clear scratch is removed after success or failure")
}
TestClearing() {
    global ExplainerDoneTxt
    for spec in [["audio", ClearAudioFile], ["ocr", ClearOcrFile], ["explainer", ClearExplainerFile]] {
        path := TestSetPaths("clear-" spec[1])
        TestWrite(path ".tmp", "WORKER TEMP MUST NOT BE USED")
        TestWrite(path ".done.tmp", "STALE DONE")
        Loop 5 {
            TestWrite(path, "last valid result")
            TestWrite(path ".done", "old completion")
            locked := TestLock(path)
            try TestAssert(!spec[2].Call(), "Locked clear returns failure, not an unassigned-variable error")
            finally DllCall("CloseHandle", "ptr", locked)
            TestAssert(FileRead(path, "UTF-8") = "last valid result", "Failed clear does not damage existing output")
            TestNoClearScratch(path)
            TestAssert(spec[2].Call(), "A subsequent clear succeeds after the lock is released")
            TestAssert(FileRead(path, "UTF-8") = "", "Clear really creates empty content")
            TestAssert(FileRead(path ".tmp", "UTF-8") = "WORKER TEMP MUST NOT BE USED", "Worker temporary file is untouched")
            TestAssert(FileRead(path ".done.tmp", "UTF-8") = "STALE DONE", "Worker completion temporary file is untouched")
            TestNoClearScratch(path)
            if spec[1] = "explainer"
                TestAssert(FileRead(path ".done", "UTF-8") = "", "Explainer completion marker also clears")
        }
        FileDelete(path)
        TestAssert(spec[2].Call(), "Clear creates a missing output file")
        TestAssert(FileRead(path, "UTF-8") = "", "Created output is empty")
    }
    path := TestSetPaths("clear-done-lock")
    TestWrite(path, "explanation"), TestWrite(ExplainerDoneTxt, "done")
    locked := TestLock(ExplainerDoneTxt)
    try TestAssert(!ClearExplainerFile(), "Locked completion marker is a recoverable clear failure")
    finally DllCall("CloseHandle", "ptr", locked)
    TestAssert(ClearExplainerFile(), "Completion marker clear can be retried")
    TestNoClearScratch(ExplainerDoneTxt)
    TestAssert(!ClearOverlayOutputFile(A_ScriptDir "\missing-directory\out.txt"), "Creation failure returns false safely")
}
TestResetPollState(kind, pending := "") {
    global __LastAudioRaw, __LastOcrRaw, __LastOcrDoneRaw, __LastExplainRaw, __LastExplainDoneRaw
    global __OcrText, __AudioText, __TranslationRequestId, __TranslationPid, __TranslationStartedAt, __TranslationExitedAt
    global TestRenders, TestStatusHides, TestScrolls
    __LastAudioRaw := "old", __LastOcrRaw := "old", __LastExplainRaw := "old"
    __LastOcrDoneRaw := "old-done", __LastExplainDoneRaw := "old-done"
    __OcrText := "old OCR", __AudioText := "old audio"
    __TranslationRequestId := pending, __TranslationPid := 123, __TranslationStartedAt := 456, __TranslationExitedAt := 789
    TestRenders := 0, TestStatusHides := 0, TestScrolls := 0
}
TestAssertPollUntouched() {
    global __LastAudioRaw, __LastOcrRaw, __LastOcrDoneRaw, __LastExplainRaw, __LastExplainDoneRaw
    global __OcrText, __AudioText, TestRenders, TestStatusHides, TestScrolls
    TestAssert(__OcrText = "old OCR" && __AudioText = "old audio", "Failed read retains both text channels")
    TestAssert(__LastAudioRaw = "old" && __LastOcrRaw = "old" && __LastExplainRaw = "old", "Failed read retains content caches")
    TestAssert(__LastOcrDoneRaw = "old-done" && __LastExplainDoneRaw = "old-done", "Failed read retains completion caches")
    TestAssert(TestRenders = 0 && TestStatusHides = 0 && TestScrolls = 0, "Failed read does not render, scroll, or finish busy status")
}
TestPolling() {
    global __OcrText, __AudioText, __LastOcrRaw, __LastOcrDoneRaw
    global __TranslationRequestId, __TranslationPid, __TranslationStartedAt, __TranslationExitedAt
    global TestRenders, TestStatusHides
    for spec in [["audio", PollAudioSubtitle], ["ocr", PollOcrFile], ["explainer", PollExplainerFile]] {
        path := TestSetPaths("poll-" spec[1])
        for lockMarker in [false, true] {
            if lockMarker && spec[1] = "audio"
                continue
            TestWrite(path, "new text"), TestWrite(path ".done", "new-done")
            TestResetPollState(spec[1])
            locked := TestLock(lockMarker ? path ".done" : path)
            try {
                Loop 10 {
                    spec[2].Call()
                    TestAssertPollUntouched()
                }
            } finally DllCall("CloseHandle", "ptr", locked)
            spec[2].Call()
            TestAssert(TestRenders = 1, "Next successful poll renders once")
            TestAssert((spec[1] = "audio" ? __AudioText : __OcrText) = "new text", "Poll recovers new output after unlock")
            spec[2].Call()
            TestAssert(TestRenders = 1, "Unchanged output is not rendered twice")
            TestWrite(path, "")
            spec[2].Call()
            TestAssert((spec[1] = "audio" ? __AudioText : __OcrText) = "", "Successfully read empty output still clears")
        }
        TestResetPollState(spec[1])
        FileDelete(path)
        spec[2].Call()
        TestAssertPollUntouched()
        TestWrite(path, "without completion marker")
        FileDelete(path ".done")
        spec[2].Call()
        TestAssert(TestRenders = 1, "Missing optional completion marker permits idle text updates")
    }
    path := TestSetPaths("poll-pending")
    TestWrite(path, "new translation"), TestWrite(path ".done", "new-id")
    TestResetPollState("ocr", "new-id")
    locked := TestLock(path)
    try PollOcrFile()
    finally DllCall("CloseHandle", "ptr", locked)
    TestAssertPollUntouched()
    TestAssert(__TranslationRequestId = "new-id" && __TranslationPid = 123
        && __TranslationStartedAt = 456 && __TranslationExitedAt = 789, "Read failure does not prematurely finish a pending request")
    PollOcrFile()
    TestAssert(__TranslationRequestId = "" && __TranslationPid = 0 && __OcrText = "new translation", "Successful matching result completes the request")
    TestResetPollState("ocr", "different-id")
    PollOcrFile()
    TestAssertPollUntouched()
    TestAssert(__TranslationRequestId = "different-id", "Mismatched completion ID remains ignored")
    TestResetPollState("ocr", "new-id")
    __LastOcrRaw := "new translation"
    PollOcrFile()
    TestAssert(TestRenders = 0 && TestStatusHides = 1 && __TranslationRequestId = "", "Same-text completion finishes without a redundant render")
}
TestMainSettings() {
    defFontSize := 22, defOverlayTrans := 220
    ; @MAIN_SETTINGS@
    return Map("font", fontSize, "explainerFont", fontSize_EW, "opacity", overlayTrans,
        "explainerOpacity", overlayTrans_EW, "maxKB", capMaxKB, "x", ewX, "y", ewY, "w", ewW, "h", ewH)
}
TestLivePollingTimers() {
    global TestTimerTicks, TestRenders
    for spec in [["audio", PollAudioSubtitle], ["ocr", PollOcrFile], ["explainer", PollExplainerFile]] {
        path := TestSetPaths("live-timer-" spec[1])
        TestWrite(path, "new text"), TestWrite(path ".done", "new-done")
        TestResetPollState(spec[1])
        TestTimerTicks := 0
        callback := TestPollTick.Bind(spec[2])
        locked := TestLock(path)
        SetTimer(callback, 15)
        try {
            deadline := A_TickCount + 1000
            while TestTimerTicks < 3 && A_TickCount < deadline
                Sleep(15)
            TestAssert(TestTimerTicks >= 3, "Production polling callbacks run on real timers under a file lock")
            TestAssertPollUntouched()
        } finally {
            SetTimer(callback, 0)
            DllCall("CloseHandle", "ptr", locked)
        }
        SetTimer(callback, 15)
        try {
            deadline := A_TickCount + 1000
            while TestRenders < 1 && A_TickCount < deadline
                Sleep(15)
            TestAssert(TestRenders = 1, "Live polling resumes after a transient lock is released")
        } finally SetTimer(callback, 0)
    }
}
TestPollTick(callback) {
    global TestTimerTicks
    TestTimerTicks += 1
    callback.Call()
}
TestOverlaySettings() {
    ; @OVERLAY_SETTINGS@
    return Map("font", FONT_SIZE, "bold", FONT_BOLD, "maxKB", Cap_MaxKB, "top", IsTop)
}
TestSettings() {
    global iniPath, BoundsIni, __EXPLAIN_MODE
    for entry in [["", 22], ["not-a-number", 22], ["12oops", 22], ["1e999", 22], ["9223372036854775808", 22],
        ["24", 24], [" 25 ", 25], ["24.9", 24], ["0x20", 32], ["0", 6], ["-10", 6], ["999", 128]] {
        IniWrite(entry[1], iniPath, "cfg", "fontSize")
        TestAssert(ReadIniInt(iniPath, "cfg", "fontSize", 22, 6, 128) = entry[2], "Integer reader validates and clamps: " entry[1])
        TestAssert(GameProfileReadInt(iniPath, "cfg", "fontSize", 22, 6, 128) = entry[2], "Profiles use identical validation")
    }
    TestAssert(ReadIniInt(iniPath, "absent", "key", 37) = 37, "Missing key uses default")
    TestAssert(ReadIniInt(iniPath, "absent", "key", "") = "", "Missing optional geometry stays unset")
    IniWrite("9007199254740993", iniPath, "precision", "id")
    TestAssert(ReadIniInt(iniPath, "precision", "id", 0) = 9007199254740993, "Valid integer IDs retain precision")
    for bad in ["", "bad", "1e999"] {
        for section in ["cfg", "cfg_explainer"]
            for key in ["fontSize", "fontBold", "overlayTrans", "winTop"]
                IniWrite(bad, iniPath, section, key)
        IniWrite(bad, iniPath, "capture", "maxKB")
        for key in ["x", "y", "w", "h"]
            IniWrite(bad, iniPath, "explainer_bounds", key)
        for key in ["x", "y", "w", "h", "dpi"]
            IniWrite(bad, BoundsIni, "win", key)
        before := FileRead(iniPath, "UTF-8")
        settings := TestMainSettings()
        TestAssert(settings["font"] = 22 && settings["explainerFont"] = 22, "Main startup tolerates malformed font sizes")
        TestAssert(settings["opacity"] = 220 && settings["explainerOpacity"] = 220, "Malformed opacity uses defaults")
        TestAssert(settings["maxKB"] = 1400, "Malformed capture limit uses default")
        for key in ["x", "y", "w", "h"]
            TestAssert(settings[key] = "", "Invalid saved Explainer geometry stays unset")
        for mode in [false, true] {
            __EXPLAIN_MODE := mode
            settings := TestOverlaySettings()
            TestAssert(settings["font"] = 22 && settings["bold"] = 0 && settings["maxKB"] = 1400, "Overlay startup uses safe numeric defaults")
            TestAssert(settings["top"] = !mode, "Overlay topmost default respects mode")
        }
        bounds := LoadOverlayBounds()
        TestAssert(bounds["x"] = 120 && bounds["y"] = 120 && bounds["w"] = 900 && bounds["h"] = 500, "Malformed overlay bounds use valid defaults")
        TestAssert(FileRead(iniPath, "UTF-8") = before, "Validation does not rewrite saved settings")
    }
    IniWrite(160, iniPath, "cfg", "fontSize"), IniWrite(160, iniPath, "cfg_explainer", "fontSize")
    IniWrite(-1920, iniPath, "explainer_bounds", "x"), IniWrite(-120, iniPath, "explainer_bounds", "y")
    IniWrite(960, iniPath, "explainer_bounds", "w"), IniWrite(720, iniPath, "explainer_bounds", "h")
    settings := TestMainSettings()
    TestAssert(settings["font"] = 128 && settings["explainerFont"] = 160, "Main honors distinct Translator/Explainer font limits")
    TestAssert(settings["x"] = -1920 && settings["y"] = -120 && settings["w"] = 960 && settings["h"] = 720, "Valid negative-monitor bounds are preserved")
    for mode in [false, true] {
        __EXPLAIN_MODE := mode
        TestAssert(TestOverlaySettings()["font"] = (mode ? 160 : 128), "Overlay uses the same per-mode limits")
    }
    IniWrite(-1920, BoundsIni, "win", "x"), IniWrite(-120, BoundsIni, "win", "y")
    IniWrite(-1, BoundsIni, "win", "w"), IniWrite(999999, BoundsIni, "win", "h")
    bounds := LoadOverlayBounds()
    TestAssert(bounds["x"] = -1920 && bounds["y"] = -120 && bounds["w"] = 200 && bounds["h"] = 32767, "Overlay bounds clamp dimensions without losing negative coordinates")
}
TestClosedWindows() {
    fixture := Gui("+ToolWindow", "JRPG private lifecycle test")
    button := fixture.AddButton(, "Copy...")
    fixture.Show("Hide w300 h180")
    state := Map("gui", fixture, "copyButton", button, "copyFeedbackSerial", 1, "sizeCallback", (*) => 0)
    state["closed"] := true
    TestAssert(!StudyDialogFocusIfAlive(state, button), "Late focus skips a closing dialog before native destruction")
    state["closed"] := false
    fixture.Destroy()
    TestAssert(!StudyLibraryStateAlive(state) && !StudyCandidatesGuiAlive(state), "Study guards reject a destroyed GUI")
    StudyReaderRestoreCopyButton(state, 1)
    StudyWindowEnableLiveResize(state)
    TestAssert(!StudyDialogFocusIfAlive(state, button, true), "Late opening focus skips a destroyed control and selection reset")
    TestAssert(true, "Late Study callbacks safely return after destruction")
}
HideOverlayStatus(*) {
    global TestStatusHides
    TestStatusHides += 1
}
RenderCombined(*) {
    global TestRenders
    TestRenders += 1
}
ScrollOutputToTop(*) {
    global TestScrolls
    TestScrolls += 1
}
ScrollOutputToBottom(*) {
    global TestScrolls
    TestScrolls += 1
}
Dbg(*) {
}
DbgRect(*) {
}
EnsureBoundsOnScreen(bounds, *) => bounds
