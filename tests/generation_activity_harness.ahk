; Focused, isolated tests of the real constructors and request lifecycle bodies.
TestGenerationActivity() {
    global GenerationFixture, CPStudyLibraryState, CPStudyReaderState
    library := TestDesktopStudyLibrary(), reader := TestDesktopStudyReader()
    CPStudyLibraryState := 0
    reader["currentGroupId"] := 1, reader["currentVersion"] := 1
    reader["database"] := A_ScriptDir "\fixture.db", reader["outputDir"] := A_ScriptDir
    reader["sections"] := [Map("key", "source", "content", "Synthetic Japanese context")]
    reader["versions"] := [Map("version", 1)]
    reader["source"].Value := "お城には行けました？"
    DirCreate(A_ScriptDir "\scripts")
    FileAppend("# fixture only", A_ScriptDir "\scripts\example_sentence.py")
    for kind in ["example", "version"] {
        for outcome in ["success", "failure", "exception", "closed", "ownerClosed"] {
            if outcome = "ownerClosed" && kind = "example"
                continue
            s := kind = "example" ? TestDesktopAnkiControls(reader) : TestDesktopNewVersionControls(reader)
            DesktopAssert(!DesktopShown(s["progressBar"]), kind " is initially idle")
            sizes := kind = "example" ? [[960, 700], [1040, 760]] : [[820, 620], [920, 740]]
            for size in sizes {
                DesktopDialogFixtureShow(s, size[1], size[2])
                s["status"].Value := kind = "example"
                    ? "Generating a learner-friendly example sentence... This may take a moment."
                    : "Generating another explanation... This may take a moment."
                StudySetActivity(s, true)
                TestGenerationBarBounds(s, size[2])
                StudySetActivity(s, false)
                DesktopAssert(!DesktopShown(s["progressBar"]), kind " returns to idle after resize")
            }
            GenerationFixture := Map("state", s, "kind", kind, "outcome", outcome, "calls", 0, "notices", 0)
            originalBack := s.Has("back") ? s["back"].Value : ""
            if kind = "example"
                TestStudyReaderGenerateVocabularyExample(s)
            else
                TestStudyReaderGenerateNewVersion(s)
            DesktopAssert(GenerationFixture["calls"] = 1, kind " invokes the helper once: " outcome)
            DesktopAssert(!s["modelActivity"], kind " clears activity: " outcome)
            if outcome = "ownerClosed"
                reader["closed"] := false ; Simulated owner teardown, without destroying the fixture.
            if !s["closed"] {
                DesktopAssert(!DesktopShown(s["progressBar"]), kind " hides animation: " outcome)
                button := kind = "example" ? s["exampleButton"] : s["generateButton"]
                DesktopAssert(button.Enabled, kind " re-enables generation: " outcome)
                if kind = "example" {
                    DesktopAssert(s["addButton"].Enabled, "Anki action restored")
                    DesktopAssert(outcome = "success" ? InStr(s["back"].Value, "Example sentence:") : s["back"].Value = originalBack,
                        "Only a successful example changes the reviewed card")
                    StudyReaderCloseAnkiAddDialog(s)
                } else
                    StudyReaderCloseNewVersionDialog(s)
            }
            DesktopAssert(GenerationFixture["notices"] = (outcome = "failure" || outcome = "exception" ? 1 : 0),
                kind " reports errors only for failed requests: " outcome)
            StudySetActivity(s, false)
            StudySetActivity(s, true)
            DesktopAssert(!s["modelActivity"], "Late activity calls cannot revive a destroyed dialog")
        }
    }
    TestGenerationFullscreen(reader)
    ToastDestroy()
    reader["gui"].Destroy(), library["gui"].Destroy()
    CPStudyReaderState := 0
}

TestGenerationFullscreen(reader) {
    reader["bigBoxPresentation"] := true
    for kind in ["example", "version"] {
        s := kind = "example" ? TestDesktopAnkiControls(reader, true, true, true) : TestDesktopNewVersionControls(reader)
        g := s["gui"], form := s["bigBoxForm"]
        DesktopAssert(!s.Has("desktop"), "Fullscreen generation retains its existing shell")
        for size in [[1280, 720], [1920, 1080], [3840, 2160]] {
            w := size[1], h := size[2]
            g.Show("Hide x-12000 y-12000 w" w " h" h)
            StudyCandidatesRecommendationBigBoxResize(form, g, 0, w, h)
            CPApplyOwnedDialogTheme(g)
            g.Show("NA x-12000 y-12000 w" w " h" h)
            s["status"].Value := "Generating... This may take a moment."
            StudySetActivity(s, true)
            s["progressBar"].GetPos(&x, &y, &bw, &bh)
            s["status"].GetPos(&sx, &sy, &sw, &sh)
            primary := kind = "example" ? s["addButton"] : s["generateButton"]
            primary.GetPos(, &buttonY)
            DesktopAssert(DesktopShown(s["progressBar"]) && x = sx && bw = sw && y >= sy + sh,
                "Fullscreen activity fits below status: " kind " " w)
            DesktopAssert(bh > 0 && x >= 0 && x + bw <= w && y + bh < buttonY,
                "Fullscreen activity leaves room for actions: " kind " " w)
            StudySetActivity(s, false)
            DesktopAssert(!DesktopShown(s["progressBar"]), "Fullscreen activity stops")
        }
        if kind = "example"
            StudyReaderCloseAnkiAddDialog(s)
        else
            StudyReaderCloseNewVersionDialog(s)
    }
    reader["bigBoxPresentation"] := false
}

TestGenerationBarBounds(s, height) {
    bar := s["progressBar"], status := s["status"]
    DesktopAssert(DesktopShown(bar) && (WinGetStyle(bar.Hwnd) & 0x8), "Native marquee is visible")
    DesktopAssert(!(WinGetStyle(bar.Hwnd) & 0x10000), "Activity bar does not take keyboard focus")
    bar.GetPos(&x, &y, &w, &h), status.GetPos(&sx, &sy, &sw, &sh)
    DesktopAssert(x = sx && w = sw && y >= sy + sh && h >= 4, "Bar fits below the status text")
    DesktopAssert(y + h <= height - 16, "Bar stays inside the dialog")
    button := s.Get("generateButton", s["addButton"])
    button.GetPos(&bx, &by, &bw, &bh)
    DesktopAssert(y + h <= by || x + w < bx, "Bar never overlaps the primary action")
}

TestGenerationWait(*) {
    global GenerationFixture
    f := GenerationFixture, s := f["state"], kind := f["kind"]
    f["calls"] += 1
    DesktopAssert(s["modelActivity"] && DesktopShown(s["progressBar"]), "Animation starts before waiting")
    button := kind = "example" ? s["exampleButton"] : s["generateButton"]
    DesktopAssert(!button.Enabled && !s["addButton"].Enabled, "Request actions disabled while waiting")
    ; Recursive attempts cannot start another helper while the first is busy.
    if kind = "example"
        TestStudyReaderGenerateVocabularyExample(s)
    else
        TestStudyReaderGenerateNewVersion(s)
    if f["outcome"] = "success" {
        for frame in [1, 2] {
            Sleep(150)
            s["gui"].GetClientPos(,, &w, &h)
            dpi := GetWindowDPI(s["gui"].Hwnd) / 96
            TestDesktopStudyCapture(s["gui"], "activity-" kind "-" frame ".png", Round(w * dpi), Round(h * dpi))
            if frame = 1 {
                s["progressBar"].GetPos(&x, &y, &bw, &bh)
                FileAppend(kind "|" Round(x * dpi) "|" Round(y * dpi) "|" Round(bw * dpi) "|" Round(bh * dpi) "`n",
                    A_ScriptDir "\activity-bounds.txt")
            }
        }
        if kind = "example"
            FileAppend("ok`t" StudyLibraryHexEncode("お城へ行きます。") "`t"
                . StudyLibraryHexEncode("おしろへいきます。") "`t" StudyLibraryHexEncode("I will go to the castle."),
                EnvGet("EXAMPLE_RESULT_FILE"), "UTF-8-RAW")
        return 0
    }
    if f["outcome"] = "exception"
        throw Error("Synthetic launch failure")
    if f["outcome"] = "closed" {
        if kind = "example"
            StudyReaderCloseAnkiAddDialog(s)
        else
            StudyReaderCloseNewVersionDialog(s)
        return 0
    }
    if f["outcome"] = "ownerClosed" {
        s["readerState"]["closed"] := true
        return 0
    }
    return 1
}

TestGenerationNotice(*) {
    global GenerationFixture
    GenerationFixture["notices"] += 1
    DesktopAssert(!GenerationFixture["state"]["modelActivity"], "Animation stops before showing the failure")
}

TestGenerationLoadGroup(reader, *) {
    reader["versions"].Push(Map("version", 2))
}
