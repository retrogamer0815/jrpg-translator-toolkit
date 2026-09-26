; Loaded by test_desktop_layout.ps1 -PromptUnsavedOnly. All windows, files,
; messages, and simulated button/key events belong to this isolated process.
TestPromptUnsavedChanges() {
    global TestPromptReplies := [], TestPromptMessageCount := 0, TestPromptErrorCount := 0
    global TestPromptCurrent := 0, TestPromptCapture := false, controlDarkMode
    global promptsDir, explainPromptsDir, ddlPrompt, ddlEPr, ui
    SetTimer(TestPromptMessageReply, 30)
    try {
        for dark in [1, 0] {
            controlDarkMode := dark
            for route in ["translation", "explanation", "legacy-explanation", "classic", "reader", "reader-fullscreen"] {
                FileAppend("Testing " route " (dark=" dark ")`n", "*")
                path := route = "translation" ? PromptFilePath("unsaved-fixture")
                    : route = "legacy-explanation" ? ExplainPromptFilePath() : ExplainProfilePath("unsaved-fixture")
                original := "Explain Japanese. 日本語 {jp}`nKeep ALL spaces.  "
                SaveTextAtomic(path, original, false)
                ddlPrompt.Delete(), ddlPrompt.Add(["unsaved-fixture"]), ddlPrompt.Choose(1)
                ddlEPr.Delete(), ddlEPr.Add(["unsaved-fixture"]), ddlEPr.Choose(1)
                s := TestPromptOpen(route, path, original)
                TestPromptCurrent := s
                c := s["controls"], editor := c["editor"], hwnd := s["gui"].Hwnd
                messageCountBefore := TestPromptMessageCount
                editor.Value := StrReplace(original, "ALL", "all") ; Case-only edit is dirty.
                draft := editor.Value
                for response in ["Cancel", "window-close", "escape"] {
                    TestPromptReplies.Push(response)
                    TestPromptClose(s, "window-close")
                    DesktopAssert(!s["closed"] && DllCall("user32\IsWindowVisible", "ptr", hwnd),
                        route ": cancelled X-close leaves editor visible")
                    DesktopAssert(editor.Value == draft && FileRead(path, "UTF-8") == original,
                        route ": cancel preserves both the draft and saved file")
                    DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", hwnd), route ": editor re-enabled")
                }
                DesktopAssert(TestPromptMessageCount = messageCountBefore + 3, route ": each dirty close asks exactly once")
                editor.Value := StrReplace(original, "`n", "`r`n")
                TestPromptClose(s, "button")
                DesktopAssert(s["closed"] && TestPromptMessageCount = messageCountBefore + 3,
                    route ": reverted edit / equivalent newlines close without a warning")

                s := TestPromptOpen(route, path, original), TestPromptCurrent := s
                s["controls"]["editor"].Value := original "`n保存した変更"
                SendMessage(0xF5, 0, 0, s["controls"]["save"].Hwnd)
                Sleep(80)
                saved := FileRead(path, "UTF-8")
                DesktopAssert(InStr(saved, "保存した変更") && !s["closed"], route ": Save persists and leaves editor open")
                DesktopAssert(FileRead(path ".bak", "UTF-8") == original, route ": Save preserves previous prompt backup")
                messageCountBefore := TestPromptMessageCount
                TestPromptClose(s, "escape")
                DesktopAssert(s["closed"] && TestPromptMessageCount = messageCountBefore, route ": saved editor closes silently with Escape")

                s := TestPromptOpen(route, path, saved), TestPromptCurrent := s
                s["controls"]["editor"].Value := "Save and close. 日本語`n{jp}"
                TestPromptReplies.Push("Yes")
                TestPromptClose(s, "button")
                DesktopAssert(s["closed"] && CPPromptEditorComparableText(FileRead(path, "UTF-8"))
                    == "Save and close. 日本語`n{jp}", route ": Save and close writes before closing")
                DesktopAssert(FileRead(path ".bak", "UTF-8") == saved, route ": close-save also backs up the previous prompt")

                saved := FileRead(path, "UTF-8"), backup := FileRead(path ".bak", "UTF-8")
                s := TestPromptOpen(route, path, saved), TestPromptCurrent := s
                s["controls"]["editor"].Value := saved " " ; Whitespace-only edit is dirty.
                TestPromptReplies.Push("No")
                TestPromptClose(s, "escape")
                DesktopAssert(s["closed"] && FileRead(path, "UTF-8") == saved
                    && FileRead(path ".bak", "UTF-8") == backup, route ": Discard never writes file or backup")

                s := TestPromptOpen(route, path, saved), TestPromptCurrent := s
                s["controls"]["editor"].Value := "Keep this draft if saving fails. 日本語"
                errorsBefore := TestPromptErrorCount
                FileSetAttrib("+R", path)
                try {
                    TestPromptReplies.Push("Yes")
                    TestPromptClose(s, "button")
                    DesktopAssert(!s["closed"] && s["controls"]["editor"].Value == "Keep this draft if saving fails. 日本語",
                        route ": failed save keeps editor and draft")
                    DesktopAssert(FileRead(path, "UTF-8") == saved && TestPromptErrorCount = errorsBefore + 1,
                        route ": failed save reports error and preserves saved file")
                } finally {
                    FileSetAttrib("-R", path)
                    ; FileCopy carries the artificial read-only flag into the
                    ; backup too. Restore this fixture before the next route.
                    if FileExist(path ".bak")
                        FileSetAttrib("-R", path ".bak")
                }
                messageCountBefore := TestPromptMessageCount
                ; BM_CLICK synthesizes mouse messages; a second immediate click
                ; on the classic button would be classified as DoubleClick.
                Sleep(DllCall("user32\GetDoubleClickTime") + 50)
                DesktopAssert(DllCall("user32\IsWindowEnabled", "ptr", s["gui"].Hwnd),
                    route ": failed-save error restores the editor")
                TestPromptReplies.Push("No")
                TestPromptClose(s, "button")
                DesktopAssert(s["closed"] && TestPromptMessageCount = messageCountBefore + 1,
                    route ": failed save does not reset dirty state")
            }
        }
        ; New prompts start empty but still protect text once the user types.
        path := PromptFilePath("blank-new")
        SaveTextAtomic(path, "", false)
        ddlPrompt.Delete(), ddlPrompt.Add(["blank-new"]), ddlPrompt.Choose(1)
        s := OpenPromptEditor(), TestPromptCurrent := s
        s["controls"]["editor"].Value := "New instructions"
        TestPromptReplies.Push("Yes")
        TestPromptClose(s, "window-close")
        DesktopAssert(s["closed"] && FileRead(path, "UTF-8") == "New instructions", "New blank prompt is protected")
        DesktopAssert(TestPromptReplies.Length = 0, "All expected confirmations were shown")
    } finally {
        SetTimer(TestPromptMessageReply, 0)
        TestPromptCurrent := 0
    }
}

TestPromptOpen(route, path, text) {
    global ui, ddlEPr
    if route = "translation"
        return OpenPromptEditor()
    if route = "explanation"
        return OpenExplainPromptEditor_Multi()
    if route = "legacy-explanation"
        return OpenExplainPromptEditor()
    if route = "classic" {
        g := Gui("+Resize +Owner" ui.Hwnd, "Synthetic classic prompt")
        g.CPDialogPresentation := "classic"
        editor := g.AddEdit("w680 h420 WantReturn", text)
        save := g.AddButton("w100", "Save"), close := g.AddButton("x+8 w100", "Close")
        s := CPPromptEditorWireLegacy(g, editor, save, close, path, "Saved fixture", "the test prompt")
        CPShowTextEditorDialog(g, editor, ui.Hwnd)
        return s
    }
    parent := Map("gui", ui, "prompt", ddlEPr, "bigBoxPresentation", route = "reader-fullscreen")
    s := StudyReaderPromptDialogCreate(parent, "edit", text, "unsaved-fixture")
    s["saveAction"] := StudyReaderSaveNewVersionPrompt.Bind(parent, s["gui"],
        s["controls"]["editor"], path, "unsaved-fixture")
    s["controls"]["save"].OnEvent("Click", s["saveAction"])
    if route = "reader-fullscreen"
        StudyLibraryBigBoxFormShow(s["bigBoxForm"], s["controls"]["editor"])
    else
        StudyDesktopDialogShow(s, 1000, 780, s["controls"]["editor"])
    return s
}

TestPromptClose(s, route) {
    global TestPromptReplies
    ; BM_CLICK may be ignored by an inactive native dialog after a nested
    ; error closes. Match a real user's next click by activating this fixture.
    WinActivate("ahk_id " s["gui"].Hwnd)
    if route = "button" {
        c := s["controls"]
        SendMessage(0xF5, 0, 0, c.Has("close") ? c["close"].Hwnd : c["cancel"].Hwnd)
    } else if route = "window-close"
        PostMessage(0x10, 0, 0, s["gui"].Hwnd)
    else
        ControlSend("{Escape}", s["controls"]["editor"])
    deadline := A_TickCount + 6000
    Sleep(100)
    while (TestPromptReplies.Length || s.Get("confirmingClose", false)) && A_TickCount < deadline
        Sleep(25)
    DesktopAssert(!TestPromptReplies.Length && !s.Get("confirmingClose", false),
        "Close action completed: " route " closed=" s["closed"] " pending=" TestPromptReplies.Length
        " confirming=" s.Get("confirmingClose", false))
}

TestPromptMessageReply() {
    global TestPromptReplies, TestPromptMessageCount, TestPromptErrorCount, TestPromptCurrent, TestPromptCapture
    static visibleSince := Map()
    for hwnd in WinGetList("ahk_pid " DllCall("GetCurrentProcessId")) {
        if !DllCall("user32\IsWindowVisible", "ptr", hwnd)
            continue
        g := GuiFromHwnd(hwnd)
        if !IsObject(g) {
            ; The classic light theme uses a native owned error message.
            if IsObject(TestPromptCurrent) && TestPromptCurrent.Get("confirmingClose", false)
                && WinGetClass("ahk_id " hwnd) = "#32770"
                && InStr(WinGetText("ahk_id " hwnd), "Could not save") {
                TestPromptErrorCount++
                PostMessage(0x111, 1, 0, hwnd) ; IDOK, only this synthetic process.
            }
            continue
        }
        buttons := Map()
        for ctrl in g
            if ctrl.Type = "Button"
                buttons[ctrl.Text] := ctrl
        if g.Title = "Unsaved prompt changes" && buttons.Has("Save and close") {
            ; A timer may interrupt Show/WinActivate before initial focus is set.
            if !visibleSince.Has(hwnd) {
                visibleSince[hwnd] := A_TickCount
                continue
            }
            if A_TickCount - visibleSince[hwnd] < 200
                continue
            visibleSince.Delete(hwnd)
            DesktopAssert(TestPromptReplies.Length > 0, "No unexpected unsaved-changes modal")
            DesktopAssert(buttons.Has("Discard changes") && buttons.Has("Cancel"), "Standard three explicit choices")
            DesktopAssert(IsObject(g.FocusedCtrl) && g.FocusedCtrl.Hwnd = buttons["Cancel"].Hwnd,
                "Cancel is the safe initial focus")
            owner := DllCall("user32\GetWindow", "ptr", hwnd, "uint", 4, "ptr")
            DesktopAssert(owner = TestPromptCurrent["gui"].Hwnd && !DllCall("user32\IsWindowEnabled", "ptr", owner),
                "Confirmation is modal to the correct prompt editor")
            DesktopAssert(!CPPromptEditorConfirmClose(TestPromptCurrent), "Repeated close cannot create a nested confirmation")
            if !TestPromptCapture && IsObject(StudyDesktopContext(hwnd)) {
                WinGetClientPos(, , &w, &h, "ahk_id " hwnd)
                TestDesktopStudyCapture(g, "unsaved-prompt-confirmation.png", w, h)
                TestPromptCapture := true
            }
            TestPromptMessageCount++
            response := TestPromptReplies.RemoveAt(1)
            if response = "window-close"
                PostMessage(0x10, 0, 0, hwnd)
            else if response = "escape"
                ControlSend("{Escape}", buttons["Cancel"])
            else
                PostMessage(0xF5, 0, 0, buttons[response = "Yes" ? "Save and close" : response = "No" ? "Discard changes" : "Cancel"].Hwnd)
            return
        }
        ; Save failure opens a second, owned error dialog; never dismiss the
        ; prompt itself or any unrelated window.
        if IsObject(TestPromptCurrent) && TestPromptCurrent.Get("confirmingClose", false)
            && buttons.Has("OK") && g.Title != "Unsaved prompt changes" {
            owner := DllCall("user32\GetWindow", "ptr", hwnd, "uint", 4, "ptr")
            if owner = TestPromptCurrent["gui"].Hwnd {
                TestPromptErrorCount++
                PostMessage(0xF5, 0, 0, buttons["OK"].Hwnd)
                return
            }
        }
    }
}
