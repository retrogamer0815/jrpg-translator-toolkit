; Real profile application, native desktop controls and toast. All files/windows
; are private fixtures; no user profiles, overlays, providers or Anki are used.
TestProfileFeedbackBegin() {
    global ui, capMode := "region", capRect := "10,10,100,100", capWinInfo := ""
    global captureDir := A_ScriptDir "\captures", overlayDir := A_ScriptDir "\overlay"
    global appDir := A_ScriptDir "\Settings"
    global studyLibraryDefaultDir := A_ScriptDir "\study-fixture"
    global studyLibrariesRoot := A_ScriptDir "\study-libraries", studyLibraryDir := studyLibraryDefaultDir
    global ewX := "", ewY := "", ewW := "", ewH := ""
    global ew_lastX := "", ew_lastY := "", ew_lastW := "", ew_lastH := "", ew_bounds_watch_running := false
    global TestProfileFeedback, CPDesktop
    DirCreate(captureDir), DirCreate(overlayDir)
    TestProfileFeedback := Map("step", 0)
    for name in ["Demo profile", "Alternate profile"] {
        path := GameProfilePath(name)
        IniWrite(1, path, "profile", "schemaVersion")
        IniWrite(name = "Demo profile" ? "default_with_kanji_reading_en" : "literal", path, "screenshot", "promptProfile")
        IniWrite(name = "Demo profile" ? "default_en" : "detailed_grammar", path, "explanation", "promptProfile")
        IniWrite(name = "Demo profile" ? "202020" : "FF0000", path, "translator", "boxBg")
        IniWrite(name = "Demo profile" ? 220 : 255, path, "translator", "overlayTrans")
    }
    ui.Show("NA x0 y0 w1120 h760")
    CPDesktopNavigate(7)
    OnExit(ToastDestroy)
    SetTimer(TestProfileFeedbackStep, -500)
    return true
}

TestProfileFeedbackStep(*) {
    global TestProfileFeedback, ddlGameProfile, iniPath, CPDesktop, CPToastGui, CPToastText, controlDarkMode, ui
    try {
        step := TestProfileFeedback["step"]
        FileAppend("STEP " step "`n", A_ScriptDir "\profile-events.txt")
        if Mod(step, 3) = 0 {
            if step = 6 || step = 12 {
                controlDarkMode := step = 12
                CPDesktopTheme()
                if step = 12
                    WinSetTransparent(220, ui.Hwnd)
            }
            name := Mod(step // 3, 2) = 0 ? "Alternate profile" : "Demo profile"
            TestProfileFeedback["name"] := name
            TestProfileFeedback["message"] := "Profile applied: " name
            FileAppend("before apply`n", A_ScriptDir "\profile-events.txt")
            if step = 9 || step = 15 {
                ; Save with a notification, then switch via the header while it
                ; is still visible. Saving also avoids an unsaved-changes modal.
                active := IniRead(iniPath, "game_profiles", "active")
                DesktopAssert(GameProfileSave(active), "Real profile save succeeds")
                DesktopAssert(!GameProfileHasUnsavedChanges(active), "Saved fixture has no pending edits")
                TestProfileFeedback["message"] := "Profile saved: " active
                CPDesktop["chrome"]["profile"].Choose(name)
                CPDesktopProfileSelectionChanged(CPDesktop["chrome"]["profile"])
            } else {
                if step = 18 {
                    Toast("Generating explanation…")
                    previousToast := CPToastGui
                }
                SetComboToExistingItem(ddlGameProfile, ListGameProfiles(), name)
                GameProfileUpdateSummary()
                ApplySelectedGameProfile()
                if step = 18
                    DesktopAssert(CPToastGui != previousToast, "Profile notification replaces another workflow's toast")
            }
            FileAppend("after apply`n", A_ScriptDir "\profile-events.txt")
            DesktopAssert(A_IsCritical = 0, "Profile switch restores interruptibility")
            DesktopAssert(IniRead(iniPath, "game_profiles", "active") = name, "Profile is persisted")
            DesktopAssert(CPDesktop["chrome"]["profile"].Text = name, "Header reflects applied profile")
            TestProfileFeedback["toastHwnd"] := CPToastGui.Hwnd
            TestProfileFeedback["step"] += 1
            SetTimer(TestProfileFeedbackStep, -350)
        } else if Mod(step, 3) = 1 {
            DesktopAssert(CPToastText.Text = TestProfileFeedback["message"], "Toast names the saved/applied profile")
            DesktopAssert(!CPToastText.Visible, "Toast does not require native child painting")
            TestToastFeedbackPixels(controlDarkMode)
            TestAudioFeedbackPixels(CPDesktop["chrome"]["title"], controlDarkMode)
            ; Let native page repainting run with the notification still up.
            CPDesktopNavigate(1)
            CPDesktopNavigate(7)
            TestProfileFeedback["step"] += 1
            SetTimer(TestProfileFeedbackStep, -1800)
        } else {
            DesktopAssert(!DllCall("user32\IsWindow", "ptr", TestProfileFeedback["toastHwnd"]), "Profile toast dismisses while idle")
            if step >= 23 {
                process := DllCall("kernel32\GetCurrentProcess", "ptr")
                beforeGdi := DllCall("user32\GetGuiResources", "ptr", process, "uint", 0)
                beforeUser := DllCall("user32\GetGuiResources", "ptr", process, "uint", 1)
                Loop 40 {
                    Toast("Profile applied: Stress " A_Index)
                    ToastDestroy()
                }
                DesktopAssert(DllCall("user32\GetGuiResources", "ptr", process, "uint", 0) <= beforeGdi + 2,
                    "Repeated notifications release their GDI bitmaps, brushes and DCs")
                DesktopAssert(DllCall("user32\GetGuiResources", "ptr", process, "uint", 1) <= beforeUser + 2,
                    "Repeated notifications release their windows")
                Toast("Profile applied: 日本語 & English`nExplanation ready.")
                TestProfileFeedback["toastHwnd"] := CPToastGui.Hwnd
                SetTimer(TestProfileFeedbackUnicode, -350)
                return
            }
            TestProfileFeedback["step"] += 1
            SetTimer(TestProfileFeedbackStep, -100)
        }
    } catch as ex {
        FileAppend("FAIL: " ex.Message "`n" ex.Stack "`n", "*")
        ExitApp(1)
    }
}

TestProfileFeedbackUnicode(*) {
    global controlDarkMode
    TestToastFeedbackPixels(controlDarkMode)
    SetTimer(TestProfileFeedbackDelete, -1800)
}

TestProfileFeedbackDelete(*) {
    global TestProfileFeedback, iniPath, ddlGameProfile, CPToastGui, CPToastText, CPDesktop
    DesktopAssert(!DllCall("user32\IsWindow", "ptr", TestProfileFeedback["toastHwnd"]), "Multiline replacement dismisses while idle")
    name := IniRead(iniPath, "game_profiles", "active")
    SetComboToExistingItem(ddlGameProfile, ListGameProfiles(), name)
    SetTimer(TestProfileFeedbackConfirmDelete, 80)
    try DeleteSelectedGameProfile()
    finally SetTimer(TestProfileFeedbackConfirmDelete, 0)
    DesktopAssert(!FileExist(GameProfilePath(name)), "Confirmed fixture profile was deleted")
    DesktopAssert(CPDesktop["chrome"]["profile"].Text = "Current settings", "Deleting active profile refreshes header")
    DesktopAssert(CPToastText.Text = "Profile deleted: " name, "Deletion announces the deleted profile name")
    TestProfileFeedback["toastHwnd"] := CPToastGui.Hwnd
    SetTimer(TestProfileFeedbackDeletedPixels, -350)
}

TestProfileFeedbackConfirmDelete(dialogTitle := "Profiles", *) {
    ; Respond only to this isolated fixture's real themed confirmation dialog.
    for hwnd, dialog in StudyDesktopRegistry() {
        if dialog["kind"] = "message" && dialog["state"]["gui"].Title = dialogTitle && dialog["width"] > 0 {
            SendMessage(0xF5, 0, 0, dialog["state"]["addButton"].Hwnd)
            return
        }
    }
}

TestProfileFeedbackDeletedPixels(*) {
    global controlDarkMode
    TestToastFeedbackPixels(controlDarkMode)
    SetTimer(TestProfileFeedbackFinish, -1800)
}

TestProfileFeedbackFinish(*) {
    global TestProfileFeedback, TestAssertions, ui
    DesktopAssert(!DllCall("user32\IsWindow", "ptr", TestProfileFeedback["toastHwnd"]), "Deletion notification dismisses while idle")
    index := TestProfileFeedback.Get("promptDeleteIndex", 0)
    if index < 2 {
        TestProfileFeedback["promptDeleteIndex"] := index + 1
        TestProfileFeedbackDeletePrompt(index + 1)
        return
    }
    TestProfileFeedbackCreatePrompts()
    FileAppend("PASS: " TestAssertions " real profile-feedback assertions.`n", "*")
    ui.Destroy()
    ExitApp(0)
}

TestProfileFeedbackDeletePrompt(index) {
    global TestProfileFeedback, ddlPrompt, ddlEPr, CPToastGui, CPToastText
    explanation := index = 2
    name := explanation ? "Explanation fixture" : "Translation fixture"
    path := explanation ? ExplainProfilePath(name) : PromptFilePath(name)
    FileAppend("Synthetic prompt instructions.", path, "UTF-8")
    if explanation
        RefreshExplainPromptProfilesList(name)
    else
        RefreshPromptProfilesList(name)
    CPDesktopNavigate(explanation ? 4 : 1)
    confirm := TestProfileFeedbackConfirmDelete.Bind(explanation ? "Delete EXPLAIN prompt" : "Delete prompt")
    SetTimer(confirm, 80)
    try {
        if explanation
            DeleteExplainPromptProfile()
        else
            DeletePromptProfile()
    } finally SetTimer(confirm, 0)
    DesktopAssert(!FileExist(path), "Confirmed fixture prompt was deleted")
    DesktopAssert((explanation ? ddlEPr : ddlPrompt).Text != name, "Deleted prompt no longer appears in selector")
    DesktopAssert(CPToastText.Text = (explanation ? "Explanation" : "Translation") " prompt deleted: " name,
        "Notification identifies the deleted prompt type and name")
    TestProfileFeedback["toastHwnd"] := CPToastGui.Hwnd
    SetTimer(TestProfileFeedbackDeletedPixels, -350)
}

TestProfileFeedbackCreatePrompts() {
    global ddlPrompt, ddlEPr
    for explanation in [false, true] {
        name := (explanation ? "Explanation" : "Translation") " draft 日本語"
        path := explanation ? ExplainProfilePath(name) : PromptFilePath(name)
        CPDesktopNavigate(explanation ? 4 : 1)
        finishName := TestProfileFeedbackNamePrompt.Bind(explanation ? "New EXPLAIN prompt" : "New prompt", name)
        SetTimer(finishName, 80)
        try {
            if explanation
                NewExplainPromptProfile()
            else
                NewPromptProfile()
        } finally SetTimer(finishName, 0)
        editor := 0, count := 0
        for hwnd, dialog in StudyDesktopRegistry() {
            if dialog["kind"] = "textEditor" && dialog["state"].Get("path", "") = path {
                editor := dialog["state"]
                count += 1
            }
        }
        DesktopAssert(count = 1, "Naming a new prompt immediately opens exactly one editor")
        DesktopAssert(editor["ready"] && DllCall("user32\IsWindowVisible", "ptr", editor["gui"].Hwnd),
            "New prompt editor is visible and ready")
        DesktopAssert(editor["controls"]["editor"].Value = "", "New prompt editor has no prefilled instructions")
        DesktopAssert(FileRead(path, "UTF-8") = "", "New prompt file starts empty")
        DesktopAssert((explanation ? ddlEPr : ddlPrompt).Text = name, "New prompt remains selected")
        if explanation {
            editor["gui"].GetClientPos(,, &w, &h)
            dpi := GetWindowDPI(editor["gui"].Hwnd) / 96
            TestDesktopStudyCapture(editor["gui"], "new-explanation-prompt.png", Round(w * dpi), Round(h * dpi))
        }
        text := "Custom Japanese instructions. 日本語 {jp}"
        editor["controls"]["editor"].Value := text
        SendMessage(0xF5, 0, 0, editor["controls"]["save"].Hwnd)
        Sleep(25)
        DesktopAssert(FileRead(path, "UTF-8") = text, "Explicit Save persists the new prompt's edited text")
        editorHwnd := editor["gui"].Hwnd
        CPTextEditorDialogClose(editor)
        DesktopAssert(!DllCall("user32\IsWindow", "ptr", editorHwnd), "New prompt editor closes cleanly")
        ToastDestroy()
    }
}

TestProfileFeedbackNamePrompt(title, name) {
    ; Only answer the naming dialog created by this isolated test process.
    for hwnd, dialog in StudyDesktopRegistry() {
        g := dialog["state"]["gui"]
        if g.Title != title || !g.HasOwnProp("CPInputDialog") || !g.CPInputDialog.Get("ready", false)
            continue
        s := g.CPInputDialog
        s["controls"]["editor"].Value := name
        SendMessage(0xF5, 0, 0, s["controls"]["save"].Hwnd)
        return
    }
}
