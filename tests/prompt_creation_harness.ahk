#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
; Production creation functions and real fixture files; dialog input and editor
; presentation are substituted to test branching without user data or APIs.
global TestAssertions := 0, TestEditors := [], TestNotices := [], TestEvents := []
global TestInput := {Result: "OK", Value: ""}, TestDesktop := true
global ddlPrompt := {Text: ""}, ddlEPr := {Text: ""}
global promptsDir := A_ScriptDir "\translation", explainPromptsDir := A_ScriptDir "\explanation"
global CPBigBoxSetupState := Map(), CPBigBoxControls := Map("setup_raw", {Value: ""})
DirCreate(promptsDir), DirCreate(explainPromptsDir)
try {
    TestPromptCreation()
    FileAppend("PASS: " TestAssertions " prompt creation assertions (" A_PtrSize * 8 "-bit).`n", "*")
    ExitApp(0)
} catch as testFailure {
    FileAppend("FAIL: " testFailure.Message "`n" testFailure.Stack "`n", "*")
    ExitApp(1)
}

Check(condition, message) {
    global TestAssertions
    TestAssertions += 1
    if !condition
        throw Error(message)
}

ResetCreation(name, result := "OK") {
    global TestInput, TestEditors, TestNotices, TestEvents
    TestInput := {Result: result, Value: name}
    TestEditors := [], TestNotices := [], TestEvents := []
}

TestPromptCreation() {
    global TestDesktop, TestEditors, TestNotices, TestEvents, ddlPrompt, ddlEPr
    global CPBigBoxSetupState, CPBigBoxControls
    for route in ["translation", "explanation", "reader-desktop", "reader-classic", "reader-fullscreen"] {
        TestDesktop := route != "reader-classic" && route != "reader-fullscreen"
        state := Map("gui", {Hwnd: 0}, "prompt", {Text: "previous"},
            "bigBoxPresentation", route = "reader-fullscreen")
        action := route = "translation" ? NewPromptProfile : route = "explanation" ? NewExplainPromptProfile
            : StudyReaderAddNewVersionPrompt.Bind(state)
        pathFor := route = "translation" ? PromptFilePath : ExplainProfilePath
        for result in ["Cancel", "OK"] {
            name := result = "Cancel" ? route " cancelled" : "   "
            ResetCreation(name, result)
            action.Call()
            Check(TestEditors.Length = 0 && TestEvents.Length = 0, route ": cancelled/empty name opens no editor")
            Check(!FileExist(pathFor.Call(name)), route ": cancelled/empty name creates no prompt")
        }
        for typedName in [route " new", route " 日本語 & notes", route ":name"] {
            name := StrReplace(typedName, ":", "_")
            ResetCreation("  " typedName "  ")
            action.Call()
            path := pathFor.Call(name)
            Check(FileExist(path), route ": creates the selected prompt")
            Check(FileRead(path, "UTF-8") = "", route ": new prompt starts with no template")
            Check(TestEditors.Length = 1 && TestEditors[1]["name"] = name && TestEditors[1]["text"] = "",
                route ": immediately opens exactly one editor for the new empty prompt")
            Check(TestEvents[1] = (route = "translation" ? "translation-refresh" : route = "explanation" ? "explanation-refresh" : "reader-refresh")
                && TestEvents[TestEvents.Length] = "edit", route ": selects the new prompt before opening its editor")
            Check(TestNotices.Length = 0, route ": valid creation raises no error")
        }
        name := route " existing", path := pathFor.Call(name)
        FileAppend("Keep these instructions.", path, "UTF-8")
        ResetCreation(name)
        action.Call()
        Check(FileRead(path, "UTF-8") = "Keep these instructions.", route ": duplicate name never overwrites existing prompt")
        Check(TestEditors.Length = 0 && TestNotices.Length = 1, route ": duplicate is reported, not silently recreated")
    }
    for domain in ["translation", "explanation"] {
        CPBigBoxSetupState := Map("domain", domain)
        CPBigBoxControls["setup_raw"].Value := "Previous editor text"
        ResetCreation("")
        name := "fullscreen-new-" domain
        CPBigBoxStartPromptText(name)
        Check(CPBigBoxSetupState["flow"] = "promptText" && CPBigBoxSetupState["name"] = name,
            domain ": fullscreen advances from name to editor")
        Check(CPBigBoxControls["setup_raw"].Value = "", domain ": fullscreen editor is blank, not stale or prefilled")
        Check(CPBigBoxSetupState["isNew"] && CPBigBoxSetupState["stamp"] = "missing",
            domain ": retains fullscreen unsaved-prompt state")
        Check(TestEvents.Length = 1 && TestEvents[1] = "setupEdit", domain ": opens the fullscreen editor page")
        Check(!FileExist((domain = "translation" ? PromptFilePath : ExplainProfilePath).Call(name)),
            domain ": fullscreen still waits for explicit text save before writing a file")
    }
}

CPDesktopActive() {
    global TestDesktop
    return TestDesktop
}
CPDialogDefaultOwner() => 0
CPThemedInputBox(*) {
    global TestInput
    return TestInput
}
CPAdaptiveOwnedMessage(owner, message, *) {
    global TestNotices
    TestNotices.Push(message)
    return "No"
}
CPThemedOwnedMessage(args*) => CPAdaptiveOwnedMessage(args*)
StudyReaderPromptDialogCreate(parent, mode) => Map("mode", mode)
StudyReaderPromptDialogRun(state) {
    global TestInput
    Check(state["mode"] = "name", "Reader asks for a name before editing")
    return TestInput
}
RefreshPromptProfilesList(name) {
    global ddlPrompt, TestEvents
    ddlPrompt.Text := name
    TestEvents.Push("translation-refresh")
}
RefreshExplainPromptProfilesList(name) {
    global ddlEPr, TestEvents
    ddlEPr.Text := name
    TestEvents.Push("explanation-refresh")
}
StudyReaderRefreshNewVersionPromptList(state, name) {
    global TestEvents
    state["prompt"].Text := name
    TestEvents.Push("reader-refresh")
}
StudyReaderRefreshMainExplainPromptList() {
    global TestEvents
    TestEvents.Push("main-refresh")
}
OpenPromptEditor(*) {
    global ddlPrompt
    RecordEditor(ddlPrompt.Text, PromptFilePath(ddlPrompt.Text))
}
OpenExplainPromptEditor_Multi(*) {
    global ddlEPr
    RecordEditor(ddlEPr.Text, ExplainProfilePath(ddlEPr.Text))
}
StudyReaderEditNewVersionPrompt(state) => RecordEditor(state["prompt"].Text, ExplainProfilePath(state["prompt"].Text))
RecordEditor(name, path) {
    global TestEditors, TestEvents
    TestEditors.Push(Map("name", name, "text", FileRead(path, "UTF-8")))
    TestEvents.Push("edit")
}
CPBigBoxSetupEnter(page) {
    global TestEvents
    TestEvents.Push(page)
}
