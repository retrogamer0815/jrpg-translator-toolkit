#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
; Real deletion and INI operations, confined to generated test fixtures.
; Only dialog responses, UI refresh and notification presentation are stubbed.
global TestAssertions := 0, TestToasts := [], TestDialogs := [], TestLogs := [], TestEvents := []
global TestAnswer := "Yes", ddlGameProfile := {Text: ""}
global gameProfilesDir := A_ScriptDir "\profiles", iniPath := A_ScriptDir "\settings.ini"
global promptsDir := A_ScriptDir "\translation-prompts", explainPromptsDir := A_ScriptDir "\explanation-prompts"
global ddlPrompt := {Text: ""}, ddlEPr := {Text: ""}
DirCreate(gameProfilesDir)
DirCreate(promptsDir), DirCreate(explainPromptsDir)
try {
    RunDeletionTests()
    RunPromptDeletionTests()
    FileAppend("PASS: " TestAssertions " profile/prompt deletion notification assertions (" A_PtrSize * 8 "-bit).`n", "*")
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

ResetDeletion(name := "", answer := "Yes", active := "") {
    global ddlGameProfile, TestAnswer, TestToasts, TestDialogs, TestLogs, TestEvents, iniPath
    ddlGameProfile.Text := name, TestAnswer := answer
    TestToasts := [], TestDialogs := [], TestLogs := [], TestEvents := []
    IniWrite(active, iniPath, "game_profiles", "active")
}

RunDeletionTests() {
    global TestToasts, TestDialogs, TestLogs, TestEvents, iniPath
    ResetDeletion()
    DeleteSelectedGameProfile()
    Check(TestDialogs.Length = 0 && TestToasts.Length = 0 && TestEvents.Length = 0,
        "No selection has no side effects or notification")

    for answer in ["No", "Cancel"] {
        name := "Cancelled " answer, path := GameProfilePath(name)
        IniWrite(name, path, "profile", "name")
        ResetDeletion(name, answer, name)
        DeleteSelectedGameProfile()
        Check(FileExist(path), "Declining preserves the profile")
        Check(IniRead(iniPath, "game_profiles", "active") = name, "Declining preserves active selection")
        Check(TestDialogs.Length = 1 && TestToasts.Length = 0 && TestEvents.Length = 0,
            "Declining never refreshes or announces a deletion")
    }

    for name in ["Active fixture", "Inactive fixture", "日本語 & English"] {
        path := GameProfilePath(name)
        IniWrite(name, path, "profile", "name")
        active := name = "Inactive fixture" ? "Other profile" : name
        ResetDeletion(name, "Yes", active)
        DeleteSelectedGameProfile()
        Check(!FileExist(path), "Confirmed profile is actually deleted")
        Check(TestToasts.Length = 1 && TestToasts[1] = "Profile deleted: " name,
            "Exactly one success notification includes the deleted name")
        Check(IniRead(iniPath, "game_profiles", "active") = (active = name ? "" : active),
            "Only deletion of the active profile clears its active marker")
        Check(TestEvents.Length = 2 && TestEvents[1] = "refresh" && TestEvents[2] = "toast",
            "Success is announced after the profile list refreshes")
        Check(TestDialogs.Length = 1 && TestLogs.Length = 0, "Successful deletion shows no error")
    }

    name := "Locked fixture", path := GameProfilePath(name)
    IniWrite(name, path, "profile", "name")
    ResetDeletion(name, "Yes", name)
    handle := DllCall("kernel32\CreateFileW", "wstr", path, "uint", 0x80000000,
        "uint", 0, "ptr", 0, "uint", 3, "uint", 0x80, "ptr", 0, "ptr")
    Check(handle != -1 && handle != 0, "Exclusive fixture file lock succeeds")
    try DeleteSelectedGameProfile()
    finally DllCall("kernel32\CloseHandle", "ptr", handle)
    Check(FileExist(path), "Failed deletion preserves the profile")
    Check(IniRead(iniPath, "game_profiles", "active") = name, "Failed deletion preserves active marker")
    Check(TestToasts.Length = 0 && TestEvents.Length = 0, "Failed deletion never reports success or refreshes selection")
    Check(TestDialogs.Length = 2 && InStr(TestDialogs[2], "Could not delete profile '" name "'"),
        "Failed deletion explains the error")
    Check(TestLogs.Length = 1, "Failed deletion is logged")
}

CPDialogDefaultOwner() => 0
CPThemedOwnedMessage(args*) => CPAdaptiveOwnedMessage(args*)
CPAdaptiveOwnedMessage(owner, message, title, buttons := "ok", *) {
    global TestDialogs, TestAnswer
    TestDialogs.Push(message)
    return buttons = "yesno" ? TestAnswer : "OK"
}
IniWriteRetry(value, file, section, key) => IniWrite(value, file, section, key)
RefreshGameProfilesList(*) {
    global TestEvents
    TestEvents.Push("refresh")
}
Toast(message) {
    global TestToasts, TestEvents
    TestToasts.Push(message), TestEvents.Push("toast")
}
DbgCP(message) {
    global TestLogs
    TestLogs.Push(message)
}

RunPromptDeletionTests() {
    global ddlPrompt, ddlEPr, TestToasts, TestDialogs, TestLogs, TestEvents
    for route in ["translation", "explanation", "reader"] {
        ctrl := route = "translation" ? ddlPrompt : ddlEPr
        srState := Map("gui", {Hwnd: 0}, "prompt", ctrl)
        action := route = "translation" ? DeletePromptProfile
            : route = "explanation" ? DeleteExplainPromptProfile : StudyReaderDeleteNewVersionPrompt.Bind(srState)
        pathFor := route = "translation" ? PromptFilePath : ExplainProfilePath
        otherPathFor := route = "translation" ? ExplainProfilePath : PromptFilePath
        prefix := route = "translation" ? "Translation" : "Explanation"
        ResetDeletion()
        ctrl.Text := ""
        action.Call()
        Check(TestDialogs.Length = 1 && TestToasts.Length = 0 && TestEvents.Length = 0,
            route ": empty selection never announces a deletion")

        for answer in ["No", "Cancel"] {
            name := route " cancelled " answer, path := pathFor.Call(name)
            IniWrite("fixture", path, "test", "text")
            ResetDeletion(, answer)
            ctrl.Text := name
            action.Call()
            Check(FileExist(path) && ctrl.Text = name, route ": declining preserves file and selection")
            Check(TestDialogs.Length = 1 && TestToasts.Length = 0 && TestEvents.Length = 0,
                route ": declining neither refreshes nor announces success")
        }

        for name in [route " fixture", route " 日本語 & English"] {
            path := pathFor.Call(name), otherPath := otherPathFor.Call(name)
            IniWrite("fixture", path, "test", "text")
            IniWrite("keep", otherPath, "test", "text")
            ResetDeletion()
            ctrl.Text := name
            action.Call()
            Check(!FileExist(path), route ": confirmed prompt is actually deleted")
            Check(FileExist(otherPath), route ": identically named prompt of other type is untouched")
            Check(TestToasts.Length = 1 && TestToasts[1] = prefix " prompt deleted: " name,
                route ": one success toast identifies type and original name")
            refreshCount := route = "reader" ? 2 : 1
            Check(TestEvents.Length = refreshCount + 1 && TestEvents[TestEvents.Length] = "toast",
                route ": success follows all necessary selector refreshes")
            Check(TestEvents[1] = route "-refresh" && (route != "reader" || TestEvents[2] = "explanation-refresh"),
                route ": refreshes the correct prompt selectors")
            Check(ctrl.Text != name, route ": success uses captured name, not refreshed selection")
        }

        name := route " locked", path := pathFor.Call(name)
        IniWrite("fixture", path, "test", "text")
        ResetDeletion()
        ctrl.Text := name
        handle := DllCall("kernel32\CreateFileW", "wstr", path, "uint", 0x80000000,
            "uint", 0, "ptr", 0, "uint", 3, "uint", 0x80, "ptr", 0, "ptr")
        Check(handle != -1 && handle != 0, route ": exclusive fixture lock succeeds")
        try action.Call()
        finally DllCall("kernel32\CloseHandle", "ptr", handle)
        Check(FileExist(path) && ctrl.Text = name, route ": failed deletion preserves file and selection")
        Check(TestToasts.Length = 0 && TestEvents.Length = 0, route ": failed deletion does not announce success")
        Check(TestDialogs.Length = 2 && InStr(TestDialogs[2], "Could not delete " StrLower(prefix) " prompt '" name "'"),
            route ": failed deletion shows an accurate error")
        Check(TestLogs.Length = 1, route ": failed deletion is logged")
    }
}

RefreshPromptProfilesList(*) {
    global TestEvents, ddlPrompt
    TestEvents.Push("translation-refresh")
    ddlPrompt.Text := "Remaining translation prompt"
}
RefreshExplainPromptProfilesList(*) {
    global TestEvents, ddlEPr
    TestEvents.Push("explanation-refresh")
    ddlEPr.Text := "Remaining explanation prompt"
}
StudyReaderRefreshNewVersionPromptList(state, *) {
    global TestEvents
    TestEvents.Push("reader-refresh")
    state["prompt"].Text := "Remaining reader prompt"
}
StudyReaderRefreshMainExplainPromptList() => RefreshExplainPromptProfilesList()
