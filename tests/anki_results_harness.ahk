#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
global TestRows := [], TestMessages := []
try {
    state := Map("outputDir", A_ScriptDir)
    result := StudyAnkiReadAddResult(state)
    Assert(result["code"] = "uncertain" && InStr(result["message"], "may already"), "Missing bridge result is uncertain")
    TestRows := [["uncertain", "", "Synthetic warning"]]
    result := StudyAnkiReadAddResult(state)
    Assert(result["message"] = "Synthetic warning", "Bridge uncertainty message survives decoding")
    review := Map("status", {Value: ""}, "addButton", {Enabled: true}, "exampleButton", {Enabled: true})
    StudyAnkiShowUncertainAdd(review, state, result["message"])
    Assert(review["uncertain"] && !review["addButton"].Enabled, "Uncertain review prevents retry")
    Assert(!review["exampleButton"].Enabled, "Example generator cannot re-enable Add")
    Assert(review["status"].Value = "Synthetic warning" && TestMessages.Length = 1, "Status and warning are displayed")
    review.Delete("exampleButton")
    StudyAnkiShowUncertainAdd(review, state, result["message"])
    Assert(TestMessages.Length = 2, "Explanation-only review needs no vocabulary button")
    for code in ["added", "duplicate", "added_unlinked", "blocked"] {
        TestRows := [[code, "123", "result message"]]
        result := StudyAnkiReadAddResult(state)
        Assert(result["code"] = code && result["noteId"] = "123", "Existing result codes preserved")
    }
    FileAppend("PASS: 10 Anki result/UI-state assertions.`n", "*")
    ExitApp(0)
} catch as testFailure {
    FileAppend("FAIL: " testFailure.Message "`n" testFailure.Stack "`n", "*")
    ExitApp(1)
}
Assert(condition, message) {
    if !condition
        throw Error(message)
}
StudyLibraryReadRows(*) {
    global TestRows
    return TestRows
}
StudyLibraryHexDecode(value) => value
StudyReaderAnkiMessage(reader, message, *) {
    global TestMessages
    TestMessages.Push(message)
}
