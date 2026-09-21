#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn All, StdOut
global checks := 0, mode := "", notices := [], messages := [], launches := 0, locked := 0
global envPath := A_ScriptDir "\synthetic-api.env", envSavedOpenAI := "old", envSavedGemini := "old"
global eOpenAI := {Value: "synthetic-new"}, eGemini := {Value: "synthetic-gemini"}, cbApiInApp := {Value: 1}
global ShotBuf := [], pythonExe := A_ScriptFullPath, translatorPy := A_ScriptFullPath, explainScript := A_ScriptFullPath
global ControlIni := A_ScriptDir "\synthetic-control.ini", iniPath := ControlIni
global __TranslationPid := 0, __TranslationRequestId := "", __TranslationStartedAt := 0, __TranslationExitedAt := 0
global keysConfigured := true, tip := "", debugMode := false, explainsDir := A_ScriptDir, studyLibraryDir := A_ScriptDir, overlayDir := A_ScriptDir
global explainProvider := "openai", explainOpenAIModel := "fixture", explainGeminiModel := "fixture", ddlEPr := {Text: "fixture"}
global destroyOnTheme := false, explanationRuns := 0
try {
    FileAppend("OLD SYNTHETIC FILE", envPath, "UTF-8")
    mode := "save-denied"
    SaveApiEnv()
    Check(FileRead(envPath, "UTF-8") = "OLD SYNTHETIC FILE", "Denied replacement preserves old keys")
    Check(envSavedOpenAI = "old", "Denied replacement leaves saved-state cache alone")
    locked := LockFile(envPath)
    mode := ""
    SaveApiEnv()
    DllCall("CloseHandle", "ptr", locked), locked := 0
    Check(FileRead(envPath, "UTF-8") = "OLD SYNTHETIC FILE", "Exclusive destination lock preserves old keys")
    mode := "save-reentrant"
    SaveApiEnv()
    Check(InStr(FileRead(envPath, "UTF-8"), "OPENAI_API_KEY=synthetic-new") > 0, "Later save succeeds after errors")
    Check(envSavedOpenAI = "synthetic-new", "Saved-state cache advances after commit")
    Check(notices.Length = 1, "Reentrant save rejected without duplicate commit")
    leftovers := 0
    Loop Files envPath ".*.tmp"
        leftovers += 1
    Check(leftovers = 0, "Only temporary copies are cleaned up")

    for rejection in ["missing-script", "missing-python", "missing-key", "busy", "launch-failed"] {
        mode := rejection
        pythonExe := rejection = "missing-python" ? A_ScriptDir "\missing.exe" : A_ScriptFullPath
        translatorPy := rejection = "missing-script" ? A_ScriptDir "\missing.py" : A_ScriptFullPath
        keysConfigured := rejection != "missing-key"
        __TranslationPid := rejection = "busy" ? DllCall("GetCurrentProcessId", "uint") : 0
        ShotBuf := ["one.png", "two.png"]
        FlushBufferedScreenshots()
        Check(ShotBuf.Length = 2 && ShotBuf[1] = "one.png", rejection " retains queued captures")
        Check(!InStr(tip, "Sent 2"), rejection " does not falsely report sent")
    }
    mode := "append-during-launch", __TranslationPid := 0
    FlushBufferedScreenshots()
    Check(ShotBuf.Length = 1 && ShotBuf[1] = "new.png", "Accepted snapshot preserves later captures")
    Check(tip = "Sent 2 screenshot(s).", "Accepted request reports sent")
    mode := "", __TranslationPid := 0
    Check(flushTranslate(["oneshot.png"]), "One-shot accepted")
    Check(ShotBuf.Length = 1 && ShotBuf[1] = "new.png", "One-shot does not discard buffered captures")
    ShotBuf := ["one.png", "one.png", "two.png"]
    RemoveAcceptedCaptures(["one.png"])
    Check(ShotBuf.Length = 2 && ShotBuf[1] = "one.png", "Only accepted occurrence removed")

    EnvSet("EXPLAIN_ERROR_FILE", "prior synthetic value")
    for scenario in ["launch-failed", "locked-error", "success", "success"] {
        mode := scenario
        before := explanationRuns
        ExplainNow()
        if locked
            DllCall("CloseHandle", "ptr", locked), locked := 0
        Check(explanationRuns = before + 1, "Explanation can start: " scenario)
        Check(EnvGet("EXPLAIN_ERROR_FILE") = "prior synthetic value", "Explanation restores scoped environment: " scenario)
    }
    mode := "recursive-explain"
    before := explanationRuns
    ExplainNow()
    Check(explanationRuns = before + 1, "Nested Explanation rejected while request active")

    Loop 30 {
        dialog := Gui()
        dialog.AddText(, "Synthetic lifecycle fixture")
        dialog.Show("Hide w260 h100")
        destroyOnTheme := Mod(A_Index, 2) = 0
        beforeCritical := A_IsCritical
        revealed := StudyWindowRevealFinished(dialog, "x-10000 y-10000")
        Check(revealed = !destroyOnTheme, "Reveal tolerates native destruction")
        Check(A_IsCritical = beforeCritical, "Reveal restores thread interruptibility")
        try dialog.Destroy()
    }
    FileAppend("PASS: " checks " reliability assertions (synthetic I/O and real locks/windows).`n", "*")
    ExitApp(0)
} catch as fixtureError {
    if locked
        DllCall("CloseHandle", "ptr", locked)
    FileAppend("FAIL: " fixtureError.Message " at " fixtureError.Line "`n" fixtureError.Stack "`n", "*")
    ExitApp(1)
}
Check(condition, message) {
    global checks
    checks += 1
    if !condition
        throw Error(message)
}
LockFile(path) {
    handle := DllCall("CreateFileW", "str", path, "uint", 0x80000000, "uint", 0, "ptr", 0,
        "uint", 3, "uint", 0x80, "ptr", 0, "ptr")
    if handle = -1
        throw OSError(A_LastError)
    return handle
}
TestReplaceApiEnvFile(temporary, destination) {
    global mode
    if mode = "save-denied"
        throw OSError(5)
    if mode = "save-reentrant"
        SaveApiEnv()
    CPReplaceApiEnvFile(temporary, destination)
}
TestTranslateRun(cmd, directory, flags, &processId) {
    global mode, ShotBuf, launches
    launches += 1
    if mode = "launch-failed"
        throw OSError(5)
    if mode = "append-during-launch"
        ShotBuf.Push("new.png")
    processId := 0
}
TestExplainRun(cmd, directory, flags, &processId) {
    global explanationRuns, mode, locked, overlayDir
    explanationRuns += 1
    if mode = "launch-failed"
        throw OSError(5)
    if mode = "locked-error" {
        path := EnvGet("EXPLAIN_ERROR_FILE")
        FileAppend("Synthetic error", path, "UTF-8")
        locked := LockFile(path)
    } else {
        done := FileOpen(overlayDir "\explainer.done", "w", "UTF-8")
        done.Write(EnvGet("JRPG_REQUEST_ID")), done.Close()
    }
    processId := 0
}
TestExplainAlive(*) => false
CPApplyOwnedDialogTheme(dialog) {
    global destroyOnTheme
    if destroyOnTheme
        dialog.Destroy()
}
CPAdaptiveOwnedMessage(args*) {
    global messages
    messages.Push(args)
}
CPDialogDefaultOwner(*) => 0
ToggleApiKeyControls(*) => 0
UpdateEnvDirty(*) => 0
CPApiKeysSetNotice(message) {
    global notices
    notices.Push(message)
}
ResolvePath(value) => value
RegisterScreenshotForStartupCleanup(*) => 0
OverlayApiKeyConfigured(*) => keysConfigured
CPApiKeyConfigured(*) => keysConfigured
OverlayMissingApiKeyText(*) => "Synthetic missing key"
CPShowMissingApiKey(*) => 0
showText(*) => 0
Dbg(*) => 0
DbgCP(*) => 0
TestTip(value := "") {
    global tip
    tip := value
}
ShowOverlayStatus(*) => 0
ExportGlossaryEnv(*) => 0
WriteTranslationTerminalResult(*) => 0
CPWriteExplainerTerminalResult(*) => 0
CPSyncExplanationSelectionFromControls(*) => "openai"
ExplainProfilePath(*) => ""
ReadIniInt(path, section, key, fallback) => fallback
StudyLibraryCurrentChapter(*) => ""
GameProfileSafeName(value) => value
Toast(*) => 0
SignalExplainerBusy(*) {
    global mode
    if mode = "recursive-explain"
        ExplainNow()
}
