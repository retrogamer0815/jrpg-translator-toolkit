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
global toasts := [], terminalMessage := "", terminalToken := "", closedChild := false
global warningCues := 0
global launchedEnvironment := Map(), launchedCommand := ""
global fixtureBindings := Map(), fixtureNavigating := false
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
    for scenario in ["launch-failed", "locked-error", "504", "404", "503", "429", "401", "no-result", "timeout", "slow-success", "success", "success"] {
        mode := scenario
        before := explanationRuns
        beforeMessages := messages.Length
        beforeCues := warningCues
        ExplainNow()
        if locked
            DllCall("CloseHandle", "ptr", locked), locked := 0
        Check(explanationRuns = before + 1, "Explanation can start: " scenario)
        Check(EnvGet("EXPLAIN_ERROR_FILE") = "prior synthetic value", "Explanation restores scoped environment: " scenario)
        Check(messages.Length = beforeMessages, "No blocking failure dialog: " scenario)
        if scenario != "success" && scenario != "slow-success" {
            Check(CPExplanationFailureStatus() != "", "Retained failure status: " scenario)
            Check(toasts[-1][2] = 15000 && toasts[-1][3] = 2, "Long priority failure toast: " scenario)
            Check(terminalToken = launchedEnvironment["JRPG_REQUEST_ID"] && terminalMessage != "", "Full details and completion retained: " scenario)
            Check(warningCues = beforeCues + 1, "One nonblocking warning sound on failure")
        } else {
            Check(CPExplanationFailureStatus() = "", "Successful retry clears failure status")
            Check(toasts[-1][1] = "Detailed explanation ready", "Only confirmed success says ready")
            Check(warningCues = beforeCues, "Success has no warning sound")
        }
    }
    Check(closedChild, "Hard timeout stops only its synthetic child")
    Check(InStr(CPExplanationFailureSummary("Gemini explanation failed (error 504). DEADLINE_EXCEEDED"), "timed out"), "504 is a timeout, not an absent model")
    Check(InStr(CPExplanationFailureSummary("Error 404 NOT_FOUND"), "unavailable"), "Unavailable model gets actionable summary")
    Check(InStr(CPExplanationFailureSummary("Error 429 RESOURCE_EXHAUSTED"), "limit"), "Quota gets a distinct summary")
    Check(InStr(CPExplanationFailureSummary("The model process ended without returning a result. Check the API key and connection."), "Check the Explainer"), "Generic failure is not misreported as authentication failure")
    Check(InStr(CPExplanationFailureSummary("File not found: synthetic.py"), "Check the Explainer"), "Local missing file is not misreported as unavailable model")
    mode := "signal-error"
    beforeMessages := messages.Length
    ExplainNow()
    Check(CPExplanationFailureStatus() != "" && messages.Length = beforeMessages, "Unexpected failure uses nonmodal recovery")
    mode := "recursive-explain"
    before := explanationRuns
    ExplainNow()
    Check(explanationRuns = before + 1, "Nested Explanation rejected while request active")
    FileAppend("Vocabulary and grammar: {jp}", A_ScriptDir "\quick-prompt.txt", "UTF-8")
    FileAppend("日本語", overlayDir "\last_jp.txt", "UTF-8")
    FileAppend("fixture image bytes", A_ScriptDir "\capture.png")
    FileAppend(A_ScriptDir "\capture.png", overlayDir "\last_src.txt", "UTF-8")
    IniWrite(1, iniPath, "cfg", "saveExplains")
    IniWrite(1, iniPath, "cfg", "saveStudyLibrary")
    IniWrite(0, iniPath, "quick_explanation", "useDetailedModel")
    IniWrite("gemini", iniPath, "quick_explanation", "provider")
    IniWrite("fixture-quick", iniPath, "quick_explanation", "geminiModel")
    EnvSet("SAVE_STUDY_LIBRARY", "original-archive-setting")
    EnvSet("GEMINI_EXPLAIN_MODEL", "original-gemini-model")
    for scenario in ["success", "launch-failed", "504", "timeout", "recursive-explain", "success"] {
        mode := scenario, before := explanationRuns
        QuickExplainNow()
        Check(explanationRuns = before + 1, "Quick can retry, rejecting overlapping detailed/quick: " scenario)
        Check(InStr(launchedCommand, " --quick-explanation"), "Quick uses explicit child command mode")
        Check(launchedEnvironment["SAVE_EXPLAINS"] = "0" && launchedEnvironment["SAVE_STUDY_LIBRARY"] = "0", "Quick disables only child saving")
        Check(launchedEnvironment["EXPLAIN_PROVIDER"] = "gemini" && launchedEnvironment["GEMINI_EXPLAIN_MODEL"] = "fixture-quick", "Quick model override reaches child")
        Check(EnvGet("SAVE_STUDY_LIBRARY") = "original-archive-setting" && EnvGet("GEMINI_EXPLAIN_MODEL") = "original-gemini-model", "Quick restores parent environment after success/failure")
        Check(!FileExist(launchedEnvironment["EXPLAIN_SOURCE_TEXT_FILE"]), "Request snapshot cleaned after completion/failure")
        Check(IniRead(iniPath, "cfg", "saveStudyLibrary") = "1", "Detailed saving preference preserved")
        if scenario = "success"
            Check(toasts[-1][1] = "Quick explanation ready", "Quick completion has distinct notice")
    }
    mode := "success"
    ExplainNow()
    Check(!InStr(launchedCommand, "--quick-explanation") && launchedEnvironment["SAVE_STUDY_LIBRARY"] = "1", "Detailed after Quick keeps its mode and archive saving")
    Check(launchedEnvironment["EXPLAIN_PROVIDER"] = "openai" && launchedEnvironment["EXPLAIN_MODEL"] = "fixture", "Quick does not alter detailed model")
    IniWrite(1, iniPath, "quick_explanation", "useDetailedModel")
    QuickExplainNow()
    Check(launchedEnvironment["EXPLAIN_PROVIDER"] = "openai" && launchedEnvironment["EXPLAIN_MODEL"] = "fixture", "Quick can reuse current detailed model")
    Rebind_QuickExplanation()
    Check(fixtureBindings.Count = 0, "Quick shortcut starts unassigned")
    IniWrite("^!F24", iniPath, "hotkeys", "quick_explanation")
    Rebind_QuickExplanation()
    before := explanationRuns
    fixtureBindings["^!F24"].Call()
    Check(explanationRuns = before + 1 && InStr(launchedCommand, "--quick-explanation"), "Configured keyboard shortcut dispatches Quick")
    fixtureNavigating := true
    fixtureBindings["^!F24"].Call()
    Check(explanationRuns = before + 1, "Quick gameplay shortcut is suppressed during UI navigation")
    fixtureNavigating := false
    IniWrite("", iniPath, "hotkeys", "quick_explanation")
    Rebind_QuickExplanation()
    Check(fixtureBindings.Count = 0, "Disabling Quick unregisters its previous shortcut")
    CPControllerDispatchAction("quick_explanation")
    Check(explanationRuns = before + 2 && InStr(launchedCommand, "--quick-explanation"), "Direct controller binding dispatches Quick")
    CPControllerDispatchAction("explain_last_translation")
    Check(explanationRuns = before + 3 && !InStr(launchedCommand, "--quick-explanation"), "Existing controller binding still dispatches Detailed")

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
TestHotkey(key, callback) {
    global fixtureBindings
    if callback = "Off"
        fixtureBindings.Delete(key)
    else
        fixtureBindings[key] := callback
}
CPControllerNavigationTarget() => fixtureNavigating
SafeCall(callback) => callback.Call()
CPControllerTranslatorAction(*) => 0
CPControllerToggleTranslator(*) => 0
CPControllerToggleExplainer(*) => 0
ToggleControlPanel(*) => 0
CP_LaunchExplainerRequest(*) => 0
StartStopAudio(*) => 0
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
    global explanationRuns, explanationPolls, mode, locked, overlayDir, launchedEnvironment, launchedCommand
    launchedCommand := cmd
    for key in ["JRPG_REQUEST_ID", "EXPLAIN_PROVIDER", "EXPLAIN_MODEL", "GEMINI_EXPLAIN_MODEL", "SAVE_EXPLAINS", "SAVE_STUDY_LIBRARY", "EXPLAIN_SOURCE_TEXT_FILE", "EXPLAIN_SOURCE_PATHS_FILE"]
        launchedEnvironment[key] := EnvGet(key)
    if FileExist(overlayDir "\last_jp.txt") {
        Check(FileRead(launchedEnvironment["EXPLAIN_SOURCE_TEXT_FILE"], "UTF-8") = "日本語", "Request snapshots Japanese source")
        copiedImage := Trim(FileRead(launchedEnvironment["EXPLAIN_SOURCE_PATHS_FILE"], "UTF-8"), "`r`n")
        Check(copiedImage != A_ScriptDir "\capture.png" && FileRead(copiedImage) = "fixture image bytes", "Request owns an independent image snapshot")
    }
    explanationRuns += 1
    explanationPolls := 0
    if mode = "launch-failed"
        throw OSError(5)
    if mode = "locked-error" {
        path := EnvGet("EXPLAIN_ERROR_FILE")
        FileAppend("Synthetic error", path, "UTF-8")
        locked := LockFile(path)
    } else if mode != "no-result" && mode != "timeout" {
        done := FileOpen(overlayDir "\explainer.done", "w", "UTF-8")
        done.Write(EnvGet("JRPG_REQUEST_ID")), done.Close()
        if RegExMatch(mode, "^\d{3}$")
            FileAppend("Gemini explanation failed (error " mode ").`nDetails: synthetic provider failure", EnvGet("EXPLAIN_ERROR_FILE"), "UTF-8")
    }
    processId := 0
}
TestExplainAlive(*) {
    global explanationPolls, mode
    explanationPolls += 1
    return mode = "timeout" || (mode = "slow-success" && explanationPolls = 1)
}
TestExplainStartedAt() => mode = "timeout" ? A_TickCount - CPModelRequestTimeoutMs("explanation") - 1
    : mode = "slow-success" ? A_TickCount - 120001 : A_TickCount
TestExplainClose(*) {
    global closedChild
    closedChild := true
}
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
CPWriteExplainerTerminalResult(message, token) {
    global terminalMessage, terminalToken
    terminalMessage := message, terminalToken := token
}
CPRefreshExplanationFailureStatus(*) => 0
TestWarningCue(*) {
    global warningCues
    warningCues += 1
}
CPSyncExplanationSelectionFromControls(*) => "openai"
ExplainProfilePath(*) => ""
ReadIniInt(path, section, key, fallback) => Integer(IniRead(path, section, key, fallback))
CPQuickExplanationPromptPath(*) => A_ScriptDir "\quick-prompt.txt"
StudyLibraryCurrentChapter(*) => ""
GameProfileSafeName(value) => value
Toast(args*) {
    global toasts
    toasts.Push(args)
}
SignalExplainerBusy(*) {
    global mode
    if mode = "recursive-explain" {
        ExplainNow()
        QuickExplainNow()
    }
    if mode = "signal-error"
        throw Error("Synthetic signal failure")
}
