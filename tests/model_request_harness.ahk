#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
global assertions := 0, notices := [], warningCues := 0, overlayDir := A_ScriptDir
global OcrTxt := A_ScriptDir "\ocr.txt", OcrDoneTxt := A_ScriptDir "\ocr.done"
try {
    for mode in ["success", "failure", "hang", "batches", "repeated", "cancel", "closed", "ownerClosed"] {
        state := Map("closed", false), progress := A_ScriptDir "\" mode ".progress"
        cmd := Format('"{1}" /ErrorStdOut "{2}" "{3}" "{4}"', A_AhkPath,
            A_ScriptDir "\model_request_worker.ahk", mode, progress)
        if mode = "cancel"
            SetTimer(CPModelRequestCancel.Bind(state), -100)
        if mode = "closed"
            SetTimer((*) => state["closed"] := true, -100)
        if mode = "ownerClosed" {
            state["readerState"] := Map("closed", false)
            SetTimer((*) => state["readerState"]["closed"] := true, -100)
        }
        started := A_TickCount
        result := CPModelRequestRun(cmd, state, progress, 300)
        expected := mode = "success" || mode = "batches" ? 0 : mode = "failure" ? 7
            : mode = "hang" || mode = "repeated" ? -3 : -2
        Check(result = expected, mode " returns the correct outcome: " result)
        Check(!state.Has("modelJob") && CPModelRequestJobs().Count = 0, mode " releases handles and job state")
        Check(A_TickCount - started < 3000, mode " bounded completion")
        pid := CPModelRequestRead(progress ".pid")
        Check(pid != "" && !ProcessExist(Integer(pid)), mode " leaves no running child")
        Check(!FileExist(progress ".late"), mode " cannot write a late result")
        if mode = "batches"
            Check(A_TickCount - started > 450, "Several healthy batches may exceed one batch deadline")
        if mode = "hang"
            Check(InStr(CPModelRequestError(state, ""), "timed out"), "Timeout has actionable diagnostic")
    }
    state := Map()
    try {
        CPModelRequestRun('"Z:\missing-jrpg-fixture.exe"', state)
        throw Error("Expected a launch error")
    } catch OSError {
        Check(!state.Has("modelJob"), "Launch failure releases guard")
    }
    locked := A_ScriptDir "\locked.txt"
    FileAppend("locked", locked)
    handle := DllCall("kernel32\CreateFileW", "str", locked, "uint", 0x80000000,
        "uint", 0, "ptr", 0, "uint", 3, "uint", 0, "ptr", 0, "ptr")
    try Check(CPModelRequestRead(locked, "unreadable") = "unreadable", "Locked diagnostics cannot raise an AHK error")
    finally DllCall("kernel32\CloseHandle", "ptr", handle)
    Check(CPModelRequestRead(A_ScriptDir "\missing.txt", "missing") = "missing", "Missing diagnostic fallback")
    for message in ["504 DEADLINE_EXCEEDED", "timed out", "408 request timeout"]
        Check(InStr(CPModelFailureSummary("Example sentence", message), "timed out"), "Classifies timeout")
    s := Map("closed", true)
    CPModelRequestFailure(s, "Recommendations", "504 DEADLINE_EXCEEDED")
    Check(notices.Length = 1 && warningCues = 1, "A generation failure emits a visible warning and sound")
    Check(s["modelFailureDetails"] = "504 DEADLINE_EXCEEDED", "Full diagnostics retained in state")
    Check(notices[1][2] = 15000, "Notification duration is fifteen seconds")
    WriteStatus("request-1`n504 DEADLINE_EXCEEDED")
    CPPollTranslationStatus()
    Check(InStr(CPTranslationFailureStatus(), "timed out"), "Translation failure retained outside hidden overlay")
    count := notices.Length
    CPPollTranslationStatus()
    Check(notices.Length = count, "Failure not notified twice")
    WriteStatus("request-2`n")
    CPPollTranslationStatus()
    Check(CPTranslationFailureStatus() = "" && notices.Length = count, "Success clears stale failure without error toast")
    WriteStatus("request-3`n503 unavailable")
    CPPollTranslationStatus()
    Check(notices.Length = count + 1, "Next failure is independently notified")
    WriteTranslationTerminalResult("Translation timed out.", "request-4")
    CPPollTranslationStatus()
    Check(notices.Length = count + 2, "Overlay watchdog failure reaches the main notification")
    Check(CPModelRequestRead(OcrDoneTxt) = "request-4", "Watchdog completion token matches notification")
    Check(CPModelRequestRead(OcrTxt) = "Translation timed out.", "Watchdog retains error details in the overlay")
    TestRecommendationBridgeFailures()
    FileAppend("PASS: " assertions " model request supervision assertions.`n", "*")
    ExitApp(0)
} catch as err {
    FileAppend("FAIL: " err.Message " at " err.Line "`n" err.Stack "`n", "*")
    ExitApp(1)
}
Check(condition, label) {
    global assertions
    assertions += 1
    if !condition
        throw Error(label)
}
WriteStatus(text) {
    global overlayDir
    try FileDelete(overlayDir "\translation.status")
    FileAppend(text, overlayDir "\translation.status", "UTF-8")
}
StudyLibraryStateAlive(state) => !state.Get("closed", false)
Toast(message, duration, position) {
    global notices
    notices.Push([message, duration, position])
}
TestWarningCue() {
    global warningCues
    warningCues += 1
}
CPBigBoxUpdateAIContent() {
}
CPBigBoxUpdateSettingsHint() {
}

TestRecommendationBridgeFailures() {
    global notices
    for mode in ["failure", "providerTimeout", "hang", "cancel", "launch", "success"] {
        state := Map("closed", false, "closeRequested", false, "busyCount", 0,
            "outputDir", A_ScriptDir, "generatingRecommendations", true)
        progress := A_ScriptDir "\bridge-" mode ".progress"
        state["command"] := mode = "launch" ? '"Z:\missing-jrpg-fixture.exe"'
            : Format('"{1}" /ErrorStdOut "{2}" "{3}" "{4}"', A_AhkPath,
                A_ScriptDir "\model_request_worker.ahk", mode, progress)
        errorPath := A_ScriptDir "\candidate_recommendation_error.txt"
        try FileDelete(errorPath)
        FileAppend("Stale error from a previous request", errorPath)
        count := notices.Length
        if mode = "cancel"
            SetTimer((*) => (state["closeRequested"] := true, CPModelRequestCancel(state)), -100)
        result := StudyCandidatesRunBridge(state, "generate-recommendations")
        Check(result = (mode = "success"), "Recommendation bridge result: " mode)
        Check(state["busyCount"] = 0 && state["controlsEnabled"], "Recommendation controls released: " mode)
        Check(!state.Has("modelJob"), "Recommendation job released: " mode)
        Check(notices.Length = count + (mode = "cancel" || mode = "success" ? 0 : 1),
            "Recommendation failure notified once, cancellation silently: " mode)
        if mode = "providerTimeout" || mode = "hang"
            Check(InStr(state["modelFailureSummary"], "timed out"), "Recommendation timeout is actionable: " mode)
        if mode = "failure"
            Check(!InStr(state["modelFailureDetails"], "Stale"), "An early bridge failure cannot report stale diagnostics")
    }
}
TestBridgeRun(command, state, progress) => CPModelRequestRun(command, state, progress, 300)
StudyCandidatesBridgeCommand(state, *) => Map("ok", true, "command", state["command"])
StudyCandidatesGuiAlive(state) => !state.Get("closed", false)
StudyCandidatesBeginWork(state) {
    state["busyCount"] += 1
    state["controlsEnabled"] := false
    return true
}
StudyCandidatesEndWork(state) {
    state["busyCount"] -= 1
    state["controlsEnabled"] := true
}
StudyCandidatesReportBridgeUnavailable(*) => CPThemedOwnedMessage()
CPThemedOwnedMessage(*) {
    throw Error("Model failures must never open a blocking dialog")
}
