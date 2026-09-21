#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
global TestAssertions := 0, TestToasts := [], TestLog := []
global TestRunning := false, TestResult := true, TestFailure := ""
global TestReenter := false, TestNestedResult := "", TestChecks := []
global TestStarts := 0, TestStops := 0

try {
    TestShortcutNotifications()
    FileAppend("PASS: " TestAssertions " audio shortcut notification assertions (" A_PtrSize * 8 "-bit).`n", "*")
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

ResetAudio(running := false, result := true, failure := "") {
    global TestToasts, TestLog, TestRunning, TestResult, TestFailure
    global TestReenter, TestNestedResult, TestChecks, TestStarts, TestStops
    TestToasts := [], TestLog := [], TestChecks := []
    TestRunning := running, TestResult := result, TestFailure := failure
    TestReenter := false, TestNestedResult := "", TestStarts := 0, TestStops := 0
}

TestShortcutNotifications() {
    global TestToasts, TestLog, TestRunning, TestReenter, TestNestedResult
    global TestChecks, TestStarts, TestStops
    for viaController in [false, true] {
        for stopping in [false, true] {
            for succeeded in [false, true] {
                ResetAudio(stopping, succeeded)
                if viaController
                    CPControllerDispatchAction("start_stop_audio")
                else
                    shortcutResult := StartStopAudio("synthetic hotkey")
                expected := succeeded
                    ? (stopping ? "Audio Translation Off" : "Audio Translation On")
                    : (stopping ? "Could not stop audio translation" : "Could not start audio translation")
                Check(TestToasts.Length = 1 && TestToasts[1] = expected,
                    "Exactly one accurate notification for each shortcut outcome")
                if !viaController
                    Check(shortcutResult = succeeded, "Hotkey returns the operation outcome")
                Check(TestStarts = !stopping && TestStops = stopping, "Only the requested transition runs")
                Check(TestChecks.Length = 1 && TestChecks[1], "Shortcut retains verified recovery scan")
                Check(TestRunning = (succeeded ? !stopping : stopping), "Failed action does not change state")
            }
        }
    }

    for failure in ["check", "start", "stop"] {
        ResetAudio(failure = "stop", true, failure)
        Check(!StartStopAudio(), "Unexpected audio error returns failure")
        Check(TestToasts.Length = 1 && TestToasts[1] = "Could not change audio translation",
            "Unexpected error never claims success")
        Check(TestLog.Length = 1, "Unexpected failure is logged")
        ResetAudio()
        Check(StartStopAudio() && TestStarts = 1, "Shortcut guard released after error")
    }

    for stopping in [false, true] {
        ResetAudio(stopping)
        TestReenter := true
        Check(StartStopAudio(), "Outer transition succeeds during duplicate shortcut")
        Check(TestNestedResult = false && TestStarts + TestStops = 1, "Reentrant shortcut cannot toggle twice")
        Check(TestToasts.Length = 1, "Reentrant shortcut does not duplicate notification")
    }

    for stopping in [false, true] {
        ResetAudio(stopping)
        ToggleAudioFromButton()
        Check(TestToasts.Length = 0, "UI button remains silent")
        Check(TestRunning = !stopping, "UI button still toggles")
    }
    ResetAudio()
    Loop 20 {
        Check(StartStopAudio(), "Repeated shortcut succeeds")
        Check(TestToasts[A_Index] = (Mod(A_Index, 2) ? "Audio Translation On" : "Audio Translation Off"),
            "Repeated shortcuts announce alternating confirmed states")
    }
}

; Only worker/status/toast boundaries are simulated. The actual shortcut,
; controller dispatch, button handler and start/stop wrappers are extracted
; from production. No audio devices, provider keys, or user processes are used.
AudioIsRunning(allowRecoveryScan := false) {
    global TestRunning, TestFailure, TestChecks
    TestChecks.Push(allowRecoveryScan)
    if TestFailure = "check"
        throw Error("Synthetic status failure")
    return TestRunning
}
StartAudioCore(bigBox := false) {
    global TestStarts
    TestStarts += 1
    return TestTransition(false)
}
StopAudioCore(announce := true, recoveryScan := true) {
    global TestStops
    TestStops += 1
    return TestTransition(true)
}
TestTransition(stopping) {
    global TestRunning, TestResult, TestFailure, TestReenter, TestNestedResult
    if TestFailure = (stopping ? "stop" : "start")
        throw Error("Synthetic worker failure")
    if TestReenter {
        TestReenter := false
        TestNestedResult := StartStopAudio()
    }
    if TestResult
        TestRunning := !stopping
    return TestResult
}
Toast(message) {
    global TestToasts
    TestToasts.Push(message)
}
DbgCP(message) {
    global TestLog
    TestLog.Push(message)
}
CPControllerTranslatorAction(*) => 0
ExplainNow(*) => 0
CPControllerToggleTranslator(*) => 0
CPControllerToggleExplainer(*) => 0
ToggleControlPanel(*) => 0
CP_LaunchExplainerRequest(*) => 0
