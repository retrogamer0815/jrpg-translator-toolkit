#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
global gPidAudio := 0, gAudioProcess := 0, gLastAction := "", gJustStoppedUntil := 0
global pythonExe := A_AhkPath, audioScript := A_ScriptDir "\audio_runtime_fixture.ahk"
global gAudioSessionFile := A_ScriptDir "\audio-session.pid"
global TestCommands := Map(), TestCommandFailure := false, TestAssertions := 0
global TestRecovery := []
EnvSet("JRPG_TEST_AUDIO_RUNTIME_LOG", A_ScriptDir "\child.log")
EnvSet("JRPG_TEST_AUDIO_RUNTIME_MODE", "wait")
try {
    TestOwnership()
    FileAppend("PASS: " TestAssertions " audio process-ownership assertions (" A_PtrSize * 8 "-bit).`n", "*")
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

WriteMarker(value) {
    global gAudioSessionFile
    try FileDelete(gAudioSessionFile)
    FileAppend(value, gAudioSessionFile, "UTF-8-RAW")
}

TestOwnership() {
    global gPidAudio, gAudioProcess, pythonExe, audioScript, gAudioSessionFile
    global TestCommands, TestCommandFailure, TestRecovery
    benign := AudioLaunchProcess(pythonExe, audioScript)
    worker := AudioLaunchProcess(pythonExe, audioScript)
    TestCommands[worker["pid"]] := '"' pythonExe '" "' audioScript '"'
    TestCommands[benign["pid"]] := '"' pythonExe '" "' A_ScriptDir '\unrelated.ahk"'
    try {
        Check(AudioProcessAlive(worker) && AudioProcessAlive(benign), "Synthetic children started")
        for marker in [String(benign["pid"]), "999999999999999999999999999999", "bad marker", "JRPG_AUDIO_V2`t4294967296`t0000000000000000"] {
            WriteMarker(marker)
            gPidAudio := benign["pid"] ; Poisoned cached PID is never authority.
            Check(!AudioIsRunning(), "Invalid/unrelated marker is not adopted")
            Check(StopAudioCore(false, false) && AudioProcessAlive(benign), "Stop never kills poisoned PID")
        }
        WriteMarker(benign["marker"])
        Check(!AudioIsRunning(), "Correct executable and timestamp with wrong script is rejected")
        Check(AudioProcessAlive(benign), "Unrelated same-executable process survives")
        staleMarker := "JRPG_AUDIO_V2`t" worker["pid"] "`t0000000000000000"
        Check(!AudioOpenVerifiedProcess(worker["pid"], staleMarker), "PID reuse/creation-time mismatch rejected")
        originalPython := pythonExe
        pythonExe := A_ScriptDir "\other-python.exe"
        Check(!AudioOpenVerifiedProcess(worker["pid"], worker["marker"]), "Wrong executable path rejected")
        pythonExe := originalPython
        TestCommandFailure := true
        Check(!AudioOpenVerifiedProcess(worker["pid"], worker["marker"]), "Unavailable process inspection fails closed")
        TestCommandFailure := false
        command := TestCommands[worker["pid"]]
        Check(AudioCommandMatches(command, pythonExe, audioScript), "Exact two-argument launch accepted")
        for other in [command " --test-audio", command " --list-speakers",
            '"' pythonExe '" "' audioScript '.bak"', '"' pythonExe '" -c "' audioScript '"',
            '"' pythonExe '" "' audioScript ' extra"']
            Check(!AudioCommandMatches(other, pythonExe, audioScript), "Substring/diagnostic command rejected")

        WriteMarker(worker["marker"])
        Check(AudioIsRunning() && gPidAudio = worker["pid"], "Validated marker adopts held worker handle")
        ; Prove status uses that HANDLE, not a subsequently changed cached PID.
        gPidAudio := benign["pid"]
        Check(AudioIsRunning(), "Owned handle governs status despite poisoned display PID")
        Check(StopAudioCore(false, false), "Verified recovered worker stops")
        Check(!AudioProcessAlive(worker) && AudioProcessAlive(benign), "Only verified worker terminates")
        Check(!FileExist(gAudioSessionFile), "Stopped worker's own marker is cleared")

        ; An exited handle never switches to another process if a PID is reused.
        AudioSetProcess(AudioProcessRecord(DllCall("kernel32\OpenProcess", "uint", 0x101001,
            "int", false, "uint", benign["pid"], "ptr"), benign["pid"]))
        AudioSetProcess() ; Exercise releasing a live handle without termination.
        Check(AudioProcessAlive(benign), "Releasing tracking does not kill the process")
        AudioSetProcess(worker)
        gPidAudio := benign["pid"]
        Check(StopAudioCore(false, false) && AudioProcessAlive(benign), "Exited handle cannot redirect termination via cached PID")
        newerMarker := benign["marker"]
        WriteMarker(newerMarker)
        AudioSessionClear(staleMarker)
        Check(FileRead(gAudioSessionFile, "UTF-8") = newerMarker, "Clear preserves another session's marker")
        try FileDelete(gAudioSessionFile)

        ; Open/verify/release paths must not leak native process handles.
        countBefore := 0, countAfter := 0
        DllCall("kernel32\GetProcessHandleCount", "ptr", -1, "uint*", &countBefore)
        Loop 100 {
            Check(!AudioOpenVerifiedProcess(benign["pid"], newerMarker), "Repeated wrong-script rejection")
            AudioStartOwnedProcess(pythonExe, audioScript)
            Check(StopAudioCore(false, false), "Repeated owned launch/stop")
        }
        DllCall("kernel32\GetProcessHandleCount", "ptr", -1, "uint*", &countAfter)
        Check(countAfter <= countBefore + 2, "Process handles stable: " countBefore " -> " countAfter)
    } finally {
        AudioSetProcess()
        for process in [worker, benign] {
            if AudioProcessAlive(process) {
                DllCall("kernel32\TerminateProcess", "ptr", process["handle"], "uint", 1)
                DllCall("kernel32\WaitForSingleObject", "ptr", process["handle"], "uint", 1000)
            }
            AudioProcessRelease(process)
        }
    }
}

; Only OS command-line inspection is stubbed: actual handles, creation times,
; image paths, liveness and termination all use production Win32 code.
AudioProcessCommandLine(pid) {
    global TestCommands, TestCommandFailure
    if TestCommandFailure
        throw Error("Synthetic process-inspection denial")
    return TestCommands.Get(pid, "")
}
AudioProcessesByScript() {
    global TestRecovery
    return TestRecovery
}
ResolvePath(path) => path
DbgCP(*) => 0
_UpdateStatus(*) => 0
