#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
; Isolated child-process fixture: never loads audio drivers or makes API calls.
mode := EnvGet("JRPG_TEST_AUDIO_MODE")
if mode = "hang" {
    Sleep(10000)
    ExitApp(0)
}
Sleep(120)
listing := A_Args.Length && A_Args[1] = "--list-speakers"
resultPath := EnvGet(listing ? "AUDIO_DEVICE_LIST_RESULT_FILE" : "AUDIO_TEST_RESULT_FILE")
if mode = "missing"
    ExitApp(1)
if listing {
    result := mode = "error" ? "JRPG_AUDIO_DEVICES:ERROR`nSimulated driver error`n"
        : "JRPG_AUDIO_DEVICES:OK`nGame speakers`nHeadphones`n日本語 ＆ 音声`n"
} else {
    result := mode = "silent" ? "JRPG_AUDIO_TEST:SILENT:peak=0.000000`n"
        : mode = "error" ? "JRPG_AUDIO_TEST:ERROR:Simulated device error`n"
        : "JRPG_AUDIO_TEST:DETECTED:peak=0.500000`n"
}
FileAppend(result, resultPath, "UTF-8-RAW")
ExitApp(mode = "error" ? 1 : 0)
