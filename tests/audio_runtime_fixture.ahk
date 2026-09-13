#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
; A short-lived test child: never loads Python, audio drivers or model clients.
FileAppend("started`n", EnvGet("JRPG_TEST_AUDIO_RUNTIME_LOG"), "UTF-8")
if EnvGet("JRPG_TEST_AUDIO_RUNTIME_MODE") = "exit"
    DllCall("kernel32\ExitProcess", "uint", 1)
Sleep(10000)
ExitApp(0)
