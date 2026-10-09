#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
mode := A_Args[1], progress := A_Args[2]
FileAppend(DllCall("kernel32\GetCurrentProcessId", "uint"), progress ".pid")
if mode = "success"
    ExitApp(0)
if mode = "failure"
    ExitApp(7)
if mode = "providerTimeout" {
    FileAppend("504 DEADLINE_EXCEEDED", A_ScriptDir "\candidate_recommendation_error.txt")
    ExitApp(1)
}
if mode = "batches" || mode = "repeated" {
    Loop mode = "batches" ? 3 : 20 {
        tmp := progress ".tmp"
        try FileDelete(tmp)
        FileAppend(mode = "batches" ? A_Index : 1, tmp)
        FileMove(tmp, progress, true)
        Sleep(180)
    }
    ExitApp(0)
}
Sleep(15000)
FileAppend("late result", progress ".late")
ExitApp(0)
