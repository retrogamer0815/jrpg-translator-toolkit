#Requires AutoHotkey v2.0
#SingleInstance Off
Persistent()
kind := A_Args[1]
destination := A_Args[2]
if InStr(kind, "parent") {
    childDestination := A_Args[3]
    Run('"' A_AhkPath '" /ErrorStdOut "' A_ScriptFullPath '" child "' childDestination '"', A_ScriptDir, "Hide")
}
if kind = "hung-parent"
    OnMessage(0x0010, (*) => 0)
FileAppend(DllCall("GetCurrentProcessId", "uint") "`nREADY", destination, "UTF-8")
if kind = "exit-parent"
    ExitApp()
