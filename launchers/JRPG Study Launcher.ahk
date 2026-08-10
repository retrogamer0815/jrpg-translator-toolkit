#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
;@Ahk2Exe-SetVersion 0.9.5.0
;@Ahk2Exe-SetName JRPG Translator Study Launcher
;@Ahk2Exe-SetDescription Opens the JRPG Translator Study Library or Reader
;@Ahk2Exe-SetCopyright Copyright (c) 2025 retrogamer0815
; Compile this source under either of these names:
;   JRPG Translator - Study Library.exe
;   JRPG Translator - Study Reader.exe
; The executable name selects which shared JRPG Translator mode is opened.

launcherMode := InStr(StrLower(A_ScriptName), "reader")
    ? "--study-reader" : "--study-library"
launcherDirectory := A_ScriptDir
translatorPath := launcherDirectory "\JRPG Translator.exe"
if !FileExist(translatorPath) {
    launcherDirectory := A_ScriptDir "\.."
    translatorPath := launcherDirectory "\JRPG Translator.exe"
}
if !FileExist(translatorPath) {
    MsgBox(
        "JRPG Translator.exe could not be found next to this launcher.",
        "JRPG Translator",
        "OK Iconx"
    )
    ExitApp
}
Run('"' translatorPath '" ' launcherMode, launcherDirectory)
