#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force

^+WheelUp::
if WinActive("ahk_exe EXCEL.EXE")
    Send ^{PgUp}
return

^+WheelDown::
if WinActive("ahk_exe EXCEL.EXE")
    Send ^{PgDn}
return
