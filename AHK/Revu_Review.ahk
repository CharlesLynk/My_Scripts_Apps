#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force

; ===== Speed / Reliability Settings =====
#InstallKeybdHook
#UseHook
SetBatchLines, -1
SetKeyDelay, -1, -1             ; no key delay, no press duration
SetWinDelay, -1
SetControlDelay, -1
SetMouseDelay, -1
SendMode, Input                  ; fastest reliable send method

; ===== Helper: run your sequence and then Enter → same number → Esc =====
; number: the digit you already send after "S"
DoRevu(number) {
    ; Open context menu and give Revu a tiny moment to render it
    SendInput, {RButton}
    Sleep, 30                    ; 20–40 ms is typically enough for Revu

    ; Send S and the chosen number
    SendInput, s%number%

    ; Immediately confirm and finalize per your requirement:
    ; Enter → same number again → Escape
    ; Most apps can absorb this chain without sleeps; add a tiny Sleep if needed.
    SendInput, {Enter}%number%{Esc}
}

; ===== Hotkeys / Calls =====

^8::
if WinActive("ahk_exe REVU.EXE")
{
    DoRevu(4)
}
return

^6::
if WinActive("ahk_exe REVU.EXE")
{
    DoRevu(3)
}
return

^0::
if WinActive("ahk_exe REVU.EXE")
{
    DoRevu(2)
}
return

^4::
if WinActive("ahk_exe REVU.EXE")
{
    DoRevu(1)
}
return

/*

twoSetsInMenu := false


F7::
if WinActive("ahk_exe REVU.EXE")
    twoSetsInMenu := !twoSetsInMenu
	if WinActive("ahk_exe REVU.EXE") && twoSetsInMenu
	{
		ToolTip, Default Is In Context Menu
		SetTimer, RemoveToolTip, -1000
		return
	}
	else if WinActive("ahk_exe REVU.EXE") && !twoSetsInMenu
	{
		ToolTip, No Default In Context Menu
		SetTimer, RemoveToolTip, -1000
		return
	}
return



RemoveToolTip:
	ToolTip
return


#If !twoSetsInMenu
^8::
if WinActive("ahk_exe REVU.EXE")
    Send {RButton}
	Sleep 20
	Send S
	Sleep 20
	;Send R
	;Sleep 20
	;Send {Enter}
	;Sleep 20
	Send 4
return

^6::
if WinActive("ahk_exe REVU.EXE")
    Send {RButton}
	Sleep 20
	Send S
	Sleep 20
	;Send R
	;Sleep 20
	;Send {Enter}
	;Sleep 20
	Send 3
return

^0::
if WinActive("ahk_exe REVU.EXE")
    Send {RButton}
	Sleep 20
	Send S
	Sleep 20
	;Send R
	;Sleep 20
	;Send {Enter}
	;Sleep 20
	Send 2
return

^4::
if WinActive("ahk_exe REVU.EXE")
    Send {RButton}
	Sleep 20
	Send S
	Sleep 20
	;Send R
	;Sleep 20
	;Send {Enter}
	;Sleep 20
	Send 1
return
#If

#If twoSetsInMenu
^8::
if WinActive("ahk_exe REVU.EXE")
    Send {RButton}
	Sleep 20
	Send S
	Sleep 20
	Send {Enter}
	Sleep 20
	;Send R
	;Sleep 20
	;Send {Enter}
	;Sleep 20
	Send 4
return

^6::
if WinActive("ahk_exe REVU.EXE")
    Send {RButton}
	Sleep 20
	Send S
	Sleep 20
	Send {Enter}
	Sleep 20
	;Send R
	;Sleep 20
	;Send {Enter}
	;Sleep 20
	Send 3
return

^0::
if WinActive("ahk_exe REVU.EXE")
    Send {RButton}
	Sleep 20
	Send S
	Sleep 20
	Send {Enter}
	Sleep 20
	;Send R
	;Sleep 20
	;Send {Enter}
	;Sleep 20
	Send 2
return

^4::
if WinActive("ahk_exe REVU.EXE")
    Send {RButton}
	Sleep 20
	Send S
	Sleep 20
	Send {Enter}
	Sleep 20
	;Send R
	;Sleep 20
	;Send {Enter}
	;Sleep 20
	Send 1
return
#If

*/