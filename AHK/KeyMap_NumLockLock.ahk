#warn
Persistent
#SingleInstance Force
OnExit Closing

/*
Loop
{
    If (WinActive("ahk_exe acad.exe") or WinActive("ahk_exe AcCefSubprocess.exe"))
    {
        SetNumLockState 0
    }
    Else
    {
        SetNumLockState 1
    }

    Sleep(10000)
}
*/

#HotIf GetKeyState("NumLock", "T") = 0
NumpadIns::Numpad0
NumpadEnd::Numpad1
NumpadDown::Numpad2
NumpadPgDn::Numpad3
NumpadLeft::Numpad4
NumpadClear::Numpad5
NumpadRight::Numpad6
NumpadHome::Numpad7
NumpadUp::Numpad8
NumpadPgUp::Numpad9
NumpadDel::NumpadDot
#HotIf


Closing(Exit, *){
	SetNumLockState 1
}