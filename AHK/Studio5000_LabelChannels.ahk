#Requires AutoHotkey v2.0

^[::
StartNumber := InputBox("Enter Start Number", "Please enter the start number:")
EndNumber := InputBox("Enter End Number", "Please enter the end number:")
return

^]::
Loop EndNumber - StartNumber + 1
{
    X := StartNumber + A_Index - 1
    Clipboard := "Channel " X
    Send("^d")
    Sleep 250
    Send("^v")
    Sleep 250
    if (A_Index < EndNumber - StartNumber + 1)
    {
        KeyWait("LButton", "D")
        Send("{Left}")
        Sleep 250
        Send("{Down}")
        Sleep 250
    }
}
return
	
