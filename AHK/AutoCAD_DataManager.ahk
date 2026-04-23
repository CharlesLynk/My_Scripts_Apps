#SingleInstance Force
Persistent
SendMode ("Input")  ; Recommended for new scripts due to its superior speed and reliability.

; Define variables to store the find and replace strings
findString := "100"
replaceString := "999"
programArmed := 0
oldLines := ""
newLines := ""
loopBreaker := 0
delayTime := 250


#HotIf WinActive("ahk_exe acad.exe") or WinActive("ahk_exe AcCefSubprocess.exe")


^Numpad7::
{
	global programArmed := 0
	global oldLines := ""
	global newLines := ""
	
    ; Prompt the user for the find string
    Global findString := InputBox("Enter the string to find:", "Find String").Value

    ; Prompt the user for the replace string
    Global replaceString := InputBox("Enter the string to replace with:", "Replace String").Value

	return
}

^!Numpad7::
{
	Global delayTime := InputBox("Enter the required delay time in ms, recommend 2 seconds for line groups and .25 for others:", "Delay Time").Value

	return
}

^!Numpad3::
{
	If loopBreaker = 0
	{
		global loopBreaker := 1
		return
	}
}
return


^Numpad3::
{
	if (programArmed != 1){
		displayString := ""
		clipboardContent := ""
		global oldLines := ""
		global newLines := ""
		
		; Retrieve the contents of the clipboard
		clipboardContent := A_Clipboard

		; Split the clipboard contents into an array of lines. Keep the original values for comparison
		global oldLines := StrSplit(clipboardContent, "`n", "`r`n `t")
		global newLines := StrSplit(clipboardContent, "`n", "`r`n `t")
		
		; Loop through each line and replace the find string with the replace string
		for index, line in newLines
		{
			; If the findString is an asterisk, replace the rightmost characters with replaceString.
			if (findString = "*")
			{
				newLines[index] := Substr(newLines[index], 1, Max(0, StrLen(newLines[index]) - StrLen(replaceString))) . replaceString
			}
			else
			{
				newLines[index] := StrReplace(newLines[index], findString, replaceString)
			}
			
			; Generate the preview string for user approval before arming the paste function
			if (index < newLines.length)
			{
				displayString .= oldLines[index] " -> " newLines[index] "," "`r`n"
			}
			else
			{
				displayString .= oldLines[index] " -> " newLines[index]
			}
		}

		; Show the array of strings to the user for approval
		if (MsgBox(displayString, "Arm the Program?", 257) = "OK")
		{
			global programArmed := 1
		}
		else
		{
			global programArmed := 0
		}
	}
	else
	{
		global loopBreaker := 0
		; Loop through each string in the array and send it followed by an "Enter" keypress
			for index, line in newLines
			{
				if (loopBreaker = 1)
				{
					global loopBreaker := 0
					break					
				}
				; Send the string followed by an "Enter" keypress
				;SendInput(newLines[index])
				A_Clipboard := ""
				;ClipWait (2)
				A_Clipboard := newLines[index]
				if !ClipWait(2)
				{
					MsgBox "The attempt to copy text onto the clipboard failed."
					break
				}
				Send ("^v")
				Sleep(delayTime)
				Send ("{Enter}")
				
			}
		global programArmed := 0
	}
	return
}

#HotIf 