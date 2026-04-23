; Initialize the counter
counter := 1

; Define the hotkey for CTRL+J
^+j::
{
    
    
    ; Send CTRL+P
    Send, ^p
    Sleep, 150


    ; Send ENTER
    Send, {Enter}
    Sleep, 50
    
    ; Type the current counter value
    Send, %counter%
    Sleep, 50
    
    ; Send ENTER
    Send, {Enter}

; Increment the counter
    counter++
}
return

; Define the hotkey for CTRL+SHIFT+J
^+k::
{
    ; Prompt the user for a number
    InputBox, customCounter, Enter a number, Please enter the next number to use:
    
    ; Set the counter to the custom value
    counter := customCounter
}
return