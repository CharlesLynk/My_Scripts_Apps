## Guide: Transferring Executables via Text Copy-Paste (Base64 Method)
This guide explains how to convert a compiled binary file (like a Windows .exe) into a safe text format (Base64), copy it as text, and reconstruct it on another computer. [1, 2] 
This bypasses environment restrictions where file transfers are blocked but text copy-pasting is allowed. [3] 
------------------------------
## 🛠️ Step 1: Encode Binary to Text (Source Computer)
Run this process on the computer that currently holds the compiled file.

   1. Open PowerShell.
   2. Execute the following command (replace the paths with your actual file locations):

[Convert]::ToBase64String([IO.File]::ReadAllBytes("C:\path\to\your\source_program.exe")) | Out-File "C:\path\to\encoded_text.txt"


   1. Open encoded_text.txt.
   2. Select all text (Ctrl + A) and copy it (Ctrl + C).

------------------------------
## 📥 Step 2: Decode Text to Binary (Target Computer)
Run this process on the destination computer where you want to run the program.

   1. Open a text editor (like Notepad).
   2. Paste the copied text and save the file as incoming_text.txt.
   3. Open PowerShell.
   4. Execute the following command to rebuild your executable:

[IO.File]::WriteAllBytes("C:\path\to\restored_program.exe", [Convert]::FromBase64String((Get-Content "C:\path\to\incoming_text.txt" -Raw).Trim()))

------------------------------
## 💡 Technical Considerations & Best Practices## Why This Is Necessary

* Opening an .exe directly in Notepad corrupts the file.
* Text editors alter invisible or unmappable binary characters.
* Base64 maps binary data into 64 safe ASCII characters (A-Z, a-z, 0-9, +, /). [4, 5, 6] 

## Data Overhead

* Base64 encoding increases total file/text size by roughly 33%.
* For small C# applications, this size increase is negligible. [7] 

## Troubleshooting Clipboard Limits

* Large files can freeze a remote clipboard interface.
* Break the text into smaller chunks if the clipboard hangs.
* Alternative Option: If PowerShell commands are blocked, use Command Prompt (cmd) and the built-in Windows utility certutil:
* Encode: certutil -encode source.exe text.txt
   * Decode: certutil -decode text.txt restored.exe

