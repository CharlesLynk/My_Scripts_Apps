# StreamDock label size: device test steps

Run all commands from `VSD_MacroPad\Tools` in PowerShell.

The tool uses the current account's `%APPDATA%\HotSpot\StreamDock\profiles` by
default. Pass `-ProfilesRoot` to use another installation or an offline copy
of an unpacked profile tree; the checked-in `.SDProfile` backups are not that
tree. Set `-BackupRoot` and `-GeneratedRoot` to separate writable directories
for an offline check. Defaults are `VSD_MacroPad\Backups\Auto` and
`VSD_MacroPad\Generated`, relative to the script, not the current directory.
Use Windows PowerShell 5.1; rendering uses Windows `System.Drawing`. Device
checks require VSD Craft and a StreamDock.

## Before you start

- Tested offline on temp copies only. Your live profiles have not been changed.
- Writes are refused while VSD Craft is running unless you add `-Force`.
- Every write makes a backup first. Restore is always available (see the last section).

## 1. Check the current state (no changes)

```powershell
.\Set-StreamDockFontSize.ps1 -Mode List
```

## 2. Render previews (no changes to profiles)

```powershell
.\Set-StreamDockFontSize.ps1 -Mode Render -ProfileName 'VSD Guidance'
```

Open `..\Generated\VSD Guidance\p1\sheet.png` and check the text.

## 3. Quick, safe fix first (changes FontSize only)

Pick a fixed size and apply it to one profile:

```powershell
.\Set-StreamDockFontSize.ps1 -Mode Preset -ProfileName 'AutoCAD' -PresetName Fixed-8 -WhatIf
.\Set-StreamDockFontSize.ps1 -Mode Preset -ProfileName 'AutoCAD' -PresetName Fixed-8 -Force
```

Close VSD Craft before the `-Force` run. Without `-Force` the tool refuses to write while VSD Craft is open.

## 4. Image test on the device (start with the smallest profile)

Close VSD Craft, then run:

```powershell
.\Set-StreamDockFontSize.ps1 -Mode Apply -ProfileName 'VSD Guidance' -WhatIf
.\Set-StreamDockFontSize.ps1 -Mode Apply -ProfileName 'VSD Guidance' -Force
```

Apply changes the title to images and hides the device's own text. It affects all pages of the profile.

Then open VSD Craft and check on the device:
- Are the labels sharp and at the right size?
- Do they stay correct after VSD Craft restarts, and after the device sits idle or is reconnected?
- Do the images still show after a restart? (Unknown: VSD Craft may need the profile re-loaded.)

If the images look good, run it on AutoCAD the same way. If not, restore.

## Restore

Each write prints a backup stamp. You can also find it as a folder name under `..\Backups\Auto`.

```powershell
.\Set-StreamDockFontSize.ps1 -Mode Restore -BackupStamp <stamp> -Force
```

Restore reverts the manifest files. New image files are left in place, which is harmless.

## Open questions to answer during testing

- Does FontSize map to pixels or em on the device?
- Does a 72×72 image display 1:1, or is it scaled?
- Does VSD Craft reload profile files from disk on restart?
- Is the key layout right (column 0–5, row 0–2)?
