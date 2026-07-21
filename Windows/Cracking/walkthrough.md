
## Standard Windows RE (If not bitlocker encrypted)
---
    bcdedit /set {default} safeboot minimal
    bcdedit /set {default} safebootalternateshell yes
    bcdedit /set {default} nointegritychecks on
    bcdedit /set {default} disableelamdrivers yes

    copy C:\Windows\System32\cmd.exe C:\Windows\System32\utilman.exe /y

    wpeutil reboot
    OR
    shutdown /r /t 0

---

## If bcdedit cannot find the store:
    
    diskpart
    list volume
    exit

    diskpart
    select volume <NUMBER>
    assign letter=S
    exit

    bcdedit /store S:\EFI\Microsoft\Boot\BCD /enum
    bcdedit /store S:\EFI\Microsoft\Boot\BCD /set {default} safeboot minimal
    bcdedit /store S:\EFI\Microsoft\Boot\BCD /set {default} safebootalternateshell yes
    bcdedit /store S:\EFI\Microsoft\Boot\BCD /set {default} nointegritychecks on
    bcdedit /store S:\EFI\Microsoft\Boot\BCD /set {default} disableelamdrivers yes

    copy C:\Windows\System32\cmd.exe C:\Windows\System32\utilman.exe /y

    wpeutil reboot
    OR
    shutdown /r /t 0

---

## After booting into Windows:

    net user Administrator /active:yes
    net user Administrator password
    OR
    net user NewAdmin Password123 /add
    net localgroup Administrators NewAdmin /add

    HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\PasswordLess\Device
        DevicePasswordLessBuildVersion 0
    HKEY_LOCAL_MACHINE\SAM\SAM\Domains\Account\Users\Names\
        Find offending user, note the hex ID
    HKEY_LOCAL_MACHINE\SAM\SAM\Domains\Account\Users\
        Delete all "Internet..." keys

    net user OffendingUser Password
    
---

## When ready:

    bcdedit /deletevalue {default} safeboot
    bcdedit /deletevalue {default} safebootalternateshell
    bcdedit /set {default} nointegritychecks off