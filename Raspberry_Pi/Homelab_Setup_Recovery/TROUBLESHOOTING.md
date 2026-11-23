# Troubleshooting Guide

Quick diagnostic guide for common issues with the Raspberry Pi setup scripts and root switcher. USB boot sections apply to Raspberry Pi 4 and 5.

## 📖 Reading This Guide

After running the setup, this file should be available on the boot partition and in /root.

Recommended (pretty formatting):
```bash
glow /root/TROUBLESHOOTING.md
```

Plain text:
```bash
less /root/TROUBLESHOOTING.md
cat /root/TROUBLESHOOTING.md
```

Main documentation:
```bash
glow /root/README.md
```

## 🔍 Quick Health Check

Run these commands to get a fast overview:

```bash
# Boot + storage
findmnt / /boot/firmware
lsblk -o NAME,SIZE,TYPE,FSTYPE,MOUNTPOINT,PARTUUID

# Network
ping -c 1 1.1.1.1
ip addr
ip route
```

---

## ❌ Script Won’t Run

### Symptoms
- Permission denied
- Command not found
- No such file or directory

### Fix
```bash
chmod +x homelab-setup.sh switch-root.sh

# Root switcher (interactive menu)
sudo ./switch-root.sh
```

### Check shebangs
```bash
head -1 homelab-setup.sh
head -1 switch-root.sh
```

---

## ❌ Can’t Copy Files to /boot/firmware

### Symptoms
- “failed to preserve ownership” when using mv or cp
- “Operation not permitted” when copying to /boot/firmware

### Cause
/boot/firmware is FAT32, which doesn’t support Unix ownership or permissions.

### Fix
```bash
# Use cp (not mv)
sudo cp homelab-setup.sh switch-root.sh README.md TROUBLESHOOTING.md /boot/firmware/

# If cp complains about ownership, ignore the warning
ls -l /boot/firmware/*.sh /boot/firmware/*.md
```

Best practice:
- Keep working copies in /root or /home/admin
- Only copy to /boot/firmware for recovery

---

## ❌ USB Storage Device Not Detected (Pi 4, switch-root.sh)

### Symptoms
- “No USB storage devices detected”
- Device not shown in selection menu

### Diagnostic
```bash
lsblk -o NAME,MODEL,TRAN,TYPE,SIZE
lsusb
dmesg | grep -i usb
dmesg | grep -i sd
```

### Common fixes
- Use a powered USB hub
- Try a different USB cable or port
- Restart udev: sudo systemctl restart udev

---

## ❌ Partitioning or Device Node Issues (Pi 4/5)

### Symptoms
- “Device is busy”
- “Partition device nodes not available”
- “Failed to retrieve PARTUUIDs”

### Diagnostic
```bash
mount | grep /dev/sd
sudo lsof | grep /dev/sd
sudo fuser -m /dev/sdX
swapon --show
ls -la /dev/sdX*
sudo partprobe /dev/sdX
sudo udevadm settle
```

### Fix
```bash
sudo umount /dev/sdX*
sudo fuser -km /dev/sdX
sudo swapoff /dev/sdX1
sudo partprobe /dev/sdX
sudo udevadm settle
```

Notes:
- The script already retries partition detection and PARTUUID reads.
- If it still fails, the device or enclosure may be unstable.

---

## ❌ PARTUUID Conflicts (Pi 4/5)

### Symptoms
- “PARTUUID conflict detected”
- Boot loops after cloning or flashing

### Diagnostic
```bash
sudo blkid -s PARTUUID
```

### Fix
```bash
# Regenerate GUIDs on the target disk
sudo sgdisk --randomize-guids /dev/sdX
sudo partprobe /dev/sdX
sudo udevadm settle
```

The scripts automatically check and fail safe when a conflict is detected.

---

## ❌ USB Boot Issues (Pi 4 / Power)

### Symptoms
- USB drive not detected during boot
- “No bootable device”
- Intermittent USB boot failures

### What the script does
- Adds USB power settings to config.txt on the current boot device
- Scans and configures ALL USB boot partitions (option 3 or after options 4–6)

### Verify power settings
```bash
grep -E "(usb_max_current|dtoverlay=dwc2)" /boot/firmware/config.txt
```

Expected:
- dtoverlay=dwc2,dr_mode=host
- usb_max_current_enable=1

### EEPROM boot order (manual only)
Only needed if you want full USB boot without SD fallback.
Use raspi-config or rpi-eeprom-config. Option 8 in switch-root.sh also manages this on Pi 4.

---

## ❌ System Won’t Boot After Migration (Pi 4/5)

### Emergency recovery
1) Boot from SD only (remove USB drive)
2) Restore cmdline.txt backup

```bash
# switch-root.sh backup file name
sudo cp /boot/firmware/cmdline.txt.bak /boot/firmware/cmdline.txt
sudo reboot
```

### Diagnostic (from SD boot)
```bash
cat /boot/firmware/cmdline.txt
cat /boot/firmware/config.txt | grep -E "(usb_max_current|dtoverlay=dwc2)"
```

If needed, run switch-root.sh option 3 to reconfigure USB boot partitions.

---

## ❌ switch-root.sh Download/Flash Issues (Option 7, Pi 4/5)

### Symptoms
- Download stalls
- SHA256 mismatch
- xz integrity check fails
- dd write errors

### Controls
During download, press Ctrl+C for the interactive menu:
- Resume
- Retry download
- View log
- Abort

### Fixes
```bash
# Check network
ping -c 3 1.1.1.1

# Check storage space for image download
df -h /tmp

# If checksum fails, retry download
```

The script validates:
- SHA256 checksum (if available)
- XZ integrity
- Partition detection after write

---

## ❌ SSH or WiFi Not Working After Setup

### Diagnostic
```bash
sudo ./switch-root.sh   # then choose option 2
ip addr
systemctl status ssh
```

### Fix
- Re-run option 2 to review SSH and WiFi state and update first-boot seed files.
- If the target already completed first boot, seed-file edits will not reconfigure it. Connect with a console (keyboard/HDMI) and fix the settings there.

---

## 🐌 Slow Performance

### Diagnostic
```bash
findmnt /
free -h
uptime
```

### Fix
- Migrate to USB storage (switch-root.sh option 6, Pi 4/5 only)
- Use an SSD or faster USB adapter

---

## ❌ Camera Not Detected (Pi Zero 2 W, IMX519)

See [camera/README.md](camera/README.md) for the full camera setup. The camera checks are in `camera-setup.sh --verify`.

### Symptoms
- `rpicam-still --list-cameras` does not list imx519
- Test capture fails

### Fix
- Reseat the ribbon. The Zero 2 W needs a 22-to-15-pin adapter. Make sure the blue tab faces the correct way for your adapter.
- Confirm `/boot/firmware/config.txt` has `camera_auto_detect=0` and `dtoverlay=imx519` under `[all]`.
- Reboot after changing `config.txt`.

---

## ❌ RTSP Stream Not Running (camera-stream service)

### Diagnostic
```bash
systemctl status camera-stream
journalctl -u camera-stream -e
```

### Fix
- Run `sudo bash camera-setup.sh --verify` to see which step fails.
- If the stream is up but Frigate cannot read it, check the password in the homelab `frigate/.env` matches `sudo cat /etc/mediamtx/frigate-password` on the Pi.

---

## 🆘 Emergency Recovery Checklist

### If boot fails (Pi 4/5)
1) Boot from SD only
2) Restore /boot/firmware/cmdline.txt.bak
3) Re-run switch-root.sh option 3 or 4

### If network is broken
```bash
ip addr
ip route
ping -c 1 1.1.1.1
```

---

## 📞 Getting Help (Diagnostics Bundle)

Run this to create a diagnostic file you can share:

```bash
{
  echo "=== System Info ==="
  cat /proc/device-tree/model
  cat /etc/os-release | grep VERSION

  echo -e "\n=== Script Versions ==="
  head -20 /root/homelab-setup.sh
  head -20 /boot/firmware/switch-root.sh

  echo -e "\n=== Boot Config ==="
  findmnt / /boot/firmware
  cat /boot/firmware/cmdline.txt

  echo -e "\n=== Disk ==="
  lsblk -o NAME,SIZE,TYPE,FSTYPE,MOUNTPOINT,PARTUUID

  echo -e "\n=== Network ==="
  ip addr
  ip route

  echo -e "\n=== Recent Errors ==="
  journalctl -p err -n 20 --no-pager
} > diagnostic.txt

less diagnostic.txt
```

Include:
1) Pi model and OS version
2) Script version
3) What you were doing when it failed
4) Exact error output
5) diagnostic.txt contents