# Troubleshooting Guide

Quick diagnostic guide for common issues with the homelab setup script and root switcher.

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
# Services
docker ps -a

# Boot + storage
findmnt / /boot/firmware
lsblk -o NAME,SIZE,TYPE,FSTYPE,MOUNTPOINT,PARTUUID

# Network
ping -c 1 1.1.1.1
tailscale status

# Firewall (optional)
sudo systemctl is-active nftables
sudo nft list ruleset

# Logs
docker logs --tail 50 pihole
docker logs --tail 50 homeassistant
docker logs --tail 50 homeway
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

# Homelab setup (requires mode)
sudo ./homelab-setup.sh install-docker
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

## ❌ USB Storage Device Not Detected (switch-root.sh)

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
- Use a powered USB hub (Pi 3/4 especially)
- Try a different USB cable or port
- Restart udev: sudo systemctl restart udev

---

## ❌ Partitioning or Device Node Issues

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

## ❌ PARTUUID Conflicts

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

## ❌ USB Boot Issues (Pi 5 / Power)

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
Use raspi-config or rpi-eeprom-config.

---

## ❌ System Won’t Boot After Migration

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

## ❌ switch-root.sh Download/Flash Issues (Option 7)

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

## ❌ Docker Install / Docker Daemon Issues

### Symptoms
- Docker command missing after install
- Docker daemon not running

### Diagnostic
```bash
docker --version
systemctl status docker
journalctl -u docker -n 50
```

### Fix
```bash
sudo systemctl start docker
sudo systemctl enable docker
```

If installation failed, re-run:
```bash
sudo ./homelab-setup.sh install-docker
```

---

## ❌ Pi-hole Won’t Start

### Symptoms
- Container crashes
- Port 53 conflict
- Port 80 conflict

### Diagnostic
```bash
sudo ss -tulpn | grep :53
sudo ss -tulpn | grep :80
systemctl status systemd-resolved
docker logs pihole
```

### Fix (common)
```bash
# Disable systemd-resolved
sudo systemctl stop systemd-resolved
sudo systemctl disable systemd-resolved

# Replace resolv.conf
sudo rm /etc/resolv.conf
echo "nameserver 1.1.1.1" | sudo tee /etc/resolv.conf
echo "nameserver 1.0.0.1" | sudo tee -a /etc/resolv.conf

# Restart Pi-hole
docker restart pihole
```

---

## ❌ Home Assistant Won’t Start

### Symptoms
- Container keeps restarting
- “Unhealthy” status
- Very slow first start

### Diagnostic
```bash
docker ps -a | grep homeassistant
docker logs --tail 100 homeassistant
ls -la /homelab/homeassistant
df -h /homelab
```

### Fix
```bash
# Permissions
sudo chown -R root:root /homelab/homeassistant

# Fresh start (destructive)
sudo rm -rf /homelab/homeassistant/*
```

Note: First start can take 10–15 minutes on slower storage.

---

## ❌ Homeway Won’t Start / HOMEWAY_TOKEN Issues

### Symptoms
- HOMEWAY_TOKEN not set
- Container restarts

### Diagnostic
```bash
docker logs homeway
ls -la /homelab/compose/.env
grep HOMEWAY_TOKEN /homelab/compose/.env
```

### Fix
```bash
# Run configure mode to populate .env
sudo ./homelab-setup.sh configure

# Or set manually
cd /homelab/compose
sed -i '/^HOMEWAY_TOKEN=/d' .env
echo "HOMEWAY_TOKEN=your-token" >> .env
chmod 600 .env

docker compose -f homeway-compose.yml up -d
```

Notes:
- install-docker creates a placeholder .env with empty HOMEWAY_TOKEN
- configure mode fills the token
- .env persists across reboots and restores

---

## ❌ Tailscale Won’t Connect

### Diagnostic
```bash
tailscale status
systemctl status tailscaled
journalctl -u tailscaled -n 50
```

### Fix
```bash
sudo systemctl restart tailscaled
sudo tailscale up --ssh --accept-dns=false
```

If auth key expired, generate a new one and re-run configure mode.

---

## ❌ Locked Out After Firewall Configuration

### Cause
The firewall (nftables) allows only LAN subnet + Tailscale.
If the subnet was wrong, SSH is blocked.

### Recovery
```bash
# Local console
sudo nft flush ruleset
sudo systemctl stop nftables

# Or remove config
sudo rm /etc/nftables.conf
```

Then re-run:
```bash
sudo ./homelab-setup.sh configure
```

---

## ❌ Backup Fails

### Diagnostic
```bash
df -h /homelab
ls -lh /homelab/backups
journalctl -p err -n 20
```

### Fix
```bash
# Free space
docker system prune -a --volumes

# Remove old backups
sudo rm -rf /homelab/backups/homelab-backup-OLD*.tar.gz
```

---

## ❌ Restore Fails

### Diagnostic
```bash
ls -lh /homelab/backups
cd /homelab/backups
docker logs pihole
docker logs homeassistant
```

### Fix
- Try an older backup
- Verify Docker is running
- Check disk space

---

## 🐌 Slow Performance

### Diagnostic
```bash
findmnt /
free -h
uptime
```

### Fix
- Migrate to USB storage (switch-root.sh option 6)
- Use an SSD or faster USB adapter
- Reduce container load or disable Homeway

---

## 🆘 Emergency Recovery Checklist

### If boot fails
1) Boot from SD only
2) Restore /boot/firmware/cmdline.txt.bak
3) Re-run switch-root.sh option 3 or 4

### If Docker is broken
```bash
sudo systemctl restart docker
journalctl -u docker -n 100
```

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

  echo -e "\n=== Docker Containers ==="
  docker ps -a

  echo -e "\n=== Tailscale ==="
  tailscale status

  echo -e "\n=== Firewall ==="
  sudo systemctl is-active nftables
  sudo nft list ruleset 2>/dev/null | head -40

  echo -e "\n=== Homeway Token ==="
  if [[ -f /homelab/compose/.env ]]; then
    sed 's/HOMEWAY_TOKEN=.*/HOMEWAY_TOKEN=***MASKED***/' /homelab/compose/.env
  else
    echo "Missing /homelab/compose/.env"
  fi

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
