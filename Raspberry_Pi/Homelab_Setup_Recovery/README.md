# Raspberry Pi Headless Setup Scripts

Setup scripts for Raspberry Pi OS (Bookworm/Trixie) for first-time headless configuration: SSH and WiFi pre-checks, first-boot seed files, and helper functions. Pi-hole, Home Assistant, and Tailscale now run on the Windows Hyper-V homelab, so they are not set up by these scripts.

## 🚀 Features

- **Boot and Storage Manager** (`switch-root.sh`)
  - SSH and WiFi pre-checks before boot changes
  - First-boot seed file updates for the next boot target (`ssh`, `network-config`, `user-data`)
  - Safe `cmdline.txt` backup and rollback
  - Fresh OS flash workflow with SHA256 verification (option 7)
  - **USB boot features (options 3–8) are Raspberry Pi 4 and 5 only; the script refuses to run them on other boards**
- **Shared Helpers** (`homelab-setup.sh`)
  - Logging, prompts, disk-space checks, and retry helpers for headless setup
  - Currently has no setup modes. It is kept as a helper library and placeholder entry point.

## 📋 Prerequisites

### Hardware
- Raspberry Pi 4 or 5 for USB boot features (options 3–8)
- Raspberry Pi 4 or 5 for generic first-boot setup
- Raspberry Pi Zero 2W is not covered by the USB boot options (it has no USB-A port; USB boot is untested on it and not supported by these options)
- SD card with Raspberry Pi OS Lite (Bookworm/Trixie)
- Ethernet recommended for first-time setup
- USB storage device (SSD/USB flash/NVMe) if using USB boot (Pi 4 or 5)

### Network & Credentials
- WiFi SSID and password (if not using Ethernet)
- Username and password for the new system

## 🎯 Quick Start

### First-Time Remote Setup (SSH from Windows PowerShell)
Use this sequence to place the scripts in a location `switch-root.sh` already searches. This puts them in `/boot/firmware`, which is a primary search location and survives root migration.

```powershell
# From Windows PowerShell (replace USER and PI_IP)
$PiUser   = "admin"
$PiHost   = "192.168.1.25"
$SourceDir = "D:\Utilities\Transfer"
$TempDir   = "/home/$PiUser/pi-setup-scripts"

# 1) Nuke keygen if it locks you out
ssh-keygen.exe -R $PiHost

# 2) Create a temp directory on the Pi
ssh "$PiUser@$PiHost" "mkdir -p '$TempDir'"

# 3) Copy local files to the temp directory (recursively)
# The '\.' ensures all contents are copied, not the directory name itself
scp -r "$SourceDir\." "$($PiUser)@$($PiHost):$TempDir"

# 4) Move scripts into /boot/firmware and make them executable
ssh "$PiUser@$PiHost" "sudo cp '$TempDir'/* /boot/firmware/ && sudo rm '$TempDir'/* && sudo rmdir '$TempDir' && sudo chmod +x /boot/firmware/switch-root.sh /boot/firmware/homelab-setup.sh"

```

### Fast Path: Fresh OS on USB (Pi 4/5)
1) Flash Raspberry Pi OS Lite to USB using Pi Imager **or** `switch-root.sh` option 7.
2) Boot from USB.
3) Run `sudo ./switch-root.sh` and choose option 2 to verify SSH and WiFi.

### Migration Path: Keep Existing System (Pi 4/5)
1) Run `switch-root.sh` option 6 (Clone OS to USB target) and choose whether to copy your filesystem.
2) Reboot to USB.
3) Run `sudo ./switch-root.sh` and choose option 2 to verify SSH and WiFi.

## 📦 Script Inventory

### `switch-root.sh` – Boot + Storage Manager

Run:
```bash
sudo ./switch-root.sh
```

Menu options:
1) Show current setting
2) Verify initial configuration (SSH and WiFi)
3) Configure USB boot partitions (add power settings to config.txt) — **Pi 4/5 only**
4) Switch to current-device-only (boot+root on current device) — **Pi 4/5 only**
5) Switch to Hybrid (boot on current device, root on another) — **Pi 4/5 only**
6) Clone OS to USB target (boot+root) — **Pi 4/5 only**
7) Flash fresh Raspberry Pi OS to USB (wget + sha256sum + dd) — **Pi 4/5 only**
8) Configure EEPROM boot order (SD/USB priority) — **Pi 4/5 only**

Key behaviors:
- Performs SSH/WiFi checks before options 4–7.
- Option 2 can prompt for target-first-boot SSH enable/disable and WiFi SSID/password, then update target boot seed files (`ssh`, `network-config`, `user-data`) when a next boot target is detected.
- Option 2 warns when the target appears to have already completed first boot (seed-file edits may not reconfigure an already-initialized system).
- Creates a backup of `/boot/firmware/cmdline.txt` before changes.
- Options 3–6 configure USB boot partitions; option 7 configures only the device it flashed.
- Option 7 applies Fresh_Install-formatted `cmdline.txt`, `config.txt`, `network-config`, `user-data`, and `ssh` files after flash.
- Option 7 prompts for plaintext WiFi SSID/password, username/password, and hostname, then dynamically rewrites target PARTUUID/fstab.

### `homelab-setup.sh` – Shared Helpers

Currently has no setup modes. Running it prints a notice and exits. Its helper functions (logging, prompts, disk-space checks, retry wrappers) are available for new setup modes. The file is kept in the boot partition alongside `switch-root.sh` so both are available after a root migration.

### `camera/` – Arducam IMX519 RTSP Camera (Pi Zero 2 W)

Planned setup for a Pi Zero 2 W with an Arducam IMX519. The Pi streams RTSP. Motion detection, recording, and viewing run on the homelab in Frigate. See [camera/README.md](camera/README.md).

- `camera/camera-setup.sh` – installs camera packages, enables the IMX519 overlay, verifies the camera, and runs the RTSP stream (MediaMTX) as a systemd service
- `camera/frigate/` – Docker Compose files and Frigate config for the homelab Linux VM

## 🧭 Detailed Usage

### `switch-root.sh` option 2 – Verify + pre-boot target connectivity update
- Verifies local SSH/WiFi state.
- Checks boot-target readiness when applicable.
- Offers connectivity sync for validated targets.
- Optionally writes first-boot seed connectivity files (`ssh`, `network-config`, `user-data`) on the next boot target.
- Shows a first-boot applicability warning before writing seed files if cloud-init appears already finalized.

### `switch-root.sh` option 6 – Clone OS to USB target (Pi 4/5 only)
- Partitions USB (512MB FAT32 boot + ext4 root)
- Clones /boot/firmware to USB boot
- Updates USB cmdline.txt root=PARTUUID
- Optionally copies root filesystem (optimized rsync + excludes)
- Writes fstab on USB root
- Configures power settings and scripts on USB boot

### `switch-root.sh` option 7 – Flash fresh OS to USB (Pi 4/5 only)
- Downloads Raspberry Pi OS image and SHA256 checksum
- Validates XZ integrity and verifies checksum when available (prompts if checksum download fails)
- Writes image via `dd`, regenerates partition identifiers, and rewrites boot/root mappings
- Applies Fresh_Install-style first-boot files with prompted WiFi/user/hostname values

### `switch-root.sh` option 8 – Configure EEPROM boot order (Pi 4/5 only)
- Reads and validates current EEPROM boot order.
- Lets you choose SD/USB preference safely.
- Re-checks target boot continuity/readiness after change.

## 📁 Script Files on the Boot Partition

```
/boot/firmware/
├── switch-root.sh
├── homelab-setup.sh
├── README.md
├── TROUBLESHOOTING.md
├── cmdline.txt
├── cmdline.txt.bak      (created by switch-root.sh)
└── ...
```

## 🔧 Common Tasks

Check status:
```bash
findmnt / /boot/firmware
df -h
lsblk
```

Check SSH and WiFi state:
```bash
sudo ./switch-root.sh   # then choose option 2
```

## 🔒 Security Notes

- Use a strong password for the new user account.
- Option 7 writes the WiFi password in plaintext to first-boot `network-config` and stores the user password as a hash in `user-data`. Remove or secure those files on the boot partition once the Pi has finished first boot.

## 🛠️ Troubleshooting

Quick diagnostics:
```bash
findmnt / /boot/firmware
df -h
lsblk -f
```

Common issues:
- Boot fails after migration (Pi 4/5): restore `/boot/firmware/cmdline.txt.bak`

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for more.

## 🔄 Customization

Edit constants in `homelab-setup.sh` before running if needed:
- `CONNECTION_RETRY_MAX_ATTEMPTS` / `CONNECTION_RETRY_DELAY`
- `APT_RETRY_ATTEMPTS` / `APT_RETRY_DELAY`
- `MIN_DISK_SPACE_MB`

## 📚 Resources

- Script help: `sudo ./switch-root.sh`
- Raspberry Pi docs: https://www.raspberrypi.com/documentation/

---

**Version:** 1.0.0 | **License:** MIT