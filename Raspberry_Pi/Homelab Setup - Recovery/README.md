# Raspberry Pi Homelab Setup Scripts

Automated setup scripts for Raspberry Pi OS (Bookworm/Trixie) that help you migrate from SD to USB storage and deploy a Docker-based homelab stack with secure access.

## 🚀 Features

- **USB Migration and Boot Management** (`switch-root.sh`)
  - Current-device-only, hybrid, and USB-only configurations
  - Full USB flash workflow (download + SHA256 verification + write)
  - USB boot partition auto-configuration (power settings + scripts)
  - Safe cmdline backup and rollback
  - Pre-checks for SSH/WiFi before boot changes
- **Docker Homelab Stack** (`homelab-setup.sh`)
  - Pi-hole, Home Assistant, Homeway, Tailscale
  - Optional nftables firewall with LAN subnet auto-detection
  - Robust retry logic for downloads and image pulls
- **Backup and Restore**
  - Single tar.gz backup file per run
  - Integrity verification and safe restore flow
  - Keeps the newest 3 backups

## 📋 Prerequisites

### Hardware
- Raspberry Pi 4 or 5 (Pi 3 may work but untested)
- USB storage device: SSD/USB flash/NVMe (powered hub recommended on Pi 3)
- SD card with Raspberry Pi OS Lite (Bookworm/Trixie)
- Ethernet recommended

### Network & Credentials
- Static IP for Pi-hole (recommended): `192.168.1.199`
- Tailscale auth key: https://login.tailscale.com/admin/settings/keys
- Home Assistant long-lived access token (used by Homeway) generated from HA UI

## 🎯 Quick Start

### First-Time Remote Setup (SSH from Windows PowerShell)
Use this sequence to place the scripts in a location `switch-root.sh` already searches. This puts them in `/boot/firmware`, which is a primary search location and survives root migration.

```powershell
# From Windows PowerShell (replace USER and PI_IP)
$PiUser   = "admin"
$PiHost   = "192.168.1.25"
$SourceDir = "D:\Utilities\Transfer"
$TempDir   = "/home/$PiUser/homelab-scripts"

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

### Fast Path: Fresh OS on USB (Recommended)
1) Flash Raspberry Pi OS Lite to USB using Pi Imager **or** `switch-root.sh` option 7.
2) Boot from USB.
3) Run:
   ```bash
   sudo ./homelab-setup.sh install-docker
   sudo ./homelab-setup.sh configure
   sudo ./homelab-setup.sh backup
   ```

### Migration Path: Keep Existing System
1) Run `switch-root.sh` option 6 (USB-only) and choose whether to copy your filesystem.
2) Reboot to USB.
3) Run:
   ```bash
   sudo ./homelab-setup.sh install-docker
   sudo ./homelab-setup.sh configure
   sudo ./homelab-setup.sh backup
   ```

## 📦 Script Inventory

### `switch-root.sh` – Boot + Storage Manager

Run:
```bash
sudo ./switch-root.sh
```

Menu options:
1) Show current setting
2) Verify initial configuration (SSH and WiFi)
3) Configure USB boot partitions (power settings + scripts)
4) Switch to current-device-only (boot+root on current device)
5) Switch to Hybrid (boot on current device, root on another)
6) Switch to USB-only (full USB storage setup with boot+root)
7) Flash fresh Raspberry Pi OS to USB (wget + sha256sum + dd)

Key behaviors:
- Performs SSH/WiFi checks before options 4–7.
- Creates a backup of `/boot/firmware/cmdline.txt` before changes.
- Options 4–6 auto-configure all USB boot partitions; option 7 configures only the flashed device.
- Option 7 can optionally copy your user, SSH keys, and network config to the new image.

### `homelab-setup.sh` – Docker Stack Installer

Run:
```bash
sudo ./homelab-setup.sh <mode>
```

Modes:
- `install-docker`: Install Docker, deploy Pi-hole + Home Assistant, create Homeway compose + .env, install Tailscale
- `configure`: Configure Tailscale, optional firewall (nftables), collect HA token for Homeway, optionally lock root
- `backup`: Create a single tar.gz backup (keeps latest 3)
- `restore`: Restore from a selected backup

## 🧭 Detailed Usage

### `switch-root.sh` option 6 – USB-only setup
- Partitions USB (512MB FAT32 boot + ext4 root)
- Clones /boot/firmware to USB boot
- Updates USB cmdline.txt root=PARTUUID
- Optionally copies root filesystem (optimized rsync + excludes)
- Writes fstab on USB root
- Configures power settings and scripts on USB boot

### `switch-root.sh` option 7 – Flash fresh OS to USB
- Downloads Raspberry Pi OS image and SHA256 checksum
- Validates XZ integrity and verifies checksum when available (prompts if checksum download fails)
- Writes image via `dd` and reconfigures USB boot
- Optionally copies user/SSH/network configs from the current system

### `homelab-setup.sh install-docker`
What it does:
- Installs prerequisites and Docker + Compose plugin
- Detects and resolves port conflicts
- Pulls Pi-hole and Home Assistant images with retry
- Starts Pi-hole and Home Assistant (host networking)
- Creates Homeway compose file and `.env` placeholder
- Installs Tailscale
- Installs `glow` when available for terminal Markdown viewing
- Adds Hagezi blocklist to Pi-hole

Service URLs:
- Pi-hole: http://<pi-ip>/admin (password from script)
- Home Assistant: http://<pi-ip>:8123

### `homelab-setup.sh configure`
What it does:
- Connects Tailscale (`--ssh --accept-dns=false`)
- Optional firewall (nftables) with LAN subnet auto-detection
- Prompts for HA long-lived token and stores it in `/homelab/compose/.env` as `HOMEWAY_TOKEN`
- Starts Homeway
- Optional root account password lock (sudo still works)

### `homelab-setup.sh backup`
- Creates: `/homelab/backups/homelab-backup-YYYY-MM-DD_HHMMSS.tar.gz`
- Includes: Pi-hole config, Home Assistant config, compose files (including `.env`), nftables config (if present), Tailscale state
- Verifies archive integrity
- Keeps newest 3 backups

### `homelab-setup.sh restore`
- Selects a backup from the list
- Stops containers + Tailscale
- Restores `/homelab` and system configs (`/etc/nftables.conf`, `/var/lib/tailscale/tailscaled.state` when present)
- Restarts containers and Tailscale

## 📁 Directory Layout

```
/homelab/
├── compose/
│   ├── .env
│   ├── pihole-compose.yml
│   ├── homeassistant-compose.yml
│   └── homeway-compose.yml
├── pihole/
│   └── etc-pihole/
├── homeassistant/
├── homeway/
└── backups/
```

## 🔧 Common Tasks

Check status:
```bash
docker ps
tailscale status
```

Update containers:
```bash
cd /homelab/compose
docker compose -f pihole-compose.yml pull && docker compose -f pihole-compose.yml up -d
docker compose -f homeassistant-compose.yml pull && docker compose -f homeassistant-compose.yml up -d
docker compose -f homeway-compose.yml pull && docker compose -f homeway-compose.yml up -d
```

Change Pi-hole password:
```bash
docker exec -it pihole pihole -a -p
```

## 🔒 Security Notes

- Change Pi-hole password immediately.
- Use Tailscale ACLs for remote access control.
- Consider enabling the optional firewall during `configure`.
- `.env` is permissioned to 600 and contains the Homeway token.

## 🛠️ Troubleshooting

Quick diagnostics:
```bash
docker ps -a
findmnt / /boot/firmware
df -h
tailscale status
docker logs pihole
docker logs homeassistant
```

Common issues:
- Pi-hole won’t start: disable `systemd-resolved` and restart container
- Boot fails after migration: restore `/boot/firmware/cmdline.txt.bak`
- Homeway won’t start: verify `/homelab/compose/.env` has `HOMEWAY_TOKEN=`

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for more.

## 🔄 Customization

Edit constants in `homelab-setup.sh` before running if needed:
- `PIHOLE_PASSWORD`
- `PIHOLE_DNS_UPSTREAM`
- `PIHOLE_STATIC_IP`
- `HAGEZI_BLOCKLIST`

## 📚 Resources

- Script help: `./homelab-setup.sh help`
- Raspberry Pi docs: https://www.raspberrypi.com/documentation/
- Pi-hole docs: https://docs.pi-hole.net/
- Home Assistant docs: https://www.home-assistant.io/docs/
- Tailscale docs: https://tailscale.com/kb/

---

**Version:** 1.0.0 | **License:** MIT
