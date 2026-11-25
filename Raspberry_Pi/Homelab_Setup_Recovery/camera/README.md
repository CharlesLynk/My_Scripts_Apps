# Arducam IMX519 Security Camera (Pi Zero 2 W)

Status: planned. The files are written and reviewed statically but have not been run on hardware.

## Design

- **Pi Zero 2 W** captures video with the Arducam IMX519 and publishes an RTSP stream (H.264, 1280x720 @ 15 fps). It does no motion detection or recording.
- **Homelab Linux VM (Hyper-V)** runs Frigate in Docker. Frigate pulls the stream, runs motion and object detection, records events, and serves the web UI.

Keeping detection on the homelab leaves the Zero 2 W free to stream. It also keeps the camera's storage and retention rules in one place.

## Hardware

- Raspberry Pi Zero 2 W with microSD (Raspberry Pi OS Lite, Bookworm or Trixie)
- Arducam IMX519 camera module
- 22-pin to 15-pin camera ribbon adapter. The Zero 2 W has the smaller camera connector.
- Wi-Fi or Ethernet adapter. The Zero 2 W has Wi-Fi but no Ethernet port.

## Pi setup

Copy the `camera` folder to the Pi, then run:

```bash
cd camera
sudo bash camera-setup.sh          # installs packages, enables the IMX519 overlay, offers reboot
sudo bash camera-setup.sh --verify  # after reboot: detects the camera, takes a test still
sudo bash camera-setup.sh --stream  # installs MediaMTX and starts the camera-stream service
```

The script changes `/boot/firmware/config.txt` (backed up first) to add:

```
camera_auto_detect=0
[all]
dtoverlay=imx519
```

`--stream` prints the RTSP URL with the generated password. The password is also stored in `/etc/mediamtx/frigate-password`.

## Homelab setup (Frigate)

Run these on the Hyper-V Linux VM with Docker Engine and the Docker Compose plugin (`docker compose`). Work from `camera/frigate/` so Compose picks up the local `.env`, `config/`, and `mosquitto/` files. The checkout includes the configuration files, but not the credentials or writable media directory.

1. On a fresh checkout, copy `.env.example` to `.env`. On recovery, keep or restore the existing `.env` instead of overwriting it. Set:
   - `FRIGATE_RTSP_PASSWORD`: the password printed by `camera-setup.sh --stream`
   - `FRIGATE_MQTT_PASSWORD`: the password matching the recovered broker file, or a new value for a fresh broker (for example, `openssl rand -hex 24`)
2. Keep or restore an existing `mosquitto/passwd`. **Only if the file is absent on a fresh setup**, create it with the password from step 1. `-c` overwrites an existing file:
   ```bash
   if [ ! -e ./mosquitto/passwd ] && [ ! -L ./mosquitto/passwd ]; then
     docker compose run --rm --entrypoint mosquitto_passwd mosquitto -c /mosquitto/config/passwd frigate
   fi
   ```
   The broker must be able to read `mosquitto/passwd` through the config bind mount. Keep the file private and preserve its ownership and permissions when restoring it.
3. In `config/config.yml`, replace the `PI_ADDRESS` placeholder with the Pi's Tailscale IP if it has not already been set. On the default layout, ensure `./storage` exists and is writable by Frigate (`mkdir -p ./storage` on a fresh setup). For an external drive, see [Optional media location](#optional-media-location).
4. Validate the Compose definition without starting containers or printing resolved credentials:
   ```bash
   docker compose config --quiet
   ```
   This checks the definition, not mount availability, credentials, or Frigate's application config.
5. Start the stack:
   ```bash
   docker compose up -d
   ```
6. Open `http://<vm-address>:5000`. The camera appears as `pi_zero_cam`.

Detection runs on the CPU detector in the current config. Frigate's current docs recommend the camera setup wizard and UI-based detector settings. Check `config.yml` against the version you pin before relying on it.

### Optional media location

Leave `FRIGATE_STORAGE_PATH` unset to keep using `camera/frigate/storage/`. To use an external filesystem, set it in `.env` to the absolute Linux VM path of an existing writable directory. Before starting Frigate, verify the intended filesystem is mounted and the directory is writable by Frigate. Compose's bind mount can otherwise create a directory on the VM's root filesystem and record there instead.

To migrate existing media, stop Frigate, back up its media and `config/` (including the database), then copy the media tree to the new location, preserving ownership and permissions. Change only `FRIGATE_STORAGE_PATH`; `config/` and `mosquitto/` stay in place. Keep the original media until playback and new recordings have been verified. To roll back, stop Frigate before unsetting the override; reconcile recordings made after the switch with the saved database and media as needed.

### Dependency versions

Compose uses the checked-in Frigate `0.18.1` tag. The broker still defaults to the moving `eclipse-mosquitto:2` tag; the checkout does not identify a tested patch or digest. For repeatable recovery, record a verified `eclipse-mosquitto@sha256:…` reference from the working deployment and set `MOSQUITTO_IMAGE` in `.env`. This pins the broker image used by both password-file creation and the stack. The Frigate release tag is versioned, not an immutable digest.

On the Pi, `camera-setup.sh` already defaults `MEDIAMTX_VERSION` to `1.21.4`. It installs `rpicam-apps`, `ffmpeg`, `curl`, and `ca-certificates` from the target Raspberry Pi OS repositories; exact package versions depend on that OS release and its repositories.

## Recording policy

- Continuous recording is off (`record.retain.days: 0`).
- Alert and detection clips are kept for 14 days.
- Tracked objects: person, car, dog, cat.

## Security notes

- The stream uses RTSP with a generated password. The password file is owned by root and readable by the Pi's service user (mode 640). Only localhost can publish. Frigate's read-only user is the only RTSP client the config grants.
- Do not expose port 8554 on the Pi to the internet. Use Tailscale or the LAN only.
- Keep `/etc/mediamtx/frigate-password` and `frigate/.env` out of git. The root `.gitignore` already excludes `.env` files.

## Bandwidth and performance

- A 1280x720 @ 15 fps H.264 stream at 2 Mbps uses about 2 Mbps of upload from the Pi. Check that the Wi-Fi link can sustain it.
- The Zero 2 W is a 2.4 GHz device. Expect Wi-Fi throughput to vary. Lower `STREAM_BITRATE` and `STREAM_FPS` in `camera-setup.sh` if the stream stutters.

## Known gaps (not yet tested)

- `rpicam-vid | ffmpeg` publishing into MediaMTX through `runOnInit` has not been run on a Pi.
- The Frigate config schema was written from the docs at the time of writing. Verify it against the `0.18.1` image configured in `docker-compose.yml`.
- The IMX519 tuning file path (`/usr/share/libcamera/ipa/rpi/vc4/imx519.json`) is assumed. `--verify` warns if it is missing.
