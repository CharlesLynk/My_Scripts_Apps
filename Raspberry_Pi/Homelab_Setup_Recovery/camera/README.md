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

Run these on the Hyper-V Linux VM, in `camera/frigate/`.

1. Copy `.env.example` to `.env` and fill in:
   - `FRIGATE_RTSP_PASSWORD`: the password printed by `camera-setup.sh --stream`
   - `FRIGATE_MQTT_PASSWORD`: a new value (for example, `openssl rand -hex 24`)
2. Create the Mosquitto user. You will be prompted for the password from step 1:
   ```bash
   docker compose run --rm --entrypoint mosquitto_passwd mosquitto -c /mosquitto/config/passwd frigate
   ```
3. In `config/config.yml`, replace `PI_ADDRESS` with the Pi's Tailscale IP.
4. Start the stack:
   ```bash
   docker compose up -d
   ```
5. Open `http://<vm-address>:5000`. The camera appears as `pi_zero_cam`.

Detection runs on the CPU detector in the current config. Frigate's current docs recommend the camera setup wizard and UI-based detector settings. Check `config.yml` against the version you pin before relying on it.

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
- The Frigate config schema was written from the docs at the time of writing. Verify against the Frigate version you pin. `stable` is used in `docker-compose.yml`.
- The IMX519 tuning file path (`/usr/share/libcamera/ipa/rpi/vc4/imx519.json`) is assumed. `--verify` warns if it is missing.
