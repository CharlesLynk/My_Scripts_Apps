#!/bin/bash
#
################################################################################
# Raspberry Pi Camera Stream Setup (Arducam IMX519, Pi Zero 2 W)
#
# Turns the Pi into an RTSP camera. Motion detection, recording, and viewing
# happen on the homelab (Frigate). See README.md in this folder.
#
#   sudo bash camera-setup.sh           Install packages, enable IMX519, reboot
#   sudo bash camera-setup.sh --verify  Check detection, test still, test stream
#   sudo bash camera-setup.sh --stream  Install MediaMTX and start the RTSP stream
#
# Run the default mode, reboot, then --verify, then --stream.
################################################################################

set -euo pipefail

readonly SCRIPT_VERSION="2.0.0"
readonly CONFIG_FILE="/boot/firmware/config.txt"
readonly TUNING_FILE="/usr/share/libcamera/ipa/rpi/vc4/imx519.json"
readonly MEDIAMTX_VERSION="${MEDIAMTX_VERSION:-1.21.4}"
readonly MEDIAMTX_BIN="/usr/local/bin/mediamtx"
readonly MEDIAMTX_CONF_DIR="/etc/mediamtx"
readonly MEDIAMTX_CONF="${MEDIAMTX_CONF_DIR}/mediamtx.yml"
readonly STREAM_SECRET_FILE="${MEDIAMTX_CONF_DIR}/frigate-password"
readonly SERVICE_NAME="camera-stream"
readonly STREAM_PATH="cam"
readonly RTSP_PORT=8554
readonly STREAM_WIDTH=1280
readonly STREAM_HEIGHT=720
readonly STREAM_FPS=15
readonly STREAM_BITRATE=2000000
readonly APT_RETRY_ATTEMPTS=5
readonly APT_RETRY_DELAY=4
readonly MIN_DISK_SPACE_MB=500

readonly RED='\033[0;31m'
readonly GREEN='\033[0;32m'
readonly YELLOW='\033[1;33m'
readonly BLUE='\033[0;34m'
readonly CYAN='\033[0;36m'
readonly BOLD='\033[1m'
readonly NC='\033[0m'

log_info() {
    echo -e "${BLUE}[i]${NC} $*"
}

log_success() {
    echo -e "${GREEN}[✓]${NC} $*"
}

log_warning() {
    echo -e "${YELLOW}[⚠]${NC} $*"
}

log_error() {
    echo -e "${RED}[✗]${NC} $*" >&2
}

log_step() {
    echo -e "\n${CYAN}[→]${NC} ${BOLD}$*${NC}"
}

command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Prompt user for yes/no confirmation (default no)
prompt_yn() {
    local answer=""
    read -p "$1 [y/n]: " -r answer
    [[ "$answer" =~ ^[Yy]$ ]]
}

usage() {
    cat <<EOF
Camera Stream Setup v${SCRIPT_VERSION} (Arducam IMX519)

Usage: sudo bash $(basename "$0") [option]

  (no option)  Install packages and enable the IMX519 overlay, then reboot
  --verify     Check camera detection, take a test still, probe the stream
  --stream     Install MediaMTX and start the RTSP stream for Frigate
  -h, --help   Show this help
EOF
}

check_root() {
    if [[ $EUID -ne 0 ]]; then
        log_error "Run with sudo: sudo bash $0 $*"
        exit 1
    fi
}

# The camera runs as the normal user, so test files and the stream service do too
resolve_run_user() {
    RUN_USER="${SUDO_USER:-}"
    if [[ -z "$RUN_USER" || "$RUN_USER" == "root" ]]; then
        log_error "Run this script with sudo from your normal user account, not as root."
        exit 1
    fi
    USER_HOME="$(getent passwd "$RUN_USER" | cut -d: -f6)"
}

check_board() {
    local model=""
    if [[ -r /proc/device-tree/model ]]; then
        model="$(tr -d '\0' < /proc/device-tree/model)"
    fi
    log_info "Board: ${model:-unknown}"
    if [[ "$model" != *"Raspberry Pi"* ]]; then
        log_warning "This does not look like a Raspberry Pi. Continuing anyway."
    fi
}

check_disk_space() {
    local avail_mb
    avail_mb=$(df -Pm / | awk 'NR==2 {print $4}')
    if (( avail_mb < MIN_DISK_SPACE_MB )); then
        log_error "Only ${avail_mb} MB free on /; need at least ${MIN_DISK_SPACE_MB} MB"
        exit 1
    fi
}

# Retry a command (apt mirrors, DNS, and GitHub are often flaky on first boot)
retry_cmd() {
    local attempt=1
    until "$@"; do
        if (( attempt >= APT_RETRY_ATTEMPTS )); then
            log_error "Command failed after ${attempt} attempts: $*"
            return 1
        fi
        log_warning "Attempt ${attempt} failed, retrying in ${APT_RETRY_DELAY}s..."
        sleep "$APT_RETRY_DELAY"
        attempt=$((attempt + 1))
    done
}

install_packages() {
    log_step "Installing camera packages"
    export DEBIAN_FRONTEND=noninteractive
    retry_cmd apt-get update
    retry_cmd apt-get install -y --no-install-recommends \
        rpicam-apps ffmpeg curl ca-certificates
    log_success "Packages installed"
}

ensure_video_group() {
    if [[ " $(id -nG "$RUN_USER") " == *" video "* ]]; then
        return
    fi
    usermod -aG video "$RUN_USER"
    log_success "Added ${RUN_USER} to the video group (takes effect after reboot)"
}

configure_boot() {
    log_step "Configuring ${CONFIG_FILE}"
    if [[ ! -f "$CONFIG_FILE" ]]; then
        log_error "${CONFIG_FILE} not found. This script expects Raspberry Pi OS Bookworm or Trixie."
        exit 1
    fi

    local backup="${CONFIG_FILE}.bak.$(date +%Y%m%d-%H%M%S)"
    cp "$CONFIG_FILE" "$backup"
    log_info "Backup saved to ${backup}"

    # camera_auto_detect is global, so a new line goes at the top of the file.
    # Appending it could land inside a conditional [section].
    if grep -qE '^[[:space:]]*camera_auto_detect=' "$CONFIG_FILE"; then
        sed -i -E 's/^[[:space:]]*camera_auto_detect=.*/camera_auto_detect=0/' "$CONFIG_FILE"
    else
        sed -i '1i camera_auto_detect=0' "$CONFIG_FILE"
    fi
    log_success "camera_auto_detect=0 set"

    if grep -qE '^[[:space:]]*dtoverlay=imx519' "$CONFIG_FILE"; then
        log_info "dtoverlay=imx519 already present"
    else
        # A fresh [all] header makes the overlay apply on every model
        printf '\n# Arducam IMX519 (added by camera-setup.sh)\n[all]\ndtoverlay=imx519\n' >> "$CONFIG_FILE"
        log_success "Added dtoverlay=imx519 under [all]"
    fi
}

run_setup() {
    check_board
    check_disk_space
    install_packages
    ensure_video_group
    configure_boot

    echo
    echo "Setup complete. The IMX519 overlay loads on the next boot."
    echo "After reboot: sudo bash $(basename "$0") --verify"
    echo
    if prompt_yn "Reboot now?"; then
        reboot
    fi
}

verify_camera() {
    log_step "Checking camera detection"
    if ! command_exists rpicam-still; then
        log_error "rpicam-still not found. Run setup first: sudo bash $0"
        return 1
    fi

    local listing
    listing="$(rpicam-still --list-cameras 2>&1 || true)"
    echo "$listing"
    if ! grep -qi 'imx519' <<< "$listing"; then
        log_error "IMX519 not detected. Check the camera ribbon (Pi Zero 2 W needs a 22-to-15-pin adapter),"
        log_error "confirm camera_auto_detect=0 and dtoverlay=imx519 in ${CONFIG_FILE}, and reboot."
        return 1
    fi
    log_success "IMX519 detected"

    if [[ -f "$TUNING_FILE" ]]; then
        log_success "Tuning file present"
    else
        log_warning "Tuning file ${TUNING_FILE} missing. Colors may look off."
    fi

    log_step "Taking a test still"
    local out="${USER_HOME}/camera-test.jpg"
    if runuser -u "$RUN_USER" -- rpicam-still -n -t 2000 --width 1280 --height 720 -o "$out"; then
        log_success "Test image saved to ${out}"
    else
        log_error "Test capture failed. Check the ribbon cable and the camera troubleshooting notes in README.md."
        return 1
    fi
}

mediamtx_arch() {
    case "$(uname -m)" in
        aarch64) echo "arm64" ;;
        armv7l)  echo "armv7" ;;
        armv6l)  echo "armv6" ;;
        *)
            log_error "Unsupported architecture: $(uname -m)"
            exit 1
            ;;
    esac
}

install_mediamtx() {
    log_step "Installing MediaMTX v${MEDIAMTX_VERSION}"
    local arch url tmp
    arch="$(mediamtx_arch)"
    url="https://github.com/bluenviron/mediamtx/releases/download/v${MEDIAMTX_VERSION}/mediamtx_v${MEDIAMTX_VERSION}_linux_${arch}.tar.gz"
    tmp="$(mktemp -d)"

    retry_cmd curl -fsSL -o "${tmp}/mediamtx.tar.gz" "$url"
    # Extract the whole archive; member paths in release tarballs vary
    tar -xzf "${tmp}/mediamtx.tar.gz" -C "$tmp"
    install -m 755 "${tmp}/mediamtx" "$MEDIAMTX_BIN"
    rm -rf "$tmp"
    log_success "Installed ${MEDIAMTX_BIN}"
}

# Hex keeps the password URL-safe for rtsp:// URLs
ensure_stream_password() {
    install -d -m 755 "$MEDIAMTX_CONF_DIR"
    if [[ -s "$STREAM_SECRET_FILE" ]]; then
        log_info "Reusing the existing stream password"
        return
    fi
    (umask 077; openssl rand -hex 24 > "$STREAM_SECRET_FILE")
    log_success "Generated stream password in ${STREAM_SECRET_FILE}"
}

write_stream_config() {
    log_step "Writing ${MEDIAMTX_CONF}"
    local pass cmd
    pass="$(< "$STREAM_SECRET_FILE")"

    # MediaMTX runs this command when the path starts. It publishes back to the
    # same server over localhost. -c copy avoids re-encoding on the Zero 2 W.
    cmd="rpicam-vid -t 0 -n --inline --codec h264 --width ${STREAM_WIDTH} --height ${STREAM_HEIGHT} --framerate ${STREAM_FPS} --bitrate ${STREAM_BITRATE} -o - | ffmpeg -loglevel warning -f h264 -framerate ${STREAM_FPS} -i - -c copy -f rtsp -rtsp_transport tcp rtsp://127.0.0.1:${RTSP_PORT}/${STREAM_PATH}"

    cat > "$MEDIAMTX_CONF" <<EOF
# Generated by camera-setup.sh. Frigate reads the stream as user "frigate".
rtspAddress: :${RTSP_PORT}
rtmp: false
hls: false
webrtc: false
srt: false
api: false
metrics: false
pprof: false

authMethod: internal
# Anonymous access is limited to publishing from localhost (the runOnInit
# ffmpeg). Frigate reads with its own user. Empty path means all paths.
authInternalUsers:
  - user: any
    pass:
    ips: ['127.0.0.1', '::1']
    permissions:
      - action: publish
        path:
  - user: frigate
    pass: ${pass}
    ips: []
    permissions:
      - action: read
        path:

paths:
  ${STREAM_PATH}:
    runOnInit: "${cmd}"
    runOnInitRestart: true
EOF

    # The service user needs to read the config and password; others should not
    chown "root:${RUN_USER}" "$MEDIAMTX_CONF" "$STREAM_SECRET_FILE"
    chmod 640 "$MEDIAMTX_CONF" "$STREAM_SECRET_FILE"
    log_success "Config written"
}

install_service() {
    log_step "Installing ${SERVICE_NAME} service"
    cat > "/etc/systemd/system/${SERVICE_NAME}.service" <<EOF
[Unit]
Description=RTSP camera stream for Frigate (Arducam IMX519 via MediaMTX)
After=network-online.target
Wants=network-online.target

[Service]
User=${RUN_USER}
SupplementaryGroups=video
ExecStart=${MEDIAMTX_BIN} ${MEDIAMTX_CONF}
Restart=on-failure
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF
    systemctl daemon-reload
    systemctl enable "${SERVICE_NAME}.service"
    systemctl restart "${SERVICE_NAME}.service"
    sleep 5

    if systemctl is-active --quiet "${SERVICE_NAME}.service"; then
        log_success "Service is running"
    else
        log_error "Service failed to start. See: journalctl -u ${SERVICE_NAME} -e"
        exit 1
    fi
}

stream_address() {
    local ip=""
    if command_exists tailscale; then
        ip="$(tailscale ip -4 2>/dev/null | head -n 1 || true)"
    fi
    if [[ -z "$ip" ]]; then
        ip="$(hostname -I 2>/dev/null | awk '{print $1}')"
    fi
    echo "${ip:-<pi-address>}"
}

setup_stream() {
    log_step "Preparing the RTSP stream"
    if ! command_exists rpicam-vid || ! command_exists ffmpeg; then
        log_error "Missing rpicam-vid or ffmpeg. Run setup first: sudo bash $0"
        exit 1
    fi
    local listing
    listing="$(rpicam-still --list-cameras 2>&1 || true)"
    if ! grep -qi 'imx519' <<< "$listing"; then
        log_error "IMX519 not detected. Run --verify and fix the camera before streaming."
        exit 1
    fi

    install_mediamtx
    ensure_stream_password
    write_stream_config
    install_service

    local pass
    pass="$(< "$STREAM_SECRET_FILE")"
    echo
    log_success "Stream URL for Frigate:"
    echo "    rtsp://frigate:${pass}@$(stream_address):${RTSP_PORT}/${STREAM_PATH}"
    echo
    log_info "Put the password in the homelab's frigate/.env as FRIGATE_RTSP_PASSWORD."
    log_info "Re-print it any time with: sudo cat ${STREAM_SECRET_FILE}"
    log_info "Logs: journalctl -u ${SERVICE_NAME} -f"
}

verify_stream() {
    log_step "Checking the RTSP stream"
    if [[ ! -f "/etc/systemd/system/${SERVICE_NAME}.service" ]]; then
        log_info "Stream service not installed yet. Run: sudo bash $0 --stream"
        return 0
    fi
    if ! systemctl is-active --quiet "${SERVICE_NAME}.service"; then
        log_error "Service ${SERVICE_NAME} is not running. See: journalctl -u ${SERVICE_NAME} -e"
        return 1
    fi
    if ! command_exists ffprobe; then
        log_warning "ffprobe not found; skipping the stream probe"
        return 0
    fi

    local url info
    url="rtsp://frigate:$(< "$STREAM_SECRET_FILE")@127.0.0.1:${RTSP_PORT}/${STREAM_PATH}"
    if info="$(timeout 20 ffprobe -v error -rtsp_transport tcp \
            -show_entries stream=codec_name,width,height -of csv=p=0 "$url" 2>&1)"; then
        log_success "Stream is live: ${info}"
    else
        log_error "Stream probe failed: ${info}"
        return 1
    fi
}

main() {
    local mode="${1:-}"

    case "$mode" in
        -h|--help)
            usage
            exit 0
            ;;
    esac

    check_root "$@"
    resolve_run_user

    case "$mode" in
        "")
            run_setup
            ;;
        --verify)
            verify_camera
            verify_stream
            ;;
        --stream)
            setup_stream
            ;;
        *)
            log_error "Unknown option: $mode"
            usage
            exit 1
            ;;
    esac
}

main "$@"
