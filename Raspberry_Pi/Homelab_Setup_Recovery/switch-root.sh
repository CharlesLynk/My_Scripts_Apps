#!/bin/bash
#
# === Raspberry Pi Root Switcher ===
#
# Usage:
#   sudo ./switch-root.sh
#
# Options (interactive menu):
#   1) Show current setting
#   2) Verify initial configuration (SSH and WiFi)
#   3) Configure USB boot partitions (add power settings to config.txt)
#   4) Switch to current-device-only (boot+root on current device)
#   5) Switch to Hybrid (boot on current device, root on another)
#   6) Clone OS to USB target (boot+root)
#   7) Flash fresh Raspberry Pi OS to USB (wget + sha256sum + dd)
#   8) Configure EEPROM boot order (SD/USB priority)
#   q) Quit
#
# Notes:
#   - This script edits /boot/firmware/cmdline.txt in place.
#   - Before the first change, it saves a copy as cmdline.txt.bak.
#   - If a backup already exists, you will be prompted whether to overwrite it.
#   - If the Pi fails to boot, mount the SD/USB on another computer
#     and restore by copying cmdline.txt.bak back to cmdline.txt.
#   - Scripts are always stored in boot partition (/boot/firmware/) for consistency.
#   - Options 3-6 check and configure ALL USB boot partitions AFTER the main
#     operation completes successfully (power settings and scripts).
#   - Option 7 configures only the newly flashed USB device.
#   - Initial configuration checks (SSH and WiFi) run before options 4-7 (boot changes)
#     to ensure you maintain remote access after switching boot configuration.
#

set -euo pipefail

# Debug flag for timestamps (set with --debug flag)
DEBUG_MODE=false

BOOT_DIR="/boot/firmware"
CMDLINE="$BOOT_DIR/cmdline.txt"
BACKUP="$BOOT_DIR/cmdline.txt.bak"

# Memory and size constants
MIN_MEMORY_COPY_MB=512
MIN_MEMORY_FLASH_MB=256
MIN_DISK_SPACE_FLASH_MB=2048
DEFAULT_FLASH_CACHE_DIR="/var/cache/switch-root/images"

# WiFi verification settings (connection not required)
readonly WIFI_SSID="LynkNet"
readonly WIFI_PSK="Tay98Chase97Wifi"

# Connection retry and progress monitoring constants
readonly CONNECTION_RETRY_MAX_ATTEMPTS=15
readonly CONNECTION_RETRY_DELAY=2
readonly WGET_RETRY_ATTEMPTS=15
readonly WGET_RETRY_DELAY=2
readonly PROGRESS_CHECK_INTERVAL=5

# Default password to enforce when enabling SSH password authentication
readonly DEFAULT_SSH_PASSWORD="password"

# Enforce SSH continuity when switching boot/root targets
REQUIRE_SSH_CONTINUITY=true

# PARTUUID duplication handling mode: fail|fix|warn
PARTUUID_DUP_MODE="fail"

# Terminal color codes
readonly RED='\033[0;31m'
readonly GREEN='\033[0;32m'
readonly YELLOW='\033[1;33m'
readonly BLUE='\033[0;34m'
readonly CYAN='\033[0;36m'
readonly BOLD='\033[1m'
readonly NC='\033[0m' # No Color

# Set up logging
LOG_TIMESTAMP=$(date "+%Y%m%d_%H%M%S")
LOG_FILE="/tmp/switch-root-${LOG_TIMESTAMP}.log"
exec > >(tee -a "$LOG_FILE") 2>&1

# Download control state (used by Ctrl+C menu)
DOWNLOAD_MENU_ACTION=""

# Boot-order continuity check summary (updated by check_boot_order_ssh_continuity)
BOOT_ORDER_LAST_CHECKED=0
BOOT_ORDER_LAST_FAILURES=0
BOOT_ORDER_LAST_SKIPPED=0
BOOT_ORDER_VALIDATED_BOOT_PARTS=()
BOOT_ORDER_VALIDATED_ROOT_PARTS=()
BOOT_ORDER_NEXT_TARGET_BOOT=""
BOOT_ORDER_NEXT_TARGET_ROOT=""

# EEPROM USB boot profile targets (Pi 4 USB boot hardening)
readonly EEPROM_USB_DISCOVER_TIMEOUT_TARGET=20000
readonly EEPROM_USB_PWR_OFF_TIME_TARGET=1000

# Conditional timestamp - only shown when --debug flag is active
timestamp() {
  if [[ "$DEBUG_MODE" == "true" ]]; then
    date "+%F %T"
  fi
}

# Helper: Unified error handler with optional debug context
# Usage: err "message" [--debug]
# Optional: Set ERROR_HINT="hint text" before calling for context-specific help
err() {
  local show_debug=false
  local message=""
  
  # Parse arguments
  while [[ $# -gt 0 ]]; do
    case "$1" in
      --debug)
        show_debug=true
        shift
        ;;
      *)
        message="$*"
        break
        ;;
    esac
  done
  
  log_error "$message"
  echo "" >&2
  
  # Show context-specific hint if available
  if [[ -n "${ERROR_HINT:-}" ]]; then
    log_info "Suggestion:"
    echo "  $ERROR_HINT" >&2
    echo "" >&2
  fi
  
  if [[ "$show_debug" == "true" ]]; then
    echo "=== Debug Information ===" >&2
    echo "Active mounts (boot/mnt):" >&2
    mount | grep -E "(boot|/mnt)" >&2 || echo "  None found" >&2
    echo "" >&2
    echo "Block devices:" >&2
    lsblk -o NAME,SIZE,TRAN,TYPE,FSTYPE,LABEL,MOUNTPOINT,PARTUUID >&2 || true
    echo "" >&2
    echo "Recent kernel messages (last 15):" >&2
    dmesg | tail -15 >&2 || true
    echo "========================" >&2
    echo "" >&2
    echo "Recovery steps:" >&2
    echo "  1. Review log file: cat $LOG_FILE" >&2
    echo "  2. Check device connections and power" >&2
    echo "  3. Verify partition table: sudo parted /dev/sda print (or /dev/mmcblk0)" >&2
    echo "  4. Check for kernel errors: dmesg | tail -50" >&2
    echo "  5. Unmount stuck devices: for d in /mnt/*/; do sudo umount -l \$d; done" >&2
    echo "========================" >&2
    echo "" >&2
  else
    echo "Log file saved to: $LOG_FILE" >&2
    echo "Review the log for complete details." >&2
    echo "Tip: Re-run with debug context if needed" >&2
    echo "" >&2
  fi
  
  cleanup_failed_mounts 2>/dev/null || true
  exit 1
}

# Helper: Success/info messages with color and timestamp
log_success() { echo -e "$(timestamp) ${GREEN}[✓]${NC} $*"; }

# Additional logging functions for consistency
log_info()    { echo -e "$(timestamp) ${BLUE}[i]${NC} $*"; }
log_warning() { echo -e "$(timestamp) ${YELLOW}[⚠]${NC} $*"; }
log_error()   { echo -e "$(timestamp) ${RED}[✗]${NC} $*" >&2; }
log_step()    { echo -e "\n$(timestamp) ${CYAN}[→]${NC} ${BOLD}$*${NC}"; }

# Print a separator line
print_separator() {
    echo "================================================================================"
}

# Helper: Prompt for yes/no confirmation
# Usage: prompt_yn "Question?" [default_yes]
# Returns: 0 (success) for yes, 1 (failure) for no
# Note: This follows Unix exit code convention where 0 = success/true
prompt_yn() {
  local prompt="$1"
  local default_yes="${2:-false}"
  local answer=""
  local answer_normalized=""
  
  if [[ "$default_yes" == "true" ]]; then
    read -rp "$prompt [Y/n]: " answer
    answer_normalized="${answer,,}"
    answer_normalized="${answer_normalized//[[:space:]]/}"
    [[ -z "$answer_normalized" || "$answer_normalized" == "y" || "$answer_normalized" == "yes" ]]
  else
    read -rp "$prompt [y/N]: " answer
    answer_normalized="${answer,,}"
    answer_normalized="${answer_normalized//[[:space:]]/}"
    [[ "$answer_normalized" == "y" || "$answer_normalized" == "yes" ]]
  fi
}

# Helper: Prompt with formatted context section
# Usage: prompt_with_context "Title" "Description" "Question?"
# Returns: 0 for yes, 1 for no
prompt_with_context() {
  local title="$1"
  local description="$2"
  local prompt="$3"
  
  echo
  echo "=== $title ==="
  if [[ -n "$description" ]]; then
    echo "$description"
    echo
  fi
  prompt_yn "$prompt"
}

# Print a boxed message for critical warnings or important information
# Usage: print_box "Title" "Line 1" "Line 2" ...
print_box() {
  local title="$1"
  shift
  local lines=("$@")
  
  echo
  echo "╔══════════════════════════════════════════════════════════════════════════╗"
  printf "║  %-72s║\n" "$title"
  echo "╠══════════════════════════════════════════════════════════════════════════╣"
  for line in "${lines[@]}"; do
    printf "║  %-72s║\n" "$line"
  done
  echo "╚══════════════════════════════════════════════════════════════════════════╝"
  echo
}

################################################################################
# NETWORK RETRY AND DOWNLOAD PROGRESS HELPERS
################################################################################

################################################################################
# TWO-STAGE RETRY SYSTEM FOR NETWORK OPERATIONS
################################################################################

# Stage 1: Retry connection establishment with configurable attempts
# Usage: retry_until_connected [--attempts N] [--delay SECONDS] command args...
# Returns: 0 if connected, 1 if all attempts exhausted
retry_until_connected() {
    local max_attempts=${CONNECTION_RETRY_MAX_ATTEMPTS}
    local retry_delay=${CONNECTION_RETRY_DELAY}
    local attempt=1
    local exitCode=0
    
    # Parse optional arguments
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --attempts)
                max_attempts="$2"
                shift 2
                ;;
            --delay)
                retry_delay="$2"
                shift 2
                ;;
            *)
                break
                ;;
        esac
    done
    
    while (( attempt <= max_attempts )); do
        if "$@" >/dev/null 2>&1; then
            # Clear any progress line before returning
            if (( attempt > 1 )); then
                printf "\r%80s\r" " "
                log_success "Connection established on attempt ${attempt}"
            fi
            return 0
        else
            exitCode=$?
        fi
        
        if (( attempt < max_attempts )); then
            printf "\r${YELLOW}[⚠] Retry${NC} Connection attempt %d/%d failed, retrying in %ds...     " \
                "$attempt" "$max_attempts" "$retry_delay"
            sleep $retry_delay
        fi
        attempt=$((attempt + 1))
    done
    
    # Clear the progress line and print final error
    printf "\r%80s\r" " "
    log_error "Failed to establish connection after ${max_attempts} attempts"
    return $exitCode
}

# Stage 2: Monitor download progress with user interaction
# Usage: monitor_download_progress "operation_name" command args...
# Returns: 0 if completed successfully, 1 if user aborted or failed, 2 if user requested retry
monitor_download_progress() {
    local operation_name="$1"
    shift
    local temp_log="/tmp/download_progress_$$.log"
    local pid_file="/tmp/download_pid_$$.pid"

  DOWNLOAD_MENU_ACTION=""
    
    # Start operation in background and capture PID
    "$@" > "$temp_log" 2>&1 &
    local download_pid=$!
    echo $download_pid > "$pid_file"
    
    log_info "${operation_name} started (PID: ${download_pid})"
    log_info "Monitoring progress - checking every ${PROGRESS_CHECK_INTERVAL} seconds"
    echo
    log_info "Commands available while downloading:"
    log_info "  - Press Ctrl+C to pause and show options menu"
    echo
    
    local last_size=0
    local stall_count=0
    local check_count=0
    local start_time=$(date +%s)
    
    # Set up trap for Ctrl+C to show menu instead of exiting
    local original_int_trap=$(trap -p INT)
    trap 'show_download_menu "$operation_name" "$download_pid" "$temp_log" "$pid_file"' INT
    
    while kill -0 $download_pid 2>/dev/null; do
      if [[ -n "$DOWNLOAD_MENU_ACTION" ]]; then
        case "$DOWNLOAD_MENU_ACTION" in
          retry)
            log_warning "Download stopped by user"
            kill $download_pid 2>/dev/null
            wait $download_pid 2>/dev/null
            rm -f "$temp_log" "$pid_file" 2>/dev/null
            if [[ -n "$original_int_trap" ]]; then
              eval "$original_int_trap"
            else
              trap - INT
            fi
            return 2
            ;;
          abort)
            log_warning "User requested abort"
            kill $download_pid 2>/dev/null
            wait $download_pid 2>/dev/null
            rm -f "$temp_log" "$pid_file" 2>/dev/null
            if [[ -n "$original_int_trap" ]]; then
              eval "$original_int_trap"
            else
              trap - INT
            fi
            err "Script aborted by user"
            ;;
        esac
      fi

        sleep $PROGRESS_CHECK_INTERVAL
        check_count=$((check_count + 1))
        
        local current_time=$(date +%s)
        local elapsed=$((current_time - start_time))
        local elapsed_min=$((elapsed / 60))
        local elapsed_sec=$((elapsed % 60))
        
        # Check if process is making progress (reading/writing data)
        local current_size=0
        if [[ -f "/proc/${download_pid}/io" ]]; then
            current_size=$(awk '/^read_bytes:/ {sum+=$2} /^write_bytes:/ {sum+=$2} END {print sum}' /proc/${download_pid}/io 2>/dev/null || echo 0)
        fi
        
        # Detect stalls (no I/O activity)
        if (( current_size == last_size && last_size > 0 )); then
            stall_count=$((stall_count + 1))
            if (( stall_count >= 3 )); then
                log_warning "Download appears stalled (no I/O activity for $((stall_count * PROGRESS_CHECK_INTERVAL))s)"
                log_info "Press Ctrl+C for options menu"
                stall_count=0  # Reset to avoid spam
            fi
        else
            stall_count=0
        fi
        last_size=$current_size
        
        # Show progress update every 3 checks (15s with default interval)
        if (( check_count % 3 == 0 )); then
            local size_mb=$((current_size / 1024 / 1024))
            printf "${BLUE}[PROGRESS]${NC} %s - %dm %ds elapsed, %dMB transferred\n" \
                "$operation_name" "$elapsed_min" "$elapsed_sec" "$size_mb"
        fi
    done
    
    # Restore previous INT handler
    if [[ -n "$original_int_trap" ]]; then
      eval "$original_int_trap"
    else
      trap - INT
    fi
    
    # Check exit status
    wait $download_pid
    local exit_code=$?
    
    # Cleanup
    rm -f "$temp_log" "$pid_file" 2>/dev/null
    
    if (( exit_code == 0 )); then
        local total_time=$(($(date +%s) - start_time))
        log_success "${operation_name} completed successfully in ${total_time}s"
        return 0
    else
        log_error "${operation_name} failed with exit code ${exit_code}"
        return 1
    fi
}

# Interactive menu shown when user presses Ctrl+C during download
# Usage: show_download_menu "operation_name" pid log_file pid_file
show_download_menu() {
    local operation_name="$1"
    local download_pid="$2"
    local temp_log="$3"
    local pid_file="$4"
    
    echo
    echo "═══════════════════════════════════════════════════════════════"
    echo "  Download Control Menu - ${operation_name}"
    echo "═══════════════════════════════════════════════════════════════"
    echo
    
    # Check if process is still running
    if kill -0 $download_pid 2>/dev/null; then
        log_info "Status: RUNNING (PID: ${download_pid})"
    else
        log_info "Status: STOPPED"
    fi
    echo
    
    # Show recent log output
    if [[ -f "$temp_log" ]]; then
        echo "Recent output:"
        tail -n 5 "$temp_log" 2>/dev/null | sed 's/^/  │ /'
        echo
    fi
    
    echo "Options:"
    echo "  1) Resume monitoring (continue waiting)"
    echo "  2) Stop and retry download (kills current attempt)"
    echo "  3) View full log"
    echo "  4) Abort and exit script"
    echo
    read -p "Enter choice [1-4]: " choice
    
    case "$choice" in
        1)
            log_success "Resuming monitoring..."
            echo
            # Return to monitoring loop (trap will be re-established)
            ;;
        2)
            if prompt_yn "Retry download now?"; then
          DOWNLOAD_MENU_ACTION="retry"
            else
          DOWNLOAD_MENU_ACTION="abort"
            fi
            ;;
        3)
            if [[ -f "$temp_log" ]]; then
                echo
                echo "═══════════════════════════════════════════════════════════════"
                echo "  Full Log Output"
                echo "═══════════════════════════════════════════════════════════════"
                cat "$temp_log"
                echo "═══════════════════════════════════════════════════════════════"
                echo
                read -p "Press Enter to return to menu..."
                # Show menu again
                show_download_menu "$operation_name" "$download_pid" "$temp_log" "$pid_file"
            else
                log_warning "Log file not found"
                sleep 2
            fi
            ;;
        4)
          log_warning "User requested abort"
          DOWNLOAD_MENU_ACTION="abort"
            ;;
        *)
            log_warning "Invalid choice, resuming monitoring..."
            echo
            ;;
    esac
}

# Set up cleanup trap (only INT/TERM to avoid interfering with intentional mounts)
trap 'cleanup_failed_mounts 2>/dev/null || true; echo "\nInterrupted - log saved to: $LOG_FILE" >&2' INT TERM

# Rsync exclude patterns for filesystem copy operations
RSYNC_EXCLUDES=(
  "/boot"
  "/boot/firmware"
  "/proc/*"
  "/sys/*"
  "/dev/*"
  "/run/*"
  "/tmp/*"
  "/mnt/*"
  "/media/*"
  "/lost+found"
  "/swapfile"
  "/var/swap"
  "/var/cache/*"
  "/var/tmp/*"
  "/var/log/*.log"
  "/var/log/*.log.*"
  "/home/*/.cache"
  "/home/*/.local/share/Trash"
  "/root/.cache"
  "/var/lib/docker"
  "/var/lib/containerd"
  "/var/lib/apt/lists/*"
)

################################################################################
# SCRIPT DISCOVERY AND COPYING
################################################################################

# Helper: Get list of homelab script files
get_script_files() {
  echo "homelab-setup.sh" "switch-root.sh" "README.md" "TROUBLESHOOTING.md"
}

# Helper: Get script search locations
get_script_locations() {
  echo "$BOOT_DIR" "/root" "/home/admin" "$(dirname "$(readlink -f "$0")")" "/usr/local/bin"
}

# Prompt and copy scripts to boot partition (accessible even if root won't boot)
# Note: This function prompts the user before copying
copy_scripts_to_boot() {
  local boot_mount="$1"
  local target_name="$2"  # Description for logging
  
  # Find scripts in current system using helper
  local script_locations=($(get_script_locations))
  local script_files=($(get_script_files))
  
  local files_found=()
  local files_to_copy=()
  
  for file in "${script_files[@]}"; do
    for loc in "${script_locations[@]}"; do
      if [[ -f "$loc/$file" ]]; then
        files_found+=("$file")
        files_to_copy+=("$loc/$file")
        break  # Found this file, move to next
      fi
    done
  done
  
  if [[ ${#files_to_copy[@]} -eq 0 ]]; then
    log_info "No homelab scripts found to copy"
    return 0
  fi
  
  # Check what already exists on target boot partition
  local missing_files=()
  
  for file in "${files_found[@]}"; do
    [[ ! -f "${boot_mount}/${file}" ]] && missing_files+=("$file")
  done
  
  if [[ ${#missing_files[@]} -eq 0 ]]; then
    log_success "All homelab scripts already present on boot partition of ${target_name}"
    return 0
  fi
  
  log_info "Scripts missing on boot partition of ${target_name}: ${missing_files[*]}"
  
  if ! prompt_yn "Copy missing scripts to boot partition?"; then
    log_info "Skipped script copy"
    return 0
  fi
  
  # Copy missing files
  local copied=0
  local failed=0
  for i in "${!files_found[@]}"; do
    local filename="${files_found[$i]}"
    local filepath="${files_to_copy[$i]}"
    
    # Skip if already exists
    [[ -f "${boot_mount}/${filename}" ]] && continue
    
    if cp "$filepath" "${boot_mount}/${filename}" 2>/dev/null; then
      if [[ "$filename" == *.sh ]]; then
        chmod 755 "${boot_mount}/${filename}" 2>/dev/null
      else
        chmod 644 "${boot_mount}/${filename}" 2>/dev/null
      fi
      log_success "Copied ${filename} to boot partition"
      copied=$((copied + 1))
    else
      log_error "Failed to copy ${filename}"
      failed=$((failed + 1))
    fi
  done
  
  if [[ $copied -gt 0 ]]; then
    log_success "  Copied $copied script(s) to boot partition of ${target_name}"
  fi
  
  if [[ $failed -gt 0 ]]; then
    log_warning "Failed to copy $failed file(s)"
    return 1
  fi
  
  return 0
}

################################################################################
# FILESYSTEM AND MOUNT HELPERS
################################################################################

# Helper: Wait for partitions to appear after partitioning
# Usage: wait_for_partitions "/dev/sda" 2
wait_for_partitions() {
  local device="$1"
  local expected_count="$2"
  local max_wait="${3:-15}"
  local attempt=1
  
  log_info "Waiting for $expected_count partition(s) to be recognized on $device..."
  
  while (( attempt <= max_wait )); do
    local current_count=$(lsblk -ln "$device" 2>/dev/null | grep -c part || echo 0)
    
    if (( current_count >= expected_count )); then
      log_success "All partitions detected"
      return 0
    fi
    
    # Show progress on every wait iteration
    log_info "Waiting for partitions... (${attempt}/${max_wait}) - ${current_count}/${expected_count} detected"
    
    if (( attempt < max_wait )); then
      partprobe 2>/dev/null || true
      udevadm settle 2>/dev/null || true
      sleep 1
    fi
    
    attempt=$((attempt + 1))
  done
  
  log_error "Only $(lsblk -ln "$device" 2>/dev/null | grep -c part || echo 0) of $expected_count partitions detected"
  return 1
}

# Helper: Check if mount point is writable
# Usage: check_writable "/mnt/target"
check_writable() {
  local mount_point="$1"
  local test_file="${mount_point}/.write_test_$$"
  
  if ! touch "$test_file" 2>/dev/null; then
    log_error "$mount_point is read-only or not writable"
    return 1
  fi
  rm "$test_file" 2>/dev/null
  return 0
}

# Helper: Verify filesystem has required files/directories
# Usage: verify_filesystem "/mnt/partition" "boot|root"
verify_filesystem() {
  local mount_point="$1"
  local fs_type="$2"
  local missing=()
  
  case "$fs_type" in
    boot)
      local patterns=("config.txt" "cmdline.txt" "start*.elf" "fixup*.dat")
      for pattern in "${patterns[@]}"; do
        if ! ls "$mount_point"/$pattern >/dev/null 2>&1; then
          missing+=("$pattern")
        fi
      done
      ;;
    root)
      local dirs=("/etc" "/home" "/usr" "/var" "/bin" "/sbin")
      for dir in "${dirs[@]}"; do
        if [[ ! -d "${mount_point}${dir}" ]]; then
          missing+=("$dir")
        fi
      done
      
      # Additional root-specific checks
      if [[ ! -f "${mount_point}/etc/fstab" ]]; then
        missing+=("/etc/fstab")
      fi
      ;;
    *)
      log_error "Unknown filesystem type: $fs_type"
      return 1
      ;;
  esac
  
  if [[ ${#missing[@]} -gt 0 ]]; then
    log_error "Missing ${fs_type} requirements: ${missing[*]}"
    return 1
  fi
  
  return 0
}

# Helper: Check available memory for operations
# Usage: check_memory_available 512
# Returns: 1 if below threshold (warning only, not fatal), 0 if sufficient
check_memory_available() {
  local required_mb="$1"
  local available_mb=$(free -m | awk '/^Mem:/ {print $7}')
  
  if (( available_mb < required_mb )); then
    log_warning "Low memory: ${available_mb}MB available, ${required_mb}MB recommended"
    log_warning "Operations may be slower than usual"
    return 1
  fi
  log_success "Sufficient memory available: ${available_mb}MB"
  return 0
}

# Helper: Detect and disable swap on target device
# Usage: disable_swap_if_active "/dev/sda2"
disable_swap_if_active() {
  local device="$1"
  
  if grep -q "$device" /proc/swaps 2>/dev/null; then
    log_success "Disabling active swap on $device"
    if ! swapoff "$device" 2>/dev/null; then
      log_error "Failed to disable swap on $device"
      return 1
    fi
  fi
  return 0
}

# Helper: Add USB power settings to config.txt if not present
# Returns: 0=success (configured or already present), 1=error
add_usb_power_settings() {
  local config_file="$1"
  
  [[ ! -f "$config_file" ]] && return 1
  
  # Remove legacy forced host overlay from older script versions.
  # This overlay is not required for Pi 4 USB boot and can reduce compatibility.
  remove_legacy_usb_host_overlay "$config_file" || return 1

  # Apply model-specific USB power tweak when appropriate.
  if ! append_usb_power_settings "$config_file"; then
    return 1
  fi

  sync
  return 0
}

# Helper: Append USB power settings block without heredoc (avoid editor indentation issues)
append_usb_power_settings() {
  local config_file="$1"

  # usb_max_current_enable is relevant to Pi 5 firmware behavior.
  # Do not add model-specific power flags on other boards.
  if ! is_raspberry_pi_5; then
    return 0
  fi

  if grep -q '^\s*usb_max_current_enable\s*=\s*1\s*$' "$config_file" 2>/dev/null; then
    return 0
  fi

  printf '\n# USB power tweak for Pi 5 (added by switch-root.sh)\n' >> "$config_file" || return 1
  printf 'usb_max_current_enable=1\n' >> "$config_file" || return 1
  return 0
}

# Helper: Detect Raspberry Pi 5 from device-tree model string
is_raspberry_pi_5() {
  local model=""
  model=$(tr -d '\0' < /proc/device-tree/model 2>/dev/null || true)
  [[ "$model" == *"Raspberry Pi 5"* ]]
}

# Helper: Remove legacy forced host overlay added by old versions of this script
remove_legacy_usb_host_overlay() {
  local config_file="$1"
  local tmp_file

  [[ -f "$config_file" ]] || return 1

  if ! grep -qE '^\s*dtoverlay=dwc2,dr_mode=host\s*$' "$config_file" 2>/dev/null && \
     ! grep -q 'Unlock power for USB Boot on USB-Compliant power supply (added by switch-root.sh)' "$config_file" 2>/dev/null; then
    return 0
  fi

  tmp_file=$(mktemp /tmp/switch-root-config-XXXXXX) || return 1
  if ! sed \
      -e '/^\s*dtoverlay=dwc2,dr_mode=host\s*$/d' \
      -e '/Unlock power for USB Boot on USB-Compliant power supply (added by switch-root.sh)/d' \
      "$config_file" > "$tmp_file"; then
    rm -f "$tmp_file" 2>/dev/null || true
    return 1
  fi

  if ! cp "$tmp_file" "$config_file"; then
    rm -f "$tmp_file" 2>/dev/null || true
    return 1
  fi

  rm -f "$tmp_file" 2>/dev/null || true
  log_warning "Removed legacy dtoverlay=dwc2,dr_mode=host from $(basename "$config_file") for safer USB boot compatibility"
  return 0
}

# Helper: Unified mount wrapper - perform action with temporary mount, auto-cleanup on exit
# Usage: with_mount "/dev/sda1" callback_function [--verify-boot (must be first)] [additional_args...]
# Note: If using --verify-boot, it must be the first additional argument after callback_function
with_mount() {
  local device="$1"
  local callback="$2"
  shift 2
  
  local verify_boot=false
  if [[ "${1:-}" == "--verify-boot" ]]; then
    verify_boot=true
    shift
  fi
  
  local mount_point="/mnt/temp_mount_$$_$(date +%s)"
  
  if ! safe_mount "$device" "$mount_point"; then
    log_error "Failed to mount $device"
    return 1
  fi
  
  # Verify boot files if requested
  if [[ "$verify_boot" == "true" ]]; then
    if ! verify_filesystem "$mount_point" "boot"; then
      safe_umount "$mount_point"
      log_error "Boot files verification failed on $device"
      return 1
    fi
  fi
  
  # Execute callback with mount point as first argument, plus any additional args
  local result=0
  "$callback" "$mount_point" "$device" "$@" || result=$?
  
  # Cleanup - use forced unmount on error to prevent orphaned mounts
  if (( result != 0 )); then
    safe_umount "$mount_point" || umount -l "$mount_point" 2>/dev/null || log_warning "Failed to unmount $mount_point"
  else
    safe_umount "$mount_point" || log_warning "Failed to unmount $mount_point"
  fi
  
  return $result
}

# Helper: Unified retry helper with customizable condition check
# Usage: retry_until "description" max_attempts interval check_function [check_args...]
retry_until() {
  local description="$1"
  local max_attempts="$2"
  local interval="${3:-1}"
  local check_function="$4"
  shift 4
  local attempt=1
  
  while (( attempt <= max_attempts )); do
    if "$check_function" "$@"; then
      return 0
    fi
    
    # Progress notification on every attempt
    if (( attempt > 1 )); then
      log_info "$description... (${attempt}/${max_attempts})"
    fi
    
    if (( attempt < max_attempts )); then
      partprobe 2>/dev/null || true
      udevadm settle 2>/dev/null || true
      sleep "$interval"
    fi
    
    attempt=$((attempt + 1))
  done
  
  log_error "Timeout: $description failed after ${max_attempts} attempts"
  return 1
}

# Helper: Resolve a disk basename from any block device path.
# Usage: resolve_disk_name_for_device "/dev/sda1" -> "sda"
resolve_disk_name_for_device() {
  local device="$1"
  local base="${device##*/}"

  if [[ -b "$device" ]]; then
    local type
    type=$(lsblk -no TYPE "$device" 2>/dev/null | head -n1 || true)
    if [[ "$type" == "part" ]]; then
      local pkname
      pkname=$(lsblk -no PKNAME "$device" 2>/dev/null | head -n1 || true)
      if [[ -n "$pkname" ]]; then
        echo "$pkname"
        return 0
      fi
    fi
  fi

  echo "$base"
}

# Helper: Read disk transport with kernel-name fallbacks for blank TRAN values.
get_disk_transport_for_name() {
  local disk_name="$1"
  local transport
  transport=$(lsblk -no TRAN "/dev/${disk_name}" 2>/dev/null | head -n1 || true)

  if [[ -z "$transport" && "$disk_name" =~ ^mmcblk ]]; then
    transport="mmc"
  fi
  if [[ -z "$transport" && "$disk_name" =~ ^nvme ]]; then
    transport="nvme"
  fi

  echo "$transport"
}

get_disk_hotplug_for_name() {
  local disk_name="$1"
  lsblk -ndo HOTPLUG "/dev/${disk_name}" 2>/dev/null | head -n1 || echo "0"
}

get_disk_removable_for_name() {
  local disk_name="$1"
  lsblk -ndo RM "/dev/${disk_name}" 2>/dev/null | head -n1 || echo "0"
}

is_disk_usb_bridge_for_name() {
  local disk_name="$1"
  local sys_path
  sys_path=$(readlink -f "/sys/block/${disk_name}" 2>/dev/null || true)
  [[ "$sys_path" == *"/usb"* ]]
}

# Helper: Classify a disk basename by transport class (sd|usb|pcie).
# Usage: disk_matches_transport_class "sda" "usb"
disk_matches_transport_class() {
  local disk_name="$1"
  local class="$2"
  local transport
  local hotplug
  local removable
  local usb_bridge=false

  transport=$(get_disk_transport_for_name "$disk_name")
  hotplug=$(get_disk_hotplug_for_name "$disk_name")
  removable=$(get_disk_removable_for_name "$disk_name")
  if is_disk_usb_bridge_for_name "$disk_name"; then
    usb_bridge=true
  fi

  case "$class" in
    sd)
      [[ "$transport" == "mmc" || "$disk_name" =~ ^mmcblk ]]
      ;;
    usb)
      # Never classify built-in SD/eMMC as USB.
      if [[ "$disk_name" =~ ^mmcblk ]]; then
        return 1
      fi
      [[ "$transport" == "usb" ]] && return 0
      [[ ( "$transport" == "ata" || "$transport" == "sata" ) && ( "$hotplug" == "1" || "$usb_bridge" == "true" ) ]] && return 0
      [[ -z "$transport" && "$removable" == "1" && "$hotplug" == "1" ]] && return 0
      [[ "$usb_bridge" == "true" ]] && return 0
      return 1
      ;;
    pcie)
      [[ "$transport" == "pci" || "$transport" == "pcie" || "$transport" == "nvme" || "$disk_name" =~ ^nvme ]]
      ;;
    *)
      return 1
      ;;
  esac
}

# Helper: Unified device validation with configurable check levels
# Usage: validate_target_device "/dev/sda" [check_level]
# Levels: minimal (exists+readable), standard (minimal+USB), full (standard+no_critical_mounts)
validate_target_device() {
  local device="$1"
  local check_level="${2:-full}"
  
  # Level: minimal - device exists and is readable
  if ! [[ -b "$device" ]]; then
    log_error "Device not found: $device"
    return 1
  fi
  
  if ! blockdev --getsize64 "$device" >/dev/null 2>&1; then
    log_error "Cannot read device: $device"
    return 1
  fi
  
  # Level: standard - check for acceptable removable/portable devices
  if [[ "$check_level" == "standard" || "$check_level" == "full" ]]; then
    local disk_name
    disk_name=$(resolve_disk_name_for_device "$device")
    local transport
    transport=$(get_disk_transport_for_name "$disk_name")
    local removable
    removable=$(get_disk_removable_for_name "$disk_name")

    # Accept USB-like disks, removable media, and SD/eMMC.
    if ! disk_matches_transport_class "$disk_name" "usb" && [[ "$removable" != "1" ]] && [[ ! "$disk_name" =~ ^mmcblk ]]; then
      log_error "Device $device is not detected as USB/removable/SD (transport: $transport, removable: $removable)"
      log_error "This appears to be internal storage - refusing to proceed"
      ERROR_HINT="Use a USB drive/adapter or SD card. Check: lsblk -o NAME,TRAN,RM"
      return 1
    fi
  fi
  
  # Level: full - also check for critical mounts
  if [[ "$check_level" == "full" ]]; then
    local critical_mounts=("/" "/boot" "/boot/firmware" "/home" "/var" "/usr" "/opt" "/srv")
    
    log_info "Checking for critical mounts on $device..."
    
    local has_critical=false
    while IFS= read -r part; do
      [[ -z "$part" ]] && continue
      local part_dev="/dev/$part"
      [[ -b "$part_dev" ]] || continue
      
      local mount_point=$(findmnt -n -o TARGET "$part_dev" 2>/dev/null || echo "")
      [[ -z "$mount_point" ]] && continue
      
      for critical in "${critical_mounts[@]}"; do
        # Check exact match OR if mount_point starts with critical path
        if [[ "$mount_point" == "$critical" ]] || [[ "$mount_point" == "$critical"/* ]]; then
          log_error "CRITICAL: $part_dev is mounted as $mount_point (under critical path $critical)"
          has_critical=true
        fi
      done
    done < <(lsblk -ln -o NAME "$device" | grep -v "^$(basename "$device")$")
  
    if [[ "$has_critical" == "true" ]]; then
      log_error "Cannot perform destructive operations on device with critical mounts"
      ERROR_HINT="Unmount critical filesystems first or select a different device"
      return 1
    fi
  
    log_success "No critical mounts detected on $device"
  fi
  
  return 0
}

# Helper: Determine parent disk name for a device (handles mappers/partitions)
get_parent_disk_name() {
  local dev="$1"
  [[ -n "$dev" ]] || return 0

  local resolved
  resolved=$(readlink -f "$dev" 2>/dev/null || echo "$dev")

  local parent
  parent=$(lsblk -no NAME -s "$resolved" 2>/dev/null | tail -n1)
  if [[ -n "$parent" ]]; then
    echo "$parent"
  fi
}

# Helper: Determine disk name backing a mountpoint (/, /boot/firmware)
get_mount_disk_name() {
  local mnt="$1"
  local src

  src=$(findmnt -n -o SOURCE -T "$mnt" 2>/dev/null || true)
  if [[ -z "$src" ]]; then
    echo ""
    return 0
  fi

  if [[ "$src" == "/dev/root" ]]; then
    src=$(lsblk -no PATH,MOUNTPOINT -l | awk -v m="$mnt" '$2==m {print $1; exit}')
  fi

  get_parent_disk_name "$src"
}

# Helper: Select and validate target storage device
# Usage: select_target_storage_device "prompt" "check_level"
# Sets USB_NAME and USB_DEV globals
select_target_storage_device() {
  local prompt_text="$1"
  local check_level="$2"
  local device_name=""

  echo
  log_info "Available storage devices:"
  print_storage_overview
  echo
  read -rp "$prompt_text" device_name

  # Match: sda-sdz, sdaa-sdzz (SCSI), mmcblk0-99 (SD/eMMC), nvme0n1-nvme9n9 (NVMe)
  local valid_disk_regex='^(sd[a-z]{1,2}|mmcblk[0-9]+|nvme[0-9]+n[0-9]+)$'
  [[ "$device_name" =~ $valid_disk_regex ]] || err "Invalid device name (expected: sda, sdaa, mmcblk0, nvme0n1, etc.)"

  USB_NAME="$device_name"
  USB_DEV="/dev/${USB_NAME}"

  # Validate target device
  if ! validate_target_device "$USB_DEV" "$check_level"; then
    err "Safety check failed: $USB_DEV validation failed"
  fi

  # Verify it's not the current root disk
  local current_root_disk
  current_root_disk="$(get_mount_disk_name "/")"
  if [[ -z "$current_root_disk" ]]; then
    err "Unable to determine current root disk - refusing to proceed"
  fi
  [[ "$USB_NAME" == "$current_root_disk" ]] && err "Cannot use current root disk"

  # Verify it's not the current boot disk
  local current_boot_dev
  local current_boot_disk=""
  current_boot_dev="$(findmnt -no SOURCE "$BOOT_DIR" 2>/dev/null || true)"
  current_boot_disk="$(get_mount_disk_name "$BOOT_DIR")"
  if [[ -z "$current_boot_disk" ]]; then
    err "Unable to determine current boot disk - refusing to proceed"
  fi
  if [[ "$USB_NAME" == "$current_boot_disk" ]]; then
    err "Cannot use current boot disk"
  fi

  log_success "Verified $USB_DEV is a suitable target"
}

# Helper: Unmount all partitions for a given device name (e.g., sda)
unmount_device_partitions() {
  local device_name="$1"

  for node in $(lsblk -ln -o NAME,PKNAME,TYPE | awk -v p="$device_name" '$3=="part" && $2==p {print "/dev/"$1}'); do
    disable_swap_if_active "$node" || true
    if mountpoint -q "$node" 2>/dev/null || grep -q "$node" /proc/mounts 2>/dev/null; then
      log_success "Unmounting $node"
      umount "$node" 2>/dev/null || umount -l "$node" 2>/dev/null || err "Failed to unmount $node"
    fi
  done
}

# Helper: Check if block device exists and is readable (used by retry_until)
check_device_ready() {
  local device="$1"
  validate_target_device "$device" "minimal"
}

# Helper: Wait for block device to become available with retry logic
# Usage: wait_for_device "/dev/sda1" [max_attempts]
wait_for_device() {
  local device="$1"
  local max_attempts="${2:-10}"
  retry_until "Waiting for $device" "$max_attempts" 1 check_device_ready "$device"
}

# Helper: Get partition name based on device type
# Usage: get_partition_name "/dev/sda" 1
get_partition_name() {
  local disk="$1"
  local partition_num="$2"
  
  if [[ "$disk" =~ (mmcblk|nvme|loop) ]]; then
    echo "${disk}p${partition_num}"
  else
    echo "${disk}${partition_num}"
  fi
}

# Helper: Check if PARTUUID is available (used by retry_until)
check_partuuid_available() {
  local device="$1"
  blkid -s PARTUUID -o value "$device" 2>/dev/null | grep -q .
}

# Helper: Safely get PARTUUID with retries
# Usage: get_partuuid_safe "/dev/sda1"
get_partuuid_safe() {
  local device="$1"
  local mode="${2:-$PARTUUID_DUP_MODE}"
  
  if retry_until "Getting PARTUUID for $device" 5 0.5 check_partuuid_available "$device"; then
    local partuuid
    partuuid=$(blkid -s PARTUUID -o value "$device" 2>/dev/null || true)
    if [[ -n "$partuuid" ]]; then
      handle_partuuid_uniqueness "$partuuid" "$device" "$mode" || return 1
      echo "$partuuid"
      return 0
    fi
    return 0
  fi
  return 1
}

# Helper: Safe mount with cleanup of stale mounts
# Usage: safe_mount "/dev/sda1" "/mnt/target"
safe_mount() {
  local device="$1"
  local mount_point="$2"
  
  # Clean up if mount point is already mounted
  if mountpoint -q "$mount_point" 2>/dev/null; then
    umount "$mount_point" 2>/dev/null || return 1
  fi
  
  # Ensure mount point exists and is empty
  if [[ -d "$mount_point" ]]; then
    if [[ "$(ls -A "$mount_point" 2>/dev/null)" ]]; then
      # Directory not empty and not mounted - shouldn't happen
      log_error "Mount point not empty: $mount_point"
      return 1
    fi
  else
    mkdir -p "$mount_point" || return 1
  fi
  
  # Attempt mount
  if ! mount "$device" "$mount_point" 2>/dev/null; then
    rmdir "$mount_point" 2>/dev/null || true
    return 1
  fi
  
  return 0
}

# Helper: Safe unmount with cleanup
# Usage: safe_umount "/mnt/target"
safe_umount() {
  local mount_point="$1"
  
  sync
  
  if mountpoint -q "$mount_point" 2>/dev/null; then
    if ! umount "$mount_point" 2>/dev/null; then
      # Try lazy unmount as fallback
      umount -l "$mount_point" 2>/dev/null || return 1
    fi
  fi
  
  # Remove directory if empty
  if [[ -d "$mount_point" ]] && [[ ! "$(ls -A "$mount_point" 2>/dev/null)" ]]; then
    rmdir "$mount_point" 2>/dev/null || true
  fi
  
  return 0
}

# Helper: Cleanup orphaned temporary mounts from failed runs
# Usage: cleanup_failed_mounts
cleanup_failed_mounts() {
  local cleaned=0
  
  for mnt in /mnt/*_temp_* /mnt/*_temp /mnt/temp_mount_*; do
    [[ -d "$mnt" ]] || continue
    
    if mountpoint -q "$mnt" 2>/dev/null; then
      log_success "Cleaning up orphaned mount: $mnt"
      umount -l "$mnt" 2>/dev/null || true
      cleaned=$((cleaned + 1))
    fi
    
    if [[ -d "$mnt" ]] && [[ ! "$(ls -A "$mnt" 2>/dev/null)" ]]; then
      rmdir "$mnt" 2>/dev/null || true
    fi
  done
  
  if (( cleaned > 0 )); then
    log_success "Cleaned up $cleaned orphaned mount(s)"
  fi
}

# Helper: Cleanup on error - unmount devices and remove partial configurations
# Usage: cleanup_on_error "device_path" ["description"]
cleanup_on_error() {
  local device="${1:-}"
  local description="${2:-operation}"
  
  log_info "Cleaning up after failed $description..."
  
  # Unmount any partitions from this device
  if [[ -n "$device" && -b "$device" ]]; then
    for part in "${device}"*; do
      [[ -b "$part" ]] || continue
      if mountpoint -q "$part" 2>/dev/null || grep -q "$part" /proc/mounts 2>/dev/null; then
        umount -l "$part" 2>/dev/null || true
      fi
    done
  fi
  
  # Clean up temp mounts
  cleanup_failed_mounts 2>/dev/null || true
  
  log_info "Cleanup complete. Device left in safe-to-retry state."
}

# Helper: Check available space before operations
# Usage: check_space_available "/target/path" required_mb
check_space_available() {
  local target_path="$1"
  local required_mb="$2"
  
  local available_kb=$(df --output=avail "$target_path" | tail -1)
  local available_mb=$((available_kb / 1024))
  
  if (( available_mb < required_mb )); then
    log_error "Insufficient space: ${available_mb}MB available, ${required_mb}MB required"
    return 1
  fi
  
  return 0
}

# Helper: Regenerate PARTUUIDs to avoid conflicts after cloning/partitioning
# Usage: regenerate_partuuids "/dev/sda"
regenerate_partuuids() {
  local device="$1"
  local part_table_type
  part_table_type=$(lsblk -ndo PTTYPE "$device" 2>/dev/null | head -n1 || true)

  # Raspberry Pi OS images are frequently MBR (dos) rather than GPT.
  # For MBR, randomize the disk identifier so derived PARTUUID values differ.
  if [[ "$part_table_type" == "dos" ]]; then
    if command -v sfdisk >/dev/null 2>&1; then
      local new_disk_id
      new_disk_id="0x$(od -An -N4 -tx4 /dev/urandom 2>/dev/null | tr -d ' \n')"

      if [[ -n "$new_disk_id" ]] && sfdisk --disk-id "$device" "$new_disk_id" >/dev/null 2>&1; then
        log_success "Updated MBR disk identifier on $device ($new_disk_id)"
        partprobe "$device" 2>/dev/null || true
        udevadm settle 2>/dev/null || true
        sleep 1
        return 0
      fi

      log_error "Failed to update MBR disk identifier on $device"
      return 1
    fi

    log_warning "sfdisk not available, skipping MBR disk identifier update"
    return 0
  fi
  
  if ! command -v sgdisk >/dev/null 2>&1; then
    log_warning "sgdisk not available, skipping GPT GUID regeneration"
    return 0
  fi
  
  log_success "Regenerating PARTUUIDs on $device to avoid conflicts..."
  if sgdisk --randomize-guids "$device" >/dev/null 2>&1; then
    partprobe "$device" 2>/dev/null || true
    udevadm settle 2>/dev/null || true
    sleep 1
    return 0
  else
    log_error "Failed to regenerate PARTUUIDs"
    return 1
  fi
}

# Helper: Handle PARTUUID uniqueness based on mode (fail|fix|warn)
# Usage: handle_partuuid_uniqueness "abc-123-def" "/dev/sda2" [mode]
handle_partuuid_uniqueness() {
  local partuuid="$1"
  local device="${2:-device}"
  local mode="${3:-$PARTUUID_DUP_MODE}"

  [[ -n "$partuuid" ]] || { log_error "PARTUUID is empty"; return 1; }

  # Force fresh probe to avoid stale cache after partition operations
  partprobe 2>/dev/null || true
  udevadm settle 2>/dev/null || true

  local count
  count=$(blkid -s PARTUUID | grep -c "PARTUUID=\"$partuuid\"" || echo 0)

  if (( count <= 1 )); then
    return 0
  fi

  case "$mode" in
    fix)
      ensure_partuuid_unique_or_fix "$partuuid" "$device"
      ;;
    warn)
      log_warning "PARTUUID conflict detected (non-fatal): $partuuid"
      log_warning "Devices with this PARTUUID:"
      blkid -s PARTUUID | grep "$partuuid" | sed 's/^/  /' >&2 || true
      return 1
      ;;
    *)
      log_error "ERROR: PARTUUID $partuuid appears on $count devices"
      log_error "This will cause boot conflicts. Run: sudo blkid | grep $partuuid"
      ERROR_HINT="Regenerate PARTUUIDs with: sudo sgdisk --randomize-guids /dev/sdX"
      log_error "Devices with this PARTUUID:"
      blkid -s PARTUUID | grep "$partuuid" | sed 's/^/  /' >&2 || true
      err --debug "PARTUUID conflict on $device (PARTUUID=$partuuid) - cannot proceed safely"
      ;;
  esac
}

# Helper: Build a PARTUUID map file (old -> new) for a disk
# Usage: create_partuuid_map_file "/tmp/pre.map" "/dev/sda" "/tmp/map"
create_partuuid_map_file() {
  local pre_map_file="$1"
  local disk="$2"
  local map_file="$3"

  declare -A old_by_part
  while read -r part old_uuid; do
    [[ -n "$part" && -n "$old_uuid" ]] && old_by_part["$part"]="$old_uuid"
  done < "$pre_map_file"

  declare -A new_by_part
  while read -r part new_uuid; do
    [[ -n "$part" && -n "$new_uuid" ]] && new_by_part["$part"]="$new_uuid"
  done < <(lsblk -nrpo NAME,PARTUUID "$disk" 2>/dev/null || true)

  : > "$map_file"
  for part in "${!old_by_part[@]}"; do
    if [[ -n "${new_by_part[$part]:-}" ]]; then
      echo "${old_by_part[$part]} ${new_by_part[$part]}" >> "$map_file"
    fi
  done
}

# Helper: Update cmdline.txt root PARTUUID if it is mapped
# Usage: update_cmdline_root_from_map "/mnt/boot" "/dev/sda1" "/tmp/map"
update_cmdline_root_from_map() {
  local boot_mount="$1"
  local device="$2"
  local map_file="$3"
  local cmdline_path="$boot_mount/cmdline.txt"

  [[ -f "$cmdline_path" ]] || return 0
  [[ -s "$map_file" ]] || return 0

  local cmdline_root
  local cmdline_root_value
  local cmdline_partuuid
  local new_root_partuuid

  cmdline_root="$(grep -o 'root=[^ ]*' "$cmdline_path" || true)"
  cmdline_root_value="${cmdline_root#root=}"

  if [[ "$cmdline_root_value" == PARTUUID=* ]]; then
    cmdline_partuuid="${cmdline_root_value#PARTUUID=}"
    new_root_partuuid=$(awk -v old="$cmdline_partuuid" '$1==old {print $2; exit}' "$map_file")
    if [[ -n "$new_root_partuuid" ]]; then
      log_info "Updating cmdline.txt on ${boot_mount} to root=PARTUUID=${new_root_partuuid}"
      update_cmdline_root "$cmdline_path" "$new_root_partuuid" "PARTUUID" || return 1
    fi
  fi

  return 0
}

# Helper: Read required PARTUUID with configurable acquire and verify modes.
# Usage: read_required_partuuid_checked "/dev/sda1" "label" [acquire_mode] [verify_mode]
# Returns PARTUUID on stdout or non-zero on failure.
read_required_partuuid_checked() {
  local device="$1"
  local label="$2"
  local acquire_mode="${3:-$PARTUUID_DUP_MODE}"
  local verify_mode="${4:-}"

  local partuuid
  partuuid="$(get_partuuid_safe "$device" "$acquire_mode" 2>/dev/null || true)"
  if [[ -z "$partuuid" ]]; then
    log_error "Could not determine PARTUUID of ${label} (${device})"
    return 1
  fi

  if [[ -n "$verify_mode" ]]; then
    handle_partuuid_uniqueness "$partuuid" "$device" "$verify_mode" || return 1
  fi

  echo "$partuuid"
}

# Helper: Update fstab entries for PARTUUIDs (either direct or from a map file)
# Usage: update_fstab_partuuid_entries "/etc/fstab" "boot_partuuid" "root_partuuid" [map_file]
update_fstab_partuuid_entries() {
  local fstab_path="$1"
  local boot_partuuid="${2:-}"
  local root_partuuid="${3:-}"
  local map_file="${4:-}"

  [[ -f "$fstab_path" ]] || { log_error "fstab not found: $fstab_path"; return 1; }

  if [[ -n "$map_file" && -f "$map_file" ]]; then
    cp "$fstab_path" "${fstab_path}.bak" 2>/dev/null || true
    while read -r old new; do
      [[ -n "$old" && -n "$new" ]] || continue
      sed -i -E "s#PARTUUID=${old}#PARTUUID=${new}#g" "$fstab_path" 2>/dev/null || true
    done < "$map_file"
    return 0
  fi

  cp "$fstab_path" "${fstab_path}.bak" 2>/dev/null || return 1

  [[ -n "$boot_partuuid" && -n "$root_partuuid" ]] || {
    log_error "Missing PARTUUIDs (boot=$boot_partuuid, root=$root_partuuid)"
    return 1
  }

  if grep -q "^\s*[^#]*/boot/firmware\s*" "$fstab_path"; then
    sed -i -E "s#^[^#]*\s+(/boot/firmware\s+)#PARTUUID=$boot_partuuid  \1#" "$fstab_path"
  else
    echo "PARTUUID=$boot_partuuid  /boot/firmware  vfat  defaults  0  2" >> "$fstab_path"
  fi

  if grep -qE "^[^#]+\s+/\s+[a-z]" "$fstab_path" 2>/dev/null; then
    sed -i -E "s#^([^#]*\s+)(/\s+)#PARTUUID=$root_partuuid  \2#" "$fstab_path"
  else
    echo "PARTUUID=$root_partuuid  /  ext4  defaults,noatime  0  1" >> "$fstab_path"
  fi

  grep -q "PARTUUID=$boot_partuuid" "$fstab_path" && \
  grep -q "PARTUUID=$root_partuuid" "$fstab_path" || {
    log_error "fstab verification failed"
    return 1
  }

  return 0
}

# Helper: Update fstab on mounted root using a PARTUUID map
# Usage: update_fstab_from_map "/mnt/root" "/dev/sda2" "/tmp/map"
update_fstab_from_map() {
  local root_mount="$1"
  local device="$2"
  local map_file="$3"
  local fstab_path="$root_mount/etc/fstab"

  [[ -f "$fstab_path" ]] || return 1
  update_fstab_partuuid_entries "$fstab_path" "" "" "$map_file"
}

# Helper: Apply PARTUUID mapping updates for cmdline/fstab on a disk
# Usage: update_partuuid_map_for_disk "/dev/sda" "/tmp/pre.map"
update_partuuid_map_for_disk() {
  local disk="$1"
  local pre_map_file="$2"
  local strict_mode="${3:-false}"

  partprobe "$disk" 2>/dev/null || true
  udevadm settle 2>/dev/null || true
  sleep 1

  local map_file
  map_file=$(mktemp /tmp/partuuid-map-XXXXXX)
  create_partuuid_map_file "$pre_map_file" "$disk" "$map_file"

  if [[ ! -s "$map_file" ]]; then
    log_warning "No PARTUUID changes detected on ${disk} (map empty)"
    rm -f "$map_file" 2>/dev/null || true
    return 0
  fi

  local boot_updated=false
  while read -r part fstype; do
    [[ -z "$part" || -z "$fstype" ]] && continue
    case "$fstype" in
      vfat|fat|msdos)
        if with_mount "$part" update_cmdline_root_from_map --verify-boot "$map_file"; then
          boot_updated=true
        fi
        ;;
    esac
  done < <(lsblk -nrpo NAME,FSTYPE "$disk" 2>/dev/null || true)

  local root_updated=false
  while read -r part fstype; do
    [[ -z "$part" || -z "$fstype" ]] && continue
    case "$fstype" in
      vfat|fat|msdos|swap)
        continue
        ;;
    esac

    if with_mount "$part" update_fstab_from_map "$map_file"; then
      log_info "Updated fstab on root partition: ${part}"
      root_updated=true
      break
    fi
  done < <(lsblk -nrpo NAME,FSTYPE "$disk" 2>/dev/null || true)

  if [[ "$boot_updated" != "true" ]]; then
    log_warning "No boot cmdline.txt updated on ${disk} (none found or root=PARTUUID not present)"
  fi
  if [[ "$root_updated" != "true" ]]; then
    log_warning "No root fstab updated on ${disk} (no root filesystem found)"
  fi

  if [[ "$strict_mode" == "true" ]]; then
    if [[ "$boot_updated" != "true" || "$root_updated" != "true" ]]; then
      log_error "Strict PARTUUID map update failed on ${disk} (boot_updated=${boot_updated}, root_updated=${root_updated})"
      rm -f "$map_file" 2>/dev/null || true
      return 1
    fi
  fi

  rm -f "$map_file" 2>/dev/null || true
  return 0
}

# Helper: Ensure PARTUUID is unique system-wide, auto-resolving duplicates
# Usage: ensure_partuuid_unique_or_fix "abc-123-def" "/dev/sda2"
ensure_partuuid_unique_or_fix() {
  local partuuid="$1"
  local preferred_device="${2:-}"

  [[ -n "$partuuid" ]] || { log_error "PARTUUID is empty"; return 1; }

  local -a devices=()
  mapfile -t devices < <(blkid -t "PARTUUID=${partuuid}" -o device 2>/dev/null || true)

  if [[ ${#devices[@]} -le 1 ]]; then
    return 0
  fi

  if ! command -v sgdisk >/dev/null 2>&1; then
    log_error "PARTUUID conflict detected but sgdisk is not available"
    err --debug "Install sgdisk (gdisk package) to auto-resolve PARTUUID conflicts"
  fi

  local keep_device=""
  if [[ -n "$preferred_device" ]]; then
    local resolved_pref
    resolved_pref=$(readlink -f "$preferred_device" 2>/dev/null || echo "$preferred_device")
    for dev in "${devices[@]}"; do
      if [[ "$(readlink -f "$dev" 2>/dev/null || echo "$dev")" == "$resolved_pref" ]]; then
        keep_device="$dev"
        break
      fi
    done
  fi

  if [[ -z "$keep_device" && -n "${CURRENT_ROOT:-}" ]]; then
    local resolved_root
    resolved_root=$(readlink -f "$CURRENT_ROOT" 2>/dev/null || echo "$CURRENT_ROOT")
    for dev in "${devices[@]}"; do
      if [[ "$(readlink -f "$dev" 2>/dev/null || echo "$dev")" == "$resolved_root" ]]; then
        keep_device="$dev"
        break
      fi
    done
  fi

  if [[ -z "$keep_device" ]]; then
    keep_device="${devices[0]}"
  fi

  log_warning "Duplicate PARTUUID detected: ${partuuid}"
  log_warning "Keeping: ${keep_device}"

  if [[ -z "$preferred_device" ]]; then
    log_error "Refusing automatic fix without explicit target device"
    err --debug "Attach only target media or provide explicit preferred target partition for conflict repair"
  fi

  local resolved_preferred
  resolved_preferred=$(readlink -f "$preferred_device" 2>/dev/null || echo "$preferred_device")
  local preferred_in_conflict=false
  for dev in "${devices[@]}"; do
    if [[ "$(readlink -f "$dev" 2>/dev/null || echo "$dev")" == "$resolved_preferred" ]]; then
      preferred_in_conflict=true
      break
    fi
  done
  if [[ "$preferred_in_conflict" != "true" ]]; then
    log_error "Preferred device ${preferred_device} does not match conflicted PARTUUID entries"
    err --debug "Refusing to modify non-target disks; detach unrelated drives and retry"
  fi

  local preferred_disk_name
  preferred_disk_name=$(lsblk -no PKNAME "$preferred_device" 2>/dev/null | head -n1)
  if [[ -z "$preferred_disk_name" ]]; then
    preferred_disk_name=$(basename "$preferred_device")
  fi
  local preferred_disk="/dev/${preferred_disk_name}"

  local current_root_dev
  current_root_dev=$(findmnt -no SOURCE / 2>/dev/null || true)
  local current_root_disk=""
  if [[ -n "$current_root_dev" ]]; then
    current_root_disk=$(lsblk -no PKNAME "$current_root_dev" 2>/dev/null | head -n1 || true)
  fi

  local current_boot_dev
  current_boot_dev=$(findmnt -no SOURCE "$BOOT_DIR" 2>/dev/null || true)
  local current_boot_disk=""
  if [[ -n "$current_boot_dev" ]]; then
    current_boot_disk=$(lsblk -no PKNAME "$current_boot_dev" 2>/dev/null | head -n1 || true)
  fi

  if [[ -n "$current_root_disk" && "$preferred_disk_name" == "$current_root_disk" ]]; then
    log_error "Refusing to randomize current root disk: ${preferred_disk}"
    err --debug "Conflict requires manual resolution; do not modify the running root disk"
  fi
  if [[ -n "$current_boot_disk" && "$preferred_disk_name" == "$current_boot_disk" ]]; then
    log_error "Refusing to randomize current boot disk: ${preferred_disk}"
    err --debug "Conflict requires manual resolution; do not modify the current boot disk"
  fi

  local pre_map_file
  pre_map_file=$(mktemp /tmp/partuuid-pre-XXXXXX)
  lsblk -nrpo NAME,PARTUUID "$preferred_disk" 2>/dev/null > "$pre_map_file" || true

  log_warning "Regenerating GUIDs only on target disk: ${preferred_disk}"
  if sgdisk --randomize-guids "$preferred_disk" >/dev/null 2>&1; then
    if ! update_partuuid_map_for_disk "$preferred_disk" "$pre_map_file" "true"; then
      rm -f "$pre_map_file" 2>/dev/null || true
      err --debug "Randomized GUIDs on ${preferred_disk} but failed strict cmdline/fstab remap"
    fi
    log_success "PARTUUID conflict repair applied to target disk only: ${preferred_disk}"
  else
    rm -f "$pre_map_file" 2>/dev/null || true
    log_error "Failed to regenerate GUIDs on ${preferred_disk}"
    err --debug "Unable to resolve PARTUUID conflict on target disk"
  fi

  rm -f "$pre_map_file" 2>/dev/null || true

  local count
  count=$(blkid -s PARTUUID | grep -c "PARTUUID=\"${partuuid}\"" || echo 0)
  if (( count > 1 )); then
    log_error "PARTUUID conflict persists after target-only fix (PARTUUID=${partuuid})"
    log_error "Unmodified conflicting devices still present. Detach non-target drives and retry."
    err --debug "Target-only PARTUUID fix incomplete due to remaining duplicates on other disks"
  fi

  return 0
}

# Helper: Ensure boot partition is properly configured (power settings + scripts)
# Usage: ensure_boot_configured "/mnt/boot" "USB device"
# Returns: 0=success, 1=error
ensure_boot_configured() {
  local boot_mount="$1"
  local target_name="$2"
  
  # Check writable
  if ! check_writable "$boot_mount"; then
    log_error "Cannot configure $target_name - boot partition is read-only"
    return 1
  fi
  
  # Add USB power settings
  if [[ -f "$boot_mount/config.txt" ]]; then
    if add_usb_power_settings "$boot_mount/config.txt"; then
      log_success "USB power settings configured on $target_name"
    else
      log_warning "Failed to configure USB power settings on $target_name"
    fi
  else
    log_warning "config.txt not found on $target_name"
  fi
  
  # Copy scripts
  copy_scripts_to_boot "$boot_mount" "$target_name"
  
  sync
  return 0
}

# Helper: Prepare boot partition for a specific root identifier
# Usage: prepare_boot_partition_for_root "/mnt/boot" "label" "root_partuuid" [cmdline_path] [root_id_type]
prepare_boot_partition_for_root() {
  local boot_mount="$1"
  local target_label="$2"
  local root_partuuid="$3"
  local cmdline_path="${4:-$boot_mount/cmdline.txt}"
  local root_id_type="${5:-PARTUUID}"

  [[ -n "$root_partuuid" ]] || { log_error "Root PARTUUID missing - cannot update cmdline.txt"; return 1; }
  [[ -f "$cmdline_path" ]] || { log_error "cmdline.txt not found: $cmdline_path"; return 1; }

  if ! update_cmdline_root "$cmdline_path" "$root_partuuid" "$root_id_type"; then
    return 1
  fi

  if ! ensure_boot_configured "$boot_mount" "$target_label"; then
    log_warning "Could not fully configure boot partition on ${target_label}"
  fi

  return 0
}

# Helper: Configure boot partition for a target root (optional boot copy)
# Usage: configure_boot_partition_for_target "/dev/sda1" "root_partuuid" "label" "copy_boot"
configure_boot_partition_for_target() {
  local boot_dev="$1"
  local root_partuuid="$2"
  local target_label="$3"
  local copy_boot="${4:-false}"

  [[ -n "$root_partuuid" ]] || { log_error "Root PARTUUID missing - cannot update cmdline.txt"; return 1; }

  configure_boot_partition() {
    local boot_mount="$1"

    if ! check_writable "$boot_mount"; then
      return 1
    fi

    if [[ "$copy_boot" == "true" ]]; then
      log_success "Copying boot files (this may take 1-2 minutes)..."
      local fstype
      fstype=$(findmnt -n -o FSTYPE "$boot_mount" 2>/dev/null || echo "")
      local rsync_cmd=(rsync -aH -x --info=progress2)
      if [[ "$fstype" =~ ^(vfat|fat|msdos)$ ]]; then
        rsync_cmd+=(--no-perms --no-owner --no-group --no-acls --no-xattrs)
      fi
      if ! "${rsync_cmd[@]}" "$BOOT_DIR/" "$boot_mount/"; then
        log_error "Boot file copy failed"
        return 1
      fi
      sync

      if ! verify_filesystem "$boot_mount" "boot"; then
        log_error "Boot file verification failed - copy incomplete"
        return 1
      fi
      log_success "Boot files verified"
    else
      if ! verify_filesystem "$boot_mount" "boot"; then
        log_error "Boot file verification failed"
        return 1
      fi
    fi

    prepare_boot_partition_for_root "$boot_mount" "$target_label" "$root_partuuid" "$boot_mount/cmdline.txt"
  }

  if ! with_mount "$boot_dev" configure_boot_partition; then
    return 1
  fi

  return 0
}

# Helper: Update cmdline.txt with new root identifier (with rollback and validation)
# Usage: update_cmdline_root "/boot/firmware/cmdline.txt" "abc-123-def" [PARTUUID|UUID]
update_cmdline_root() {
  local target_cmdline="$1"
  local root_id="$2"
  local root_id_type="${3:-PARTUUID}"
  
  [[ -f "$target_cmdline" ]] || { log_error "cmdline.txt not found: $target_cmdline"; return 1; }
  [[ -n "$root_id" ]] || { log_error "Root identifier is empty"; return 1; }
  root_id_type="${root_id_type^^}"
  if [[ "$root_id_type" != "PARTUUID" && "$root_id_type" != "UUID" ]]; then
    log_error "Invalid root identifier type: $root_id_type"
    return 1
  fi
  
  # Create temporary backup
  local temp_backup="${target_cmdline}.tmp_backup_$$"
  if ! cp "$target_cmdline" "$temp_backup"; then
    log_error "Failed to create temporary backup"
    return 1
  fi
  
  # Perform modification
  if ! sed -i -E "s#root=[^ ]+#root=${root_id_type}=${root_id}#" "$target_cmdline"; then
    cp "$temp_backup" "$target_cmdline" 2>/dev/null
    rm "$temp_backup" 2>/dev/null
    log_error "sed command failed"
    return 1
  fi

  # Enforce rootwait for USB/media timing reliability during early boot.
  if ! grep -qE '(^| )rootwait( |$)' "$target_cmdline"; then
    if ! sed -i -E 's#$# rootwait#' "$target_cmdline"; then
      cp "$temp_backup" "$target_cmdline" 2>/dev/null
      rm "$temp_backup" 2>/dev/null
      log_error "Failed to append rootwait to cmdline.txt"
      return 1
    fi
  fi
  
  # Validate: must be single line
  local line_count=$(wc -l < "$target_cmdline" 2>/dev/null || echo 999)
  if (( line_count > 1 )); then
    cp "$temp_backup" "$target_cmdline"
    rm "$temp_backup" 2>/dev/null
    log_error "Validation failed: cmdline.txt must be a single line (found $line_count lines)"
    return 1
  fi
  
  # Validate: PARTUUID must be present
  if ! grep -q "root=${root_id_type}=${root_id}" "$target_cmdline"; then
    cp "$temp_backup" "$target_cmdline"
    rm "$temp_backup" 2>/dev/null
    log_error "Validation failed: root=${root_id_type} not found in updated cmdline.txt"
    return 1
  fi
  
  # Success - remove temp backup and sync
  rm "$temp_backup" 2>/dev/null
  sync
  
  log_success "Updated cmdline.txt with root=${root_id_type}=${root_id}"
  return 0
}

# Helper: Print planned cmdline/fstab changes
# Usage: print_change_plan "context" "cmdline_path" "root_id" "boot_partuuid" "fstab_path" [root_id_type]
print_change_plan() {
  local context="$1"
  local cmdline_path="$2"
  local root_id="$3"
  local boot_partuuid="$4"
  local fstab_path="$5"
  local root_id_type="${6:-PARTUUID}"

  root_id_type="${root_id_type^^}"
  log_info "Planned changes (${context}):"
  if [[ -n "$cmdline_path" && -n "$root_id" ]]; then
    log_info "  cmdline: root=${root_id_type}=${root_id} (${cmdline_path})"
  fi
  if [[ -n "$fstab_path" ]]; then
    if [[ -n "$boot_partuuid" ]]; then
      log_info "  fstab: /boot/firmware -> PARTUUID=${boot_partuuid} (${fstab_path})"
    fi
    if [[ -n "$root_id" ]]; then
      log_info "  fstab: / -> PARTUUID=${root_id} (${fstab_path})"
    fi
  fi
}

# Helper: Ensure backup exists before modifying cmdline.txt
# Call this before any operation that modifies cmdline.txt
ensure_cmdline_backup() {
  if [[ ! -f "$CMDLINE" ]]; then
    err "$CMDLINE not found"
  fi
  
  if [[ ! -f "$BACKUP" ]]; then
    cp "$CMDLINE" "$BACKUP"
    log_success "Created backup: $BACKUP"
  else
    if cmp -s "$CMDLINE" "$BACKUP"; then
      log_success "Backup is current (no changes needed)"
    else
      echo "A backup already exists at $BACKUP"
      echo "Current cmdline.txt has changed since that backup was created."
      if prompt_yn "Replace existing backup with current cmdline.txt state?"; then
        cp "$CMDLINE" "$BACKUP"
        log_success "Updated backup: $BACKUP"
      else
        log_success "Using existing backup (no changes made)"
      fi
    fi
  fi
}

update_fstab() {
  local root_mount_point="$1"
  local boot_partuuid="$2"  # Accept boot PARTUUID as parameter
  local root_partuuid="$3"  # Accept root PARTUUID as parameter
  local fstab_path="${root_mount_point}/etc/fstab"

  if ! update_fstab_partuuid_entries "$fstab_path" "$boot_partuuid" "$root_partuuid"; then
    return 1
  fi

  sync
  log_success "Updated fstab: boot=PARTUUID=$boot_partuuid, root=PARTUUID=$root_partuuid"
  return 0
}

################################################################################
# SYSTEM CONFIGURATION CHECKS (SSH/WIFI)
################################################################################

# Helper: Collect sshd_config files in load order (main + drop-ins + Include)
collect_sshd_config_files() {
  local main_file="$1"
  local dropin_dir="$2"
  local -a files=()
  local -A seen=()

  add_file() {
    local f="$1"
    [[ -f "$f" ]] || return 0
    if [[ -z "${seen[$f]:-}" ]]; then
      files+=("$f")
      seen[$f]=1
    fi
  }

  # Seed with main and drop-ins
  add_file "$main_file"
  if [[ -d "$dropin_dir" ]]; then
    mapfile -t dropins < <(ls "$dropin_dir"/*.conf 2>/dev/null | sort)
    for f in "${dropins[@]}"; do
      add_file "$f"
    done
  fi

  # Follow Include directives (one-pass expansion with recursion guard)
  local i=0
  while (( i < ${#files[@]} )); do
    local f="${files[$i]}"
    mapfile -t includes < <(
      awk '
        BEGIN { IGNORECASE=1 }
        { line=$0; sub(/#.*/, "", line); gsub(/^[ \t]+|[ \t]+$/, "", line) }
        tolower(line) ~ /^include[ \t]+/ {
          sub(/^[^ \t]+[ \t]+/, "", line)
          if (line != "") print line
        }
      ' "$f" 2>/dev/null
    )

    for inc in "${includes[@]}"; do
      # Resolve relative paths per sshd_config convention (/etc/ssh)
      local resolved="$inc"
      if [[ "$resolved" != /* ]]; then
        resolved="/etc/ssh/$resolved"
      fi

      # Expand globs
      mapfile -t expanded < <(compgen -G "$resolved" 2>/dev/null || true)
      if [[ ${#expanded[@]} -eq 0 && -f "$resolved" ]]; then
        expanded=("$resolved")
      fi
      for ef in "${expanded[@]}"; do
        add_file "$ef"
      done
    done

    i=$((i + 1))
  done

  printf '%s\n' "${files[@]}"
}

# Helper: Detect Match blocks (settings may vary per connection/user)
sshd_config_has_match() {
  local main_file="$1"
  local dropin_dir="$2"
  mapfile -t config_files < <(collect_sshd_config_files "$main_file" "$dropin_dir")
  if [[ ${#config_files[@]} -eq 0 ]]; then
    return 1
  fi
  grep -i -qE '^[[:space:]]*Match[[:space:]]+' "${config_files[@]}" 2>/dev/null
}

# Helper: Get effective sshd_config setting value from main + drop-ins
# Usage: get_effective_sshd_setting "/etc/ssh/sshd_config" "/etc/ssh/sshd_config.d" "PasswordAuthentication"
get_effective_sshd_setting() {
  local main_file="$1"
  local dropin_dir="$2"
  local key="$3"

  mapfile -t config_files < <(collect_sshd_config_files "$main_file" "$dropin_dir")
  if [[ ${#config_files[@]} -eq 0 ]]; then
    echo ""
    return 0
  fi

  awk -v key="$key" '
    BEGIN { IGNORECASE=1 }
    {
      line=$0
      sub(/#.*/, "", line)
      gsub(/^[ \t]+|[ \t]+$/, "", line)
      if (line == "") next
      if (tolower(line) ~ "^" tolower(key) "[ \t]+") {
        sub(/^[^ \t]+[ \t]+/, "", line)
        val=line
      }
    }
    END { if (val != "") print val }
  ' "${config_files[@]}"
}

# Helper: Validate SSH auth/access policy using effective sshd settings
# Usage: evaluate_sshd_access_policy "/etc/ssh/sshd_config" "/etc/ssh/sshd_config.d" "username" "group1 group2" [context_prefix]
SSHD_POLICY_ISSUES=()
evaluate_sshd_access_policy() {
  local ssh_config_main="$1"
  local ssh_config_d="$2"
  local target_user="${3:-}"
  local target_groups="${4:-}"
  local context_prefix="${5:-}"

  SSHD_POLICY_ISSUES=()

  local effective_password_auth
  local effective_kbd_auth
  local effective_challenge_auth
  local effective_use_pam
  local effective_auth_methods
  local effective_allow_users
  local effective_deny_users
  local effective_allow_groups
  local effective_deny_groups
  local effective_permit_root

  if sshd_config_has_match "$ssh_config_main" "$ssh_config_d"; then
    log_warning "${context_prefix}SSH config contains Match blocks - settings may be conditional"
  fi

  effective_password_auth=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "PasswordAuthentication")
  effective_kbd_auth=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "KbdInteractiveAuthentication")
  effective_challenge_auth=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "ChallengeResponseAuthentication")
  effective_use_pam=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "UsePAM")
  effective_auth_methods=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "AuthenticationMethods")
  effective_allow_users=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "AllowUsers")
  effective_deny_users=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "DenyUsers")
  effective_allow_groups=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "AllowGroups")
  effective_deny_groups=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "DenyGroups")
  effective_permit_root=$(get_effective_sshd_setting "$ssh_config_main" "$ssh_config_d" "PermitRootLogin")

  if [[ -n "$effective_auth_methods" ]]; then
    if ! grep -qiE '(password|keyboard-interactive)' <<< "$effective_auth_methods"; then
      SSHD_POLICY_ISSUES+=("${context_prefix}SSH AuthenticationMethods does not allow password or keyboard-interactive authentication")
    fi
  fi

  if [[ "${effective_password_auth,,}" == "no" ]]; then
    SSHD_POLICY_ISSUES+=("${context_prefix}SSH config disables PasswordAuthentication")
  fi

  if [[ "${effective_password_auth,,}" == "no" && "${effective_kbd_auth,,}" == "no" && "${effective_challenge_auth,,}" == "no" ]]; then
    SSHD_POLICY_ISSUES+=("${context_prefix}SSH config disables all password-based auth methods")
  fi

  if [[ "${effective_use_pam,,}" == "no" && ( "${effective_password_auth,,}" == "yes" || "${effective_kbd_auth,,}" == "yes" || "${effective_challenge_auth,,}" == "yes" ) ]]; then
    log_warning "${context_prefix}UsePAM is disabled; password authentication may fail"
  fi

  if [[ -n "$target_user" ]]; then
    if [[ -n "$effective_allow_users" && ! " $effective_allow_users " =~ [[:space:]]${target_user}[[:space:]] ]]; then
      SSHD_POLICY_ISSUES+=("${context_prefix}SSH config AllowUsers does not include ${target_user}")
    fi

    if [[ -n "$effective_deny_users" && " $effective_deny_users " =~ [[:space:]]${target_user}[[:space:]] ]]; then
      SSHD_POLICY_ISSUES+=("${context_prefix}SSH config DenyUsers includes ${target_user}")
    fi

    if [[ -n "$effective_allow_groups" ]]; then
      local allowed=false
      local grp
      for grp in $effective_allow_groups; do
        if [[ " $target_groups " =~ [[:space:]]${grp}[[:space:]] ]]; then
          allowed=true
          break
        fi
      done
      if [[ "$allowed" != "true" ]]; then
        SSHD_POLICY_ISSUES+=("${context_prefix}SSH config AllowGroups does not include ${target_user}")
      fi
    fi

    if [[ -n "$effective_deny_groups" ]]; then
      local grp
      for grp in $effective_deny_groups; do
        if [[ " $target_groups " =~ [[:space:]]${grp}[[:space:]] ]]; then
          SSHD_POLICY_ISSUES+=("${context_prefix}SSH config DenyGroups includes ${target_user} (group: ${grp})")
          break
        fi
      done
    fi
  fi

  if [[ -n "$effective_permit_root" && "${effective_permit_root,,}" == "yes" ]]; then
    log_warning "${context_prefix}SSH config allows root login (PermitRootLogin yes)"
  fi

  (( ${#SSHD_POLICY_ISSUES[@]} == 0 ))
}

# Check if SSH is properly configured with default credentials
check_ssh_config() {
  local errors=()
  local target_user=""
  local ssh_config_main="/etc/ssh/sshd_config"
  local ssh_config_d="/etc/ssh/sshd_config.d"
    
    echo
    log_success "Checking SSH configuration..."

    # Determine the user to validate (prefer sudo-invoking user)
    target_user=$(detect_invoking_user)
    if [[ -z "$target_user" ]]; then
      log_warning "Could not determine a non-root user; user-specific checks will be skipped"
      target_user=""
    fi
    
    # Check if SSH service is enabled and running
    if ! systemctl is-enabled ssh >/dev/null 2>&1 && ! systemctl is-enabled sshd >/dev/null 2>&1; then
        errors+=("SSH service is not enabled")
    fi
    
    if ! systemctl is-active ssh >/dev/null 2>&1 && ! systemctl is-active sshd >/dev/null 2>&1; then
        errors+=("SSH service is not running")
    fi
    
    # Check if target user exists and has sudo access
    if [[ -n "$target_user" ]]; then
      if ! id "$target_user" &>/dev/null; then
        errors+=("User '$target_user' does not exist")
      else
        # Check if target user is in sudo group
        if ! groups "$target_user" | grep -q sudo; then
          errors+=("User '$target_user' is not in sudo group")
        fi
            
        # Verify password is set (check shadow file for non-locked account)
        if [[ -f /etc/shadow ]]; then
          local user_shadow=$(grep "^${target_user}:" /etc/shadow 2>/dev/null)
          if [[ -z "$user_shadow" ]]; then
            errors+=("User '$target_user' has no password entry in shadow file")
          elif [[ "$user_shadow" =~ ^${target_user}:[*!]: ]]; then
            errors+=("User '$target_user' account is locked or has no password")
          fi
        fi
      fi
    fi
    
    # Check SSH configuration allows password authentication (main + drop-ins + Include)
    if [[ -f "$ssh_config_main" || -d "$ssh_config_d" ]]; then
      local user_groups=""
      if [[ -n "$target_user" ]]; then
        user_groups=$(id -nG "$target_user" 2>/dev/null || echo "")
      fi

      if ! evaluate_sshd_access_policy "$ssh_config_main" "$ssh_config_d" "$target_user" "$user_groups"; then
        errors+=("${SSHD_POLICY_ISSUES[@]}")
      fi
    fi
    
    # Report results
    if [[ ${#errors[@]} -eq 0 ]]; then
        log_success "SSH configuration check: PASSED"
        log_success "  ✓ SSH service is enabled and running"
        if [[ -n "$target_user" ]]; then
          log_success "  ✓ User '$target_user' exists with sudo access"
        else
          log_success "  ✓ User check skipped (no non-root user detected)"
        fi
        log_success "  ✓ Password authentication is allowed"
        return 0
    else
        echo
        log_error "SSH configuration check: FAILED"
        for error in "${errors[@]}"; do
            echo "  ✗ $error"
        done
        echo
        log_info "Expected SSH configuration:"
        if [[ -n "$target_user" ]]; then
          log_info "  - Username: $target_user"
          log_info "  - User is in sudo group"
        else
          log_info "  - A non-root sudo user should exist"
        fi
        log_info "  - SSH service: enabled and running"
        log_info "  - Password authentication: enabled"
        return 1
    fi
}

# Check if WiFi is properly configured
check_wifi_config() {
  local errors=()
  local warnings=()
    
  echo
  log_success "Checking WiFi configuration..."
    
  # Check if WiFi tools exist
  if ! command -v iwconfig >/dev/null 2>&1; then
    warnings+=("iwconfig not installed (wireless-tools package)")
  fi
    
  # Detect ethernet connectivity (default route interface, non-wlan)
  local default_iface=$(ip route | awk '/default/ {print $5; exit}')
  local ethernet_up=false
  if [[ -n "$default_iface" && ! "$default_iface" =~ ^wlan ]]; then
    if ip -4 addr show "$default_iface" | grep -q "inet "; then
      ethernet_up=true
      log_success "  ✓ Ethernet detected on ${default_iface}"
    fi
  fi
    
  # Detect WiFi interface
  local wifi_iface=$(ls /sys/class/net/ | grep -m1 '^wlan' || echo "")
  if [[ -z "$wifi_iface" ]]; then
    if [[ "$ethernet_up" == "true" ]]; then
      warnings+=("No WiFi interface detected (Ethernet is active)")
    else
      errors+=("No WiFi interface detected and no active Ethernet connection")
    fi
  else
    log_success "  Found WiFi interface: $wifi_iface"
        
    # Check current connection (optional)
    if command -v iwgetid >/dev/null 2>&1; then
      local current_ssid=$(iwgetid -r 2>/dev/null || echo "")
      if [[ "$current_ssid" == "$WIFI_SSID" ]]; then
        log_success "  ✓ Connected to ${WIFI_SSID}"
      elif [[ -n "$current_ssid" ]]; then
        warnings+=("Connected to different network: $current_ssid (expected: ${WIFI_SSID})")
      else
        warnings+=("Not currently connected to any WiFi network")
      fi
    fi
        
    # Check wpa_supplicant configuration
    local wifi_configured=false
    local nm_configured=false
    local nm_active_match=false
    if [[ -f /etc/wpa_supplicant/wpa_supplicant.conf ]]; then
      if grep -q "ssid=\"${WIFI_SSID}\"" /etc/wpa_supplicant/wpa_supplicant.conf 2>/dev/null; then
        log_success "  ✓ ${WIFI_SSID} found in wpa_supplicant.conf"
                
        # Verify password is configured (we can't see the actual PSK easily)
        if grep -A2 "ssid=\"${WIFI_SSID}\"" /etc/wpa_supplicant/wpa_supplicant.conf 2>/dev/null | grep -q 'psk='; then
          log_success "  ✓ Password configured for ${WIFI_SSID}"
          wifi_configured=true
        else
          warnings+=("${WIFI_SSID} configured but password (psk) not found")
        fi
      else
        warnings+=("${WIFI_SSID} not found in /etc/wpa_supplicant/wpa_supplicant.conf")
      fi
    else
      if [[ "$nm_active_match" != "true" && "$nm_configured" != "true" ]]; then
        warnings+=("/etc/wpa_supplicant/wpa_supplicant.conf not found")
      fi
    fi

    # Check NetworkManager configuration (common on Raspberry Pi OS Desktop)
    if [[ -d /etc/NetworkManager/system-connections ]]; then
      if grep -rl "^ssid=${WIFI_SSID}$" /etc/NetworkManager/system-connections 2>/dev/null | grep -q .; then
        log_success "  ✓ ${WIFI_SSID} found in NetworkManager profiles"
        nm_configured=true
      elif grep -rl "ssid=${WIFI_SSID}" /etc/NetworkManager/system-connections 2>/dev/null | grep -q .; then
        log_success "  ✓ ${WIFI_SSID} found in NetworkManager profiles"
        nm_configured=true
      else
        if [[ "$nm_active_match" != "true" ]]; then
          warnings+=("${WIFI_SSID} not found in /etc/NetworkManager/system-connections")
        fi
      fi
    else
      if [[ "$nm_active_match" != "true" && "$nm_configured" != "true" ]]; then
        warnings+=("/etc/NetworkManager/system-connections not found")
      fi
    fi

    # Check active NetworkManager WiFi connection
    if command -v nmcli >/dev/null 2>&1; then
      local nm_active_ssid
      nm_active_ssid=$(nmcli -t -f active,ssid dev wifi 2>/dev/null | awk -F: '$1=="yes" {print $2; exit}')
      if [[ "$nm_active_ssid" == "$WIFI_SSID" ]]; then
        log_success "  ✓ NetworkManager active on ${WIFI_SSID}"
        nm_active_match=true
        nm_configured=true
        wifi_configured=true
      elif [[ -n "$nm_active_ssid" ]]; then
        warnings+=("NetworkManager active on different SSID: $nm_active_ssid (expected: ${WIFI_SSID})")
      fi
    fi

    # If NetworkManager is actively connected, suppress static-profile warnings.
    if [[ "$nm_active_match" == "true" ]]; then
      local filtered_warnings=()
      local warning
      for warning in "${warnings[@]}"; do
        case "$warning" in
          "/etc/wpa_supplicant/wpa_supplicant.conf not found"|"${WIFI_SSID} not found in /etc/NetworkManager/system-connections"|"/etc/NetworkManager/system-connections not found")
            ;;
          *)
            filtered_warnings+=("$warning")
            ;;
        esac
      done
      warnings=("${filtered_warnings[@]}")
    fi

    if [[ "$nm_configured" == "true" ]]; then
      wifi_configured=true
    fi
        
    if [[ "$wifi_configured" != "true" && "$ethernet_up" != "true" ]]; then
      errors+=("WiFi not configured and no active Ethernet connection")
    fi
  fi
    
  # Report results
  echo
  if [[ ${#errors[@]} -eq 0 ]]; then
    log_success "WiFi configuration check: PASSED"
  else
    log_error "WiFi configuration check: FAILED"
    for error in "${errors[@]}"; do
      echo "  ✗ $error"
    done
  fi
    
  if [[ ${#warnings[@]} -gt 0 ]]; then
    echo
    echo "$(timestamp) [~] WiFi warnings:"
    for warning in "${warnings[@]}"; do
      echo "  ⚠ $warning"
    done
  fi
    
  if [[ ${#errors[@]} -gt 0 && "$ethernet_up" == "true" ]]; then
    log_warning "Ethernet is active; continuing despite WiFi issues"
    return 0
  fi
    
  if [[ ${#errors[@]} -gt 0 ]]; then
    echo
    log_info "Expected WiFi configuration:"
    log_info "  - Network name (SSID): ${WIFI_SSID}"
    log_info "  - Password: ${WIFI_PSK}"
    log_info "  - Configuration file (wpa_supplicant): /etc/wpa_supplicant/wpa_supplicant.conf"
    log_info "  - Configuration directory (NetworkManager): /etc/NetworkManager/system-connections"
    return 1
  fi
    
  return 0
}

# Check initial Pi configuration (SSH and WiFi)
check_boot_target_readiness() {
    echo
    log_success "Checking configured boot targets (read-only)..."
  log_info "Mode 2 target check includes boot readiness diagnostics (cmdline/fstab/root mapping) and SSH continuity"

    if ! eeprom_tool_available; then
      log_warning "EEPROM tools not available - skipping target boot readiness check"
      return 0
    fi

    local current_order
    current_order=$(get_eeprom_boot_order)
    if [[ -z "$current_order" ]]; then
      log_warning "Could not read EEPROM BOOT_ORDER - skipping target boot readiness check"
      return 0
    fi

    if check_boot_order_ssh_continuity "$current_order"; then
      if (( BOOT_ORDER_LAST_CHECKED == 0 )); then
        log_warning "Boot target readiness is inconclusive: 0 devices validated"
        log_info "Current system SSH/WiFi may be healthy, but the next boot target was not verified"
        log_info "To verify a specific target now, attach it and run option 8 or option 7/6 followed by option 2"
        log_warning "Boot target readiness check: INCONCLUSIVE"
        return 0
      fi
      if [[ -n "$BOOT_ORDER_NEXT_TARGET_BOOT" && -n "$BOOT_ORDER_NEXT_TARGET_ROOT" ]]; then
        log_success "Next non-current boot candidate validated: ${BOOT_ORDER_NEXT_TARGET_BOOT} -> ${BOOT_ORDER_NEXT_TARGET_ROOT}"
      else
        log_info "No validated non-current boot candidate found in BOOT_ORDER (current device remains primary)"
      fi
      log_success "Boot target readiness check: PASSED"
      return 0
    fi

    log_error "Boot target readiness check: FAILED"
    return 1
}

# Check initial Pi configuration (SSH and WiFi)
check_initial_config() {
    local include_boot_targets=false
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --include-boot-targets)
          include_boot_targets=true
          ;;
      esac
      shift
    done

    echo
    print_separator
    echo "Initial Configuration Verification"
    print_separator
    
    local ssh_status=0
    local wifi_status=0
    local target_status=0
    
    check_ssh_config || ssh_status=$?
    check_wifi_config || wifi_status=$?
    if [[ "$include_boot_targets" == "true" ]]; then
      check_boot_target_readiness || target_status=$?
    fi
    
    echo
    print_separator
    
    if [[ $ssh_status -eq 0 ]] && [[ $wifi_status -eq 0 ]] && [[ $target_status -eq 0 ]]; then
        if [[ "$include_boot_targets" == "true" ]] && (( BOOT_ORDER_LAST_CHECKED > 0 )); then
          offer_sync_connectivity_to_validated_targets || true
          offer_preboot_connectivity_update_to_next_target || true
        fi

        if [[ "$include_boot_targets" == "true" ]] && (( BOOT_ORDER_LAST_CHECKED == 0 )); then
          log_warning "Current-device checks passed, but boot-target readiness is inconclusive"
        else
          log_success "All configuration checks passed!"
        fi
        print_separator
        return 0
    else
        echo
        log_error "Configuration verification failed"
        echo
        
        if [[ $ssh_status -ne 0 ]]; then
            log_info "SSH issues detected. To fix:"
            log_info "  - Enable SSH: sudo systemctl enable --now ssh"
            log_info "  - Ensure user 'admin' exists and is in sudo group"
            log_info "  - Verify PasswordAuthentication is enabled in /etc/ssh/sshd_config"
        fi

        if [[ $wifi_status -ne 0 ]]; then
            log_info "WiFi issues detected. To fix:"
          log_info "  - Configure SSID/PSK in /etc/wpa_supplicant/wpa_supplicant.conf"
          log_info "  - Or add a NetworkManager profile in /etc/NetworkManager/system-connections"
            log_info "  - Ensure wlan interface is up and connected"
        fi

        if [[ $target_status -ne 0 ]]; then
            log_info "Boot target readiness issues detected. To fix:"
            log_info "  - Run option 7 again to re-apply SSH/user/network settings on target"
            log_info "  - Verify EEPROM BOOT_ORDER points to your intended boot device"
            log_info "  - Use option 8 to validate target SSH continuity in selected boot order"
        fi

        print_separator
        return 1
    fi
}

# Helper: Unified preflight checks for boot/root operations
# Usage: preflight_boot_change "operation description" [--ensure-current-ssh]
preflight_boot_change() {
  local operation="$1"
  shift
  local ensure_current_ssh=false
  local allow_failure=false

  while [[ $# -gt 0 ]]; do
    case "$1" in
      --ensure-current-ssh)
        ensure_current_ssh=true
        shift
        ;;
      --allow-failure)
        allow_failure=true
        shift
        ;;
      *)
        shift
        ;;
    esac
  done

  echo
  log_success "Verifying SSH and WiFi configuration before ${operation}..."
  if ! check_initial_config; then
    if [[ "$allow_failure" == "true" ]]; then
      return 1
    fi
    echo
    err "Cannot proceed with ${operation} - fix SSH/WiFi configuration first"
  fi

  if [[ "$ensure_current_ssh" == "true" ]]; then
    ensure_ssh_continuity "/" "$CURRENT_ROOT" "fatal"
  fi

  return 0
}

# Helper: Unified SSH continuity wrapper
# Usage: ensure_ssh_continuity "/mnt/root" "target label" [fatal|nonfatal] [apply|check]
ensure_ssh_continuity() {
  local root_mount_point="$1"
  local target_label="$2"
  local mode="${3:-fatal}"
  local action="${4:-apply}"

  if ! ensure_ssh_continuity_on_target "$root_mount_point" "$target_label" "$action"; then
    if [[ "$mode" == "fatal" ]]; then
      err "SSH continuity check failed for ${target_label}"
    else
      log_warning "SSH continuity check failed for ${target_label}"
      return 1
    fi
  fi

  return 0
}

# Helper: Copy SSH server config (sshd_config + drop-ins) to target root
# Usage: copy_ssh_server_config_to_target "/mnt/target" "target label"
copy_ssh_server_config_to_target() {
  local root_mount_point="$1"
  local target_label="$2"
  local ssh_target_dir="${root_mount_point}/etc/ssh"
  local ssh_config_d_src="/etc/ssh/sshd_config.d"

  if [[ ! -d "$ssh_target_dir" ]]; then
    log_warning "Target SSH directory missing on ${target_label}: ${ssh_target_dir}"
    return 1
  fi

  if [[ -f /etc/ssh/sshd_config ]]; then
    if cp /etc/ssh/sshd_config "${ssh_target_dir}/sshd_config" 2>/dev/null; then
      log_success "Copied sshd_config to ${target_label}"
    else
      log_warning "Failed to copy sshd_config to ${target_label}"
      return 1
    fi
  else
    log_warning "Source /etc/ssh/sshd_config not found - skipping"
  fi

  if [[ -d "$ssh_config_d_src" ]]; then
    local target_dropin_dir="${ssh_target_dir}/sshd_config.d"
    mkdir -p "$target_dropin_dir"

    local copied_dropins=0
    local conf
    for conf in "$ssh_config_d_src"/*.conf; do
      [[ -f "$conf" ]] || continue
      if cp "$conf" "$target_dropin_dir/" 2>/dev/null; then
        copied_dropins=$((copied_dropins + 1))
      fi
    done

    if (( copied_dropins > 0 )); then
      chmod 644 "$target_dropin_dir"/*.conf 2>/dev/null || true
      chown root:root "$target_dropin_dir"/*.conf 2>/dev/null || true
      log_success "Copied ${copied_dropins} sshd_config.d drop-in file(s)"
    fi
  fi

  return 0
}

# Helper: Apply shared connectivity baseline on target root
# Usage: apply_connectivity_baseline_to_target "/mnt/target" "target label" [copy_network_mode] [enable_ssh_mode]
apply_connectivity_baseline_to_target() {
  local root_mount_point="$1"
  local target_label="$2"
  local copy_network_mode="${3:-default}"
  local enable_ssh_mode="${4:-nonfatal}"
  local enable_failed=0

  if ! enable_ssh_service_on_target "$root_mount_point" "$target_label"; then
    log_warning "Failed to enable SSH service on ${target_label}"
    enable_failed=1
  fi

  if [[ "$copy_network_mode" == "force" ]]; then
    copy_network_config "$root_mount_point" "force" || true
  else
    copy_network_config "$root_mount_point" || true
  fi

  if [[ "$enable_ssh_mode" == "fatal" && $enable_failed -ne 0 ]]; then
    return 1
  fi

  if [[ $enable_failed -ne 0 ]]; then
    return 1
  fi

  return 0
}

# Helper: Apply shared connectivity sync steps (SSH config copy + baseline apply).
# Usage: run_connectivity_sync_steps "/mnt/target" "target label" [copy_network_mode] [baseline_mode]
# baseline_mode: fatal|nonfatal
run_connectivity_sync_steps() {
  local root_mount_point="$1"
  local target_label="$2"
  local copy_network_mode="${3:-default}"
  local baseline_mode="${4:-nonfatal}"

  if ! copy_ssh_server_config_to_target "$root_mount_point" "$target_label"; then
    log_warning "SSH server config copy incomplete for ${target_label}"
  fi

  if ! apply_connectivity_baseline_to_target "$root_mount_point" "$target_label" "$copy_network_mode" "$baseline_mode"; then
    if [[ "$baseline_mode" == "fatal" ]]; then
      return 1
    fi
    log_warning "Connectivity baseline apply incomplete for ${target_label}"
    return 1
  fi

  return 0
}

# Helper: Sync current SSH and WiFi connectivity settings to a mounted target root
# Usage (via with_mount): sync_connectivity_on_mounted_root "/mnt/target" "/dev/sda2" "target label"
sync_connectivity_on_mounted_root() {
  local root_mount_point="$1"
  local mounted_device="$2"
  local target_label="${3:-$mounted_device}"

  log_info "Syncing connectivity settings to ${target_label}"

  run_connectivity_sync_steps "$root_mount_point" "$target_label" "default" "nonfatal" || true

  # Read-only verify SSH continuity after sync
  if ! ensure_ssh_continuity_on_target "$root_mount_point" "$target_label" "check"; then
    log_warning "Post-sync SSH continuity check did not fully pass on ${target_label}"
  fi

  return 0
}

# Helper: Check whether connectivity sync is needed for a mounted target root.
# Returns: 0 if sync is needed, 1 if target is already aligned.
# Usage: connectivity_sync_needed_on_mounted_root "/mnt/target" "/dev/sda2"
connectivity_sync_needed_on_mounted_root() {
  local root_mount_point="$1"
  local target_label="$2"
  local -a drift_items=()

  compare_if_present() {
    local src="$1"
    local rel="$2"
    local dst="${root_mount_point}${rel}"

    [[ -f "$src" ]] || return 0
    if [[ ! -f "$dst" ]]; then
      drift_items+=("${rel} (missing on target)")
      return 0
    fi
    if ! cmp -s "$src" "$dst" 2>/dev/null; then
      drift_items+=("${rel} (content differs)")
    fi
  }

  compare_if_present "/etc/ssh/sshd_config" "/etc/ssh/sshd_config"
  compare_if_present "/etc/hostname" "/etc/hostname"
  compare_if_present "/etc/hosts" "/etc/hosts"

  # Static network config is only copied when interface stanzas are present.
  if [[ -f /etc/dhcpcd.conf ]] && grep -q "^interface" /etc/dhcpcd.conf; then
    compare_if_present "/etc/dhcpcd.conf" "/etc/dhcpcd.conf"
  fi

  if [[ -d /etc/ssh/sshd_config.d ]]; then
    local src_dropin
    for src_dropin in /etc/ssh/sshd_config.d/*.conf; do
      [[ -f "$src_dropin" ]] || continue
      compare_if_present "$src_dropin" "/etc/ssh/sshd_config.d/$(basename "$src_dropin")"
    done
  fi

  local nm_profiles_present=false
  if [[ -d /etc/NetworkManager/system-connections ]]; then
    local src_nm
    for src_nm in /etc/NetworkManager/system-connections/*; do
      [[ -f "$src_nm" ]] || continue
      nm_profiles_present=true
      compare_if_present "$src_nm" "/etc/NetworkManager/system-connections/$(basename "$src_nm")"
    done
  fi

  compare_if_present "/etc/wpa_supplicant/wpa_supplicant.conf" "/etc/wpa_supplicant/wpa_supplicant.conf"

  # Baseline enablement checks that sync would enforce.
  if [[ ! -e "${root_mount_point}/etc/systemd/system/multi-user.target.wants/ssh.service" ]]; then
    drift_items+=("/etc/systemd/system/multi-user.target.wants/ssh.service (missing)")
  fi

  if [[ "$nm_profiles_present" == "true" && -f "${root_mount_point}/lib/systemd/system/NetworkManager.service" ]]; then
    if [[ ! -e "${root_mount_point}/etc/systemd/system/multi-user.target.wants/NetworkManager.service" ]]; then
      drift_items+=("/etc/systemd/system/multi-user.target.wants/NetworkManager.service (missing)")
    fi
  fi

  if (( ${#drift_items[@]} == 0 )); then
    return 1
  fi

  log_info "Connectivity drift detected on ${target_label}:"
  local item
  for item in "${drift_items[@]}"; do
    log_info "  - ${item}"
  done
  return 0
}

# Helper: Assess a mounted target and emit "needs" or "same" to output file.
# Usage (via with_mount): assess_connectivity_sync_need_on_mounted_root "/mnt/target" "/dev/sda2" "/tmp/out" "label"
assess_connectivity_sync_need_on_mounted_root() {
  local root_mount_point="$1"
  local mounted_device="$2"
  local output_file="$3"
  local target_label="${4:-$mounted_device}"

  if connectivity_sync_needed_on_mounted_root "$root_mount_point" "$target_label"; then
    printf '%s\n' "needs" > "$output_file"
  else
    printf '%s\n' "same" > "$output_file"
  fi

  return 0
}

# Helper: Offer to sync connectivity settings from current system to validated target roots
offer_sync_connectivity_to_validated_targets() {
  local -A unique_targets=()
  local root_dev

  for root_dev in "${BOOT_ORDER_VALIDATED_ROOT_PARTS[@]}"; do
    [[ -z "$root_dev" ]] && continue
    [[ "$root_dev" == "$CURRENT_ROOT" ]] && continue
    unique_targets["$root_dev"]=1
  done

  local targets=()
  local target
  for target in "${!unique_targets[@]}"; do
    targets+=("$target")
  done

  if (( ${#targets[@]} == 0 )); then
    log_info "No non-current validated target roots found for connectivity sync"
    return 0
  fi

  local targets_needing_sync=()
  local targets_already_synced=()
  local assessment_failed=()

  for target in "${targets[@]}"; do
    local need_file
    need_file=$(mktemp /tmp/connectivity-sync-need-XXXXXX)

    if with_mount "$target" assess_connectivity_sync_need_on_mounted_root "$need_file" "$target"; then
      local need_status
      need_status=$(cat "$need_file" 2>/dev/null || true)
      if [[ "$need_status" == "needs" ]]; then
        targets_needing_sync+=("$target")
      else
        targets_already_synced+=("$target")
      fi
    else
      assessment_failed+=("$target")
    fi

    rm -f "$need_file" 2>/dev/null || true
  done

  if (( ${#targets_already_synced[@]} > 0 )); then
    log_info "Connectivity already aligned on validated target(s):"
    for target in "${targets_already_synced[@]}"; do
      log_info "  - ${target}"
    done
  fi

  if (( ${#assessment_failed[@]} > 0 )); then
    log_warning "Could not assess connectivity drift on:"
    for target in "${assessment_failed[@]}"; do
      log_warning "  - ${target}"
    done
  fi

  if (( ${#targets_needing_sync[@]} == 0 )); then
    if (( ${#assessment_failed[@]} == 0 )); then
      log_success "All validated target roots already match current SSH/WiFi settings"
    else
      log_info "No assessed targets require connectivity sync"
    fi
    return 0
  fi

  echo
  log_info "Validated target root devices that need connectivity sync:"
  for target in "${targets_needing_sync[@]}"; do
    log_info "  - ${target}"
  done

  if ! prompt_yn "Copy current SSH/WiFi settings to these target root devices now?"; then
    log_info "Skipped connectivity sync"
    return 0
  fi

  local synced=0
  local failed=0
  for target in "${targets_needing_sync[@]}"; do
    if with_mount "$target" sync_connectivity_on_mounted_root "$target"; then
      synced=$((synced + 1))
      log_success "Connectivity settings synced to ${target}"
    else
      failed=$((failed + 1))
      log_warning "Connectivity sync failed for ${target}"
    fi
  done

  log_info "Connectivity sync results: ${synced} synced, ${failed} failed"
  return 0
}

# Ensure all USB boot partitions have power settings and scripts
ensure_all_usb_boot_partitions() {
  log_step "Checking all USB boot partitions"
  
  # Find all USB storage devices (not the current boot device)
  local current_boot_dev="$(findmnt -no SOURCE "$BOOT_DIR")"
  local current_boot_disk="$(lsblk -no PKNAME "$current_boot_dev" 2>/dev/null || echo "")"
  
  # Safety check: if we can't determine current boot disk, fail safely
  if [[ -z "$current_boot_disk" ]]; then
    log_warning "Could not determine current boot disk - skipping USB partition check for safety"
    log_info "Current boot device: $current_boot_dev"
    return 0
  fi
  
  # Get USB boot partitions using parent-disk transport classification.
  # Note: TRAN is often blank on partition rows (only populated on parent disks),
  # so we reuse the same robust detector used by BOOT_ORDER checks.
  mapfile -t partition_devices < <(list_boot_partitions_for_transport_class "usb")

  # Exclude partitions on the current boot disk.
  local filtered_partitions=()
  local part_dev
  for part_dev in "${partition_devices[@]}"; do
    local pkname
    pkname=$(lsblk -no PKNAME "$part_dev" 2>/dev/null | head -n1 || true)
    [[ -z "$pkname" ]] && continue
    if [[ "$pkname" != "$current_boot_disk" ]]; then
      filtered_partitions+=("$part_dev")
    fi
  done
  partition_devices=("${filtered_partitions[@]}")
  
  if [[ ${#partition_devices[@]} -eq 0 ]]; then
    log_info "No additional USB devices detected (other than current boot device)."
    return 0
  fi
  
  log_success "Found ${#partition_devices[@]} USB boot partition(s) to check"
  
  local configured=0
  local failed=0
  
  for part_dev in "${partition_devices[@]}"; do
    log_info "Checking $part_dev..."
    
    # Mount, verify as boot partition, configure, and unmount
    if with_mount "$part_dev" ensure_boot_configured --verify-boot "$part_dev"; then
      configured=$((configured + 1))
      log_success "Configured $part_dev"
    else
      failed=$((failed + 1))
      log_error "Failed to configure $part_dev"
    fi
  done
  
  echo
  if [[ $configured -gt 0 ]]; then
    log_success "USB boot partitions: $configured configured successfully"
  fi
  
  if [[ $failed -gt 0 ]]; then
    log_warning "$failed partition(s) could not be configured"
    return 1
  fi
  
  if [[ $configured -eq 0 && $failed -eq 0 ]]; then
    log_info "No additional USB boot partitions found to configure"
  fi
  
  return 0
}

# Sync system configuration files between devices
sync_system_configs() {
  local root_mount_point="$1"
  local target_name="$2"
  
  log_step "Configuration Sync for ${target_name}"
  echo
  
  local configs_to_check=(
    "/etc/hostname"
    "/etc/hosts"
    "/etc/timezone"
    "/etc/locale.gen"
    "/etc/nftables.conf"
    "/etc/ssh/sshd_config"
  )
  
  local different_configs=()
  
  for config in "${configs_to_check[@]}"; do
    # Skip if source doesn't exist
    [[ ! -f "$config" ]] && continue
    
    # Skip if target doesn't exist (will be noted)
    if [[ ! -f "${root_mount_point}${config}" ]]; then
      different_configs+=("$config (missing on target)")
      continue
    fi
    
    # Compare files
    if ! cmp -s "$config" "${root_mount_point}${config}" 2>/dev/null; then
      different_configs+=("$config")
    fi
  done
  
  if [[ ${#different_configs[@]} -eq 0 ]]; then
    log_info "All configuration files match between current system and ${target_name}"
    return 0
  fi
  
  log_info "Configuration differences detected:"
  for cfg in "${different_configs[@]}"; do
    log_info "  - $cfg"
  done

  echo
  log_info "Summary (source → target):"
  log_info "  Path | Source size/mtime | Target size/mtime"
  for config in "${configs_to_check[@]}"; do
    [[ ! -f "$config" ]] && continue
    if [[ ! -f "${root_mount_point}${config}" ]]; then
      local src_size src_time
      src_size=$(stat -c %s "$config" 2>/dev/null || echo "?")
      src_time=$(stat -c %y "$config" 2>/dev/null | cut -d'.' -f1 || echo "?")
      log_info "  ${config} | ${src_size} bytes @ ${src_time} | (missing)"
      continue
    fi

    if ! cmp -s "$config" "${root_mount_point}${config}" 2>/dev/null; then
      local src_size src_time tgt_size tgt_time
      src_size=$(stat -c %s "$config" 2>/dev/null || echo "?")
      src_time=$(stat -c %y "$config" 2>/dev/null | cut -d'.' -f1 || echo "?")
      tgt_size=$(stat -c %s "${root_mount_point}${config}" 2>/dev/null || echo "?")
      tgt_time=$(stat -c %y "${root_mount_point}${config}" 2>/dev/null | cut -d'.' -f1 || echo "?")
      log_info "  ${config} | ${src_size} bytes @ ${src_time} | ${tgt_size} bytes @ ${tgt_time}"
    fi
  done
  log_info "Action: If you proceed, source files will overwrite target files listed above."
  
  if ! prompt_with_context "Configuration Sync" "" "Copy current system configurations to ${target_name}?"; then
    echo "Configuration sync skipped"
    return 0
  fi
  
  local synced=0
  for config in "${configs_to_check[@]}"; do
    [[ ! -f "$config" ]] && continue
    
    if cp "$config" "${root_mount_point}${config}" 2>/dev/null; then
      log_success "Synced $config"
      synced=$((synced + 1))
    fi
  done
  
  # SSH host keys are intentionally not copied (minimal access settings only)
  
  log_success "Synced $synced configuration file(s)"
  return 0
}

################################################################################
# USER/SSH CONFIGURATION HELPERS (FRESH INSTALLS)
################################################################################

# Helper: Copy user account to target system
# Usage: copy_user_account "/mnt/fresh_root" "username"
copy_user_account() {
  local root_mount_point="$1"
  local username="$2"
  
  [[ -z "$username" ]] && return 1
  
  # Verify user exists on source system
  if ! id "$username" &>/dev/null; then
    log_error "User '$username' does not exist on source system"
    return 1
  fi

  # Ensure target auth files exist and are regular files
  local target_passwd="$root_mount_point/etc/passwd"
  local target_shadow="$root_mount_point/etc/shadow"
  local target_group="$root_mount_point/etc/group"

  for f in "$target_passwd" "$target_shadow" "$target_group"; do
    if [[ ! -f "$f" ]]; then
      log_error "Missing target auth file: $f"
      return 1
    fi
  done

  if [[ ! -w "$target_passwd" || ! -w "$target_shadow" || ! -w "$target_group" ]]; then
    log_error "Target auth files are not writable"
    return 1
  fi
  
  log_success "Configuring user: $username"
  
  # Get actual UID/GID for this user
  local user_uid=$(id -u "$username")
  local user_gid=$(id -g "$username")
  
  # Check if user has admin privileges (sudo group)
  local has_sudo=false
  if groups "$username" | grep -q "\bsudo\b"; then
    has_sudo=true
    log_success "User $username has sudo privileges"
  fi
  
  # Ensure user has sudo access (auto-add for streamlined continuity)
  if [[ "$has_sudo" == "false" ]]; then
    log_warning "User '$username' is NOT in the sudo group - adding now"
    if usermod -aG sudo "$username"; then
      log_success "Added $username to sudo group"
      has_sudo=true
    else
      log_error "Failed to add user to sudo group"
      return 1
    fi
  fi
  
  # Only copy user if not already present on target (prevent duplicates)
  if ! grep -q "^${username}:" "$target_passwd" 2>/dev/null; then
    local src_passwd
    local src_shadow
    local src_group

    src_passwd=$(grep -m1 "^${username}:" /etc/passwd 2>/dev/null || true)
    src_shadow=$(grep -m1 "^${username}:" /etc/shadow 2>/dev/null || true)
    src_group=$(grep -m1 "^${username}:" /etc/group 2>/dev/null || true)

    if [[ -z "$src_passwd" || -z "$src_shadow" || -z "$src_group" ]]; then
      log_error "Failed to read source account entries for $username"
      return 1
    fi

    printf '%s\n' "$src_passwd" >> "$target_passwd" || return 1
    printf '%s\n' "$src_shadow" >> "$target_shadow" || return 1
    printf '%s\n' "$src_group" >> "$target_group" || return 1
  else
    log_success "[i] User $username already exists on target, leaving password unchanged"
  fi
  
  # Copy sudo group membership if user should have sudo
  if [[ "$has_sudo" == "true" ]]; then
    if ! grep -q "^sudo:.*${username}" "$target_group"; then
      sed -i "s/^sudo:\([^:]*\):\([^:]*\):\(.*\)/sudo:\1:\2:\3,${username}/" "$target_group"
      log_success "Added $username to sudo group on target"
    fi
  fi
  
  # Create home directory if doesn't exist (use actual UID/GID)
  if [[ ! -d "$root_mount_point/home/$username" ]]; then
    mkdir -p "$root_mount_point/home/$username"
    chown "$user_uid:$user_gid" "$root_mount_point/home/$username"
    chmod 755 "$root_mount_point/home/$username"
    log_success "Created home directory with correct ownership (UID:$user_uid GID:$user_gid)"
  fi
  
  # Note: authorized_keys is intentionally not copied (minimal access settings only)
  
  return 0
}

# Helper: Enable SSH service on target system (no key copy)
# Usage: enable_ssh_service_on_target "/mnt/fresh_root" "target label"
enable_ssh_service_on_target() {
  local root_mount_point="$1"
  local target_label="$2"

  mkdir -p "$root_mount_point/etc/systemd/system/multi-user.target.wants"
  ln -sf /lib/systemd/system/ssh.service "$root_mount_point/etc/systemd/system/multi-user.target.wants/ssh.service" 2>/dev/null || true
  log_success "SSH service enabled on ${target_label}"
  return 0
}

# Helper: Detect the invoking (non-root) user for configuration copy
detect_invoking_user() {
  local detected_user="${SUDO_USER:-}"
  if [[ -z "$detected_user" || "$detected_user" == "root" ]]; then
    detected_user=$(logname 2>/dev/null || true)
  fi
  if [[ -z "$detected_user" || "$detected_user" == "root" ]]; then
    detected_user=$(whoami 2>/dev/null || true)
  fi
  [[ "$detected_user" == "root" ]] && detected_user=""

  echo "$detected_user"
}

# Helper: Get user groups from a mounted root filesystem
# Usage: get_user_groups_from_root "/mnt/target_root" "username"
get_user_groups_from_root() {
  local root_mount_point="$1"
  local username="$2"
  local passwd_file="${root_mount_point}/etc/passwd"
  local group_file="${root_mount_point}/etc/group"

  [[ -f "$passwd_file" && -f "$group_file" ]] || { echo ""; return 0; }

  local primary_gid
  primary_gid=$(awk -F: -v u="$username" '$1==u {print $4; exit}' "$passwd_file" 2>/dev/null)

  local groups=""
  if [[ -n "$primary_gid" ]]; then
    local primary_group
    primary_group=$(awk -F: -v gid="$primary_gid" '$3==gid {print $1; exit}' "$group_file" 2>/dev/null)
    [[ -n "$primary_group" ]] && groups="$primary_group"
  fi

  local supplemental
  supplemental=$(awk -F: -v u="$username" 'BEGIN{found=""} $4 ~ ("(^|,)"u"(,|$)") {print $1}' "$group_file" 2>/dev/null | paste -sd ' ' -)
  if [[ -n "$supplemental" ]]; then
    groups="$groups $supplemental"
  fi

  echo "$groups"
}

# Helper: Ensure target root has SSH continuity (single user, password auth, root disabled)
# Usage: ensure_ssh_continuity_on_target "/mnt/target_root" "target label" [apply|check]
ensure_ssh_continuity_on_target() {
  local root_mount_point="$1"
  local target_label="$2"
  local action="${3:-apply}"
  local detected_user
  local target_sshd_config="${root_mount_point}/etc/ssh/sshd_config"
  local sshd_config_d_target="${root_mount_point}/etc/ssh/sshd_config.d"
  local access_dropin="${sshd_config_d_target}/99-switch-root-access.conf"

  if [[ "$action" == "check" ]]; then
    log_step "Checking SSH continuity on ${target_label} (read-only)"
  else
    log_step "Ensuring SSH continuity on ${target_label}"
  fi

  detected_user=$(detect_invoking_user)
  if [[ -z "$detected_user" ]]; then
    if [[ "$action" == "check" ]]; then
      log_warning "No non-root user detected for check-only mode"
    else
      read -rp "  Enter username to configure on target system: " detected_user
      if [[ -z "$detected_user" ]]; then
        log_error "No username provided - cannot guarantee SSH access"
        return 1
      fi
    fi
  fi

  # Ensure user exists on target (authorized_keys not copied)
  if [[ "$action" != "check" ]]; then
    if ! copy_user_account "$root_mount_point" "$detected_user"; then
      log_error "Failed to configure user '$detected_user' on ${target_label}"
      return 1
    fi
  else
    if [[ -n "$detected_user" ]]; then
      if ! grep -q "^${detected_user}:" "${root_mount_point}/etc/passwd" 2>/dev/null; then
        log_error "User '${detected_user}' not present on ${target_label} (check-only)"
        return 1
      fi
    fi
  fi

  # Enforce single-user access, password auth, and disable root SSH
  if [[ "$action" != "check" ]]; then
    mkdir -p "$sshd_config_d_target"
    printf 'PermitRootLogin no\nPasswordAuthentication yes\nAllowUsers %s\n' "$detected_user" > "$access_dropin"
    chmod 644 "$access_dropin" 2>/dev/null || true
    chown root:root "$access_dropin" 2>/dev/null || true
    log_success "Applied SSH access rules (PermitRootLogin no; AllowUsers ${detected_user}; PasswordAuthentication yes)"
  fi

  if [[ ! -f "$target_sshd_config" ]]; then
    if [[ "$action" == "check" ]]; then
      log_error "Target sshd_config missing (check-only) - cannot validate SSH continuity"
    else
      log_error "Target sshd_config missing - cannot guarantee SSH access"
    fi
    return 1
  fi

  local target_groups=""
  if [[ -n "$detected_user" ]]; then
    target_groups=$(get_user_groups_from_root "$root_mount_point" "$detected_user")
  fi

  if ! evaluate_sshd_access_policy "$target_sshd_config" "$sshd_config_d_target" "$detected_user" "$target_groups" "Target "; then
    local policy_issue
    for policy_issue in "${SSHD_POLICY_ISSUES[@]}"; do
      log_error "$policy_issue"
    done
    return 1
  fi

  if [[ "$action" != "check" ]]; then
    # Apply shared SSH service + network baseline
    if ! run_connectivity_sync_steps "$root_mount_point" "$target_label" "force" "fatal"; then
      log_error "Failed to apply connectivity baseline on ${target_label}"
      return 1
    fi

    # Force default password to ensure login continuity
    if ! set_target_user_password "$root_mount_point" "$detected_user" "$DEFAULT_SSH_PASSWORD" "$target_label"; then
      log_error "Failed to set default SSH password on ${target_label}"
      return 1
    fi

    log_success "SSH access configured for user: ${detected_user}"
    log_success "Default SSH password set for ${detected_user}"
  else
    log_success "SSH continuity check passed (read-only)"
  fi
  return 0
}

# Helper: Set a user password in the target root
# Usage: set_target_user_password "/mnt/target_root" "username" "password" "target label"
set_target_user_password() {
  local root_mount_point="$1"
  local username="$2"
  local new_password="$3"
  local target_label="$4"

  if [[ -z "$username" ]]; then
    log_error "No username provided - cannot set password on ${target_label}"
    return 1
  fi

  if [[ "$root_mount_point" == "/" ]]; then
    if echo "${username}:${new_password}" | chpasswd 2>/dev/null; then
      log_success "Password set for ${username} on current system"
      return 0
    fi
  else
    if [[ -x "${root_mount_point}/usr/sbin/chpasswd" ]]; then
      if chroot "$root_mount_point" /usr/sbin/chpasswd <<< "${username}:${new_password}" 2>/dev/null; then
        log_success "Password set for ${username} on ${target_label}"
        return 0
      fi
    elif [[ -x "${root_mount_point}/sbin/chpasswd" ]]; then
      if chroot "$root_mount_point" /sbin/chpasswd <<< "${username}:${new_password}" 2>/dev/null; then
        log_success "Password set for ${username} on ${target_label}"
        return 0
      fi
    fi

    # Fallback: update /etc/shadow directly using openssl hash
    if command -v openssl >/dev/null 2>&1; then
      local hash
      hash=$(openssl passwd -6 "$new_password" 2>/dev/null || true)
      if [[ -n "$hash" && -f "${root_mount_point}/etc/shadow" ]]; then
        if sed -i -E "s#^${username}:[^:]*:#${username}:${hash}:#" "${root_mount_point}/etc/shadow"; then
          log_success "Password set for ${username} on ${target_label} (shadow update)"
          return 0
        fi
      fi
    fi
  fi

  log_error "Failed to set password for ${username} on ${target_label}"
  return 1
}


# Helper: Extract root device from a boot partition's cmdline.txt
# Usage: extract_root_device_from_boot_partition "/mnt/boot" "/dev/sda1" "/tmp/outfile"
extract_root_device_from_boot_partition() {
  local boot_mount_point="$1"
  local boot_device="$2"
  local output_file="$3"
  local cmdline_file="${boot_mount_point}/cmdline.txt"

  if [[ ! -f "$cmdline_file" ]]; then
    log_warning "cmdline.txt not found on boot partition: ${boot_device}"
    return 1
  fi

  local cmdline_root
  local cmdline_root_value
  local cmdline_root_device

  cmdline_root="$(grep -o 'root=[^ ]*' "$cmdline_file" || true)"
  cmdline_root_value="${cmdline_root#root=}"
  cmdline_root_device="$(resolve_cmdline_root_device "$cmdline_root_value")"

  if [[ -z "$cmdline_root_device" ]]; then
    log_warning "Unable to resolve root device from ${cmdline_file}"
    return 1
  fi

  echo "$cmdline_root_device" > "$output_file"
  return 0
}

# Helper: List bootable FAT partitions by transport class (sd|usb|pcie)
list_boot_partitions_for_transport_class() {
  local class="$1"
  local part
  local pkname
  local fstype

  while read -r part; do
    [[ -z "$part" ]] && continue

    # Derive parent disk per partition and normalize to basename (sda, mmcblk0, nvme0n1).
    pkname=$(lsblk -no PKNAME "$part" 2>/dev/null | head -n1 || true)
    pkname="${pkname##*/}"
    [[ -n "$pkname" ]] || continue

    # Read FSTYPE per-partition to avoid column-shift parsing issues when fields are blank.
    fstype=$(lsblk -no FSTYPE "$part" 2>/dev/null | head -n1 || true)
    [[ -z "$fstype" ]] && fstype=$(blkid -s TYPE -o value "$part" 2>/dev/null | head -n1 || true)

    # Accept common FAT filesystem labels for boot partitions.
    case "${fstype,,}" in
      vfat|fat|fat16|fat32|msdos) ;;
      *) continue ;;
    esac

    if disk_matches_transport_class "$pkname" "$class"; then
      printf '%s\n' "$part"
    fi
  done < <(lsblk -nrpo NAME,TYPE | awk '$2=="part" {print $1}')
}

# Resolve a filesystem identifier (PARTUUID/UUID/LABEL/PARTLABEL/device path) to a block device.
resolve_device_identifier() {
  local identifier="$1"
  local resolved=""

  case "$identifier" in
    PARTUUID=*)
      resolved=$(blkid -t "PARTUUID=${identifier#PARTUUID=}" -o device 2>/dev/null | head -n1)
      ;;
    UUID=*)
      resolved=$(blkid -t "UUID=${identifier#UUID=}" -o device 2>/dev/null | head -n1)
      ;;
    LABEL=*)
      resolved=$(blkid -t "LABEL=${identifier#LABEL=}" -o device 2>/dev/null | head -n1)
      ;;
    PARTLABEL=*)
      resolved=$(blkid -t "PARTLABEL=${identifier#PARTLABEL=}" -o device 2>/dev/null | head -n1)
      ;;
    /dev/*)
      resolved="$identifier"
      ;;
    *)
      resolved=""
      ;;
  esac

  echo "$resolved"
}

# Verify boot cmdline points to the expected root device.
# Usage (via with_mount): verify_boot_cmdline_targets_root "/mnt/boot" "/dev/sda1" "/dev/sda2" "root_partuuid" "label"
verify_boot_cmdline_targets_root() {
  local boot_mount_point="$1"
  local boot_device="$2"
  local expected_root_device="$3"
  local expected_root_partuuid="$4"
  local target_label="$5"
  local cmdline_file="${boot_mount_point}/cmdline.txt"

  [[ -f "$cmdline_file" ]] || { log_error "cmdline.txt missing on ${target_label}"; return 1; }

  local cmdline_root
  local cmdline_root_value
  local resolved_root
  cmdline_root="$(grep -o 'root=[^ ]*' "$cmdline_file" || true)"
  cmdline_root_value="${cmdline_root#root=}"

  log_info "[diag] ${target_label}: cmdline root token=${cmdline_root_value:-missing}"

  if [[ -z "$cmdline_root_value" || "$cmdline_root_value" == "$cmdline_root" ]]; then
    log_error "root=... not found in ${cmdline_file} (${target_label})"
    return 1
  fi

  resolved_root="$(resolve_device_identifier "$cmdline_root_value")"
  if [[ -z "$resolved_root" ]]; then
    log_error "Could not resolve cmdline root '${cmdline_root_value}' on ${target_label}"
    return 1
  fi

  log_info "[diag] ${target_label}: cmdline root resolves to ${resolved_root}"

  if [[ "$resolved_root" != "$expected_root_device" ]]; then
    log_error "Cmdline root mismatch on ${target_label}: ${cmdline_root_value} -> ${resolved_root}, expected ${expected_root_device}"
    return 1
  fi

  if [[ "$cmdline_root_value" == PARTUUID=* && -n "$expected_root_partuuid" ]]; then
    local cmd_partuuid="${cmdline_root_value#PARTUUID=}"
    if [[ "$cmd_partuuid" != "$expected_root_partuuid" ]]; then
      log_error "Cmdline PARTUUID mismatch on ${target_label}: ${cmd_partuuid}, expected ${expected_root_partuuid}"
      return 1
    fi
  fi

  if ! grep -qE '(^| )rootwait( |$)' "$cmdline_file"; then
    log_error "rootwait is missing in ${cmdline_file} (${target_label})"
    return 1
  fi

  return 0
}

# Verify root fstab points to the expected boot/root devices.
# Usage (via with_mount): verify_root_fstab_targets_boot_and_root "/mnt/root" "/dev/sda2" "/dev/sda1" "root_partuuid" "boot_partuuid" "label"
verify_root_fstab_targets_boot_and_root() {
  local root_mount_point="$1"
  local mounted_root_device="$2"
  local expected_boot_device="$3"
  local expected_root_partuuid="$4"
  local expected_boot_partuuid="$5"
  local target_label="$6"
  local fstab_path="${root_mount_point}/etc/fstab"

  [[ -f "$fstab_path" ]] || { log_error "fstab missing on ${target_label}"; return 1; }

  local root_source
  local boot_source
  root_source=$(awk '!/^\s*#/ && $2=="/" {print $1; exit}' "$fstab_path")
  boot_source=$(awk '!/^\s*#/ && $2=="/boot/firmware" {print $1; exit}' "$fstab_path")

  log_info "[diag] ${target_label}: fstab / source=${root_source:-missing}"
  log_info "[diag] ${target_label}: fstab /boot/firmware source=${boot_source:-missing}"

  [[ -n "$root_source" ]] || { log_error "Root fstab entry missing on ${target_label}"; return 1; }
  [[ -n "$boot_source" ]] || { log_error "Boot fstab entry missing on ${target_label}"; return 1; }

  local resolved_root
  resolved_root="$(resolve_device_identifier "$root_source")"
  if [[ -z "$resolved_root" ]]; then
    log_error "Could not resolve fstab root source '${root_source}' on ${target_label}"
    return 1
  fi
  log_info "[diag] ${target_label}: fstab / resolves to ${resolved_root}"
  if [[ "$resolved_root" != "$mounted_root_device" ]]; then
    log_error "fstab root mismatch on ${target_label}: ${root_source} -> ${resolved_root}, expected ${mounted_root_device}"
    return 1
  fi

  local resolved_boot
  resolved_boot="$(resolve_device_identifier "$boot_source")"
  if [[ -z "$resolved_boot" ]]; then
    log_error "Could not resolve fstab boot source '${boot_source}' on ${target_label}"
    return 1
  fi
  log_info "[diag] ${target_label}: fstab /boot/firmware resolves to ${resolved_boot}"
  if [[ "$resolved_boot" != "$expected_boot_device" ]]; then
    log_error "fstab boot mismatch on ${target_label}: ${boot_source} -> ${resolved_boot}, expected ${expected_boot_device}"
    return 1
  fi

  if [[ "$root_source" == PARTUUID=* && -n "$expected_root_partuuid" ]]; then
    local fstab_root_partuuid="${root_source#PARTUUID=}"
    if [[ "$fstab_root_partuuid" != "$expected_root_partuuid" ]]; then
      log_error "fstab root PARTUUID mismatch on ${target_label}: ${fstab_root_partuuid}, expected ${expected_root_partuuid}"
      return 1
    fi
  fi

  if [[ "$boot_source" == PARTUUID=* && -n "$expected_boot_partuuid" ]]; then
    local fstab_boot_partuuid="${boot_source#PARTUUID=}"
    if [[ "$fstab_boot_partuuid" != "$expected_boot_partuuid" ]]; then
      log_error "fstab boot PARTUUID mismatch on ${target_label}: ${fstab_boot_partuuid}, expected ${expected_boot_partuuid}"
      return 1
    fi
  fi

  return 0
}

# Verify that a boot/root partition pair is internally consistent for boot.
# Usage: diagnose_boot_pair_readiness "/dev/sda1" "/dev/sda2" "label"
diagnose_boot_pair_readiness() {
  local boot_device="$1"
  local root_device="$2"
  local target_label="${3:-${boot_device} -> ${root_device}}"

  local boot_partuuid=""
  local root_partuuid=""

  log_step "Boot Readiness Diagnostics: ${target_label}"

  boot_partuuid="$(get_partuuid_safe "$boot_device" "warn" 2>/dev/null || true)"
  root_partuuid="$(get_partuuid_safe "$root_device" "warn" 2>/dev/null || true)"

  log_info "Expected boot device: ${boot_device} (PARTUUID=${boot_partuuid:-unknown})"
  log_info "Expected root device: ${root_device} (PARTUUID=${root_partuuid:-unknown})"

  print_boot_readiness_failure_guidance() {
    local bdev="$1"
    local rdev="$2"
    local label="$3"
    echo
    log_info "Boot readiness recovery hints for ${label}:"
    log_info "  1. Verify boot/root identifiers:"
    log_info "     - sudo blkid ${bdev} ${rdev}"
    log_info "  2. Re-run diagnostics and target checks:"
    log_info "     - option 2 (verify initial configuration)"
    log_info "     - option 8 (EEPROM boot order)"
    log_info "  3. Re-apply target config if identifiers drifted:"
    log_info "     - option 7 (fresh flash) or option 6 (clone)"
    log_info "  4. Verify EEPROM points to intended media first (USB+SD or SD+USB)"
    echo
  }

  if ! with_mount "$boot_device" verify_boot_cmdline_targets_root --verify-boot "$root_device" "$root_partuuid" "$target_label"; then
    log_error "Boot cmdline diagnostic failed for ${target_label}"
    print_boot_readiness_failure_guidance "$boot_device" "$root_device" "$target_label"
    return 1
  fi

  if ! with_mount "$root_device" verify_root_fstab_targets_boot_and_root "$boot_device" "$root_partuuid" "$boot_partuuid" "$target_label"; then
    log_error "Root fstab diagnostic failed for ${target_label}"
    print_boot_readiness_failure_guidance "$boot_device" "$root_device" "$target_label"
    return 1
  fi

  log_success "Boot readiness diagnostics passed for ${target_label}"
  return 0
}

# Check whether boot partition has cloud-init seed files staged for first boot.
# Usage (via with_mount): boot_partition_has_first_boot_seed "/mnt/boot" "/dev/sda1"
boot_partition_has_first_boot_seed() {
  local boot_mount_point="$1"

  [[ -f "${boot_mount_point}/user-data" ]] || return 1
  [[ -f "${boot_mount_point}/network-config" ]] || return 1
  return 0
}

# Check whether root partition still appears to be pre-first-boot cloud-init state.
# Usage (via with_mount): root_partition_cloud_init_pending "/mnt/root" "/dev/sda2"
root_partition_cloud_init_pending() {
  local root_mount_point="$1"

  if [[ -f "${root_mount_point}/var/lib/cloud/instance/boot-finished" ]]; then
    return 1
  fi

  if [[ -d "${root_mount_point}/etc/cloud" || -d "${root_mount_point}/var/lib/cloud" ]]; then
    return 0
  fi

  return 1
}

# Treat non-current targets with staged cloud-init as pending until first boot completes.
# Usage: is_target_ssh_continuity_pending_first_boot "/dev/sda1" "/dev/sda2"
is_target_ssh_continuity_pending_first_boot() {
  local boot_part="$1"
  local root_dev="$2"

  if ! with_mount "$boot_part" boot_partition_has_first_boot_seed --verify-boot; then
    return 1
  fi

  if ! with_mount "$root_dev" root_partition_cloud_init_pending; then
    return 1
  fi

  return 0
}

# Helper: Check SSH continuity across all bootable devices in the desired EEPROM boot order
check_boot_order_ssh_continuity() {
  local desired="$1"
  local order="${desired#0x}"
  order="${order,,}"

  BOOT_ORDER_LAST_CHECKED=0
  BOOT_ORDER_LAST_FAILURES=0
  BOOT_ORDER_LAST_SKIPPED=0
  BOOT_ORDER_VALIDATED_BOOT_PARTS=()
  BOOT_ORDER_VALIDATED_ROOT_PARTS=()
  BOOT_ORDER_NEXT_TARGET_BOOT=""
  BOOT_ORDER_NEXT_TARGET_ROOT=""

  log_step "Boot Target Readiness Check (EEPROM Boot Order)"
  log_info "Read-only check: no changes will be made to any device"
  log_info "Checks include: boot files, cmdline root mapping, fstab boot/root mapping, SSH continuity"
  log_info "Desired BOOT_ORDER: ${desired}"
  log_info "Boot order: $(describe_boot_order "$desired")"

  local checked=0
  local failures=0
  local skipped_nonboot=0
  local pending_first_boot=0
  declare -A seen_boot_parts=()

  for (( i=${#order}-1; i>=0; i-- )); do
    local digit="${order:$i:1}"
    local class=""

    case "$digit" in
      0|1) class="sd" ;;
      4|5) class="usb" ;;
      6) class="pcie" ;;
      e|f)
        log_info "Reached BOOT_ORDER terminator ($(boot_order_digit_label "$digit")); no further boot device checks"
        break
        ;;
      2|3|7)
        log_info "Skipping $(boot_order_digit_label "$digit") (no local storage to verify)"
        continue
        ;;
      *)
        log_warning "Unknown boot order digit: ${digit}"
        continue
        ;;
    esac

    log_info "Checking ${class^^} bootable devices..."
    mapfile -t boot_parts < <(list_boot_partitions_for_transport_class "$class")
    if [[ ${#boot_parts[@]} -eq 0 ]]; then
      log_info "No ${class} boot partitions found"
      continue
    fi

    for boot_part in "${boot_parts[@]}"; do
      if [[ -n "${seen_boot_parts[$boot_part]:-}" ]]; then
        log_info "Skipping ${boot_part} (already checked via earlier BOOT_ORDER class)"
        continue
      fi
      seen_boot_parts[$boot_part]=1

      log_info "Evaluating boot partition: ${boot_part}"

      local root_dev_file
      root_dev_file=$(mktemp /tmp/boot-root-dev-XXXXXX)

      if ! with_mount "$boot_part" extract_root_device_from_boot_partition --verify-boot "$root_dev_file"; then
        log_info "Skipping ${boot_part} (not a Raspberry Pi boot partition or cmdline could not be parsed)"
        rm -f "$root_dev_file"
        skipped_nonboot=$((skipped_nonboot + 1))
        continue
      fi

      local root_dev
      root_dev=$(cat "$root_dev_file" 2>/dev/null || true)
      rm -f "$root_dev_file"

      if [[ -z "$root_dev" ]]; then
        log_info "Skipping ${boot_part} (root device unresolved from cmdline)"
        skipped_nonboot=$((skipped_nonboot + 1))
        continue
      fi

      log_info "Root device for ${boot_part}: ${root_dev}"

      if ! diagnose_boot_pair_readiness "$boot_part" "$root_dev" "${boot_part} -> ${root_dev}"; then
        failures=$((failures + 1))
        continue
      fi

      if [[ "$root_dev" == "$CURRENT_ROOT" ]]; then
        if ! ensure_ssh_continuity "/" "$root_dev" "nonfatal" "check"; then
          log_error "SSH continuity check failed for current root (${root_dev})"
          failures=$((failures + 1))
        else
          checked=$((checked + 1))
          BOOT_ORDER_VALIDATED_BOOT_PARTS+=("$boot_part")
          BOOT_ORDER_VALIDATED_ROOT_PARTS+=("$root_dev")
        fi
        continue
      fi

      if [[ ! -b "$root_dev" ]]; then
        log_warning "Root device not found: ${root_dev}"
        failures=$((failures + 1))
        continue
      fi

      if ! with_mount "$root_dev" ensure_ssh_continuity "nonfatal" "check"; then
        if is_target_ssh_continuity_pending_first_boot "$boot_part" "$root_dev"; then
          log_warning "SSH continuity on ${root_dev} is pending first boot cloud-init"
          log_info "Run one successful boot on this target, then re-run option 2/8 for strict SSH validation"
          pending_first_boot=$((pending_first_boot + 1))
          continue
        fi

        log_error "SSH continuity check failed for ${root_dev}"
        failures=$((failures + 1))
      else
        checked=$((checked + 1))
        BOOT_ORDER_VALIDATED_BOOT_PARTS+=("$boot_part")
        BOOT_ORDER_VALIDATED_ROOT_PARTS+=("$root_dev")
        if [[ "$root_dev" != "$CURRENT_ROOT" && -z "$BOOT_ORDER_NEXT_TARGET_ROOT" ]]; then
          BOOT_ORDER_NEXT_TARGET_BOOT="$boot_part"
          BOOT_ORDER_NEXT_TARGET_ROOT="$root_dev"
          log_info "Next non-current boot candidate: ${BOOT_ORDER_NEXT_TARGET_BOOT} -> ${BOOT_ORDER_NEXT_TARGET_ROOT}"
        fi
      fi
    done
  done

  if (( failures > 0 )); then
    BOOT_ORDER_LAST_CHECKED=$checked
    BOOT_ORDER_LAST_FAILURES=$failures
    BOOT_ORDER_LAST_SKIPPED=$skipped_nonboot
    log_error "Boot target readiness check failed for ${failures} device(s)"
    echo
    log_error "One or more boot targets are not safe to boot"
    log_info "Recommended next steps:"
    log_info "  - Run option 8 to confirm EEPROM BOOT_ORDER"
    log_info "  - Run option 7 (fresh flash) or option 6 (clone) to rebuild target mapping"
    log_info "  - Re-run option 2 and confirm diagnostics pass"
    return 1
  fi

  if (( skipped_nonboot > 0 )); then
    log_info "Ignored ${skipped_nonboot} non-boot USB/SD candidate partition(s)"
  fi

  if (( pending_first_boot > 0 )); then
    log_info "Deferred SSH continuity on ${pending_first_boot} target(s) pending first boot cloud-init"
  fi

  if (( checked == 0 )); then
    log_warning "No bootable devices were validated for BOOT_ORDER ${desired}"
    log_info "This often means the intended target media is not attached or not yet configured as a Pi boot device"
    log_info "Tip: Run option 7 (flash) or option 6 (clone), then re-run option 2 or option 8"
  else
    log_info "Validated SSH target root devices:"
    local idx
    for idx in "${!BOOT_ORDER_VALIDATED_ROOT_PARTS[@]}"; do
      log_info "  - boot ${BOOT_ORDER_VALIDATED_BOOT_PARTS[$idx]} -> root ${BOOT_ORDER_VALIDATED_ROOT_PARTS[$idx]}"
    done
  fi

  BOOT_ORDER_LAST_CHECKED=$checked
  BOOT_ORDER_LAST_FAILURES=$failures
  BOOT_ORDER_LAST_SKIPPED=$skipped_nonboot

  log_success "Boot target readiness verified for ${checked} bootable device(s)"
  return 0
}

# Helper: Copy network configuration to target system
# Usage: copy_network_config "/mnt/fresh_root"
copy_network_config() {
  local root_mount_point="$1"
  local wifi_copied=false
  local nm_profiles_copied=false
  local nm_service_file="${root_mount_point}/lib/systemd/system/NetworkManager.service"
  local nm_service_link="${root_mount_point}/etc/systemd/system/multi-user.target.wants/NetworkManager.service"

  if [[ "$root_mount_point" == "/" ]]; then
    return 0
  fi

  # Preserve host identity so mDNS and router naming stay consistent after flash.
  if [[ -f /etc/hostname ]]; then
    cp /etc/hostname "$root_mount_point/etc/hostname" 2>/dev/null || true
    log_success "hostname copied"
  fi

  if [[ -f /etc/hosts ]]; then
    cp /etc/hosts "$root_mount_point/etc/hosts" 2>/dev/null || true
    log_success "hosts file copied"
  fi

  # Copy network configuration if using static IP
  if [[ -f /etc/dhcpcd.conf ]] && grep -q "^interface" /etc/dhcpcd.conf; then
    cp /etc/dhcpcd.conf "$root_mount_point/etc/dhcpcd.conf" 2>/dev/null || {
      log_error "Failed to copy network configuration"
      return 1
    }
    log_success "Network configuration copied"
  fi

  # Copy NetworkManager WiFi profiles if present
  if [[ -d /etc/NetworkManager/system-connections ]]; then
    mkdir -p "$root_mount_point/etc/NetworkManager/system-connections"
    cp /etc/NetworkManager/system-connections/* "$root_mount_point/etc/NetworkManager/system-connections/" 2>/dev/null || true
    chmod 600 "$root_mount_point/etc/NetworkManager/system-connections/"* 2>/dev/null || true
    chown root:root "$root_mount_point/etc/NetworkManager/system-connections/"* 2>/dev/null || true
    log_success "NetworkManager profiles copied"
    nm_profiles_copied=true
    wifi_copied=true
  fi

  # If NetworkManager profiles were copied, make sure NetworkManager starts at boot.
  if [[ "$nm_profiles_copied" == "true" && -f "$nm_service_file" ]]; then
    mkdir -p "${root_mount_point}/etc/systemd/system/multi-user.target.wants"
    ln -sf /lib/systemd/system/NetworkManager.service "$nm_service_link" 2>/dev/null || true
    log_success "NetworkManager service enabled on target"
  fi

  # Copy wpa_supplicant configuration if present (headless WiFi continuity)
  if [[ -f /etc/wpa_supplicant/wpa_supplicant.conf ]]; then
    mkdir -p "$root_mount_point/etc/wpa_supplicant"
    cp /etc/wpa_supplicant/wpa_supplicant.conf "$root_mount_point/etc/wpa_supplicant/wpa_supplicant.conf" 2>/dev/null || {
      log_error "Failed to copy wpa_supplicant.conf"
      return 1
    }
    chmod 600 "$root_mount_point/etc/wpa_supplicant/wpa_supplicant.conf" 2>/dev/null || true
    chown root:root "$root_mount_point/etc/wpa_supplicant/wpa_supplicant.conf" 2>/dev/null || true
    log_success "wpa_supplicant.conf copied"
    wifi_copied=true
  fi

  if [[ "$wifi_copied" != "true" ]]; then
    log_warning "No WiFi configuration copied (NetworkManager profiles and wpa_supplicant.conf not copied)"
  fi
  
  return 0
}


################################################################################
# FRESH INSTALL TEMPLATE HELPERS (OPTION 7)
################################################################################

FLASH_WIFI_SSID=""
FLASH_WIFI_PASSWORD=""
FLASH_USERNAME=""
FLASH_USER_PASSWORD=""
FLASH_HOSTNAME=""

is_valid_linux_username() {
  local username="$1"
  [[ "$username" =~ ^[a-z_][a-z0-9_-]{0,31}$ ]]
}

is_valid_hostname_label() {
  local hostname="$1"
  [[ "$hostname" =~ ^[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?$ ]]
}

yaml_escape_double_quotes() {
  local value="$1"
  value="${value//\\/\\\\}"
  value="${value//\"/\\\"}"
  printf '%s' "$value"
}

# Generate a cloud-init compatible password hash from plaintext.
generate_cloud_password_hash() {
  local plaintext_password="$1"

  if command -v mkpasswd >/dev/null 2>&1; then
    mkpasswd -m yescrypt "$plaintext_password" 2>/dev/null && return 0
  fi

  if command -v openssl >/dev/null 2>&1; then
    openssl passwd -6 "$plaintext_password" 2>/dev/null && return 0
  fi

  return 1
}

# Resolve Fresh_Install template directory when available.
resolve_fresh_install_template_dir() {
  local script_dir
  script_dir="$(dirname "$(readlink -f "$0")")"

  local -a candidates=()
  candidates+=("${FRESH_INSTALL_DIR:-}")
  candidates+=("${script_dir}/../Fresh_Install")
  candidates+=("${script_dir}/Fresh_Install")
  candidates+=("$(pwd)/../Fresh_Install")
  if [[ -n "${SUDO_USER:-}" ]]; then
    candidates+=("/home/${SUDO_USER}/My_Scripts_Apps/Raspberry_Pi/Fresh_Install")
  fi
  candidates+=("/home/admin/My_Scripts_Apps/Raspberry_Pi/Fresh_Install")
  candidates+=("/root/My_Scripts_Apps/Raspberry_Pi/Fresh_Install")

  local candidate
  for candidate in "${candidates[@]}"; do
    [[ -n "$candidate" ]] || continue
    if [[ -f "$candidate/cmdline.txt" && -f "$candidate/config.txt" ]]; then
      echo "$candidate"
      return 0
    fi
  done

  return 1
}

prompt_fresh_flash_identity() {
  local confirm=""

  FLASH_WIFI_SSID=""
  FLASH_WIFI_PASSWORD=""
  FLASH_USERNAME=""
  FLASH_USER_PASSWORD=""
  FLASH_HOSTNAME=""

  echo
  print_box "Mode 7 First-Boot Personalization" \
    "" \
    "Enter plaintext values for first boot provisioning:" \
    "  - WiFi SSID/password" \
    "  - Username/password" \
    "  - Hostname" \
    ""

  while true; do
    read -rp "WiFi SSID: " FLASH_WIFI_SSID
    [[ -n "$FLASH_WIFI_SSID" ]] && break
    log_warning "WiFi SSID cannot be empty"
  done

  while true; do
    read -rsp "WiFi password: " FLASH_WIFI_PASSWORD
    echo
    [[ -n "$FLASH_WIFI_PASSWORD" ]] || { log_warning "WiFi password cannot be empty"; continue; }

    read -rsp "Confirm WiFi password: " confirm
    echo
    [[ "$FLASH_WIFI_PASSWORD" == "$confirm" ]] && break
    log_warning "WiFi passwords do not match - try again"
  done

  while true; do
    read -rp "Username: " FLASH_USERNAME
    if ! is_valid_linux_username "$FLASH_USERNAME"; then
      log_warning "Invalid username. Use lowercase letters, digits, _ or -, and start with a letter/_"
      continue
    fi
    if [[ "$FLASH_USERNAME" == "root" ]]; then
      log_warning "Username 'root' is not allowed for this workflow"
      continue
    fi
    break
  done

  while true; do
    read -rsp "User password: " FLASH_USER_PASSWORD
    echo
    [[ -n "$FLASH_USER_PASSWORD" ]] || { log_warning "User password cannot be empty"; continue; }

    read -rsp "Confirm user password: " confirm
    echo
    [[ "$FLASH_USER_PASSWORD" == "$confirm" ]] && break
    log_warning "User passwords do not match - try again"
  done

  while true; do
    read -rp "Hostname: " FLASH_HOSTNAME
    if ! is_valid_hostname_label "$FLASH_HOSTNAME"; then
      log_warning "Invalid hostname label (letters, digits, hyphen; cannot start/end with hyphen)"
      continue
    fi
    break
  done

  echo
  log_info "First-boot settings queued for flash target:"
  log_info "  - Hostname: ${FLASH_HOSTNAME}"
  log_info "  - Username: ${FLASH_USERNAME}"
  log_info "  - WiFi SSID: ${FLASH_WIFI_SSID}"
}

write_default_fresh_cmdline_template() {
  local output_file="$1"
  printf '%s\n' 'console=serial0,115200 console=tty1 root=PARTUUID=7d23366b-02 rootfstype=ext4 fsck.repair=yes rootwait resize cfg80211.ieee80211_regdom=US' > "$output_file"
}

write_default_fresh_config_template() {
  local output_file="$1"
  cat > "$output_file" <<'EOF'
# For more options and information see
# http://rptl.io/configtxt
# Some settings may impact device functionality. See link above for details

# Uncomment some or all of these to enable the optional hardware interfaces
#dtparam=i2c_arm=on
#dtparam=i2s=on
#dtparam=spi=on

# Enable audio (loads snd_bcm2835)
dtparam=audio=on

# Additional overlays and parameters are documented
# /boot/firmware/overlays/README

# Automatically load overlays for detected cameras
camera_auto_detect=1

# Automatically load overlays for detected DSI displays
display_auto_detect=1

# Automatically load initramfs files, if found
auto_initramfs=1

# Enable DRM VC4 V3D driver
dtoverlay=vc4-kms-v3d
max_framebuffers=2

# Don't have the firmware create an initial video= setting in cmdline.txt.
# Use the kernel's default instead.
disable_fw_kms_setup=1

# Run in 64-bit mode
arm_64bit=1

# Disable compensation for displays with overscan
disable_overscan=1

# Run as fast as firmware / board allows
arm_boost=1

[cm4]
# Enable host mode on the 2711 built-in XHCI USB controller.
# This line should be removed if the legacy DWC2 controller is required
# (e.g. for USB device mode) or if USB support is not required.
otg_mode=1

[cm5]
dtoverlay=dwc2,dr_mode=host

[all]
EOF
}

write_default_fresh_network_template() {
  local output_file="$1"
  local wifi_ssid_escaped="$2"
  local wifi_password_escaped="$3"

  cat > "$output_file" <<EOF
network:
  version: 2
  ethernets:
    eth0:
      dhcp4: true
      dhcp6: true
      optional: true
  wifis:
    wlan0:
      dhcp4: true
      regulatory-domain: "US"
      access-points:
        "${wifi_ssid_escaped}":
          password: "${wifi_password_escaped}"
      optional: true
EOF
}

write_default_fresh_user_data_template() {
  local output_file="$1"
  local hostname="$2"
  local username="$3"
  local password_hash_escaped="$4"

  cat > "$output_file" <<EOF
#cloud-config
manage_resolv_conf: false
hostname: ${hostname}
manage_etc_hosts: true
packages:
- avahi-daemon
apt:
  preserve_sources_list: true
  conf: |
    Acquire {
      Check-Date "false";
    };
timezone: America/Chicago
keyboard:
  model: pc105
  layout: "us"
users:
- name: ${username}
  groups: users,adm,dialout,audio,netdev,video,plugdev,cdrom,games,input,gpio,spi,i2c,render,sudo
  shell: /bin/bash
  lock_passwd: false
  passwd: "${password_hash_escaped}"
enable_ssh: true
ssh_pwauth: true
rpi:
  interfaces:
    serial: true
EOF
}

render_network_from_template() {
  local template_file="$1"
  local output_file="$2"
  local wifi_ssid_escaped="$3"
  local wifi_password_escaped="$4"

  local ssid_done=false
  local password_done=false
  local in_access_points=false

  : > "$output_file"
  while IFS= read -r line || [[ -n "$line" ]]; do
    if [[ "$line" =~ ^[[:space:]]*access-points:[[:space:]]*$ ]]; then
      in_access_points=true
      printf '%s\n' "$line" >> "$output_file"
      continue
    fi

    if [[ "$in_access_points" == "true" && "$ssid_done" == "false" && "$line" =~ ^([[:space:]]*)\".*\":[[:space:]]*$ ]]; then
      printf '%s"%s":\n' "${BASH_REMATCH[1]}" "$wifi_ssid_escaped" >> "$output_file"
      ssid_done=true
      continue
    fi

    if [[ "$in_access_points" == "true" && "$password_done" == "false" && "$line" =~ ^([[:space:]]*)password:[[:space:]]*.*$ ]]; then
      printf '%spassword: "%s"\n' "${BASH_REMATCH[1]}" "$wifi_password_escaped" >> "$output_file"
      password_done=true
      continue
    fi

    printf '%s\n' "$line" >> "$output_file"
  done < "$template_file"

  [[ "$ssid_done" == "true" && "$password_done" == "true" ]]
}

render_user_data_from_template() {
  local template_file="$1"
  local output_file="$2"
  local hostname="$3"
  local username="$4"
  local password_hash_escaped="$5"

  local hostname_done=false
  local username_done=false
  local password_done=false

  : > "$output_file"
  while IFS= read -r line || [[ -n "$line" ]]; do
    if [[ "$hostname_done" == "false" && "$line" =~ ^hostname:[[:space:]]*.*$ ]]; then
      printf 'hostname: %s\n' "$hostname" >> "$output_file"
      hostname_done=true
      continue
    fi

    if [[ "$username_done" == "false" && "$line" =~ ^([[:space:]]*)-[[:space:]]name:[[:space:]]*.*$ ]]; then
      printf '%s- name: %s\n' "${BASH_REMATCH[1]}" "$username" >> "$output_file"
      username_done=true
      continue
    fi

    if [[ "$password_done" == "false" && "$line" =~ ^([[:space:]]*)passwd:[[:space:]]*.*$ ]]; then
      printf '%spasswd: "%s"\n' "${BASH_REMATCH[1]}" "$password_hash_escaped" >> "$output_file"
      password_done=true
      continue
    fi

    printf '%s\n' "$line" >> "$output_file"
  done < "$template_file"

  [[ "$hostname_done" == "true" && "$username_done" == "true" && "$password_done" == "true" ]]
}

# Apply Fresh_Install-formatted first-boot files to a mounted boot partition.
# Usage (via with_mount): apply_fresh_install_boot_files "/mnt/boot" "/dev/sda1" "root_partuuid" "ssid" "wifi_pass" "username" "user_pass" "hostname"
apply_fresh_install_boot_files() {
  local boot_mount_point="$1"
  local mounted_boot_device="$2"
  local root_partuuid="$3"
  local wifi_ssid="$4"
  local wifi_password="$5"
  local username="$6"
  local user_password="$7"
  local hostname="$8"

  local template_dir=""
  local password_hash=""
  local wifi_ssid_escaped=""
  local wifi_password_escaped=""
  local password_hash_escaped=""

  [[ -n "$root_partuuid" ]] || { log_error "Root PARTUUID is empty"; return 1; }
  [[ -f "$boot_mount_point/cmdline.txt" ]] || { log_error "cmdline.txt missing on ${mounted_boot_device}"; return 1; }

  if ! check_writable "$boot_mount_point"; then
    log_error "Boot partition is not writable: ${mounted_boot_device}"
    return 1
  fi

  if ! password_hash=$(generate_cloud_password_hash "$user_password"); then
    log_error "Failed to generate password hash (install whois for mkpasswd or openssl)"
    return 1
  fi

  template_dir="$(resolve_fresh_install_template_dir || true)"
  if [[ -n "$template_dir" ]]; then
    log_info "Using Fresh_Install templates from: ${template_dir}"
  else
    log_warning "Fresh_Install directory not found - using embedded template defaults"
  fi

  if [[ -n "$template_dir" && -f "$template_dir/cmdline.txt" ]]; then
    cp "$template_dir/cmdline.txt" "$boot_mount_point/cmdline.txt" || return 1
  else
    write_default_fresh_cmdline_template "$boot_mount_point/cmdline.txt" || return 1
  fi

  if [[ -n "$template_dir" && -f "$template_dir/config.txt" ]]; then
    cp "$template_dir/config.txt" "$boot_mount_point/config.txt" || return 1
  else
    write_default_fresh_config_template "$boot_mount_point/config.txt" || return 1
  fi

  if ! update_cmdline_root "$boot_mount_point/cmdline.txt" "$root_partuuid" "PARTUUID"; then
    return 1
  fi

  if ! add_usb_power_settings "$boot_mount_point/config.txt"; then
    log_warning "Failed to apply USB power settings to config.txt"
  fi

  wifi_ssid_escaped="$(yaml_escape_double_quotes "$wifi_ssid")"
  wifi_password_escaped="$(yaml_escape_double_quotes "$wifi_password")"
  password_hash_escaped="$(yaml_escape_double_quotes "$password_hash")"

  if [[ -n "$template_dir" && -f "$template_dir/network-config" ]]; then
    if ! render_network_from_template "$template_dir/network-config" "$boot_mount_point/network-config" "$wifi_ssid_escaped" "$wifi_password_escaped"; then
      log_warning "Could not render network-config from template - using built-in fallback"
      write_default_fresh_network_template "$boot_mount_point/network-config" "$wifi_ssid_escaped" "$wifi_password_escaped"
    fi
  else
    write_default_fresh_network_template "$boot_mount_point/network-config" "$wifi_ssid_escaped" "$wifi_password_escaped"
  fi

  if [[ -n "$template_dir" && -f "$template_dir/user-data" ]]; then
    if ! render_user_data_from_template "$template_dir/user-data" "$boot_mount_point/user-data" "$hostname" "$username" "$password_hash_escaped"; then
      log_warning "Could not render user-data from template - using built-in fallback"
      write_default_fresh_user_data_template "$boot_mount_point/user-data" "$hostname" "$username" "$password_hash_escaped"
    fi
  else
    write_default_fresh_user_data_template "$boot_mount_point/user-data" "$hostname" "$username" "$password_hash_escaped"
  fi

  chmod 644 "$boot_mount_point/network-config" "$boot_mount_point/user-data" 2>/dev/null || true
  touch "$boot_mount_point/ssh" 2>/dev/null || true

  sync
  log_success "Applied Fresh_Install-format files on ${mounted_boot_device}"
  return 0
}

# Prompt for option 2 target connectivity settings.
# Populates OPTION2_SSH_ENABLED/OPTION2_WIFI_SSID/OPTION2_WIFI_PASSWORD globals.
OPTION2_SSH_ENABLED=""
OPTION2_WIFI_SSID=""
OPTION2_WIFI_PASSWORD=""

prompt_option2_target_connectivity_settings() {
  local confirm=""
  local ssh_choice=""

  OPTION2_SSH_ENABLED="true"
  OPTION2_WIFI_SSID=""
  OPTION2_WIFI_PASSWORD=""

  echo
  print_box "Mode 2 Target Boot Connectivity" \
    "" \
    "This updates pre-boot connectivity seed files on the target boot drive:" \
    "  - ssh/ssh.txt flag file" \
    "  - network-config (WiFi SSID/password)" \
    "  - user-data (enable_ssh / ssh_pwauth)" \
    ""

  while true; do
    read -rp "Enable SSH on target first boot? [Y/n]: " ssh_choice
    ssh_choice="${ssh_choice,,}"
    ssh_choice="${ssh_choice//[[:space:]]/}"
    case "$ssh_choice" in
      ""|y|yes)
        OPTION2_SSH_ENABLED="true"
        break
        ;;
      n|no)
        OPTION2_SSH_ENABLED="false"
        break
        ;;
      *)
        log_warning "Please answer y or n"
        ;;
    esac
  done

  while true; do
    read -rp "WiFi SSID for target first boot: " OPTION2_WIFI_SSID
    [[ -n "$OPTION2_WIFI_SSID" ]] && break
    log_warning "WiFi SSID cannot be empty"
  done

  while true; do
    read -rsp "WiFi password for target first boot: " OPTION2_WIFI_PASSWORD
    echo
    [[ -n "$OPTION2_WIFI_PASSWORD" ]] || { log_warning "WiFi password cannot be empty"; continue; }

    read -rsp "Confirm WiFi password: " confirm
    echo
    [[ "$OPTION2_WIFI_PASSWORD" == "$confirm" ]] && break
    log_warning "WiFi passwords do not match - try again"
  done

  log_info "Target connectivity settings captured (SSH=${OPTION2_SSH_ENABLED}, SSID=${OPTION2_WIFI_SSID})"
}

# Update a top-level cloud-init YAML boolean key or append if missing.
# Usage: set_cloud_init_bool_key "/path/user-data" "enable_ssh" "true"
set_cloud_init_bool_key() {
  local yaml_file="$1"
  local key="$2"
  local bool_value="$3"

  [[ -f "$yaml_file" ]] || return 1

  if grep -qE "^${key}:[[:space:]]*(true|false)" "$yaml_file" 2>/dev/null; then
    sed -i -E "s#^${key}:[[:space:]]*(true|false).*#${key}: ${bool_value}#" "$yaml_file" || return 1
  elif grep -qE "^${key}:" "$yaml_file" 2>/dev/null; then
    sed -i -E "s#^${key}:.*#${key}: ${bool_value}#" "$yaml_file" || return 1
  else
    printf '%s\n' "${key}: ${bool_value}" >> "$yaml_file" || return 1
  fi

  return 0
}

# Apply option 2 connectivity settings to a mounted target boot partition.
# Usage (via with_mount): apply_option2_target_connectivity_settings "/mnt/boot" "/dev/sda1" "true|false" "ssid" "wifi_password"
apply_option2_target_connectivity_settings() {
  local boot_mount_point="$1"
  local mounted_boot_device="$2"
  local ssh_enabled="$3"
  local wifi_ssid="$4"
  local wifi_password="$5"

  local template_dir=""
  local source_network=""
  local source_user_data=""
  local tmp_network=""
  local tmp_user_data=""
  local wifi_ssid_escaped=""
  local wifi_password_escaped=""

  [[ -n "$wifi_ssid" && -n "$wifi_password" ]] || {
    log_error "WiFi SSID/password are required for target connectivity update"
    return 1
  }

  if ! check_writable "$boot_mount_point"; then
    log_error "Boot partition is not writable: ${mounted_boot_device}"
    return 1
  fi

  template_dir="$(resolve_fresh_install_template_dir || true)"
  wifi_ssid_escaped="$(yaml_escape_double_quotes "$wifi_ssid")"
  wifi_password_escaped="$(yaml_escape_double_quotes "$wifi_password")"

  source_network="$boot_mount_point/network-config"
  if [[ ! -f "$source_network" && -n "$template_dir" && -f "$template_dir/network-config" ]]; then
    source_network="$template_dir/network-config"
  fi

  if [[ -f "$source_network" ]]; then
    tmp_network=$(mktemp /tmp/option2-network-XXXXXX)
    if render_network_from_template "$source_network" "$tmp_network" "$wifi_ssid_escaped" "$wifi_password_escaped"; then
      cp "$tmp_network" "$boot_mount_point/network-config" || { rm -f "$tmp_network"; return 1; }
      rm -f "$tmp_network"
    else
      rm -f "$tmp_network"
      write_default_fresh_network_template "$boot_mount_point/network-config" "$wifi_ssid_escaped" "$wifi_password_escaped" || return 1
    fi
  else
    write_default_fresh_network_template "$boot_mount_point/network-config" "$wifi_ssid_escaped" "$wifi_password_escaped" || return 1
  fi

  source_user_data="$boot_mount_point/user-data"
  if [[ ! -f "$source_user_data" && -n "$template_dir" && -f "$template_dir/user-data" ]]; then
    source_user_data="$template_dir/user-data"
  fi

  if [[ -f "$source_user_data" ]]; then
    tmp_user_data=$(mktemp /tmp/option2-userdata-XXXXXX)
    cp "$source_user_data" "$tmp_user_data" || { rm -f "$tmp_user_data"; return 1; }
    if ! set_cloud_init_bool_key "$tmp_user_data" "enable_ssh" "$ssh_enabled"; then
      rm -f "$tmp_user_data"
      return 1
    fi
    if [[ "$ssh_enabled" == "true" ]]; then
      set_cloud_init_bool_key "$tmp_user_data" "ssh_pwauth" "true" || { rm -f "$tmp_user_data"; return 1; }
    else
      set_cloud_init_bool_key "$tmp_user_data" "ssh_pwauth" "false" || { rm -f "$tmp_user_data"; return 1; }
    fi
    cp "$tmp_user_data" "$boot_mount_point/user-data" || { rm -f "$tmp_user_data"; return 1; }
    rm -f "$tmp_user_data"
  else
    cat > "$boot_mount_point/user-data" <<EOF
#cloud-config
enable_ssh: ${ssh_enabled}
ssh_pwauth: ${ssh_enabled}
EOF
  fi

  if [[ "$ssh_enabled" == "true" ]]; then
    touch "$boot_mount_point/ssh" 2>/dev/null || true
    rm -f "$boot_mount_point/ssh.txt" 2>/dev/null || true
  else
    rm -f "$boot_mount_point/ssh" "$boot_mount_point/ssh.txt" 2>/dev/null || true
  fi

  chmod 644 "$boot_mount_point/network-config" "$boot_mount_point/user-data" 2>/dev/null || true
  sync
  log_success "Updated target boot connectivity files on ${mounted_boot_device}"
  return 0
}

# Option 2 helper: prompt and apply pre-boot connectivity to next validated boot target.
offer_preboot_connectivity_update_to_next_target() {
  local next_boot="${BOOT_ORDER_NEXT_TARGET_BOOT:-}"
  local next_root="${BOOT_ORDER_NEXT_TARGET_ROOT:-}"

  if [[ -z "$next_boot" || -z "$next_root" ]]; then
    log_info "No next non-current boot target available for pre-boot connectivity update"
    return 0
  fi

  echo
  log_info "Next boot target candidate: ${next_boot} -> ${next_root}"

  if ! prompt_yn "Configure pre-boot SSH/WiFi files on this boot target now?"; then
    log_info "Skipped pre-boot target connectivity update"
    return 0
  fi

  # These files are consumed at first boot; warn if target already appears to have completed cloud-init.
  if ! is_target_ssh_continuity_pending_first_boot "$next_boot" "$next_root"; then
    log_warning "Target does not appear to be in first-boot cloud-init state"
    log_warning "Editing boot seed files may not change an already-initialized root filesystem"
    if ! prompt_yn "Continue updating boot seed files anyway?"; then
      log_info "Skipped pre-boot target connectivity update"
      return 0
    fi
  fi

  prompt_option2_target_connectivity_settings

  if ! with_mount "$next_boot" apply_option2_target_connectivity_settings --verify-boot \
      "$OPTION2_SSH_ENABLED" "$OPTION2_WIFI_SSID" "$OPTION2_WIFI_PASSWORD"; then
    log_warning "Failed to update pre-boot connectivity files on ${next_boot}"
    OPTION2_WIFI_PASSWORD=""
    return 1
  fi

  OPTION2_WIFI_PASSWORD=""
  log_success "Pre-boot connectivity update complete for ${next_boot}"
  return 0
}


################################################################################
# MAIN OPERATIONS
################################################################################

# Find current root partition and PARTUUID
CURRENT_ROOT="$(findmnt -no SOURCE /)"
CURRENT_PARTUUID="$(get_partuuid_safe "$CURRENT_ROOT" "warn" 2>/dev/null || true)"

# Show log location
echo "$(timestamp) [i] Logging to: $LOG_FILE"
echo

################################################################################
# EEPROM BOOT ORDER HELPERS
################################################################################

eeprom_tool_available() {
  command -v rpi-eeprom-config >/dev/null 2>&1
}

normalize_boot_order() {
  local order="$1"
  order="${order,,}"
  order="${order//[[:space:]]/}"
  if [[ -z "$order" ]]; then
    echo ""
    return
  fi
  if [[ "$order" =~ ^0x ]]; then
    echo "$order"
  else
    echo "0x${order}"
  fi
}

eeprom_numeric_to_decimal() {
  local raw="${1:-}"
  raw="${raw,,}"
  raw="${raw//[[:space:]]/}"

  if [[ -z "$raw" ]]; then
    return 1
  fi

  if [[ "$raw" =~ ^0x[0-9a-f]+$ ]]; then
    echo "$((16#${raw#0x}))"
    return 0
  fi

  if [[ "$raw" =~ ^[0-9]+$ ]]; then
    echo "$raw"
    return 0
  fi

  return 1
}

boot_order_includes_usb_msd() {
  local order="${1#0x}"
  order="${order,,}"
  local i

  for (( i=${#order}-1; i>=0; i-- )); do
    local digit="${order:$i:1}"
    case "$digit" in
      e|f)
        break
        ;;
      4|5)
        return 0
        ;;
    esac
  done

  return 1
}

eeprom_upsert_setting() {
  local conf_file="$1"
  local key="$2"
  local value="$3"

  if grep -q "^${key}=" "$conf_file"; then
    sed -i -E "s#^${key}=.*#${key}=${value}#" "$conf_file"
  else
    echo "${key}=${value}" >> "$conf_file"
  fi
}

get_eeprom_config_value() {
  local key="$1"
  if ! eeprom_tool_available; then
    echo ""
    return 1
  fi

  rpi-eeprom-config 2>/dev/null | awk -F= -v k="$key" '$1==k {print $2; exit}' | sed 's/#.*//; s/[[:space:]]*$//'
}

validate_usb_boot_eeprom_profile() {
  local order="$1"
  local context_label="${2:-EEPROM USB boot profile}"

  if ! boot_order_includes_usb_msd "$order"; then
    return 0
  fi

  local discover_raw=""
  local pwr_raw=""
  local discover_dec=""
  local pwr_dec=""
  local failed=0

  discover_raw="$(get_eeprom_config_value "USB_MSD_DISCOVER_TIMEOUT" || true)"
  pwr_raw="$(get_eeprom_config_value "USB_MSD_PWR_OFF_TIME" || true)"

  if ! discover_dec="$(eeprom_numeric_to_decimal "$discover_raw" 2>/dev/null || true)" || [[ -z "$discover_dec" ]]; then
    log_error "${context_label}: USB_MSD_DISCOVER_TIMEOUT is missing or invalid"
    failed=1
  elif (( discover_dec < EEPROM_USB_DISCOVER_TIMEOUT_TARGET )); then
    log_error "${context_label}: USB_MSD_DISCOVER_TIMEOUT=${discover_raw} is below required ${EEPROM_USB_DISCOVER_TIMEOUT_TARGET}"
    failed=1
  else
    log_success "${context_label}: USB_MSD_DISCOVER_TIMEOUT=${discover_raw}"
  fi

  if ! pwr_dec="$(eeprom_numeric_to_decimal "$pwr_raw" 2>/dev/null || true)" || [[ -z "$pwr_dec" ]]; then
    log_error "${context_label}: USB_MSD_PWR_OFF_TIME is missing or invalid"
    failed=1
  elif (( pwr_dec < EEPROM_USB_PWR_OFF_TIME_TARGET )); then
    log_error "${context_label}: USB_MSD_PWR_OFF_TIME=${pwr_raw} is below required ${EEPROM_USB_PWR_OFF_TIME_TARGET}"
    failed=1
  else
    log_success "${context_label}: USB_MSD_PWR_OFF_TIME=${pwr_raw}"
  fi

  if (( failed > 0 )); then
    log_error "${context_label} failed USB boot hardening checks"
    return 1
  fi

  return 0
}

# Validate EEPROM BOOT_ORDER syntax and return normalized value on success.
# Usage: validate_eeprom_boot_order "0xf14" [context_label]
validate_eeprom_boot_order() {
  local order="$1"
  local context_label="${2:-EEPROM BOOT_ORDER}"
  local normalized

  normalized=$(normalize_boot_order "$order")
  local body="${normalized#0x}"

  if [[ -z "$body" ]]; then
    log_error "${context_label} is empty"
    return 1
  fi

  if [[ ! "$body" =~ ^[0-7ef]+$ ]]; then
    log_error "${context_label} contains invalid digits: ${normalized}"
    log_error "Allowed BOOT_ORDER digits are 0-7, e, and f"
    return 1
  fi

  if [[ ! "$body" =~ [ef] ]]; then
    log_error "${context_label} must include a terminal nibble (e or f): ${normalized}"
    return 1
  fi

  if (( ${#body} > 8 )); then
    log_error "${context_label} is too long (${#body} nibbles): ${normalized}"
    return 1
  fi

  echo "$normalized"
}

get_eeprom_boot_order() {
  if ! eeprom_tool_available; then
    echo ""
    return 1
  fi
  local order
  order=$(rpi-eeprom-config 2>/dev/null | awk -F= '/^BOOT_ORDER=/{print $2; exit}' | sed 's/#.*//')
  local normalized
  if ! normalized=$(validate_eeprom_boot_order "$order" "EEPROM BOOT_ORDER readback"); then
    return 1
  fi
  echo "$normalized"
}

boot_order_digit_label() {
  case "$1" in
    0) echo "SD card detect" ;;
    1) echo "SD card" ;;
    2) echo "Network" ;;
    3) echo "RPIBOOT" ;;
    4) echo "USB-MSD" ;;
    5) echo "BCM-USB-MSD" ;;
    6) echo "PCIe storage" ;;
    7) echo "HTTP" ;;
    e|E) echo "Stop" ;;
    f|F) echo "Restart" ;;
    *) echo "Unknown" ;;
  esac
}

describe_boot_order() {
  local order="${1#0x}"
  order="${order,,}"
  if [[ -z "$order" ]]; then
    echo "Unknown"
    return
  fi

  local i
  local labels=()
  for (( i=${#order}-1; i>=0; i-- )); do
    labels+=("$(boot_order_digit_label "${order:$i:1}")")
  done

  local out=""
  for i in "${!labels[@]}"; do
    if [[ $i -eq 0 ]]; then
      out="${labels[$i]}"
    else
      out+=" → ${labels[$i]}"
    fi
  done
  echo "$out"
}

apply_eeprom_boot_order() {
  local desired="$1"
  if ! desired=$(validate_eeprom_boot_order "$desired" "Requested EEPROM BOOT_ORDER"); then
    return 1
  fi

  local tmp_conf
  tmp_conf=$(mktemp /tmp/eeprom-bootconf-XXXXXX)

  if ! rpi-eeprom-config > "$tmp_conf" 2>/dev/null; then
    rm -f "$tmp_conf"
    return 1
  fi

  if grep -q '^BOOT_ORDER=' "$tmp_conf"; then
    sed -i -E "s/^BOOT_ORDER=.*/BOOT_ORDER=${desired}/" "$tmp_conf"
  else
    echo "BOOT_ORDER=${desired}" >> "$tmp_conf"
  fi

  if boot_order_includes_usb_msd "$desired"; then
    log_info "Enforcing EEPROM USB boot profile for BOOT_ORDER=${desired}"
    eeprom_upsert_setting "$tmp_conf" "USB_MSD_DISCOVER_TIMEOUT" "$EEPROM_USB_DISCOVER_TIMEOUT_TARGET"
    eeprom_upsert_setting "$tmp_conf" "USB_MSD_PWR_OFF_TIME" "$EEPROM_USB_PWR_OFF_TIME_TARGET"
  fi

  log_info "Applying EEPROM BOOT_ORDER=${desired}"
  if rpi-eeprom-config --apply "$tmp_conf"; then
    rm -f "$tmp_conf"
    local readback=""
    local verify_attempt=1

    while (( verify_attempt <= 5 )); do
      readback=$(get_eeprom_boot_order 2>/dev/null || true)
      if [[ "$readback" == "$desired" ]]; then
        log_success "EEPROM BOOT_ORDER validated: ${readback}"
        log_info "Boot order: $(describe_boot_order "$readback")"
        if ! validate_usb_boot_eeprom_profile "$readback" "EEPROM readback"; then
          log_error "EEPROM BOOT_ORDER applied, but USB boot profile validation failed"
          return 1
        fi
        log_info "Reboot required for EEPROM changes to take effect"
        return 0
      fi
      sleep 1
      verify_attempt=$((verify_attempt + 1))
    done

    # On many Pi models, EEPROM values do not reflect the staged update until reboot.
    # Treat a pending EEPROM update as a successful staging operation.
    local update_status=""
    update_status=$(rpi-eeprom-update 2>/dev/null || true)
    if [[ "$update_status" =~ [Uu]pdates?[[:space:]]+pending ]] || [[ -f "/boot/firmware/pieeprom.upd" ]]; then
      log_success "EEPROM update staged successfully (pending reboot)"
      if [[ -n "$readback" ]]; then
        log_info "Current readback before reboot: ${readback}"
      fi
      if boot_order_includes_usb_msd "$desired"; then
        log_warning "USB boot profile will be validated after reboot when staged EEPROM settings are active"
      fi
      log_info "Reboot to apply staged EEPROM update, then re-run option 1 or 8 to confirm BOOT_ORDER=${desired}"
      return 0
    fi

    if [[ -n "$readback" ]]; then
      log_error "EEPROM BOOT_ORDER verification mismatch"
      log_error "Expected: ${desired}"
      log_error "Readback: ${readback}"
    else
      log_error "Could not read back EEPROM BOOT_ORDER after apply"
    fi
    return 1
  fi

  rm -f "$tmp_conf"
  return 1
}

verify_or_prompt_eeprom_boot_order() {
  local desired="$1"
  local mode_label="$2"

  if ! desired=$(validate_eeprom_boot_order "$desired" "Desired EEPROM BOOT_ORDER (${mode_label})"); then
    log_warning "Skipping EEPROM update due to invalid desired BOOT_ORDER"
    return 1
  fi

  if ! eeprom_tool_available; then
    log_warning "EEPROM tools not available on this system"
    log_info "If your Pi supports EEPROM boot order, use: raspi-config → Advanced Options → Boot Order"
    return 0
  fi

  local current
  current=$(get_eeprom_boot_order)
  if [[ -z "$current" ]]; then
    log_warning "Unable to read EEPROM BOOT_ORDER"
    return 0
  fi

  if [[ "$current" == "$desired" ]]; then
    log_success "EEPROM BOOT_ORDER already matches ${mode_label} (${current})"
    log_info "Boot order: $(describe_boot_order "$current")"
    if ! validate_usb_boot_eeprom_profile "$current" "Current EEPROM"; then
      log_warning "Current EEPROM BOOT_ORDER matches, but USB boot profile is not hardened"
      if prompt_yn "Repair EEPROM USB boot profile now?" true; then
        if ! apply_eeprom_boot_order "$desired"; then
          log_warning "EEPROM USB boot profile repair failed - manual fix may be required"
        fi
      fi
    fi
    return 0
  fi

  log_warning "EEPROM BOOT_ORDER does not match ${mode_label}"
  log_info "Current: ${current} ($(describe_boot_order "$current"))"
  log_info "Desired: ${desired} ($(describe_boot_order "$desired"))"
  if prompt_yn "Update EEPROM boot order now?"; then
    if ! apply_eeprom_boot_order "$desired"; then
      log_warning "EEPROM update failed - you may need to adjust manually"
    fi
  else
    log_info "Skipped EEPROM update"
  fi
}

choose_eeprom_boot_order() {
  local desired_label="$1"
  case "$desired_label" in
    "SD+USB") echo "0xf41" ;;
    "USB+SD") echo "0xf14" ;;
    "SD Only") echo "0xf1" ;;
    "USB Only") echo "0xf4" ;;
    *) echo "" ;;
  esac
}

recommend_eeprom_for_current_boot() {
  local preference
  preference=$(get_boot_device_preference)
  if [[ "$preference" == "sd" ]]; then
    verify_or_prompt_eeprom_boot_order "0xf41" "SD+USB"
  elif [[ "$preference" == "usb" ]]; then
    verify_or_prompt_eeprom_boot_order "0xf14" "USB+SD"
  else
    log_warning "Could not determine boot device type to recommend EEPROM boot order"
    if prompt_yn "Open EEPROM boot order menu now?"; then
      configure_eeprom_boot_order
    fi
  fi
}

get_boot_device_preference() {
  local boot_dev
  boot_dev=$(findmnt -no SOURCE "$BOOT_DIR" 2>/dev/null || true)
  if [[ -z "$boot_dev" ]]; then
    echo "unknown"
    return
  fi

  local boot_disk
  boot_disk=$(lsblk -no PKNAME "$boot_dev" 2>/dev/null | head -n1)
  [[ -n "$boot_disk" ]] || { echo "unknown"; return; }

  if disk_matches_transport_class "$boot_disk" "sd"; then
    echo "sd"
  elif disk_matches_transport_class "$boot_disk" "usb" || disk_matches_transport_class "$boot_disk" "pcie"; then
    echo "usb"
  else
    echo "unknown"
  fi
}

################################################################################
# DEVICE INFO HELPERS (SHOW CURRENT)
################################################################################

print_storage_overview() {
  local root_dev="$CURRENT_ROOT"
  local boot_dev
  boot_dev="$(findmnt -no SOURCE "$BOOT_DIR" 2>/dev/null || true)"

  local root_disk=""
  local boot_disk=""
  root_disk="$(lsblk -no PKNAME "$root_dev" 2>/dev/null || true)"
  if [[ -n "$boot_dev" ]]; then
    boot_disk="$(lsblk -no PKNAME "$boot_dev" 2>/dev/null | head -n1)"
  fi

  echo
  log_info "Current root device: ${root_dev:-unknown} ${root_disk:+(disk: /dev/$root_disk)}"
  log_info "Current boot device: ${boot_dev:-unknown} ${boot_disk:+(disk: /dev/$boot_disk)}"
  echo
  lsblk -o NAME,SIZE,TRAN,TYPE,FSTYPE,LABEL,MOUNTPOINT,PARTUUID 2>/dev/null || true
}

print_device_tree() {
  local dev="$1"
  [[ -n "$dev" ]] || return 0
  lsblk -o NAME,SIZE,TRAN,TYPE,FSTYPE,LABEL,MOUNTPOINT,PARTUUID "$dev" 2>/dev/null || true
}

print_device_details() {
  local dev="$1"
  local title="$2"

  echo
  echo "=== $title ==="

  if [[ ! -b "$dev" ]]; then
    log_warning "Device not found: $dev"
    return 0
  fi

  lsblk -o NAME,SIZE,TRAN,TYPE,FSTYPE,LABEL,MOUNTPOINT,PARTUUID "$dev" 2>/dev/null || true
}

resolve_cmdline_root_device() {
  local root_value="$1"

  if [[ "$root_value" == PARTUUID=* ]]; then
    local partuuid
    partuuid="${root_value#PARTUUID=}"
    handle_partuuid_uniqueness "$partuuid" "cmdline root" "warn" || true
  fi

  resolve_device_identifier "$root_value"
}

show_current() {
  log_step "Current Boot Configuration"

  local cmdline_root
  local cmdline_root_value
  local cmdline_root_device

  cmdline_root="$(grep -o 'root=[^ ]*' "$CMDLINE" || true)"
  cmdline_root_value="${cmdline_root#root=}"
  cmdline_root_device="$(resolve_cmdline_root_device "$cmdline_root_value")"

  echo "Cmdline root setting: ${cmdline_root:-not found}"
  if [[ -n "$cmdline_root_device" ]]; then
    echo "Cmdline root device: $cmdline_root_device"
  fi

  echo "Live system root device: $CURRENT_ROOT (PARTUUID=${CURRENT_PARTUUID:-unknown})"

  local boot_device
  boot_device="$(findmnt -no SOURCE "$BOOT_DIR" 2>/dev/null || true)"
  if [[ -n "$boot_device" ]]; then
    echo "Boot partition device: $boot_device"
  else
    log_warning "Unable to determine boot partition device"
  fi

  echo
  log_info "Mounted filesystems (root + boot):"
  findmnt / "$BOOT_DIR" 2>/dev/null || true

  if [[ -f /etc/fstab ]]; then
    echo
    log_info "fstab entries (root + boot):"
    grep -E '^[^#].*\s+(/boot/firmware|/)\s+' /etc/fstab || true
  fi

  if [[ -n "$boot_device" ]]; then
    print_device_details "$boot_device" "Boot Partition Details"
    local boot_disk
    boot_disk=$(lsblk -no PKNAME "$boot_device" 2>/dev/null | head -n1)
    [[ -n "$boot_disk" ]] && print_device_details "/dev/$boot_disk" "Boot Disk Details"
  fi

  print_device_details "$CURRENT_ROOT" "Root Partition Details"
  local root_disk
  root_disk=$(lsblk -no PKNAME "$CURRENT_ROOT" 2>/dev/null | head -n1)
  [[ -n "$root_disk" ]] && print_device_details "/dev/$root_disk" "Root Disk Details"

  if [[ -n "$cmdline_root_device" && "$cmdline_root_device" != "$CURRENT_ROOT" ]]; then
    print_device_details "$cmdline_root_device" "Cmdline Root Device Details"
  fi

  echo
  if eeprom_tool_available; then
    log_info "EEPROM bootloader configuration:"
    rpi-eeprom-config 2>/dev/null || true

    local current_boot_order
    current_boot_order=$(get_eeprom_boot_order)
    if [[ -n "$current_boot_order" ]]; then
      echo
      log_info "EEPROM BOOT_ORDER: ${current_boot_order}"
      log_info "Boot order: $(describe_boot_order "$current_boot_order")"
      if ! validate_usb_boot_eeprom_profile "$current_boot_order" "Current EEPROM"; then
        log_warning "EEPROM USB boot profile is not hardened for USB boot"
      fi
    else
      log_warning "EEPROM BOOT_ORDER not found"
    fi
  else
    log_info "EEPROM tools not available on this system"
  fi
}

configure_eeprom_boot_order() {
  log_step "EEPROM Boot Order"

  if ! eeprom_tool_available; then
    log_warning "EEPROM tools not available on this system"
    log_info "If your Pi supports EEPROM boot order, use: raspi-config → Advanced Options → Boot Order"
    return 1
  fi

  local current
  current=$(get_eeprom_boot_order)
  if [[ -n "$current" ]]; then
    log_info "Current BOOT_ORDER: ${current}"
    log_info "Boot order: $(describe_boot_order "$current")"
  else
    log_warning "Current BOOT_ORDER not found"
  fi

  echo
  echo "Select EEPROM boot order:"
  echo "  1) SD+USB (0xf41)"
  echo "  2) USB+SD (0xf14)"
  echo "  3) SD Only (0xf1)"
  echo "  4) USB Only (0xf4)"
  echo "  q) Cancel"
  read -rp "Enter choice: " sel
  sel="${sel,,}"

  local desired_label=""
  case "$sel" in
    1) desired_label="SD+USB" ;;
    2) desired_label="USB+SD" ;;
    3) desired_label="SD Only" ;;
    4) desired_label="USB Only" ;;
    q) log_info "Cancelled"; return 0 ;;
    *) err "Invalid selection" ;;
  esac

  local desired
  desired=$(choose_eeprom_boot_order "$desired_label")
  [[ -n "$desired" ]] || err "Invalid BOOT_ORDER selection"
  desired=$(validate_eeprom_boot_order "$desired" "Selected EEPROM BOOT_ORDER") || err "Invalid BOOT_ORDER value"

  if [[ -n "$current" && "$current" == "$desired" ]]; then
    log_success "EEPROM BOOT_ORDER already set to ${desired} (${desired_label})"
    return 0
  fi

  log_info "Selected BOOT_ORDER=${desired} (${desired_label})"

  # Ensure SSH continuity across all bootable devices in the new order
  if [[ "$REQUIRE_SSH_CONTINUITY" == "true" ]]; then
    if ! check_boot_order_ssh_continuity "$desired"; then
      err "SSH continuity check failed for BOOT_ORDER ${desired}"
    fi
  else
    check_boot_order_ssh_continuity "$desired" || log_warning "SSH continuity check failed for BOOT_ORDER ${desired}"
  fi

  if prompt_yn "Apply EEPROM boot order now?" true; then
    if ! apply_eeprom_boot_order "$desired"; then
      log_warning "EEPROM update failed - you may need to adjust manually"
    fi
  else
    log_info "Skipped EEPROM update"
  fi
}

switch_current_only() {
  preflight_boot_change "boot configuration change" --ensure-current-ssh
  
  # Ensure backup exists before modifying
  ensure_cmdline_backup
  
  # Set cmdline.txt to boot and run from current device (SD, USB, etc.)
  CURRENT_PARTUUID_FRESH="$(read_required_partuuid_checked "$CURRENT_ROOT" "current root" "$PARTUUID_DUP_MODE" || true)"
  [[ -n "$CURRENT_PARTUUID_FRESH" ]] || err "Could not determine PARTUUID of current root"

  BOOT_DEV="$(findmnt -no SOURCE "$BOOT_DIR")"
  BOOT_PARTUUID="$(read_required_partuuid_checked "$BOOT_DEV" "boot device" "$PARTUUID_DUP_MODE" || true)"
  [[ -n "$BOOT_PARTUUID" ]] || err "Could not determine PARTUUID of boot device $BOOT_DEV"

  print_change_plan "current-device-only" "$CMDLINE" "$CURRENT_PARTUUID_FRESH" "$BOOT_PARTUUID" "/etc/fstab"
  
  # Ensure boot partition is configured (already mounted at $BOOT_DIR)
  if ! prepare_boot_partition_for_root "$BOOT_DIR" "current boot device" "$CURRENT_PARTUUID_FRESH" "$CMDLINE"; then
    err "Failed to configure boot partition"
  fi
  
  # Update current fstab if needed
  FSTAB_PATH="/etc/fstab"
  if [[ -f "$FSTAB_PATH" ]]; then
    echo
    log_info "Current fstab entries (boot + root):"
    grep -E '^[^#].*\s+(/boot/firmware|/)\s+' "$FSTAB_PATH" || true

    echo
    log_info "Target devices for fstab update (current boot/root):"
    local device_list=()
    [[ -n "$BOOT_DEV" && -b "$BOOT_DEV" ]] && device_list+=("$BOOT_DEV")
    [[ -n "$CURRENT_ROOT" && -b "$CURRENT_ROOT" ]] && device_list+=("$CURRENT_ROOT")
    if (( ${#device_list[@]} > 0 )); then
      lsblk -o NAME,SIZE,TRAN,TYPE,FSTYPE,LABEL,MOUNTPOINT,PARTUUID "${device_list[@]}" 2>/dev/null || true
    else
      log_warning "Unable to determine boot/root devices for display"
    fi
    log_info "These devices will be used for updated fstab entries."

    if prompt_yn "Update /etc/fstab to match current boot configuration?"; then
      BOOT_PARTUUID="$(read_required_partuuid_checked "$BOOT_DEV" "boot device" "$PARTUUID_DUP_MODE" "fail" || true)"
      CURRENT_PARTUUID_FRESH="$(read_required_partuuid_checked "$CURRENT_ROOT" "current root" "$PARTUUID_DUP_MODE" "fail" || true)"

      if [[ -n "$BOOT_PARTUUID" && -n "$CURRENT_PARTUUID_FRESH" ]]; then
        if update_fstab "/" "$BOOT_PARTUUID" "$CURRENT_PARTUUID_FRESH"; then
          log_success "Updated fstab with boot=PARTUUID=$BOOT_PARTUUID, root=PARTUUID=$CURRENT_PARTUUID_FRESH"
        else
          log_warning "fstab update failed"
        fi
      else
        log_warning "Could not determine PARTUUIDs, fstab not updated"
      fi
    fi
  fi
  
  echo
  log_success "Current device configuration complete!"
  log_info "Boot and root are both on the current device."
  log_info "Scripts available at: /boot/firmware/"

  recommend_eeprom_for_current_boot
  
  # Ensure all USB devices are also configured
  ensure_all_usb_boot_partitions
}

switch_hybrid() {
  preflight_boot_change "boot configuration change" --ensure-current-ssh
  
  log_step "Hybrid Boot Setup"
  log_info "Boot from current device, root from another device"
  echo
  log_info "Scanning for candidate root partitions (ext4, not current root)..."
  # Grab NAME, FSTYPE, MOUNTPOINT, SIZE, PARTUUID and sort by NAME
  mapfile -t CANDIDATES < <(
    lsblk -ln -o NAME,FSTYPE,MOUNTPOINT,SIZE,PARTUUID | \
    awk -v cur="$(basename "$CURRENT_ROOT")" '$2=="ext4" && $1!=cur {print $0}' | \
    sort -k1,1
  )

  if (( ${#CANDIDATES[@]} == 0 )); then
    err "No candidate ext4 partitions found"
  fi

  echo "Select a root partition:"
  i=1
  for line in "${CANDIDATES[@]}"; do
    # Columns: NAME FSTYPE MOUNTPOINT SIZE PARTUUID
    NAME=$(awk '{print $1}' <<<"$line")
    FSTYPE=$(awk '{print $2}' <<<"$line")
    MOUNTPOINT=$(awk '{print $3}' <<<"$line")
    SIZE=$(awk '{print $4}' <<<"$line")
    PARTUUID=$(awk '{print $5}' <<<"$line")

    [[ -z "$MOUNTPOINT" ]] && MOUNTPOINT="-"
    [[ -z "$PARTUUID" ]] && PARTUUID="-"

    echo "  $i) /dev/$NAME  size=$SIZE  mount=$MOUNTPOINT  fstype=$FSTYPE  PARTUUID=$PARTUUID"
    i=$((i + 1))
  done

  read -rp "Enter number: " sel
  if ! [[ "$sel" =~ ^[0-9]+$ ]] || (( sel < 1 || sel > ${#CANDIDATES[@]} )); then
    err "Invalid selection"
  fi

  ROOT_NAME=$(awk '{print $1}' <<<"${CANDIDATES[$((sel-1))]}")
  ROOT_DEV="/dev/$ROOT_NAME"
  
  # Disable swap if active on target
  if ! disable_swap_if_active "$ROOT_DEV"; then
    log_warning "Could not disable swap, mount may fail"
  fi
  
  # Ensure backup exists before modifying
  ensure_cmdline_backup
  
  # Get PARTUUIDs early to avoid race conditions
  PARTUUID="$(read_required_partuuid_checked "$ROOT_DEV" "$ROOT_DEV" "$PARTUUID_DUP_MODE" "fail" || true)"
  [[ -n "$PARTUUID" ]] || err "Could not determine PARTUUID of $ROOT_DEV"
  
  BOOT_DEV="$(findmnt -no SOURCE "$BOOT_DIR")"
  BOOT_PARTUUID="$(read_required_partuuid_checked "$BOOT_DEV" "boot device" "$PARTUUID_DUP_MODE" "fail" || true)"
  [[ -n "$BOOT_PARTUUID" ]] || err "Could not determine PARTUUID of boot device $BOOT_DEV"

  print_change_plan "hybrid boot" "$CMDLINE" "$PARTUUID" "$BOOT_PARTUUID" "${ROOT_DEV}:/etc/fstab"
  
  # Define callback to configure hybrid root
  configure_hybrid_root() {
    local root_mount="$1"
    
    # Check if writable
    if ! check_writable "$root_mount"; then
      log_error "Target partition is read-only - cannot configure"
      return 1
    fi
    
    echo
    log_info "Verifying target partition..."
    
    # Verify it's a valid root filesystem
    if ! verify_filesystem "$root_mount" "root"; then
      echo
      log_warning "Target partition does not appear to have a complete OS installation"
      if ! prompt_yn "  Continue anyway?"; then
        return 1
      fi
    else
      log_success "Target partition has valid OS installation"
    fi
    
    # Ensure SSH continuity on target root
    if ! ensure_ssh_continuity "$root_mount" "$ROOT_DEV" "nonfatal"; then
      log_error "SSH continuity check failed for ${ROOT_DEV}"
      return 1
    fi

    # Update fstab on target
    echo
    log_info "Updating fstab on target device..."
    if update_fstab "$root_mount" "$BOOT_PARTUUID" "$PARTUUID"; then
      log_success "Target fstab configured"
    else
      log_warning "fstab update failed"
    fi
    
    # Offer to sync configurations
    sync_system_configs "$root_mount" "$ROOT_DEV"
    
    return 0
  }
  
  # Execute hybrid configuration with automatic cleanup
  if ! with_mount "$ROOT_DEV" configure_hybrid_root; then
    err --debug "Failed to configure hybrid root"
  fi
  
  # Configure boot partition (BOOT_DEV is already mounted at $BOOT_DIR, configure directly)
  echo
  log_info "Configuring boot partition..."
  if ! prepare_boot_partition_for_root "$BOOT_DIR" "current boot device" "$PARTUUID" "$CMDLINE"; then
    err "Failed to configure boot partition"
  fi
  
  echo
  log_success "Hybrid configuration complete!"
  log_info "Current boot device will load root from: $ROOT_DEV"
  log_info "Reboot to test: sudo reboot"

  recommend_eeprom_for_current_boot
  
  # Ensure all USB devices are configured
  ensure_all_usb_boot_partitions
}

clone_os_to_usb_target() {
  preflight_boot_change "clone operation"
  
  log_step "Clone OS to USB Target"
  log_info "This will prepare a USB storage device with both boot and root partitions."
  echo
  
  # Select target USB storage device
  select_target_storage_device "Enter USB storage device name (whole disk, e.g., sda or nvme0n1): " "standard"
  
  echo
  log_info "Current partitions on $USB_DEV:"
  print_device_tree "$USB_DEV"
  echo
  log_info "This will:"
  log_info "  1. Create boot partition (~512MB, FAT32)"
  log_info "  2. Create root partition (remaining space, ext4)"
  log_info "  3. Clone current /boot/firmware to new boot partition"
  log_info "  4. Update cmdline.txt on USB storage boot partition"
  echo
  
  # Pre-flight checks before confirming
  log_info "Pre-flight checks:"
  if ! check_memory_available "$MIN_MEMORY_FLASH_MB"; then
    log_warning "  [!] WARNING: Low memory may cause issues"
  else
    log_success "Sufficient memory available"
  fi
  
  local boot_space_kb=$(df --output=avail "$BOOT_DIR" | tail -1)
  if (( boot_space_kb < 524288 )); then
    log_warning "  [!] WARNING: Boot partition may be too small to clone"
  else
    log_success "Boot partition has sufficient space"
  fi
  echo
  
  print_box "DESTRUCTIVE OPERATION WARNING" \
    "" \
    "This will COMPLETELY ERASE all data on: $USB_DEV" \
    "" \
    "Operations to be performed:" \
    "  1. Create new partition table (GPT)" \
    "  2. Create boot partition (~512MB, FAT32)" \
    "  3. Create root partition (remaining space, ext4)" \
    "  4. Clone /boot/firmware to new boot partition" \
    "  5. Update cmdline.txt for USB boot" \
    "" \
    "All existing partitions and data will be PERMANENTLY DELETED" \
    "" \
    "Device: $USB_DEV" \
    "Transport: USB" \
    ""
  
  read -rp "Type 'yes' to continue: " CONFIRM
  if [[ "${CONFIRM,,}" != "yes" ]]; then
    log_info "Aborted"
    return 0
  fi
  
  # Unmount any existing partitions
  unmount_device_partitions "$USB_NAME"
  
  # Partition the USB storage device
  log_success "Creating partitions on $USB_DEV"
  parted -s "$USB_DEV" mklabel gpt
  parted -s "$USB_DEV" mkpart primary fat32 1MiB 513MiB
  parted -s "$USB_DEV" set 1 boot on
  parted -s "$USB_DEV" mkpart primary ext4 513MiB 100%
  
  log_success "Triggering partition table rescan..."
  partprobe "$USB_DEV" 2>/dev/null || true
  udevadm settle 2>/dev/null || true
  
  # Regenerate PARTUUIDs to avoid conflicts with existing devices
  regenerate_partuuids "$USB_DEV" || log_warning "Could not regenerate PARTUUIDs"
  
  # Detect partition names
  USB_BOOT="$(get_partition_name "$USB_DEV" 1)"
  USB_ROOT="$(get_partition_name "$USB_DEV" 2)"
  
  # Wait for partitions to appear with timeout and notifications
  if ! wait_for_partitions "$USB_DEV" 2 15; then
    cleanup_on_error "$USB_DEV" "USB partition creation"
    err --debug "Partitions not detected after timeout"
  fi
  
  # Verify specific partitions exist
  if ! wait_for_device "$USB_BOOT" 5; then
    err --debug "Boot partition not detected: $USB_BOOT"
  fi
  if ! wait_for_device "$USB_ROOT" 5; then
    err --debug "Root partition not detected: $USB_ROOT"
  fi
  log_success "All partitions detected and ready"
  
  # Format partitions
  log_success "Formatting boot partition: $USB_BOOT"
  mkfs.vfat -F 32 -n BOOT "$USB_BOOT"
  
  log_success "Formatting root partition: $USB_ROOT"
  mkfs.ext4 -L rootfs "$USB_ROOT"
  
  # Wait for filesystem to be ready and re-read PARTUUIDs after formatting
  udevadm settle 2>/dev/null || true
  sleep 2
  
  local previous_dup_mode="$PARTUUID_DUP_MODE"
  PARTUUID_DUP_MODE="fix"

  # Get PARTUUIDs after regeneration and formatting (must be done after regenerate_partuuids)
  log_success "Reading PARTUUIDs after regeneration..."
  USB_BOOT_PARTUUID="$(read_required_partuuid_checked "$USB_BOOT" "$USB_BOOT" "$PARTUUID_DUP_MODE" "fail" || true)"
  [[ -n "$USB_BOOT_PARTUUID" ]] || err "Could not determine PARTUUID of $USB_BOOT"
  
  USB_ROOT_PARTUUID="$(read_required_partuuid_checked "$USB_ROOT" "$USB_ROOT" "$PARTUUID_DUP_MODE" "fail" || true)"
  [[ -n "$USB_ROOT_PARTUUID" ]] || err "Could not determine PARTUUID of $USB_ROOT"

  print_change_plan "clone OS to USB" "${USB_BOOT}:/cmdline.txt" "$USB_ROOT_PARTUUID" "$USB_BOOT_PARTUUID" "${USB_ROOT}:/etc/fstab"
  
  # Clone boot partition
  log_success "Preparing boot partition on $USB_BOOT..."
  if ! configure_boot_partition_for_target "$USB_BOOT" "$USB_ROOT_PARTUUID" "USB boot partition" "true"; then
    err --debug "Failed to configure USB boot partition"
  fi

  log_success "USB boot partition ready! Boot files copied and configured."
  log_info "Root partition is formatted but empty - starting root copy next."
  echo
  
  # Detect source type and show performance tip BEFORE operation
  local source_root=$(findmnt -n -o SOURCE /)
  local source_type="unknown"
  if [[ $source_root =~ mmcblk ]]; then
    source_type="SD card"
  elif [[ $source_root =~ sd[a-z] ]]; then
    source_type="USB device"
  fi
  
  # Show performance tip for USB-to-USB migrations
  if [[ "$source_type" == "USB device" ]]; then
    echo
    log_info "💡 Performance Note: You're migrating from USB to USB"
    log_info "   Alternative approach (faster, ~2 min vs 20 min):"
    log_info "   1. Flash fresh Raspberry Pi OS directly to target SSD using Pi Imager"
    log_info "   2. Boot from new SSD and run homelab-setup.sh to restore configs"
    log_info "   This avoids the slow USB-to-USB copy step entirely"
    echo
  fi
  
  log_success "Copying root filesystem from $source_type ($source_root) to $USB_ROOT"
  log_success "This may take 5-30 minutes depending on data size and source device speed..."
  echo
  
  # Define callback for root filesystem copy
  copy_root_filesystem() {
    local root_mount="$1"
    
    # Pre-flight checks
    log_success "Running pre-flight checks..."
    
    # Calculate required space (current usage + 1GB buffer)
    local used_space_mb=$(df -m / | tail -1 | awk '{print $3}')
    local required_mb=$((used_space_mb + 1024))
    log_success "Current system size: ${used_space_mb}MB, required space: ${required_mb}MB"
    
    if ! check_space_available "$root_mount" "$required_mb"; then
      return 1
    fi
    log_success "Sufficient space available"
    
    # Check available memory
    if ! check_memory_available "$MIN_MEMORY_COPY_MB"; then
      log_warning "  [!] Low memory detected - copy will proceed but may be slower"
    fi
    
    # Check writable
    if ! check_writable "$root_mount"; then
      return 1
    fi
    
    log_success "Pre-flight checks passed - starting filesystem copy..."
    echo
  
    # Build rsync command with exclude patterns
    local rsync_cmd=(rsync -aAXH --numeric-ids --info=progress2 --inplace --no-whole-file --delete)
    for exclude in "${RSYNC_EXCLUDES[@]}"; do
      rsync_cmd+=(--exclude="$exclude")
    done
    rsync_cmd+=(/ "$root_mount/")
    
    # Optimized rsync - exclude caches and temporary files
    if "${rsync_cmd[@]}"; then
      log_success "Root filesystem copied successfully"
      
      # Validate critical directories were copied
      log_success "Validating filesystem copy..."
      local missing_dirs=()
      for critical_dir in "/etc" "/usr" "/var" "/bin" "/sbin" "/lib"; do
        if [[ ! -d "${root_mount}${critical_dir}" ]]; then
          missing_dirs+=("$critical_dir")
        fi
      done
      
      if [[ ${#missing_dirs[@]} -gt 0 ]]; then
        log_error "Critical directories missing after copy: ${missing_dirs[*]}"
        log_error "Filesystem copy incomplete - cannot proceed"
        return 1
      fi
      log_success "Filesystem validation passed"
      
      # Ensure SSH continuity on copied root
      if ! ensure_ssh_continuity "$root_mount" "$USB_ROOT" "nonfatal"; then
        log_error "SSH continuity check failed for ${USB_ROOT}"
        return 1
      fi

      # Create fstab
      log_success "Updating fstab on USB root partition"
      local BOOT_PARTUUID
      BOOT_PARTUUID="$(read_required_partuuid_checked "$USB_BOOT" "$USB_BOOT" "$PARTUUID_DUP_MODE" "fail" || true)"
      [[ -n "$BOOT_PARTUUID" ]] || { log_error "Failed to get boot PARTUUID"; return 1; }
      if ! update_fstab "$root_mount" "$BOOT_PARTUUID" "$USB_ROOT_PARTUUID"; then
        log_error "Failed to update fstab"
        return 1
      fi
      sync
      log_success "Updated fstab"
      return 0
    else
      log_error "rsync failed"
      return 1
    fi
  }
  
  # Execute root filesystem copy with automatic cleanup
  if ! with_mount "$USB_ROOT" copy_root_filesystem; then
    err --debug "Failed to copy root filesystem"
  fi
  log_success "USB storage device prepared successfully!"
  echo
  log_info "Next steps:"
  log_info "  1. Reboot to test: sudo reboot"
  log_info "  2. Verify boot: findmnt / should show $USB_ROOT"
  log_info "  3. Optional: Configure EEPROM for full USB boot (no SD card needed)"
  echo
  log_info "USB Boot: $USB_BOOT (PARTUUID=$(get_partuuid_safe "$USB_BOOT" || echo unknown))"
  log_info "USB Root: $USB_ROOT (PARTUUID=$USB_ROOT_PARTUUID)"

  verify_or_prompt_eeprom_boot_order "0xf4" "USB Only"
  
  # Ensure all USB devices are configured
  ensure_all_usb_boot_partitions

  PARTUUID_DUP_MODE="$previous_dup_mode"
}

configure_usb_boot_partitions() {
  # Brief health check (non-blocking)
  log_step "Configure USB Boot Partitions"
  log_info "This will check all USB boot partitions and ensure they have:"
  log_info "  - USB power settings in config.txt"
  log_info "  - Homelab setup scripts"
  echo
  log_info "Note: This only updates files, it does NOT change boot order."
  echo
  
  # Configure current boot partition explicitly
  if ! ensure_boot_configured "$BOOT_DIR" "current boot device"; then
    log_warning "Could not fully configure current boot partition"
  fi

  # Then configure all other USB boot partitions
  local result=0
  ensure_all_usb_boot_partitions || result=$?
  
  echo
  if [[ $result -eq 0 ]]; then
    log_success "All USB boot partitions checked and configured successfully!"
  else
    log_warning "USB boot partitions checked - some configurations failed (see above)"
  fi
}

# Flash fresh Raspberry Pi OS to USB using wget and dd
flash_fresh_usb() {
  local missing_tools=()
  local required_tool
  for required_tool in wget sha256sum xz dd lsblk blkid partprobe udevadm mount umount findmnt; do
    if ! command -v "$required_tool" >/dev/null 2>&1; then
      missing_tools+=("$required_tool")
    fi
  done
  if (( ${#missing_tools[@]} > 0 )); then
    err "Missing required tools for mode 7: ${missing_tools[*]}"
  fi

  # Verify SSH and WiFi before flashing (user may want to copy these configs)
  if ! preflight_boot_change "flashing" --allow-failure; then
    echo
    log_success "Warning: SSH/WiFi configuration issues detected."
    log_success "You can still flash, but you may lose remote access after reboot."
    if ! prompt_yn "Continue anyway?"; then
      err "Aborted - fix SSH/WiFi configuration first"
    fi
  fi

  prompt_fresh_flash_identity
  
  log_step "Flash Fresh Raspberry Pi OS to USB"
  log_info "This will download and write a fresh Raspberry Pi OS image to USB storage."
  log_info "Uses wget for downloading, sha256sum for verification, and dd for writing."
  echo
  
  # Select target USB storage device
  select_target_storage_device "Enter USB storage device name (whole disk, e.g., sda or nvme0n1): " "full"
  
  echo
  log_info "Current partitions on $USB_DEV:"
  print_device_tree "$USB_DEV"
  echo
  
  # Get image URL
  log_info "Enter image URL (without .img.xz extension)"
  log_info "Example: https://downloads.raspberrypi.com/raspios_lite_arm64_latest"
  log_info "Or press Enter for default: raspios_lite_arm64_latest"
  echo
  read -rp "Image URL: " IMAGE_URL
  
  # Use default if empty
  if [[ -z "$IMAGE_URL" ]]; then
    IMAGE_URL="https://downloads.raspberrypi.com/raspios_lite_arm64_latest"
    log_success "Using default: $IMAGE_URL"
  fi

  if [[ "$IMAGE_URL" =~ \.img\.xz$ || "$IMAGE_URL" =~ _latest$ ]]; then
    IMAGE_XZ_URL="$IMAGE_URL"
  else
    IMAGE_XZ_URL="${IMAGE_URL}.img.xz"
    log_info "Normalized image URL to: $IMAGE_XZ_URL"
  fi
  
  # Construct full URLs
  SHA256_URL="${IMAGE_XZ_URL}.sha256"
  
  echo
  print_box "DESTRUCTIVE OPERATION WARNING" \
    "" \
    "This will:" \
    "  1. Download Raspberry Pi OS image (~2GB download)" \
    "  2. Verify SHA256 checksum" \
    "  3. COMPLETELY ERASE all data on $USB_DEV" \
    "  4. Write fresh OS image to device" \
    "  5. Apply Fresh_Install-style first-boot files" \
    "     (cmdline/config/network-config/user-data/ssh)" \
    "" \
    "Current partitions on $USB_DEV will be PERMANENTLY DELETED" \
    "Make sure you have selected the CORRECT device!" \
    "" \
    "Device: $USB_DEV" \
    "Transport: USB" \
    ""
  
  read -rp "Type 'yes' to continue: " CONFIRM
  [[ "${CONFIRM,,}" == "yes" ]] || err "Aborted"
  
  # Unmount any existing partitions
  unmount_device_partitions "$USB_NAME"
  
  # Create temporary directory for transient files
  TEMP_DIR=$(mktemp -d -t flash-usb-XXXXXX)
  local USE_PERSISTENT_CACHE=false
  local CACHE_DIR=""
  local CACHE_ENTRY_DIR=""
  
  # Safe cleanup function that checks for mounted filesystems
  cleanup_temp_dir() {
    if [[ -d "$TEMP_DIR" ]]; then
      # Check if anything is mounted under TEMP_DIR
      if mount | grep -q "$TEMP_DIR"; then
        log_warning "Mounted filesystem detected in $TEMP_DIR, unmounting..."
        umount -R "$TEMP_DIR" 2>/dev/null || umount -l "$TEMP_DIR" 2>/dev/null || true
      fi
      rm -rf "$TEMP_DIR"
    fi
  }
  
  # Save original trap and set up temp dir cleanup (will be restored at function exit)
  local original_trap=$(trap -p EXIT)
  trap cleanup_temp_dir EXIT
  
  local temp_dir_parent="$(dirname "$TEMP_DIR")"
  local temp_available_mb=$(df -m "$temp_dir_parent" | tail -1 | awk '{print $4}')
  log_info "Available temporary space in $temp_dir_parent: ${temp_available_mb}MB"

  if (( temp_available_mb < MIN_DISK_SPACE_FLASH_MB )); then
    log_warning "Less than ${MIN_DISK_SPACE_FLASH_MB}MB free in temporary storage (${temp_available_mb}MB available)"
    log_warning "This is common when /tmp is RAM-backed on smaller Raspberry Pi models"

    if prompt_yn "Use a persistent image cache directory instead?" "true"; then
      read -rp "Cache directory [${DEFAULT_FLASH_CACHE_DIR}]: " CACHE_DIR
      CACHE_DIR="${CACHE_DIR:-$DEFAULT_FLASH_CACHE_DIR}"

      mkdir -p "$CACHE_DIR" 2>/dev/null || err "Could not create cache directory: $CACHE_DIR"
      if ! check_writable "$CACHE_DIR"; then
        err "Cache directory is not writable: $CACHE_DIR"
      fi

      USE_PERSISTENT_CACHE=true
      log_success "Using persistent cache directory: $CACHE_DIR"
    else
      if ! prompt_yn "Continue with temporary directory anyway?"; then
        cleanup_temp_dir
        err "Aborted - insufficient temporary disk space"
      fi
    fi
  fi

  if [[ "$USE_PERSISTENT_CACHE" == "true" ]]; then
    local cache_key
    cache_key=$(printf "%s" "$IMAGE_XZ_URL" | sha256sum | awk '{print substr($1,1,16)}')
    CACHE_ENTRY_DIR="${CACHE_DIR}/${cache_key}"
    mkdir -p "$CACHE_ENTRY_DIR" || err "Could not create cache entry directory: $CACHE_ENTRY_DIR"

    IMAGE_FILE="$CACHE_ENTRY_DIR/raspios.img.xz"
    SHA256_FILE="$CACHE_ENTRY_DIR/raspios.img.xz.sha256"
    printf '%s\n' "$IMAGE_XZ_URL" > "$CACHE_ENTRY_DIR/source_url.txt" 2>/dev/null || true
  else
    IMAGE_FILE="$TEMP_DIR/raspios.img.xz"
    SHA256_FILE="$TEMP_DIR/raspios.img.xz.sha256"
  fi

  local SHA256_LATEST_FILE="$SHA256_FILE"
  if [[ "$USE_PERSISTENT_CACHE" == "true" ]]; then
    SHA256_LATEST_FILE="$TEMP_DIR/latest.sha256"
  fi

  # Check available disk space for the chosen image destination
  local image_dir="$(dirname "$IMAGE_FILE")"
  local available_mb=$(df -m "$image_dir" | tail -1 | awk '{print $4}')
  log_info "Available space in $image_dir: ${available_mb}MB"

  if (( available_mb < MIN_DISK_SPACE_FLASH_MB )); then
    log_warning "Less than ${MIN_DISK_SPACE_FLASH_MB}MB free space (${available_mb}MB available)"
    log_warning "Image download may fail if disk runs out of space"
    if ! prompt_yn "Continue anyway?"; then
      cleanup_temp_dir
      err "Aborted - insufficient disk space"
    fi
  else
    log_success "Sufficient disk space available"
  fi

  # Download checksum first so cached images can be reused if still current
  local expected_hash=""
  local checksum_downloaded=false
  local should_download_image=true

  log_success "Downloading SHA256 checksum..."
  if retry_until_connected --attempts ${WGET_RETRY_ATTEMPTS} --delay ${WGET_RETRY_DELAY} \
      wget -O "$SHA256_LATEST_FILE" "$SHA256_URL"; then
    expected_hash=$(awk '{print $1}' "$SHA256_LATEST_FILE" | head -1)
    if [[ -n "$expected_hash" ]]; then
      checksum_downloaded=true
      if [[ "$USE_PERSISTENT_CACHE" == "true" ]]; then
        cp "$SHA256_LATEST_FILE" "$SHA256_FILE" 2>/dev/null || true
      fi
      log_success "SHA256 checksum downloaded"
    else
      log_warning "Downloaded checksum file did not contain a valid hash"
    fi
  else
    log_warning "Could not download SHA256 checksum"
  fi

  # Reuse cached image if checksum still matches the latest release
  if [[ "$USE_PERSISTENT_CACHE" == "true" && -f "$IMAGE_FILE" ]]; then
    log_info "Cached image found: $IMAGE_FILE"
    local cached_actual_hash
    cached_actual_hash=$(sha256sum "$IMAGE_FILE" | awk '{print $1}')

    if [[ "$checksum_downloaded" == "true" && -n "$expected_hash" ]]; then
      if [[ "$cached_actual_hash" == "$expected_hash" ]]; then
        should_download_image=false
        log_success "Cached image matches latest Raspberry Pi OS checksum - reusing cache"
      else
        log_info "Cached image does not match latest checksum - downloading updated image"
      fi
    elif [[ -f "$SHA256_FILE" ]]; then
      local cached_expected_hash
      cached_expected_hash=$(awk '{print $1}' "$SHA256_FILE" | head -1)
      if [[ -n "$cached_expected_hash" && "$cached_actual_hash" == "$cached_expected_hash" ]]; then
        if prompt_yn "Could not verify newest version online. Use existing cached image anyway?" "true"; then
          expected_hash="$cached_expected_hash"
          should_download_image=false
          log_warning "Using cached image without online update verification"
        fi
      fi
    fi
  fi

  if [[ "$should_download_image" == "true" ]]; then
    # Download image with two-stage retry
    log_success "Downloading image from: $IMAGE_XZ_URL"
    log_success "This may take 5-60 minutes depending on image size and connection speed..."
    echo

    local download_attempts=0
    local max_download_attempts=3

    while (( download_attempts < max_download_attempts )); do
      download_attempts=$((download_attempts + 1))

      if (( download_attempts > 1 )); then
        echo
        log_success "Download retry ${download_attempts}/${max_download_attempts}..."
      fi

      # Stage 1: Test connection with HEAD request
      log_success "[Stage 1] Testing connection to download server..."
      if ! retry_until_connected --attempts ${WGET_RETRY_ATTEMPTS} --delay ${WGET_RETRY_DELAY} \
          wget --spider --server-response "$IMAGE_XZ_URL"; then

        log_error "Could not establish connection after ${WGET_RETRY_ATTEMPTS} attempts"

        if (( download_attempts < max_download_attempts )); then
          if prompt_yn "Try again with DNS flush?"; then
            # Attempt DNS flush (may require systemd-resolved)
            systemctl restart systemd-resolved 2>/dev/null || true
            sleep 2
            continue
          else
            cleanup_temp_dir
            err "Failed to establish connection to $IMAGE_XZ_URL"
          fi
        else
          cleanup_temp_dir
          err "Failed to establish connection after ${max_download_attempts} attempts"
        fi
      fi

      echo
      log_success "[Stage 2] Connection established, starting download with progress monitoring..."

      # Stage 2: Download with interactive progress monitoring
      local download_result=0
      monitor_download_progress "Raspberry Pi OS image" \
          wget --progress=bar:force -O "$IMAGE_FILE" "$IMAGE_XZ_URL" || download_result=$?

      # Check for special retry signal from interactive menu
      if (( download_result == 2 )); then
        log_success "User requested retry, attempting again..."
        rm -f "$IMAGE_FILE" 2>/dev/null
        sleep 2
        continue
      elif (( download_result != 0 )); then
        log_error "Download failed or was aborted"

        if (( download_attempts < max_download_attempts )); then
          echo
          log_success "Cleaning up partial download..."
          rm -f "$IMAGE_FILE" 2>/dev/null

          if prompt_yn "Retry download?"; then
            sleep 2
            continue
          else
            cleanup_temp_dir
            err "Download aborted by user"
          fi
        else
          cleanup_temp_dir
          err "Failed to download image after ${max_download_attempts} attempts"
        fi
      fi

      # Success - exit retry loop
      break
    done

    if [[ "$USE_PERSISTENT_CACHE" == "true" ]]; then
      log_success "Downloaded image saved to cache for future flashes"
    fi
  else
    log_success "Skipping image download and using cached file"
  fi

  echo

  # Verify checksum when available
  if [[ -n "$expected_hash" ]]; then
    log_success "Verifying SHA256 checksum..."
    local actual_hash
    actual_hash=$(sha256sum "$IMAGE_FILE" | awk '{print $1}')
    if [[ "$expected_hash" == "$actual_hash" ]]; then
      log_success "SHA256 checksum verified successfully"
    else
      err "SHA256 checksum mismatch - image may be corrupted or tampered"
    fi
  else
    log_warning "Could not verify SHA256 checksum with a valid hash"
    if ! prompt_yn "  Continue without checksum verification?"; then
      err "Aborted - checksum verification required"
    fi
  fi

  # Validate downloaded image
  log_success "Validating XZ compression integrity..."
  if ! xz -t "$IMAGE_FILE" 2>/dev/null; then
    err "Downloaded image appears corrupted (failed xz integrity check)"
  fi
  log_success "Image validation passed"
  
  # Write image using dd
  echo
  log_success "Writing image to $USB_DEV using dd..."
  log_success "This typically takes 5-15 minutes depending on image size and USB speed"
  log_success "Progress will be shown below..."
  echo
  
  if ! xz -dc "$IMAGE_FILE" | dd of="$USB_DEV" bs=16M iflag=fullblock status=progress oflag=direct conv=fsync; then
    err "Failed to write image to $USB_DEV"
  fi
  
  sync
  log_success "Image written successfully"
  
  # Wait for kernel to recognize new partitions
  log_success "Triggering partition table rescan..."
  partprobe "$USB_DEV" 2>/dev/null || true
  udevadm settle 2>/dev/null || true
  
  # Wait for partitions with timeout
  if ! wait_for_partitions "$USB_DEV" 2 20; then
    cleanup_on_error "$USB_DEV" "fresh OS flash"
    err --debug "Partitions not detected after image write"
  fi
  
  # Regenerate PARTUUIDs to avoid conflicts with source image
  regenerate_partuuids "$USB_DEV" || log_warning "Could not regenerate PARTUUIDs"
  
  # Detect partition names
  USB_BOOT="$(get_partition_name "$USB_DEV" 1)"
  USB_ROOT="$(get_partition_name "$USB_DEV" 2)"
  
  # Wait for expected partition nodes specifically
  if ! wait_for_device "$USB_BOOT" 10; then
    err --debug "Boot partition not detected: $USB_BOOT"
  fi
  if ! wait_for_device "$USB_ROOT" 10; then
    err --debug "Root partition not detected: $USB_ROOT"
  fi
  log_success "Boot partition detected: $USB_BOOT"
  log_success "Root partition detected: $USB_ROOT"

  local boot_fstype
  local root_fstype
  boot_fstype="$(blkid -s TYPE -o value "$USB_BOOT" 2>/dev/null || true)"
  root_fstype="$(blkid -s TYPE -o value "$USB_ROOT" 2>/dev/null || true)"

  if [[ ! "$boot_fstype" =~ ^(vfat|fat|msdos)$ ]]; then
    err "Unexpected boot partition filesystem on $USB_BOOT: ${boot_fstype:-unknown} (expected vfat/fat)"
  fi
  if [[ "$root_fstype" != "ext4" ]]; then
    err "Unexpected root partition filesystem on $USB_ROOT: ${root_fstype:-unknown} (expected ext4)"
  fi
  log_success "Post-flash filesystem signatures look valid (boot=$boot_fstype, root=$root_fstype)"
  
  local previous_dup_mode="$PARTUUID_DUP_MODE"
  PARTUUID_DUP_MODE="fix"

  # Re-settle partition state after PARTUUID regeneration.
  if ! wait_for_partitions "$USB_DEV" 2 20; then
    err --debug "Partitions became unavailable after PARTUUID regeneration"
  fi

  # Re-read PARTUUIDs after regeneration (they were just changed by sgdisk)
  log_info "Reading new PARTUUIDs after regeneration..."
  USB_BOOT_PARTUUID="$(read_required_partuuid_checked "$USB_BOOT" "$USB_BOOT" "$PARTUUID_DUP_MODE" "fail" || true)"
  USB_ROOT_PARTUUID="$(read_required_partuuid_checked "$USB_ROOT" "$USB_ROOT" "$PARTUUID_DUP_MODE" "fail" || true)"
  [[ -n "$USB_BOOT_PARTUUID" ]] || log_warning "Could not read boot PARTUUID"
  [[ -n "$USB_ROOT_PARTUUID" ]] || err "Could not read root PARTUUID"

  print_change_plan "flash fresh USB" "${USB_BOOT}:/cmdline.txt" "$USB_ROOT_PARTUUID" "$USB_BOOT_PARTUUID" "${USB_ROOT}:/etc/fstab"

  log_success "Applying Fresh_Install-style first-boot configuration..."
  if ! with_mount "$USB_BOOT" apply_fresh_install_boot_files --verify-boot \
      "$USB_ROOT_PARTUUID" "$FLASH_WIFI_SSID" "$FLASH_WIFI_PASSWORD" \
      "$FLASH_USERNAME" "$FLASH_USER_PASSWORD" "$FLASH_HOSTNAME"; then
    err --debug "Failed to apply Fresh_Install files on $USB_BOOT"
  fi

  # Keep script copy behavior for recovery convenience.
  copy_scripts_on_fresh_boot() {
    local boot_mount="$1"
    copy_scripts_to_boot "$boot_mount" "newly flashed USB drive"
  }
  if ! with_mount "$USB_BOOT" copy_scripts_on_fresh_boot --verify-boot; then
    log_warning "Could not copy scripts to newly flashed boot partition"
  fi

  # Ensure fstab is aligned with regenerated PARTUUIDs
  update_fstab_on_usb_root() {
    local root_mount="$1"
    update_fstab "$root_mount" "$USB_BOOT_PARTUUID" "$USB_ROOT_PARTUUID"
  }
  if ! with_mount "$USB_ROOT" update_fstab_on_usb_root; then
    err "Failed to update fstab on root partition"
  fi

  log_step "Boot Readiness Diagnostics"
  if ! diagnose_boot_pair_readiness "$USB_BOOT" "$USB_ROOT" "freshly flashed USB target"; then
    err "Boot readiness diagnostics failed for freshly flashed USB target"
  fi
  
  # Cleanup handled by trap, but call explicitly for clarity
  cleanup_temp_dir
  
  # Restore original trap
  eval "$original_trap"
  
  echo
  print_separator
  log_success "Fresh Raspberry Pi OS successfully flashed to $USB_DEV!"
  print_separator
  echo
  log_info "Next steps:"
  log_info "  1. Power off: poweroff"
  log_info "  2. Remove SD card (optional - can leave as fallback)"
  log_info "  3. Power on - Pi should boot from USB"
  log_info "  4. If Pi doesn't boot from USB, configure EEPROM boot order:"
  log_info "     raspi-config -> Advanced Options -> Boot Order -> USB Boot"
  log_info "  5. Scripts are in /boot/firmware/ (e.g., /boot/firmware/homelab-setup.sh)"
  echo

  # Clear sensitive prompted values before leaving mode 7.
  FLASH_WIFI_PASSWORD=""
  FLASH_USER_PASSWORD=""

  PARTUUID_DUP_MODE="$previous_dup_mode"
}

# Cleanup any orphaned mounts from previous failed runs
cleanup_failed_mounts

# Parse command line arguments for debug flag
if [[ "${1:-}" == "--debug" ]]; then
  DEBUG_MODE=true
  log_info "Debug mode enabled - timestamps will be shown"
  echo
fi

# Main menu (config checks run per-option as needed)
log_step "Raspberry Pi Root Switcher"
echo "1) Show current setting"
echo "2) Verify initial configuration (SSH and WiFi)"
echo "3) Configure USB boot partitions (add power settings to config.txt)"
echo "4) Switch to current-device-only (boot+root on current device)"
echo "5) Switch to Hybrid (boot on current device, root on another)"
echo "6) Clone OS to USB target (boot+root)"
echo "7) Flash fresh Raspberry Pi OS to USB (wget + sha256sum + dd)"
echo "8) Configure EEPROM boot order (SD/USB priority)"
echo "q) Quit"
read -rp "Enter choice: " choice
choice="${choice,,}"

case "$choice" in
  1) show_current ;;
  2) check_initial_config --include-boot-targets ;;
  3) configure_usb_boot_partitions ;;
  4) switch_current_only ;;
  5) switch_hybrid ;;
  6) clone_os_to_usb_target ;;
  7) flash_fresh_usb ;;
  8) configure_eeprom_boot_order ;;
  q) echo "Exiting." ;;
  *) err "Invalid choice" ;;
esac
