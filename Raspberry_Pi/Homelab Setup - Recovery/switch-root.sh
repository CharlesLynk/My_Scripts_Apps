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
#   6) Switch to USB-only (full USB storage setup with boot+root)
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

# WiFi verification settings (connection not required)
readonly WIFI_SSID="LynkNet"
readonly WIFI_PSK="Tay98Chase97Wifi"

# Connection retry and progress monitoring constants
readonly CONNECTION_RETRY_MAX_ATTEMPTS=15
readonly CONNECTION_RETRY_DELAY=2
readonly WGET_RETRY_ATTEMPTS=15
readonly WGET_RETRY_DELAY=2
readonly PROGRESS_CHECK_INTERVAL=5

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
    lsblk -o NAME,SIZE,TYPE,FSTYPE,MOUNTPOINT >&2 || true
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
  
  if [[ "$default_yes" == "true" ]]; then
    read -rp "$prompt (Y/n): " answer
    [[ -z "$answer" || "$answer" =~ ^[Yy]$ ]]
  else
    read -rp "$prompt (y/n): " answer
    [[ "$answer" =~ ^[Yy]$ ]]
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
        ((check_count++))
        
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
            ((stall_count++))
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
            
            # Show last few lines of log for context
            if [[ -f "$temp_log" ]]; then
                local last_line=$(tail -n 1 "$temp_log" 2>/dev/null | tr -d '\n' | cut -c1-80)
                if [[ -n "$last_line" ]]; then
                    echo "  Last: $last_line"
                fi
            fi
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
      ((copied++))
    else
      log_error "Failed to copy ${filename}"
      ((failed++))
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
    
    ((attempt++))
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
  
  # Already present - success
  if grep -q "usb_max_current_enable" "$config_file" 2>/dev/null; then
    return 0
  fi
  
  # Add settings
  if ! append_usb_power_settings "$config_file"; then
    return 1  # Error writing file
  fi
  
  sync
  return 0  # Success
}

# Helper: Append USB power settings block without heredoc (avoid editor indentation issues)
append_usb_power_settings() {
  local config_file="$1"

  printf '\n# Unlock power for USB Boot on USB-Compliant power supply (added by switch-root.sh)\n' >> "$config_file" || return 1
  printf 'dtoverlay=dwc2,dr_mode=host\n' >> "$config_file" || return 1
  printf 'usb_max_current_enable=1\n' >> "$config_file" || return 1
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
  if [[ "$1" == "--verify-boot" ]]; then
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
    
    ((attempt++))
  done
  
  log_error "Timeout: $description failed after ${max_attempts} attempts"
  return 1
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
    local transport=$(lsblk -ndo TRAN "$device" 2>/dev/null || echo "unknown")
    local removable=$(lsblk -ndo RM "$device" 2>/dev/null || echo "0")
    local is_mmc=false
    [[ "$device" =~ /dev/mmcblk ]] && is_mmc=true

    # Accept USB, removable devices, and SD/eMMC (mmcblk)
    if [[ "$transport" != "usb" && "$removable" != "1" && "$is_mmc" != "true" ]]; then
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
  
  if retry_until "Getting PARTUUID for $device" 5 0.5 check_partuuid_available "$device"; then
    blkid -s PARTUUID -o value "$device" 2>/dev/null
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
      ((cleaned++))
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
  
  if ! command -v sgdisk >/dev/null 2>&1; then
    log_warning "sgdisk not available, skipping PARTUUID regeneration"
    return 1
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

# Helper: Verify PARTUUID uniqueness in system
# Usage: check_partuuid_unique "abc-123-def"
check_partuuid_unique() {
  local partuuid="$1"
  
  # Force fresh probe to avoid stale cache after partition operations
  partprobe 2>/dev/null || true
  udevadm settle 2>/dev/null || true
  
  local count=$(blkid -s PARTUUID | grep -c "PARTUUID=\"$partuuid\"" || echo 0)
  
  if (( count > 1 )); then
    log_error "ERROR: PARTUUID $partuuid appears on $count devices"
    log_error "This will cause boot conflicts. Run: sudo blkid | grep $partuuid"
    ERROR_HINT="Regenerate PARTUUIDs with: sudo sgdisk --randomize-guids /dev/sdX"
    return 1
  fi
  return 0
}

# Helper: Check PARTUUID uniqueness and fail with debug info if conflict detected
# Usage: check_partuuid_unique_or_fail "abc-123-def" "/dev/sda2"
check_partuuid_unique_or_fail() {
  local partuuid="$1"
  local device="${2:-device}"
  
  if ! check_partuuid_unique "$partuuid"; then
    log_error "PARTUUID conflict detected: $partuuid"
    log_error "Devices with this PARTUUID:"
    blkid -s PARTUUID | grep "$partuuid" | sed 's/^/  /' >&2 || true
    err --debug "PARTUUID conflict on $device (PARTUUID=$partuuid) - cannot proceed safely"
  fi
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

# Helper: Update cmdline.txt with new root PARTUUID (with rollback and validation)
# Usage: update_cmdline_root "/boot/firmware/cmdline.txt" "abc-123-def"
update_cmdline_root() {
  local target_cmdline="$1"
  local partuuid="$2"
  
  [[ -f "$target_cmdline" ]] || { log_error "cmdline.txt not found: $target_cmdline"; return 1; }
  [[ -n "$partuuid" ]] || { log_error "PARTUUID is empty"; return 1; }
  
  # Create temporary backup
  local temp_backup="${target_cmdline}.tmp_backup_$$"
  if ! cp "$target_cmdline" "$temp_backup"; then
    log_error "Failed to create temporary backup"
    return 1
  fi
  
  # Perform modification
  if ! sed -i -E "s#root=[^ ]+#root=PARTUUID=$partuuid#" "$target_cmdline"; then
    cp "$temp_backup" "$target_cmdline" 2>/dev/null
    rm "$temp_backup" 2>/dev/null
    log_error "sed command failed"
    return 1
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
  if ! grep -q "root=PARTUUID=$partuuid" "$target_cmdline"; then
    cp "$temp_backup" "$target_cmdline"
    rm "$temp_backup" 2>/dev/null
    log_error "Validation failed: PARTUUID not found in updated cmdline.txt"
    return 1
  fi
  
  # Success - remove temp backup and sync
  rm "$temp_backup" 2>/dev/null
  sync
  
  log_success "Updated cmdline.txt with root=PARTUUID=$partuuid"
  return 0
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
  
  # Validate inputs
  [[ -f "$fstab_path" ]] || { log_error "fstab not found: $fstab_path"; return 1; }
  [[ -n "$boot_partuuid" && -n "$root_partuuid" ]] || {
    log_error "Missing PARTUUIDs (boot=$boot_partuuid, root=$root_partuuid)"
    return 1
  }
  
  # Backup fstab before modification
  cp "$fstab_path" "${fstab_path}.bak" || return 1
  
  # Update boot partition (usually /boot/firmware)
  if grep -q "^\s*[^#]*/boot/firmware\s*" "$fstab_path"; then
    # Preserve filesystem type and options, only update device
    sed -i -E "s#^[^#]*\s+(/boot/firmware\s+)#PARTUUID=$boot_partuuid  \1#" "$fstab_path"
  else
    echo "PARTUUID=$boot_partuuid  /boot/firmware  vfat  defaults  0  2" >> "$fstab_path"
  fi
  
  # Update root partition (match lines with exactly / as mount point, not subdirectories)
  if grep -qE "^[^#]+\s+/\s+[a-z]" "$fstab_path" 2>/dev/null; then
    # Preserve filesystem type and options, only update device
    sed -i -E "s#^([^#]*\s+)(/\s+)#PARTUUID=$root_partuuid  \2#" "$fstab_path"
  else
    echo "PARTUUID=$root_partuuid  /  ext4  defaults,noatime  0  1" >> "$fstab_path"
  fi
  
  # Verify changes were applied
  grep -q "PARTUUID=$boot_partuuid" "$fstab_path" && \
  grep -q "PARTUUID=$root_partuuid" "$fstab_path" || {
    log_error "fstab verification failed"
    return 1
  }
  
  sync
  log_success "Updated fstab: boot=PARTUUID=$boot_partuuid, root=PARTUUID=$root_partuuid"
  return 0
}

################################################################################
# SYSTEM CONFIGURATION CHECKS (SSH/WIFI)
################################################################################

# Check if SSH is properly configured with default credentials
check_ssh_config() {
  local errors=()
  local target_user=""
    
    echo
    log_success "Checking SSH configuration..."

    # Determine the user to validate (prefer sudo-invoking user)
    target_user="${SUDO_USER:-}"
    if [[ -z "$target_user" || "$target_user" == "root" ]]; then
      target_user=$(logname 2>/dev/null || true)
    fi
    if [[ -z "$target_user" || "$target_user" == "root" ]]; then
      target_user=$(whoami 2>/dev/null || true)
    fi
    if [[ -z "$target_user" || "$target_user" == "root" ]]; then
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
    
    # Check SSH configuration allows password authentication
    if [[ -f /etc/ssh/sshd_config ]]; then
        if grep -qE '^PasswordAuthentication[[:space:]]+no' /etc/ssh/sshd_config; then
            errors+=("SSH password authentication is disabled in sshd_config")
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
      warnings+=("/etc/wpa_supplicant/wpa_supplicant.conf not found")
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
        warnings+=("${WIFI_SSID} not found in /etc/NetworkManager/system-connections")
      fi
    else
      warnings+=("/etc/NetworkManager/system-connections not found")
    fi

    # Check active NetworkManager WiFi connection
    if command -v nmcli >/dev/null 2>&1; then
      local nm_active_ssid
      nm_active_ssid=$(nmcli -t -f active,ssid dev wifi 2>/dev/null | awk -F: '$1=="yes" {print $2; exit}')
      if [[ "$nm_active_ssid" == "$WIFI_SSID" ]]; then
        log_success "  ✓ NetworkManager active on ${WIFI_SSID}"
        nm_configured=true
        wifi_configured=true
      elif [[ -n "$nm_active_ssid" ]]; then
        warnings+=("NetworkManager active on different SSID: $nm_active_ssid (expected: ${WIFI_SSID})")
      fi
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
check_initial_config() {
    echo
    print_separator
    echo "Initial Configuration Verification"
    print_separator
    
    local ssh_status=0
    local wifi_status=0
    
    check_ssh_config || ssh_status=$?
    check_wifi_config || wifi_status=$?
    
    echo
    print_separator
    
    if [[ $ssh_status -eq 0 ]] && [[ $wifi_status -eq 0 ]]; then
        log_success "All configuration checks passed!"
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

        print_separator
        return 1
    fi
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
  
  # Get all USB FAT32 partitions in one query, excluding current boot disk
  mapfile -t partition_devices < <(
    lsblk -nrpo NAME,FSTYPE,TRAN,PKNAME | \
    awk -v boot="$current_boot_disk" \
      '$2=="vfat" && $3=="usb" && $4!=boot && $4!="" {print $1}'
  )
  
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
      ((configured++))
      log_success "Configured $part_dev"
    else
      ((failed++))
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
  
  if ! prompt_with_context "Configuration Sync" "" "Copy current system configurations to ${target_name}?"; then
    echo "Configuration sync skipped"
    return 0
  fi
  
  local synced=0
  for config in "${configs_to_check[@]}"; do
    [[ ! -f "$config" ]] && continue
    
    if cp "$config" "${root_mount_point}${config}" 2>/dev/null; then
      log_success "Synced $config"
      ((synced++))
    fi
  done
  
  # Special handling for SSH keys
  if [[ -d "/etc/ssh" ]] && [[ -d "${root_mount_point}/etc/ssh" ]]; then
    if prompt_with_context "SSH Host Keys" "" "Copy SSH host keys to maintain same identity?"; then
      for keyfile in /etc/ssh/ssh_host_*_key*; do
        [[ -f "$keyfile" ]] || continue
        if cp "$keyfile" "${root_mount_point}/etc/ssh/$(basename "$keyfile")" 2>/dev/null; then
          log_success "Copied $(basename "$keyfile")"
        fi
      done
    fi
  fi
  
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
  
  # Handle users without sudo access
  if [[ "$has_sudo" == "false" ]]; then
    echo
    log_warning "User '$username' is NOT in the sudo group"
    log_warning "This user will have LIMITED access on the target system"
    echo
    
    if prompt_with_context "Grant Sudo Access" \
      "User will have LIMITED access without sudo privileges." \
      "Add $username to sudo group now?"; then
      
      if usermod -aG sudo "$username"; then
        log_success "Added $username to sudo group"
        has_sudo=true
      else
        log_error "Failed to add user to sudo group"
        if ! prompt_yn "Continue anyway?"; then
          log_info "Skipping user configuration"
          return 1
        fi
      fi
    elif ! prompt_yn "Continue without sudo privileges?"; then
      log_info "Skipping user configuration"
      return 1
    fi
  fi
  
  # Only copy user if not already present on target (prevent duplicates)
  if ! grep -q "^${username}:" "$root_mount_point/etc/passwd" 2>/dev/null; then
    grep "^${username}:" /etc/passwd >> "$root_mount_point/etc/passwd" 2>/dev/null || true
    grep "^${username}:" /etc/shadow | tee -a "$root_mount_point/etc/shadow" >/dev/null 2>&1 || true
    grep "^${username}:" /etc/group >> "$root_mount_point/etc/group" 2>/dev/null || true
  else
    log_success "[i] User $username already exists on target, updating password only"
    # Update password in shadow file
    local user_shadow=$(grep "^${username}:" /etc/shadow)
    sed -i "/^${username}:/d" "$root_mount_point/etc/shadow"
    echo "$user_shadow" | tee -a "$root_mount_point/etc/shadow" >/dev/null
  fi
  
  # Copy sudo group membership if user should have sudo
  if [[ "$has_sudo" == "true" ]]; then
    if ! grep -q "^sudo:.*${username}" "$root_mount_point/etc/group"; then
      sed -i "s/^sudo:\([^:]*\):\([^:]*\):\(.*\)/sudo:\1:\2:\3,${username}/" "$root_mount_point/etc/group"
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
  
  # Copy SSH authorized_keys if exists (use actual UID/GID)
  if [[ -f "/home/$username/.ssh/authorized_keys" ]]; then
    mkdir -p "$root_mount_point/home/$username/.ssh"
    cp "/home/$username/.ssh/authorized_keys" "$root_mount_point/home/$username/.ssh/"
    chown -R "$user_uid:$user_gid" "$root_mount_point/home/$username/.ssh"
    chmod 700 "$root_mount_point/home/$username/.ssh"
    chmod 600 "$root_mount_point/home/$username/.ssh/authorized_keys"
    log_success "Copied SSH authorized_keys"
  fi
  
  return 0
}

# Helper: Copy SSH host keys to target system
# Usage: copy_ssh_keys "/mnt/fresh_root"
copy_ssh_keys() {
  local root_mount_point="$1"
  
  # Copy SSH host keys (preserves SSH fingerprint)
  if [[ -d /etc/ssh ]] && [[ -d "$root_mount_point/etc/ssh" ]]; then
    log_success "Copying SSH host keys..."
    cp /etc/ssh/ssh_host_* "$root_mount_point/etc/ssh/" 2>/dev/null || true
    log_success "SSH host keys copied (SSH fingerprint preserved)"
  else
    log_error "SSH directory not found, skipping host key copy"
    return 1
  fi
  
  # Enable SSH if currently enabled
  if systemctl is-enabled ssh >/dev/null 2>&1 || systemctl is-enabled sshd >/dev/null 2>&1; then
    # Create systemd symlink to enable SSH on first boot
    mkdir -p "$root_mount_point/etc/systemd/system/multi-user.target.wants"
    ln -sf /lib/systemd/system/ssh.service "$root_mount_point/etc/systemd/system/multi-user.target.wants/ssh.service" 2>/dev/null || true
    log_success "SSH service enabled"
  fi
  
  return 0
}

# Helper: Copy network configuration to target system
# Usage: copy_network_config "/mnt/fresh_root"
copy_network_config() {
  local root_mount_point="$1"
  
  # Copy network configuration if using static IP
  if [[ -f /etc/dhcpcd.conf ]] && grep -q "^interface" /etc/dhcpcd.conf; then
    if prompt_with_context "Network Configuration" "" "Copy network configuration (static IP)?"; then
      cp /etc/dhcpcd.conf "$root_mount_point/etc/dhcpcd.conf" 2>/dev/null || {
        log_error "Failed to copy network configuration"
        return 1
      }
      log_success "Network configuration copied"
    fi
  fi

  # Copy NetworkManager WiFi profiles if present
  if [[ -d /etc/NetworkManager/system-connections ]]; then
    if prompt_with_context "Network Configuration" "" "Copy NetworkManager WiFi profiles?"; then
      mkdir -p "$root_mount_point/etc/NetworkManager/system-connections"
      cp /etc/NetworkManager/system-connections/* "$root_mount_point/etc/NetworkManager/system-connections/" 2>/dev/null || true
      chmod 600 "$root_mount_point/etc/NetworkManager/system-connections/"* 2>/dev/null || true
      chown root:root "$root_mount_point/etc/NetworkManager/system-connections/"* 2>/dev/null || true
      log_success "NetworkManager profiles copied"
    fi
  fi
  
  return 0
}

# Helper: Write fstab without heredoc (avoid editor indentation issues)
# Usage: write_fstab "$root_mount" "$boot_partuuid" "$root_partuuid"
write_fstab() {
  local root_mount_point="$1"
  local boot_partuuid="$2"
  local root_partuuid="$3"
  local fstab_path="${root_mount_point}/etc/fstab"

  [[ -n "$root_mount_point" && -n "$boot_partuuid" && -n "$root_partuuid" ]] || return 1

  printf '# Generated by switch-root.sh\n' > "$fstab_path" || return 1
  printf 'proc            /proc           proc    defaults          0       0\n' >> "$fstab_path" || return 1
  printf 'PARTUUID=%s  /boot/firmware  vfat    defaults          0       2\n' "$boot_partuuid" >> "$fstab_path" || return 1
  printf 'PARTUUID=%s  /               ext4    defaults,noatime  0       1\n' "$root_partuuid" >> "$fstab_path" || return 1
  return 0
}

# Helper: Configure user and SSH on fresh OS installation
# Usage: configure_user_and_ssh "/mnt/fresh_root"
configure_user_and_ssh() {
  local root_mount_point="$1"
  
  if ! prompt_with_context "User and SSH Configuration" \
    "The fresh image has default user settings. You can:\n  1. Keep fresh image defaults (you'll need keyboard/monitor on first boot)\n  2. Copy your current user/SSH configuration (recommended for headless)" \
    "Copy current user and SSH config to target?"; then
    log_info "Skipped user/SSH configuration"
    log_info "First boot will require keyboard/monitor or fresh SSH setup"
    return 0
  fi
  
  log_success "Copying user and SSH configuration..."
  
  # Get current username (skip root)
  local detected_user="${SUDO_USER:-}"
  if [[ -z "$detected_user" || "$detected_user" == "root" ]]; then
    detected_user=$(logname 2>/dev/null || true)
  fi
  if [[ -z "$detected_user" || "$detected_user" == "root" ]]; then
    detected_user=$(who am i | awk '{print $1}' 2>/dev/null || true)
  fi
  [[ "$detected_user" == "root" ]] && detected_user=""
  
  # If detection failed, try to find a non-root user with home directory
  if [[ -z "$detected_user" ]]; then
    detected_user=$(awk -F: '$3 >= 1000 && $3 < 65534 && $1 != "nobody" {print $1; exit}' /etc/passwd)
  fi
  
  # Last resort: ask user
  if [[ -z "$detected_user" ]]; then
    read -rp "  Enter username to configure on target system: " detected_user
    if [[ -z "$detected_user" ]]; then
      log_info "No username provided, skipping user configuration"
      return 0
    fi
  fi
  
  # Copy user account using helper function
  if ! copy_user_account "$root_mount_point" "$detected_user"; then
    log_error "Failed to copy user account"
    return 1
  fi
  
  # Copy SSH keys and enable service
  copy_ssh_keys "$root_mount_point" || log_warning "SSH configuration may be incomplete"
  
  # Copy network configuration
  copy_network_config "$root_mount_point" || true
  
  sync
  log_success "User and SSH configuration complete"
  return 0
}

################################################################################
# MAIN OPERATIONS
################################################################################

# Find current root partition and PARTUUID
CURRENT_ROOT="$(findmnt -no SOURCE /)"
CURRENT_PARTUUID="$(blkid -s PARTUUID -o value "$CURRENT_ROOT" 2>/dev/null || true)"

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

get_eeprom_boot_order() {
  if ! eeprom_tool_available; then
    echo ""
    return 1
  fi
  local order
  order=$(rpi-eeprom-config 2>/dev/null | awk -F= '/^BOOT_ORDER=/{print $2; exit}' | sed 's/#.*//')
  normalize_boot_order "$order"
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

  log_info "Applying EEPROM BOOT_ORDER=${desired}"
  if rpi-eeprom-config --apply "$tmp_conf"; then
    rm -f "$tmp_conf"
    log_success "EEPROM boot order staged successfully"
    log_info "Reboot required for EEPROM changes to take effect"
    return 0
  fi

  rm -f "$tmp_conf"
  return 1
}

verify_or_prompt_eeprom_boot_order() {
  local desired="$1"
  local mode_label="$2"

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

  local transport
  transport=$(lsblk -no TRAN "/dev/$boot_disk" 2>/dev/null | head -n1)

  case "$transport" in
    mmc) echo "sd" ;;
    usb|sata|ata|pci|pcie) echo "usb" ;;
    *) echo "unknown" ;;
  esac
}

################################################################################
# DEVICE INFO HELPERS (SHOW CURRENT)
################################################################################

transport_label() {
  case "$1" in
    usb) echo "USB" ;;
    mmc) echo "SD/eMMC" ;;
    sata|ata) echo "SATA" ;;
    pci|pcie) echo "PCIe storage" ;;
    *) echo "Unknown" ;;
  esac
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

  local info
  info=$(lsblk -P -o NAME,SIZE,TYPE,FSTYPE,LABEL,PARTUUID,UUID,MODEL,TRAN,MOUNTPOINT "$dev" 2>/dev/null | head -n1)
  if [[ -z "$info" ]]; then
    log_warning "Unable to read device info for $dev"
    return 0
  fi

  # shellcheck disable=SC2086
  eval "$info"
  local transport
  transport=$(transport_label "${TRAN:-}")

  echo "  Device: /dev/${NAME} (${TYPE:-unknown})"
  echo "  Size: ${SIZE:-unknown}"
  echo "  Filesystem: ${FSTYPE:-unknown}"
  echo "  Label: ${LABEL:-none}"
  echo "  PARTUUID: ${PARTUUID:-unknown}"
  echo "  UUID: ${UUID:-unknown}"
  echo "  Model: ${MODEL:-unknown}"
  echo "  Transport: ${transport}"
  echo "  Mountpoint: ${MOUNTPOINT:-none}"
}

resolve_cmdline_root_device() {
  local root_value="$1"
  local resolved=""

  case "$root_value" in
    PARTUUID=*)
      resolved=$(blkid -t "PARTUUID=${root_value#PARTUUID=}" -o device 2>/dev/null | head -n1)
      ;;
    UUID=*)
      resolved=$(blkid -t "UUID=${root_value#UUID=}" -o device 2>/dev/null | head -n1)
      ;;
    /dev/*)
      resolved="$root_value"
      ;;
    *)
      resolved=""
      ;;
  esac

  echo "$resolved"
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

  local desired_label=""
  case "$sel" in
    1) desired_label="SD+USB" ;;
    2) desired_label="USB+SD" ;;
    3) desired_label="SD Only" ;;
    4) desired_label="USB Only" ;;
    q|Q) log_info "Cancelled"; return 0 ;;
    *) err "Invalid selection" ;;
  esac

  local desired
  desired=$(choose_eeprom_boot_order "$desired_label")
  [[ -n "$desired" ]] || err "Invalid BOOT_ORDER selection"

  if [[ -n "$current" && "$current" == "$desired" ]]; then
    log_success "EEPROM BOOT_ORDER already set to ${desired} (${desired_label})"
    return 0
  fi

  log_info "Selected BOOT_ORDER=${desired} (${desired_label})"
  if prompt_yn "Apply EEPROM boot order now?" true; then
    if ! apply_eeprom_boot_order "$desired"; then
      log_warning "EEPROM update failed - you may need to adjust manually"
    fi
  else
    log_info "Skipped EEPROM update"
  fi
}

switch_current_only() {
  # Verify SSH and WiFi before changing boot configuration
  echo
  log_success "Verifying SSH and WiFi configuration before boot configuration change..."
  if ! check_initial_config; then
    echo
    err "Cannot proceed with boot configuration change - fix SSH/WiFi configuration first"
  fi
  
  # Ensure backup exists before modifying
  ensure_cmdline_backup
  
  # Set cmdline.txt to boot and run from current device (SD, USB, etc.)
  [[ -n "$CURRENT_PARTUUID" ]] || err "Could not determine PARTUUID of current root"
  
  # Update cmdline.txt using helper with rollback
  if ! update_cmdline_root "$CMDLINE" "$CURRENT_PARTUUID"; then
    err "Failed to update $CMDLINE"
  fi
  
  # Ensure boot partition is configured (already mounted at $BOOT_DIR)
  if ! ensure_boot_configured "$BOOT_DIR" "current boot device"; then
    log_warning "Could not fully configure boot partition"
  fi
  
  # Update current fstab if needed
  FSTAB_PATH="/etc/fstab"
  if [[ -f "$FSTAB_PATH" ]]; then
    if prompt_yn "Update /etc/fstab to match current boot configuration?"; then
      cp "$FSTAB_PATH" "${FSTAB_PATH}.bak"
      log_success "Backed up fstab to ${FSTAB_PATH}.bak"
      
      BOOT_DEV="$(findmnt -no SOURCE "$BOOT_DIR")"
      BOOT_PARTUUID="$(get_partuuid_safe "$BOOT_DEV" || true)"
      
      # Verify PARTUUIDs are unique
      if [[ -n "$BOOT_PARTUUID" ]]; then
        check_partuuid_unique_or_fail "$BOOT_PARTUUID" "$BOOT_DEV"
      fi
      if [[ -n "$CURRENT_PARTUUID" ]]; then
        check_partuuid_unique_or_fail "$CURRENT_PARTUUID" "$CURRENT_ROOT"
      fi
      
      if [[ -n "$BOOT_PARTUUID" && -n "$CURRENT_PARTUUID" ]]; then
        # Update boot partition line
        if grep -q "^\s*[^#]*/boot/firmware\s*" "$FSTAB_PATH"; then
          sed -i -E "s#^[^#]*\s+/boot/firmware\s+#PARTUUID=$BOOT_PARTUUID  /boot/firmware  #" "$FSTAB_PATH"
        else
          echo "PARTUUID=$BOOT_PARTUUID  /boot/firmware  vfat  defaults  0  2" >> "$FSTAB_PATH"
        fi
        
        # Update root partition line (match lines with exactly / as mount point)
        if grep -qE "^[^#]+\s+/\s+[a-z]" "$FSTAB_PATH" 2>/dev/null; then
          sed -i -E "s#^([^#]*\s+)(/\s+)#PARTUUID=$CURRENT_PARTUUID  \2#" "$FSTAB_PATH"
        else
          echo "PARTUUID=$CURRENT_PARTUUID  /  ext4  defaults,noatime  0  1" >> "$FSTAB_PATH"
        fi
        
        log_success "Updated fstab with boot=PARTUUID=$BOOT_PARTUUID, root=PARTUUID=$CURRENT_PARTUUID"
      else
        log_warning "Could not determine PARTUUIDs, fstab not updated"
      fi
    fi
  fi
  
  echo
  log_success "Current device configuration complete!"
  log_info "Boot and root are both on the current device."
  log_info "Scripts available at: /boot/firmware/"

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
  
  # Ensure all USB devices are also configured
  ensure_all_usb_boot_partitions
}

switch_hybrid() {
  # Verify SSH and WiFi before changing boot configuration
  echo
  log_success "Verifying SSH and WiFi configuration before boot configuration change..."
  if ! check_initial_config; then
    echo
    err "Cannot proceed with boot configuration change - fix SSH/WiFi configuration first"
  fi
  
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
    ((i++))
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
  PARTUUID="$(get_partuuid_safe "$ROOT_DEV" || true)"
  [[ -n "$PARTUUID" ]] || err "Could not determine PARTUUID of $ROOT_DEV"
  
  # Verify PARTUUID is unique (MUST be called AFTER PARTUUID assignment)
  check_partuuid_unique_or_fail "$PARTUUID" "$ROOT_DEV"
  
  BOOT_DEV="$(findmnt -no SOURCE "$BOOT_DIR")"
  BOOT_PARTUUID="$(get_partuuid_safe "$BOOT_DEV" || true)"
  [[ -n "$BOOT_PARTUUID" ]] || err "Could not determine PARTUUID of boot device $BOOT_DEV"
  
  # Verify boot PARTUUID is unique
  check_partuuid_unique_or_fail "$BOOT_PARTUUID" "$BOOT_DEV"
  
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
  if ! ensure_boot_configured "$BOOT_DIR" "current boot device"; then
    log_warning "Could not configure boot partition"
  fi

  # Update cmdline.txt on current boot device using helper
  echo
  log_info "Updating boot configuration..."
  if ! update_cmdline_root "$CMDLINE" "$PARTUUID"; then
    err "Failed to update boot configuration"
  fi
  
  echo
  log_success "Hybrid configuration complete!"
  log_info "Current boot device will load root from: $ROOT_DEV"
  log_info "Reboot to test: sudo reboot"

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
  
  # Ensure all USB devices are configured
  ensure_all_usb_boot_partitions
}

switch_usb_only() {
  # Verify SSH and WiFi before USB boot setup
  echo
  log_success "Verifying SSH and WiFi configuration before USB boot setup..."
  if ! check_initial_config; then
    echo
    err "Cannot proceed with USB boot setup - fix SSH/WiFi configuration first"
  fi
  
  log_step "USB-Only Boot Setup"
  log_info "This will prepare a USB storage device with both boot and root partitions."
  echo
  
  # Select target USB storage device
  log_info "Available storage devices:"
  lsblk -d -o NAME,SIZE,MODEL,TRAN
  echo
  read -rp "Enter USB storage device name (whole disk, e.g., sda or nvme0n1): " USB_NAME
  
  # Match: sda-sdz, sdaa-sdzz (SCSI), mmcblk0-99 (SD/eMMC), nvme0n1-nvme9n9 (NVMe)
  valid_disk_regex='^(sd[a-z]{1,2}|mmcblk[0-9]+|nvme[0-9]+n[0-9]+)$'
  [[ "$USB_NAME" =~ $valid_disk_regex ]] || err "Invalid device name (expected: sda, sdaa, mmcblk0, nvme0n1, etc.)"
  
  USB_DEV="/dev/${USB_NAME}"
  
  # Validate it's a USB device (standard checks: exists, readable, USB transport)
  if ! validate_target_device "$USB_DEV" "standard"; then
    err "Safety check failed: $USB_DEV validation failed"
  fi
  log_success "Verified $USB_DEV is a USB device"
  
  # Verify it's not the current root disk
  CURRENT_ROOT_DISK="$(lsblk -no PKNAME "$CURRENT_ROOT")"
  [[ "$USB_NAME" == "$CURRENT_ROOT_DISK" ]] && err "Cannot use current root disk"
  
  echo
  log_info "Current partitions on $USB_DEV:"
  lsblk "$USB_DEV" || true
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
  
  read -rp "Type YES to continue: " CONFIRM
  if [[ "$CONFIRM" != "YES" ]]; then
    log_info "Aborted"
    return 0
  fi
  
  # Unmount any existing partitions
  for node in $(lsblk -ln -o NAME,PKNAME,TYPE | awk -v p="$USB_NAME" '$3=="part" && $2==p {print "/dev/"$1}'); do
    disable_swap_if_active "$node" || true
    if mountpoint -q "$node" 2>/dev/null || grep -q "$node" /proc/mounts 2>/dev/null; then
      log_success "Unmounting $node"
      umount "$node" 2>/dev/null || umount -l "$node" 2>/dev/null || err "Failed to unmount $node"
    fi
  done
  
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
  
  # Get PARTUUIDs after regeneration and formatting (must be done after regenerate_partuuids)
  log_success "Reading PARTUUIDs after regeneration..."
  USB_BOOT_PARTUUID="$(get_partuuid_safe "$USB_BOOT" || true)"
  [[ -n "$USB_BOOT_PARTUUID" ]] || err "Could not determine PARTUUID of $USB_BOOT"
  
  USB_ROOT_PARTUUID="$(get_partuuid_safe "$USB_ROOT" || true)"
  [[ -n "$USB_ROOT_PARTUUID" ]] || err "Could not determine PARTUUID of $USB_ROOT"
  
  # Verify PARTUUIDs are unique - fail if conflict detected
  check_partuuid_unique_or_fail "$USB_BOOT_PARTUUID" "$USB_BOOT"
  check_partuuid_unique_or_fail "$USB_ROOT_PARTUUID" "$USB_ROOT"
  
  # Clone boot partition
  log_success "Preparing boot partition on $USB_BOOT..."
  
  # Define callback for boot partition configuration
  configure_usb_boot() {
    local boot_mount="$1"
    
    # Check writable
    if ! check_writable "$boot_mount"; then
      return 1
    fi
    
    log_success "Copying boot files (this may take 1-2 minutes)..."
    rsync -aHAXx --info=progress2 "$BOOT_DIR/" "$boot_mount/"
    sync
    
    # Verify boot files were copied
    if ! verify_filesystem "$boot_mount" "boot"; then
      log_error "Boot file verification failed - copy incomplete"
      return 1
    fi
    log_success "Boot files verified"
    
    # Update cmdline.txt on USB storage boot partition using helper
    local usb_cmdline="$boot_mount/cmdline.txt"
    if ! update_cmdline_root "$usb_cmdline" "$USB_ROOT_PARTUUID"; then
      return 1
    fi
    
    # Configure boot partition (power settings + scripts)
    if ! ensure_boot_configured "$boot_mount" "USB boot partition"; then
      log_warning "Could not fully configure USB boot partition"
    fi
    
    return 0
  }
  
  # Execute boot configuration with automatic cleanup
  if ! with_mount "$USB_BOOT" configure_usb_boot; then
    err --debug "Failed to configure USB boot partition"
  fi
  
  log_success "USB boot partition ready! Boot files copied and configured."
  log_info "Root partition is formatted but empty - choose whether to copy data next."
  echo
  
  local COPY_FS_DONE=false
  
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
  
  if prompt_with_context "Copy Root Filesystem" \
    "Do you want to copy the current system to the USB root partition?\nThis will take 10-30 minutes depending on data size." \
    "Copy filesystem now?"; then
    COPY_FS_DONE=true
    
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
        
        # Create fstab
        log_success "Creating fstab on USB root partition"
        local BOOT_PARTUUID="$(get_partuuid_safe "$USB_BOOT" || true)"
        [[ -n "$BOOT_PARTUUID" ]] || { log_error "Failed to get boot PARTUUID"; return 1; }
        if ! write_fstab "$root_mount" "$BOOT_PARTUUID" "$USB_ROOT_PARTUUID"; then
          log_error "Failed to write fstab"
          return 1
        fi
        sync
        log_success "Created fstab"
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
  else
    log_info "Skipped filesystem copy"
    log_info "You'll need to manually copy or restore data to $USB_ROOT"
  fi
  log_success "USB storage device prepared successfully!"
  echo
  log_info "Next steps:"
  if [[ "$COPY_FS_DONE" == "true" ]]; then
    log_info "  1. Reboot to test: sudo reboot"
    log_info "  2. Verify boot: findmnt / should show $USB_ROOT"
    log_info "  3. Optional: Configure EEPROM for full USB boot (no SD card needed)"
  else
    log_info "  1. Copy your system to $USB_ROOT"
    log_info "  2. Update fstab on $USB_ROOT"
    log_info "  3. Reboot and test"
  fi
  echo
  log_info "USB Boot: $USB_BOOT (PARTUUID=$(get_partuuid_safe "$USB_BOOT" || echo unknown))"
  log_info "USB Root: $USB_ROOT (PARTUUID=$USB_ROOT_PARTUUID)"

  verify_or_prompt_eeprom_boot_order "0xf4" "USB Only"
  
  # Ensure all USB devices are configured
  ensure_all_usb_boot_partitions
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
  
  # Just call the universal function
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
  # Verify SSH and WiFi before flashing (user may want to copy these configs)
  echo
  log_success "Verifying SSH and WiFi configuration before flashing..."
  if ! check_initial_config; then
    echo
    log_success "Warning: SSH/WiFi configuration issues detected."
    log_success "You can still flash, but you may lose remote access after reboot."
    if ! prompt_yn "Continue anyway?"; then
      err "Aborted - fix SSH/WiFi configuration first"
    fi
  fi
  
  log_step "Flash Fresh Raspberry Pi OS to USB"
  log_info "This will download and write a fresh Raspberry Pi OS image to USB storage."
  log_info "Uses wget for downloading, sha256sum for verification, and dd for writing."
  echo
  
  # Select target USB storage device
  echo
  log_info "Available storage devices:"
  lsblk -d -o NAME,SIZE,MODEL,TRAN
  echo
  read -rp "Enter USB storage device name (whole disk, e.g., sda or nvme0n1): " USB_NAME
  
  valid_disk_regex='^(sd[a-z]{1,2}|mmcblk[0-9]+|nvme[0-9]+n[0-9]+)$'
  [[ "$USB_NAME" =~ $valid_disk_regex ]] || err "Invalid device name (expected: sda, sdaa, mmcblk0, nvme0n1, etc.)"
  
  USB_DEV="/dev/${USB_NAME}"
  
  # Validate device with full checks (exists, readable, USB transport, no critical mounts)
  if ! validate_target_device "$USB_DEV" "full"; then
    err "Safety check failed: $USB_DEV validation failed"
  fi
  log_success "Verified $USB_DEV is a USB device with no critical mounts"
  
  # Verify it's not the current root disk
  CURRENT_ROOT_DISK="$(lsblk -no PKNAME "$CURRENT_ROOT")"
  [[ "$USB_NAME" == "$CURRENT_ROOT_DISK" ]] && err "Cannot use current root disk"
  
  echo
  log_info "Current partitions on $USB_DEV:"
  lsblk "$USB_DEV" || true
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
  
  # Construct full URLs
  IMAGE_XZ_URL="${IMAGE_URL}"
  SHA256_URL="${IMAGE_XZ_URL}.sha256"
  
  echo
  print_box "DESTRUCTIVE OPERATION WARNING" \
    "" \
    "This will:" \
    "  1. Download Raspberry Pi OS image (~2GB download)" \
    "  2. Verify SHA256 checksum" \
    "  3. COMPLETELY ERASE all data on $USB_DEV" \
    "  4. Write fresh OS image to device" \
    "  5. Configure USB power settings and scripts" \
    "" \
    "Current partitions on $USB_DEV will be PERMANENTLY DELETED" \
    "Make sure you have selected the CORRECT device!" \
    "" \
    "Device: $USB_DEV" \
    "Transport: USB" \
    ""
  
  read -rp "Type YES to continue: " CONFIRM
  [[ "$CONFIRM" == "YES" ]] || err "Aborted"
  
  # Unmount any existing partitions
  for node in $(lsblk -ln -o NAME,PKNAME,TYPE | awk -v p="$USB_NAME" '$3=="part" && $2==p {print "/dev/"$1}'); do
    disable_swap_if_active "$node" || true
    if mountpoint -q "$node" 2>/dev/null || grep -q "$node" /proc/mounts 2>/dev/null; then
      log_success "Unmounting $node"
      umount "$node" 2>/dev/null || umount -l "$node" 2>/dev/null || true
    fi
  done
  
  # Create temporary directory for downloads
  TEMP_DIR=$(mktemp -d -t flash-usb-XXXXXX)
  
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
  
  IMAGE_FILE="$TEMP_DIR/raspios.img.xz"
  SHA256_FILE="$TEMP_DIR/raspios.img.xz.sha256"
  
  # Check available disk space (need ~2GB for compressed image)
  local temp_dir_parent="$(dirname "$TEMP_DIR")"
  local available_mb=$(df -m "$temp_dir_parent" | tail -1 | awk '{print $4}')
  log_info "Available space in $temp_dir_parent: ${available_mb}MB"
  
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
  
  # Download image with two-stage retry
  log_success "Downloading image from: $IMAGE_XZ_URL"
  log_success "This may take 5-60 minutes depending on image size and connection speed..."
  echo
  
  local download_attempts=0
  local max_download_attempts=3
  
  while (( download_attempts < max_download_attempts )); do
    ((download_attempts++))
    
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
  
  echo
  
  # Download SHA256 checksum with retry
  log_success "Downloading SHA256 checksum..."
  if retry_until_connected --attempts ${WGET_RETRY_ATTEMPTS} --delay ${WGET_RETRY_DELAY} \
      wget -O "$SHA256_FILE" "$SHA256_URL"; then
    log_success "Verifying SHA256 checksum..."
    
    # Compare expected hash to downloaded file (independent of filename)
    local expected_hash=$(awk '{print $1}' "$SHA256_FILE" | head -1)
    local actual_hash=$(sha256sum "$IMAGE_FILE" | awk '{print $1}')
    if [[ -n "$expected_hash" && "$expected_hash" == "$actual_hash" ]]; then
      log_success "SHA256 checksum verified successfully"
    else
      err "SHA256 checksum mismatch - image may be corrupted or tampered"
    fi
  else
    log_warning "Could not download SHA256 checksum, skipping verification"
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
  
  if ! xz -dc "$IMAGE_FILE" | dd of="$USB_DEV" bs=16M status=progress oflag=direct conv=fsync; then
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
  
  # Wait for boot partition specifically
  if ! wait_for_device "$USB_BOOT" 10; then
    err --debug "Boot partition not detected: $USB_BOOT"
  fi
  log_success "Boot partition detected: $USB_BOOT"
  
  # Re-read PARTUUIDs after regeneration (they were just changed by sgdisk)
  log_info "Reading new PARTUUIDs after regeneration..."
  USB_BOOT_PARTUUID="$(get_partuuid_safe "$USB_BOOT" || true)"
  USB_ROOT_PARTUUID="$(get_partuuid_safe "$USB_ROOT" || true)"
  [[ -n "$USB_BOOT_PARTUUID" ]] || log_warning "Could not read boot PARTUUID"
  [[ -n "$USB_ROOT_PARTUUID" ]] || log_warning "Could not read root PARTUUID"
  
  # Verify PARTUUIDs are unique
  if [[ -n "$USB_BOOT_PARTUUID" ]]; then
    check_partuuid_unique_or_fail "$USB_BOOT_PARTUUID" "$USB_BOOT"
  fi
  if [[ -n "$USB_ROOT_PARTUUID" ]]; then
    check_partuuid_unique_or_fail "$USB_ROOT_PARTUUID" "$USB_ROOT"
  fi
  
  # Configure boot partition
  log_success "Configuring boot partition..."
  
  if ! with_mount "$USB_BOOT" ensure_boot_configured --verify-boot "newly flashed USB drive"; then
    err --debug "Failed to configure boot partition $USB_BOOT"
  fi
  log_success "Boot partition configured"
  
  # Offer to configure user and SSH from current system
  [[ -b "$USB_ROOT" ]] || err "Root partition not detected: $USB_ROOT"
  
  # Configure user and SSH on fresh installation
  if ! with_mount "$USB_ROOT" configure_user_and_ssh; then
    log_warning "Failed to configure user/SSH on root partition"
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
}

# Cleanup any orphaned mounts from previous failed runs
cleanup_failed_mounts

# Check and update current boot config.txt with USB power settings
if [[ -f "$BOOT_DIR/config.txt" ]]; then
  if ! grep -q "usb_max_current_enable" "$BOOT_DIR/config.txt" 2>/dev/null; then
    log_success "Adding USB power settings to current boot config.txt..."
    if ! add_usb_power_settings "$BOOT_DIR/config.txt"; then
      log_warning "Failed to add USB power settings"
    else
      log_success "USB power settings configured"
    fi
  fi
fi

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
echo "6) Switch to USB-only (full USB storage setup with boot+root)"
echo "7) Flash fresh Raspberry Pi OS to USB (wget + sha256sum + dd)"
echo "8) Configure EEPROM boot order (SD/USB priority)"
echo "q) Quit"
read -rp "Enter choice: " choice

case "$choice" in
  1) show_current ;;
  2) check_initial_config ;;
  3) configure_usb_boot_partitions ;;
  4) switch_current_only ;;
  5) switch_hybrid ;;
  6) switch_usb_only ;;
  7) flash_fresh_usb ;;
  8) configure_eeprom_boot_order ;;
  q|Q) echo "Exiting." ;;
  *) err "Invalid choice" ;;
esac
