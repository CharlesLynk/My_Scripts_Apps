#!/bin/bash
#
################################################################################
# Raspberry Pi Homelab Docker Stack Setup
# 
# Automated setup script for Raspberry Pi OS Trixie/Bookworm that:
# - Installs Docker-based homelab stack (Pi-hole, Home Assistant, Homeway)
# - Configures Tailscale for secure remote access
# - Provides backup and restore functionality
#

# Usage: ./homelab-setup.sh <mode> [options]
#
# Modes:
#   install-docker  - Install Docker stack (Pi-hole, HA, Homeway, Tailscale)
#   configure       - Post-setup configuration (collect tokens/keys)
#   backup          - Backup all configurations
#   restore         - Restore from backup

################################################################################
# CONSTANTS AND CONFIGURATION
################################################################################

readonly SCRIPT_VERSION="1.0.0"

################################################################################
# Directory Structure
################################################################################
readonly HOMELAB_DIR="/homelab"
readonly BACKUP_DIR="${HOMELAB_DIR}/backups"
readonly COMPOSE_DIR="${HOMELAB_DIR}/compose"

################################################################################
# Connection & Network Retry Configuration
################################################################################
# Fast connection establishment (for DNS issues)
readonly CONNECTION_RETRY_MAX_ATTEMPTS=15
readonly CONNECTION_RETRY_DELAY=2

# HTTP download retries (curl/wget operations)
readonly CURL_RETRY_ATTEMPTS=15
readonly CURL_RETRY_DELAY=2

# apt repository/package retries
readonly APT_RETRY_ATTEMPTS=5
readonly APT_RETRY_DELAY=4

# Image pull retry strategy
readonly IMAGE_PULL_OUTER_RETRY_MAX=3
readonly PROGRESS_CHECK_INTERVAL=5

################################################################################
# System Operation Delays
################################################################################
# Post-operation wait times (seconds)
readonly DNS_REFRESH_DELAY=2
readonly CLEANUP_DELAY=2
readonly DOCKER_VERIFY_DELAY=3
readonly SERVICE_STATUS_CHECK_DELAY=5

# File system operations
readonly FILE_LOCK_RETRY_DELAY=3
readonly FILE_LOCK_RETRY_MAX_TIME=15

# Container lifecycle management
readonly CONTAINER_STOP_WAIT_SECONDS=30
readonly CONTAINER_STOP_CHECK_INTERVAL=1

################################################################################
# Service Readiness Checks
################################################################################
# Pi-hole startup (attempts × delay = max wait time)
readonly PIHOLE_READY_CHECK_ATTEMPTS=24    # 24 × 5s = 2 minutes
readonly PIHOLE_READY_CHECK_DELAY=5

# Home Assistant startup (can take 10-15 minutes on slow storage)
readonly HA_READY_CHECK_ATTEMPTS=90        # 90 × 10s = 15 minutes
readonly HA_READY_CHECK_DELAY=10

################################################################################
# Resource Requirements
################################################################################
readonly MIN_DISK_SPACE_MB=100
readonly DOCKER_IMAGE_MIN_SPACE_MB=3072

################################################################################
# Docker Images
################################################################################
readonly PIHOLE_IMAGE="pihole/pihole:v6.2.2"
readonly HOMEASSISTANT_IMAGE="ghcr.io/home-assistant/home-assistant:2025.10.4"
readonly HOMEWAY_IMAGE="homewayio/homeway:latest"

################################################################################
# Application Configuration
################################################################################
# Pi-hole settings
readonly PIHOLE_PASSWORD="password"
readonly PIHOLE_DNS_UPSTREAM="1.1.1.1;1.0.0.1"
readonly PIHOLE_STATIC_IP="192.168.1.199"
readonly HAGEZI_BLOCKLIST="https://raw.githubusercontent.com/hagezi/dns-blocklists/main/hosts/pro.txt"

################################################################################
# Terminal Color Codes
################################################################################
readonly RED='\033[0;31m'
readonly GREEN='\033[0;32m'
readonly YELLOW='\033[1;33m'
readonly BLUE='\033[0;34m'
readonly CYAN='\033[0;36m'
readonly BOLD='\033[1m'
readonly NC='\033[0m' # No Color

# Download control state (used by Ctrl+C menu)
DOWNLOAD_MENU_ACTION=""

################################################################################
# UTILITY FUNCTIONS
################################################################################

# Print formatted messages
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

# Print a separator line
print_separator() {
    echo "================================================================================"
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

# Prompt user for yes/no confirmation
# Usage: prompt_yn "Question?" [default_yes]
# Returns: 0 for yes, 1 for no
prompt_yn() {
    local prompt="$1"
    local default_yes="${2:-false}"
    local answer=""
    
    if [[ "$default_yes" == "true" ]]; then
        read -p "${prompt} [Y/n]: " -r answer
        [[ -z "$answer" || "$answer" =~ ^[Yy]$ ]]
    else
        read -p "${prompt} [y/n]: " -r answer
        [[ "$answer" =~ ^[Yy]$ ]]
    fi
}

# Prompt with formatted context section
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

# Check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Check available disk space
# Usage: check_disk_space "/path" required_mb "description"
check_disk_space() {
    local check_path="$1"
    local required_mb="$2"
    local description="${3:-operation}"

    # Resolve to nearest existing parent to avoid df failures on non-existent paths
    local resolved_path="$check_path"
    while [[ ! -d "$resolved_path" ]]; do
        resolved_path=$(dirname "$resolved_path")
        [[ -z "$resolved_path" || "$resolved_path" == "." ]] && resolved_path="/" && break
    done
    [[ -z "$resolved_path" ]] && resolved_path="/"
    
    local available_kb=$(df -Pk "$resolved_path" 2>/dev/null | tail -1 | awk '{print $4}')
    available_kb="${available_kb:-0}"
    local available_mb=$((available_kb / 1024))
    
    if [[ $available_mb -lt $required_mb ]]; then
        log_error "Insufficient disk space for ${description}"
        log_error "Available: ${available_mb}MB, need at least ${required_mb}MB"
        return 1
    fi
    return 0
}

# Generate timestamp for filenames
timestamp_filename() {
    date "+%Y-%m-%d_%H%M%S"
}

# Unified error handler with cleanup
# Usage: err "message"
# Optional: Set ERROR_HINT="hint text" before calling for context-specific help
err() {
    log_error "$*"
    echo
    
    # Show context-specific hint if available
    if [[ -n "${ERROR_HINT:-}" ]]; then
        echo "${CYAN}Suggestion:${NC}" >&2
        echo "  $ERROR_HINT" >&2
        echo
    fi
    
    cleanup_temp_files 2>/dev/null || true
    exit 1
}

# Cleanup temporary files on exit
cleanup_temp_files() {
    rm -f /tmp/get-docker.sh 2>/dev/null || true
    rm -f /tmp/tailscale-install.sh 2>/dev/null || true
}

# Register cleanup on exit
trap cleanup_temp_files EXIT INT TERM

# Print header for each mode
print_header() {
    local mode=$1
    clear
    print_separator
    echo -e "${CYAN}Raspberry Pi Homelab Setup - Mode: ${mode}${NC}"
    echo -e "Version: ${SCRIPT_VERSION}"
    print_separator
    echo
}

# Check if script is run as root
check_root() {
    if [[ $EUID -ne 0 ]]; then
        log_error "This script must be run as root (use sudo)"
        err "Root privileges required"
    fi
}

# Helper function to check if a Docker container is running
# Usage: is_container_running "container_name"
is_container_running() {
    local container_name="$1"
    docker ps --filter "name=${container_name}" --format "{{.Names}}" 2>/dev/null | grep -q "^${container_name}$"
}

# Helper function to check if Docker daemon is running
# Usage: verify_docker_running
verify_docker_running() {
    if ! docker info >/dev/null 2>&1; then
        log_error "Docker daemon is not running"
        log_error "Try: sudo systemctl start docker"
        return 1
    fi
    return 0
}

# Helper function to create directory with error handling
# Usage: create_directory_safe "/path/to/dir" "Description"
create_directory_safe() {
    local dir_path="$1"
    local description="${2:-directory}"
    
    local parent_dir=$(dirname "$dir_path")
    if ! check_disk_space "$parent_dir" "$MIN_DISK_SPACE_MB" "${description}"; then
        log_error "Free up space: sudo apt clean && sudo apt autoremove"
        return 1
    fi
    
    if ! mkdir -p "$dir_path" 2>/dev/null; then
        log_error "Failed to create ${description}: ${dir_path}"
        log_error "Check permissions and disk space"
        return 1
    fi
    
    # Verify directory was actually created
    if [[ ! -d "$dir_path" ]]; then
        log_error "Directory creation reported success but ${dir_path} doesn't exist"
        return 1
    fi
    
    return 0
}

# Check and report status of all homelab containers
# Returns 0 if all critical services are running, 1 if any issues detected
# Note: Homeway is considered optional and doesn't affect return status
report_container_status() {
    local all_ok=true
    
    if is_container_running "pihole"; then
        log_success "Pi-hole: Running"
    else
        log_warning "Pi-hole: Not running (check logs: docker logs pihole)"
        all_ok=false
    fi
    
    if is_container_running "homeassistant"; then
        log_success "Home Assistant: Running"
    else
        log_warning "Home Assistant: Not running (check logs: docker logs homeassistant)"
        all_ok=false
    fi
    
    if is_container_running "homeway"; then
        log_success "Homeway: Running"
    else
        log_info "Homeway: Not running (normal if not configured)"
    fi
    
    if tailscale status >/dev/null 2>&1; then
        log_success "Tailscale: Connected"
    else
        log_warning "Tailscale: Not connected (run: sudo tailscale up)"
        all_ok=false
    fi
    
    echo
    if [ "$all_ok" = true ]; then
        return 0  # All critical services healthy
    else
        return 1  # One or more critical services have issues
    fi
}

# Unified service management function
# Usage: manage_docker_services stop|start [service1 service2 ...]
# If no services specified, operates on all services (pihole, homeassistant, homeway)
manage_docker_services() {
    local action="$1"
    shift
    local services=("$@")
    
    # Default to all services if none specified
    if [[ ${#services[@]} -eq 0 ]]; then
        services=("pihole" "homeassistant" "homeway")
    fi
    
    case "$action" in
        stop)
            # Check if Docker is running
            if ! verify_docker_running; then
                log_warning "Docker daemon not running. Skipping container operations."
                return 0
            fi
            
            for container in "${services[@]}"; do
                if is_container_running "$container"; then
                    log_info "Stopping ${container}..."
                    docker stop -t 10 "$container" 2>/dev/null || true
                    
                    # Wait for container to stop
                    local stopped=false
                    local max_checks=$((CONTAINER_STOP_WAIT_SECONDS / CONTAINER_STOP_CHECK_INTERVAL))
                    for ((i=1; i<=max_checks; i++)); do
                        if ! is_container_running "$container"; then
                            stopped=true
                            break
                        fi
                        sleep $CONTAINER_STOP_CHECK_INTERVAL
                    done
                    
                    if [[ "$stopped" == "false" ]]; then
                        log_warning "${container} failed to stop gracefully, forcing..."
                        docker kill "$container" 2>/dev/null || true
                    fi
                fi
            done
            ;;
            
        start)
            local started=0
            local failed=0
            
            for service_name in "${services[@]}"; do
                local compose_file="${COMPOSE_DIR}/${service_name}-compose.yml"
                
                # Get display name (capitalize first letter)
                local display_name="${service_name^}"
                case "$service_name" in
                    pihole) display_name="Pi-hole" ;;
                    homeassistant) display_name="Home Assistant" ;;
                    homeway) display_name="Homeway" ;;
                esac
                
                # Check if compose file exists
                if [[ ! -f "$compose_file" ]]; then
                    log_warning "${display_name} compose file not found, skipping"
                    continue
                fi
                
                # Special check for Homeway - requires .env file
                if [[ "$service_name" == "homeway" ]] && [[ ! -f "${COMPOSE_DIR}/.env" ]]; then
                    log_warning "Homeway .env file not found, skipping"
                    continue
                fi
                
                log_info "Starting ${display_name}..."
                if docker compose -f "$compose_file" up -d 2>/dev/null; then
                    ((started++))
                else
                    log_error "Failed to start ${display_name}"
                    ((failed++))
                fi
            done
            
            if [[ $started -gt 0 ]] || [[ $failed -gt 0 ]]; then
                log_info "Service start results: ${started} started, ${failed} failed"
            fi
            
            if [[ $failed -gt 0 ]]; then
                return 1
            fi
            ;;
            
        *)
            log_error "Invalid action: $action (use 'start' or 'stop')"
            return 1
            ;;
    esac
    
    return 0
}

# Helper: Write Pi-hole Docker Compose file without heredoc
# Usage: write_pihole_compose "/path/to/pihole-compose.yml" "timezone"
write_pihole_compose() {
    local output_path="$1"
    local timezone="$2"

    local lines=(
        "services:"
        "  pihole:"
        "    image: ${PIHOLE_IMAGE}"
        "    container_name: pihole"
        "    network_mode: \"host\""
        "    environment:"
        "      TZ: \"${timezone}\""
        "      FTLCONF_webserver_api_password: \"${PIHOLE_PASSWORD}\""
        "      FTLCONF_dns_upstreams: \"${PIHOLE_DNS_UPSTREAM}\""
        "    volumes:"
        "      - ${HOMELAB_DIR}/pihole/etc-pihole:/etc/pihole"
        "    restart: unless-stopped"
        "    dns:"
        "      - 127.0.0.1"
        "      - 1.1.1.1"
    )

    printf '%s\n' "${lines[@]}" > "$output_path" || return 1
    return 0
}

# Helper: Write Home Assistant Docker Compose file without heredoc
# Usage: write_homeassistant_compose "/path/to/homeassistant-compose.yml" "timezone"
write_homeassistant_compose() {
    local output_path="$1"
    local timezone="$2"

    local lines=(
        "services:"
        "  homeassistant:"
        "    image: ${HOMEASSISTANT_IMAGE}"
        "    container_name: homeassistant"
        "    network_mode: \"host\""
        "    environment:"
        "      TZ: \"${timezone}\""
        "    volumes:"
        "      - ${HOMELAB_DIR}/homeassistant:/config"
        "    restart: unless-stopped"
        "    privileged: true"
    )

    printf '%s\n' "${lines[@]}" > "$output_path" || return 1
    return 0
}

# Helper: Write Homeway Docker Compose file without heredoc
# Usage: write_homeway_compose "/path/to/homeway-compose.yml"
write_homeway_compose() {
    local output_path="$1"

    local lines=(
        "services:"
        "  homeway:"
        "    image: ${HOMEWAY_IMAGE}"
        "    container_name: homeway"
        "    restart: unless-stopped"
        "    network_mode: host"
        "    environment:"
        '      - HOMEWAY_TOKEN=${HOMEWAY_TOKEN}'
        "    volumes:"
        "      - ${HOMELAB_DIR}/homeway:/data"
    )

    printf '%s\n' "${lines[@]}" > "$output_path" || return 1
    return 0
}

# Helper: Write nftables configuration without heredoc
# Usage: write_nftables_conf "lan_subnet"
write_nftables_conf() {
    local lan_subnet="$1"
    local output_path="/etc/nftables.conf"

    local lines=(
        "#!/usr/sbin/nft -f"
        ""
        "flush ruleset"
        ""
        "table inet filter {"
        "    chain input {"
        "        type filter hook input priority 0; policy drop;"
        "        "
        "        # Allow established/related connections"
        "        ct state established,related accept"
        "        "
        "        # Allow loopback"
        "        iifname \"lo\" accept"
        "        "
        "        # Allow Tailscale (Tailscale handles its own access control via ACLs)"
        "        iifname \"tailscale0\" accept"
        "        "
        "        # Allow from LAN subnet"
        "        ip saddr ${lan_subnet} accept"
        "        "
        "        # Allow ICMP (ping)"
        "        ip protocol icmp accept"
        "        ip6 nexthdr icmpv6 accept"
        "        "
        "        # Note: SSH is allowed via LAN (above) and Tailscale interface"
        "        # Remote SSH access is controlled by Tailscale ACLs, not this firewall"
        "        "
        "        # Log dropped packets (optional, comment out if too noisy)"
        "        # log prefix \"nftables-dropped: \" drop"
        "    }"
        "    "
        "    chain forward {"
        "        type filter hook forward priority 0; policy drop;"
        "    }"
        "    "
        "    chain output {"
        "        type filter hook output priority 0; policy accept;"
        "    }"
        "}"
    )

    printf '%s\n' "${lines[@]}" > "$output_path" || return 1
    return 0
}

################################################################################
# CENTRALIZED PREREQUISITES CHECK
# Checks and installs all required tools at script start with clear documentation
################################################################################

check_all_prerequisites() {
    local mode=$1  # Operating mode to determine which tools are needed
    local missing_packages=()
    local missing_commands=()
    
    log_step "Checking prerequisites for mode: ${mode}"
    
    # First: Validate critical system utilities that can't be installed via apt
    # Core system management (required for script operation)
    if ! command_exists systemctl; then
        missing_commands+=("systemctl (part of systemd)")
    fi
    
    # Package management (needed to install other prerequisites)
    if ! command_exists apt-get; then
        missing_commands+=("apt-get (package manager)")
    fi
    
    # Mount point detection (used in backup and install modes)
    if ! command_exists findmnt; then
        missing_commands+=("findmnt (part of util-linux)")
    fi
    
    # Port conflict detection (used in install-docker mode)
    if ! command_exists netstat && ! command_exists ss; then
        log_warning "Neither 'netstat' nor 'ss' found - port conflict detection may not work"
        log_info "Install net-tools or iproute2 package if needed"
    fi
    
    # Fail if critical system utilities are missing
    if [[ ${#missing_commands[@]} -gt 0 ]]; then
        log_error "Missing critical system utilities: ${missing_commands[*]}"
        log_error "Your system may be misconfigured or using an unsupported OS"
        ERROR_HINT="This script requires Raspberry Pi OS (Bookworm/Trixie) or Debian-based system"
        return 1
    fi
    
    # Second: Check installable packages
    if ! command_exists bc; then
        log_info "bc: Required for retry timing calculations"
        missing_packages+=(bc)
    fi
    
    if [[ "$mode" =~ ^(install-docker|configure)$ ]]; then
        if ! command_exists curl; then
            log_info "curl: Required for downloading Docker and Tailscale"
            missing_packages+=(curl)
        fi
        
        if ! command_exists ipcalc; then
            log_info "ipcalc: Required for network subnet validation"
            missing_packages+=(ipcalc)
        fi
    fi
    
    if [[ "$mode" =~ ^(backup|restore)$ ]]; then
        if ! command_exists tar; then
            log_info "tar: Required for backup compression"
            missing_packages+=(tar)
        fi
    fi
    
    # Install missing packages if any
    if [[ ${#missing_packages[@]} -gt 0 ]]; then
        log_warning "Missing packages detected: ${missing_packages[*]}"
        log_info "Installing missing packages..."
        
        if ! retry_apt_update; then
            err "Failed to update package lists. Check internet connection."
        fi
        
        # Deduplicate and install all packages in one command
        local unique_packages=($(echo "${missing_packages[@]}" | tr ' ' '\n' | sort -u | tr '\n' ' '))
        
        if ! retry_apt_install "${unique_packages[@]}"; then
            log_error "Failed to install packages: ${unique_packages[*]}"
            log_error "Try manually: sudo apt-get install ${unique_packages[*]}"
            err "Package installation failed"
        fi
        
        log_success "All missing packages installed successfully"
    else
        log_success "All prerequisites satisfied"
    fi
    
    return 0
}

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
        if "$@" 2>&1; then
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
            printf "\r${YELLOW}[RETRY]${NC} Connection attempt %d/%d failed, retrying in %ds...     " \
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

# Retry apt metadata refresh with transient-network tolerance
retry_apt_update() {
    retry_until_connected --attempts ${APT_RETRY_ATTEMPTS} --delay ${APT_RETRY_DELAY} \
        apt-get -o Acquire::Retries=3 update -qq
}

# Retry apt package installation with transient-network tolerance
retry_apt_install() {
    [[ $# -eq 0 ]] && return 0
    retry_until_connected --attempts ${APT_RETRY_ATTEMPTS} --delay ${APT_RETRY_DELAY} \
        apt-get -o Acquire::Retries=3 install -y "$@"
}

# Stage 2: Monitor download progress with user interaction
# Usage: monitor_download_progress "operation_name" command args...
# Returns: 0 if completed successfully, 1 if user aborted or failed
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
    log_info "  • Press Ctrl+C to pause and show options menu"
    echo
    
    local last_size=0
    local stall_count=0
    local check_count=0
    local start_time=$(date +%s)
    
    # Set up trap for Ctrl+C to show menu instead of exiting
    local original_int_trap=$(trap -p INT)
    trap 'show_download_menu "$operation_name" "$download_pid" "$temp_log" "$pid_file"' INT
    
    while kill -0 $download_pid 2>/dev/null; do
        if [[ -n "${DOWNLOAD_MENU_ACTION:-}" ]]; then
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
                    err "Setup aborted by user"
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
        echo "Status: RUNNING (PID: ${download_pid})"
    else
        echo "Status: STOPPED"
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
    echo "  4) Abort and exit setup script"
    echo
    read -p "Enter choice [1-4]: " choice
    
    case "$choice" in
        1)
            log_info "Resuming monitoring..."
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

# Legacy compatibility wrapper - maps old retry_with_backoff calls to new system
# Usage: retry_with_backoff [--attempts N] [--initial-delay SECONDS] [--max-time SECONDS] [--fixed-delay] command args...
retry_with_backoff() {
    local max_attempts=5
    local retry_delay=2
    
    # Parse optional arguments (ignore max-time and initial-delay, use our fixed values)
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --attempts)
                max_attempts="$2"
                shift 2
                ;;
            --initial-delay|--max-time)
                # Ignore these - we use fixed retry strategy now
                shift 2
                ;;
            --fixed-delay)
                # This is our default behavior now
                shift
                ;;
            *)
                break
                ;;
        esac
    done
    
    # Just use simple connection retry for legacy calls
    retry_until_connected --attempts "$max_attempts" --delay "$retry_delay" "$@"
}

# Helper: Check registry access for images in a compose file
# Usage: check_registry_access "compose-file.yml"
check_registry_access() {
    local compose_file="$1"
    local images=""

    images=$(docker compose -f "$compose_file" config --images 2>/dev/null) || return 1
    [[ -n "$images" ]] || return 1

    for image in $images; do
        if ! docker manifest inspect "$image" >/dev/null 2>&1; then
            return 1
        fi
    done

    return 0
}

# Helper function to pull and verify Docker images with two-stage retry logic
# Usage: pull_and_verify_docker_image "image:tag" "compose-file.yml" "context_description"
pull_and_verify_docker_image() {
    local image_name="$1"
    local compose_file="$2"
    local context_desc="$3"
    
    local outer_attempts=0
    local max_outer_attempts=${IMAGE_PULL_OUTER_RETRY_MAX}
    
    while [[ $outer_attempts -lt $max_outer_attempts ]]; do
        outer_attempts=$((outer_attempts + 1))
        
        if ! check_disk_space "/var/lib/docker" "$DOCKER_IMAGE_MIN_SPACE_MB" "Docker image pull"; then
            log_error "Free up space: docker system prune -a"
            return 1
        fi
        
        if [[ $outer_attempts -gt 1 ]]; then
            echo
            log_warning "Outer retry ${outer_attempts}/${max_outer_attempts}${context_desc}..."
        fi
        
        # Stage 1: Retry connection establishment (up to 15 quick attempts)
        log_step "Stage 1: Establishing connection to registry"
        if ! retry_until_connected --attempts ${CONNECTION_RETRY_MAX_ATTEMPTS} --delay ${CONNECTION_RETRY_DELAY} \
            check_registry_access "$compose_file"; then
            
            log_error "Could not establish connection to registry after ${CONNECTION_RETRY_MAX_ATTEMPTS} attempts"
            
            if [[ $outer_attempts -lt $max_outer_attempts ]]; then
                if prompt_yn "Try again with fresh DNS resolution?"; then
                    # Flush Docker's DNS cache (systemd-resolved is disabled for Pi-hole)
                    log_info "Restarting Docker to flush DNS cache..."
                    systemctl restart docker 2>/dev/null || true
                    sleep $DNS_REFRESH_DELAY
                    continue
                else
                    return 1
                fi
            else
                return 1
            fi
        fi
        
        echo
        log_step "Stage 2: Downloading image with progress monitoring"
        
        # Stage 2: Monitor download progress interactively
        local download_result=0
        monitor_download_progress "${image_name} download" docker compose -f "$compose_file" pull || download_result=$?
        
        # Check for special retry signal from interactive menu
        if (( download_result == 2 )); then
            log_info "User requested retry, attempting again..."
            # Cleanup and retry
            docker rmi "$image_name" 2>/dev/null || true
            docker image prune -f >/dev/null 2>&1 || true
            sleep $CLEANUP_DELAY
            continue
        elif (( download_result != 0 )); then
            log_error "Download failed or was aborted"
            
            if [[ $outer_attempts -lt $max_outer_attempts ]]; then
                echo
                log_warning "Cleaning up partial download..."
                docker rmi "$image_name" 2>/dev/null || true
                docker image prune -f >/dev/null 2>&1 || true
                
                if prompt_yn "Retry download?"; then
                    sleep $CLEANUP_DELAY
                    continue
                else
                    return 1
                fi
            else
                return 1
            fi
        fi
        
        # Verify image was actually pulled
        echo
        log_info "Verifying image integrity..."
        if docker image inspect "$image_name" >/dev/null 2>&1; then
            log_success "${image_name} pulled and verified successfully"
            return 0
        else
            log_warning "Image pull completed but image not found locally"
            log_info "This may indicate a corrupted download"
            
            if [[ $outer_attempts -lt $max_outer_attempts ]]; then
                if prompt_yn "Retry download?"; then
                    docker rmi "$image_name" 2>/dev/null || true
                    continue
                else
                    return 1
                fi
            else
                return 1
            fi
        fi
    done
    
    log_error "Failed to pull image after ${max_outer_attempts} complete attempts"
    return 1
}

# Check internet connectivity
check_internet() {
    log_info "Checking internet connectivity..."
    if ping -c 1 -W 5 1.1.1.1 >/dev/null 2>&1; then
        log_success "Internet connection: OK"
        return 0
    fi

    log_warning "ICMP ping blocked or failed, attempting HTTP connectivity check..."

    if command_exists curl; then
        if curl -fsI https://www.cloudflare.com >/dev/null 2>&1 || curl -fsI https://1.1.1.1 >/dev/null 2>&1; then
            log_success "Internet connection: OK (HTTP)"
            return 0
        fi
    elif command_exists wget; then
        if wget --spider -q https://www.cloudflare.com >/dev/null 2>&1 || wget --spider -q https://1.1.1.1 >/dev/null 2>&1; then
            log_success "Internet connection: OK (HTTP)"
            return 0
        fi
    fi

    log_error "No internet connection detected. Please check your network."
    log_error "This script requires internet to download Docker, Tailscale, and updates."
    return 1
}

# Get current IP address of the Pi
get_current_ip() {
    local default_iface=$(ip route | grep default | awk '{print $5}' | head -1)
    if [[ -n "$default_iface" ]]; then
        ip -4 addr show "$default_iface" | grep -oP '(?<=inet\s)\d+(\.\d+){3}' | head -1
    fi
}

# Get the IP address of the SSH client (if connected via SSH)
get_ssh_client_ip() {
    # Method 1: Check SSH_CLIENT environment variable (most reliable)
    if [[ -n "$SSH_CLIENT" ]]; then
        echo "$SSH_CLIENT" | awk '{print $1}'
        return 0
    fi
    
    # Method 2: Check SSH_CONNECTION environment variable
    if [[ -n "$SSH_CONNECTION" ]]; then
        echo "$SSH_CONNECTION" | awk '{print $1}'
        return 0
    fi
    
    # Not an SSH session or unable to detect
    return 1
}

# Check if an IP address is within a CIDR subnet
ip_in_subnet() {
    local ip="$1"
    local subnet="$2"
    
    # Use ipcalc if available for accurate checking
    if command_exists ipcalc; then
        ipcalc -c "$ip" "$subnet" >/dev/null 2>&1
        return $?
    fi
    
    log_warning "ipcalc not available; cannot verify subnet safely"
    return 1
}

# Check if current SSH connection will work with proposed firewall rules
verify_ssh_access_with_firewall() {
    local lan_subnet="$1"
    
    # Get SSH client IP
    local ssh_client_ip=$(get_ssh_client_ip)
    
    if [[ -z "$ssh_client_ip" ]]; then
        log_info "Not connected via SSH or unable to detect client IP"
        log_info "Firewall rules will allow access from: ${lan_subnet} and Tailscale"
        return 0  # Not SSH, assume user knows what they're doing
    fi
    
    log_info "Detected SSH connection from: ${ssh_client_ip}"
    
    # Check if SSH client is in the LAN subnet
    if ip_in_subnet "$ssh_client_ip" "$lan_subnet"; then
        log_success "✓ SSH client IP is within allowed LAN subnet (${lan_subnet})"
        return 0
    fi
    
    # Check if SSH client is a Tailscale IP (100.x.x.x)
    if [[ "$ssh_client_ip" =~ ^100\. ]]; then
        if tailscale status >/dev/null 2>&1; then
            log_success "✓ SSH client appears to be a Tailscale IP and Tailscale is connected"
            return 0
        else
            log_warning "SSH client appears to be Tailscale IP but Tailscale is not connected"
            log_warning "You may lose access if Tailscale doesn't reconnect automatically"
        fi
    fi
    
    # SSH client is NOT in allowed subnets
    log_error "WARNING: SSH client IP ${ssh_client_ip} is NOT in allowed subnet ${lan_subnet}"
    log_error "YOU WILL BE LOCKED OUT if you apply these firewall rules!"
    echo
    log_info "Your options:"
    log_info "  1. Cancel and reconnect from LAN (${lan_subnet})"
    log_info "  2. Cancel and fix subnet detection"
    log_info "  3. Connect via Tailscale and ensure it's working"
    log_info "  4. Skip firewall configuration (not recommended)"
    echo
    
    return 1
}

# Detect LAN subnet from active network interface
detect_lan_subnet() {
    log_info "Detecting LAN subnet from active network interfaces..."
    
    # Ensure ipcalc is available for reliable subnet detection
    if ! command_exists ipcalc; then
        log_info "ipcalc not found, installing..."
        if ! retry_apt_update || ! retry_apt_install ipcalc >/dev/null 2>&1; then
            log_error "Failed to install ipcalc package"
            log_error "Cannot safely detect network subnet without ipcalc"
            return 1
        fi
        log_success "ipcalc installed successfully"
    fi
    
    # Check for multiple active network interfaces
    local active_interfaces=($(ip route | grep default | awk '{print $5}'))
    
    if [[ ${#active_interfaces[@]} -gt 1 ]]; then
        log_warning "Multiple active network interfaces detected: ${active_interfaces[*]}"
        log_info "This may affect subnet detection accuracy"
    fi
    
    # Get the default route interface
    local default_iface=$(ip route | grep default | awk '{print $5}' | head -1)
    
    if [[ -z "$default_iface" ]]; then
        log_warning "Could not detect default network interface"
        return 1
    fi
    
    log_info "Default network interface: ${default_iface}"
    
    # Get IP address and netmask for the interface
    local ip_info=$(ip -4 addr show "$default_iface" | grep -oP '(?<=inet\s)\d+(\.\d+){3}/\d+' | head -1)
    
    if [[ -z "$ip_info" ]]; then
        log_warning "Could not get IP address from ${default_iface}"
        return 1
    fi
    
    local ip_address=$(echo "$ip_info" | cut -d'/' -f1)
    local cidr=$(echo "$ip_info" | cut -d'/' -f2)
    
    # Calculate network address from IP and CIDR using ipcalc
    local network_subnet=""
    network_subnet=$(ipcalc -n "$ip_info" | grep -oP '(?<=Network:\s+)\S+')
    
    if [[ -z "$network_subnet" ]]; then
        log_error "Failed to calculate network subnet"
        return 1
    fi
    
    log_success "Detected LAN subnet: ${network_subnet}"
    log_info "  Interface: ${default_iface}"
    log_info "  Your IP: ${ip_address}/${cidr}"
    
    # If multiple interfaces, show all detected subnets
    if [[ ${#active_interfaces[@]} -gt 1 ]]; then
        echo
        log_info "All active interfaces:"
        for iface in "${active_interfaces[@]}"; do
            local iface_ip=$(ip -4 addr show "$iface" | grep -oP '(?<=inet\s)\d+(\.\d+){3}/\d+' | head -1)
            if [[ -n "$iface_ip" ]]; then
                local iface_subnet=""
                iface_subnet=$(ipcalc -n "$iface_ip" | grep -oP '(?<=Network:\s+)\S+')
                log_info "  - ${iface}: ${iface_ip} -> ${iface_subnet}"
            fi
        done
        echo
        log_warning "Using primary interface (${default_iface}) subnet: ${network_subnet}"
    fi
    
    # Export for use in other functions
    export DETECTED_LAN_SUBNET="$network_subnet"
    export DETECTED_IP="$ip_address"
    return 0
}

################################################################################
# MODE 1: INSTALL-DOCKER
# Install Docker and deploy the homelab stack
################################################################################

mode_install_docker() {
    print_header "INSTALL-DOCKER"
    
    # Check prerequisites - install-docker mode has specific requirements
    # but mostly installs its own dependencies (Docker, Tailscale)
    if ! check_all_prerequisites "install-docker"; then
        err "Prerequisites check failed. Cannot continue."
    fi
    
    # Check if we're on USB storage (recommended but not required)
    local root_device=$(findmnt -n -o SOURCE /)
    if [[ "$root_device" =~ "mmcblk" ]]; then
        log_warning "You appear to be running from SD card"
        log_warning "It's recommended to migrate to USB storage first for better performance"
        if ! prompt_yn "Continue anyway?"; then
            log_info "Installation cancelled by user"
            exit 0
        fi
    fi
    
    check_internet || exit 1
    
    log_step "Creating directory structure"
    create_directory_safe "${HOMELAB_DIR}/pihole/etc-pihole" "Pi-hole configuration" || err "Failed to create Pi-hole directory"
    create_directory_safe "${HOMELAB_DIR}/homeassistant" "Home Assistant configuration" || err "Failed to create Home Assistant directory"
    create_directory_safe "${COMPOSE_DIR}" "Docker Compose files" || err "Failed to create Docker Compose directory"
    create_directory_safe "${BACKUP_DIR}" "Backup storage" || err "Failed to create backup directory"
    
    log_success "Created ${HOMELAB_DIR} directory structure"
    
    log_step "Checking for port conflicts"
    
    # Check port 80 and 443 for web servers
    log_info "Checking for web server conflicts on ports 80/443..."
    local port_80_used=""
    local port_443_used=""
    
    # Use ss if available (modern), fallback to netstat (legacy)
    if command_exists ss; then
        port_80_used=$(ss -tulpn 2>/dev/null | grep ':80 ' || true)
        port_443_used=$(ss -tulpn 2>/dev/null | grep ':443 ' || true)
    elif command_exists netstat; then
        port_80_used=$(netstat -tulpn 2>/dev/null | grep ':80 ' || true)
        port_443_used=$(netstat -tulpn 2>/dev/null | grep ':443 ' || true)
    else
        log_warning "Neither 'ss' nor 'netstat' available - skipping port conflict check"
        log_info "Install iproute2 or net-tools for port checking"
    fi
    
    if [[ -n "$port_80_used" ]]; then
        log_warning "Port 80 is already in use:"
        echo "$port_80_used"
        log_warning "Pi-hole requires port 80 for its web interface"
        
        # Check for common web servers
        for service in apache2 nginx lighttpd; do
            if systemctl is-active --quiet "$service" 2>/dev/null; then
                log_warning "Found running service: $service"
                if prompt_yn "Stop and disable $service?" true; then
                    systemctl stop "$service"
                    systemctl disable "$service"
                    log_success "$service stopped and disabled"
                fi
            fi
        done
    fi
    
    if [[ -n "$port_443_used" ]]; then
        log_info "Port 443 is in use (Pi-hole will use HTTP only; HTTPS requires port 443)"
    fi
    
    # Check if systemd-resolved is using port 53
    # systemd-resolved runs a DNS stub listener on port 53 which conflicts with Pi-hole
    # Pi-hole needs exclusive access to port 53 to function as the network's DNS server
    if systemctl is-active --quiet systemd-resolved; then
        log_warning "systemd-resolved is running and will conflict with Pi-hole (port 53)"
        log_info "Pi-hole requires exclusive access to port 53 for DNS queries"
        log_info "systemd-resolved can be re-enabled later if you uninstall Pi-hole"
        echo
        if prompt_yn "Disable systemd-resolved to allow Pi-hole to use port 53?" true; then
            log_info "Stopping and disabling systemd-resolved..."
            systemctl disable systemd-resolved
            systemctl stop systemd-resolved
            
            # Replace resolv.conf with static DNS (using Cloudflare)
            rm -f /etc/resolv.conf
            echo "nameserver 1.1.1.1" > /etc/resolv.conf
            echo "nameserver 1.0.0.1" >> /etc/resolv.conf
            
            log_success "systemd-resolved disabled"
            log_info "To re-enable later: sudo systemctl enable systemd-resolved && sudo systemctl start systemd-resolved"
        else
            log_error "Cannot proceed - Pi-hole requires port 53"
            log_info "Options:"
            log_info "  1. Disable systemd-resolved (recommended)"
            log_info "  2. Configure systemd-resolved to not use port 53 (advanced)"
            log_info "  3. Cancel installation"
            err "Pi-hole installation cancelled - port 53 conflict"
        fi
    fi
    
    log_step "Installing Docker"
    
    if command_exists docker; then
        log_info "Docker is already installed"
        docker --version
    else
        log_info "Downloading Docker installation script..."
        # NOTE: Docker's get.docker.com uses HTTPS for transport security (no checksums published)
        if ! retry_until_connected --attempts ${CURL_RETRY_ATTEMPTS} --delay ${CURL_RETRY_DELAY} curl -fsSL https://get.docker.com -o /tmp/get-docker.sh; then
            err "Failed to download Docker installation script after multiple attempts. Check internet connection and try again."
        fi
        
        log_info "Installing Docker..."
        if ! sh /tmp/get-docker.sh; then
            rm /tmp/get-docker.sh
            log_error "Docker installation failed"
            log_error "Possible causes:"
            log_error "  1. Unsupported OS version (must be Raspberry Pi OS Bookworm/Trixie)"
            log_error "  2. Network issues during package download"
            log_error "  3. Insufficient disk space (check: df -h /)"
            log_error "  4. Repository key verification failure"
            log_error "Troubleshooting:"
            log_error "  - Check install log: cat /tmp/get-docker.sh"
            log_error "  - Manual install: https://docs.docker.com/engine/install/debian/"
            log_error "  - Check system logs: journalctl -xe"
            err "Docker installation failed"
        fi
        
        # Add the user who invoked sudo to docker group
        local target_user="${SUDO_USER:-}"
        
        if [[ -z "$target_user" ]]; then
            log_warning "Could not determine user (script not run via sudo)"
            log_info "Add your user to docker group manually: sudo usermod -aG docker <username>"
        elif ! id "$target_user" &>/dev/null; then
            # Check user existence first (inverted logic for clarity)
            log_error "User '${target_user}' does not exist"
            log_error "SUDO_USER environment variable may be incorrect"
            log_info "Add your user to docker group manually: sudo usermod -aG docker <username>"
        else
            # User exists, try to add to docker group
            log_info "Adding user '${target_user}' to docker group..."
            if usermod -aG docker "$target_user" 2>/dev/null; then
                log_success "User '${target_user}' added to docker group"
                log_warning "User '${target_user}' must log out and back in for group changes to take effect"
                log_info "Or run: newgrp docker (in current session only)"
            else
                log_error "Failed to add '${target_user}' to docker group"
                log_error "This may cause permission issues when running docker commands"
                log_info "Try manually: sudo usermod -aG docker ${target_user}"
            fi
        fi
        
        rm /tmp/get-docker.sh
        
        # Verify Docker was installed successfully
        if ! command_exists docker; then
            err "Docker command not found after installation"
        fi
        
        log_success "Docker installed successfully"
        docker --version
    fi
    
    # Verify Docker daemon is running
    log_info "Verifying Docker daemon status..."
    if ! verify_docker_running; then
        log_info "Starting Docker daemon..."
        systemctl start docker
        sleep $DOCKER_VERIFY_DELAY
        if ! verify_docker_running; then
            ERROR_HINT="Check Docker status: systemctl status docker"
            err "Docker daemon failed to start"
        fi
    fi
    log_success "Docker daemon is running"
    
    log_step "Installing Docker Compose plugin"
    
    if docker compose version >/dev/null 2>&1; then
        log_info "Docker Compose plugin already installed"
        docker compose version
    else
        log_info "Installing docker-compose-plugin..."
        if ! retry_apt_update; then
            err "Failed to update package lists"
        fi
        if ! retry_apt_install docker-compose-plugin; then
            log_error "Try running: sudo apt-get update && sudo apt-get install docker-compose-plugin"
            err "Failed to install docker-compose-plugin"
        fi
        
        # Verify installation
        if ! docker compose version >/dev/null 2>&1; then
            err "Docker Compose plugin not working after installation"
        fi
        
        log_success "Docker Compose plugin installed"
        docker compose version
    fi
    
    log_step "Installing Markdown viewer (glow)"
    
    if command_exists glow; then
        log_info "glow is already installed"
        glow --version
    else
        log_info "Installing glow for reading documentation..."
        if ! retry_apt_install glow 2>/dev/null; then
            log_warning "glow not available in apt repos, trying manual installation..."
            
            # Detect architecture
            local arch=$(dpkg --print-architecture)
            local glow_version="1.5.1"
            local glow_url=""
            
            case "$arch" in
                arm64|aarch64)
                    glow_url="https://github.com/charmbracelet/glow/releases/download/v${glow_version}/glow_${glow_version}_linux_arm64.tar.gz"
                    ;;
                armhf|armv7l)
                    glow_url="https://github.com/charmbracelet/glow/releases/download/v${glow_version}/glow_${glow_version}_linux_armv7.tar.gz"
                    ;;
                *)
                    log_warning "Architecture ${arch} not supported for glow binary download"
                    log_info "You can still read markdown with: less /root/README.md"
                    glow_url=""
                    ;;
            esac
            
            if [[ -n "$glow_url" ]]; then
                local temp_dir="$(mktemp -d)"
                if [[ -z "$temp_dir" ]] || [[ ! -d "$temp_dir" ]]; then
                    log_warning "Failed to create temporary directory for glow installation"
                    log_info "You can still read markdown with: less /root/README.md"
                elif retry_until_connected --attempts ${CURL_RETRY_ATTEMPTS} --delay ${CURL_RETRY_DELAY} \
                    curl -fsSL "$glow_url" -o "${temp_dir}/glow.tar.gz" && \
                   tar -xzf "${temp_dir}/glow.tar.gz" -C "${temp_dir}" && \
                   mv "${temp_dir}/glow" /usr/local/bin/glow && \
                   chmod +x /usr/local/bin/glow; then
                    log_success "glow installed from GitHub releases"
                    [[ -n "$temp_dir" ]] && rm -rf "${temp_dir}"
                else
                    log_warning "Failed to install glow manually"
                    log_info "You can still read markdown with: less /root/README.md"
                    [[ -n "$temp_dir" ]] && rm -rf "${temp_dir}"
                fi
            fi
        else
            log_success "glow installed successfully"
            glow --version
        fi
    fi
    
    log_step "Detecting timezone"
    local timezone="UTC"
    
    if [[ -f /etc/timezone ]]; then
        # File exists, read and clean content (remove whitespace/newlines)
        local detected_tz=$(cat /etc/timezone 2>/dev/null | tr -d '\n\r' | xargs)
        
        if [[ -n "$detected_tz" ]]; then
            timezone="$detected_tz"
            log_info "Detected timezone from /etc/timezone: ${timezone}"
        else
            log_warning "/etc/timezone exists but is empty"
            log_warning "Defaulting to UTC - services will use UTC time"
            log_info "Fix with: sudo timedatectl set-timezone America/New_York (or your timezone)"
        fi
    else
        log_warning "Could not find /etc/timezone file"
        log_warning "Defaulting to UTC - services will use UTC time"
        log_info "Fix with: sudo timedatectl set-timezone America/New_York (or your timezone)"
    fi
    
    log_step "Creating Pi-hole Docker Compose file"
    
        if ! write_pihole_compose "${COMPOSE_DIR}/pihole-compose.yml" "${timezone}"; then
                err "Failed to create Pi-hole compose file"
        fi
    
    log_success "Created Pi-hole compose file"
    
    log_step "Creating Home Assistant Docker Compose file"
    
        if ! write_homeassistant_compose "${COMPOSE_DIR}/homeassistant-compose.yml" "${timezone}"; then
                err "Failed to create Home Assistant compose file"
        fi
    
    log_success "Created Home Assistant compose file"
    
    log_step "Starting Pi-hole"
    log_info "Pulling Pi-hole image and starting container..."
    
    cd "${COMPOSE_DIR}" || err "Failed to change to compose directory: ${COMPOSE_DIR}"
    if ! pull_and_verify_docker_image "${PIHOLE_IMAGE}" "pihole-compose.yml" ""; then
        log_error "Failed to pull Pi-hole image after multiple attempts"
        echo
        log_error "Troubleshooting:"
        log_error "  1. Check internet: ping -c 3 1.1.1.1"
        log_error "  2. Test Docker Hub: curl -I https://hub.docker.com"
        log_error "  3. Check disk space: df -h /var/lib/docker"
        log_error "  4. View Docker logs: journalctl -u docker -n 50"
        echo
        log_info "Manual recovery:"
        log_info "  - Wait a few minutes and re-run: sudo $0 install-docker"
        log_info "  - Try different DNS: echo 'nameserver 8.8.8.8' > /etc/resolv.conf"
        log_info "  - Use mobile hotspot if ISP is throttling"
        err "Pi-hole image pull failed"
    fi

    if ! docker compose -f pihole-compose.yml up -d; then
        log_error "Failed to start Pi-hole container"
        echo
        log_info "Common causes:"
        log_info "  1. Port 53 conflict - check: sudo ss -tulpn | grep :53"
        log_info "  2. Port 80 conflict - check: sudo ss -tulpn | grep :80"
        log_info "  3. systemd-resolved still running - check: systemctl status systemd-resolved"
        echo
        ERROR_HINT="View detailed logs: docker logs pihole"
        err "Pi-hole container failed to start"
    fi
    
    # Verify container actually started (with retry - container needs time to initialize)
    log_info "Verifying Pi-hole container started..."
    # Container startup typically takes 2-5 seconds
    if ! retry_until_connected --attempts ${CONNECTION_RETRY_MAX_ATTEMPTS} --delay ${CONNECTION_RETRY_DELAY} is_container_running pihole; then
        log_error "Pi-hole container is not running after docker compose up"
        log_error "Container may have started and immediately crashed"
        log_error "Check logs: docker logs pihole"
        err "Pi-hole container failed to start or crashed immediately"
    fi
    
    log_info "Waiting for Pi-hole to be ready..."
    
    # Helper function to check if Pi-hole is responding
    check_pihole_ready() {
        curl -s http://localhost/admin/ >/dev/null 2>&1
    }
    
    # Pi-hole container startup: lighttpd web server initialization
    # Typical: 5-15 seconds, worst case under load: 30-45 seconds
    # Use fixed polling to check readiness (Pi-hole needs time to initialize)
    if retry_until_connected --attempts ${PIHOLE_READY_CHECK_ATTEMPTS} --delay ${PIHOLE_READY_CHECK_DELAY} check_pihole_ready; then
        local current_ip=$(get_current_ip)
        log_success "Pi-hole is running!"
        if [[ -n "$current_ip" ]]; then
            log_info "Web interface: http://${current_ip}/admin"
        else
            log_info "Web interface: http://localhost/admin (or use your Pi's IP address)"
        fi
        log_info "Password: ${PIHOLE_PASSWORD}"
        if [[ "$current_ip" != "$PIHOLE_STATIC_IP" ]] && [[ -n "$current_ip" ]]; then
            log_warning "Note: Configure static IP ${PIHOLE_STATIC_IP} in your router for consistent access"
        fi
    else
        log_warning "Pi-hole may still be starting up. Check with: docker logs pihole"
        log_warning "If it doesn't start within 2 minutes, there may be a configuration issue"
    fi
    
    log_step "Adding Hagezi blocklist to Pi-hole"
    
    # Note: Pi-hole v6 API requires authentication which complicates automation
    # Using CLI approach is more reliable for initial setup
    log_info "Adding Hagezi blocklist via Pi-hole CLI..."
    
    # Add blocklist using sqlite3 directly (Pi-hole's database)
    if docker exec pihole sqlite3 /etc/pihole/gravity.db \
        "INSERT OR IGNORE INTO adlist (address, enabled, comment) VALUES ('${HAGEZI_BLOCKLIST}', 1, 'Hagezi Pro Blocklist');" 2>/dev/null; then
        log_info "Blocklist added to database, updating gravity..."
        if docker exec pihole pihole -g >/dev/null 2>&1; then
            log_success "Hagezi blocklist added and gravity updated"
        else
            log_warning "Blocklist added but gravity update failed"
            log_info "Run manually: docker exec pihole pihole -g"
        fi
    else
        log_warning "Failed to add blocklist automatically"
        log_info "You can add it manually in the web UI:"
        log_info "  1. Go to: http://${PIHOLE_STATIC_IP}/admin"
        log_info "  2. Navigate to: Adlists → Add New Adlist"
        log_info "  3. Paste: ${HAGEZI_BLOCKLIST}"
        log_info "  4. Click 'Add' and then update gravity"
        echo
        if ! prompt_yn "Continue with setup?"; then
            err "Setup cancelled by user"
        fi
    fi

    log_step "Starting Home Assistant"
    log_info "Pulling Home Assistant image and starting container..."
    log_warning "First start may take several minutes to initialize..."
    log_warning "Home Assistant image is 1.5-2GB - this will take a while on slow connections"
    
    # Home Assistant image is ~1.5-2GB compressed - much larger than Pi-hole!
    # User has full control via interactive progress monitoring
    if ! pull_and_verify_docker_image "${HOMEASSISTANT_IMAGE}" "homeassistant-compose.yml" " (large image, may take 30-60 min)"; then
        log_error "Failed to pull Home Assistant image after multiple attempts"
        echo
        log_error "Troubleshooting:"
        log_error "  1. Check internet: ping -c 3 1.1.1.1"
        log_error "  2. Test GitHub Registry: curl -I https://ghcr.io"
        log_error "  3. Check disk space: df -h /var/lib/docker (need ~4GB free)"
        log_error "  4. View Docker logs: journalctl -u docker -n 50"
        echo
        log_info "Manual recovery:"
        log_info "   1. Wait for better network conditions and re-run: sudo $0 install-docker"
        log_info "   2. Download on another machine and transfer:"
        log_info "      a. On machine with image: docker save ${HOMEASSISTANT_IMAGE} | gzip > ha.tar.gz"
        log_info "      b. Transfer to Pi, then: gunzip -c ha.tar.gz | docker load"
        err "Home Assistant image pull failed"
    fi
    
    if ! docker compose -f homeassistant-compose.yml up -d; then
        log_error "Failed to start Home Assistant container"
        echo
        log_info "Common causes:"
        log_info "  1. Port 8123 conflict - check: sudo ss -tulpn | grep :8123"
        log_info "  2. Insufficient disk space - check: df -h ${HOMELAB_DIR}"
        log_info "  3. Permission issues - check: ls -ld ${HOMELAB_DIR}/homeassistant"
        echo
        ERROR_HINT="View detailed logs: docker logs homeassistant"
        err "Home Assistant container failed to start"
    fi
    
    # Verify container actually started (with retry - container needs time to initialize)
    log_info "Verifying Home Assistant container started..."
    # Container startup typically takes 2-5 seconds
    if ! retry_until_connected --attempts ${CONNECTION_RETRY_MAX_ATTEMPTS} --delay ${CONNECTION_RETRY_DELAY} is_container_running homeassistant; then
        log_error "Home Assistant container is not running after docker compose up"
        log_error "Container may have started and immediately crashed"
        log_error "Check logs: docker logs homeassistant"
        err "Home Assistant container failed to start or crashed immediately"
    fi
    
    log_info "Waiting for Home Assistant to be ready..."
    
    # Helper function to check if Home Assistant is responding
    check_ha_ready() {
        curl -s http://localhost:8123/ >/dev/null 2>&1
    }
    
    # Home Assistant first boot: unpacks Python env, initializes DB, loads integrations
    # Pi 5 + SSD: 2-4 minutes | Pi 4 + good SD: 3-5 minutes | Pi 4 + slow SD: 10-15 minutes
    # Use fixed 10s polling to check every 10 seconds (up to 90 checks in 15 minutes)
    if retry_until_connected --attempts ${HA_READY_CHECK_ATTEMPTS} --delay ${HA_READY_CHECK_DELAY} check_ha_ready; then
        local current_ip=$(get_current_ip)
        log_success "Home Assistant is running!"
        if [[ -n "$current_ip" ]]; then
            log_info "Web interface: http://${current_ip}:8123"
        else
            log_info "Web interface: http://localhost:8123 (or use your Pi's IP address)"
        fi
    else
        log_warning "Home Assistant may still be starting up. Check with: docker logs homeassistant"
        log_warning "First boot can take 5-10 minutes. Be patient!"
    fi
    
    log_step "Creating Homeway Docker Compose file"
    
    # Ensure Homeway directory exists
    create_directory_safe "${HOMELAB_DIR}/homeway" "Homeway data" || err "Failed to create Homeway directory"
    
        if ! write_homeway_compose "${COMPOSE_DIR}/homeway-compose.yml"; then
                err "Failed to create Homeway compose file"
        fi
    
    log_success "Created Homeway compose file (requires token configuration in next step)"
    
    # Create placeholder .env file
    if [[ ! -f "${COMPOSE_DIR}/.env" ]]; then
        echo "# Docker Compose environment variables" > "${COMPOSE_DIR}/.env"
        echo "# This file will be populated during configure mode" >> "${COMPOSE_DIR}/.env"
        echo "HOMEWAY_TOKEN=" >> "${COMPOSE_DIR}/.env"
        chmod 600 "${COMPOSE_DIR}/.env"
        log_info "Created ${COMPOSE_DIR}/.env placeholder file"
    fi
    
    log_step "Installing Tailscale"
    
    if command_exists tailscale; then
        log_info "Tailscale is already installed"
        tailscale version
    else
        log_info "Installing Tailscale..."
        # NOTE: Tailscale's install.sh uses HTTPS for transport security (open-source: github.com/tailscale/tailscale)
        if ! retry_until_connected --attempts ${CURL_RETRY_ATTEMPTS} --delay ${CURL_RETRY_DELAY} curl -fsSL https://tailscale.com/install.sh -o /tmp/tailscale-install.sh; then
            err "Failed to download Tailscale installation script after multiple attempts"
        fi
        if ! sh /tmp/tailscale-install.sh; then
            err "Failed to install Tailscale"
        fi
        rm /tmp/tailscale-install.sh
        log_success "Tailscale installed (requires configuration in next step)"
    fi
    
    log_step "Enabling Docker to start on boot"
    systemctl enable docker
    
    echo
    print_separator
    log_success "Docker stack installation complete!"
    print_separator
    echo
    local current_ip=$(get_current_ip)
    local display_ip="${current_ip:-<your-pi-ip>}"
    log_info "Services installed:"
    log_info "  ✓ Pi-hole:         http://${display_ip}/admin (password: ${PIHOLE_PASSWORD})"
    log_info "  ✓ Home Assistant:  http://${display_ip}:8123"
    log_info "  ✓ Homeway:         Installed (needs configuration)"
    log_info "  ✓ Tailscale:       Installed (needs configuration)"
    if [[ "$current_ip" != "$PIHOLE_STATIC_IP" ]] && [[ -n "$current_ip" ]]; then
        echo
        log_warning "Current IP: ${current_ip}"
        log_warning "Recommended: Configure static IP ${PIHOLE_STATIC_IP} in your router"
    fi
    echo
    log_info "Documentation:"
    if command_exists glow; then
        log_info "  Read docs with:    glow /root/README.md"
        log_info "  Troubleshooting:   glow /root/TROUBLESHOOTING.md"
    else
        log_info "  Read docs with:    less /root/README.md"
        log_info "  Troubleshooting:   less /root/TROUBLESHOOTING.md"
    fi
    echo
    log_info "Next steps:"
    log_info "  1. Configure Tailscale and collect credentials: sudo $0 configure"
    log_info "  2. Set up your router to use Pi-hole as DNS (${PIHOLE_STATIC_IP})"
    log_info "  3. Complete Home Assistant onboarding in web browser"
    echo
}

################################################################################
# MODE 2: CONFIGURE
# Post-setup configuration (collect tokens and keys)
################################################################################

mode_configure() {
    print_header "CONFIGURE"
    
    # Check prerequisites for this mode
    if ! check_all_prerequisites "configure"; then
        err "Prerequisites check failed. Cannot continue."
    fi
    
    log_info "This mode will configure Tailscale and collect credentials for Homeway"
    echo
    
    # Check if services are running
    log_step "Checking service status"
    report_container_status
    
    # Configure Tailscale
    log_step "Configuring Tailscale"
    
    if tailscale status >/dev/null 2>&1; then
        log_info "Tailscale is already configured and connected"
        tailscale status
        echo
        if ! prompt_yn "Reconfigure Tailscale?"; then
            log_info "Skipping Tailscale configuration"
        else
            tailscale down
        fi
    fi
    
    if ! tailscale status >/dev/null 2>&1; then
        echo
        log_info "To get a Tailscale auth key:"
        log_info "  1. Go to: https://login.tailscale.com/admin/settings/keys"
        log_info "  2. Click 'Generate auth key'"
        log_info "  3. Copy the key (starts with 'tskey-auth-...')"
        echo
        
        read -p "Enter Tailscale auth key: " tailscale_key
        
        if [[ -z "$tailscale_key" ]]; then
            log_warning "No auth key provided, skipping Tailscale configuration"
        else
            log_info "Connecting to Tailscale..."
            
            # Configure with SSH enabled and DNS disabled (Pi-hole handles DNS)
            tailscale up --authkey="$tailscale_key" --ssh --accept-dns=false
            
            if tailscale status >/dev/null 2>&1; then
                log_success "Tailscale connected successfully!"
                echo
                tailscale status
                echo
                log_info "Your Tailscale IP and MagicDNS hostname are shown above"
            else
                log_error "Failed to connect to Tailscale"
                log_error "Possible causes:"
                log_error "  1. Invalid or expired auth key"
                log_error "  2. Network connectivity issues (check: ping 1.1.1.1)"
                log_error "  3. Firewall blocking UDP port 41641"
                log_error "  4. Time sync issues (check: timedatectl status)"
                log_error "Troubleshooting:"
                log_error "  - Check Tailscale logs: journalctl -u tailscaled -n 50"
                log_error "  - Generate new key: https://login.tailscale.com/admin/settings/keys"
                log_error "  - Test connectivity: tailscale ping <another-node>"
            fi
        fi
    fi
    
    # Configure firewall
    log_step "Configuring firewall (nftables)"
    
    echo
    log_info "The firewall will restrict access to:"
    log_info "  • Your LAN subnet (auto-detected)"
    log_info "  • Tailscale VPN"
    log_info "  • All other networks will be blocked"
    echo
    log_warning "Note: Raspberry Pi OS has no firewall by default (all access allowed)"
    echo
    
    if ! prompt_yn "Do you want to configure the firewall?" true; then
        log_info "Skipping firewall configuration"
        log_warning "Your Pi will be accessible from all networks (default Pi OS behavior)"
        log_info "You can configure firewall later by running: sudo $0 configure"
    else
        # Detect LAN subnet dynamically
        echo
        local lan_subnet=""
        if detect_lan_subnet; then
            lan_subnet="$DETECTED_LAN_SUBNET"
            log_success "Auto-detected LAN subnet: ${lan_subnet}"
            echo
            log_warning "CRITICAL: Firewall will ONLY allow access from this subnet and Tailscale"
            log_warning "If this subnet is incorrect, you will be LOCKED OUT after applying firewall rules!"
            log_info "Your current IP: ${DETECTED_IP}"
            log_info "Detected subnet: ${lan_subnet}"
            echo
            
            if ! prompt_yn "Is ${lan_subnet} correct for your LAN?" true; then
                log_warning "Subnet detection appears incorrect."
                lan_subnet=""
            fi
        fi
        
        # Manual subnet entry if auto-detection failed or user rejected
        if [ -z "$lan_subnet" ]; then
            log_warning "Failed to auto-detect LAN subnet or detection was rejected"
            echo
            read -p "Enter your LAN subnet manually (e.g., 192.168.1.0/24): " lan_subnet
            
            if [ -z "$lan_subnet" ]; then
                log_error "No subnet provided. Skipping firewall configuration."
                log_info "You can configure firewall later by running: sudo $0 configure"
                return 0
            fi
        fi
        
        # Validate subnet format
        if [[ ! "$lan_subnet" =~ ^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+/[0-9]+$ ]]; then
            log_error "Invalid subnet format: ${lan_subnet}"
            log_error "Expected format: 192.168.1.0/24"
            log_warning "Skipping firewall configuration to prevent lockout"
            log_info "You can manually configure nftables later"
            return 0
        fi
        
        # Validate that it's a proper network address (not a host IP)
        if command_exists ipcalc; then
            local network_check=$(ipcalc -n "$lan_subnet" 2>/dev/null | grep -oP '(?<=Network:\s+)\S+')
            if [[ -n "$network_check" ]] && [[ "$network_check" != "$lan_subnet" ]]; then
                log_warning "Subnet address validation: ${lan_subnet} != ${network_check}"
                log_warning "You may have entered a host IP instead of network address"
                log_info "For 192.168.1.x network, use: 192.168.1.0/24 (not 192.168.1.5/24)"
                echo
                if ! prompt_yn "Use detected network address ${network_check} instead?" true; then
                    log_warning "Keeping user-provided subnet: ${lan_subnet}"
                    log_warning "Firewall may not work as expected!"
                else
                    lan_subnet="$network_check"
                    log_info "Using corrected subnet: ${lan_subnet}"
                fi
            fi
        fi

        # Validate that current IP is within the chosen subnet (automatic safety check)
        local current_ip_for_check="${DETECTED_IP:-$(get_current_ip)}"
        if command_exists ipcalc && [[ -n "$current_ip_for_check" ]]; then
            if ! ipcalc -c "$current_ip_for_check" "$lan_subnet" >/dev/null 2>&1; then
                log_warning "Current IP (${current_ip_for_check}) is NOT within subnet ${lan_subnet}"
                log_warning "Applying firewall rules may lock you out"
                if prompt_yn "Enter a different subnet manually?" true; then
                    read -p "Enter your LAN subnet (e.g., 192.168.1.0/24): " lan_subnet
                fi
            fi
        fi
        
        log_info "Using LAN subnet: ${lan_subnet}"
        echo
        
        # Verify SSH access before applying firewall rules
        log_step "Verifying SSH access compatibility"
        if ! verify_ssh_access_with_firewall "$lan_subnet"; then
            if ! prompt_yn "Apply firewall rules anyway? (YOU MAY BE LOCKED OUT)"; then
                log_info "Firewall configuration cancelled for safety"
                log_info "You can configure firewall later by running: sudo $0 configure"
                return 0
            else
                log_warning "Proceeding with firewall configuration despite SSH access warning..."
            fi
        fi
        
        # Final warning before applying firewall rules
        print_box "⚠  CRITICAL WARNING  ⚠" \
            "" \
            "You are about to apply firewall rules that will:" \
            "  • BLOCK all access except from: ${lan_subnet}" \
            "  • BLOCK all access except via: Tailscale VPN" \
            "" \
            "If the subnet is WRONG, you will be IMMEDIATELY LOCKED OUT via SSH" \
            "" \
            "═══════════════════════════════════════════════════════════════════════" \
            "RECOVERY IF LOCKED OUT:" \
            "  1. Power off Raspberry Pi" \
            "  2. Remove SD card and mount on another computer" \
            "  3. Mount the ROOT partition (usually partition 2, labeled 'rootfs')" \
            "  4. Delete file: /etc/nftables.conf from the root filesystem" \
            "  5. OR: Re-flash Raspberry Pi OS to SD card (15 minutes)" \
            "═══════════════════════════════════════════════════════════════════════" \
            "" \
            "Current detection:" \
            "  Your IP: ${DETECTED_IP:-unknown}" \
            "  Subnet:  ${lan_subnet}" \
            ""
        
        if ! prompt_yn "I understand the risks. Apply firewall rules now?"; then
            log_info "Firewall configuration cancelled"
            log_info "You can configure firewall later by running: sudo $0 configure"
            return 0
        fi
        
        echo
        
        if [[ -f /etc/nftables.conf ]]; then
            log_info "Backing up existing nftables.conf..."
            cp /etc/nftables.conf /etc/nftables.conf.backup
        fi
        
        log_info "Creating firewall rules to allow LAN and Tailscale access..."
    
        if ! write_nftables_conf "$lan_subnet"; then
            log_error "Failed to write /etc/nftables.conf"
            return 0
        fi
        
        # Ensure firewall config is persisted to disk
        sync
        
        # Verify syntax before loading
        if ! nft -c -f /etc/nftables.conf; then
            log_error "nftables.conf syntax validation failed"
            if [[ -f /etc/nftables.conf.backup ]]; then
                log_info "Restoring previous nftables.conf from backup"
                mv /etc/nftables.conf.backup /etc/nftables.conf
            else
                log_info "Removing broken nftables.conf (no previous config to restore)"
                rm -f /etc/nftables.conf
            fi
            log_error "Firewall configuration aborted to prevent lockout"
            return 0
        fi
        
        # Load rules
        if nft -f /etc/nftables.conf; then
            log_success "Firewall rules loaded and verified"
            
            # Clean up backup file after successful configuration
            if [[ -f /etc/nftables.conf.backup ]]; then
                rm -f /etc/nftables.conf.backup
            fi
            
            # Enable nftables service
            systemctl enable nftables
            
            log_info "Firewall configuration:"
            log_info "  ✓ LAN subnet (${lan_subnet}): Full access"
            log_info "  ✓ Tailscale: Full access"
            log_info "  ✓ Other networks: Blocked"
        else
            log_error "Failed to load firewall rules"
            log_error "Possible causes:"
            log_error "  1. Syntax error in nftables.conf (check: nft -c -f /etc/nftables.conf)"
            log_error "  2. nftables not installed (check: which nft)"
            log_error "  3. Kernel module not loaded (check: lsmod | grep nf_tables)"
            log_error "  4. Invalid subnet format: ${lan_subnet}"
            log_error "Troubleshooting:"
            log_error "  - Test syntax: sudo nft -c -f /etc/nftables.conf"
            log_error "  - View current rules: sudo nft list ruleset"
            log_error "  - Check system log: journalctl -xe"
            if [[ -f /etc/nftables.conf.backup ]]; then
                log_info "Restoring backup..."
                mv /etc/nftables.conf.backup /etc/nftables.conf
            else
                log_info "Removing broken nftables.conf (no previous config to restore)"
                rm -f /etc/nftables.conf
            fi
        fi
    fi  # End of "Do you want to configure the firewall?" block
    
    echo
    
    # Collect Home Assistant token
    log_step "Home Assistant Configuration"
    
    local current_ip=$(get_current_ip)
    local display_ip="${current_ip:-<your-pi-ip>}"
    echo
    log_info "To generate a Home Assistant Long-Lived Access Token:"
    log_info "  1. Open Home Assistant: http://${display_ip}:8123"
    log_info "  2. Complete initial onboarding if needed"
    log_info "  3. Click your profile (bottom left)"
    log_info "  4. Scroll to 'Long-Lived Access Tokens'"
    log_info "  5. Click 'Create Token'"
    log_info "  6. Give it a name (e.g., 'Homeway')"
    log_info "  7. Copy the token (you won't see it again!)"
    echo
    
    read -p "Enter Home Assistant Long-Lived Access Token (or press Enter to skip): " ha_token
    
    if [[ -n "$ha_token" ]]; then
        # Save token to .env file for Docker Compose
        log_info "Saving Home Assistant token to ${COMPOSE_DIR}/.env..."
        
        # Ensure compose directory exists
        create_directory_safe "${COMPOSE_DIR}" "Docker Compose directory" || err "Failed to create Docker Compose directory"
        
        if [[ -f "${COMPOSE_DIR}/.env" ]]; then
            # Update existing .env file - remove old token line
            grep -v "^HOMEWAY_TOKEN=" "${COMPOSE_DIR}/.env" > "${COMPOSE_DIR}/.env.tmp" 2>/dev/null || true
            mv "${COMPOSE_DIR}/.env.tmp" "${COMPOSE_DIR}/.env"
            echo "HOMEWAY_TOKEN=${ha_token}" >> "${COMPOSE_DIR}/.env"
        else
            # Create new .env file
            echo "# Docker Compose environment variables" > "${COMPOSE_DIR}/.env"
            echo "HOMEWAY_TOKEN=${ha_token}" >> "${COMPOSE_DIR}/.env"
        fi
        chmod 600 "${COMPOSE_DIR}/.env"
        log_success "Token saved to .env file (persists across container restarts)"
        
        # Configure and start Homeway if compose file exists
        if [[ -f "${COMPOSE_DIR}/homeway-compose.yml" ]]; then
            log_step "Configuring Homeway"
            log_info "Starting Homeway container with Home Assistant token..."
            
            # Verify Docker is running
            if ! verify_docker_running; then
                log_warning "Docker daemon not running. Cannot start Homeway."
                return 0
            fi
            
            cd "${COMPOSE_DIR}" || { log_error "Failed to change to compose directory"; return 0; }
            if [[ -f "${COMPOSE_DIR}/homeway-compose.yml" ]]; then
                if docker compose -f homeway-compose.yml up -d 2>/dev/null; then
                    log_success "Homeway container started successfully!"
                    log_info "Homeway will connect to Home Assistant automatically"
                    log_info "Token persisted in .env - container will restart automatically"
                else
                    log_warning "Failed to start Homeway container. Check logs: docker logs homeway"
                fi
            fi
        else
            log_warning "Homeway compose file not found. Run 'install-docker' mode first."
        fi
    else
        log_warning "No token provided. Homeway will not be configured."
        log_info "You can configure it later by running: sudo $0 configure"
    fi
    
    echo
    
    # Secure root account
    log_step "Securing system accounts"
    
    # Verify there's another account with sudo access
    local sudo_user="${SUDO_USER:-}"
    if [[ -z "$sudo_user" ]]; then
        log_warning "Cannot determine original user account (not run via sudo)"
        log_warning "Skipping root account lockdown for safety"
    else
        log_info "Detected user account: ${sudo_user}"
        log_info "Disabling password login for root account..."
        log_info "This ensures '${sudo_user}' is the only account that can log in with password"
        log_info "(Tailscale will handle SSH authentication separately)"
        echo
        
        if prompt_yn "Disable root password login?" true; then
            # Verify the user exists and has sudo access
            if ! id "$sudo_user" &>/dev/null; then
                log_error "User '${sudo_user}' does not exist - cannot safely lock root"
                log_warning "Skipping root account lockdown"
            elif ! groups "$sudo_user" | grep -q '\bsudo\b'; then
                log_warning "User '${sudo_user}' is not in sudo group"
                log_warning "Locking root would prevent administrative access"
                if ! prompt_yn "Lock root anyway? (NOT RECOMMENDED)"; then
                    log_info "Skipping root account lockdown"
                    sudo_user=""  # Clear to skip locking
                fi
            fi
            
            if [[ -n "$sudo_user" ]]; then
                # Lock the root account password
                passwd -l root >/dev/null 2>&1
                
                if passwd -S root 2>/dev/null | grep -q "^root L"; then
                    log_success "Root password login disabled"
                    log_info "Root account is now locked - only '${sudo_user}' can log in with password"
                    log_info "Root can still be accessed via: sudo -i"
                else
                    log_warning "Failed to lock root account"
                fi
            fi
        else
            log_info "Skipping root account lockdown"
        fi
    fi
    
    echo
    print_separator
    log_success "Configuration complete!"
    print_separator
    echo
    log_info "Your homelab is now configured for:"
    log_info "  ✓ Local LAN access to all services"
    log_info "  ✓ Remote access via Tailscale VPN"
    log_info "  ✓ SSH access on LAN and via Tailscale"
    log_info "  ✓ Firewall protecting against unauthorized access"
    local current_ip=$(get_current_ip)
    local display_ip="${current_ip:-<your-pi-ip>}"
    echo
    log_info "Access your services:"
    log_info "  - Pi-hole:        http://${display_ip}/admin"
    log_info "  - Home Assistant: http://${display_ip}:8123"
    echo
    log_info "To create backups, run: sudo $0 backup"
    echo
}

################################################################################
# MODE 3: BACKUP
# Create comprehensive backup of all configurations
################################################################################

mode_backup() {
    print_header "BACKUP"
    
    # Check prerequisites for this mode
    if ! check_all_prerequisites "backup"; then
        err "Prerequisites check failed. Cannot continue."
    fi
    
    # Ensure backup directory exists
    if [[ ! -d "${BACKUP_DIR}" ]]; then
        create_directory_safe "${BACKUP_DIR}" "Backup directory" || err "Failed to create backup directory"
    fi
    
    local timestamp=$(timestamp_filename)
    local backup_file="${BACKUP_DIR}/homelab-backup-${timestamp}.tar.gz"
    
    log_info "Creating backup: homelab-backup-${timestamp}.tar.gz"
    echo
    
    # Stop containers to ensure consistent backup
    log_step "Stopping containers for consistent backup"
    local containers_stopped=false
    if verify_docker_running 2>/dev/null; then
        manage_docker_services stop
        log_success "Containers stopped"
        containers_stopped=true
    else
        log_info "Docker not running, skipping container stop"
    fi
    
    log_step "Creating backup archive"
    log_info "Backing up essential configurations..."
    
    # Create single compressed archive with all essential configs
    # Use tar's ability to continue on missing files (--ignore-failed-read)
    local tar_files=()
    local tar_base_dirs=()
    local source_count=0
    
    # Pi-hole configuration (DNS settings, blocklists, custom records)
    if [[ -d "${HOMELAB_DIR}/pihole/etc-pihole" ]]; then
        tar_files+=("pihole/etc-pihole")
        tar_base_dirs+=("${HOMELAB_DIR}")
        ((source_count++))
        log_info "  - Pi-hole configuration"
    fi
    
    # Home Assistant configuration (automations, integrations, dashboards)
    if [[ -d "${HOMELAB_DIR}/homeassistant" ]]; then
        tar_files+=("homeassistant")
        tar_base_dirs+=("${HOMELAB_DIR}")
        ((source_count++))
        log_info "  - Home Assistant configuration"
    fi
    
    # Docker Compose files and .env (includes tokens)
    if [[ -d "${COMPOSE_DIR}" ]]; then
        tar_files+=("compose")
        tar_base_dirs+=("${HOMELAB_DIR}")
        ((source_count++))
        log_info "  - Docker Compose files and tokens"
    fi
    
    # Firewall rules
    if [[ -f /etc/nftables.conf ]]; then
        tar_files+=("etc/nftables.conf")
        tar_base_dirs+=("/")
        ((source_count++))
        log_info "  - Firewall rules"
    fi
    
    # Tailscale VPN identity
    if [[ -f /var/lib/tailscale/tailscaled.state ]]; then
        tar_files+=("var/lib/tailscale/tailscaled.state")
        tar_base_dirs+=("/")
        ((source_count++))
        log_info "  - Tailscale identity"
    fi
    
    # Create the backup archive
    echo
    
    # Build tar command with proper directory changes
    local tar_cmd=("tar" "czf" "$backup_file")
    
    # Add homelab directory contents
    if [[ -d "${HOMELAB_DIR}/pihole/etc-pihole" ]] || [[ -d "${HOMELAB_DIR}/homeassistant" ]] || [[ -d "${COMPOSE_DIR}" ]]; then
        tar_cmd+=("-C" "${HOMELAB_DIR}")
        [[ -d "${HOMELAB_DIR}/pihole/etc-pihole" ]] && tar_cmd+=("pihole/etc-pihole")
        [[ -d "${HOMELAB_DIR}/homeassistant" ]] && tar_cmd+=("homeassistant")
        [[ -d "${COMPOSE_DIR}" ]] && tar_cmd+=("compose")
    fi
    
    # Add system files from root
    if [[ -f /etc/nftables.conf ]] || [[ -f /var/lib/tailscale/tailscaled.state ]]; then
        tar_cmd+=("-C" "/")
        [[ -f /etc/nftables.conf ]] && tar_cmd+=("etc/nftables.conf")
        [[ -f /var/lib/tailscale/tailscaled.state ]] && tar_cmd+=("var/lib/tailscale/tailscaled.state")
    fi

    if [[ $source_count -eq 0 ]]; then
        log_warning "No backup sources found. Nothing to archive."
        log_info "Verify ${HOMELAB_DIR} contents and system config paths before retrying."
        if [[ "$containers_stopped" == "true" ]] && verify_docker_running 2>/dev/null; then
            manage_docker_services start
            log_success "Containers restarted"
        fi
        return 0
    fi
    
    # Execute tar command
    if "${tar_cmd[@]}" 2>/dev/null; then
        # Verify archive integrity by testing extraction
        if tar tzf "$backup_file" >/dev/null 2>&1; then
            local backup_size=$(du -h "$backup_file" | cut -f1)
            log_success "Backup created successfully: ${backup_size}"
        else
            log_error "Backup verification failed - archive is corrupted"
            rm -f "$backup_file"
            err "Backup creation failed - archive integrity check failed"
        fi
    else
        log_error "Failed to create backup archive"
        rm -f "$backup_file"
        err "Backup creation failed - check disk space and permissions"
    fi
    
    # Restart containers
    log_step "Restarting containers"
    if verify_docker_running 2>/dev/null; then
        manage_docker_services start
        log_success "Containers restarted"
    fi
    
    # Simple backup rotation: keep last 3 backups
    log_step "Managing backup rotation"
    local all_backups=($(ls -1t "${BACKUP_DIR}"/homelab-backup-*.tar.gz 2>/dev/null || true))
    local backup_count=${#all_backups[@]}
    
    if [[ $backup_count -gt 3 ]]; then
        log_info "Keeping 3 most recent backups, removing $((backup_count - 3)) old backup(s)"
        # Remove backups beyond the 3rd (index 3+)
        for ((i=3; i<backup_count; i++)); do
            local old_backup="${all_backups[$i]}"
            log_info "  Removing: $(basename "$old_backup")"
            rm -f "$old_backup" 2>/dev/null || log_warning "Failed to remove old backup (not critical)"
        done
    fi
    
    echo
    print_separator
    log_success "Backup complete!"
    print_separator
    echo
    log_info "Backup file: $(basename "$backup_file")"
    log_info "Location: ${BACKUP_DIR}"
    log_info "Size: $(du -h "$backup_file" | cut -f1)"
    echo
    
    # Show all current backups
    local current_backups=($(ls -1t "${BACKUP_DIR}"/homelab-backup-*.tar.gz 2>/dev/null || true))
    if [[ ${#current_backups[@]} -gt 0 ]]; then
        log_info "Available backups (${#current_backups[@]}/3):"
        for b in "${current_backups[@]}"; do
            local b_size=$(du -h "$b" | cut -f1)
            log_info "  - $(basename "$b") (${b_size})"
        done
    fi
    echo
    log_info "To restore this backup, run: sudo $0 restore"
    echo
}

################################################################################
# MODE 4: RESTORE
# Restore from backup
################################################################################

mode_restore() {
    print_header "RESTORE"
    
    # Check prerequisites for this mode
    if ! check_all_prerequisites "restore"; then
        err "Prerequisites check failed. Cannot continue."
    fi
    
    # List available backups
    if [[ ! -d "$BACKUP_DIR" ]]; then
        err "Backup directory not found: ${BACKUP_DIR}"
    fi
    
    local backups=($(ls -1t "${BACKUP_DIR}"/homelab-backup-*.tar.gz 2>/dev/null || true))
    
    if [[ ${#backups[@]} -eq 0 ]]; then
        err "No backups found in ${BACKUP_DIR}"
    fi
    
    log_info "Available backups:"
    echo
    
    local i=1
    for backup in "${backups[@]}"; do
        local size=$(du -h "$backup" | cut -f1)
        local filename=$(basename "$backup")
        printf "%2d) %-40s  %s\n" "$i" "$filename" "$size"
        i=$((i+1))
    done
    
    echo
    read -p "Select backup to restore (1-${#backups[@]}): " selection
    
    if [[ ! "$selection" =~ ^[0-9]+$ ]] || [[ $selection -lt 1 ]] || [[ $selection -gt ${#backups[@]} ]]; then
        err "Invalid selection"
    fi
    
    local backup_file="${backups[$((selection-1))]}"
    
    log_info "Selected: $(basename "$backup_file")"
    echo
    
    # Quick integrity check using tar's built-in verification
    log_step "Verifying backup integrity"
    if tar tzf "$backup_file" >/dev/null 2>&1; then
        log_success "Backup integrity verified"
    else
        log_error "Backup archive is corrupted or invalid"
        if ! prompt_yn "Attempt restore anyway? (likely to fail)"; then
            log_info "Restore cancelled"
            exit 0
        fi
    fi
    
    echo
    log_warning "WARNING: This will overwrite current configurations"
    echo
    if ! prompt_yn "Proceed with restore?"; then
        log_info "Restore cancelled"
        exit 0
    fi
    
    log_step "Stopping services"
    
    if verify_docker_running 2>/dev/null; then
        manage_docker_services stop
        log_success "Containers stopped"
    else
        log_info "Docker not running, skipping container stop"
    fi
    
    if tailscale status >/dev/null 2>&1; then
        log_info "Stopping Tailscale..."
        tailscale down
    fi
    
    log_step "Clearing existing configurations"
    
    # Wait for file locks to be released after container stop
    sleep $CLEANUP_DELAY
    
    # Remove existing directories (with retry logic for file locks)
    local lock_attempts=$((FILE_LOCK_RETRY_MAX_TIME / FILE_LOCK_RETRY_DELAY))
    
    if [[ -d "${HOMELAB_DIR}/pihole/etc-pihole" ]]; then
        log_info "Removing Pi-hole configuration..."
        retry_until_connected --attempts $lock_attempts --delay $FILE_LOCK_RETRY_DELAY rm -rf "${HOMELAB_DIR}/pihole/etc-pihole" || \
            log_warning "Failed to remove Pi-hole directory (may still be locked)"
    fi
    
    if [[ -d "${HOMELAB_DIR}/homeassistant" ]]; then
        log_info "Removing Home Assistant configuration..."
        retry_until_connected --attempts $lock_attempts --delay $FILE_LOCK_RETRY_DELAY rm -rf "${HOMELAB_DIR}/homeassistant" || \
            log_warning "Failed to remove Home Assistant directory (may still be locked)"
    fi
    
    if [[ -d "${COMPOSE_DIR}" ]]; then
        log_info "Removing Docker Compose files..."
        rm -rf "${COMPOSE_DIR}" 2>/dev/null || log_warning "Failed to remove compose directory"
    fi
    
    log_step "Extracting backup archive"
    
    # Extract entire backup to root - tar will recreate directory structure
    if ! tar xzf "$backup_file" -C / 2>/dev/null; then
        log_error "Failed to extract backup archive"
        err "Restore failed - check archive integrity and disk space"
    fi
    
    log_success "Backup extracted successfully"
    
    # Verify critical directories were restored
    local restore_ok=true
    
    if [[ ! -d "${HOMELAB_DIR}/pihole/etc-pihole" ]] && [[ ! -d "${HOMELAB_DIR}/homeassistant" ]]; then
        log_warning "Neither Pi-hole nor Home Assistant configs found in backup"
        restore_ok=false
    fi
    
    if [[ ! -d "${COMPOSE_DIR}" ]]; then
        log_warning "Docker Compose files not found in backup"
        restore_ok=false
    fi
    
    if [[ "$restore_ok" == "false" ]]; then
        log_warning "Backup appears incomplete or from different system"
        if ! prompt_yn "Continue with service startup anyway?"; then
            log_info "Restore cancelled"
            exit 0
        fi
    fi
    
    # Restore firewall rules if present
    if [[ -f /etc/nftables.conf ]]; then
        log_info "Firewall rules restored"
        if systemctl is-enabled nftables >/dev/null 2>&1; then
            nft -f /etc/nftables.conf 2>/dev/null || log_warning "Failed to reload firewall rules"
        fi
    fi
    
    # Set Tailscale permissions if state was restored
    if [[ -f /var/lib/tailscale/tailscaled.state ]]; then
        chown root:root /var/lib/tailscale/tailscaled.state 2>/dev/null
        chmod 600 /var/lib/tailscale/tailscaled.state 2>/dev/null
        log_info "Tailscale state restored"
    fi
    
    log_step "Starting services"
    
    manage_docker_services start
    
    if [[ -f /var/lib/tailscale/tailscaled.state ]]; then
        log_info "Starting Tailscale..."
        tailscale up --ssh --accept-dns=false
    fi
    
    log_success "Services started"
    
    log_step "Verifying services"
    sleep $SERVICE_STATUS_CHECK_DELAY
    
    report_container_status
    local status_ok=$?
    
    print_separator
    if [[ $status_ok -eq 0 ]]; then
        log_success "Restore complete! All services are running."
    else
        log_warning "Restore complete with warnings. Check service logs."
    fi
    print_separator
    echo
}

################################################################################
# MAIN SCRIPT ENTRY POINT
################################################################################

show_usage() {
        local lines=(
                "Raspberry Pi Homelab Docker Stack Setup v${SCRIPT_VERSION}"
                ""
                "Usage: $0 <mode> [options]"
                ""
                "Modes:"
                "  install-docker  - Install Docker stack (Pi-hole, Home Assistant, Homeway, Tailscale)"
                "  configure       - Post-setup configuration (collect API tokens and keys)"
                "  backup          - Backup all Docker configurations and volumes"
                "  restore         - Restore Docker configurations from backup"
                ""
                "Examples:"
                "  sudo $0 install-docker"
                "  sudo $0 configure"
                "  sudo $0 backup"
                "  sudo $0 restore"
                ""
                "For more information, visit: https://github.com/yourusername/homelab-setup"
                ""
        )

        printf '%s\n' "${lines[@]}"
}

main() {
    # Check for root except for help
    if [[ "$1" != "-h" ]] && [[ "$1" != "--help" ]] && [[ "$1" != "help" ]]; then
        check_root
    fi
    
    # Parse command
    local mode="${1:-}"
    
    case "$mode" in
        install-docker)
            mode_install_docker
            ;;
        configure)
            mode_configure
            ;;
        backup)
            mode_backup
            ;;
        restore)
            mode_restore
            ;;
        help|-h|--help)
            show_usage
            ;;
        "")
            log_error "No mode specified"
            echo
            show_usage
            exit 1
            ;;
        *)
            log_error "Unknown mode: ${mode}"
            echo
            show_usage
            exit 1
            ;;
    esac
}

# Run main function
main "$@"
