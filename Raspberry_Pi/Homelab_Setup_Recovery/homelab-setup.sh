#!/bin/bash
#
################################################################################
# Raspberry Pi Headless Setup Helpers
#
# Shared helper functions for first-time headless Raspberry Pi setup.
################################################################################
################################################################################
# CONSTANTS AND CONFIGURATION
################################################################################

readonly SCRIPT_VERSION="1.0.0"

################################################################################
# Connection & Network Retry Configuration
################################################################################
# Fast connection establishment (for DNS issues)
readonly CONNECTION_RETRY_MAX_ATTEMPTS=15
readonly CONNECTION_RETRY_DELAY=2

# apt repository/package retries
readonly APT_RETRY_ATTEMPTS=5
readonly APT_RETRY_DELAY=4

################################################################################
# Resource Requirements
################################################################################
readonly MIN_DISK_SPACE_MB=100
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

# Unified error handler
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
    
    exit 1
}

# Check if script is run as root
check_root() {
    if [[ $EUID -ne 0 ]]; then
        log_error "This script must be run as root (use sudo)"
        err "Root privileges required"
    fi
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
    log_error "This script requires internet access to download packages."
    return 1
}

################################################################################
# MAIN SCRIPT ENTRY POINT
################################################################################

main() {
    log_info "homelab-setup.sh has no setup modes yet."
    log_info "Docker stack modes were removed. Shared helpers remain for headless Pi setup."
    log_info "For boot, USB, and first-boot configuration, use switch-root.sh."
}

main "$@"