
#################################################################################
    Windows
#################################################################################
New-Item -ItemType Directory -Force "$env:USERPROFILE\.ssh"
ssh-keygen -t ed25519 -a 100 -C "$env:USERNAME@$env:COMPUTERNAME" -f "$env:USERPROFILE\.ssh\id_ed25519_ubuntu_homelab"
Get-Content "$env:USERPROFILE\.ssh\id_ed25519.pub" | Set-Clipboard


#################################################################################
    Ubuntu SSH Key Only
#################################################################################
mkdir -p ~/.ssh
chmod 700 ~/.ssh
touch ~/.ssh/authorized_keys
chmod 600 ~/.ssh/authorized_keys

cat >> ~/.ssh/authorized_keys <<'EOF'
PASTE_YOUR_PUBLIC_KEY_HERE
EOF

chown -R "$USER:$USER" ~/.ssh
chmod 700 ~/.ssh
chmod 600 ~/.ssh/authorized_keys

sudo cp /etc/ssh/sshd_config /etc/ssh/sshd_config.backup.$(date +%Y%m%d-%H%M%S)

sudo tee /etc/ssh/sshd_config.d/00-key-only-auth.conf > /dev/null <<'EOF'
# Key-only SSH authentication
PubkeyAuthentication yes
PasswordAuthentication no
KbdInteractiveAuthentication no
ChallengeResponseAuthentication no
AuthenticationMethods publickey
PermitRootLogin no
EOF

sudo sshd -t
sudo systemctl reload ssh


#################################################################################
    Ubuntu First Boot
#################################################################################
sudo bash -c '
set -euo pipefail

echo "Updating package lists"
apt update

echo "Upgrading installed packages"
DEBIAN_FRONTEND=noninteractive apt full-upgrade -y

echo "Installing baseline server tools"
DEBIAN_FRONTEND=noninteractive apt install -y \
    ca-certificates \
    curl \
    wget \
    gnupg \
    lsb-release \
    software-properties-common \
    apt-transport-https \
    unzip \
    htop \
    iotop \
    iftop \
    ncdu \
    nano \
    vim \
    tmux \
    git \
    net-tools \
    dnsutils \
    traceroute \
    iputils-ping \
    openssh-server \
    unattended-upgrades

echo "Installing Hyper-V optimized kernel and tools"
DEBIAN_FRONTEND=noninteractive apt install -y \
    linux-virtual \
    linux-tools-virtual \
    linux-cloud-tools-virtual

echo "Trying to install kernel-specific Hyper-V tools"
apt install -y "linux-tools-$(uname -r)" "linux-cloud-tools-$(uname -r)" || true

echo "Adding Hyper-V modules to initramfs module list"
for mod in hv_vmbus hv_storvsc hv_netvsc hv_utils hv_balloon; do
    if ! grep -Fxq "$mod" /etc/initramfs-tools/modules; then
        echo "$mod" >> /etc/initramfs-tools/modules
    fi
done

echo "Updating initramfs"
update-initramfs -u

echo "Enabling SSH"
systemctl enable --now ssh

echo "Enabling Hyper-V daemons if present"
systemctl enable --now hv-kvp-daemon.service 2>/dev/null || true
systemctl enable --now hv-vss-daemon.service 2>/dev/null || true
systemctl enable --now hv-fcopy-daemon.service 2>/dev/null || true

echo "Cleaning packages"
apt autoremove -y
apt autoclean

echo "First boot setup complete. Reboot recommended."
'


#################################################################################
    Ubuntu Swap
#################################################################################
    Google it



#################################################################################
    Ubuntu Static IP
#################################################################################
IFACE="$(ip route | awk '/default/ {print $5; exit}')"
echo "Detected interface: $IFACE"

sudo tee /etc/netplan/99-static-ip.yaml > /dev/null <<EOF
network:
  version: 2
  renderer: networkd
  ethernets:
    $IFACE:
      dhcp4: false
      addresses:
        - 192.168.1.210/24
      routes:
        - to: default
          via: 192.168.1.1
      nameservers:
        addresses:
          - 192.168.1.1
          - 1.1.1.1
          - 8.8.8.8
EOF
sudo chmod 600 /etc/netplan/99-static-ip.yaml
sudo netplan try
    