EDITING NOTE: Before creating or editing ASCII graphics in this file, check the ASCII DRAWING STANDARDS section at the bottom and follow those symbols/spacing rules exactly.

HOME LAN

NETWORK RECOVERY MAP (ASCII)
Last updated: 2026-04-28

=========================================================================
TOPOLOGY OVERVIEW
=========================================================================


                           [ Router / LAN GW ]
                               192.168.1.1
                               /         \
                              /           \
                             /             \
                   Wi-Fi (DHCP)           Ethernet LAN 192.168.1.0/24
                          │                             │
                          │                             │
               ┌─────────────────────┐       ┌─────────────────────┐
               │ Laptop Wi-Fi Adapter│       │ Desktop Network     │
               │ DHCP (company VPN)  │       │ Bridge              │
               └──────────┬──────────┘       │ IP: 192.168.1.95    │
                          │                  │ GW: 192.168.1.1     │
                          │                  └──────────┬──────────┘
                          │                             │
                          │                             │ bridge member
                          │                    ┌────────┴────────┐
                          │                    │ Desktop Eth NIC │
                          │                    └────────┬────────┘
                          │                             │
                          │                             │ USB bridge member
                          │                  ┌──────────┴─────────┐
                          │                  │    USB Ethernet    │
                          │                  │  Adapter + Switch  │
                          │                  └──────────┬─────────┘
                          │                             │
                          │                             │ cable
                          │                    ┌────────┴───────┐
                          │                    │ Laptop Eth NIC │
                          │                    │ 192.168.1.90   │
                          │                    │ GW: blank      │
                          │                    │ DNS: blank     │
                          │                    └────────────────┘
                          │
                          │ host internet path + VPN tunnel path
                          │
        ┌────────────────────────────────────────────────────────────┐
        │ VMware VM on Laptop                                        │
        │                                                            │
        │ NIC #1: VMnet0 (Auto-Bridged -> Laptop Ethernet Adapter)   │
        │   VM IP: 192.168.1.91 (DHCP from LAN)                      │
        │   Use: Local LAN reachability (192.168.1.x side)           │
        │                                                            │
        │ NIC #2: VMnet8 (NAT -> Laptop Wi-Fi Adapter)               │
        │   VMware NAT path to internet via host Wi-Fi               │
        │   Use: Internet access even if LAN side has no gateway     │
        └────────────────────────────────────────────────────────────┘


=========================================================================
ADAPTER QUICK REFERENCE
=========================================================================

1) Laptop Ethernet Adapter
- IP: 192.168.1.90 (static)
- Subnet: likely 255.255.255.0
- Default Gateway: blank
- DNS: blank
- Purpose: local wired segment + VMnet0 bridge target

2) Laptop Wi-Fi Adapter
- IP: DHCP (from router)
- Default Gateway: router-provided
- DNS: router/company-provided
- Special: company VPN tunnel carries traffic
- Purpose: internet + VPN path; host path used by VMnet8 NAT

3) VMware VM Adapter (VMnet0)
- Type: auto-bridged
- Bridge target: Laptop Ethernet Adapter
- VM IP: 192.168.1.91 (DHCP)
- Purpose: appears directly on 192.168.1.x LAN

4) VMware VM Adapter (VMnet8)
- Type: NAT
- Upstream: Laptop Wi-Fi Adapter
- Purpose: VM internet via host networking stack

5) Desktop Network Bridge
- Bridge includes: Desktop Ethernet + USB Ethernet connection
- Bridge IP: 192.168.1.95
- Default Gateway: 192.168.1.1
- DNS: blank
- Purpose: LAN + internet access on desktop through bridge


=========================================================================
RECOVERY CHECKLIST (FROM SCRATCH)
=========================================================================

1) Physical layer
- Connect laptop Ethernet port to USB Ethernet switch.
- Connect switch to desktop via USB Ethernet path.

2) Laptop adapter config
- Set laptop Ethernet static IP to 192.168.1.90/24.
- Leave laptop Ethernet gateway and DNS blank.
- Confirm laptop Wi-Fi is connected and gets DHCP.
- Confirm company VPN comes up on Wi-Fi path.

3) VMware network config
- VM Adapter 1 = VMnet0 (Auto-Bridged).
- Confirm VMnet0 bridges to laptop Ethernet adapter.
- VM Adapter 2 = VMnet8 (NAT).
- Boot VM and verify it gets 192.168.1.91 on bridged NIC.

4) Desktop bridge config
- Ensure desktop network bridge exists and includes both ports.
- Set bridge IP to 192.168.1.95/24.
- Set bridge gateway to 192.168.1.1.
- Keep bridge DNS blank (as currently designed).

5) Validation tests
- From laptop: ping 192.168.1.95 and 192.168.1.91.
- From desktop: ping 192.168.1.1, then test internet.
- From VM:
	- ping 192.168.1.90 and 192.168.1.95 via bridged NIC.
	- test internet via NAT NIC.


=========================================================================
TROUBLESHOOTING SNAPSHOTS
=========================================================================

- VM cannot reach LAN:
	- Verify VMnet0 is bridged to laptop Ethernet, not Wi-Fi.
	- Check VM has 192.168.1.x address and correct subnet.

- VM has LAN but no internet:
	- Confirm VMnet8 adapter is connected and set to NAT.
	- Confirm laptop Wi-Fi and VPN are up.

- Desktop has LAN but no internet:
	- Verify desktop bridge gateway is 192.168.1.1.
	- Confirm router path from bridged segment is available.

- Name resolution fails:
	- DNS is intentionally blank on wired/bridge interfaces.
	- Test by IP first; add DNS only if needed for hostnames.


=========================================================================
ASCII DRAWING STANDARDS
=========================================================================

- Approved line/box symbols for graphics: │ ─ ┤ ┴ ├ ┬ ┼ ┌ ┐ └ ┘ / \
- On any line that contains ASCII/box graphics, use non-breaking spaces " " for indentation/alignment instead of regular spaces.
- Keep all diagram edits visually aligned in a monospaced editor view.

