# Homelab Tiny/Mini/Micro CPU & Platform Reference

---

## CPU TIERS

### Tier 3 (Recommended baseline / sweet spot)

| CPU | Core/Thread Notes | iGPU | Approx PassMark Range |
| :--- | :--- | :--- | :--- |
| **i5-9500T** | No HT (6C/6T) | Intel UHD 630 | ~7,000 |
| **i7-9700T / i7-9700** | No HT (8C/8T) | Intel UHD 630 | ~11,000 - 13,000 |
| **i5-10500T / i5-10500** | (6C/12T) | Intel UHD 630 | ~10,000 - 13,400 |
| **i5-10600T / i5-10600** | (6C/12T) | Intel UHD 630 | ~11,000 - 13,500 |
| **i7-8700T / i7-8700** | (6C/12T) | Intel UHD 630 | ~11,000 - 13,000 |
| **i7-10700T / i7-10700** | (8C/16T) | Intel UHD 630 | ~13,500 - 17,000 |
| **Ryzen PRO 4650GE / 4650G** | (6C/12T) | Vega 7 | ~15,000 - 16,000 |
| **i5-11400T / i5-11400** | (6C/12T) | Intel UHD 730 | ~12,000 - 17,000 |
| **i5-11500T / i5-11500** | (6C/12T) | Intel UHD 750 | ~13,500 - 17,500 |
| **Ryzen PRO 4650GE / 4650G** | (6C/12T) | Radeon Vega 7 | ~15,000 - 16,000 |

---

### Tier 4 (High-performance compact)

| CPU | Core/Thread Notes | iGPU | Approx PassMark Range |
| :--- | :--- | :--- | :--- |
| **Ryzen 5600GE / 5600G** | (6C/12T) | Radeon Vega 7 | ~18,000 - 19,800 |
| **Ryzen 5800U** | (8C/16T) | Radeon Vega 8 | ~19,000 |
| **Ryzen PRO 4750GE / 4750G / 5700GE** | (8C/16T) | Radeon Vega 8 | ~19,500 - 21,000 |
| **i5-11600T / i5-11600** | (6C/12T) | Intel UHD 750 | ~14,000 - 18,000 |
| **i5-12400T / i5-12400** | (6 P-Cores / 12T) | Intel UHD 730 | ~15,500 - 19,400 |
| **Ryzen PRO 5650GE / 5650G** | (6C/12T) | Vega 7 | ~18,000 - 19,600 |
| **i5-12500T / i5-12500** | (6 P-Cores / 12T) | Intel UHD 770 | ~16,300 - 20,000 |
| **i5-12600T / i5-12600** | (6 P-Cores / 12T) | Intel UHD 770 | ~17,400 - 21,500 |
| **Ryzen 5700G / 5800H / 5900HX** | (8C/16T) | Radeon Vega 8 | ~21,500 - 22,000 |
| **Ryzen PRO 5750GE / 5750G** | (8C/16T) | Vega 8 | ~22,000 - 22,500 |
| **i5-13400T / i5-13400** | (10C/16T) | Intel UHD 730 | ~18,000 - 22,000 |
| **Ryzen 7640HS** | (6C/12T) | Radeon 760M | ~23,000 |
| **Ryzen 6800H / 6850H / 7735HS** | (8C/16T) | Radeon 680M | ~24,000 |
| **i5-12600K** | (10C/16T) | Intel UHD 770 | ~24,500 |
| **i5-13500T / i5-13500** | (14C/20T) | Intel UHD 770 | ~22,500 - 26,500 |
| **Ryzen 7840HS / Ryzen 7940HS** | (8C/16T) | Radeon 780M | ~29,500 |
| **i7-12700T / i7-12700** | (12C/20T) | Intel UHD 770 | ~23,500 - 31,000 |
| **i5-13600T / i5-13600** | (14C/20T) | Intel UHD 770 | ~27,100 - 31,500 |
| **Ryzen 7 7700** | (8C/16T) | AMD Radeon | ~34,300 |
| **i7-13700T / i7-13700** | (16C/24T) | Intel UHD 770 | ~27,500 - 38,000 |

---

## PLATFORM FAMILIES

### 1L Tiny / Mini / Micro (Primary homelab nodes)

**Dell OptiPlex Micro**
- 7050 / 7060 → 8th gen
- 7070 / 7080 → 9th–10th gen
- 7090 / 7000 → 10th–11th gen

**HP EliteDesk / ProDesk Mini**
- 800 G5 → 8th/9th gen
- 800 G6 → 10th gen
- 800 G7 → 11th gen

**Lenovo ThinkCentre Tiny**
- M920q / M920x → 8th gen
- M720q → 8th gen
- M70q Gen 2 → 10th gen
- M75q Gen 2 → Ryzen 4000
- M75q Gen 3 → Ryzen 5000

Notes:
- Best overall value / availability
- Limited expansion (except Lenovo PCIe riser models)

---

## Tiny Workstations (Same size, much more capability)

**Lenovo ThinkStation Tiny**
- P330 Tiny → 8th/9th gen + Quadro
- P340 Tiny → 10th gen
- P360 Tiny → 12th gen

**HP Z2 Mini**
- G4 → 8th/9th gen
- G5 → 10th gen / Xeon
- G9 → 12th gen

**Dell Precision Compact / Micro**
- Precision 3430 Compact
- Precision 3240 Compact

Notes:
- GPU-capable
- Higher RAM ceilings
- Best for heavy workloads, AI, cameras

---

## Small Form Factor (SFF)

- OptiPlex 7060 / 7070 / 7080 SFF
- EliteDesk 800 G5 / G6 SFF
- ThinkCentre M720s / M920s

Notes:
- PCIe expansion (NICs, storage, GPU)
- Supports full Tier 3 + some Tier 4 CPUs

---

## Modern Mini-PC (Non-enterprise)

**Minisforum**
- UM590 / UM690 → Ryzen 5000–6000
- UM790 Pro → Ryzen 7040
- MS-01 → high-end Intel + PCIe + 10GbE

**Beelink**
- SER6 / SER7 → Ryzen 6000–7000
- SER8 → Ryzen 8845HS

**ASUS NUC**
- NUC 12 → 12th gen Intel
- NUC 13 → 13th gen Intel

Notes:
- Highest performance-per-watt
- Less enterprise consistency than Dell/HP/Lenovo

---

## QUICK SELECTION GUIDE

FLOOR:
- i5-9500T
- i7-8700

TARGET:
- i5-10500T
- i7-9700
- Ryzen 5 5600G

HIGH:
- i5-12500
- i7-12700
- Ryzen 7 5700G+

---

## ROLE MAPPING

Primary Node:
- OptiPlex / EliteDesk / ThinkCentre (Tier 3)

Upgrade Node:
- Same class or newer

Specialized Node:
- Z2 Mini / ThinkStation Tiny

Expansion Node:
- SFF systems

---

## NOTES

- Prefer NVMe support
- Target 32GB RAM minimum
- Avoid pre-8th gen Intel unless extremely cheap
- AMD Tiny systems (M75q Gen 2+) are best enterprise Ryzen option

---

(End of reference)