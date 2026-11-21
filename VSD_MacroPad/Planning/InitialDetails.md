# Project Summary: VSDinside Stream Dock Automatic Text Scaling

## Hardware and Software

- Device: VSDinside Stream Dock / Macro Keypad (exact model not yet confirmed).
- Configuration software: VSD Craft.
- Device is connected through a USB dock.
- Device contains LCD buttons whose labels are configured through VSD Craft and stored in local configuration files.

## Original Issue

The device remains powered when disconnected from the PC because the dock continues to provide USB power.

Instead of turning off, the Stream Dock enters an idle/demo mode and continues displaying content on its button LCDs.

Investigation was started into:
1. Device sleep behavior.
2. Display scaling behavior.
3. Configuration automation opportunities.

---

## Windows DPI Scaling Investigation

A Windows application associated with the device appeared correctly sized on a 1080p monitor but did not properly scale on a 4K monitor.

Windows Compatibility settings were modified:

**Properties → Compatibility → Change High DPI Settings → Override High DPI Scaling Behavior**

Testing was performed using:
- Application
- System
- System (Enhanced)

System (Enhanced) was selected/tested because it generally produces sharper results on older Win32 applications than System scaling.

### Side Effect

After DPI scaling changes were applied, text rendered on the Stream Dock LCD buttons became smaller than intended for some buttons.

It is not yet clear whether:

- VSD Craft stores font size directly in configuration data.
- The software recalculates button layouts based on Windows DPI.
- Rendered button images are regenerated when DPI changes.
- The actual LCD assets are bitmap-based rather than text-based.

---

## Desired End State

The user wants button labels to be automatically readable regardless of text length.

### Primary Goal

Automatically scale text size so that labels fit inside square keypad displays.

Examples:

- Short labels should use larger font sizes.
- Long labels should shrink automatically.
- No text clipping.
- No manual per-button font adjustment.

---

## Potential Future Enhancement

Instead of storing text directly:

1. Generate a square image for each button.
2. Render text into the image dynamically.
3. Apply automatic font-size reduction until text fits within predefined margins.
4. Export image to the format required by VSD Craft.
5. Assign generated image to the button automatically.

### Advantages

- Consistent appearance.
- Independent of Windows DPI behavior.
- Full control over typography, alignment, margins, and wrapping.
- Easier long-term automation.

---

## Investigation Required

Determine where VSD Craft stores button definitions.

Need to identify:

- Configuration file locations.
- File formats (JSON, XML, SQLite, binary, proprietary).
- Asset storage locations.
- Font-size fields.
- Text label fields.
- Image attachment fields.

### Useful Approach

1. Create several test buttons.
2. Export or save a profile.
3. Search profile directories for modified files.
4. Compare before/after versions.
5. Determine the underlying schema.

---

## Desired Automation Options

### Option A: Batch Text Scaling

Create a utility that:

- Reads VSD Craft profile/configuration data.
- Calculates optimal font size for every button label.
- Updates profile data automatically.
- Reimports or overwrites the profile.

### Option B: Automatic Image Generation (Preferred)

Create a utility that:

- Reads button labels from a profile.
- Generates square PNG images.
- Auto-scales text to fit.
- Supports:
  - Multiline wrapping
  - Configurable margins
  - Centered alignment
  - Optional icons
  - Dark/light themes
- Updates button image references in the profile automatically.

---

## Technical Questions

1. Where are VSD Craft profile files stored?
2. What file format is used?
3. Are button labels stored as text or rendered images?
4. Is there a global font-size parameter?
5. Can multiple button definitions be edited simultaneously outside the GUI?
6. Is there an import/export API or plugin SDK?
7. Can profile files be modified while VSD Craft is closed and then reloaded?

---

## Proposed Development Path

### Phase 1: Reverse Engineering

- Locate VSD Craft profile files.
- Determine profile structure.
- Identify where text, font settings, and image references are stored.

### Phase 2: Batch Font-Size Editor

- Develop a utility that modifies font sizes programmatically.
- Add automatic fit calculations.

### Phase 3: Text-to-Image Pipeline

- Generate button graphics from labels.
- Create files suitable for direct import.

### Phase 4: Intelligent Sizing Algorithm

Implement an algorithm that:

1. Measures rendered text dimensions.
2. Uses binary search (or equivalent) to determine the maximum font size that fits.
3. Wraps text when necessary.
4. Applies configurable margins.
5. Exports the final image.

### Phase 5: Themes and Templates

- Icon overlays.
- Color themes.
- Template-driven button generation.
- Batch regeneration of entire profiles.

---

## Preferred Long-Term Architecture

The preferred long-term solution is image generation rather than direct font-size manipulation.

Reasons:

- Eliminates dependency on VSD Craft's text-rendering behavior.
- Eliminates DPI-scaling side effects.
- Ensures consistent appearance across software versions.
- Makes advanced layouts possible.
- Allows future templating and bulk profile generation.

---

## Recommended First Technical Task

Determine whether VSD Craft stores profiles as:

- JSON
- XML
- SQLite
- Binary blobs
- Other proprietary formats

If the profiles are JSON or XML, batch-editing and automation can likely be implemented very quickly.

If labels are already stored as image references, an automated image-generation workflow may be easier and more robust than manipulating text properties directly.