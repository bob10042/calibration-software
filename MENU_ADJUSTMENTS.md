# Menu Adjustments Required

## Menu Structure Changes
1. Move Meters Menu to Replace Spare3
- Current location: Meters menu is defined in panelSideMenu (lines 240-328)
- Target location: Replace btnSpare3 in panelCalibrationMenu (around line 400)
- Keep all existing meter buttons:
  - APS M2003
  - Fluke 8508A
  - Keysight 34471A
  - Vitrek 920A

## Visual Style Adjustments

### Font Size Fixes
1. Increase dropdown font sizes to match other menu items
- Current font: Microsoft Sans Serif, 9F
- Update font size to match other menu items for consistency
- Affected components:
  - btnMeters
  - btnApsM2003
  - btnFluke8508A
  - btnKeysight34471A
  - btnVitrek920A

### Color Consistency
1. Menu Button Colors
- Main menu buttons should use: Color.FromArgb(((int)(((byte)(1)))), ((int)(((byte)(147)))), ((int)(((byte)(207)))))
- Submenu buttons should use: Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))))
- Text colors:
  - Main menu: SystemColors.ButtonHighlight
  - Submenu: Color.Black

### Button Styling
1. Button Properties to Match
- FlatStyle = FlatStyle.Flat
- FlatAppearance.BorderSize = 0
- TextAlign = ContentAlignment.MiddleLeft
- Padding = new Padding(35, 0, 0, 0)

## Code Locations
1. Menu Structure:
- panelCalibrationMenu: Around line 400
- btnSpare3 definition: Around line 400-420
- Meters menu components: Lines 240-328

2. Visual Properties:
- Font definitions: Throughout component definitions
- Color definitions: Throughout button style definitions
- Button styling: In each button's property section

Note: Line numbers are approximate and may vary. Use the component names to locate the relevant sections.
