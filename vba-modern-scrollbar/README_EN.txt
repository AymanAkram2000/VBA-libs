### النسخة الإنجليزية (English README - `README.md`)

```markdown
# VBA Modern ScrollBar

A lightweight, modern, flat custom scrollbar component for Excel VBA UserForms (MSForms). Built to replace outdated default scrollbars with sleek vertical and horizontal slim bars, full mouse drag physics, keyboard navigation, and custom color theming.

---

### Key Features
* **Custom Color Theming:** Define custom colors using CSS Hex (`#RRGGBB`) or standard RGB `(R, G, B)` for the track, default thumb, and active/pressed thumb states.
* **Dual-Axis Support:** Run vertical-only, horizontal-only, or both simultaneously without corner collisions.
* **Fixed Viewport Pinning:** The vertical bar stays permanently docked to the frame's visible edge, even during horizontal scrolling.
* **Full Keyboard Accessibility:** Arrow key navigation (Up/Down/Left/Right) automatically engages upon clicking the frame or scrollbar.
* **Compact Shorthand API:** Fast single-line setup using intuitive shorthand flags (`"VH"`, `"V"`, `"H"`).
* **Configurable Thickness & Steps:** Customize bar thickness and scroll speed per arrow key press.

---

### Required Files
* `CModernScrollBar.cls`: The standalone core scrollbar engine class.

---

### Quick Start

Inside your `UserForm` code module:

```vba
Option Explicit

Private Scroll1 As CModernScrollBar

Private Sub UserForm_Initialize()
    Set Scroll1 = New CModernScrollBar
    
    ' Apply dual scrollbars with custom Hex colors:
    ' Apply(Target, [Mode], [VSide], [HSide], [BarThickness], [StepSize], [Theme])
    
    Scroll1.Apply Me.Frame1, "VH", "R", "B", 6, 20, _
                  "#F0ECE8;#C5BFB7;#8E857B"
End Sub
Parameter Reference:
Target: Target MSForms.Frame control (Required).

Mode:

"VH": Enable both Vertical and Horizontal (Default).

"V": Vertical scrollbar only.

"H": Horizontal scrollbar only.

VSide: Vertical bar placement ("R" for Right [default], "L" for Left).

HSide: Horizontal bar placement ("B" for Bottom [default], "T" for Top).

BarThickness: Bar thickness in points/pixels (Default: 6).

StepSize: Scrolling distance per arrow key press (Default: 20).

Theme: Semicolon-delimited color string supporting Hex (#RRGGBB) or RGB (R, G, B):

Order: TrackColor;ThumbNormal;ThumbActive.