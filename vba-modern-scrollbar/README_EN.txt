2. النسخة الإنجليزية (English README - `README.md`)

```markdown
# VBA Modern ScrollBar

A lightweight, modern, flat custom scrollbar component for Excel VBA UserForms (MSForms). Built to replace outdated Windows 95-style scrollbars with sleek vertical and horizontal slim bars, RTL (Right-To-Left) Arabic layout support, custom Hex/RGB theming, and full keyboard navigation.

---

### Key Features
* **Native RTL (Right-to-Left) Support:** When setting `RTL:=True`, the horizontal scrollbar originates from the far right, the vertical bar docks to the left edge to preserve text readability, and arrow keys invert navigation direction naturally for Arabic interfaces.
* **Custom Color Theming:** Define custom palettes using CSS Hex (`#RRGGBB`) or standard RGB `(R, G, B)` for the track, default thumb, and active/pressed thumb states.
* **Fixed Viewport Pinning:** The vertical bar stays permanently docked to the frame's visible edge, even during deep horizontal scrolling.
* **Dual-Axis Support:** Run vertical-only, horizontal-only, or both simultaneously without corner collisions.
* **Full Keyboard Accessibility:** Arrow key navigation (Up/Down/Left/Right) automatically engages upon clicking the frame or scrollbar.
* **Compact Shorthand API:** Fast single-line setup using intuitive shorthand flags (`"VH"`, `"V"`, `"H"`).
* **Configurable Thickness & Step Size:** Easily customize bar thickness and scroll distance per keypress.

---

### Repository Structure
* `CModernScrollBar.cls`: The standalone core scrollbar engine class.

---

### Installation
1. Press `ALT + F11` to open the VBA IDE.
2. Right-click inside the Project Explorer pane and select **Import File...**.
3. Choose `CModernScrollBar.cls`.

---

### Quick Start

Inside your `UserForm` code module:

```vba
Option Explicit

Private Scroll1 As CModernScrollBar

Private Sub UserForm_Initialize()
    Set Scroll1 = New CModernScrollBar
    
    ' Quick RTL setup for Arabic UserForms:
    Scroll1.Apply Me.Frame1, "VH", RTL:=True
    
    ' Advanced setup with custom Hex theme:
    ' Scroll1.Apply Me.Frame1, "VH", "L", "B", 6, 20, True, _
    '               "#F0ECE8;#C5BFB7;#8E857B"
End Sub



Scroll1.Apply Target, [Mode], [VSide], [HSide], [BarThickness], [StepSize], [RTL], [Theme]