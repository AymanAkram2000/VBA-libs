### النسخة الإنجليزية (English README - `README.md`)

```markdown
# VBA Modern ComboBox

A clean, responsive, flat searchable dropdown component for Excel VBA UserForms. It transforms standard `TextBox` controls into autocomplete search pickers with an integrated modern scrollbar, dynamic row height controls, precise vertical text centering, and a full custom Theming Engine.

---

### Key Features
* **Full Custom Theming:** Set custom palettes using CSS Hex (`#RRGGBB`) or standard `(R, G, B)` syntax for background, hover, font, border, and scrollbar components.
* **Row-Height & Geometry Control:** The `H` parameter in offsets directly controls individual item row heights while automatically maintaining vertical text centering.
* **Dynamic Geometry Offsets:** Full custom control over placement and dimensions (`X, Y, Width, RowHeight`) relative to the anchor textbox.
* **Instant Substring Search:** Real-time filtering across items as you type.
* **Full Keyboard Navigation:**
  * Traverse items using `Up` and `Down` arrow keys.
  * Smooth `ScrollIntoView` ensuring highlighted entries are visible.
  * Up-arrow wraps back to the textbox for quick editing.
  * Commit entries with `Enter` or `Spacebar`.
* **Global Click-Away Dismissal:** Menu automatically dismisses when clicking anywhere outside.
* **Strict Validation (`MatchRequired`):** Restricts free-text input and auto-reverts invalid values.
* **Direct Sheet Row Lookup (`.Row`):** Directly exposes the true row number from the underlying worksheet.

---

### Required Files
* `CModernCombo.cls`: Core dropdown logic, data loader, positioning, layout, and theming engine.
* `CModernItem.cls`: Individual row wrapper handling mouse clicks and active styling.
* `CGlobalWatcher.cls`: Global event hook intercepting outside clicks to auto-close the menu.

---

### Quick Start

Inside your `UserForm` code module:

```vba
Option Explicit

Private Combo1 As CModernCombo

Private Sub UserForm_Initialize()
    Set Combo1 = New CModernCombo
    
    ' Attach with positional arguments, geometry offsets, and custom hex theme:
    ' Attach(TargetBox, Source, [DisplayCol], [BoundCol], [MaxRows], [MatchReq], [Offsets], [Theme])
    
    Combo1.Attach Me.TextBox1, Sheets("Sheet1").Range("A11:B50"), 2, 1, 5, True, _
                  "0,0,30,-2", _
                  "#FFFFFF;#E8F0FE;#1A73E8;#DADCE0;#F1F3F4;#BDC1C6"
End Sub

Private Sub CommandButton1_Click()
    MsgBox "Row: " & Combo1.Row & vbCrLf & _
           "Value: " & Combo1.Value
End Sub