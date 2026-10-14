========================================================================
             VBA POPUP CALENDAR - INSTALLATION & USAGE GUIDE
========================================================================

Description:
A modern, lightweight popup calendar library for Excel UserForms.
Features 1-line hookup, direct manual typing, double-click popups, 
smart outside-click auto-closing, and rounded border design.

------------------------------------------------------------------------
[1] REQUIRED FILES (4 Files Only)
------------------------------------------------------------------------
Import these 4 files into your VBA Project:
  1. frmCalendar.frm   (Calendar Form & UI Engine)
  2. modCalendar.bas   (Bootstrap & Auto-Hooking Module)
  3. clsCalDay.cls     (Day Cell Formatting & Hover Controller)
  4. clsDateBox.cls    (Date TextBoxes Event Listener)

------------------------------------------------------------------------
[2] CRITICAL SETUP REQUIREMENT (ShowModal = False)
------------------------------------------------------------------------
Your UserForm MUST be displayed in Modeless mode for the calendar 
to open alongside it and allow seamless typing and auto-closing.

Choose either method:
  • Method A (From Properties Window):
    Select your UserForm, open Properties (F4), and set:
    ShowModal = False

  • Method B (From Code):
    When opening your form via code, use:
    YourFormName.Show vbModeless

------------------------------------------------------------------------
[3] HOW TO USE IN YOUR USERFORM
------------------------------------------------------------------------
Step 1: TextBox Naming
----------------------
Any date TextBox must have a name ending with: _Date
(e.g., txtBirth_Date, Hire_Date, Order_Date).

Step 2: Activation Code
-----------------------
Add this single line inside your UserForm's Initialize event:

    Private Sub UserForm_Initialize()
        AttachCalendar Me
    End Sub

------------------------------------------------------------------------
[4] FEATURES & SHORTCUTS
------------------------------------------------------------------------
- Single Click: Focuses the box for direct manual typing (DD/MM/YYYY).
- Double-Click or [F4]: Opens the calendar near the cursor.
- Outside Click: Clicking anywhere outside instantly auto-closes the calendar.
- Selection: Clicking any date writes the value and auto-closes.
========================================================================