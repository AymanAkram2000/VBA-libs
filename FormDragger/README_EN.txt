========================================================================
                      VBA FORM DRAGGER LIBRARY
========================================================================

Description:
A lightweight library enabling smooth dragging and moving of frameless 
Excel UserForms via form background clicks or designated controls.

------------------------------------------------------------------------
[1] REQUIRED FILES
------------------------------------------------------------------------
  1. modDragForm.bas
  2. clsDragHandler.cls

------------------------------------------------------------------------
[2] HOW TO USE
------------------------------------------------------------------------
1. Import both files into your target VBA Project.
2. Add one line inside your UserForm_Initialize event:

   ' Option A: Drag using the entire form background:
   Private Sub UserForm_Initialize()
       MakeDraggable Me, Me
   End Sub

   ' Option B: Drag using a specific control (e.g. lblTitle):
   Private Sub UserForm_Initialize()
       MakeDraggable Me.lblTitle, Me
   End Sub

========================================================================