========================================================================
                      VBA USERFORM STYLER LIBRARY
========================================================================

Description:
Removes default Windows title bar and applies smooth rounded corners 
to Excel UserForms using Windows API.

------------------------------------------------------------------------
[1] REQUIRED FILES
------------------------------------------------------------------------
  1. modStyleUserForm.bas

------------------------------------------------------------------------
[2] HOW TO USE
------------------------------------------------------------------------
1. Import 'modStyleUserForm.bas' into your VBA Project.
2. In your UserForm code, call StyleUserForm inside UserForm_Initialize:

   Private Sub UserForm_Initialize()
       StyleUserForm Me, 20
   End Sub

Note: The parameter (20) defines the corner radius in pixels.

========================================================================