VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents TargetForm As MSForms.UserForm
Attribute TargetForm.VB_VarHelpID = -1
Public WithEvents TargetLabel As MSForms.Label
Attribute TargetLabel.VB_VarHelpID = -1
Public WithEvents TargetFrame As MSForms.Frame
Attribute TargetFrame.VB_VarHelpID = -1
Public WithEvents TargetButton As MSForms.CommandButton
Attribute TargetButton.VB_VarHelpID = -1
Public WithEvents TargetOtherBox As MSForms.TextBox
Attribute TargetOtherBox.VB_VarHelpID = -1

Private Sub TargetForm_Click()
    modCalendar.CloseCalendar
End Sub

Private Sub TargetLabel_Click()
    modCalendar.CloseCalendar
End Sub

Private Sub TargetFrame_Click()
    modCalendar.CloseCalendar
End Sub

Private Sub TargetButton_Click()
    modCalendar.CloseCalendar
End Sub

Private Sub TargetOtherBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modCalendar.CloseCalendar
End Sub
