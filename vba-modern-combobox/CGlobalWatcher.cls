VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGlobalWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents WatchTextBox As MSForms.TextBox
Attribute WatchTextBox.VB_VarHelpID = -1
Public WithEvents WatchButton As MSForms.CommandButton
Attribute WatchButton.VB_VarHelpID = -1
Public WithEvents WatchLabel As MSForms.Label
Attribute WatchLabel.VB_VarHelpID = -1
Public WithEvents WatchFrame As MSForms.Frame
Attribute WatchFrame.VB_VarHelpID = -1
Public ParentCombo As CModernCombo

Private Sub WatchTextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ParentCombo.CloseDropdown
End Sub

Private Sub WatchButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ParentCombo.CloseDropdown
End Sub

Private Sub WatchLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ParentCombo.CloseDropdown
End Sub

Private Sub WatchFrame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ParentCombo.CloseDropdown
End Sub
