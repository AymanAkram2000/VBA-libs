VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModernItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents BgLabel As MSForms.Label
Attribute BgLabel.VB_VarHelpID = -1
Public WithEvents TextLabel As MSForms.Label
Attribute TextLabel.VB_VarHelpID = -1
Public ParentCombo As CModernCombo
Public RowIndex As Long
Public DisplayText As String
Public BoundVal As Variant
Public SourceRowIndex As Long

Private Sub BgLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ParentCombo.HighlightItem Me
End Sub

Private Sub BgLabel_Click()
    ParentCombo.SelectItem Me
End Sub

Private Sub TextLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ParentCombo.HighlightItem Me
End Sub

Private Sub TextLabel_Click()
    ParentCombo.SelectItem Me
End Sub
