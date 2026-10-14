VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents CellLabel As MSForms.Label
Attribute CellLabel.VB_VarHelpID = -1
Public WithEvents TextLabel As MSForms.Label
Attribute TextLabel.VB_VarHelpID = -1
Public ParentCal As Object
Public DayNumber As Integer
Public IsCurrentMonth As Boolean
Public DefaultBackColor As Long
Public DefaultTextColor As Long

Private Sub CellLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsCurrentMonth Then
        ParentCal.ApplyDayHover CellLabel
    End If
End Sub

Private Sub TextLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsCurrentMonth Then
        ParentCal.ApplyDayHover CellLabel
    End If
End Sub

Private Sub CellLabel_Click()
    If IsCurrentMonth Then
        ParentCal.SelectDay DayNumber
    End If
End Sub

Private Sub TextLabel_Click()
    If IsCurrentMonth Then
        ParentCal.SelectDay DayNumber
    End If
End Sub
