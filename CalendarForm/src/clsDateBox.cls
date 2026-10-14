VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDateBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
#Else
    Private Declare Function ReleaseCapture Lib "user32" () As Long
#End If

Public WithEvents TxtBox As MSForms.TextBox
Attribute TxtBox.VB_VarHelpID = -1
Public ParentForm As Object

Private Sub TxtBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        Dim frm As Object
        Dim isLoaded As Boolean
        
        For Each frm In UserForms
            If frm.Name = "frmCalendar" Then
                isLoaded = True
                Exit For
            End If
        Next frm
        
        If isLoaded Then
            If Not frmCalendar.TargetControl Is TxtBox Then
                Unload frmCalendar
            End If
        End If
    End If
End Sub

Private Sub TxtBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ReleaseCapture
    TxtBox.MousePointer = fmMousePointerDefault
    TriggerCalendar
End Sub

Private Sub TxtBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        TriggerCalendar
    End If
End Sub

Public Sub TriggerCalendar()
    Dim boxName As String
    boxName = LCase(Trim(TxtBox.Name))
    
    If Right(boxName, 5) = "_date" Then
        On Error Resume Next
        Unload frmCalendar
        On Error GoTo 0
        
        DoEvents
        Load frmCalendar
        frmCalendar.ShowNearCursor TxtBox
    End If
End Sub
