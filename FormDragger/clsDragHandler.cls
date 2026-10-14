VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDragHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TargetForm As Object
Private WithEvents FormObj As MSForms.UserForm
Attribute FormObj.VB_VarHelpID = -1
Private WithEvents LabelObj As MSForms.Label
Attribute LabelObj.VB_VarHelpID = -1
Private WithEvents ImageObj As MSForms.Image
Attribute ImageObj.VB_VarHelpID = -1
Private WithEvents FrameObj As MSForms.Frame
Attribute FrameObj.VB_VarHelpID = -1

Private IsMouseDown As Boolean
Private StartX As Single
Private StartY As Single

Public Sub Hook(ByVal TargetControl As Object, ByVal ParentUserForm As Object)
    Set TargetForm = ParentUserForm
    
    If TypeOf TargetControl Is MSForms.UserForm Then
        Set FormObj = TargetControl
    ElseIf TypeOf TargetControl Is MSForms.Label Then
        Set LabelObj = TargetControl
    ElseIf TypeOf TargetControl Is MSForms.Image Then
        Set ImageObj = TargetControl
    ElseIf TypeOf TargetControl Is MSForms.Frame Then
        Set FrameObj = TargetControl
    End If
End Sub

' --- Form Events ---
Private Sub FormObj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        IsMouseDown = True
        StartX = X
        StartY = Y
    End If
End Sub

Private Sub FormObj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsMouseDown Then
        TargetForm.Left = TargetForm.Left + (X - StartX)
        TargetForm.Top = TargetForm.Top + (Y - StartY)
    End If
End Sub

Private Sub FormObj_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    IsMouseDown = False
End Sub

' --- Label Events ---
Private Sub LabelObj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        IsMouseDown = True
        StartX = X
        StartY = Y
    End If
End Sub

Private Sub LabelObj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsMouseDown Then
        TargetForm.Left = TargetForm.Left + (X - StartX)
        TargetForm.Top = TargetForm.Top + (Y - StartY)
    End If
End Sub

Private Sub LabelObj_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    IsMouseDown = False
End Sub

' --- Image Events ---
Private Sub ImageObj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        IsMouseDown = True
        StartX = X
        StartY = Y
    End If
End Sub

Private Sub ImageObj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsMouseDown Then
        TargetForm.Left = TargetForm.Left + (X - StartX)
        TargetForm.Top = TargetForm.Top + (Y - StartY)
    End If
End Sub

Private Sub ImageObj_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    IsMouseDown = False
End Sub

' --- Frame Events ---
Private Sub FrameObj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        IsMouseDown = True
        StartX = X
        StartY = Y
    End If
End Sub

Private Sub FrameObj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsMouseDown Then
        TargetForm.Left = TargetForm.Left + (X - StartX)
        TargetForm.Top = TargetForm.Top + (Y - StartY)
    End If
End Sub

Private Sub FrameObj_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    IsMouseDown = False
End Sub
