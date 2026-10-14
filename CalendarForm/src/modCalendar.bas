Attribute VB_Name = "modCalendar"
Option Explicit

Public GlobalDateHandlers As Collection
Public GlobalClickHandlers As Collection

Public Sub AttachCalendar(ByRef TargetForm As Object)
    Dim ctl As MSForms.Control
    Dim dateHandler As clsDateBox
    Dim clickHandler As clsFormClick
    
    Set GlobalDateHandlers = New Collection
    Set GlobalClickHandlers = New Collection
    
    If TypeOf TargetForm Is MSForms.UserForm Then
        Set clickHandler = New clsFormClick
        Set clickHandler.TargetForm = TargetForm
        GlobalClickHandlers.Add clickHandler
    End If
    
    For Each ctl In TargetForm.Controls
        If TypeName(ctl) = "TextBox" Then
            If Right(LCase(Trim(ctl.Name)), 5) = "_date" Then
                Set dateHandler = New clsDateBox
                Set dateHandler.TxtBox = ctl
                Set dateHandler.ParentForm = TargetForm
                GlobalDateHandlers.Add dateHandler
            Else
                Set clickHandler = New clsFormClick
                Set clickHandler.TargetOtherBox = ctl
                GlobalClickHandlers.Add clickHandler
            End If
        ElseIf TypeName(ctl) = "Label" Then
            Set clickHandler = New clsFormClick
            Set clickHandler.TargetLabel = ctl
            GlobalClickHandlers.Add clickHandler
        ElseIf TypeName(ctl) = "CommandButton" Then
            Set clickHandler = New clsFormClick
            Set clickHandler.TargetButton = ctl
            GlobalClickHandlers.Add clickHandler
        ElseIf TypeName(ctl) = "Frame" Then
            Set clickHandler = New clsFormClick
            Set clickHandler.TargetFrame = ctl
            GlobalClickHandlers.Add clickHandler
        End If
    Next ctl
End Sub

Public Sub CloseCalendar()
    Dim frm As Object
    For Each frm In UserForms
        If frm.Name = "frmCalendar" Then
            Unload frmCalendar
            Exit For
        End If
    Next frm
End Sub
