VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModernScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' CModernScrollBar - Universal Flat Scrollbar (Fixed Viewport & Custom Theme)
' =========================================================================

Private WithEvents m_TargetFrame As MSForms.Frame
Attribute m_TargetFrame.VB_VarHelpID = -1
Private WithEvents m_VTrack As MSForms.Label
Attribute m_VTrack.VB_VarHelpID = -1
Private WithEvents m_VThumb As MSForms.Label
Attribute m_VThumb.VB_VarHelpID = -1
Private WithEvents m_HTrack As MSForms.Label
Attribute m_HTrack.VB_VarHelpID = -1
Private WithEvents m_HThumb As MSForms.Label
Attribute m_HThumb.VB_VarHelpID = -1
Private WithEvents m_KeyReceiver As MSForms.TextBox
Attribute m_KeyReceiver.VB_VarHelpID = -1

Private m_EnableV As Boolean
Private m_EnableH As Boolean
Private m_VAlignLeft As Boolean
Private m_HAlignTop As Boolean

Private m_BarThickness As Double
Private m_StepSize As Double

Private m_IsDraggingV As Boolean
Private m_DragStartY As Single
Private m_IsDraggingH As Boolean
Private m_DragStartX As Single

' Theme Colors
Private m_ClrTrack As Long
Private m_ClrThumbNormal As Long
Private m_ClrThumbActive As Long

' Mode: "VH" (Both), "V" (Vertical Only), "H" (Horizontal Only)
' VSide: "R" (Right - Default), "L" (Left)
' HSide: "B" (Bottom - Default), "T" (Top)
' Theme: "TrackColor; ThumbNormal; ThumbActive" (Supports Hex #RRGGBB or RGB (R,G,B))
Public Sub Apply(ByVal Target As Object, _
                 Optional ByVal Mode As String = "VH", _
                 Optional ByVal VSide As String = "R", _
                 Optional ByVal HSide As String = "B", _
                 Optional ByVal BarThickness As Double = 6, _
                 Optional ByVal StepSize As Double = 20, _
                 Optional ByVal Theme As String = "")
    
    If Not TypeOf Target Is MSForms.Frame Then Exit Sub
    
    Set m_TargetFrame = Target
    m_BarThickness = BarThickness
    m_StepSize = StepSize
    
    Mode = UCase(Trim(Mode))
    m_EnableV = (InStr(Mode, "V") > 0 Or Mode = "BOTH" Or Mode = "")
    m_EnableH = (InStr(Mode, "H") > 0 Or Mode = "BOTH")
    
    m_VAlignLeft = (UCase(Trim(VSide)) = "L")
    m_HAlignTop = (UCase(Trim(HSide)) = "T")
    
    ParseTheme Theme
    
    m_TargetFrame.ScrollBars = fmScrollBarsNone
    m_TargetFrame.KeepScrollBarsVisible = fmScrollBarsNone
    
    BuildControls
    CreateKeyReceiver
    AutoCalculateLimits
    UpdateLayout
End Sub

Public Sub Refresh()
    AutoCalculateLimits
    UpdateLayout
End Sub

Private Sub ParseTheme(ByVal rawStr As String)
    ' Default Clean Neutral Palette
    m_ClrTrack = RGB(245, 242, 238)
    m_ClrThumbNormal = RGB(197, 191, 183)
    m_ClrThumbActive = RGB(142, 133, 123)
    
    rawStr = Trim(rawStr)
    If rawStr = "" Then Exit Sub
    
    Dim tokens() As String
    If InStr(rawStr, ";") > 0 Then
        tokens = Split(rawStr, ";")
    Else
        tokens = Split(rawStr, "|")
    End If
    
    On Error Resume Next
    If UBound(tokens) >= 0 Then m_ClrTrack = ConvertToColor(tokens(0), m_ClrTrack)
    If UBound(tokens) >= 1 Then m_ClrThumbNormal = ConvertToColor(tokens(1), m_ClrThumbNormal)
    If UBound(tokens) >= 2 Then m_ClrThumbActive = ConvertToColor(tokens(2), m_ClrThumbActive)
    On Error GoTo 0
End Sub

Private Function ConvertToColor(ByVal colStr As String, ByVal defaultCol As Long) As Long
    colStr = Trim(colStr)
    If colStr = "" Then ConvertToColor = defaultCol: Exit Function
    
    On Error GoTo Fallback
    ' Hex format: #FFFFFF or FFFFFF
    If Left(colStr, 1) = "#" Then colStr = Mid(colStr, 2)
    If Len(colStr) = 6 And InStr(colStr, ",") = 0 Then
        Dim r As Long, g As Long, b As Long
        r = val("&H" & Left(colStr, 2))
        g = val("&H" & Mid(colStr, 3, 2))
        b = val("&H" & Right(colStr, 2))
        ConvertToColor = RGB(r, g, b)
        Exit Function
    End If
    
    ' RGB format: (255, 255, 255) or 255, 255, 255
    colStr = Replace(Replace(colStr, "(", ""), ")", "")
    Dim rgbParts() As String
    rgbParts = Split(colStr, ",")
    If UBound(rgbParts) >= 2 Then
        ConvertToColor = RGB(val(Trim(rgbParts(0))), val(Trim(rgbParts(1))), val(Trim(rgbParts(2))))
        Exit Function
    End If

Fallback:
    ConvertToColor = defaultCol
End Function

Private Sub AutoCalculateLimits()
    Dim ctrl As MSForms.Control
    Dim maxBottom As Double, maxRight As Double
    
    maxBottom = m_TargetFrame.InsideHeight
    maxRight = m_TargetFrame.InsideWidth
    
    For Each ctrl In m_TargetFrame.Controls
        If ctrl.Name <> "trck_V" And ctrl.Name <> "thmb_V" And _
           ctrl.Name <> "trck_H" And ctrl.Name <> "thmb_H" And _
           ctrl.Name <> "keyBox_Global" Then
            If ctrl.Visible Then
                If (ctrl.Top + ctrl.Height) > maxBottom Then maxBottom = ctrl.Top + ctrl.Height + 10
                If (ctrl.Left + ctrl.Width) > maxRight Then maxRight = ctrl.Left + ctrl.Width + 10
            End If
        End If
    Next ctrl
    
    If maxBottom > m_TargetFrame.ScrollHeight Then m_TargetFrame.ScrollHeight = maxBottom
    If maxRight > m_TargetFrame.ScrollWidth Then m_TargetFrame.ScrollWidth = maxRight
End Sub

Private Sub BuildControls()
    On Error Resume Next
    m_TargetFrame.Controls.Remove "trck_V"
    m_TargetFrame.Controls.Remove "thmb_V"
    m_TargetFrame.Controls.Remove "trck_H"
    m_TargetFrame.Controls.Remove "thmb_H"
    On Error GoTo 0
    
    Set m_VTrack = m_TargetFrame.Controls.Add("Forms.Label.1", "trck_V", True)
    With m_VTrack
        .BackColor = m_ClrTrack
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .Caption = ""
        .Visible = False
    End With
    
    Set m_VThumb = m_TargetFrame.Controls.Add("Forms.Label.1", "thmb_V", True)
    With m_VThumb
        .BackColor = m_ClrThumbNormal
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .Caption = ""
        .Visible = False
    End With
    
    Set m_HTrack = m_TargetFrame.Controls.Add("Forms.Label.1", "trck_H", True)
    With m_HTrack
        .BackColor = m_ClrTrack
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .Caption = ""
        .Visible = False
    End With
    
    Set m_HThumb = m_TargetFrame.Controls.Add("Forms.Label.1", "thmb_H", True)
    With m_HThumb
        .BackColor = m_ClrThumbNormal
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .Caption = ""
        .Visible = False
    End With
End Sub

Private Sub CreateKeyReceiver()
    On Error Resume Next
    m_TargetFrame.Controls.Remove "keyBox_Global"
    On Error GoTo 0
    
    Set m_KeyReceiver = m_TargetFrame.Controls.Add("Forms.TextBox.1", "keyBox_Global", True)
    With m_KeyReceiver
        .Left = -50
        .Top = -50
        .Width = 10
        .Height = 10
        .TabStop = False
    End With
End Sub

Public Sub UpdateLayout()
    If m_TargetFrame Is Nothing Then Exit Sub
    
    Dim viewH As Double, totalH As Double
    Dim viewW As Double, totalW As Double
    Dim hasV As Boolean, hasH As Boolean
    
    viewH = m_TargetFrame.InsideHeight
    totalH = m_TargetFrame.ScrollHeight
    viewW = m_TargetFrame.InsideWidth
    totalW = m_TargetFrame.ScrollWidth
    
    hasV = (m_EnableV And totalH > viewH And viewH > 0)
    hasH = (m_EnableH And totalW > viewW And viewW > 0)
    
    ' Position Vertical Bar (Pinned to Viewport X & Y)
    If hasV Then
        Dim vx As Double
        If m_VAlignLeft Then
            vx = m_TargetFrame.ScrollLeft + 1
        Else
            vx = m_TargetFrame.ScrollLeft + viewW - m_BarThickness - 1
        End If
        
        Dim vy As Double
        If hasH And m_HAlignTop Then
            vy = m_TargetFrame.ScrollTop + m_BarThickness + 2
        Else
            vy = m_TargetFrame.ScrollTop + 1
        End If
        
        Dim trackH As Double
        trackH = viewH - IIf(hasH, m_BarThickness + 3, 2)
        
        With m_VTrack
            .Left = vx
            .Top = vy
            .Width = m_BarThickness
            .Height = trackH
            .Visible = True
            .ZOrder 0
        End With
        
        Dim thumbH As Double
        thumbH = (viewH / totalH) * trackH
        If thumbH < 18 Then thumbH = 18
        
        Dim maxScrollY As Double: maxScrollY = totalH - viewH
        Dim ratioY As Double: If maxScrollY > 0 Then ratioY = m_TargetFrame.ScrollTop / maxScrollY Else ratioY = 0
        
        With m_VThumb
            .Left = m_VTrack.Left
            .Width = m_BarThickness
            .Height = thumbH
            .Top = m_VTrack.Top + (ratioY * (m_VTrack.Height - thumbH))
            .Visible = True
            .ZOrder 0
        End With
    Else
        If Not m_VTrack Is Nothing Then m_VTrack.Visible = False
        If Not m_VThumb Is Nothing Then m_VThumb.Visible = False
    End If
    
    ' Position Horizontal Bar (Pinned to Viewport X & Y)
    If hasH Then
        Dim hx As Double
        If hasV And m_VAlignLeft Then
            hx = m_TargetFrame.ScrollLeft + m_BarThickness + 2
        Else
            hx = m_TargetFrame.ScrollLeft + 1
        End If
        
        Dim hy As Double
        If m_HAlignTop Then
            hy = m_TargetFrame.ScrollTop + 1
        Else
            hy = m_TargetFrame.ScrollTop + viewH - m_BarThickness - 1
        End If
        
        Dim trackW As Double
        trackW = viewW - IIf(hasV, m_BarThickness + 3, 2)
        
        With m_HTrack
            .Left = hx
            .Top = hy
            .Height = m_BarThickness
            .Width = trackW
            .Visible = True
            .ZOrder 0
        End With
        
        Dim thumbW As Double
        thumbW = (viewW / totalW) * trackW
        If thumbW < 18 Then thumbW = 18
        
        Dim maxScrollX As Double: maxScrollX = totalW - viewW
        Dim ratioX As Double: If maxScrollX > 0 Then ratioX = m_TargetFrame.ScrollLeft / maxScrollX Else ratioX = 0
        
        With m_HThumb
            .Top = m_HTrack.Top
            .Height = m_BarThickness
            .Width = thumbW
            .Left = m_HTrack.Left + (ratioX * (m_HTrack.Width - thumbW))
            .Visible = True
            .ZOrder 0
        End With
    Else
        If Not m_HTrack Is Nothing Then m_HTrack.Visible = False
        If Not m_HThumb Is Nothing Then m_HThumb.Visible = False
    End If
End Sub

Private Sub FocusKeyboard()
    On Error Resume Next
    m_KeyReceiver.SetFocus
    On Error GoTo 0
End Sub

' === Vertical Events ===

Private Sub m_VThumb_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        FocusKeyboard
        m_IsDraggingV = True
        m_DragStartY = Y
        m_VThumb.BackColor = m_ClrThumbActive
    End If
End Sub

Private Sub m_VThumb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If m_IsDraggingV Then
        Dim newRelTop As Double, maxRelTop As Double, travelDist As Double
        
        newRelTop = (m_VThumb.Top - m_VTrack.Top) + (Y - m_DragStartY)
        If newRelTop < 0 Then newRelTop = 0
        
        maxRelTop = m_VTrack.Height - m_VThumb.Height
        If newRelTop > maxRelTop Then newRelTop = maxRelTop
        
        travelDist = maxRelTop
        If travelDist > 0 Then
            Dim ratio As Double
            ratio = newRelTop / travelDist
            m_TargetFrame.ScrollTop = ratio * (m_TargetFrame.ScrollHeight - m_TargetFrame.InsideHeight)
            UpdateLayout
        End If
    End If
End Sub

Private Sub m_VThumb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    m_IsDraggingV = False
    m_VThumb.BackColor = m_ClrThumbNormal
End Sub

Private Sub m_VTrack_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        FocusKeyboard
        Dim maxScroll As Double
        maxScroll = m_TargetFrame.ScrollHeight - m_TargetFrame.InsideHeight
        
        If Y < (m_VThumb.Top - m_VTrack.Top) Then
            m_TargetFrame.ScrollTop = Application.Max(0, m_TargetFrame.ScrollTop - m_TargetFrame.InsideHeight)
        Else
            m_TargetFrame.ScrollTop = Application.Min(maxScroll, m_TargetFrame.ScrollTop + m_TargetFrame.InsideHeight)
        End If
        UpdateLayout
    End If
End Sub

' === Horizontal Events ===

Private Sub m_HThumb_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        FocusKeyboard
        m_IsDraggingH = True
        m_DragStartX = X
        m_HThumb.BackColor = m_ClrThumbActive
    End If
End Sub

Private Sub m_HThumb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If m_IsDraggingH Then
        Dim newRelLeft As Double, maxRelLeft As Double, travelDistX As Double
        
        newRelLeft = (m_HThumb.Left - m_HTrack.Left) + (X - m_DragStartX)
        If newRelLeft < 0 Then newRelLeft = 0
        
        maxRelLeft = m_HTrack.Width - m_HThumb.Width
        If newRelLeft > maxRelLeft Then newRelLeft = maxRelLeft
        
        travelDistX = maxRelLeft
        If travelDistX > 0 Then
            Dim ratio As Double
            ratio = newRelLeft / travelDistX
            m_TargetFrame.ScrollLeft = ratio * (m_TargetFrame.ScrollWidth - m_TargetFrame.InsideWidth)
            UpdateLayout
        End If
    End If
End Sub

Private Sub m_HThumb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    m_IsDraggingH = False
    m_HThumb.BackColor = m_ClrThumbNormal
End Sub

Private Sub m_HTrack_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        FocusKeyboard
        Dim maxScrollX As Double
        maxScrollX = m_TargetFrame.ScrollWidth - m_TargetFrame.InsideWidth
        
        If X < (m_HThumb.Left - m_HTrack.Left) Then
            m_TargetFrame.ScrollLeft = Application.Max(0, m_TargetFrame.ScrollLeft - m_TargetFrame.InsideWidth)
        Else
            m_TargetFrame.ScrollLeft = Application.Min(maxScrollX, m_TargetFrame.ScrollLeft + m_TargetFrame.InsideWidth)
        End If
        UpdateLayout
    End If
End Sub

Private Sub m_TargetFrame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FocusKeyboard
End Sub

' === Keyboard Navigation ===

Private Sub m_KeyReceiver_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim maxScrollY As Double: maxScrollY = m_TargetFrame.ScrollHeight - m_TargetFrame.InsideHeight
    Dim maxScrollX As Double: maxScrollX = m_TargetFrame.ScrollWidth - m_TargetFrame.InsideWidth
    
    Select Case KeyCode
        Case vbKeyDown
            If maxScrollY > 0 Then
                m_TargetFrame.ScrollTop = Application.Min(maxScrollY, m_TargetFrame.ScrollTop + m_StepSize)
                UpdateLayout
            End If
            KeyCode = 0
            
        Case vbKeyUp
            If maxScrollY > 0 Then
                m_TargetFrame.ScrollTop = Application.Max(0, m_TargetFrame.ScrollTop - m_StepSize)
                UpdateLayout
            End If
            KeyCode = 0
            
        Case vbKeyRight
            If maxScrollX > 0 Then
                m_TargetFrame.ScrollLeft = Application.Min(maxScrollX, m_TargetFrame.ScrollLeft + m_StepSize)
                UpdateLayout
            End If
            KeyCode = 0
            
        Case vbKeyLeft
            If maxScrollX > 0 Then
                m_TargetFrame.ScrollLeft = Application.Max(0, m_TargetFrame.ScrollLeft - m_StepSize)
                UpdateLayout
            End If
            KeyCode = 0
    End Select
End Sub

