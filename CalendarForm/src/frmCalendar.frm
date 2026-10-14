VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "UserForm1"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3525
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" ( _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long, _
        ByVal X3 As Long, _
        ByVal Y4 As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowRgn Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal hRgn As LongPtr, _
        ByVal bRedraw As Boolean) As Long
#Else
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function ReleaseCapture Lib "user32" () As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long, _
        ByVal X3 As Long, _
        ByVal Y4 As Long) As Long
    Private Declare Function SetWindowRgn Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
#End If

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_THICKFRAME As Long = &H40000
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Private HOLD As Boolean
Private sx As Single, sy As Single
Private cx As Single, cy As Single
Private CurrentDate As Date
Public SelectedDate As String
Private DayButtons As Collection
Private LastHovered As MSForms.Label
Public TargetControl As Object
Public PreSelectedDate As Date



Private Sub lblMonthYear_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CurrentDate = Date
    RenderCalendar
End Sub

Private Sub lblMonthYear_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HOLD = True
    sx = X
    sy = Y
End Sub

Private Sub lblMonthYear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If HOLD = True Then
        cx = X: cy = Y
        Me.Top = Me.Top + (cy - sy)
        Me.Left = Me.Left + (cx - sx)
    End If
End Sub

Private Sub lblMonthYear_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HOLD = False
End Sub

Private Sub UserForm_Initialize()
    Me.fraDays.BackColor = RGB(253, 249, 249)
    ApplyFormStyle Me, 20
    CurrentDate = Date
    RenderCalendar
End Sub

Private Sub ApplyFormStyle(ByRef frm As Object, Optional ByVal Radius As Long = 20)
    #If VBA7 Then
        Dim hwnd As LongPtr
        Dim hRgn As LongPtr
    #Else
        Dim hwnd As Long
        Dim hRgn As Long
    #End If
    
    Dim currentStyle As Long
    Dim widthPx As Long
    Dim heightPx As Long
    
    hwnd = FindWindow("ThunderDFrame", frm.Caption)
    If hwnd = 0 Then Exit Sub
    
    currentStyle = GetWindowLong(hwnd, GWL_STYLE)
    currentStyle = currentStyle And Not WS_CAPTION
    currentStyle = currentStyle And Not WS_THICKFRAME
    SetWindowLong hwnd, GWL_STYLE, currentStyle
    
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
    widthPx = CLng(frm.Width * 1.3333)
    heightPx = CLng(frm.Height * 1.3333)
    
    hRgn = CreateRoundRectRgn(0, 0, widthPx, heightPx, Radius, Radius)
    SetWindowRgn hwnd, hRgn, True
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HOLD = True
    sx = X
    sy = Y
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If HOLD = True Then
        cx = X: cy = Y
        Me.Top = Me.Top + (cy - sy)
        Me.Left = Me.Left + (cx - sx)
    End If
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HOLD = False
End Sub

Public Sub ShowNearCursor(ByVal OutputControl As Object)
    ReleaseCapture
    Me.MousePointer = fmMousePointerDefault
    
    Set TargetControl = OutputControl
    PreSelectedDate = 0
    
    If Not TargetControl Is Nothing Then
        On Error Resume Next
        If IsDate(TargetControl.Text) Then
            PreSelectedDate = CDate(TargetControl.Text)
            CurrentDate = PreSelectedDate
        Else
            CurrentDate = Date
        End If
        On Error GoTo 0
    Else
        CurrentDate = Date
    End If

    RenderCalendar

    Me.StartUpPosition = 0

    Dim pt As POINTAPI
    GetCursorPos pt

    Dim mouseX As Single, mouseY As Single
    Dim screenW As Single, screenH As Single

    mouseX = pt.X * 0.75
    mouseY = pt.Y * 0.75
    screenW = GetSystemMetrics(SM_CXSCREEN) * 0.75
    screenH = GetSystemMetrics(SM_CYSCREEN) * 0.75

    If mouseX + Me.Width > screenW Then
        Me.Left = mouseX - Me.Width
    Else
        Me.Left = mouseX + 10
    End If

    If mouseY + Me.Height + 10 > screenH Then
        Me.Top = mouseY - Me.Height - 5
    Else
        Me.Top = mouseY + 10
    End If

    DoEvents

    On Error Resume Next
    Me.Show vbModeless
    If Err.Number = 401 Then
        Err.Clear
        Me.Show vbModal
    End If
    On Error GoTo 0
End Sub

Public Sub RenderCalendar()
    lblMonthYear.Caption = Format(CurrentDate, "mmmm yyyy")
    lblMonthYear.ForeColor = RGB(59, 42, 32)

    Dim c As Control
    For Each c In fraDays.Controls
        fraDays.Controls.Remove c.Name
    Next c

    Set DayButtons = New Collection
    Set LastHovered = Nothing

    Dim FirstDayOfMonth As Date
    FirstDayOfMonth = DateSerial(Year(CurrentDate), Month(CurrentDate), 1)

    Dim StartWeekday As Integer
    StartWeekday = Weekday(FirstDayOfMonth, vbSunday)

    Dim DaysInMonth As Integer
    DaysInMonth = Day(DateSerial(Year(CurrentDate), Month(CurrentDate) + 1, 0))

    Dim Row As Integer, Col As Integer
    Dim Index As Integer, DayNum As Integer
    Dim lblCell As MSForms.Label
    Dim lblText As MSForms.Label
    Dim CalObj As clsCalDay
    Dim ThisDayDate As Date

    Dim CellW As Single: CellW = 26
    Dim CellH As Single: CellH = 22
    Dim TextH As Single: TextH = 14

    For Index = 1 To 42
        Col = (Index - 1) Mod 7
        Row = (Index - 1) \ 7

        Set lblCell = fraDays.Controls.Add("Forms.Label.1", "cell_" & Index)
        With lblCell
            .Left = Col * CellW + 2
            .Top = Row * CellH + 2
            .Width = CellW - 2
            .Height = CellH - 2
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
        End With

        Set lblText = fraDays.Controls.Add("Forms.Label.1", "text_" & Index)
        With lblText
            .Left = lblCell.Left
            .Top = lblCell.Top + ((lblCell.Height - TextH) / 2) + 2
            .Width = lblCell.Width
            .Height = TextH
            .TextAlign = fmTextAlignCenter
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
        End With

        DayNum = Index - StartWeekday + 1

        Set CalObj = New clsCalDay
        Set CalObj.CellLabel = lblCell
        Set CalObj.TextLabel = lblText
        Set CalObj.ParentCal = Me

        If DayNum >= 1 And DayNum <= DaysInMonth Then
            lblText.Caption = DayNum
            CalObj.DayNumber = DayNum
            CalObj.IsCurrentMonth = True
            ThisDayDate = DateSerial(Year(CurrentDate), Month(CurrentDate), DayNum)

            If ThisDayDate = Date Then
                lblCell.BackColor = RGB(184, 138, 90)
                lblText.ForeColor = RGB(253, 250, 245)
                lblText.Font.Bold = True
                CalObj.DefaultBackColor = RGB(184, 138, 90)
                CalObj.DefaultTextColor = RGB(253, 250, 245)
            ElseIf PreSelectedDate <> 0 And ThisDayDate = PreSelectedDate Then
                lblCell.BackColor = RGB(214, 176, 140)
                lblText.ForeColor = RGB(59, 42, 32)
                lblText.Font.Bold = True
                CalObj.DefaultBackColor = RGB(214, 176, 140)
                CalObj.DefaultTextColor = RGB(59, 42, 32)
            Else
                lblCell.BackColor = RGB(253, 250, 245)
                lblText.ForeColor = RGB(59, 42, 32)
                lblText.Font.Bold = False
                CalObj.DefaultBackColor = RGB(253, 250, 245)
                CalObj.DefaultTextColor = RGB(59, 42, 32)
            End If
        Else
            lblText.Caption = ""
            lblCell.BackColor = RGB(253, 250, 245)
            CalObj.IsCurrentMonth = False
            CalObj.DefaultBackColor = RGB(253, 250, 245)
            CalObj.DefaultTextColor = RGB(59, 42, 32)
        End If

        DayButtons.Add CalObj
    Next Index
End Sub

Public Sub ApplyDayHover(ByVal TargetCell As MSForms.Label)
    Dim obj As clsCalDay
    
    If Not LastHovered Is Nothing Then
        If LastHovered.Name = TargetCell.Name Then Exit Sub
        For Each obj In DayButtons
            If obj.CellLabel.Name = LastHovered.Name Then
                LastHovered.BackColor = obj.DefaultBackColor
                obj.TextLabel.ForeColor = obj.DefaultTextColor
                Exit For
            End If
        Next obj
    End If

    TargetCell.BackColor = RGB(232, 220, 195)
    Set LastHovered = TargetCell
End Sub

Public Sub SelectDay(ByVal D As Integer)
    SelectedDate = Format(DateSerial(Year(CurrentDate), Month(CurrentDate), D), "dd/mm/yyyy")
    
    If Not TargetControl Is Nothing Then
        On Error Resume Next
        TargetControl.Text = SelectedDate
        On Error GoTo 0
    End If
    
    Unload Me
End Sub

Private Sub lblPrev_Click()
    CurrentDate = DateAdd("m", -1, CurrentDate)
    RenderCalendar
End Sub

Private Sub lblNext_Click()
    CurrentDate = DateAdd("m", 1, CurrentDate)
    RenderCalendar
End Sub

Private Sub lblToday_Click()
    CurrentDate = Date
    RenderCalendar
End Sub

Private Sub fraDays_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim obj As clsCalDay
    If Not LastHovered Is Nothing Then
        For Each obj In DayButtons
            If obj.CellLabel.Name = LastHovered.Name Then
                LastHovered.BackColor = obj.DefaultBackColor
                obj.TextLabel.ForeColor = obj.DefaultTextColor
                Exit For
            End If
        Next obj
        Set LastHovered = Nothing
    End If
End Sub

