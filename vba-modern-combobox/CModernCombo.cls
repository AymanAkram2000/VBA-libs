VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModernCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_TextBox As MSForms.TextBox
Attribute m_TextBox.VB_VarHelpID = -1
Private WithEvents m_Frame As MSForms.Frame
Attribute m_Frame.VB_VarHelpID = -1
Private WithEvents m_VTrack As MSForms.Label
Attribute m_VTrack.VB_VarHelpID = -1
Private WithEvents m_VThumb As MSForms.Label
Attribute m_VThumb.VB_VarHelpID = -1
Private m_ParentForm As Object

Private m_SourceData As Variant
Private m_StartExcelRow As Long
Private m_ColCount As Long
Private m_DisplayCol As Long
Private m_BoundCol As Long
Private m_MaxVisibleRows As Long
Private m_RowHeight As Double

' Offsets: X, Y, Width, Row Height
Private m_OffsetX As Double
Private m_OffsetY As Double
Private m_OffsetW As Double
Private m_OffsetH As Double

' Custom Theme Colors
Private m_ClrBg As Long
Private m_ClrHover As Long
Private m_ClrText As Long
Private m_ClrBorder As Long
Private m_ClrTrack As Long
Private m_ClrThumb As Long
Private m_ClrThumbHover As Long

Private m_ItemsColl As Collection
Private m_Watchers As Collection
Private m_CurrentHighlighted As CModernItem
Private m_SelectedValue As Variant
Private m_SelectedRow As Long
Private m_LastValidText As String
Private m_IsUpdating As Boolean
Private m_MatchRequired As Boolean

' Scrollbar & Navigation Variables
Private m_BarThickness As Double
Private m_VBarOnLeft As Boolean
Private m_IsDragging As Boolean
Private m_DragStartY As Single
Private m_DragStartTop As Double
Private m_HighlightedIndex As Long

' Quick Positional Setup with Geometry Offsets & Theme Colors:
' Attach(TargetBox, Source, [DisplayCol], [BoundCol], [MaxRows], [MatchReq], [Offsets], [Theme])
Public Sub Attach(ByVal TargetBox As MSForms.TextBox, _
                  ByVal Source As Variant, _
                  Optional ByVal DisplayCol As Long = 1, _
                  Optional ByVal BoundCol As Long = 1, _
                  Optional ByVal MaxRows As Long = 5, _
                  Optional ByVal MatchReq As Boolean = False, _
                  Optional ByVal Offsets As String = "0,0,0,0", _
                  Optional ByVal Theme As String = "")
    
    Set m_TextBox = TargetBox
    Set m_ParentForm = TargetBox.Parent
    m_DisplayCol = DisplayCol
    m_BoundCol = BoundCol
    m_MaxVisibleRows = MaxRows
    m_MatchRequired = MatchReq
    m_BarThickness = 7
    m_HighlightedIndex = 0
    
    m_SelectedRow = 0
    m_SelectedValue = Empty
    m_LastValidText = ""
    
    ParseOffsets Offsets
    ParseTheme Theme
    
    ' Row Height dynamically adjusted with offset H (min 12px)
    m_RowHeight = Application.Max(12, m_TextBox.Height + m_OffsetH)
    
    If m_TextBox.TextAlign = fmTextAlignRight Then
        m_VBarOnLeft = True
    Else
        m_VBarOnLeft = False
    End If
    
    Set m_ItemsColl = New Collection
    Set m_Watchers = New Collection
    
    LoadSourceData Source
    CreateOverlayFrame
    SetupBackdropCatcher
    HookAllOtherControls
    
    With m_TextBox
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = m_ClrBorder
    End With
End Sub

Public Property Get Value() As Variant
    ValidateMatch
    Value = m_SelectedValue
End Property

Public Property Get Row() As Long
    ValidateMatch
    Row = m_SelectedRow
End Property

Public Property Get MatchRequired() As Boolean
    MatchRequired = m_MatchRequired
End Property

Public Property Let MatchRequired(ByVal val As Boolean)
    m_MatchRequired = val
End Property

Public Sub CloseDropdown()
    If Not m_Frame Is Nothing Then
        If m_Frame.Visible Then
            m_Frame.Visible = False
            m_HighlightedIndex = 0
            ValidateMatch
        End If
    End If
End Sub

Public Sub ValidateMatch()
    If Not m_MatchRequired Then Exit Sub
    If m_IsUpdating Then Exit Sub
    
    Dim txt As String
    txt = Trim(m_TextBox.Text)
    
    If txt = "" Then
        m_SelectedValue = Empty
        m_SelectedRow = 0
        m_LastValidText = ""
        Exit Sub
    End If
    
    Dim r As Long
    For r = 1 To UBound(m_SourceData, 1)
        If StrComp(Trim(CStr(m_SourceData(r, m_DisplayCol))), txt, vbTextCompare) = 0 Then
            m_SelectedValue = m_SourceData(r, m_BoundCol)
            If m_StartExcelRow > 0 Then
                m_SelectedRow = m_StartExcelRow + r - 1
            Else
                m_SelectedRow = r
            End If
            m_LastValidText = txt
            Exit Sub
        End If
    Next r
    
    m_IsUpdating = True
    m_TextBox.Text = m_LastValidText
    If m_LastValidText = "" Then
        m_SelectedValue = Empty
        m_SelectedRow = 0
    End If
    m_IsUpdating = False
End Sub

Private Sub ParseOffsets(ByVal rawStr As String)
    m_OffsetX = 0: m_OffsetY = 0: m_OffsetW = 0: m_OffsetH = 0
    rawStr = Trim(rawStr)
    If rawStr = "" Then Exit Sub
    
    Dim parts() As String
    parts = Split(rawStr, ",")
    
    On Error Resume Next
    If UBound(parts) >= 0 Then m_OffsetX = val(Trim(parts(0)))
    If UBound(parts) >= 1 Then m_OffsetY = val(Trim(parts(1)))
    If UBound(parts) >= 2 Then m_OffsetW = val(Trim(parts(2)))
    If UBound(parts) >= 3 Then m_OffsetH = val(Trim(parts(3)))
    On Error GoTo 0
End Sub

Private Sub ParseTheme(ByVal rawStr As String)
    ' Default Clean Modern Theme
    m_ClrBg = RGB(255, 255, 255)
    m_ClrHover = RGB(245, 240, 235)
    m_ClrText = RGB(32, 32, 32)
    m_ClrBorder = RGB(224, 216, 208)
    m_ClrTrack = RGB(240, 236, 232)
    m_ClrThumb = RGB(184, 176, 166)
    m_ClrThumbHover = RGB(142, 133, 123)
    
    rawStr = Trim(rawStr)
    If rawStr = "" Then Exit Sub
    
    Dim tokens() As String
    If InStr(rawStr, ";") > 0 Then
        tokens = Split(rawStr, ";")
    Else
        tokens = Split(rawStr, "|")
    End If
    
    On Error Resume Next
    If UBound(tokens) >= 0 Then m_ClrBg = ConvertToColor(tokens(0), m_ClrBg)
    If UBound(tokens) >= 1 Then m_ClrHover = ConvertToColor(tokens(1), m_ClrHover)
    If UBound(tokens) >= 2 Then m_ClrText = ConvertToColor(tokens(2), m_ClrText)
    If UBound(tokens) >= 3 Then m_ClrBorder = ConvertToColor(tokens(3), m_ClrBorder)
    If UBound(tokens) >= 4 Then m_ClrTrack = ConvertToColor(tokens(4), m_ClrTrack)
    If UBound(tokens) >= 5 Then
        m_ClrThumb = ConvertToColor(tokens(5), m_ClrThumb)
        m_ClrThumbHover = m_ClrThumb
    End If
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

Private Sub LoadSourceData(ByVal Source As Variant)
    Dim raw As Variant, r As Long, c As Long
    
    m_StartExcelRow = 0
    If TypeName(Source) = "Range" Then
        raw = Source.Value
        m_StartExcelRow = Source.Row
    Else
        raw = Source
    End If
    
    If Not IsArray(raw) Then Exit Sub
    
    Dim rMin As Long, rMax As Long, cMin As Long, cMax As Long
    rMin = LBound(raw, 1): rMax = UBound(raw, 1)
    
    On Error Resume Next
    cMin = LBound(raw, 2): cMax = UBound(raw, 2)
    If Err.Number <> 0 Then
        m_ColCount = 1
        ReDim m_SourceData(1 To (rMax - rMin + 1), 1 To 1)
        For r = rMin To rMax
            m_SourceData(r - rMin + 1, 1) = raw(r)
        Next r
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    m_ColCount = cMax - cMin + 1
    ReDim m_SourceData(1 To (rMax - rMin + 1), 1 To m_ColCount)
    For r = rMin To rMax
        For c = cMin To cMax
            m_SourceData(r - rMin + 1, c - cMin + 1) = raw(r, c)
        Next c
    Next r
End Sub

Private Sub CreateOverlayFrame()
    Dim frameName As String
    frameName = "fraDrop_" & m_TextBox.Name
    
    On Error Resume Next
    m_ParentForm.Controls.Remove frameName
    On Error GoTo 0
    
    Set m_Frame = m_ParentForm.Controls.Add("Forms.Frame.1", frameName, False)
    
    With m_Frame
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = m_ClrBorder
        .BackColor = m_ClrBg
        .Caption = ""
        .ScrollBars = fmScrollBarsNone
        .KeepScrollBarsVisible = fmScrollBarsNone
        .Visible = False
        .ZOrder 0
    End With
    
    PositionFrame
End Sub

Private Sub SetupBackdropCatcher()
    Dim bgName As String
    Dim bgLbl As MSForms.Label
    Dim watcher As CGlobalWatcher
    
    bgName = "lblBackdrop_GlobalCatch"
    
    On Error Resume Next
    Set bgLbl = m_ParentForm.Controls(bgName)
    On Error GoTo 0
    
    If bgLbl Is Nothing Then
        Set bgLbl = m_ParentForm.Controls.Add("Forms.Label.1", bgName, True)
        With bgLbl
            .Left = 0
            .Top = 0
            .Width = m_ParentForm.InsideWidth
            .Height = m_ParentForm.InsideHeight
            .BackStyle = fmBackStyleTransparent
            .Caption = ""
            .ZOrder 1
        End With
    End If
    
    Set watcher = New CGlobalWatcher
    Set watcher.WatchLabel = bgLbl
    Set watcher.ParentCombo = Me
    m_Watchers.Add watcher
End Sub

Private Sub HookAllOtherControls()
    Dim ctrl As MSForms.Control
    Dim watcher As CGlobalWatcher
    
    For Each ctrl In m_ParentForm.Controls
        If ctrl.Name <> m_TextBox.Name And ctrl.Name <> m_Frame.Name And _
           Left(ctrl.Name, 7) <> "lblBg_" And Left(ctrl.Name, 8) <> "lblTxt_" And _
           Left(ctrl.Name, 9) <> "lblTrack_" And Left(ctrl.Name, 9) <> "lblThumb_" Then
            
            Set watcher = New CGlobalWatcher
            Set watcher.ParentCombo = Me
            
            If TypeOf ctrl Is MSForms.TextBox Then
                Set watcher.WatchTextBox = ctrl
                m_Watchers.Add watcher
            ElseIf TypeOf ctrl Is MSForms.CommandButton Then
                Set watcher.WatchButton = ctrl
                m_Watchers.Add watcher
            ElseIf TypeOf ctrl Is MSForms.Frame Then
                Set watcher.WatchFrame = ctrl
                m_Watchers.Add watcher
            ElseIf TypeOf ctrl Is MSForms.Label Then
                If ctrl.Name <> "lblBackdrop_GlobalCatch" Then
                    Set watcher.WatchLabel = ctrl
                    m_Watchers.Add watcher
                End If
            End If
        End If
    Next ctrl
End Sub

Private Sub PositionFrame()
    If m_Frame Is Nothing Then Exit Sub
    With m_Frame
        .Left = m_TextBox.Left + m_OffsetX
        .Top = m_TextBox.Top + m_TextBox.Height + 1 + m_OffsetY
        .Width = Application.Max(20, m_TextBox.Width + m_OffsetW)
    End With
End Sub

Private Sub FilterAndRender(ByVal SearchText As String)
    Dim r As Long, count As Long
    Dim cellText As String
    Dim itemObj As CModernItem
    Dim lblBg As MSForms.Label
    Dim lblText As MSForms.Label
    Dim itemTop As Double
    Dim actualTextHeight As Double
    Dim padding As Double
    Dim isMatch As Boolean
    
    ClearItems
    m_HighlightedIndex = 0
    
    itemTop = 0
    count = 0
    padding = 6
    
    For r = 1 To UBound(m_SourceData, 1)
        cellText = Trim(CStr(m_SourceData(r, m_DisplayCol)))
        
        If cellText <> "" Then
            If SearchText = "" Then
                isMatch = True
            Else
                isMatch = (InStr(1, cellText, SearchText, vbTextCompare) > 0)
            End If
            
            If isMatch Then
                count = count + 1
                
                Set lblBg = m_Frame.Controls.Add("Forms.Label.1", "lblBg_" & count, True)
                With lblBg
                    .Left = 0
                    .Top = itemTop
                    .Width = m_Frame.InsideWidth
                    .Height = m_RowHeight
                    .Caption = ""
                    .BackColor = m_ClrBg
                    .BackStyle = fmBackStyleOpaque
                    .BorderStyle = fmBorderStyleNone
                End With
                
                Set lblText = m_Frame.Controls.Add("Forms.Label.1", "lblTxt_" & count, True)
                With lblText
                    .Caption = cellText
                    .Font.Name = m_TextBox.Font.Name
                    .Font.Size = m_TextBox.Font.Size
                    .Font.Bold = m_TextBox.Font.Bold
                    .ForeColor = m_ClrText
                    .BackStyle = fmBackStyleTransparent
                    .BorderStyle = fmBorderStyleNone
                    .TextAlign = m_TextBox.TextAlign
                    
                    ' Measure actual text height for vertical centering
                    .AutoSize = True
                    actualTextHeight = .Height
                    .AutoSize = False
                    
                    ' Perfect vertical center inside m_RowHeight
                    .Height = actualTextHeight
                    .Top = itemTop + Application.Max(0, (m_RowHeight - actualTextHeight) / 2)
                    
                    If m_VBarOnLeft Then
                        .Left = padding + m_BarThickness + 2
                    Else
                        .Left = padding
                    End If
                    .Width = Application.Max(10, m_Frame.InsideWidth - (padding * 2) - m_BarThickness - 2)
                    .ZOrder 0
                End With
                
                Set itemObj = New CModernItem
                Set itemObj.BgLabel = lblBg
                Set itemObj.TextLabel = lblText
                Set itemObj.ParentCombo = Me
                itemObj.RowIndex = count
                itemObj.DisplayText = cellText
                itemObj.BoundVal = m_SourceData(r, m_BoundCol)
                
                If m_StartExcelRow > 0 Then
                    itemObj.SourceRowIndex = m_StartExcelRow + r - 1
                Else
                    itemObj.SourceRowIndex = r
                End If
                
                m_ItemsColl.Add itemObj
                itemTop = itemTop + m_RowHeight
            End If
        End If
    Next r
    
    If count > 0 Then
        PositionFrame
        
        Dim visibleRows As Long
        visibleRows = IIf(count > m_MaxVisibleRows, m_MaxVisibleRows, count)
        
        m_Frame.Height = (visibleRows * m_RowHeight) + 1.5
        m_Frame.ScrollHeight = itemTop
        m_Frame.ScrollTop = 0
        
        BuildInternalScrollBar count, visibleRows
        
        m_Frame.Visible = True
        m_Frame.ZOrder 0
    Else
        m_Frame.Visible = False
    End If
End Sub

Private Sub BuildInternalScrollBar(ByVal TotalItems As Long, ByVal VisibleItems As Long)
    If TotalItems > VisibleItems Then
        Dim trackX As Double
        If m_VBarOnLeft Then
            trackX = 1
        Else
            trackX = m_Frame.InsideWidth - m_BarThickness - 1
        End If
        
        Set m_VTrack = m_Frame.Controls.Add("Forms.Label.1", "lblTrack_Internal", True)
        With m_VTrack
            .Left = trackX
            .Top = 1
            .Width = m_BarThickness
            .Height = m_Frame.InsideHeight - 2
            .BackColor = m_ClrTrack
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
            .Caption = ""
            .Visible = True
            .ZOrder 0
        End With
        
        Dim thumbH As Double
        thumbH = (VisibleItems / TotalItems) * m_VTrack.Height
        If thumbH < 18 Then thumbH = 18
        
        Set m_VThumb = m_Frame.Controls.Add("Forms.Label.1", "lblThumb_Internal", True)
        With m_VThumb
            .Left = trackX
            .Top = 1
            .Width = m_BarThickness
            .Height = thumbH
            .BackColor = m_ClrThumb
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
            .Caption = ""
            .Visible = True
            .ZOrder 0
        End With
    End If
End Sub

Private Sub ClearItems()
    Dim i As Long
    Set m_CurrentHighlighted = Nothing
    Set m_VTrack = Nothing
    Set m_VThumb = Nothing
    Set m_ItemsColl = New Collection
    
    For i = m_Frame.Controls.count - 1 To 0 Step -1
        m_Frame.Controls.Remove m_Frame.Controls(i).Name
    Next i
End Sub

Public Sub HighlightItem(ByVal targetItem As CModernItem)
    If Not m_CurrentHighlighted Is Nothing Then
        If m_CurrentHighlighted.RowIndex = targetItem.RowIndex Then Exit Sub
        m_CurrentHighlighted.BgLabel.BackColor = m_ClrBg
    End If
    
    targetItem.BgLabel.BackColor = m_ClrHover
    Set m_CurrentHighlighted = targetItem
    m_HighlightedIndex = targetItem.RowIndex
End Sub

Public Sub SelectItem(ByVal targetItem As CModernItem)
    m_IsUpdating = True
    m_TextBox.Text = targetItem.DisplayText
    m_SelectedValue = targetItem.BoundVal
    m_SelectedRow = targetItem.SourceRowIndex
    m_LastValidText = targetItem.DisplayText
    CloseDropdown
    m_IsUpdating = False
    m_TextBox.SetFocus
End Sub

Private Sub HighlightByIndex(ByVal index As Long)
    If index < 1 Or index > m_ItemsColl.count Then Exit Sub
    Dim targetItem As CModernItem
    Set targetItem = m_ItemsColl(index)
    HighlightItem targetItem
    ScrollIntoView targetItem
End Sub

Private Sub ScrollIntoView(ByVal targetItem As CModernItem)
    Dim itemTop As Double, itemBottom As Double
    Dim viewTop As Double, viewBottom As Double
    
    itemTop = (targetItem.RowIndex - 1) * m_RowHeight
    itemBottom = itemTop + m_RowHeight
    
    viewTop = m_Frame.ScrollTop
    viewBottom = viewTop + m_Frame.InsideHeight
    
    If itemTop < viewTop Then
        m_Frame.ScrollTop = itemTop
    ElseIf itemBottom > viewBottom Then
        m_Frame.ScrollTop = itemBottom - m_Frame.InsideHeight
    End If
    
    SyncScrollBar
End Sub

Private Sub SyncScrollBar()
    If m_VTrack Is Nothing Or m_VThumb Is Nothing Then Exit Sub
    
    m_VTrack.Top = m_Frame.ScrollTop + 1
    
    Dim maxScroll As Double
    maxScroll = m_Frame.ScrollHeight - m_Frame.InsideHeight
    
    Dim travelDist As Double
    travelDist = m_VTrack.Height - m_VThumb.Height
    
    If maxScroll > 0 And travelDist > 0 Then
        m_VThumb.Top = m_VTrack.Top + ((m_Frame.ScrollTop / maxScroll) * travelDist)
    Else
        m_VThumb.Top = m_VTrack.Top
    End If
    
    m_VTrack.ZOrder 0
    m_VThumb.ZOrder 0
End Sub

' === Scrollbar Events ===

Private Sub m_VThumb_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        m_IsDragging = True
        m_DragStartY = Y
        m_DragStartTop = m_VThumb.Top - m_Frame.ScrollTop
        m_VThumb.BackColor = m_ClrThumbHover
    End If
End Sub

Private Sub m_VThumb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If m_IsDragging Then
        Dim newRelTop As Double
        Dim maxRelTop As Double
        Dim travelDist As Double
        
        newRelTop = (m_VThumb.Top - m_Frame.ScrollTop) + (Y - m_DragStartY)
        If newRelTop < 1 Then newRelTop = 1
        
        maxRelTop = 1 + m_VTrack.Height - m_VThumb.Height
        If newRelTop > maxRelTop Then newRelTop = maxRelTop
        
        travelDist = maxRelTop - 1
        If travelDist > 0 Then
            Dim ratio As Double
            ratio = (newRelTop - 1) / travelDist
            m_Frame.ScrollTop = ratio * (m_Frame.ScrollHeight - m_Frame.InsideHeight)
            SyncScrollBar
        End If
    End If
End Sub

Private Sub m_VThumb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    m_IsDragging = False
    m_VThumb.BackColor = m_ClrThumb
End Sub

Private Sub m_VTrack_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        Dim maxScroll As Double
        maxScroll = m_Frame.ScrollHeight - m_Frame.InsideHeight
        
        If Y < (m_VThumb.Top - m_VTrack.Top) Then
            m_Frame.ScrollTop = Application.Max(0, m_Frame.ScrollTop - m_Frame.InsideHeight)
        Else
            m_Frame.ScrollTop = Application.Min(maxScroll, m_Frame.ScrollTop + m_Frame.InsideHeight)
        End If
        SyncScrollBar
    End If
End Sub

' === TextBox Events ===

Private Sub m_TextBox_Change()
    If m_IsUpdating Then Exit Sub
    FilterAndRender m_TextBox.Text
End Sub

Private Sub m_TextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not m_Frame.Visible Then
        FilterAndRender m_TextBox.Text
    End If
End Sub

Private Sub m_TextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        CloseDropdown
        KeyCode = 0
        Exit Sub
    End If
    
    If Not m_Frame.Visible Then
        If KeyCode = vbKeyDown Then
            FilterAndRender m_TextBox.Text
            KeyCode = 0
        End If
        Exit Sub
    End If
    
    Select Case KeyCode
        Case vbKeyDown
            If m_ItemsColl.count > 0 Then
                If m_HighlightedIndex < m_ItemsColl.count Then
                    m_HighlightedIndex = m_HighlightedIndex + 1
                    HighlightByIndex m_HighlightedIndex
                End If
            End If
            KeyCode = 0
            
        Case vbKeyUp
            If m_HighlightedIndex > 1 Then
                m_HighlightedIndex = m_HighlightedIndex - 1
                HighlightByIndex m_HighlightedIndex
            ElseIf m_HighlightedIndex = 1 Then
                If Not m_CurrentHighlighted Is Nothing Then
                    m_CurrentHighlighted.BgLabel.BackColor = m_ClrBg
                    Set m_CurrentHighlighted = Nothing
                End If
                m_HighlightedIndex = 0
                m_Frame.ScrollTop = 0
                SyncScrollBar
                m_TextBox.SelStart = Len(m_TextBox.Text)
            End If
            KeyCode = 0
            
        Case vbKeyReturn, vbKeySpace
            If m_HighlightedIndex > 0 And m_HighlightedIndex <= m_ItemsColl.count Then
                SelectItem m_ItemsColl(m_HighlightedIndex)
                KeyCode = 0
            ElseIf KeyCode = vbKeyReturn Then
                CloseDropdown
                KeyCode = 0
            End If
            
        Case vbKeyTab
            If m_HighlightedIndex > 0 And m_HighlightedIndex <= m_ItemsColl.count Then
                SelectItem m_ItemsColl(m_HighlightedIndex)
            Else
                CloseDropdown
            End If
    End Select
End Sub

