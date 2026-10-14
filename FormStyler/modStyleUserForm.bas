Attribute VB_Name = "modStyleUserForm"
#If VBA7 Then
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

Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_THICKFRAME As Long = &H40000
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Public Sub StyleUserForm(frm As Object, Optional ByVal Radius As Long = 20)
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
    
    ' ≈ÌÃ«œ „ﬁ»÷ «·‰«›–… (HWND) ··›Ê—„
    hwnd = FindWindow("ThunderDFrame", frm.Caption)
    If hwnd = 0 Then Exit Sub
    
    ' ≈“«·… «·‹ Title Bar Ê«·≈ÿ«— «·Œ«—ÃÌ
    currentStyle = GetWindowLong(hwnd, GWL_STYLE)
    currentStyle = currentStyle And Not WS_CAPTION
    currentStyle = currentStyle And Not WS_THICKFRAME
    SetWindowLong hwnd, GWL_STYLE, currentStyle
    
    '  ÕœÌÀ „ŸÂ— «·‰«›–…
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
    '  ÕÊÌ· √»⁄«œ «·›Ê—„ „‰ ‰ﬁ«ÿ (Points) ≈·Ï »ﬂ”· (Pixels)
    widthPx = CLng(frm.Width * 1.3333)
    heightPx = CLng(frm.Height * 1.3333)
    
    '  ÿ»Ìﬁ «· œÊÌ— «·Œ›Ì› ⁄·Ï «·ÕÊ«›
    hRgn = CreateRoundRectRgn(0, 0, widthPx, heightPx, Radius, Radius)
    SetWindowRgn hwnd, hRgn, True
End Sub

