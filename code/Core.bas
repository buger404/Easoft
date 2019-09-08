Attribute VB_Name = "Core"
Public StrFormat(2) As Long, EditStrFormat(2) As Long
Public NowIce As Icelolly, Displaying As Boolean, AnimateValue As Single
Public GobalAnimation As New Animation, DebugB As New ColorMix
Public StartEasoft As Boolean
Public BlurEffect As Long
Public BlurParam As BlurParams
Public NowLayout As Layout
Sub IceError(ByVal Number As Long, ByVal Place As String, ByVal ErrorType As String, ByVal Des As String, Ice As Icelolly)
    If InStr(Ice.ErrorInfo, ErrorType & " in " & Place & " : " & vbCrLf & Des & " (0x" & Hex(Number) & ")") > 0 Then Exit Sub
    Ice.ErrorInfo = Ice.ErrorInfo & ErrorType & " in " & Place & " : " & vbCrLf & Des & " (0x" & Hex(Number) & ")" & vbCrLf
End Sub
Function ColorLight(ByVal R As Long, ByVal G As Long, ByVal b As Long) As Long
Dim Result As Long
If R > Result Then Result = R
If G > Result Then Result = G
If b > Result Then Result = b
ColorLight = (Result - R) ^ 2 + (Result - G) ^ 2 + (Result - b) ^ 2
End Function
Public Function HiWord(lValue As Long) As Integer
    If lValue And &H80000000 Then
        HiWord = (lValue \ 65535) - 1
    Else
        HiWord = lValue \ 65535
    End If
End Function
Public Function LoWord(lValue As Long) As Integer
    If lValue And &H8000& Then
        LoWord = &H8000 Or (lValue And &H7FFF&)
    Else
        LoWord = lValue And &HFFFF&
    End If
End Function
Public Sub SetEWindow(ByVal Hwnd As Long, Style As EWindowStyle)
    If Style = LayeredWindow Then
        If (GetWindowLongA(Hwnd, GWL_STYLE) And WS_CAPTION) = WS_CAPTION Then Err.Raise -40410, , "Windows's angry" & vbCrLf & "Set the border style = 0 , then continue ."
        SetWindowLongA Hwnd, GWL_EXSTYLE, GetWindowLongA(Hwnd, -20) Or &H80000
    End If
    
    If Style = AeroWindow Then
        If (GetWindowLongA(Hwnd, GWL_STYLE) And WS_CAPTION) = WS_CAPTION Then Err.Raise -40410, , "Windows's angry" & vbCrLf & "Set the border style = 0 , then continue ."
        SetWindowLongA Hwnd, GWL_EXSTYLE, GetWindowLongA(Hwnd, -20) Or &H80000
        BlurWindow Hwnd
    End If
End Sub
Public Sub SetWindowShadow(frm As Object, Optional ByVal Depth As Long = 7, Optional ByVal Trans As Long = 18)
    Dim Shadow As New WinShadow
    Shadow.Color = RGB(0, 0, 0)
    Shadow.Depth = Depth
    Shadow.Transparency = Trans
    Shadow.Shadow frm
    AddEShadow Shadow
End Sub
Public Function NewPos(ByVal x As Single, ByVal y As Single) As EPosition
    NewPos.x = x: NewPos.y = y
End Function
Public Function NewSize(ByVal Width As Single, ByVal Height As Single) As ESize
    NewSize.Width = Width: NewSize.Height = Height
End Function
Public Function NewAlign(Horizontal As EAlign1, Vertical As EAlign2) As EAlign
    NewAlign.Horizontal = Horizontal: NewAlign.Vertical = Vertical
End Function
Public Function NewAnimate(Obj As Object, Property As String, Delay As Long, Duration As Long, Start As Variant, Target As Variant, Optional Func As String = "linear", Optional ByVal Index As Integer = -1) As EAnimate
    With NewAnimate
        Set .Obj = Obj
        .Start = Start
        .Target = Target
        .ProperName = Property
        .Delay = Delay
        .Duration = Duration
        .FuncName = Func
        .StartTime = GetTickCount
        .Index = Index
    End With
    If TypeName(Obj) = "ColorMix" Then AddRefColor Obj
    If TypeName(Obj) = "Icelolly" Then AddRefIce Obj
    If TypeName(Obj) = "Fonter" Then AddRefFont Obj
    If TypeName(Obj) = "StyleBox" Then AddRefStyle Obj
    If TypeName(Obj) = "EImage" Then AddRefImage Obj
End Function
Public Function IsInRect(ByVal PointX As Long, ByVal PointY As Long, Pos As EPosition, Size As ESize) As Boolean
    IsInRect = (PointX >= Pos.x And PointY >= Pos.y And PointX <= Pos.x + Size.Width And PointY <= Pos.y + Size.Height)
End Function
