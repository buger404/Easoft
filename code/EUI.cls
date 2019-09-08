VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Layouts() As Layout, Graphics As Long, Draw As New Compatibles, TargetDC As Long, TarHwnd As Long
Dim OrPos() As EPosition
Dim MouseDrawing As New Compatibles, LastMouseTime As Long, MouseState As Integer
Dim MouseDraw As New Drawing, MGraphics As Long, MousePos As EPosition
Dim NeedPaint As Boolean
Dim ResultDraw As New Compatibles, ResultGraphics As Long
Public BackColor As New ColorMix, MouseColor As New ColorMix
Public LayoutClickActive As Layout
Dim mFocusIcelolly As Icelolly
Dim mLastLayout As Layout
Dim mReturnDown As Boolean
Dim SrcPoint  As POINTAPI
Dim WinSz As Size, BlendFunc As BLENDFUNCTION
Dim LastChr As Long
Public Property Get FocusIcelolly() As Icelolly
    Set FocusIcelolly = mFocusIcelolly
End Property
Public Property Set FocusIcelolly(nIcelolly As Icelolly)
    On Error Resume Next
    Set mFocusIcelolly = nIcelolly
    Set NowLayout = nIcelolly.ParentLayout
End Property
Sub Create(ByVal Hwnd As Long)
    Dim hdc As Long
    BackColor.Color = argb(255, 255, 255, 255)
    hdc = GetDC(Hwnd)
    Draw.Create hdc, Screen.Width / 15, Screen.Height / 15
    GdipCreateFromHDC Draw.CompatibleDC, Graphics
    TargetDC = hdc
    TarHwnd = Hwnd
    ReleaseDC Hwnd, hdc
    GdipGraphicsClear Graphics, BackColor.Color
    NewWndProc Hwnd, Me, "WinMessage"
    
    WinSz.cx = Screen.Width / 15
    WinSz.cy = Screen.Height / 15
    SrcPoint.x = 0
    SrcPoint.y = 0
    With BlendFunc
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    MouseDrawing.Create TargetDC, 60, 60
    MouseColor.Color = argb(255, 25, 169, 255)
    MouseColor.Width = 2
    MouseDraw.Size = NewSize(60, 60)
    GdipCreateFromHDC MouseDrawing.CompatibleDC, MGraphics
    
    ResultDraw.Create TargetDC, Screen.Width / 15, Screen.Height / 15
    GdipCreateFromHDC ResultDraw.CompatibleDC, ResultGraphics
    
    MouseColor.Color = argb(255, 255 - BackColor.R, 255 - BackColor.G, 255 - BackColor.b)
End Sub
Function LayoutByID(ByVal Id As String) As Layout
    For i = 1 To UBound(Layouts)
        If Layouts(i).Id = Id Then Set LayoutByID = Layouts(i): Exit For
    Next
End Function
Sub Refresh()
    GdipResetClip Graphics
    GdipGraphicsClear Graphics, BackColor.Color
    MouseColor.Color = argb(255, 255 - BackColor.R, 255 - BackColor.G, 255 - BackColor.b)
    NeedPaint = True
End Sub
Sub CreateLayout(ByVal FrameID As String, Pos As EPosition, Size As ESize, EventBox As Object)
    On Local Error GoTo sth
    ReDim Preserve Layouts(UBound(Layouts) + 1)
    ReDim Preserve OrPos(UBound(OrPos) + 1)
    Set Layouts(UBound(Layouts)) = New Layout
    Layouts(UBound(Layouts)).Size = Size
    Layouts(UBound(Layouts)).Pos = Pos
    Layouts(UBound(Layouts)).Id = FrameID
    Set Layouts(UBound(Layouts)).ParentUI = Me
    OrPos(UBound(OrPos)) = Pos
    Set Layouts(UBound(Layouts)).EventBox = EventBox
    
sth:
    If Err.Number <> 0 Then Err.Raise 4043, , "Creating Error" & vbCrLf & "Failed to create a layout , that's all we know ."
End Sub
Sub MakeIcelolly(ByVal GunID As String, ByVal Id As String, ByVal FrameID As String, Pos As EPosition, Size As ESize, Style As StyleBox)
    On Local Error GoTo sth
    Dim MyLayout As Layout
    For i = 0 To UBound(Layouts)
        Set MyLayout = Layouts(i)
        If Not MyLayout Is Nothing Then
            If MyLayout.Id = FrameID Then MyLayout.MakeIcelolly GunID, Id, Pos, Size, Style
        End If
    Next
    
sth:
    If Err.Number <> 0 Then Err.Raise 4044, , "Icelolly refused" & vbCrLf & "Icelolly refused your ask , that's all we know ."
End Sub
Function WinMessage_WndProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
    Dim Mousing As Boolean, MouseButton As EClickButton, CState As EClickState
    If uMsg = WM_MOUSEMOVE Then Mousing = True: MouseButton = NoButton: CState = MouseMove
    If uMsg = WM_LBUTTONDOWN Then Mousing = True: MouseButton = LButton: CState = MouseDown
    If uMsg = WM_LBUTTONUP Then Mousing = True: MouseButton = LButton: CState = MouseUp
    If uMsg = WM_RBUTTONDOWN Then Mousing = True: MouseButton = RButton: CState = MouseDown
    If uMsg = WM_RBUTTONUP Then Mousing = True: MouseButton = RButton: CState = MouseUp
    If uMsg = WM_MBUTTONDOWN Then Mousing = True: MouseButton = MButton: CState = MouseDown
    If uMsg = WM_MBUTTONUP Then Mousing = True: MouseButton = MButton: CState = MouseUp
    If uMsg = WM_MOUSELEAVE Then
        Set LayoutClickActive = Nothing
    End If
    If uMsg = WM_MOUSEWHEEL Then
        If Not mLastLayout Is Nothing Then mLastLayout.OnMouseWheel wParam / 7864320
    End If
    
    If uMsg = WM_SIZE Then
        Draw.Dispose
        GdipDeleteGraphics Graphics
        Draw.Create TargetDC, LoWord(lParam), HiWord(lParam)
        GdipCreateFromHDC Draw.CompatibleDC, Graphics
        ResultDraw.Dispose
        GdipDeleteGraphics ResultGraphics
        ResultDraw.Create TargetDC, LoWord(lParam), HiWord(lParam)
        GdipCreateFromHDC ResultDraw.CompatibleDC, ResultGraphics
        WinSz.cx = LoWord(lParam)
        WinSz.cy = HiWord(lParam)
        Call Me.Refresh
        Call Me.Display
    End If
    
    If uMsg = WM_NCPAINT Then ResultDraw.Present TargetDC, 0, 0
    
    If uMsg = WM_KEYDOWN Then
        If wParam = VK_RETURN Then
            If mReturnDown = False Then
                mReturnDown = True
                LastMouseTime = GetTickCount
                MouseState = 0
                'Call CheckClicks(MousePos, LButton, MouseDown)
            End If
        End If
        If Not FocusIcelolly Is Nothing Then FocusIcelolly.OnKeyDown wParam
        WinMessage_WndProc = True
    End If
    
    If uMsg = WM_KEYUP Then
        If wParam = VK_RETURN Then
            If mReturnDown = True Then
                LastMouseTime = GetTickCount
                MouseState = 1
                'Call CheckClicks(MousePos, LButton, MouseUp)
                mReturnDown = False
            End If
        End If
        If Not FocusIcelolly Is Nothing Then FocusIcelolly.OnKeyUp wParam
        WinMessage_WndProc = True
    End If
    
    If uMsg = WM_CHAR And GetAsyncKeyState(VK_CONTROL) = 0 And wParam <> VK_TAB Then
        Dim Ascii As Long
        If wParam <= 127 Then
            LastChr = 0
        Else
            If LastChr = 0 Then LastChr = wParam: Exit Function
        End If
        Ascii = CLng(LastChr) * 256 + wParam
        LastChr = 0
        If Not FocusIcelolly Is Nothing Then FocusIcelolly.OnInput Ascii
        WinMessage_WndProc = True
    End If
    
    If Mousing = True And mReturnDown = False Then
        If MouseButton <> NoButton Then
            LastMouseTime = GetTickCount
            MouseState = IIf(CState = MouseDown, 0, 1)
            MousePos = NewPos(LoWord(lParam), HiWord(lParam))
        End If
        CheckClicks NewPos(LoWord(lParam), HiWord(lParam)), MouseButton, CState
        WinMessage_WndProc = True
    End If
    
End Function
Public Sub Display()
    If TargetDC = 0 Then Err.Raise 4042, , "UI Error" & vbCrLf & "Display before create ."

    Dim MyLayout As Layout, FirstRect As Boolean, Result As Long
    
    GdipResetClip Graphics
    GdipGraphicsClear ResultGraphics, 0
    
    If NeedPaint = True Then
        GdipGraphicsClear Graphics, BackColor.Color
        For i = 1 To UBound(Layouts)
        Layouts(i).NeedPaint = True
        Next
        NeedPaint = False
    End If
    
    For i = 1 To UBound(Layouts)
        Set MyLayout = Layouts(i)
        If Not MyLayout Is Nothing Then
            If MyLayout.Display = True Then
                GdipSetClipRect Graphics, MyLayout.Pos.x, MyLayout.Pos.y, MyLayout.Size.Width, MyLayout.Size.Height, IIf(FirstRect = False, CombineMode.CombineModeReplace, CombineMode.CombineModeUnion)
                FirstRect = True
            End If
            If OrPos(i).x <> MyLayout.Pos.x Or OrPos(i).y <> MyLayout.Pos.y Then
                GdipSetClipRect Graphics, OrPos(i).x - 1, OrPos(i).y - 1, MyLayout.Size.Width + 2, MyLayout.Size.Height + 2, IIf(FirstRect = False, CombineMode.CombineModeReplace, CombineMode.CombineModeUnion)
                FirstRect = True
                OrPos(i) = MyLayout.Pos
            End If
        End If
    Next
    
    If FirstRect = False Then GoTo SkipDraw
    GdipGraphicsClear Graphics, BackColor.Color
    
    For i = 1 To UBound(Layouts)
        Set MyLayout = Layouts(i)
        If Not MyLayout Is Nothing Then
            If MyLayout.Visible = True Then
                GdipIsVisibleRectI Graphics, MyLayout.Pos.x, MyLayout.Pos.y, MyLayout.Size.Width, MyLayout.Size.Height, Result
                If Result Then
                    GdipDrawImage Graphics, MyLayout.Image, MyLayout.Pos.x, MyLayout.Pos.y
                End If
            End If
        End If
    Next
    
    
SkipDraw:
    Draw.Present ResultDraw.CompatibleDC, 0, 0
    
    If GetTickCount - LastMouseTime <= 300 Then
        Dim Progress As Single, DrawS As Long
        Progress = (GetTickCount - LastMouseTime) / 300
        DrawS = IIf(MouseState = 0, 60 - 60 * Progress, 60 * Progress)
        MouseColor.a = IIf(MouseState = 0, 128 * Progress, 128 - 128 * Progress)
        MouseDraw.Clear
        MouseDraw.FillEllipse NewPos(60 / 2 - DrawS / 2, 60 / 2 - DrawS / 2), NewSize(DrawS, DrawS), MouseColor
        MouseDraw.DrawEllipse NewPos(60 / 2 - DrawS / 2, 60 / 2 - DrawS / 2), NewSize(DrawS, DrawS), MouseColor
        MouseDraw.DrawEllipse NewPos(60 / 2 - DrawS / 2, 60 / 2 - DrawS / 2), NewSize(DrawS, DrawS), MouseColor
        GdipGraphicsClear MGraphics, 0
        GdipDrawImage MGraphics, MouseDraw.Image, 0, 0
        MouseDrawing.Present ResultDraw.CompatibleDC, MousePos.x - 30, MousePos.y - 30
    End If
    
    If (GetWindowLongA(TarHwnd, GWL_EXSTYLE) And WS_EX_LAYERED) = WS_EX_LAYERED Then '分层你吼
        UpdateLayeredWindow TarHwnd, TargetDC, ByVal 0&, WinSz, ResultDraw.CompatibleDC, SrcPoint, 0, BlendFunc, &H2
    Else
        ResultDraw.Present TargetDC, 0, 0
    End If
End Sub
Sub MoveWindow()
    SendMessageA TarHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    ReleaseCapture
End Sub
Private Sub CheckClicks(Pos As EPosition, Button As EClickButton, CState As EClickState)
    Dim i As Integer, Find As Integer
    If Not LayoutClickActive Is Nothing Then
        LayoutClickActive.CheckMouse NewPos(Pos.x - LayoutClickActive.Pos.x, Pos.y - LayoutClickActive.Pos.y), Button, CState
        Exit Sub
    End If
    
    For i = UBound(Layouts) To 1 Step -1
        If IsInRect(Pos.x, Pos.y, Layouts(i).Pos, Layouts(i).Size) = True Then  '找到这个布局了
            Layouts(i).CheckMouse NewPos(Pos.x - Layouts(i).Pos.x, Pos.y - Layouts(i).Pos.y), Button, CState
            Set mLastLayout = Layouts(i)
            Find = i: Exit For
        End If
    Next
    For i = 1 To UBound(Layouts)
        If i <> Find Then Layouts(i).LostMouse
    Next
    
    If Find = 0 Then
        Set mLastLayout = Nothing: MouseColor.Color = argb(255, 255 - BackColor.R, 255 - BackColor.G, 255 - BackColor.b)
        If CState = MouseUp Then Set FocusIcelolly = Nothing
    End If
End Sub
Private Sub Class_Initialize()
    If StartEasoft = False Then Err.Raise 4046, , "Blackmouth's love" & vbCrLf & "Start Easoft first ."
    ReDim Layouts(0)
    ReDim OrPos(0)
End Sub
