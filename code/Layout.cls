VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Layout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Icelollys() As Icelolly, OrPos() As EPosition, mSize As ESize, mPos As EPosition, mVisible As Boolean, VisibleChanged As Boolean
Dim LastClicking As Icelolly
Dim Draw As New Drawing
Public Oning As Boolean
Public Id As String, EventBox As Object, NeedPaint As Boolean
Public BackColor As New ColorMix
Public ParentUI As EUI
Public MouseInIcelolly As Icelolly
Public Property Get x() As Single
    x = mPos.x
End Property
Public Property Let x(nX As Single)
    mPos.x = nX
    NeedPaint = True
End Property
Public Property Get y() As Single
    y = mPos.y
End Property
Public Property Let y(nY As Single)
    mPos.y = nY
    NeedPaint = True
End Property
Public Property Get Width() As Single
    Width = mSize.Width
End Property
Public Property Let Width(nWidth As Single)
    mSize.Width = nWidth
    Draw.Size = mSize
    NeedPaint = True
End Property
Public Property Get Height() As Single
    Height = mSize.Height
End Property
Public Property Let Height(nHeight As Single)
    mSize.Height = nHeight
    Draw.Size = mSize
    NeedPaint = True
End Property
Public Property Get Visible() As Boolean
    Visible = mVisible
End Property
Public Property Let Visible(nVisible As Boolean)
    mVisible = nVisible
    VisibleChanged = True
End Property
Public Property Get Image() As Long
    Image = Draw.Image
End Property
Public Property Get Pos() As EPosition
    Pos = mPos
End Property
Public Property Let Pos(nPos As EPosition)
    mPos = nPos
End Property
Public Property Get Size() As ESize
    Size = mSize
End Property
Public Property Let Size(nSize As ESize)
    mSize = nSize
    Draw.Size = nSize
    NeedPaint = True
End Property
Sub CallEvent(ByVal EventName As String)
    On Local Error Resume Next
    CallByName EventBox, EventName, VbMethod
End Sub
Sub OnMouseWheel(ByVal Direction As Integer)
    If Not LastClicking Is Nothing Then LastClicking.OnMouseWheel Direction
End Sub
Function Display() As Boolean
    Dim RefList As String
    Dim MyIcelolly As Icelolly, FirstRect As Boolean, Result As Long
    
    GdipResetClip Draw.Hwnd
    If VisibleChanged = True Then Display = True: VisibleChanged = False
    
    For i = 1 To UBound(Icelollys)
        Set MyIcelolly = Icelollys(i)
        If Not MyIcelolly Is Nothing Then
            Set NowIce = MyIcelolly
            If IsInRect(MyIcelolly.Pos.x, MyIcelolly.Pos.y, NewPos(0, 0), Size) Then
                If MyIcelolly.Animate.Count > 0 Then MyIcelolly.NeedPaint = True
                If MyIcelolly.DrawAnimate.Count > 0 Then MyIcelolly.NeedPaint = True
                If MyIcelolly.ClickState = MouseLeave Then MyIcelolly.NeedPaint = True
                If OrPos(i).x <> MyIcelolly.Pos.x Or OrPos(i).y <> MyIcelolly.Pos.y Then
                    GdipSetClipRect Draw.Hwnd, OrPos(i).x - 1, OrPos(i).y - 1, MyIcelolly.Size.Width + 2, MyIcelolly.Size.Height + 2, IIf(FirstRect = False, CombineMode.CombineModeReplace, CombineMode.CombineModeUnion)
                    MyIcelolly.NeedPaint = True
                    OrPos(i) = MyIcelolly.Pos
                    FirstRect = True
                End If
                If MyIcelolly.NeedPaint = True Or NeedPaint = True Then
                    GdipSetClipRect Draw.Hwnd, MyIcelolly.Pos.x, MyIcelolly.Pos.y, MyIcelolly.Size.Width, MyIcelolly.Size.Height, IIf(FirstRect = False, CombineMode.CombineModeReplace, CombineMode.CombineModeUnion)
                    FirstRect = True
                    If MyIcelolly.Id <> "" Then CallEvent MyIcelolly.Id & "_Refresh"
                    If Err.Number <> 0 Then If Err.Number <> 0 Then IceError 4039, "Easoft.Icelolly.Events.Refresh", "Warning", "Failed to call .", MyIcelolly: Err.Clear
                End If
            End If
            If MyIcelolly.Id <> "" Then CallEvent MyIcelolly.Id & "_Update"
            If Err.Number <> 0 Then If Err.Number <> 0 Then IceError 4039, "Easoft.Icelolly.Events.Update", "Warning", "Failed to call .", MyIcelolly: Err.Clear
        End If
    Next

    NeedPaint = False
    If FirstRect = False Then Exit Function
    Display = True
    GdipGraphicsClear Draw.Hwnd, BackColor.Color
    
    For i = 1 To UBound(Icelollys)
        Set MyIcelolly = Icelollys(i)
        If Not MyIcelolly Is Nothing Then
            If MyIcelolly.Visible = True Then
                GdipIsVisibleRectI Draw.Hwnd, MyIcelolly.Pos.x, MyIcelolly.Pos.y, MyIcelolly.Size.Width, MyIcelolly.Size.Height, Result
                If Result Then
                    If MyIcelolly.ClickState = MouseUp Then MyIcelolly.ClickState = MouseMove
                    Set NowIce = MyIcelolly
                    MyIcelolly.Display
                    GdipDrawImage Draw.Hwnd, MyIcelolly.Draw.Image, MyIcelolly.Pos.x, MyIcelolly.Pos.y
                    MyIcelolly.NeedPaint = False
                    If MyIcelolly.ClickState = MouseLeave Then MyIcelolly.ClickState = None
                End If
            End If
            'MyIcelolly.ErrorInfo = ""
        End If
    Next
    
End Function
Sub CheckMouse(Pos As EPosition, Button As EClickButton, CState As EClickState)
    Dim Broil_Icelolly As Icelolly, FindMyTrueLove As Boolean, Color As Long
    If Not LastClicking Is Nothing Then
        If CState <> MouseUp And LastClicking.ClickState = MouseDown Then
            Set Broil_Icelolly = LastClicking
            Broil_Icelolly.ClickButton = Button: Broil_Icelolly.ClickPos = NewPos(Pos.x - Broil_Icelolly.x, Pos.y - Broil_Icelolly.y)
            Broil_Icelolly.NeedPaint = True
            Set NowIce = Broil_Icelolly: Broil_Icelolly.Display
            Exit Sub
        End If
        If CState = MouseUp And LastClicking.ClickState = MouseDown Then LastClicking.ClickState = CState: Set ParentUI.LayoutClickActive = Nothing: Set ParentUI.FocusIcelolly = LastClicking
    End If
    
    For i = UBound(Icelollys) To 1 Step -1
        Set Broil_Icelolly = Icelollys(i)
        If Not Broil_Icelolly Is Nothing Then 'If Icelolly is broil and delicious .
            If Broil_Icelolly.BlockClick And IsInRect(Broil_Icelolly.Pos.x, Broil_Icelolly.Pos.y, NewPos(0, 0), Size) Then '* Icelolly blocks the wey !
                
                'FindLove
                If FindMyTrueLove = False Then 'Where is my cute lcelolly ?
                    If IsInRect(Pos.x, Pos.y, Broil_Icelolly.Pos, Broil_Icelolly.Size) Then
                        If (Not LastClicking Is Nothing) And (Not LastClicking Is Broil_Icelolly) Then
                            LastClicking.ClickState = MouseLeave: LastClicking.ClickButton = NoButton: LastClicking.NeedPaint = True
                            Set NowIce = LastClicking
                            LastClicking.Display
                            Set LastClicking = Nothing
                        End If
                        If Broil_Icelolly.ClickState <> MouseDown Then
                            Broil_Icelolly.ClickState = IIf(Broil_Icelolly.ClickState = None, EClickState.MouseEnter, CState)
                            Color = IIf(Broil_Icelolly.CustomDraw = True, Broil_Icelolly.Style.Color(0).Color, GetMouseColor(Broil_Icelolly.Class, Broil_Icelolly))
                            If ParentUI.MouseColor.Color <> Color Then ParentUI.MouseColor.Color = Color
                        End If
                        If CState = MouseDown Then Set ParentUI.LayoutClickActive = Me
                        Set MouseInIcelolly = Broil_Icelolly
                        Broil_Icelolly.ClickButton = Button: Broil_Icelolly.ClickPos = NewPos(Pos.x - Broil_Icelolly.x, Pos.y - Broil_Icelolly.y)
                        Broil_Icelolly.NeedPaint = True
                        Set NowIce = Broil_Icelolly: Broil_Icelolly.Display
                        Set LastClicking = Broil_Icelolly
                        FindMyTrueLove = True ' Oh Yeah ~
                        Exit For
                    End If
                End If
                
            End If
        End If
    Next
    
    'All of icelolly refuse your love
    If FindMyTrueLove = False Then
        If (Not LastClicking Is Nothing) Then
            LastClicking.ClickState = MouseLeave: LastClicking.ClickButton = NoButton: LastClicking.NeedPaint = True
            Set NowIce = LastClicking
            LastClicking.Display
            Set LastClicking = Nothing
        End If
        ParentUI.MouseColor.Color = argb(255, 255 - ParentUI.BackColor.R, 255 - ParentUI.BackColor.G, 255 - ParentUI.BackColor.b)
    End If
    'Layout Clicked
End Sub
Sub LostMouse()
    If (Not LastClicking Is Nothing) Then
        LastClicking.ClickState = MouseLeave: LastClicking.ClickButton = NoButton: LastClicking.NeedPaint = True
        LastClicking.Display
        Set LastClicking = Nothing
    End If
End Sub
Sub MakeIcelolly(ByVal GunID As String, ByVal Id As String, Pos As EPosition, Size As ESize, Style As StyleBox)
    ReDim Preserve Icelollys(UBound(Icelollys) + 1)
    ReDim Preserve OrPos(UBound(OrPos) + 1)

    Set Icelollys(UBound(Icelollys)) = New Icelolly
    With Icelollys(UBound(Icelollys))
        .Class = GunID
        .Id = Id
        .Pos = Pos
        .Size = Size
        .Style = Style
        Set .ParentLayout = Me
    End With
    OrPos(UBound(OrPos)) = Pos
    On Local Error Resume Next
    Set NowIce = Icelollys(UBound(Icelollys))
    If Id <> "" Then CallByName EventBox, Id & "_Start", VbMethod
    If Err.Number <> 0 Then IceError 4039, "Easoft.Icelolly.Events.Start", "Warning", "Failed to call .", Icelollys(UBound(Icelollys))
    Err.Clear
End Sub
Function Count() As Long
    Count = UBound(Icelollys)
End Function
Function IcelollyByIndex(ByVal Index As Integer) As Icelolly
    If Not Icelollys(Index) Is Nothing Then Set IcelollyByIndex = Icelollys(Index)
End Function
Function IcelollyByID(ByVal Id As String, Optional ByVal Index As Integer = 0) As Icelolly
    Dim s As Integer
    For i = 1 To UBound(Icelollys)
        If Icelollys(i).Id = Id Then
            If s = Index Then Set IcelollyByID = Icelollys(i): Exit For
            s = s + 1
        End If
    Next
End Function
Private Sub Class_Initialize()
    If StartEasoft = False Then Err.Raise 4046, , "Blackmouth's love" & vbCrLf & "Start Easoft first ."
    ReDim Icelollys(0)
    ReDim OrPos(0)
    NeedPaint = True
    BackColor.Color = argb(0, 255, 255, 255)
    mVisible = True
End Sub
