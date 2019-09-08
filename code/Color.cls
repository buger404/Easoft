VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Type EPosition
    x As Single
    y As Single
End Type
Public Type ESize
    Width As Single
    Height As Single
End Type
Public Type EAnimate
    Target As Long
    FuncName As String
    Obj As Object
    Start As Long
    ProperName As String
    Delay As Long
    Duration As Long
    StartTime As Long
    Index As Integer
End Type
Public Enum EColors
    EBlue = -11969281
    EGreen = -12072237
    ERed = -119009
End Enum
Public Enum EFontStyle
    Regular = 0
    Bold = 1
    Italic = 2
    BoldItalic = 3
    Underline = 4
    Light = 6
    Strikeout = 8
End Enum
Public Enum EAlign1
    OnLeft = 0
    OnCenter = 1
    OnRight = 2
End Enum
Public Enum EAlign2
    OnTop = 0
    OnMiddle = 1
    OnBottom = 2
End Enum
Public Type EAlign
    Horizontal As EAlign1
    Vertical As EAlign2
End Type
Public Enum EClickState
    None = 0
    MouseDown = 1
    MouseMove = 2
    MouseEnter = 4
    MouseLeave = 5
    MouseUp = 3
End Enum
Public Enum EColorTheme
    Back = 0
    Fore = 1
    Border = 2
    Border2 = 3
    Active = 4
    Animate = 5
End Enum
Public Enum EClickButton
    NoButton = 0
    LButton = 1
    RButton = 2
    MButton = 3
End Enum
Public Enum EShape
    Square = 0
    Oval = 1
    RoundRect = 2
End Enum
Public Enum EWindowStyle
    LayeredWindow = 1
    AeroWindow = 2
End Enum
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
Public Function argb(ByVal a As Long, ByVal R As Long, ByVal G As Long, ByVal b As Long) As Long
    Dim Color As Long
    CopyMemory ByVal VarPtr(Color) + 3, a, 1
    CopyMemory ByVal VarPtr(Color) + 2, R, 1
    CopyMemory ByVal VarPtr(Color) + 1, G, 1
    CopyMemory ByVal VarPtr(Color), b, 1
    argb = Color
End Function
Public Function Msgbox(ByVal Context As String, ByVal Title As String, ByVal CaptionColor As Long, ByVal BColor As Long, BtnStyle As StyleBox)
    'Eg  Msgbox "hi","Warning",EColors.EGreen,MyBtn
    
    'Apply colors
    MsgWindow.TitleText.Style.Color(Fore).Color = CaptionColor
    MsgWindow.CaptionStyle.Color(Border).Color = CaptionColor
    MsgWindow.EUI.BackColor.Color = BColor
    
    'Set text
    MsgWindow.TitleText.Text = Title
    MsgWindow.Context.Text = Context
    MsgWindow.Caption = Title
    
    'Apply styles
    MsgWindow.YesBtn.Style = BtnStyle
    MsgWindow.NoBtn.Style = BtnStyle
    
    'Show
    MsgWindow.EUI.Refresh
    MsgWindow.EUI.Display
    MsgWindow.DrawTimer.Enabled = True
    MsgWindow.Show
    MsgWindow.Choice = -1
    
    'Wait
    Do While MsgWindow.Choice = -1
        Sleep 16: DoEvents
    Loop
    
    'Return
    Msgbox = MsgWindow.Choice
    MsgWindow.DrawTimer.Enabled = False
    MsgWindow.Hide
End Function
Public Sub EasoftPower(ByVal Power As Boolean)
    If Power = True Then
        If StartEasoft = True Then Err.Raise 4049, , "Blackmouth's angry" & vbCrLf & "You did not recycle resources or recycle failed last time, do you want to see the overmemory dialog?"
        If GetSetting("Easoft", "FirstUse", "First") = "" Then
            SaveSetting "Easoft", "FirstUse", "First", "True"
            For i = 1 To 4
                VBA.Beep
                Sleep 1000
            Next
            Debug.Print "* 很高兴你使用了Easoft，List还没有完成，请多关照。"
            Sleep 3000
            Debug.Print "* 你很幸运哦！Easoft的Debug对话触发条件是第一次引用Easoft并初始化呢！"
            Sleep 3000
            Debug.Print "* 也就是说，如果你在第一次使用前打开了使用Easoft的程序，那么你就不会看到这些信息。"
            Sleep 3000
            Debug.Print "* 使用者正在我面前阅读你留下的文字，这使我充满了决心。"
            Sleep 3000
        End If
        ReDim EMembers(0), EShadows(0)
        InitGDIPlus
        GdipCreateEffect2 GdipEffectType.Blur, BlurEffect
        BlurParam.Radius = 255
        GdipSetEffectParameters BlurEffect, BlurParam, LenB(BlurParam)
        StartEasoft = True
        BeginDirtyRect
        Dim Flag As Long
        For i = 0 To 2
            GdipCreateStringFormat 0, 0, StrFormat(i)
            GdipSetStringFormatAlign StrFormat(i), i

            GdipStringFormatGetGenericTypographic EditStrFormat(i)
            GdipGetStringFormatFlags EditStrFormat(i), Flag
            GdipSetStringFormatFlags EditStrFormat(i), (Flag Or StringFormatFlagsMeasureTrailingSpaces)
            'GdipSetStringFormatAlign EditStrFormat(i), i
        Next
        DebugB.Color = argb(70, 255, 0, 0)
        Load MsgWindow
    Else
        FreeWndProc
        DeleteAllMember
        For i = 0 To 2
            GdipDeleteStringFormat StrFormat(i)
            GdipDeleteStringFormat EditStrFormat(i)
        Next
        TerminateGDIPlus
        StartEasoft = False
    End If
End Sub
Public Function DebugWin() As Object
    Set DebugWin = DebugWindow
End Function
Public Function this() As Icelolly
    Set this = NowIce
End Function
