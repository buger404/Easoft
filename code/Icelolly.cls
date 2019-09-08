VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Icelolly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mPos As EPosition, mPos2 As EPosition, mSize As ESize, mMax As Long, mValue As Long, mText As String
Dim mVisible As Boolean, mCustomDraw As Boolean, mClass As String
Dim mStyle As StyleBox
Dim mClickState As EClickState, mClickButton As EClickButton, mClickPos As EPosition
Dim mMulti As Boolean, mIsOn As Boolean
Dim mItemHeight As Long
Public Id As String, BlockClick As Boolean
Public Draw As New Drawing, Src As New EImage, List As New ListMemers
Public NeedPaint As Boolean
Public ParentLayout As Layout
Public AnimateColor As New ColorMix, AnimateColor2 As New ColorMix
Public LinkWheel As Icelolly
Public Animate As New AniManager, DrawAnimate As New AniManager
Public ErrorInfo As String
Dim mTag(5) As String, mDrawTag(5) As String
Dim mStartLine As Long, mStartChr As Long, mEndLine As Long, mEndChr As Long
Public Property Get StartChr() As Long
    StartChr = mStartChr
End Property
Public Property Let StartChr(ByVal nStartChr As Long)
    mStartChr = nStartChr
    AddRefIce Me
End Property
Public Property Get EndChr() As Long
    EndChr = mEndChr
End Property
Public Property Let EndChr(ByVal nEndChr As Long)
    mEndChr = nEndChr
    AddRefIce Me
End Property
Public Property Get EndLine() As Long
    EndLine = mEndLine
End Property
Public Property Let EndLine(ByVal nEndLine As Long)
    mEndLine = nEndLine
    AddRefIce Me
End Property
Public Property Get StartLine() As Long
    StartLine = mStartLine
End Property
Public Property Let StartLine(ByVal nStartLine As Long)
    mStartLine = nStartLine
    AddRefIce Me
End Property
Public Function OnKeyDown(ByVal Code As Long) As Boolean
    Dim Cmd As Boolean
    If mClass = "Edit" Then
        Dim Lines() As String, SelMode As Boolean
        Dim StartLine As Long, StartChr As Long, EndLine As Long, EndChr As Long, temp As String
        Lines = Split(mText, vbCrLf): SelMode = ((mEndLine - mStartLine) <> 0) Or ((mEndChr - mStartChr) <> 0)
        If UBound(Lines) = -1 Then ReDim Lines(0)
        
        If (mEndLine - mStartLine) = 0 Then
            StartChr = IIf(mStartChr > mEndChr, mEndChr, mStartChr)
            EndChr = IIf(mStartChr > mEndChr, mStartChr, mEndChr)
            StartLine = mStartLine: EndLine = mEndLine
        Else
            If mStartLine > mEndLine Then
                StartLine = mEndLine: StartChr = mEndChr
                EndLine = mStartLine: EndChr = mStartChr
            Else
                StartLine = mStartLine: StartChr = mStartChr
                EndLine = mEndLine: EndChr = mEndChr
            End If
        End If
        
        If SelMode = True Then
            If (Code = VK_C Or Code = VK_X) And GetAsyncKeyState(VK_CONTROL) < 0 Then
                If StartLine <> EndLine Then
                    temp = Right(Lines(StartLine), Len(Lines(StartLine)) - StartChr)
                    For s = StartLine + 1 To EndLine - 1
                        temp = temp & vbCrLf & Lines(s)
                    Next
                    temp = temp & vbCrLf & Left(Lines(EndLine), EndChr)
                Else
                    temp = Right(Left(Lines(StartLine), EndChr), EndChr - StartChr)
                End If
                If Code = VK_X Then OnInput VK_BACK
                Clipboard.Clear
                Clipboard.SetText temp
                Cmd = True
            End If
        End If

        If Code = VK_V And GetAsyncKeyState(VK_CONTROL) < 0 Then
            temp = Clipboard.GetText
            If temp = "" Then VBA.Beep: Exit Function
            If SelMode = True Then OnInput VK_BACK
            For i = 1 To Len(temp)
                If Asc(mID(temp, i, 1)) = 13 Then
                    OnInput VK_RETURN
                Else
                    OnInput Asc(mID(temp, i, 1))
                End If
            Next
            Cmd = True
        End If
        If Code = VK_A And GetAsyncKeyState(VK_CONTROL) < 0 Then
            mStartChr = 0: mStartLine = 0: mEndChr = Len(Lines(UBound(Lines))): mEndLine = UBound(Lines)
            Cmd = True
        End If
        If Code = VK_DOWN Then
            mStartLine = mStartLine + 1
            If mStartLine > UBound(Lines) Then mStartLine = UBound(Lines)
            If mStartChr > Len(Lines(mStartLine)) Then mStartChr = Len(Lines(mStartLine))
            mEndChr = mStartChr: mEndLine = mStartLine
            Cmd = True
        End If
        If Code = VK_UP Then
            mStartLine = mStartLine - 1
            If mStartLine < 0 Then mStartLine = 0
            If mStartChr > Len(Lines(mStartLine)) Then mStartChr = Len(Lines(mStartLine))
            mEndChr = mStartChr: mEndLine = mStartLine
            Cmd = True
        End If
        If Code = VK_RIGHT Then
            mStartChr = mStartChr + 1
            If mStartChr > Len(Lines(mStartLine)) Then mStartLine = mStartLine + 1: mStartChr = 0
            If mStartLine > UBound(Lines) Then mStartLine = UBound(Lines): mStartChr = Len(Lines(UBound(Lines)))
            mEndChr = mStartChr: mEndLine = mStartLine
            Cmd = True
        End If
        If Code = VK_LEFT Then
            mStartChr = mStartChr - 1
            If mStartChr < 0 Then mStartLine = mStartLine - 1: mStartChr = Len(Lines(mStartLine))
            If mStartLine < 0 Then mStartLine = 0: mStartChr = 0
            mEndChr = mStartChr: mEndLine = mStartLine
            Cmd = True
        End If
        If Code = VK_A And GetAsyncKeyState(VK_CONTROL) < 0 Then
            mStartChr = 0: mStartLine = 0: mEndChr = Len(Lines(UBound(Lines))): mEndLine = UBound(Lines)
            Cmd = True
        End If
        If Cmd = True Then
            NeedPaint = True
            mDrawTag(0) = GetTickCount
            Exit Function
        End If
    End If
    
    If mClass = "HScroll" Then
        If Code = VK_LEFT Then Value = mValue - 1
        If Code = VK_RIGHT Then Value = mValue + 1
        Exit Function
    End If
    If mClass = "VScroll" Then
        If Code = VK_UP Then Value = mValue - 1
        If Code = VK_DOWN Then Value = mValue + 1
        Exit Function
    End If
    If mClass = "List" Then
        If Code = VK_TAB Then List.Index = List.Index + 1
        If Code = VK_UP Then List.Index = List.Index - 1
        If Code = VK_DOWN Then List.Index = List.Index + 1
        If List.Index > List.Count - 1 Then List.Index = 0
        If List.Index < 0 Then List.Index = List.Count - 1
        Exit Function
    End If
    
    OnKeyDown = True
End Function
Public Function OnKeyUp(ByVal Code As Long) As Boolean
    
    OnKeyUp = True
End Function
Public Sub OnInput(ByVal Ascii As Long)
    If mClass <> "Edit" Then
        If Ascii = VK_RETURN Then
            ParentLayout.CheckMouse mPos, LButton, MouseDown
            ParentLayout.CheckMouse mPos, LButton, MouseUp
        Else
            VBA.Beep
        End If
        Exit Sub
    End If
    'Lines：每一行的文本
    Dim Lines() As String, SelMode As Boolean
    Dim StartLine As Long, StartChr As Long, EndLine As Long, EndChr As Long
    Lines = Split(mText, vbCrLf): SelMode = ((mEndLine - mStartLine) <> 0) Or ((mEndChr - mStartChr) <> 0)
    If UBound(Lines) = -1 Then ReDim Lines(0)
    NeedPaint = True
    mDrawTag(0) = GetTickCount
    
    If (mEndLine - mStartLine) = 0 Then
        StartChr = IIf(mStartChr > mEndChr, mEndChr, mStartChr)
        EndChr = IIf(mStartChr > mEndChr, mStartChr, mEndChr)
        StartLine = mStartLine: EndLine = mEndLine
    Else
        If mStartLine > mEndLine Then
            StartLine = mEndLine: StartChr = mEndChr
            EndLine = mStartLine: EndChr = mStartChr
        Else
            StartLine = mStartLine: StartChr = mStartChr
            EndLine = mEndLine: EndChr = mEndChr
        End If
    End If
    
    If SelMode = True Then
        If StartLine <> EndLine Then
            Lines(StartLine) = Left(Lines(StartLine), StartChr) & Right(Lines(EndLine), Len(Lines(EndLine)) - EndChr)
            For i = StartLine + 1 To EndLine - 1
                Lines(i) = "[!@%#Clear#!@%]"
            Next
            Lines(EndLine) = "[!@%#Clear#!@%]"
        Else
            Lines(StartLine) = Left(Lines(StartLine), StartChr) & Right(Lines(StartLine), Len(Lines(StartLine)) - EndChr)
        End If
        mEndChr = StartChr: mEndLine = StartLine
        mStartChr = StartChr: mStartLine = StartLine
        If Ascii = VK_BACK Then GoTo Last
    End If
    
    Select Case Ascii
        Case VK_BACK
            If StartChr - 1 >= 0 Then
                Lines(StartLine) = Left(Lines(StartLine), StartChr - 1) & Right(Lines(StartLine), Len(Lines(StartLine)) - StartChr)
                mStartChr = mStartChr - 1: mEndChr = mEndChr - 1
            ElseIf StartLine - 1 >= 0 Then
                mStartChr = mStartChr + Len(Lines(StartLine - 1)): mEndChr = mStartChr
                Lines(StartLine - 1) = Lines(StartLine - 1) & Lines(StartLine)
                Lines(StartLine) = "[!@%#Clear#!@%]"
                mStartLine = mStartLine - 1: mEndLine = mEndLine - 1
            Else
                VBA.Beep
            End If
        Case VK_RETURN
            If mMulti = False Then
                VBA.Beep
            Else
                Lines(StartLine) = Left(Lines(StartLine), StartChr) & vbCrLf & Right(Lines(StartLine), Len(Lines(StartLine)) - StartChr)
                mStartLine = StartLine + 1: mEndLine = mStartLine
                mStartChr = 0: mEndChr = 0
            End If
        Case Else
        Lines(StartLine) = Left(Lines(StartLine), StartChr) & Chr(Ascii) & Right(Lines(StartLine), Len(Lines(StartLine)) - StartChr)
        mStartChr = mStartChr + 1: mEndChr = mEndChr + 1
    End Select
    
Last:
    mText = ""
    For i = 0 To UBound(Lines) - 1
        If Lines(i) <> "[!@%#Clear#!@%]" Then mText = mText & Lines(i) & vbCrLf
    Next
    If Lines(UBound(Lines)) <> "[!@%#Clear#!@%]" Then mText = mText & Lines(UBound(Lines))
End Sub
Public Sub OnMouseWheel(ByVal Direction As Integer)
    If mClass = "VScroll" Then
        Value = mValue - 5 * Direction
    End If
    If Not LinkWheel Is Nothing Then
        If LinkWheel.Class = "VScroll" Then
            LinkWheel.OnMouseWheel Direction: AddRefIce Me
        Else
            LinkWheel.NeedPaint = True
        End If
    End If
End Sub
Public Property Get Tag(ByVal Index As Integer) As String
    Tag = mTag(Index)
End Property
Public Property Let Tag(ByVal Index As Integer, ByVal nTag As String)
    mTag(Index) = nTag
End Property
Public Property Get DrawTag(ByVal Index As Integer) As String
    DrawTag = mDrawTag(Index)
End Property
Public Property Let DrawTag(ByVal Index As Integer, ByVal nDrawTag As String)
    mDrawTag(Index) = nDrawTag
End Property
Public Property Get Width() As Single
    Width = mSize.Width
End Property
Public Property Let Width(nWidth As Single)
    mSize.Width = nWidth
    Draw.Size = mSize
    AddRefIce Me
End Property
Public Property Get Height() As Single
    Height = mSize.Height
End Property
Public Property Let Height(nHeight As Single)
    mSize.Height = nHeight
    Draw.Size = mSize
    AddRefIce Me
End Property
Public Property Get ClickPos() As EPosition
    ClickPos = mClickPos
End Property
Public Property Let ClickPos(nClickPos As EPosition)
    mClickPos = nClickPos
End Property
Public Property Get ClickButton() As EClickButton
    ClickButton = mClickButton
End Property
Public Property Let ClickButton(nClickButton As EClickButton)
    mClickButton = nClickButton
End Property
Public Property Get ClickState() As EClickState
    ClickState = mClickState
End Property
Public Property Let ClickState(nClickState As EClickState)
    mClickState = nClickState
End Property
Public Property Get Style() As StyleBox
    Set Style = mStyle
End Property
Public Property Let Style(nStyle As StyleBox)
    Set mStyle = nStyle
    AddRefIce Me
End Property
Public Property Get CustomDraw() As Boolean
    CustomDraw = mCustomDraw
End Property
Public Property Let CustomDraw(nCustomDraw As Boolean)
    mCustomDraw = nCustomDraw
    AddRefIce Me
End Property
Public Property Get IsOn() As Boolean
    IsOn = mIsOn
End Property
Public Property Let IsOn(nIsOn As Boolean)
    If ParentLayout.Oning = True Then GoTo Skip
    If Id = "" Then GoTo Skip
    If Multi = False And nIsOn = False Then
        Dim Found As Integer
        For i = 1 To ParentLayout.Count
            If ParentLayout.IcelollyByIndex(i).Id = Id Then
                If ParentLayout.IcelollyByIndex(i).IsOn = True And (Not ParentLayout.IcelollyByIndex(i) Is Me) Then Found = i: Exit For
            End If
        Next
        If Found = 0 Then Exit Property
    End If
    If Multi = False Then
        ParentLayout.Oning = True
        For i = 1 To ParentLayout.Count
            If ParentLayout.IcelollyByIndex(i).Id = Id Then
                If ParentLayout.IcelollyByIndex(i).IsOn = True Then ParentLayout.IcelollyByIndex(i).IsOn = False: Exit For
            End If
        Next
        ParentLayout.Oning = False
    End If
Skip:
    mIsOn = nIsOn
    AddRefIce Me
End Property
Public Property Get Multi() As Boolean
    Multi = mMulti
End Property
Public Property Let Multi(nMulti As Boolean)
    If ParentLayout.Oning = True Then GoTo Skip
    If Id = "" Then GoTo Skip
    
    ParentLayout.Oning = True
    If nMulti = True Then
        Dim Found As Integer
        For i = 1 To ParentLayout.Count
            If ParentLayout.IcelollyByIndex(i).Id = Id Then
                If ParentLayout.IcelollyByIndex(i).IsOn = True Then Found = i: Exit For
            End If
        Next
        If Found = 0 Then mIsOn = True
    Else
        For i = 1 To ParentLayout.Count
            If ParentLayout.IcelollyByIndex(i).Id = Id Then
                If ParentLayout.IcelollyByIndex(i).IsOn = True Then ParentLayout.IcelollyByIndex(i).IsOn = False
            End If
        Next
    End If
    
    For i = 1 To ParentLayout.Count
        If ParentLayout.IcelollyByIndex(i).Id = Id Then
            ParentLayout.IcelollyByIndex(i).Multi = nMulti
        End If
    Next
    
    ParentLayout.Oning = False
Skip:
    mMulti = nMulti
    AddRefIce Me
End Property
Public Property Get Visible() As Boolean
    Visible = mVisible
End Property
Public Property Let Visible(nVisible As Boolean)
    mVisible = nVisible
    AddRefIce Me
End Property
Public Property Get Pos2() As EPosition
    Pos2 = mPos2
End Property
Public Property Let Pos2(nPos2 As EPosition)
    mPos2 = nPos2
    AddRefIce Me
End Property
Public Property Get X2() As Single
    X2 = mPos2.x
End Property
Public Property Let X2(nX2 As Single)
    mPos2.x = nX2
    AddRefIce Me
End Property
Public Property Get Y2() As Single
    Y2 = mPos2.y
End Property
Public Property Let Y2(nY2 As Single)
    mPos2.y = nY2
    AddRefIce Me
End Property
Public Property Get Pos() As EPosition
    Pos = mPos
End Property
Public Property Let Pos(nPos As EPosition)
    mPos = nPos
    AddRefIce Me
End Property
Public Property Get x() As Single
    x = mPos.x
End Property
Public Property Let x(nX As Single)
    mPos.x = nX
    AddRefIce Me
End Property
Public Property Get y() As Single
    y = mPos.y
End Property
Public Property Let y(nY As Single)
    mPos.y = nY
    AddRefIce Me
End Property
Public Property Get Size() As ESize
    Size = mSize
End Property
Public Property Let Size(nSize As ESize)
    mSize = nSize
    Draw.Size = nSize
    AddRefIce Me
End Property
Public Property Get Value() As Long
    Value = mValue
End Property
Public Property Let Value(ByVal nValue As Long)
    mValue = nValue
    If mValue < 0 Then mValue = 0
    If mValue > mMax Then mValue = mMax
    On Error Resume Next
    If Id <> "" And (mClass = "VScroll" Or mClass = "HScroll") Then CallByName ParentLayout.EventBox, Id & "_OnScroll", VbMethod
    If Not LinkWheel Is Nothing Then
        If LinkWheel.Class = "VScroll" Then
            LinkWheel.OnMouseWheel Direction: AddRefIce Me
        Else
            LinkWheel.NeedPaint = True
        End If
    End If
    AddRefIce Me
End Property
Public Property Get ItemHeight() As Long
    ItemHeight = mItemHeight
End Property
Public Property Let ItemHeight(ByVal nItemHeight As Long)
    If nItemHeight = 0 Then Err.Raise 40411, , "What the f***" & vbCrLf & "ItemHeight cannot be 0 , please try another ItemHeight ."
    mItemHeight = nItemHeight
    AddRefIce Me
End Property
Public Property Get Max() As Long
    Max = mMax
End Property
Public Property Let Max(ByVal nMax As Long)
    If nMax = 0 Then Err.Raise 40411, , "What the f***" & vbCrLf & "Max cannot be 0 , please try another max ."
    mMax = nMax
    AddRefIce Me
End Property
Public Property Get Class() As String
    Class = mClass
End Property
Public Property Let Class(ByVal nClass As String)
    mClass = nClass
    AddRefIce Me
End Property
Public Property Get Text() As String
    Text = mText
End Property
Public Property Let Text(ByVal nText As String)
    mText = nText
    AddRefIce Me
End Property
Public Sub Display()
    Set NowIce = Me
    On Local Error GoTo sth
    GdipGraphicsClear Draw.Hwnd, 0
    Animate.Update: DrawAnimate.Update
    
sth:
    If Err.Number <> 0 Then IceError 4045, "Easoft.Icelolly.Animation(Count=" & Animate.Count & "&" & DrawAnimate.Count & ")", "Drawing Error", Err.Description, NowIce
    Err.Clear
    
    If CustomDraw = False Then
        Call OurDraw(mClass)
    Else
        On Error GoTo sth2
        CallByName EventBox, Id & "_OnDraw", VbMethod
sth2:
        If Err.Number <> 0 Then IceError 4048, "Easoft.Icelolly.Drawing.Custom." & mClass, "Drawing Error", Err.Description, NowIce: Err.Raise 4048, , "Blackmouth's loooove" & vbCrLf & "Custom drawing turned on, but did not get the desired effect."
    End If
    Displaying = False
End Sub
Private Sub Class_Initialize()
    If StartEasoft = False Then Err.Raise 4046, , "Blackmouth's love" & vbCrLf & "Start Easoft first ."
    mClass = "Label"
    mVisible = True
    mMax = 100
    mItemHeight = 30
    AddIcelolly Me
    AddRefIce Me
End Sub
