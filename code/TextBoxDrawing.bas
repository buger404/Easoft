Attribute VB_Name = "TextBoxDrawing"
Sub DrawTextBox()
    If NowIce.Multi = True And NowIce.LinkWheel Is Nothing Then '当允许多行但没有绑定的滚动条时
    NowIce.Draw.DrawString NewPos(0, 0), NowIce.Size, "Set the scrollbar first !", NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore): Exit Sub
    End If
    
    'Lines：每一行的文本
    'ChrWidths：当前点击行的字符的宽度集合
    'EgHeight：当前字体的示例高度
    'TextHeight：文本总高度
    'temp：用于返回gdip字符大小时所用
    Dim Lines() As String, ChrWidths() As Long, EgHeight As Long, TextHeight As Long, temp As RECTF
    Dim DrawY As Long, DrawY2 As Long, StartDraw As Long, CanDraw As Long, SelMode As Boolean
    Dim FocusX As Single, LineWidth As Long
    Dim StartLine As Long, StartChr As Long, EndLine As Long, EndChr As Long
    Dim DrawFocus As Boolean, FocusStep As Single '画过文本框焦点了没？
    FocusStep = (GetTickCount - Val(NowIce.DrawTag(0)) + 400) Mod 1500
    '设置文本框焦点浮动
    If FocusStep <= 400 Then
        NowIce.Style.Color(Border2).a = FocusStep / 400 * 255
    ElseIf FocusStep >= 550 And FocusStep <= 550 + 600 Then
        NowIce.Style.Color(Border2).a = 255 - (FocusStep - 550) / 600 * 255
    End If

    Lines = Split(NowIce.Text, vbCrLf): SelMode = ((NowIce.EndLine - NowIce.StartLine) <> 0) Or ((NowIce.EndChr - NowIce.StartChr) <> 0)
    
    If (NowIce.EndLine - NowIce.StartLine) = 0 Then
        StartChr = IIf(NowIce.StartChr > NowIce.EndChr, NowIce.EndChr, NowIce.StartChr)
        EndChr = IIf(NowIce.StartChr > NowIce.EndChr, NowIce.StartChr, NowIce.EndChr)
        StartLine = NowIce.StartLine: EndLine = NowIce.EndLine
    Else
        If NowIce.StartLine > NowIce.EndLine Then
            StartLine = NowIce.EndLine: StartChr = NowIce.EndChr
            EndLine = NowIce.StartLine: EndChr = NowIce.StartChr
        Else
            StartLine = NowIce.StartLine: StartChr = NowIce.StartChr
            EndLine = NowIce.EndLine: EndChr = NowIce.EndChr
        End If
    End If
    
    If NowIce.ParentLayout.ParentUI.FocusIcelolly Is NowIce Then NowIce.DrawAnimate.AddByReset NewAnimate(NowIce.AnimateColor, "A", 0, 1000000, 0, 0) 'Keep drawing
    If NowIce.ClickState = MouseLeave Then NowIce.DrawAnimate.Clear
    
    '取得EgHeight
    'GdipStringFormatGetGenericTypographic StrFormat(NowIce.Style.Align.Horizontal)   '让狗屎API听话用的
    GdipMeasureString NowIce.Draw.Hwnd, StrPtr("Eg"), 2, NowIce.Style.Font.Hwnd, NewRectF(0, 0, 0, 0), EditStrFormat(NowIce.Style.Align.Horizontal), temp, 0, 0
    EgHeight = temp.Bottom
    
    '计算总文本高度&控件可画行数
    TextHeight = (UBound(Lines) + 1) * EgHeight
    CanDraw = Round(NowIce.Height / EgHeight)
    
    '当前滚动位置
    If NowIce.Multi = True And TextHeight > NowIce.Height Then DrawY = NowIce.LinkWheel.Value / NowIce.LinkWheel.Max * (TextHeight - NowIce.Height)
    
    '当前开始绘制的行
    StartDraw = Int(DrawY / EgHeight)
    
    '绘制
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Back)
    NowIce.Draw.DrawShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Border)
    NowIce.Draw.DrawImageRect NewPos(0, 0), NowIce.Size, NowIce.Src
    
    If NowIce.ParentLayout.ParentUI.FocusIcelolly Is NowIce And NowIce.Text = "" Then NowIce.Draw.DrawLine NewPos(0, 0 + EgHeight * 0.1), NewPos(0, 0 + EgHeight * 0.8), NowIce.Style.Color(Border2)
    
    For i = StartDraw To StartDraw + CanDraw
        If i > UBound(Lines) Then Exit For
        DrawY2 = i * EgHeight - DrawY
        NowIce.Draw.DrawEditString NewPos(0, DrawY2), NewSize(NowIce.Width, EgHeight), Lines(i), NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
        GdipMeasureString NowIce.Draw.Hwnd, StrPtr(Lines(i)), Len(Lines(i)), NowIce.Style.Font.Hwnd, NewRectF(0, DrawY2, 0, 0), EditStrFormat(NowIce.Style.Align.Horizontal), temp, 0, 0
        LineWidth = temp.Right
        '选取处理
        If i = StartLine Or i = EndLine Then
            '计算位置
            FocusX = 0 '清0
            GdipMeasureString NowIce.Draw.Hwnd, StrPtr(mID(Lines(i), 1, IIf(i = StartLine, StartChr, EndChr))), -1, NowIce.Style.Font.Hwnd, NewRectF(0, DrawY2, 0, 0), EditStrFormat(NowIce.Style.Align.Horizontal), temp, 0, 0
            FocusX = temp.Right

            If SelMode = False Then
                If NowIce.ParentLayout.ParentUI.FocusIcelolly Is NowIce Then
                    DrawFocus = True
                    NowIce.Draw.DrawLine NewPos(FocusX, DrawY2 + EgHeight * 0.1), NewPos(FocusX, DrawY2 + EgHeight * 0.8), NowIce.Style.Color(Border2)
                End If
            Else
                '分支
                If (EndLine - StartLine) <> 0 Then
                    NowIce.Draw.FillRect NewPos(IIf(i = StartLine, FocusX, 0), DrawY2), NewSize(IIf(i = StartLine, LineWidth - FocusX, FocusX), EgHeight), NowIce.Style.Color(Active)
                Else
                    Dim FocusX2 As Single
                    '计算位置
                    FocusX2 = 0 '清0
                    GdipMeasureString NowIce.Draw.Hwnd, StrPtr(mID(Lines(i), 1, EndChr)), -1, NowIce.Style.Font.Hwnd, NewRectF(0, DrawY2, 0, 0), EditStrFormat(NowIce.Style.Align.Horizontal), temp, 0, 0
                    FocusX2 = temp.Right
                    NowIce.Draw.FillRect NewPos(FocusX, DrawY2), NewSize(FocusX2 - FocusX, EgHeight), NowIce.Style.Color(Active)
                End If
            End If
        End If
        
        '其他
        If (EndLine - StartLine) <> 0 Then
            If i > StartLine And i < EndLine Then
                NowIce.Draw.FillRect NewPos(0, DrawY2), NewSize(LineWidth, EgHeight), NowIce.Style.Color(Active)
            End If
        End If
    Next
    
    '位置更改处理
    Dim SelLine As Long, SelChr As Long
    If NowIce.ClickState = MouseDown Then
        SelLine = Int((NowIce.ClickPos.y + DrawY) / EgHeight)
        If SelLine > UBound(Lines) Then SelLine = UBound(Lines)
        If SelLine < 0 Then SelLine = 0
        '计算位置
        FocusX = 0 '清0
        SelChr = Len(Lines(SelLine))
        For s = 1 To Len(Lines(SelLine))
            GdipMeasureString NowIce.Draw.Hwnd, StrPtr(mID(Lines(SelLine), 1, s)), -1, NowIce.Style.Font.Hwnd, NewRectF(0, 0, 0, 0), EditStrFormat(NowIce.Style.Align.Horizontal), temp, 0, 0
            FocusX = temp.Right
            If FocusX > NowIce.ClickPos.x Then SelChr = s - 1: Exit For
        Next
        If NowIce.DrawTag(5) = "" Then '刚开始的选取
            NowIce.StartLine = SelLine: NowIce.StartChr = SelChr: NowIce.DrawTag(5) = "hmmm"
            NowIce.EndLine = SelLine: NowIce.EndChr = SelChr
        Else
            NowIce.EndLine = SelLine: NowIce.EndChr = SelChr
        End If
    End If
    
    If NowIce.ClickState = MouseUp Then NowIce.DrawTag(5) = ""
    
End Sub
