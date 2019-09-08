Attribute VB_Name = "DrawIcelolly"
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Dim EverIcelolly As Boolean, LastClickTime As Long
Sub OurDraw(ByVal Class As String)
'On Error GoTo sth
    Select Case Class
        Case "Label"
            Call DrawLabel
        Case "Button"
            Call DrawButton
        Case "Option"
            Call DrawOption
        Case "Check"
            Call DrawCheck
        Case "HScroll"
            Call DrawHScroll
        Case "VScroll"
            Call DrawVScroll
        Case "Progress"
            Call DrawProgress
        Case "ArcProgress"
            Call DrawArcProgress
        Case "Shape"
            Call DrawShape
        Case "Line"
            Call DrawLine
        Case "Image"
            Call DrawImage
        Case "Slider"
            Call DrawSlider
        Case "List"
            Call DrawList
        Case "Loading"
            Call DrawLoading
        Case "Edit"
            Call DrawTextBox
        Case "TRUE_ICELOLLY_HIDDEN_EASOFT_CONTROL"
            Call DrawTrueIcelolly
    End Select
sth:
If Err.Number <> 0 Then
    IceError 4046, "Easoft.Icelolly.Drawing." & Class, "Drawing Error", Err.Description, NowIce
'    Err.Raise 4046, , "Icelolly's (╯‵□′)╯︵┻━┻" & vbCrLf & "An exception occurred while the drawing was proceeding."
End If
End Sub
Function GetMouseColor(ByVal Class As String, Ice As Icelolly) As Long
    Select Case Class
        Case "Label"
            GetMouseColor = Ice.Style.Color(Fore).Color
        Case "Button"
            GetMouseColor = Ice.Style.Color(Active).Color
        Case "Option"
            If Ice.IsOn = False Then
                GetMouseColor = Ice.Style.Color(Border2).Color
            Else
                GetMouseColor = Ice.Style.Color(Border).Color
            End If
        Case "Check"
            If Ice.IsOn = False Then
                GetMouseColor = Ice.Style.Color(Back).Color
            Else
                GetMouseColor = Ice.Style.Color(Fore).Color
            End If
        Case "HScroll"
            With Ice.ParentLayout.ParentUI.BackColor
                GetMouseColor = argb(255, 255 - .R, 255 - .G, 255 - .b)
            End With
        Case "VScroll"
            With Ice.ParentLayout.ParentUI.BackColor
                GetMouseColor = argb(255, 255 - .R, 255 - .G, 255 - .b)
            End With
        Case "Progress"
            GetMouseColor = Ice.Style.Color(Fore).Color
        Case "ArcProgress"
            GetMouseColor = Ice.Style.Color(Border).Color
        Case "Shape"
            GetMouseColor = Ice.Style.Color(Back).Color
        Case "Line"
            GetMouseColor = Ice.Style.Color(Border).Color
        Case "Slider"
            GetMouseColor = Ice.Style.Color(Fore).Color
        Case "List"
            GetMouseColor = Ice.Style.Color(Active).Color
        Case "Loading"
            GetMouseColor = Ice.Style.Color(Border).Color
        Case "Edit"
            GetMouseColor = Ice.Style.Color(Fore).Color
        Case "TRUE_ICELOLLY_HIDDEN_EASOFT_CONTROL"
            Randomize
            GetMouseColor = argb(255, Rnd * 255, Rnd * 255, 255)
    End Select
End Function
Sub DrawLoading()
    'This loading powered by 洛小羽 , design by 洛小羽 , copy by 404 .
    'DrawTag : 0 = Angle , 1 = Start , 2 = Draw Switch

    NowIce.DrawAnimate.AddByReset NewAnimate(NowIce.AnimateColor, "A", 0, 1000000, 0, 0) 'Keep drawing
    
    If NowIce.DrawTag(2) = "" Then NowIce.DrawTag(2) = -1
    If (Val(-NowIce.DrawTag(2)) = 1) And Val(NowIce.DrawTag(0)) >= 340 Then NowIce.DrawTag(2) = 1
    If (Val(NowIce.DrawTag(2)) = 1) And Val(NowIce.DrawTag(0)) <= 20 Then NowIce.DrawTag(2) = -1
    
    NowIce.DrawTag(0) = Val(NowIce.DrawTag(0)) + IIf(Val(NowIce.DrawTag(2)) = 1, -8, 8)
    NowIce.DrawTag(1) = Val(NowIce.DrawTag(1)) + IIf(Val(NowIce.DrawTag(2)) = 1, 13, 5)
    
    If Val(NowIce.DrawTag(0)) > 360 Then NowIce.DrawTag(0) = Val(NowIce.DrawTag(0)) - 360
    If Val(NowIce.DrawTag(1)) > 360 Then NowIce.DrawTag(1) = Val(NowIce.DrawTag(1)) - 360
    
    NowIce.Draw.DrawArc NewPos(3, 3), NewSize(NowIce.Width - 7, NowIce.Height - 7), Val(NowIce.DrawTag(1)), Val(NowIce.DrawTag(0)), NowIce.Style.Color(Border)
    NowIce.Draw.DrawArc NewPos(8, 8), NewSize(NowIce.Width - 17, NowIce.Height - 17), Val(NowIce.DrawTag(1)) + 180, Val(NowIce.DrawTag(0)), NowIce.Style.Color(Border)
    
End Sub
Sub DrawList()
    If NowIce.LinkWheel Is Nothing Then NowIce.Draw.DrawString NewPos(0, 0), NowIce.Size, "Set the scrollbar first !", NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore): Exit Sub
    
    Dim Start As Long, DrawY As Long, OffY As Long, CanDrawCount As Long
    Dim temp As Single
    CanDrawCount = Round(NowIce.Height / NowIce.ItemHeight)
    temp = NowIce.LinkWheel.Value / NowIce.LinkWheel.Max * (NowIce.List.Count - CanDrawCount)
    Start = Int(temp)
    If Start < 0 Then
        Start = 0
    Else
        OffY = (temp - Int(temp)) * NowIce.ItemHeight
    End If
    
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Back)
    NowIce.Draw.DrawShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Border)
    NowIce.Draw.DrawImageRect NewPos(0, 0), NowIce.Size, NowIce.Src
    
    If NowIce.List.Count = 0 Then NowIce.Draw.DrawString NewPos(0, 0), NowIce.Size, NowIce.List.NothingText, NewAlign(OnCenter, OnTop), NowIce.Style.Font, NowIce.Style.Color(Fore): Exit Sub
    
    For i = Start To Start + CanDrawCount + 1
        If i > NowIce.List.Count Then Exit For
        DrawY = (i - Start - 1) * NowIce.ItemHeight - OffY
        If NowIce.ClickState = MouseMove Or NowIce.ClickState = MouseUp Or NowIce.ClickState = MouseDown Then
            If NowIce.ClickPos.y >= DrawY And NowIce.ClickPos.y <= DrawY + NowIce.ItemHeight Then
                NowIce.Draw.FillRect NewPos(0, DrawY), NewSize(NowIce.Width, NowIce.ItemHeight), NowIce.Style.Color(Border2)
                NowIce.DrawTag(0) = i
                If NowIce.ClickState = MouseUp Then NowIce.List.Index = i
            End If
        End If
        If NowIce.List.Index = i Then NowIce.Draw.FillRect NewPos(0, DrawY), NewSize(NowIce.Width, NowIce.ItemHeight), NowIce.Style.Color(Active)
        
        If NowIce.List.Src(i) Is Nothing Then
            NowIce.Draw.DrawString NewPos(0, DrawY), NewSize(NowIce.Width, NowIce.ItemHeight), NowIce.List.Items(i), NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
        Else
            NowIce.Draw.DrawString NewPos(NowIce.ItemHeight, DrawY), NewSize(NowIce.Width - NowIce.ItemHeight * 2, NowIce.ItemHeight), NowIce.List.Items(i), NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
            NowIce.Draw.DrawImageRect NewPos(0, DrawY), NewSize(NowIce.ItemHeight, NowIce.ItemHeight), NowIce.List.Src(i)
        End If
    Next
End Sub
Sub DrawImage()
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Back)
    NowIce.Draw.DrawShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Border)
    NowIce.Draw.DrawImageRect NewPos(0, 0), NowIce.Size, NowIce.Src
    NowIce.Draw.DrawString NewPos(0, 0), NowIce.Size, NowIce.Text, NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
End Sub
Sub DrawSlider()
    NowIce.Draw.DrawLine NewPos(0, NowIce.Height / 2 - 1 / 2), NewPos(NowIce.Width, NowIce.Height / 2 - 1 / 2), NowIce.Style.Color(Border)
    NowIce.Draw.DrawLine NewPos(0, NowIce.Height / 2 - 1 / 2), NewPos(NowIce.Value / NowIce.Max * (NowIce.Size.Width - NowIce.Size.Height), NowIce.Height / 2 - 1 / 2), NowIce.Style.Color(Border2)
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(NowIce.Value / NowIce.Max * (NowIce.Size.Width - NowIce.Size.Height), 0), NewSize(NowIce.Size.Height, NowIce.Size.Height), NowIce.Style.Radian, NowIce.Style.Color(Fore)
    If NowIce.ClickState = MouseDown Then
        If NowIce.DrawTag(5) = "" Then NowIce.DrawTag(5) = NowIce.ClickPos.x - (NowIce.Value / NowIce.Max * (NowIce.Size.Width - NowIce.Size.Height))
        NowIce.Value = (NowIce.ClickPos.x - Val(NowIce.DrawTag(5))) / (NowIce.Size.Width - NowIce.Size.Height) * NowIce.Max
    End If
    If NowIce.ClickState = MouseUp Then
        NowIce.DrawTag(5) = ""
    End If
End Sub
Sub DrawLine()
    NowIce.Draw.DrawLine NowIce.Pos, NowIce.Pos2, NowIce.Style.Color(Border)
End Sub
Sub DrawShape()
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Back)
    NowIce.Draw.DrawShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Border)
    NowIce.Draw.DrawEImage NewPos(0, 0), NowIce.Size, NowIce.Style.ImgAlign, NowIce.Src
    NowIce.Draw.DrawString NewPos(0, 0), NowIce.Size, NowIce.Text, NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
End Sub
Sub DrawArcProgress()
    NowIce.Draw.FillEllipse NewPos(0, 0), NowIce.Size, NowIce.Style.Color(Back)
    NowIce.Draw.DrawEImage NewPos(0, 0), NowIce.Size, NowIce.Style.ImgAlign, NowIce.Src
    NowIce.Draw.DrawString NewPos(0, 0), NowIce.Size, NowIce.Text, NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
    NowIce.Draw.FillArc2 NewPos(0, 0), NowIce.Size, -90, NowIce.Value / NowIce.Max * 360, NowIce.Style.Color(Fore)
    NowIce.Draw.DrawArc NewPos(0, 0), NowIce.Size, -90, NowIce.Value / NowIce.Max * 360, NowIce.Style.Color(Border)
End Sub
Sub DrawProgress()
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Back)
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NewSize(NowIce.Size.Height + NowIce.Value / NowIce.Max * (NowIce.Size.Width - NowIce.Size.Height), NowIce.Size.Height), NowIce.Style.Radian, NowIce.Style.Color(Fore)
End Sub
Sub DrawVScroll()
    If NowIce.DrawTag(4) = "" Then NowIce.DrawTag(4) = 4
    If NowIce.ClickState = MouseEnter Then
        NowIce.DrawAnimate.AddByReset NewAnimate(NowIce, "DrawTag", 0, 300, NowIce.DrawTag(4), NowIce.Width, "linear", 4)
    End If

    If NowIce.DrawTag(4) <> 4 Then NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(NowIce.Width - NowIce.DrawTag(4), 0), NewSize(NowIce.DrawTag(4), NowIce.Height), NowIce.Style.Radian, NowIce.Style.Color(Back)
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(NowIce.Width - NowIce.DrawTag(4), NowIce.Value / NowIce.Max * NowIce.Size.Height * 0.7), NewSize(NowIce.DrawTag(4), NowIce.Size.Height * 0.3), NowIce.Style.Radian, NowIce.Style.Color(Fore)
    If NowIce.ClickState = MouseDown Then
        If NowIce.DrawTag(5) = "" Then NowIce.DrawTag(5) = NowIce.ClickPos.y - (NowIce.Value / NowIce.Max * NowIce.Size.Height * 0.7)
        NowIce.Value = (NowIce.ClickPos.y - Val(NowIce.DrawTag(5))) / (NowIce.Size.Height * 0.7) * NowIce.Max
    End If
    If NowIce.ClickState = MouseUp Then
        NowIce.DrawTag(5) = ""
    End If
    
    If NowIce.ClickState = MouseLeave Then
        NowIce.DrawAnimate.AddByReset NewAnimate(NowIce, "DrawTag", 0, 300, NowIce.DrawTag(4), 4, "linear", 4)
    End If
End Sub
Sub DrawHScroll()
    If NowIce.DrawTag(4) = "" Then NowIce.DrawTag(4) = 4
    If NowIce.ClickState = MouseEnter Then
        NowIce.DrawAnimate.AddByReset NewAnimate(NowIce, "DrawTag", 0, 300, NowIce.DrawTag(4), NowIce.Height, "linear", 4)
    End If
    
    If NowIce.DrawTag(4) <> 4 Then NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, NowIce.Size.Height - NowIce.DrawTag(4)), NewSize(NowIce.Width, NowIce.DrawTag(4)), NowIce.Style.Radian, NowIce.Style.Color(Back)
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(NowIce.Value / NowIce.Max * NowIce.Size.Width * 0.7, NowIce.Size.Height - NowIce.DrawTag(4)), NewSize(NowIce.Size.Width * 0.3, NowIce.DrawTag(4)), NowIce.Style.Radian, NowIce.Style.Color(Fore)
    If NowIce.ClickState = MouseDown Then
        If NowIce.DrawTag(5) = "" Then NowIce.DrawTag(5) = NowIce.ClickPos.x - (NowIce.Value / NowIce.Max * NowIce.Size.Width * 0.7)
        NowIce.Value = (NowIce.ClickPos.x - Val(NowIce.DrawTag(5))) / (NowIce.Size.Width * 0.7) * NowIce.Max
    End If
    If NowIce.ClickState = MouseUp Then
        NowIce.DrawTag(5) = ""
    End If
    
    If NowIce.ClickState = MouseLeave Then
        NowIce.DrawAnimate.AddByReset NewAnimate(NowIce, "DrawTag", 0, 300, NowIce.DrawTag(4), 4, "linear", 4)
    End If
End Sub
Sub DrawLabel()
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Back)
    NowIce.Draw.DrawRect NewPos(0, 0), NowIce.Size, NowIce.Style.Color(Border)
    NowIce.Draw.DrawEImage NewPos(0, 0), NowIce.Size, NowIce.Style.ImgAlign, NowIce.Src
    NowIce.Draw.DrawString NewPos(0, 0), NowIce.Size, NowIce.Text, NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
End Sub
Sub DrawButton()

    If NowIce.ClickState = MouseEnter Then
        NowIce.DrawAnimate.AddByReset NewAnimate(NowIce.AnimateColor, "Color", 0, 400, NowIce.AnimateColor.Color, NowIce.Style.Color(Active).Color, "linearcolor")
        'NowIce.DrawAnimate(1) = NewAnimate(NowIce, "DrawTag", 0, 400, NowIce.DrawTag(4), 0, "linearcolor", 4)
    End If
    
    If NowIce.ClickState = None And NowIce.DrawAnimate.Count = 0 And NowIce.AnimateColor.Color <> NowIce.Style.Color(Back).Color Then
        NowIce.AnimateColor.Color = NowIce.Style.Color(Back).Color
    End If
        
    If NowIce.ClickState = MouseDown Then
        NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Border2)
    Else
        NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.AnimateColor
    End If
    If NowIce.ClickState = MouseUp Then
        NowIce.DrawTag(5) = NowIce.ClickPos.x: NowIce.DrawTag(4) = NowIce.ClickPos.y: NowIce.DrawTag(3) = GetTickCount
        If NowIce.Style.Color(Active).Light <= 0.5 Then
            NowIce.AnimateColor2.Color = argb(220, 255, 255, 255)
        Else
            NowIce.AnimateColor2.Color = argb(220, 0, 0, 0)
        End If
        NowIce.DrawAnimate.AddByReset NewAnimate(NowIce.AnimateColor2, "a", 0, 800, 220, 0, "linear")
    End If
    If GetTickCount - Val(NowIce.DrawTag(3)) <= 800 Then
        Dim Progress As Single, CircleSize As Long
        Progress = (GetTickCount - Val(NowIce.DrawTag(3))) / 800
        NowIce.Draw.ClipShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian
        CircleSize = IIf(NowIce.Width > NowIce.Height, NowIce.Width, NowIce.Height) * 2 * Progress
        NowIce.Draw.FillEllipse NewPos(Val(NowIce.DrawTag(5)) - CircleSize / 2, Val(NowIce.DrawTag(4)) - CircleSize / 2), NewSize(CircleSize, CircleSize), NowIce.AnimateColor2
        NowIce.Draw.ResetClip
    End If
    NowIce.Draw.DrawShape NowIce.Style.Shape, NewPos(0, 0), NowIce.Size, NowIce.Style.Radian, NowIce.Style.Color(Border)
    NowIce.Draw.DrawEImage NewPos(0, 0), NowIce.Size, NowIce.Style.ImgAlign, NowIce.Src
    NowIce.Draw.DrawString NewPos(0, 0), NowIce.Size, NowIce.Text, NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
        
    If NowIce.ClickState = MouseLeave Then
        NowIce.DrawAnimate.AddByReset NewAnimate(NowIce.AnimateColor, "Color", 0, 400, NowIce.AnimateColor.Color, NowIce.Style.Color(Back).Color, "linearcolor")
        'NowIce.DrawAnimate(1) = NewAnimate(NowIce, "DrawTag", 0, 400, NowIce.DrawTag(4), NowIce.Style.Color(Border).a, "linearcolor", 4)
    End If
End Sub
Sub DrawOption()
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NewSize(NowIce.Size.Height, NowIce.Size.Height), NowIce.Style.Radian, NowIce.Style.Color(Back)
    
    If NowIce.IsOn = True Then
        NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(NowIce.Size.Height * 0.25, NowIce.Size.Height * 0.25), NewSize(NowIce.Size.Height * 0.5, NowIce.Size.Height * 0.5), NowIce.Style.Radian, NowIce.Style.Color(Fore)
    End If
    
    NowIce.Draw.DrawShape NowIce.Style.Shape, NewPos(0, 0), NewSize(NowIce.Size.Height, NowIce.Size.Height), NowIce.Style.Radian, IIf(NowIce.IsOn = False, NowIce.Style.Color(Border), NowIce.Style.Color(Border2))
    
    NowIce.Draw.DrawString NewPos(NowIce.Size.Height + 5, 0), NewSize(NowIce.Size.Width - NowIce.Size.Height - 5, NowIce.Size.Height), NowIce.Text, NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
    If NowIce.ClickState = MouseUp Then NowIce.IsOn = Not NowIce.IsOn
End Sub
Sub DrawCheck()
    NowIce.Draw.FillShape NowIce.Style.Shape, NewPos(0, 0), NewSize(NowIce.Size.Height, NowIce.Size.Height), NowIce.Style.Radian, NowIce.Style.Color(Back)
    
    If NowIce.IsOn = True Then
        'Draw √
        NowIce.Draw.DrawLine NewPos(NowIce.Size.Height / 4 - 0.5, NowIce.Size.Height / 2 - 1), _
                                          NewPos(NowIce.Size.Height / 2 - 0.5, NowIce.Size.Height / 4 * 3 - 1), NowIce.Style.Color(Border2)
        NowIce.Draw.DrawLine NewPos(NowIce.Size.Height / 4 * 3 - 0.5, NowIce.Size.Height / 4 - 1), _
                                          NewPos(NowIce.Size.Height / 2 - NowIce.Style.Color(Border2).Width + 1 - 0.5, NowIce.Size.Height / 4 * 3 - 1), _
                                          NowIce.Style.Color(Border2)
    End If
    
    NowIce.Draw.DrawShape NowIce.Style.Shape, NewPos(0, 0), NewSize(NowIce.Size.Height, NowIce.Size.Height), NowIce.Style.Radian, NowIce.Style.Color(Border)
    
    NowIce.Draw.DrawString NewPos(NowIce.Size.Height + 5, 0), NewSize(NowIce.Size.Width - NowIce.Size.Height - 5, NowIce.Size.Height), NowIce.Text, NowIce.Style.Align, NowIce.Style.Font, NowIce.Style.Color(Fore)
    If NowIce.ClickState = MouseUp Then NowIce.IsOn = Not NowIce.IsOn
End Sub
Sub DrawTrueIcelolly()
'嗯？ 你似乎想要看透这个方法？？

If EverIcelolly = True Then Exit Sub

    Randomize
    NowIce.AnimateColor.Color = argb(80, 255, 0, 0)
    With NowIce.Draw
        .AddArc NewPos(0, 0), NewSize(NowIce.Width / 2, NowIce.Height / 2), -180, 180
        .AddArc NewPos(NowIce.Width / 2, 0), NewSize(NowIce.Width / 2, NowIce.Height / 2), -180, 180
        .AddLine NewPos(NowIce.Width, NowIce.Height / 3), NewPos(NowIce.Width / 2, NowIce.Height)
        .AddLine NewPos(NowIce.Width / 2, NowIce.Height), NewPos(0, NowIce.Height / 3)
        .FillPath NowIce.AnimateColor
    End With
    If NowIce.ClickState = MouseUp Then
        Dim Target As Long, Red As Long, NoRed As Boolean
        LastClickTime = GetTickCount
        Target = Rnd * (NowIce.ParentLayout.Width - NowIce.Width * 2) + NowIce.Width
        Red = IIf(Target < NowIce.ParentLayout.Width / 2, 1, 2)
        On Error Resume Next
        For i = 1 To UBound(EMembers)
            If TypeName(EMembers(i)) = "ColorMix" Then
                If EMembers(i).R + Red <= 255 And EMembers(i).R + Red >= 0 Then EMembers(i).R = EMembers(i).R + Red
                If EMembers(i).a - Red <= 255 And EMembers(i).a - Red >= 30 Then EMembers(i).a = EMembers(i).a - Red
                If EMembers(i).R <> 255 And (Not EMembers(i) Is NowIce.ParentLayout.ParentUI.MouseColor) Then NoRed = True
            End If
        Next
        Beep NowIce.x * 2, 100
        If NoRed = False Then
            EverIcelolly = True
            NowIce.Class = "Button"
            For i = 1 To IceLollyCount
                GetIceLollyMember(i).Text = "火热的心"
            Next
            NowIce.ParentLayout.ParentUI.Refresh
            For i = 1000 To 6000 Step 100
                Beep i, 10
                Sleep 10: DoEvents
            Next
            Exit Sub
        End If
        NowIce.Animate.AddByReset NewAnimate(NowIce, "X", 0, 300, NowIce.x, Target)
    End If
    If LastClickTime <> 0 Then
        If GetTickCount - LastClickTime >= 1500 Then
            On Error Resume Next
            For i = 1 To UBound(EMembers)
                If TypeName(EMembers(i)) = "ColorMix" Then
                    EMembers(i).R = 0
                End If
            Next
            EverIcelolly = True
            NowIce.Class = "Button"
            For i = 1 To IceLollyCount
                GetIceLollyMember(i).Text = "寒冷的..."
            Next
            NowIce.ParentLayout.ParentUI.Refresh
            For i = 6000 To 1000 Step -100
                Beep i, 10
                Sleep 10: DoEvents
            Next
        End If
    End If
    Err.Clear
End Sub
