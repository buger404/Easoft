Attribute VB_Name = "EasoftStyle"
'***********************************************************************************************************
' Easoft Normal Style For Easoft
' Version : 1.1, 2018/9/8
' Designer : Error 404
' Maker : Error 404
'***********************************************************************************************************
Public E_Button As New StyleBox, E_Option As New StyleBox, E_Check As New StyleBox, E_Label As New StyleBox
Public E_Progress As New StyleBox, E_ArcProgress As New StyleBox
Public E_Scroll As New StyleBox, E_Slider As New StyleBox, E_Loading As New StyleBox, E_List As New StyleBox
Public Sub SetEasoftStyle(UI As EUI)
    UI.BackColor.Color = argb(255, 255, 255, 255)
    '设置各种控件的样式
     With E_Loading
        .Color(Border).Color = argb(255, 27, 191, 201)
        .Color(Border).Width = 3
    End With
    With E_List
        .Color(Back).Color = argb(255, 255, 255, 255)
        .Color(Border2).Color = argb(255, 242, 242, 242)
        .Color(Active).Color = argb(130, 27, 191, 201)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Align = NewAlign(OnLeft, OnMiddle)
    End With
    With E_Label
        .Align = NewAlign(OnLeft, OnLeft)
        .Color(Fore).Color = argb(255, 255, 255, 255)
        .Font.Size = 14
    End With
    With E_Button
        .Align = NewAlign(OnCenter, OnMiddle)
        .Color(Back).Color = argb(0, 27, 191, 201)
        .Color(Border).Color = argb(255, 232, 232, 232)
        .Color(Border2).Color = argb(255, 27, 191, 201)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Color(Active).Color = argb(150, 27, 191, 201)
        .Shape = EShape.RoundRect
        .Radian = 999
    End With
    With E_Option
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 255, 255, 255)
        .Color(Border).Color = argb(255, 222, 222, 222)
        .Color(Border2).Color = argb(255, 136, 221, 227)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Shape = Oval
    End With
    With E_Check
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 27, 191, 201)
        .Color(Border2).Color = argb(255, 255, 255, 255)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Shape = EShape.RoundRect
        .Radian = 6
    End With
    With E_Progress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 242, 242, 242)
        .Color(Fore).Color = argb(255, 27, 191, 201)
        .Shape = EShape.RoundRect
        .Radian = 60
    End With
     With E_ArcProgress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Fore).Color = argb(140, 255, 255, 255)
        .Color(Border).Color = argb(255, 27, 191, 201)
        .Color(Border).Width = 2
        .Shape = Square
        .Radian = 6
    End With
     With E_Slider
        .Color(Border).Color = argb(255, 242, 242, 242)
        .Color(Border).Width = 2
        .Color(Border2).Color = argb(255, 27, 191, 201)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 27, 191, 201)
        .Shape = Oval
    End With
    With E_Scroll
        .Color(Back).Color = argb(255, 242, 242, 242)
        .Color(Fore).Color = argb(120, 27, 191, 201)
    End With
    UI.Refresh
End Sub

