VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   711
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10050
      Top             =   900
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   10050
      Top             =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EUI As New EUI
Dim Test_Button As New StyleBox, Test_Option As New StyleBox, Test_Check As New StyleBox, Test_Progress As New StyleBox, Test_ArcProgress As New StyleBox
Dim Test_Scroll As New StyleBox, Test_Slider As New StyleBox, Test_List As New StyleBox, Test_Loading As New StyleBox
Dim Test_Edit As New StyleBox
Dim MyImg As New ImageCollection
Dim FPS As Long
Private Sub Form_Load()
    EasoftPower True
    
    '创建
    EUI.Create Me.hWnd
    EUI.BackColor.Color = argb(240, 255, 255, 255)
    EUI.Refresh

    'Aero
    'SetEWindow Me.hWnd, AeroWindow: SetWindowShadow Me
    
    '创建一个控件集合
    EUI.CreateLayout "Layout1", NewPos(40, 40), NewSize(500, 250), Me
    
    '创建一个样式
    Test_Button.Align = NewAlign(OnCenter, OnMiddle)
    Test_Button.Color(Back).Color = argb(0, 27, 191, 201)
    Test_Button.Color(Border).Color = argb(255, 232, 232, 232)
    Test_Button.Color(Border2).Color = argb(255, 27, 191, 201)
    Test_Button.Color(Fore).Color = argb(255, 27, 27, 27)
    Test_Button.Color(Active).Color = argb(150, 27, 191, 201)
    Test_Button.Shape = EShape.RoundRect
    Test_Button.Radian = 50
    
    '做冰棍吃
    EUI.MakeIcelolly "Button", "Button1", "Layout1", NewPos(0, 0), NewSize(90, 30), Test_Button
    this.Text = "Page1": this.BlockClick = True
    
    EUI.MakeIcelolly "Button", "", "Layout1", NewPos(110, 0), NewSize(90, 30), Test_Button
    this.Text = "Page2": this.BlockClick = True
    
    EUI.MakeIcelolly "Button", "Label1", "Layout1", NewPos(220, 0), NewSize(90, 30), Test_Button
    this.Text = "Page3": this.BlockClick = True
    
    '创建一个样式
    Test_Option.Align = NewAlign(OnLeft, OnMiddle)
    Test_Option.Color(Back).Color = argb(255, 255, 255, 255)
    Test_Option.Color(Border).Color = argb(255, 222, 222, 222)
    Test_Option.Color(Border2).Color = argb(255, 136, 221, 227)
    Test_Option.Color(Fore).Color = argb(255, 27, 27, 27)
    Test_Option.Shape = Oval
    
    '创建一个样式
    With Test_Check
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 27, 191, 201)
        .Color(Border2).Color = argb(255, 255, 255, 255)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Shape = EShape.RoundRect
        .Radian = 6
    End With
    
    '创建一个样式
    With Test_Progress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 242, 242, 242)
        .Color(Fore).Color = argb(255, 27, 191, 201)
        .Shape = EShape.RoundRect
        .Radian = 60
    End With
     With Test_ArcProgress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Fore).Color = argb(140, 255, 255, 255)
        .Color(Border).Color = argb(255, 27, 191, 201)
        .Color(Border).Width = 2
        .Shape = Square
        .Radian = 6
    End With
     With Test_Slider
        .Color(Border).Color = argb(255, 242, 242, 242)
        .Color(Border).Width = 2
        .Color(Border2).Color = argb(255, 27, 191, 201)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 27, 191, 201)
        .Shape = Oval
    End With
     With Test_Loading
        .Color(Border).Color = argb(255, 27, 191, 201)
        .Color(Border).Width = 3
    End With
     With Test_Edit
        .Color(Back).Color = argb(255, 242, 242, 242)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Color(Border2).Color = argb(255, 27, 27, 27)
        .Color(Active).Color = argb(110, 27, 191, 201)
        .Align = NewAlign(OnLeft, OnTop)
    End With
    
    '创建一个控件集合
    EUI.CreateLayout "Page1", NewPos(40, 90), NewSize(500, 350), Me
    
    EUI.MakeIcelolly "Progress", "Progress1", "Page1", NewPos(5, 0), NewSize(300, 18), Test_Progress
    this.Text = "": this.BlockClick = True: this.Max = 1000
    EUI.MakeIcelolly "Slider", "", "Page1", NewPos(5, 27), NewSize(300, 18), Test_Slider
    this.Text = "": this.BlockClick = True
    EUI.MakeIcelolly "ArcProgress", "Progress1", "Page1", NewPos(245, 28 + 27), NewSize(60, 60), Test_ArcProgress
    this.Text = "": this.BlockClick = True: this.Src.Path = App.Path & "\404.png": this.Src.Size = this.Size: this.Src.ClipCircle: this.Max = 1000
    EUI.MakeIcelolly "Loading", "", "Page1", NewPos(245 - 70, 28 + 27), NewSize(60, 60), Test_Loading
    EUI.MakeIcelolly "Check", "", "Page1", NewPos(5, 27 * 2), NewSize(100, 18), Test_Check
    this.Text = "Check": this.BlockClick = True
    EUI.MakeIcelolly "Option", "", "Page1", NewPos(5, 27 * 3), NewSize(100, 18), Test_Option
    this.Text = "Option": this.BlockClick = True
    EUI.MakeIcelolly "Button", "Page1Button", "Page1", NewPos(300 / 2 - 90 / 2 + 5 - 60, 27 * 5), NewSize(90, 30), Test_Button
    this.Text = "-1": this.BlockClick = True
    EUI.MakeIcelolly "Button", "Page1Button2", "Page1", NewPos(300 / 2 - 90 / 2 + 5 + 60, 27 * 5), NewSize(110, 30), Test_Button
    this.Text = "添加列表项": this.BlockClick = True
    
    EUI.MakeIcelolly "Edit", "Edit1", "Page1", NewPos(0, 27 * 5 + 40), NewSize(250, 350 - (27 * 5 + 40)), Test_Edit
    this.BlockClick = True: this.Multi = True
    this.Text = "用于测试asdgsadghsahshdahdfh" & vbCrLf & "sdhdshdfhdfhdfjdfjdfjf" & vbCrLf & "weteysyfdhdfhdfhdfhdfhg" & vbCrLf & "dfhdfhdfhdfhg"
    
    EUI.MakeIcelolly "VScroll", "", "Page1", NewPos(240, 27 * 5 + 40), NewSize(10, 350 - (27 * 5 + 40)), Test_Scroll
    this.BlockClick = True: Set EUI.LayoutByID("Page1").IcelollyByID("Edit1").LinkWheel = this
    Set this.LinkWheel = EUI.LayoutByID("Page1").IcelollyByID("Edit1")

    '创建一个控件集合
    EUI.CreateLayout "Layout2", NewPos(0, Me.ScaleHeight - 15), NewSize(Me.ScaleWidth, 15), Me
    
    '创建一个样式
    With Test_Scroll
        .Color(Back).Color = argb(255, 242, 242, 242)
        .Color(Fore).Color = argb(120, 27, 191, 201)
    End With
    
    EUI.MakeIcelolly "HScroll", "Scroll1", "Layout2", NewPos(0, 0), NewSize(Me.ScaleWidth, 15), Test_Scroll
    this.BlockClick = True
    
    '创建一个控件集合
    EUI.CreateLayout "Layout3", NewPos(Me.ScaleWidth - 300, Me.ScaleHeight - 58), NewSize(280, 18), Me
    
    EUI.MakeIcelolly "Check", "StyleCheck", "Layout3", NewPos(0, 0), NewSize(60, 18), Test_Check
    this.Text = "默认": this.BlockClick = True: this.Tag(0) = 0: this.IsOn = True
    EUI.MakeIcelolly "Check", "StyleCheck", "Layout3", NewPos(80, 0), NewSize(80, 18), Test_Check
    this.Text = "冰棍UI": this.BlockClick = True: this.Tag(0) = 1
    EUI.MakeIcelolly "Check", "StyleCheck", "Layout3", NewPos(180, 0), NewSize(80, 18), Test_Check
    this.Text = "Infinity": this.BlockClick = True: this.Tag(0) = 2
    
    'EUI.LayoutByID("Layout3").IcelollyByID("StyleCheck").Multi = True
    
    '创建一个控件集合
    EUI.CreateLayout "Layout5", NewPos(Me.ScaleWidth - 15, 0), NewSize(15, Me.ScaleHeight - 15), Me
    
    EUI.MakeIcelolly "VScroll", "Scroll2", "Layout5", NewPos(0, 0), NewSize(15, Me.ScaleHeight - 15), Test_Scroll
    this.BlockClick = True
    
    EUI.CreateLayout "Layout6", NewPos(Me.ScaleWidth - 315, 40), NewSize(290, Me.ScaleHeight - 150), Me
    
    With Test_List
        .Color(Back).Color = argb(255, 255, 255, 255)
        .Color(Border2).Color = argb(255, 242, 242, 242)
        .Color(Active).Color = argb(130, 27, 191, 201)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Align = NewAlign(OnLeft, OnMiddle)
    End With
    
    EUI.MakeIcelolly "List", "List1", "Layout6", NewPos(0, 0), NewSize(290, Me.ScaleHeight - 150), Test_List
    MyImg.AddDir App.Path & "\assets\"
    this.BlockClick = True
    
    EUI.MakeIcelolly "VScroll", "", "Layout6", NewPos(280, 0), NewSize(10, Me.ScaleHeight - 150), Test_Scroll
    this.BlockClick = True: Set EUI.LayoutByID("Layout6").IcelollyByID("List1").LinkWheel = this
    Set this.LinkWheel = EUI.LayoutByID("Layout6").IcelollyByID("List1")
    
    Timer1.Enabled = True
    
    'debugwin.Show
End Sub
Sub SetNormalStyle()
    EUI.BackColor.Color = argb(255, 255, 255, 255)
    With Test_List
        .Color(Back).Color = argb(255, 255, 255, 255)
        .Color(Border2).Color = argb(255, 242, 242, 242)
        .Color(Active).Color = argb(130, 27, 191, 201)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Align = NewAlign(OnLeft, OnMiddle)
    End With
     With Test_Slider
        .Color(Border).Color = argb(255, 242, 242, 242)
        .Color(Border).Width = 2
        .Color(Border2).Color = argb(255, 27, 191, 201)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 27, 191, 201)
        .Shape = Oval
    End With
    With Test_Progress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 242, 242, 242)
        .Color(Fore).Color = argb(255, 27, 191, 201)
        .Shape = EShape.RoundRect
        .Radian = 60
    End With
     With Test_ArcProgress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Border).Color = argb(255, 27, 191, 201)
        .Color(Border).Width = 2
        .Shape = Square
        .Radian = 6
    End With
    With Test_Option
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 255, 255, 255)
        .Color(Border).Color = argb(255, 222, 222, 222)
        .Color(Border2).Color = argb(255, 136, 221, 227)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Shape = Oval
    End With
    With Test_Button
        .Align = NewAlign(OnCenter, OnMiddle)
        .Color(Back).Color = argb(0, 27, 191, 201)
        .Color(Border).Color = argb(255, 232, 232, 232)
        .Color(Border2).Color = argb(255, 27, 191, 201)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Color(Active).Color = argb(150, 27, 191, 201)
        .Shape = EShape.RoundRect
        .Radian = 50
    End With
    With Test_Check
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 27, 191, 201)
        .Color(Border).Color = argb(0, 51, 51, 51)
        .Color(Border2).Color = argb(255, 255, 255, 255)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Shape = EShape.RoundRect
        .Radian = 6
    End With
    With Test_Scroll
        .Color(Back).Color = argb(120, 27, 191, 201)
        .Color(Fore).Color = argb(255, 27, 191, 201)
    End With
    EUI.Refresh
End Sub
Sub SetIcelollyStyle()
    EUI.BackColor.Color = argb(255, 45, 45, 48)
    With Test_List
        .Color(Back).Color = argb(255, 45, 45, 48)
        .Color(Border2).Color = argb(255, 62, 62, 64)
        .Color(Active).Color = argb(130, 0, 122, 204)
        .Color(Fore).Color = argb(255, 255, 255, 255)
        .Align = NewAlign(OnLeft, OnMiddle)
    End With
     With Test_Slider
        .Color(Border).Color = argb(255, 62, 62, 64)
        .Color(Border).Width = 2
        .Color(Border2).Color = argb(255, 0, 122, 204)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 0, 122, 204)
        .Shape = Oval
    End With
    With Test_Progress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 62, 62, 64)
        .Color(Fore).Color = argb(255, 0, 122, 204)
        .Shape = Square
    End With
     With Test_ArcProgress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Border).Color = argb(255, 0, 122, 204)
        .Color(Border).Width = 2
    End With
    With Test_Option
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 45, 45, 48)
        .Color(Border).Color = argb(255, 242, 242, 242)
        .Color(Border2).Color = argb(255, 0, 122, 204)
        .Color(Fore).Color = argb(255, 255, 255, 255)
        .Shape = Oval
    End With
    With Test_Button
        .Align = NewAlign(OnCenter, OnMiddle)
        .Color(Back).Color = argb(255, 45, 45, 48)
        .Color(Border).Color = 0
        .Color(Border2).Color = argb(255, 0, 122, 204)
        .Color(Fore).Color = argb(255, 255, 255, 255)
        .Color(Active).Color = argb(255, 62, 62, 64)
        .Shape = Square
    End With
    With Test_Check
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 45, 45, 48)
        .Color(Border).Color = argb(255, 255, 255, 255)
        .Color(Border2).Color = argb(255, 242, 242, 242)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 255, 255, 255)
        .Shape = Square
    End With
    With Test_Scroll
        .Color(Back).Color = argb(255, 62, 62, 66)
        .Color(Fore).Color = argb(255, 104, 104, 104)
    End With
    EUI.Refresh
End Sub
Sub SetInStyle()
    EUI.BackColor.Color = argb(255, 255, 255, 255)
    With Test_List
        .Color(Back).Color = argb(255, 255, 255, 255)
        .Color(Border2).Color = argb(255, 242, 242, 242)
        .Color(Active).Color = argb(130, 161, 169, 178)
        .Color(Fore).Color = argb(255, 85, 85, 85)
        .Align = NewAlign(OnLeft, OnMiddle)
    End With
     With Test_Slider
        .Color(Border).Color = argb(255, 232, 232, 232)
        .Color(Border).Width = 2
        .Color(Border2).Color = argb(255, 0, 222, 121)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 0, 222, 121)
        .Shape = Oval
    End With
    With Test_Progress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 232, 232, 232)
        .Color(Fore).Color = argb(255, 0, 222, 121)
        .Shape = EShape.RoundRect
        .Radian = 6
    End With
     With Test_ArcProgress
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Border).Color = argb(255, 0, 222, 121)
        .Color(Border).Width = 2
    End With
    With Test_Option
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 255, 255, 255)
        .Color(Border).Color = argb(255, 232, 232, 232)
        .Color(Border2).Color = argb(255, 0, 222, 121)
        .Color(Fore).Color = argb(255, 102, 102, 102)
        .Shape = Oval
    End With
    With Test_Button
        .Align = NewAlign(OnCenter, OnMiddle)
        .Color(Back).Color = argb(255, 255, 255, 255)
        .Color(Border).Color = argb(0, 232, 232, 232)
        .Color(Border2).Color = argb(255, 232, 232, 232)
        .Color(Fore).Color = argb(255, 85, 85, 85)
        .Color(Active).Color = argb(255, 238, 238, 238)
        .Shape = EShape.RoundRect
        .Radian = 10
    End With
    With Test_Check
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Back).Color = argb(255, 0, 222, 121)
        .Color(Border).Color = argb(0, 255, 255, 255)
        .Color(Border2).Color = argb(255, 255, 255, 255)
        .Color(Border2).Width = 2
        .Color(Fore).Color = argb(255, 102, 102, 102)
        .Shape = EShape.RoundRect
        .Radian = 6
    End With
    With Test_Scroll
        .Color(Back).Color = argb(255, 234, 234, 234)
        .Color(Fore).Color = argb(255, 161, 169, 178)
    End With
    EUI.Refresh
End Sub
Public Sub StyleCheck_Update()
    If this.ClickState = MouseUp Then
        Select Case Val(this.Tag(0))
            Case 0
                Call SetNormalStyle
            Case 1
                Call SetIcelollyStyle
            Case 2
                Call SetInStyle
        End Select
    End If
End Sub
Public Sub Button1_Update()

End Sub
Public Sub Page1Button_Update()
    If this.ClickState = MouseUp Then
        With EUI.LayoutByID("Page1").IcelollyByID("Progress1", 0)
            .Value = .Value - 100
        End With
        With EUI.LayoutByID("Page1").IcelollyByID("Progress1", 1)
            .Value = .Value - 100
        End With
    End If
End Sub
Public Sub Page1Button2_Update()
    If this.ClickState = MouseUp Then
        Randomize
        Select Case Int(Rnd * 3 + 1)
            Case 1
            With EUI.LayoutByID("Layout6").IcelollyByID("List1").List
                .Add "404" & IIf(Int(Rnd * 2) = 0, "~", "!")
                .Src(.Count) = MyImg.ImageByIndex(1)
            End With
            Case 2
            With EUI.LayoutByID("Layout6").IcelollyByID("List1").List
                .Add "棍棍" & IIf(Int(Rnd * 2) = 0, "~~", "!~")
                .Src(.Count) = MyImg.ImageByIndex(2)
            End With
            Case Else
            With EUI.LayoutByID("Layout6").IcelollyByID("List1").List
                .Add "黑嘴" & IIf(Int(Rnd * 2) = 0, "~", "!?")
                .Src(.Count) = MyImg.ImageByIndex(3)
            End With
        End Select
    End If
End Sub
Public Sub Label1_Update()
    If this.ClickState = MouseUp Then
        this.Class = "Button"
        this.Animate.Add NewAnimate(this, "X", 0, 1000, this.Pos.x, this.Pos.x + 100)
    End If
End Sub
Public Sub Scroll1_Update()
    If this.ClickState = MouseDown Then EUI.LayoutByID("Layout1").x = -(this.Value / 100 * (EUI.LayoutByID("Layout1").Size.Width + 40)) + 40
End Sub
Public Sub Scroll2_OnScroll()
    EUI.LayoutByID("Layout1").IcelollyByID("Label1").Text = this.Value
    EUI.LayoutByID("Layout1").y = 40 + -this.Value * 3
    EUI.LayoutByID("Page1").y = 90 + -this.Value * 3
End Sub
Public Sub List1_Update()
    If this.ClickState = MouseUp Then
        Easoft.MsgBox "吼吼吼，我来测试信息框啦，嗯...你选中的是：" & this.List.Items(this.List.Index), "我是邪恶的标题", EColors.EGreen, EUI.BackColor.Color, Test_Button
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    EasoftPower False
End Sub

Private Sub Form_Resize()
With EUI.LayoutByID("Layout2")
    .y = Me.ScaleHeight - 15
    .Width = Me.ScaleWidth
    With .IcelollyByID("Scroll1")
        .Width = Me.ScaleWidth
    End With
End With
With EUI.LayoutByID("Layout5")
    .x = Me.ScaleWidth - 15
    .Height = Me.ScaleHeight - 15
    With .IcelollyByID("Scroll2")
        .Height = Me.ScaleHeight - 15
    End With
End With
With EUI.LayoutByID("Layout3")
    .y = Me.ScaleHeight - 58
    .x = Me.ScaleWidth - 300
End With
End Sub

Private Sub Timer1_Timer()
    With EUI.LayoutByID("Page1").IcelollyByID("Progress1", 0)
        .Value = .Value + 1
    End With
    With EUI.LayoutByID("Page1").IcelollyByID("Progress1", 1)
        .Value = .Value + 1
    End With
    EUI.Display
    FPS = FPS + 1
End Sub

Private Sub Timer2_Timer()
Me.Caption = "Easoft Example , " & FPS & " fps ."
FPS = 0
End Sub
