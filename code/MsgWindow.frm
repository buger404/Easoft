VERSION 5.00
Begin VB.Form MsgWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "MessageBox"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   Icon            =   "MsgWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4800
      Top             =   150
   End
End
Attribute VB_Name = "MsgWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UI Engine
Public EUI As New EUI
'Styles
Public CaptionStyle As New StyleBox, ButtonStyle As New StyleBox
Dim TitleStyle As New StyleBox, ContextStyle As New StyleBox
'Controls
Public TitleText As Icelolly, Context As Icelolly, YesBtn As Icelolly, NoBtn As Icelolly
'Other
Public Choice As Integer
Private Sub DrawTimer_Timer()
    EUI.Display
End Sub

Private Sub Form_Load()
    'Create UI & set shadow
    EUI.Create Me.Hwnd
    SetWindowShadow Me
    
    'Set the style
    With CaptionStyle
        .Align = NewAlign(OnCenter, OnMiddle)
    End With
    With TitleStyle
        .Align = NewAlign(OnLeft, OnTop)
        .Font.Size = 20
    End With
    With ContextStyle
        .Align = NewAlign(OnLeft, OnTop)
        .Color(Fore).Color = argb(255, 127, 127, 127)
        .Font.Size = 14
    End With
    
    'Create caption aera
    EUI.CreateLayout "Caption_Aera", NewPos(0, 0), NewSize(Me.ScaleWidth, 1), Me
    EUI.MakeIcelolly "Line", "", "Caption_Aera", NewPos(0, 0), NewSize(Me.ScaleWidth, 1), CaptionStyle
    NowIce.Pos2 = NewPos(0 + Me.ScaleWidth, 0)
    
    'Create user aera
    EUI.CreateLayout "User_Aera", NewPos(20, 20), NewSize(Me.ScaleWidth - 40, Me.ScaleHeight - 40), Me
    EUI.MakeIcelolly "Label", "", "User_Aera", NewPos(0, 0), NewSize(Me.ScaleWidth - 40, 30), TitleStyle 'Index =1
    Set TitleText = NowIce
    EUI.MakeIcelolly "Label", "", "User_Aera", NewPos(0, 30), NewSize(Me.ScaleWidth - 40, Me.ScaleHeight - 40 - 30 - 40), ContextStyle 'Index =2
    Set Context = NowIce
    EUI.MakeIcelolly "Button", "YesBtn", "User_Aera", NewPos(Me.ScaleWidth - 40 - 90 * 2, Me.ScaleHeight - 30 - 40), NewSize(80, 30), ButtonStyle 'Index =3
    NowIce.BlockClick = True: Set YesBtn = NowIce: NowIce.Text = "Yes"
    EUI.MakeIcelolly "Button", "NoBtn", "User_Aera", NewPos(Me.ScaleWidth - 40 - 90, Me.ScaleHeight - 30 - 40), NewSize(80, 30), ButtonStyle 'Index =3
    NowIce.BlockClick = True: Set NoBtn = NowIce: NowIce.Text = "No"
End Sub

Sub YesBtn_Update()
    If NowIce.ClickState = MouseUp Then
        Choice = 1
    End If
End Sub

Sub NoBtn_Update()
    If NowIce.ClickState = MouseUp Then
        Choice = 0
    End If
End Sub
