VERSION 5.00
Begin VB.Form DebugWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "DebugWindow"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6300
      Top             =   150
   End
End
Attribute VB_Name = "DebugWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UI Engine
Dim EUI As New EUI
'Controls
Dim InfoText As Icelolly
'Styles
Dim Title_Text As New StyleBox, Info_Text As New StyleBox

Private Sub DrawTimer_Timer()
    If Me.Visible = False Then Exit Sub
    Dim Result As String
    If NowLayout Is Nothing Then
        Result = "None"
    Else
        Result = NowLayout.Id & "(0x" & Hex(ObjPtr(NowLayout)) & ")" & " , Events : " & "0x" & Hex(ObjPtr(NowLayout.EventBox)) & _
                    vbCrLf & "UI : 0x" & Hex(ObjPtr(NowLayout.ParentUI)) & _
                    vbCrLf & vbCrLf
        If NowLayout.ParentUI.FocusIcelolly Is Nothing Then
            Result = Result & "Icelolly : no focus ," & vbCrLf
        Else
            Result = Result & "Icelolly : focus on " & IIf(NowLayout.ParentUI.FocusIcelolly.Id = "", "[No ID]", NowLayout.ParentUI.FocusIcelolly.Id) & "(0x" & Hex(ObjPtr(NowLayout.ParentUI.FocusIcelolly)) & ") ," & _
                        vbCrLf & NowLayout.ParentUI.FocusIcelolly.ErrorInfo & vbCrLf
        End If
        If NowLayout.MouseInIcelolly Is Nothing Then
            Result = Result & "no mouse in" & vbCrLf
        Else
            Result = Result & "mouse in " & IIf(NowLayout.MouseInIcelolly.Id = "", "[No ID]", NowLayout.MouseInIcelolly.Id) & "(0x" & Hex(ObjPtr(NowLayout.MouseInIcelolly)) & ") ," & _
                        vbCrLf & NowLayout.MouseInIcelolly.ErrorInfo & vbCrLf
        End If
    End If
    InfoText.Text = Result
    EUI.Display
End Sub

Private Sub Form_Load()
    '×°ÔØ
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    'Create
    EUI.Create Me.Hwnd
    SetEWindow Me.Hwnd, AeroWindow
    SetWindowShadow Me
    EUI.BackColor.Color = argb(220, 255, 255, 255)
    
    'Create Styles
    With Title_Text
        .Align = NewAlign(OnLeft, OnMiddle)
        .Color(Fore).Color = argb(255, 27, 191, 201)
        .Font.Size = 20
    End With
    With Info_Text
        .Align = NewAlign(OnLeft, OnTop)
        .Color(Fore).Color = argb(255, 27, 27, 27)
        .Font.Size = 14
    End With
    
    'Whole
    EUI.CreateLayout "MainLayout", NewPos(20, 20), NewSize(Me.ScaleWidth - 40, Me.ScaleHeight - 40), Me
    EUI.MakeIcelolly "Label", "DebugBar", "MainLayout", NewPos(0, 0), NewSize(Me.ScaleWidth - 40, 25), Title_Text
    NowIce.Text = "Debug": NowIce.BlockClick = True
    EUI.MakeIcelolly "Label", "", "MainLayout", NewPos(0, 25), NewSize(Me.ScaleWidth - 40, Me.ScaleHeight - 40 - 25), Info_Text
    Set InfoText = NowIce
    
    EUI.Refresh
    
    DrawTimer.Enabled = True
End Sub
Sub DebugBar_Update()
    If NowIce.ClickState = MouseDown Then
        EUI.MoveWindow
    End If
End Sub
