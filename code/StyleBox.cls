VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Font As New Fonter
Dim mColor(5) As New ColorMix, mRadian As Long
Dim mAlign As EAlign, mImgAlign As EAlign, mShape As EShape
Public Property Get Radian() As Long
    Radian = mRadian
End Property
Public Property Let Radian(ByVal nRadian As Long)
    mRadian = nRadian
    AddRefStyle Me
End Property
Public Property Get Color(ByVal Theme As EColorTheme) As ColorMix
    Set Color = mColor(Theme)
End Property
Public Property Let Color(ByVal Theme As EColorTheme, nColor As ColorMix)
    Set mColor(Theme) = nColor
    AddRefStyle Me
End Property
Public Property Get Align() As EAlign
    Align = mAlign
End Property
Public Property Let Align(nAlign As EAlign)
    mAlign = nAlign
    AddRefStyle Me
End Property
Public Property Get ImgAlign() As EAlign
    ImgAlign = mImgAlign
End Property
Public Property Let ImgAlign(nImgAlign As EAlign)
    mImgAlign = nImgAlign
    AddRefStyle Me
End Property
Public Property Get Shape() As EShape
    Shape = mShape
End Property
Public Property Let Shape(ByVal nShape As EShape)
    mShape = nShape
    AddRefStyle Me
End Property
Private Sub Class_Initialize()
    If StartEasoft = False Then Err.Raise 4046, , "Blackmouth's love" & vbCrLf & "Start Easoft first ."
    mRadian = 4
End Sub
