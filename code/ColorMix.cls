VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mA As Long, mR As Long, mG As Long, mB As Long
Dim mColor As Long, mBrush As Long, mPen As Long, mWidth As Long
Dim mLight As Single
Public Property Get Light() As Long
    Light = mLight
End Property
Public Property Get brush() As Long
    brush = mBrush
End Property
Public Property Get pen() As Long
    pen = mPen
End Property
Public Property Get Color() As Long
    Color = mColor
End Property
Public Property Let Color(ByVal nColor As Long)
    mColor = nColor
    'ÄæÏò
    Dim ColorData(3) As Byte
    CopyMemory ColorData(0), nColor, 4
    mB = ColorData(0): mG = ColorData(1): mR = ColorData(2): mA = ColorData(3)
    Call ColorChanged
End Property
Public Property Get Width() As Long
    Width = mWidth
End Property
Public Property Let Width(ByVal nWidth As Long)
    mWidth = nWidth
    GdipSetPenWidth pen, nWidth
    AddRefColor Me
End Property
Public Property Get a() As Long
    a = mA
End Property
Public Property Let a(ByVal nA As Long)
    mA = nA
    Call ColorChanged
End Property
Public Property Get R() As Long
    R = mR
End Property
Public Property Let R(ByVal nR As Long)
    mR = nR
    Call ColorChanged
End Property
Public Property Get G() As Long
    G = mG
End Property
Public Property Let G(ByVal nG As Long)
    mG = nG
    Call ColorChanged
End Property
Public Property Get B() As Long
    B = mB
End Property
Public Property Let B(ByVal nB As Long)
    mB = nB
    Call ColorChanged
End Property
Private Sub ColorChanged()
    mColor = argb(mA, mR, mG, mB)
    GdipSetSolidFillColor mBrush, mColor
    GdipSetPenColor mPen, mColor
    mLight = ColorLight(mR, mG, mB) * (mA / 255) / 130050
    AddRefColor Me
End Sub
Private Sub Class_Initialize()
    If StartEasoft = False Then Err.Raise 4046, , "Blackmouth's love" & vbCrLf & "Start Easoft first ."
    mWidth = 1
    GdipCreateSolidFill 0, mBrush: GdipCreatePen1 0, 1, UnitPixel, mPen
    AddMember Me
End Sub
Public Sub Dispose()
    GdipDeleteBrush mBrush
    GdipDeletePen mPen
End Sub
