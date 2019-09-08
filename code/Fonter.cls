VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fonter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mID As String, mSize As Long, mStyle As EFontStyle
Dim mFontfam As Long, mFont As Long
Public Property Get Hwnd() As Long
    Hwnd = mFont
End Property
Public Property Get Id() As String
    Id = mID
End Property
Public Property Let Id(ByVal nID As String)
    mID = nID
    Call FontChanged
    AddRefFont Me
End Property
Public Property Get Size() As Long
    Size = mSize
End Property
Public Property Let Size(ByVal nSize As Long)
    mSize = nSize
    Call FontChanged
    AddRefFont Me
End Property
Public Property Get Style() As EFontStyle
    Style = mStyle
End Property
Public Property Let Style(nStyle As EFontStyle)
    mStyle = nStyle
    Call FontChanged
    AddRefFont Me
End Property
Private Sub FontChanged()
    GdipDeleteFont mFont
    GdipDeleteFontFamily mFontfam
    GdipCreateFontFamilyFromName StrPtr(mID), 0, mFontfam
    GdipCreateFont mFontfam, mSize, mStyle, UnitPixel, mFont
End Sub
Private Sub Class_Initialize()
    If StartEasoft = False Then Err.Raise 4046, , "Blackmouth's love" & vbCrLf & "Start Easoft first ."
    mID = "Î¢ÈíÑÅºÚ": mSize = 16: mStyle = Regular
    GdipCreateFontFamilyFromName StrPtr(mID), 0, mFontfam
    GdipCreateFont mFontfam, mSize, mStyle, UnitPixel, mFont
    AddMember Me
End Sub
