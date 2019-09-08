VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mImage() As EImage
Sub Add(ByVal Path As String)
    ReDim Preserve mImage(UBound(mImage) + 1)
    Set mImage(UBound(mImage)) = New EImage
    mImage(UBound(mImage)).Path = Path
End Sub
Sub AddDir(ByVal Path As String)
    Dim File As String
    File = Dir(Path)
    Do While File <> ""
        Add Path & File
        File = Dir()
        DoEvents
    Loop
End Sub
Property Get Count() As Long
    Count = UBound(mImage)
End Property
Property Get ImageByPath(ByVal Path As String) As EImage
    For i = 1 To UBound(mImage)
        If mImage(i).Path = Path Then Set ImageByPath = mImage(i): Exit For
    Next
End Property
Property Get ImageByIndex(ByVal Index As Integer) As EImage
    Set ImageByIndex = mImage(Index)
End Property
Private Sub Class_Initialize()
    ReDim mImage(0)
End Sub
