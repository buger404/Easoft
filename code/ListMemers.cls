VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListMemers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mIndex As Integer, mList() As String, mImage() As EImage
Dim mNothingText As String
Public Property Get NothingText() As String
    NothingText = mNothingText
End Property
Public Property Let NothingText(ByVal nNothingText As String)
    mNothingText = nNothingText
    AddRefList Me
End Property
Public Property Get Index() As Integer
    Index = mIndex
End Property
Public Property Let Index(ByVal nIndex As Integer)
    mIndex = nIndex
    AddRefList Me
End Property
Public Property Get Count() As Long
    Count = UBound(mList)
End Property
Public Property Get Items(ByVal Index As Integer) As String
    Items = mList(Index)
End Property
Public Property Let Items(ByVal Index As Integer, ByVal nItem As String)
    mList(Index) = nItem
    AddRefList Me
End Property
Public Property Get Src(ByVal Index As Integer) As EImage
    Set Src = mImage(Index)
End Property
Public Property Let Src(ByVal Index As Integer, ByVal nSrc As EImage)
    Set mImage(Index) = nSrc
    AddRefList Me
End Property
Sub Add(ByVal Item As String)
    ReDim Preserve mList(UBound(mList) + 1)
    ReDim Preserve mImage(UBound(mImage) + 1)
    mList(UBound(mList)) = Item
    AddRefList Me
End Sub
Sub Refresh()
    AddRefList Me
End Sub
Sub Remove(ByVal Item As String)
    For i = 1 To UBound(mList)
        If mList(i) = Item Then
            For s = i To UBound(mList) - 1
                mList(s) = mList(s + 1)
                Set mImage(s) = mImage(s + 1)
            Next
            ReDim Preserve mList(UBound(mList) - 1)
            ReDim Preserve mImage(UBound(mImage) - 1)
            Exit For
        End If
    Next
    AddRefList Me
End Sub
Sub Clear()
    ReDim mList(0), mImage(0)
    AddRefList Me
End Sub
Private Sub Class_Initialize()
    ReDim mList(0), mImage(0)
    mNothingText = "啊哦，什么也没有呢。"
    AddRefList Me
End Sub
