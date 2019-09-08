Attribute VB_Name = "DirtyRectangle"
Dim MyIcelolly() As Icelolly
Function IceLollyCount() As Long
    IceLollyCount = UBound(MyIcelolly)
End Function
Function GetIceLollyMember(ByVal Index As Integer) As Icelolly
    Set GetIceLollyMember = MyIcelolly(Index)
End Function
Sub BeginDirtyRect()
    ReDim MyIcelolly(0)
End Sub
Sub AddIcelolly(Ice As Icelolly)
    ReDim Preserve MyIcelolly(UBound(MyIcelolly) + 1)
    Set MyIcelolly(UBound(MyIcelolly)) = Ice
End Sub
Sub AddRefColor(Color As ColorMix)
    Dim Mem As Icelolly
    For i = 1 To UBound(MyIcelolly)
        Set Mem = MyIcelolly(i)
        If Mem.NeedPaint = False Then
            For s = 0 To 4
                If Mem.Style.Color(s) Is Color Then Mem.NeedPaint = True: Exit For
            Next
        End If
    Next
End Sub
Sub AddRefAni(Ani As AniManager)
    Dim Mem As Icelolly
    For i = 1 To UBound(MyIcelolly)
        Set Mem = MyIcelolly(i)
        If Mem.NeedPaint = False Then
            If (Mem.Animate Is Ani) Or (Mem.DrawAnimate Is Ani) Then Mem.NeedPaint = True
        End If
    Next
End Sub
Sub AddRefList(List As ListMemers)
    Dim Mem As Icelolly
    For i = 1 To UBound(MyIcelolly)
        Set Mem = MyIcelolly(i)
        If Mem.NeedPaint = False Then
            If Mem.List Is List Then Mem.NeedPaint = True
        End If
    Next
End Sub
Sub AddRefStyle(Style As StyleBox)
    Dim Mem As Icelolly
    For i = 1 To UBound(MyIcelolly)
        Set Mem = MyIcelolly(i)
        If Mem.NeedPaint = False Then
            If Mem.Style Is Style Then Mem.NeedPaint = True
        End If
    Next
End Sub
Sub AddRefIce(Ice As Icelolly)
    Dim Mem As Icelolly
    For i = 1 To UBound(MyIcelolly)
        Set Mem = MyIcelolly(i)
        If Mem.NeedPaint = False Then
            If Mem Is Ice Then Mem.NeedPaint = True
        End If
    Next
End Sub
Sub AddRefImage(Img As EImage)
    Dim Mem As Icelolly
    For i = 1 To UBound(MyIcelolly)
        Set Mem = MyIcelolly(i)
        If Mem.NeedPaint = False Then
            If Mem.Src Is Img Then Mem.NeedPaint = True
        End If
    Next
End Sub
Sub AddRefFont(Font As Fonter)
    Dim Mem As Icelolly
    For i = 1 To UBound(MyIcelolly)
        Set Mem = MyIcelolly(i)
        If Mem.NeedPaint = False Then
            If Mem.Style.Font Is Font Then Mem.NeedPaint = True
        End If
    Next
End Sub
