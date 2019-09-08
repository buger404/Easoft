Attribute VB_Name = "Dustbin"
Public EMembers() As Object, EShadows() As WinShadow
Sub AddMember(Member As Object)
    ReDim Preserve EMembers(UBound(EMembers) + 1)
    Set EMembers(UBound(EMembers)) = Member
End Sub
Sub AddEShadow(Shadow As WinShadow)
    ReDim Preserve EShadows(UBound(EShadows) + 1)
    Set EShadows(UBound(EShadows)) = Shadow
End Sub
Sub DeleteAllMember()
    On Local Error Resume Next
    
    For i = 1 To UBound(EMembers)
        EMembers(i).Dispose
    Next
    For i = 1 To UBound(EShadows)
        Set EShadows(i) = Nothing
    Next
End Sub
