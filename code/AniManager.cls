VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AniManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Animations() As EAnimate
Sub Clear()
    ReDim Animations(0)
End Sub
Sub AddByReset(Ani As EAnimate)
    ReDim Animations(1)
    Animations(UBound(Animations)) = Ani
    'AddRefAni Me
End Sub
Sub Add(Ani As EAnimate)
    ReDim Preserve Animations(UBound(Animations) + 1)
    Animations(UBound(Animations)) = Ani
    'AddRefAni Me
End Sub
Function Count() As Long
    Count = UBound(Animations)
End Function
Sub Update()
    For i = 1 To UBound(Animations)
    
        If i < 1 Then Exit For
        If i > UBound(Animations) Then Exit For

        If GetTickCount - Animations(i).StartTime <= Animations(i).Duration + Animations(i).Delay Then
            If GetTickCount - Animations(i).StartTime >= Animations(i).Delay Then
                If Animations(i).Index = -1 Then
                    CallByName Animations(i).Obj, Animations(i).ProperName, VbLet, CallByName(GobalAnimation, "Ani_" & Animations(i).FuncName, VbMethod, Animations(i))
                Else
                    CallByName Animations(i).Obj, Animations(i).ProperName, VbLet, Animations(i).Index, CallByName(GobalAnimation, "Ani_" & Animations(i).FuncName, VbMethod, Animations(i))
                End If
            Else
                If Animations(i).Index = -1 Then
                    CallByName Animations(i).Obj, Animations(i).ProperName, VbLet, Val(Animations(i).Start)
                Else
                    CallByName Animations(i).Obj, Animations(i).ProperName, VbLet, Animations(i).Index, Val(Animations(i).Start)
                End If
            End If
        Else
                If Animations(i).Index = -1 Then
                    If CallByName(Animations(i).Obj, Animations(i).ProperName, VbGet) <> Animations(i).Target Then CallByName Animations(i).Obj, Animations(i).ProperName, VbLet, Val(Animations(i).Target)
                Else
                    If CallByName(Animations(i).Obj, Animations(i).ProperName, VbGet, Animations(i).Index) <> Animations(i).Target Then CallByName Animations(i).Obj, Animations(i).ProperName, VbLet, Animations(i).Index, Val(Animations(i).Target)
                End If
            'É¾³ý²Ù×÷
            Animations(i) = Animations(UBound(Animations))
            ReDim Preserve Animations(UBound(Animations) - 1)
            i = i - 1
        End If
    Next
End Sub
Private Sub Class_Initialize()
    ReDim Animations(0)
End Sub
