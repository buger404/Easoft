Attribute VB_Name = "Animation"
Function Ani_linear(Animate As EAnimate)
    Dim h As Long, a As Single
    h = (Animate.Target - Animate.Start)
    a = -h / (-((-Animate.Duration) / 50) ^ 2)
    Ani_linear = (-((GetTickCount - Animate.StartTime - Animate.Duration) / 50) ^ 2) * a + h + Animate.Start
End Function
Function Ani_normal(Animate As EAnimate)
    Dim a As Single
    a = (GetTickCount - Animate.StartTime) / Animate.Duration
    Ani_normal = Animate.Start + (Animate.Target - Animate.Start) * a
End Function
Function Ani_color(Animate As EAnimate) As Long
    Dim a As Single
    Dim temp(3) As Byte, temp2(3) As Byte, temp3(3) As Byte, temp4 As Long
    CopyMemory temp(0), Animate.Start, 4
    CopyMemory temp2(0), Animate.Target, 4
    a = (GetTickCount - Animate.StartTime) / Animate.Duration
    For i = 0 To 3
        temp3(i) = temp(i) + (temp2(i) - temp(i)) * a
    Next
    CopyMemory temp4, temp3(0), 4
    Ani_color = temp4
End Function
Function Ani_linearcolor(Animate As EAnimate) As Long
    Dim a As Single
    Dim temp(3) As Byte, temp2(3) As Byte, temp3(3) As Byte, temp4 As Long
    Dim h As Long, a As Single
    CopyMemory temp(0), Animate.Start, 4
    CopyMemory temp2(0), Animate.Target, 4
    For i = 0 To 3
        h = temp2(i)
        h = h - temp(i)
        a = -h / (-((-Animate.Duration) / 50) ^ 2)
        temp3(i) = (-((GetTickCount - Animate.StartTime - Animate.Duration) / 50) ^ 2) * a + h + temp(i)
    Next
    CopyMemory temp4, temp3(0), 4
    Ani_linearcolor = temp4
End Function
