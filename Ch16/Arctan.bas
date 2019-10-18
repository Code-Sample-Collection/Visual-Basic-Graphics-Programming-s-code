Attribute VB_Name = "Arctan"
Option Explicit

' Return the arctan of dy/dx.
Public Function ATan2(ByVal dy As Single, ByVal dx As Single) As Single
Const PI = 3.14159265

Dim theta As Single

    If Abs(dx) < 0.01 Then
        If dy < 0 Then
            theta = -PI / 2
        Else
            theta = PI / 2
        End If
    Else
        theta = Atn(dy / dx)
        If dx < 0 Then theta = PI + theta
    End If

    ATan2 = theta
End Function
