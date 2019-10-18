Attribute VB_Name = "Arctan"
Option Explicit

' Return the arc tangent of y/x taking into
' account the proper quadrant.
Public Function Arctan2(x As Single, y As Single)
Const PI = 3.14159265
Dim theta As Single

    If x = 0 Then
        If y > 0 Then
            Arctan2 = PI / 2
        Else
            Arctan2 = -PI / 2
        End If
    Else
        theta = Atn(y / x)
        If x < 0 Then theta = PI + theta
        Arctan2 = theta
    End If
End Function
