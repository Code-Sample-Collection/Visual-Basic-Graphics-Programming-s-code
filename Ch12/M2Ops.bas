Attribute VB_Name = "M2Ops"
' Routines for manipulating 2-dimensional
' vectors and matrices.
Option Explicit

' Create a 2-dimensional identity matrix.
Public Sub m2Identity(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To 3
        For j = 1 To 3
            If i = j Then
                M(i, j) = 1
            Else
                M(i, j) = 0
            End If
        Next j
    Next i
End Sub

' Create a translation matrix for translation by
' distances tx and ty.
Public Sub m2Translate(Result() As Single, _
    ByVal tx As Single, ByVal ty As Single)

    m2Identity Result
    Result(3, 1) = tx
    Result(3, 2) = ty
End Sub

' Create a scaling matrix for scaling by factors
' of sx and sy.
Public Sub m2Scale(Result() As Single, _
    ByVal sx As Single, ByVal sy As Single)

    m2Identity Result
    Result(1, 1) = sx
    Result(2, 2) = sy
End Sub

' Create a rotation matrix for rotating by the
' given angle (in radians).
Public Sub m2Rotate(Result() As Single, ByVal theta As Single)
    m2Identity Result
    Result(1, 1) = Cos(theta)
    Result(1, 2) = Sin(theta)
    Result(2, 1) = -Result(1, 2)
    Result(2, 2) = Result(1, 1)
End Sub
' Create a rotation matrix that rotates the point
' (x, y) onto the X axis.
Public Sub m2RotateIntoX(Result() As Single, _
    ByVal X As Single, ByVal Y As Single)
Dim d As Single

    m2Identity Result
    d = Sqr(X * X + Y * Y)
    Result(1, 1) = X / d
    Result(1, 2) = -Y / d
    Result(2, 1) = -Result(1, 2)
    Result(2, 2) = Result(1, 1)
End Sub
' Create a scaling matrix for scaling by factors
' of sx and sy at the point (x, y).
Public Sub m2ScaleAt(Result() As Single, _
    ByVal sx As Single, ByVal sy As Single, _
    ByVal X As Single, ByVal Y As Single)
Dim T(1 To 3, 1 To 3) As Single
Dim S(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim M(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate T, -X, -Y

    ' Compute the inverse translation.
    m2Translate T_Inv, X, Y

    ' Scale.
    m2Scale S, sx, sy

    ' Combine the transformations.
    m2MatMultiply M, T, S           ' T * S
    m2MatMultiply Result, M, T_Inv  ' T * S * T_Inv
End Sub

' Create a matrix for reflecting across the line
' passing through (x, y) in direction <dx, dy>.
Public Sub m2ReflectAcross(Result() As Single, _
    ByVal X As Single, ByVal Y As Single, _
    ByVal dx As Single, ByVal dy As Single)
Dim T(1 To 3, 1 To 3) As Single
Dim R(1 To 3, 1 To 3) As Single
Dim S(1 To 3, 1 To 3) As Single
Dim R_Inv(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim M1(1 To 3, 1 To 3) As Single
Dim M2(1 To 3, 1 To 3) As Single
Dim M3(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate T, -X, -Y

    ' Compute the inverse translation.
    m2Translate T_Inv, X, Y

    ' Rotate so the direction vector lies in the Y axis.
    m2RotateIntoX R, dx, dy

    ' Compute the inverse translation.
    m2RotateIntoX R_Inv, dx, -dy

    ' Reflect across the X axis.
    m2Scale S, 1, -1

    ' Combine the transformations.
    m2MatMultiply M1, T, R     ' T * R
    m2MatMultiply M2, S, R_Inv ' S * R_Inv
    m2MatMultiply M3, M1, M2   ' T * R * S * R_Inv

    ' T * R * S * R_Inv * T_Inv
    m2MatMultiply Result, M3, T_Inv
End Sub

' Create a rotation matrix for rotating through
' angle theta around the point (x, y).
Public Sub m2RotateAround(Result() As Single, _
    ByVal theta As Single, _
    ByVal X As Single, ByVal Y As Single)
Dim T(1 To 3, 1 To 3) As Single
Dim R(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim M(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate T, -X, -Y

    ' Compute the inverse translation.
    m2Translate T_Inv, X, Y

    ' Rotate.
    m2Rotate R, theta

    ' Combine the transformations.
    m2MatMultiply M, T, R
    m2MatMultiply Result, M, T_Inv
End Sub

' Multiply a point and a matrix.
Public Sub m2PointMultiply(ByRef X As Single, ByRef Y As Single, A() As Single)
Dim newx As Single
Dim newy As Single

    newx = X * A(1, 1) + Y * A(2, 1) + A(3, 1)
    newy = X * A(1, 2) + Y * A(2, 2) + A(3, 2)
    X = newx
    Y = newy
End Sub
' Set copy = orig.
Public Sub m2PointCopy(copy() As Single, orig() As Single)
Dim i As Integer

    For i = 1 To 3
        copy(i) = orig(i)
    Next i
End Sub


' Set copy = orig.
Public Sub m2MatCopy(copy() As Single, orig() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To 3
        For j = 1 To 3
            copy(i, j) = orig(i, j)
        Next j
    Next i
End Sub



' Apply a transformation matrix to a point.
Public Sub m2Apply(V() As Single, A() As Single, Result() As Single)
    Result(1) = V(1) * A(1, 1) + V(2) * A(2, 1) + A(3, 1)
    Result(2) = V(1) * A(1, 2) + V(2) * A(2, 2) + A(3, 2)
    Result(3) = 1#
End Sub


' Multiply two transformation matrices.
Public Sub m2MatMultiply(Result() As Single, A() As Single, B() As Single)
    Result(1, 1) = A(1, 1) * B(1, 1) + A(1, 2) * B(2, 1)
    Result(1, 2) = A(1, 1) * B(1, 2) + A(1, 2) * B(2, 2)
    Result(1, 3) = 0#
    Result(2, 1) = A(2, 1) * B(1, 1) + A(2, 2) * B(2, 1)
    Result(2, 2) = A(2, 1) * B(1, 2) + A(2, 2) * B(2, 2)
    Result(2, 3) = 0#
    Result(3, 1) = A(3, 1) * B(1, 1) + A(3, 2) * B(2, 1) + B(3, 1)
    Result(3, 2) = A(3, 1) * B(1, 2) + A(3, 2) * B(2, 2) + B(3, 2)
    Result(3, 3) = 1#
End Sub

