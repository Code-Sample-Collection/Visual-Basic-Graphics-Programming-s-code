Attribute VB_Name = "M4Ops"
Option Explicit

' Four-dimensional matrix routines.

Type Point4D
    coord(1 To 5) As Single
    trans(1 To 5) As Single
End Type

Type Segment4D
    pt1 As Integer
    pt2 As Integer
End Type

Public Const PI = 3.14159265
Public Const INFINITY = 2147483647
' Create a transformation matrix for orthographic
' projection along the X-W plane.
Public Sub m4OrthoSide(M() As Single)
    m4Identity M
    M(1, 1) = 0
    M(3, 1) = -1
    M(3, 3) = 0
    M(4, 4) = 0
End Sub
' Create a transformation matrix for orthographic
' projection along the Y-W plane.
Public Sub m4OrthoTop(M() As Single)
    m4Identity M
    M(2, 2) = 0
    M(3, 2) = -1
    M(3, 3) = 0
    M(4, 4) = 0
End Sub

' Create a transformation matrix for orthographic
' projection along the W-Z plane.
Public Sub m4OrthoFront(M() As Single)
    m4Identity M
    M(3, 3) = 0
    M(4, 4) = 0
End Sub

' Create an identity matrix.
Public Sub m4Identity(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To 5
        For j = 1 To 5
            If i = j Then
                M(i, j) = 1
            Else
                M(i, j) = 0
            End If
        Next j
    Next i
End Sub

' Normalize a 4-D point vector.
Public Sub m4NormalizeCoords(X As Single, Y As Single, Z As Single, W As Single, S As Single)
    X = X / S
    Y = Y / S
    Z = Z / S
    W = W / S
    S = 1
End Sub

' Normalize a 4-D point vector.
Public Sub m4NormalizePoint(P() As Single)
Dim i As Integer
Dim value As Single

    value = P(5)
    For i = 1 To 4
        P(i) = P(i) / value
    Next i
    P(5) = 1
End Sub


' Normalize a 4-D transformation matrix.
Public Sub m4NormalizeMatrix(M() As Single)
Dim i As Integer
Dim j As Integer
Dim value As Single

    value = M(5, 5)
    For i = 1 To 5
        For j = 1 To 5
            M(i, j) = M(i, j) / value
        Next j
    Next i
End Sub





' Create a 4-D transformation matrix for a
' perspective projection along the W axis into
' the X-Y-Z space with focus at the origin and the
' center of projection at point (0, 0, 0, D).
Public Sub m4PerspectiveW(M() As Single, D As Single)
    m4Identity M
    If D <> 0 Then M(4, 5) = -1 / D
End Sub
' Create a 4-D transformation matrix for scaling
' by scale factors Sx, Sy, Sz, and Sw.
Public Sub m4Scale(M() As Single, Sx As Single, Sy As Single, Sz As Single, Sw As Single)
    m4Identity M
    M(1, 1) = Sx
    M(2, 2) = Sy
    M(3, 3) = Sz
    M(4, 4) = Sw
End Sub

' Create a 3-D transformation matrix for
' translation by Tx, Ty, Tz, and Tw.
Public Sub m4Translate(M() As Single, Tx As Single, Ty As Single, Tz As Single, Tw As Single)
    m4Identity M
    M(5, 1) = Tx
    M(5, 2) = Ty
    M(5, 3) = Tz
    M(5, 4) = Tw
End Sub

' Create a 4-D transformation matrix for rotation
' around the XY plane (angle measured in radians).
Public Sub m4XYRotate(M() As Single, theta As Single)
    m4Identity M
    M(3, 3) = Cos(theta)
    M(4, 4) = M(3, 3)
    M(3, 4) = Sin(theta)
    M(4, 3) = -M(3, 4)
End Sub

' Create a 4-D transformation matrix for rotation
' around the XZ plane (angle measured in radians).
Public Sub m4XZRotate(M() As Single, theta As Single)
    m4Identity M
    M(2, 2) = Cos(theta)
    M(4, 4) = M(2, 2)
    M(2, 4) = Sin(theta)
    M(4, 2) = -M(2, 4)
End Sub


' Create a 4-D transformation matrix for rotation
' around the YZ plane (angle measured in radians).
Public Sub m4YZRotate(M() As Single, theta As Single)
    m4Identity M
    M(1, 1) = Cos(theta)
    M(4, 4) = M(1, 1)
    M(1, 4) = Sin(theta)
    M(4, 1) = -M(1, 4)
End Sub
' Create a 4-D transformation matrix for rotation
' around the XW plane (angle measured in radians).
Public Sub m4XWRotate(M() As Single, theta As Single)
    m4Identity M
    M(2, 2) = Cos(theta)
    M(3, 3) = M(2, 2)
    M(2, 3) = Sin(theta)
    M(3, 2) = -M(2, 3)
End Sub


' Create a 4-D transformation matrix for rotation
' around the YW plane (angle measured in radians).
Public Sub m4YWRotate(M() As Single, theta As Single)
    m4Identity M
    M(1, 1) = Cos(theta)
    M(3, 3) = M(1, 1)
    M(3, 1) = Sin(theta)
    M(1, 3) = -M(3, 1)
End Sub

' Create a 4-D transformation matrix for rotation
' around the ZW plane (angle measured in radians).
Public Sub m4ZWRotate(M() As Single, theta As Single)
    m4Identity M
    M(1, 1) = Cos(theta)
    M(2, 2) = M(1, 1)
    M(1, 2) = Sin(theta)
    M(2, 1) = -M(1, 2)
End Sub


' Set copy = orig.
Public Sub m4MatCopy(copy() As Single, orig() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To 5
        For j = 1 To 5
            copy(i, j) = orig(i, j)
        Next j
    Next i
End Sub

' Apply a transformation matrix to a point where
' the transformation may not have 0, 0, 0, 1 in
' its final column. Normalize only the X and Y
' components of the result to preserve the Z
' information.
Public Sub m4ApplyFull(V() As Single, M() As Single, Result() As Single)
Dim i As Integer
Dim j As Integer
Dim value As Single

    For i = 1 To 5
        value = 0#
        For j = 1 To 5
            value = value + V(j) * M(j, i)
        Next j
        Result(i) = value
    Next i
    
    ' Renormalize the point.
    ' Note that value still holds Result(5).
    If value <> 0 Then
        Result(1) = Result(1) / value
        Result(2) = Result(2) / value
        Result(3) = Result(3) / value
        ' Do not transform the w component.
    Else
        ' Make the W value greater than that of
        ' the center of projection so the point
        ' will be clipped.
        Result(4) = INFINITY
    End If
    Result(5) = 1#
End Sub




' Apply a transformation matrix to a point.
Public Sub m4Apply(V() As Single, M() As Single, Result() As Single)
    Result(1) = V(1) * M(1, 1) + _
                V(2) * M(2, 1) + _
                V(3) * M(3, 1) + _
                V(4) * M(4, 1) + M(5, 1)
    Result(2) = V(1) * M(1, 2) + _
                V(2) * M(2, 2) + _
                V(3) * M(3, 2) + _
                V(4) * M(4, 2) + M(5, 2)
    Result(3) = V(1) * M(1, 3) + _
                V(2) * M(2, 3) + _
                V(3) * M(3, 3) + _
                V(4) * M(4, 3) + M(5, 3)
    Result(4) = V(1) * M(1, 4) + _
                V(2) * M(2, 4) + _
                V(3) * M(3, 4) + _
                V(4) * M(4, 4) + M(5, 4)
    Result(5) = 1#
End Sub

' Multiply two matrices together. The matrices
' may not contain 0, 0, 0, 0, 1 in their last
' columns.
Public Sub m4MatMultiplyFull(Result() As Single, A() As Single, B() As Single)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim value As Single

    For i = 1 To 5
        For j = 1 To 5
            value = 0#
            For k = 1 To 5
                value = value + A(i, k) * B(k, j)
            Next k
            Result(i, j) = value
        Next j
    Next i
End Sub
' Multiply two matrices together.
Public Sub m4MatMultiply(Result() As Single, A() As Single, B() As Single)
    Result(1, 1) = A(1, 1) * B(1, 1) + A(1, 2) * B(2, 1) + A(1, 3) * B(3, 1) + A(1, 4) * B(4, 1)
    Result(1, 2) = A(1, 1) * B(1, 2) + A(1, 2) * B(2, 2) + A(1, 3) * B(3, 2) + A(1, 4) * B(4, 2)
    Result(1, 3) = A(1, 1) * B(1, 3) + A(1, 2) * B(2, 3) + A(1, 3) * B(3, 3) + A(1, 4) * B(4, 3)
    Result(1, 4) = A(1, 1) * B(1, 4) + A(1, 2) * B(2, 4) + A(1, 3) * B(3, 4) + A(1, 4) * B(4, 4)
    Result(1, 5) = 0#
    Result(2, 1) = A(2, 1) * B(1, 1) + A(2, 2) * B(2, 1) + A(2, 3) * B(3, 1) + A(2, 4) * B(4, 1)
    Result(2, 2) = A(2, 1) * B(1, 2) + A(2, 2) * B(2, 2) + A(2, 3) * B(3, 2) + A(2, 4) * B(4, 2)
    Result(2, 3) = A(2, 1) * B(1, 3) + A(2, 2) * B(2, 3) + A(2, 3) * B(3, 3) + A(2, 4) * B(4, 3)
    Result(2, 4) = A(2, 1) * B(1, 4) + A(2, 2) * B(2, 4) + A(2, 3) * B(3, 4) + A(2, 4) * B(4, 4)
    Result(2, 5) = 0#
    Result(3, 1) = A(3, 1) * B(1, 1) + A(3, 2) * B(2, 1) + A(3, 3) * B(3, 1) + A(3, 4) * B(4, 1)
    Result(3, 2) = A(3, 1) * B(1, 2) + A(3, 2) * B(2, 2) + A(3, 3) * B(3, 2) + A(3, 4) * B(4, 2)
    Result(3, 3) = A(3, 1) * B(1, 3) + A(3, 2) * B(2, 3) + A(3, 3) * B(3, 3) + A(3, 4) * B(4, 3)
    Result(3, 4) = A(3, 1) * B(1, 4) + A(3, 2) * B(2, 4) + A(3, 3) * B(3, 4) + A(3, 4) * B(4, 4)
    Result(3, 5) = 0#
    Result(4, 1) = A(4, 1) * B(1, 1) + A(4, 2) * B(2, 1) + A(4, 3) * B(3, 1) + A(4, 4) * B(4, 1)
    Result(4, 2) = A(4, 1) * B(1, 2) + A(4, 2) * B(2, 2) + A(4, 3) * B(3, 2) + A(4, 4) * B(4, 2)
    Result(4, 3) = A(4, 1) * B(1, 3) + A(4, 2) * B(2, 3) + A(4, 3) * B(3, 3) + A(4, 4) * B(4, 3)
    Result(4, 4) = A(4, 1) * B(1, 4) + A(4, 2) * B(2, 4) + A(4, 3) * B(3, 4) + A(4, 4) * B(4, 4)
    Result(4, 5) = 0#
    Result(5, 1) = A(5, 1) * B(1, 1) + A(5, 2) * B(2, 1) + A(5, 3) * B(3, 1) + A(5, 4) * B(4, 1) + B(5, 1)
    Result(5, 2) = A(5, 1) * B(1, 2) + A(5, 2) * B(2, 2) + A(5, 3) * B(3, 2) + A(5, 4) * B(4, 2) + B(5, 2)
    Result(5, 3) = A(5, 1) * B(1, 3) + A(5, 2) * B(2, 3) + A(5, 3) * B(3, 3) + A(5, 4) * B(4, 3) + B(5, 3)
    Result(5, 4) = A(5, 1) * B(1, 4) + A(5, 2) * B(2, 4) + A(5, 3) * B(3, 4) + A(5, 4) * B(4, 4) + B(5, 4)
    Result(5, 5) = 1#
End Sub


' Give the vector the indicated length.
Public Sub m4SizeVector(ByVal L As Single, X As Single, Y As Single, Z As Single, W As Single)
    L = L / Sqr(X * X + Y * Y + Z * Z + W * W)
    X = X * L
    Y = Y * L
    Z = Z * L
    W = W * L
End Sub
