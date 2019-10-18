Attribute VB_Name = "Matrix"
Option Explicit

' Routines for making and manipulating matrices.

Public Const PI = 3.14159265358979
Public Const INFINITY = 2147483647

Public Enum ProjectionTypes
    project_Parallel
    project_Perspective
End Enum

Public Type Point3D
    Coord(1 To 4) As Single
    Trans(1 To 4) As Single
End Type

Public Type Segment3D
    pt1 As Integer
    pt2 As Integer
End Type
' Convert the spherical coordinates into
' Cartesian coordinates.
Public Sub m3SphericalToCartesian(ByVal r As Single, ByVal theta As Single, ByVal phi As Single, X As Single, Y As Single, Z As Single)
Dim r2 As Single
    
    ' Create a line to the center of projection.
    Y = r * Sin(phi)
    r2 = r * Cos(phi)
    X = r2 * Cos(theta)
    Z = r2 * Sin(theta)
End Sub
' Create a transformation matrix for an oblique
' projection onto the X-Y plane.
Public Sub m3ObliqueXY(M() As Single, ByVal S As Single, ByVal theta As Single)
    m3Identity M
    M(3, 1) = -S * Cos(theta)
    M(3, 2) = -S * Sin(theta)
    M(3, 3) = 0
End Sub


' Create a transformation matrix for orthographic
' projection along the X axis.
Public Sub m3OrthoSide(M() As Single)
    m3Identity M
    M(1, 1) = 0
    M(3, 1) = -1
    M(3, 3) = 0
End Sub
' Create a transformation matrix for orthographic
' projection along the Y axis.
Public Sub m3OrthoTop(M() As Single)
    m3Identity M
    M(2, 2) = 0
    M(3, 2) = -1
    M(3, 3) = 0
End Sub

' Create a transformation matrix for orthographic
' projection along the Z axis.
Public Sub m3OrthoFront(M() As Single)
    m3Identity M
    M(3, 3) = 0
End Sub

' Create an identity matrix.
Public Sub m3Identity(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To 4
        For j = 1 To 4
            If i = j Then
                M(i, j) = 1
            Else
                M(i, j) = 0
            End If
        Next j
    Next i
End Sub

' Normalize a 3-D point vector.
Public Sub m3NormalizeCoords(X As Single, Y As Single, Z As Single, S As Single)
    X = X / S
    Y = Y / S
    Z = Z / S
    S = 1
End Sub

' Normalize a 3-D point vector.
Public Sub m3NormalizePoint(P() As Single)
Dim i As Integer
Dim value As Single

    value = P(4)
    For i = 1 To 3
        P(i) = P(i) / value
    Next i
    P(4) = 1
End Sub


' Normalize a 3-D transformation matrix.
Public Sub m3NormalizeMatrix(M() As Single)
Dim i As Integer
Dim j As Integer
Dim value As Single

    value = M(4, 4)
    For i = 1 To 4
        For j = 1 To 4
            M(i, j) = M(i, j) / value
        Next j
    Next i
End Sub




' Create a 3-D transformation matrix for a
' perspective projection along the Z axis onto
' the X-Y plane with focus at the origin and the
' center of projection at distance (0, 0, D).
Public Sub project_PerspectiveXY(M() As Single, ByVal D As Single)
    m3Identity M
    If D <> 0 Then M(3, 4) = -1 / D
End Sub

' Create a 3-D transformation matrix for a
' projection with:
'       center of projection    (cx, cy, cz)
'       focus                   (fx, fy, fx)
'       UP vector               <ux, yx, uz>
' ptype should be project_Perspective or project_Parallel.
Public Sub m3Project(M() As Single, ByVal ptype As ProjectionTypes, ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal Fx As Single, ByVal Fy As Single, ByVal Fz As Single, ByVal ux As Single, ByVal uy As Single, ByVal uz As Single)
Static M1(1 To 4, 1 To 4) As Single
Static M2(1 To 4, 1 To 4) As Single
Static M3(1 To 4, 1 To 4) As Single
Static M4(1 To 4, 1 To 4) As Single
Static M5(1 To 4, 1 To 4) As Single
Static M12(1 To 4, 1 To 4) As Single
Static M34(1 To 4, 1 To 4) As Single
Static M1234(1 To 4, 1 To 4) As Single
Dim sin1 As Single
Dim cos1 As Single
Dim sin2 As Single
Dim cos2 As Single
Dim sin3 As Single
Dim cos3 As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim d1 As Single
Dim d2 As Single
Dim d3 As Single
Dim up1(1 To 4) As Single
Dim up2(1 To 4) As Single

    ' Translate the focus to the origin.
    m3Translate M1, -Fx, -Fy, -Fz

    A = Cx - Fx
    B = Cy - Fy
    C = Cz - Fz
    d1 = Sqr(A * A + C * C)
    If d1 <> 0 Then
        sin1 = -A / d1
        cos1 = C / d1
    End If
    d2 = Sqr(A * A + B * B + C * C)
    If d2 <> 0 Then
        sin2 = B / d2
        cos2 = d1 / d2
    End If
    
    ' Rotate around the Y axis to place the
    ' center of projection in the Y-Z plane.
    m3Identity M2
    
    ' If d1 = 0 then the center of projection
    ' already lies in the Y axis and thus the Y-Z plane.
    If d1 <> 0 Then
        M2(1, 1) = cos1
        M2(1, 3) = -sin1
        M2(3, 1) = sin1
        M2(3, 3) = cos1
    End If
    
    ' Rotate around the X axis to place the
    ' center of projection in the Z axis.
    m3Identity M3
    
    ' If d2 = 0 then the center of projection
    ' lies at the origin. This makes projection
    ' impossible.
    If d2 <> 0 Then
        M3(2, 2) = cos2
        M3(2, 3) = sin2
        M3(3, 2) = -sin2
        M3(3, 3) = cos2
    End If
    
    ' Apply the rotations to the UP vector.
    up1(1) = ux
    up1(2) = uy
    up1(3) = uz
    up1(4) = 1
    m3Apply up1, M2, up2
    m3Apply up2, M3, up1

    ' Rotate around the Z axis to put the UP
    ' vector in the Y-Z plane.
    d3 = Sqr(up1(1) * up1(1) + up1(2) * up1(2))
    m3Identity M4
    
    ' If d3 = 0 then the UP vector is a zero
    ' vector so do nothing.
    If d3 <> 0 Then
        sin3 = up1(1) / d3
        cos3 = up1(2) / d3
        M4(1, 1) = cos3
        M4(1, 2) = sin3
        M4(2, 1) = -sin3
        M4(2, 2) = cos3
    End If
    
    ' Project.
    If ptype = project_Perspective And d2 <> 0 Then
        project_PerspectiveXY M5, d2
    Else
        m3Identity M5
    End If

    ' Combine the transformations.
    m3MatMultiply M12, M1, M2
    m3MatMultiply M34, M3, M4
    m3MatMultiply M1234, M12, M34
    If ptype = project_Perspective Then
        m3MatMultiplyFull M, M1234, M5
    Else
        m3MatMultiply M, M1234, M5
    End If
End Sub



' Create a 3-D transformation matrix for a
' perspective projection with:
'       center of projection    (r, phi, theta)
'       focus                   (fx, fy, fx)
'       up vector               <ux, uy, uz>
' ptype should be project_Perspective or project_Parallel.
Public Sub m3PProject(M() As Single, ByVal ptype As ProjectionTypes, ByVal r As Single, ByVal phi As Single, ByVal theta As Single, ByVal Fx As Single, ByVal Fy As Single, ByVal Fz As Single, ByVal ux As Single, ByVal uy As Single, ByVal uz As Single)
Dim Cx As Single
Dim Cy As Single
Dim Cz As Single
Dim r2 As Single

    ' Convert to Cartesian coordinates.
    Cy = r * Sin(phi)
    r2 = r * Cos(phi)
    Cx = r2 * Cos(theta)
    Cz = r2 * Sin(theta)
    m3Project M, ptype, Cx, Cy, Cz, Fx, Fy, Fz, ux, uy, uz
End Sub

' Create a transformation matrix for reflecting
' across the plane passing through (p1, p2, p3)
' with normal vector <n1, n2, n3>.
Public Sub m3Reflect(M() As Single, ByVal p1 As Single, ByVal p2 As Single, ByVal p3 As Single, ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single)
Dim T(1 To 4, 1 To 4) As Single     ' Translate.
Dim R1(1 To 4, 1 To 4) As Single    ' Rotate 1.
Dim r2(1 To 4, 1 To 4) As Single    ' Rotate 2.
Dim S(1 To 4, 1 To 4) As Single     ' Reflect.
Dim R2i(1 To 4, 1 To 4) As Single   ' Unrotate 2.
Dim R1i(1 To 4, 1 To 4) As Single   ' Unrotate 1.
Dim Ti(1 To 4, 1 To 4) As Single    ' Untranslate.
Dim D As Single
Dim l As Single
Dim M12(1 To 4, 1 To 4) As Single
Dim M34(1 To 4, 1 To 4) As Single
Dim M1234(1 To 4, 1 To 4) As Single
Dim M56(1 To 4, 1 To 4) As Single
Dim M567(1 To 4, 1 To 4) As Single

    ' Translate the plane to the origin.
    m3Translate T, -p1, -p2, -p3
    m3Translate Ti, p1, p2, p3

    ' Rotate around Z-axis until the normal is in
    ' the Y-Z plane.
    m3Identity R1
    D = Sqr(n1 * n1 + n2 * n2)
    R1(1, 1) = n2 / D
    R1(1, 2) = n1 / D
    R1(2, 1) = -R1(1, 2)
    R1(2, 2) = R1(1, 1)
    
    m3Identity R1i
    R1i(1, 1) = R1(1, 1)
    R1i(1, 2) = -R1(1, 2)
    R1i(2, 1) = -R1(2, 1)
    R1i(2, 2) = R1(2, 2)
    
    ' Rotate around the X-axis until the normal
    ' lies along the Y axis.
    m3Identity r2
    l = Sqr(n1 * n1 + n2 * n2 + n3 * n3)
    r2(2, 2) = D / l
    r2(2, 3) = -n3 / l
    r2(3, 2) = -r2(2, 3)
    r2(3, 3) = r2(2, 2)
    
    m3Identity R2i
    R2i(2, 2) = r2(2, 2)
    R2i(2, 3) = -r2(2, 3)
    R2i(3, 2) = -r2(3, 2)
    R2i(3, 3) = r2(3, 3)

    ' Reflect across the X-Z plane.
    m3Identity S
    S(2, 2) = -1

    ' Combine the matrices.
    m3MatMultiply M12, T, R1
    m3MatMultiply M34, r2, S
    m3MatMultiply M1234, M12, M34
    m3MatMultiply M56, R2i, R1i
    m3MatMultiply M567, M56, Ti
    m3MatMultiply M, M1234, M567
End Sub


' Create a transformation atrix for rotating
' through angle theta around a line passing
' through (p1, p2, p3) in direction <d1, d2, d3>.
' Theta is measured counterclockwise as you look
' down the line opposite the line's direction.
Public Sub m3LineRotate(M() As Single, ByVal p1 As Single, ByVal p2 As Single, ByVal p3 As Single, ByVal d1 As Single, ByVal d2 As Single, ByVal d3 As Single, ByVal theta As Single)
Dim T(1 To 4, 1 To 4) As Single     ' Translate.
Dim R1(1 To 4, 1 To 4) As Single    ' Rotate 1.
Dim r2(1 To 4, 1 To 4) As Single    ' Rotate 2.
Dim Rot3(1 To 4, 1 To 4) As Single  ' Rotate.
Dim R2i(1 To 4, 1 To 4) As Single   ' Unrotate 2.
Dim R1i(1 To 4, 1 To 4) As Single   ' Unrotate 1.
Dim Ti(1 To 4, 1 To 4) As Single    ' Untranslate.
Dim D As Single
Dim l As Single
Dim M12(1 To 4, 1 To 4) As Single
Dim M34(1 To 4, 1 To 4) As Single
Dim M1234(1 To 4, 1 To 4) As Single
Dim M56(1 To 4, 1 To 4) As Single
Dim M567(1 To 4, 1 To 4) As Single

    ' Translate the plane to the origin.
    m3Translate T, -p1, -p2, -p3
    m3Translate Ti, p1, p2, p3

    ' Rotate around Z-axis until the line is in
    ' the Y-Z plane.
    m3Identity R1
    D = Sqr(d1 * d1 + d2 * d2)
    R1(1, 1) = d2 / D
    R1(1, 2) = d1 / D
    R1(2, 1) = -R1(1, 2)
    R1(2, 2) = R1(1, 1)
    
    m3Identity R1i
    R1i(1, 1) = R1(1, 1)
    R1i(1, 2) = -R1(1, 2)
    R1i(2, 1) = -R1(2, 1)
    R1i(2, 2) = R1(2, 2)
    
    ' Rotate around the X-axis until the line
    ' lies along the Y axis.
    m3Identity r2
    l = Sqr(d1 * d1 + d2 * d2 + d3 * d3)
    r2(2, 2) = D / l
    r2(2, 3) = -d3 / l
    r2(3, 2) = -r2(2, 3)
    r2(3, 3) = r2(2, 2)
    
    m3Identity R2i
    R2i(2, 2) = r2(2, 2)
    R2i(2, 3) = -r2(2, 3)
    R2i(3, 2) = -r2(3, 2)
    R2i(3, 3) = r2(3, 3)

    ' Rotate around the line (Y axis).
    m3YRotate Rot3, theta

    ' Combine the matrices.
    m3MatMultiply M12, T, R1
    m3MatMultiply M34, r2, Rot3
    m3MatMultiply M1234, M12, M34
    m3MatMultiply M56, R2i, R1i
    m3MatMultiply M567, M56, Ti
    m3MatMultiply M, M1234, M567
End Sub

' Create a 3-D transformation matrix for scaling
' by scale factors Sx, Sy, and Sz.
Public Sub m3Scale(M() As Single, ByVal Sx As Single, ByVal Sy As Single, ByVal Sz As Single)
    m3Identity M
    M(1, 1) = Sx
    M(2, 2) = Sy
    M(3, 3) = Sz
End Sub

' Create a 3-D transformation matrix for
' translation by Tx, Ty, and Tz.
Public Sub m3Translate(M() As Single, ByVal Tx As Single, ByVal Ty As Single, ByVal Tz As Single)
    m3Identity M
    M(4, 1) = Tx
    M(4, 2) = Ty
    M(4, 3) = Tz
End Sub

' Create a 3-D transformation matrix for rotation
' around the X axis (angle measured in radians).
Public Sub m3XRotate(M() As Single, ByVal theta As Single)
    m3Identity M
    M(2, 2) = Cos(theta)
    M(3, 3) = M(2, 2)
    M(2, 3) = Sin(theta)
    M(3, 2) = -M(2, 3)
End Sub

' Create a 3-D transformation matrix for rotation
' around the Y axis (angle measured in radians).
Public Sub m3YRotate(M() As Single, ByVal theta As Single)
    m3Identity M
    M(1, 1) = Cos(theta)
    M(3, 3) = M(1, 1)
    M(3, 1) = Sin(theta)
    M(1, 3) = -M(3, 1)
End Sub

' Create a 3-D transformation matrix for rotation
' around the Z axis (angle measured in radians).
Public Sub m3ZRotate(M() As Single, ByVal theta As Single)
    m3Identity M
    M(1, 1) = Cos(theta)
    M(2, 2) = M(1, 1)
    M(1, 2) = Sin(theta)
    M(2, 1) = -M(1, 2)
End Sub

' Create a matrix that rotates around the Y axis
' so the point (x, y, z) lies in the X-Z plane.
Public Sub m3YRotateIntoXZ(Result() As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim D As Single

    m3Identity Result
    D = Sqr(X * X + Y * Y)
    Result(1, 1) = X / D
    Result(1, 2) = -Y / D
    Result(2, 1) = -Result(1, 2)
    Result(2, 2) = Result(1, 1)
End Sub

' Set copy = orig.
Public Sub m3MatCopy(copy() As Single, orig() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To 4
        For j = 1 To 4
            copy(i, j) = orig(i, j)
        Next j
    Next i
End Sub

' Apply a transformation matrix to a point where
' the transformation may not have 0, 0, 0, 1 in
' its final column. Normalize only the X and Y
' components of the result to preserve the Z
' information.
Public Sub m3ApplyFull(V() As Single, M() As Single, Result() As Single)
Dim i As Integer
Dim j As Integer
Dim value As Single

    For i = 1 To 4
        value = 0#
        For j = 1 To 4
            value = value + V(j) * M(j, i)
        Next j
        Result(i) = value
    Next i

    ' Renormalize the point.
    ' Note that value still holds Result(4).
    If value <> 0 Then
        Result(1) = Result(1) / value
        Result(2) = Result(2) / value
        ' Do not transform the Z component.
    Else
        ' Make the Z value greater than that of
        ' the center of projection so the point
        ' will be clipped.
        Result(3) = INFINITY
    End If
    Result(4) = 1#
End Sub




' Apply a transformation matrix to a point.
Public Sub m3Apply(V() As Single, M() As Single, Result() As Single)
    Result(1) = V(1) * M(1, 1) + _
                V(2) * M(2, 1) + _
                V(3) * M(3, 1) + M(4, 1)
    Result(2) = V(1) * M(1, 2) + _
                V(2) * M(2, 2) + _
                V(3) * M(3, 2) + M(4, 2)
    Result(3) = V(1) * M(1, 3) + _
                V(2) * M(2, 3) + _
                V(3) * M(3, 3) + M(4, 3)
    Result(4) = 1#
End Sub

' Multiply two matrices together. The matrices
' may not contain 0, 0, 0, 1 in their last
' columns.
Public Sub m3MatMultiplyFull(Result() As Single, A() As Single, B() As Single)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim value As Single

    For i = 1 To 4
        For j = 1 To 4
            value = 0#
            For k = 1 To 4
                value = value + A(i, k) * B(k, j)
            Next k
            Result(i, j) = value
        Next j
    Next i
End Sub
' Multiply two matrices together.
Public Sub m3MatMultiply(Result() As Single, A() As Single, B() As Single)
    Result(1, 1) = A(1, 1) * B(1, 1) + A(1, 2) * B(2, 1) + A(1, 3) * B(3, 1)
    Result(1, 2) = A(1, 1) * B(1, 2) + A(1, 2) * B(2, 2) + A(1, 3) * B(3, 2)
    Result(1, 3) = A(1, 1) * B(1, 3) + A(1, 2) * B(2, 3) + A(1, 3) * B(3, 3)
    Result(1, 4) = 0#
    Result(2, 1) = A(2, 1) * B(1, 1) + A(2, 2) * B(2, 1) + A(2, 3) * B(3, 1)
    Result(2, 2) = A(2, 1) * B(1, 2) + A(2, 2) * B(2, 2) + A(2, 3) * B(3, 2)
    Result(2, 3) = A(2, 1) * B(1, 3) + A(2, 2) * B(2, 3) + A(2, 3) * B(3, 3)
    Result(2, 4) = 0#
    Result(3, 1) = A(3, 1) * B(1, 1) + A(3, 2) * B(2, 1) + A(3, 3) * B(3, 1)
    Result(3, 2) = A(3, 1) * B(1, 2) + A(3, 2) * B(2, 2) + A(3, 3) * B(3, 2)
    Result(3, 3) = A(3, 1) * B(1, 3) + A(3, 2) * B(2, 3) + A(3, 3) * B(3, 3)
    Result(3, 4) = 0#
    Result(4, 1) = A(4, 1) * B(1, 1) + A(4, 2) * B(2, 1) + A(4, 3) * B(3, 1) + B(4, 1)
    Result(4, 2) = A(4, 1) * B(1, 2) + A(4, 2) * B(2, 2) + A(4, 3) * B(3, 2) + B(4, 2)
    Result(4, 3) = A(4, 1) * B(1, 3) + A(4, 2) * B(2, 3) + A(4, 3) * B(3, 3) + B(4, 3)
    Result(4, 4) = 1#
End Sub

' Compute the cross product of two vectors.
' Set <x, y, z> = <x1, y1, z1> X <x2, y2, z2>.
Public Sub m3Cross(X As Single, Y As Single, Z As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single)
    X = y1 * z2 - z1 * y2
    Y = z1 * x2 - x1 * z2
    Z = x1 * y2 - y1 * x2
End Sub

' Give the vector the indicated length.
Public Sub m3SizeVector(ByVal l As Single, X As Single, Y As Single, Z As Single)
    l = l / Sqr(X * X + Y * Y + Z * Z)
    X = X * l
    Y = Y * l
    Z = Z * l
End Sub
