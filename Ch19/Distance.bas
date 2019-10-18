Attribute VB_Name = "Distances"
' Three-dimensional distance functions.

Option Explicit



' Return the distance between two points.
Public Function DistancePointToPoint(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single) As Single
Dim dx As Single
Dim dy As Single
Dim dz As Single

    dx = x2 - x1
    dy = y2 - y1
    dz = z2 - z1
    DistancePointToPoint = Sqr(dx * dx + dy * dy + dz * dz)
End Function
' Return the distance between a point and a line.
Public Function DistancePointToLine(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, ByVal vx2 As Single, ByVal vy2 As Single, ByVal vz2 As Single) As Single
Dim ax As Single
Dim ay As Single
Dim az As Single
Dim len_a_squared As Single
Dim a_dot_v As Single
Dim len_v As Single
Dim a_dot_v_over_len_v As Single

    ax = x2 - x1
    ay = y2 - y1
    az = z2 - z1

    len_a_squared = ax * ax + ay * ay + az * az
    a_dot_v = ax * vx2 + ay * vy2 + az * vz2
    len_v = Sqr(vx2 * vx2 + vy2 * vy2 + vz2 * vz2)
    a_dot_v_over_len_v = a_dot_v / len_v

    DistancePointToLine = Sqr( _
        len_a_squared - a_dot_v_over_len_v * a_dot_v_over_len_v)
End Function
' Return the distance between a point and a plane.
Public Function DistancePointToPlane(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, ByVal nx2 As Single, ByVal ny2 As Single, ByVal nz2 As Single) As Single
Dim ax As Single
Dim ay As Single
Dim az As Single
Dim a_dot_n As Single
Dim len_n As Single

    ax = x2 - x1
    ay = y2 - y1
    az = z2 - z1

    a_dot_n = ax * nx2 + ay * ny2 + az * nz2
    len_n = Sqr(nx2 * nx2 + ny2 * ny2 + nz2 * nz2)

    DistancePointToPlane = Abs(a_dot_n / len_n)
End Function
' Return the distance between two lines.
Public Function DistanceLineToLine(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, ByVal vx1 As Single, ByVal vy1 As Single, ByVal vz1 As Single, ByVal vx2 As Single, ByVal vy2 As Single, ByVal vz2 As Single) As Single
Dim len_a As Single
Dim len_b As Single
Dim a_dot_b As Single
Dim nx As Single
Dim ny As Single
Dim nz As Single

    ' See if the lines are parallel.
    len_a = Sqr(vx1 * vx1 + vy1 * vy1 + vz1 * vz1)
    len_b = Sqr(vx2 * vx2 + vy2 * vy2 + vz2 * vz2)
    a_dot_b = vx1 * vx2 + vy1 * vy2 + vz1 * vz2
    If a_dot_b = len_a * len_b Then
        ' The lines are parallel.
        DistanceLineToLine = _
            DistancePointToLine(x1, y1, z1, _
                x2, y2, z2, vx2, vy2, vz2)
    Else
        ' The lines are not parallel.
        ' Get the normal to both vectors.
        m3Cross nx, ny, nz, vx1, vy1, vz1, vx2, vy2, vz2

        DistanceLineToLine = _
            DistancePointToPlane(x1, y1, z1, _
                x2, y2, z2, nx, ny, nz)
    End If
End Function
