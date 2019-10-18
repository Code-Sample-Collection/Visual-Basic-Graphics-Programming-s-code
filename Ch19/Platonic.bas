Attribute VB_Name = "PlatonicSolids"
Option Explicit
' Fill in the points and segments for a cube.
Public Sub GetCube(ByRef pline As Polyline3d)
    ' Allocate the polyline.
    Set pline = New Polyline3d

    ' Make the points.
    ' Y = 1.
    pline.AddNewPoint -1, 1, -1     ' A
    pline.AddNewPoint 1, 1, -1      ' B
    pline.AddNewPoint 1, 1, 1       ' C
    pline.AddNewPoint -1, 1, 1      ' D
    ' Y = -1.
    pline.AddNewPoint -1, -1, -1    ' E
    pline.AddNewPoint 1, -1, -1     ' F
    pline.AddNewPoint 1, -1, 1      ' G
    pline.AddNewPoint -1, -1, 1     ' H

    ' Connect the points.
    ' Z = -1
    pline.AddNewSegment 1, 2
    pline.AddNewSegment 2, 3
    pline.AddNewSegment 3, 4
    pline.AddNewSegment 4, 1
    ' Z = 1
    pline.AddNewSegment 5, 6
    pline.AddNewSegment 6, 7
    pline.AddNewSegment 7, 8
    pline.AddNewSegment 8, 5
    ' Connect the layers.
    pline.AddNewSegment 1, 5
    pline.AddNewSegment 2, 6
    pline.AddNewSegment 3, 7
    pline.AddNewSegment 4, 8
End Sub
' Fill in the points and segments for a tetrahedron.
Public Sub GetTetrahedron(ByRef pline As Polyline3d)
    ' Allocate the polyline.
    Set pline = New Polyline3d

    ' Make the points.
    pline.AddNewPoint 0, 1 / Sqr(2 / 3), 0                              ' A
    pline.AddNewPoint 2 / Sqr(3), 1 / Sqr(2 / 3) - 2 * Sqr(2 / 3), 0    ' B
    pline.AddNewPoint -1 / Sqr(3), 1 / Sqr(2 / 3) - 2 * Sqr(2 / 3), -1  ' C
    pline.AddNewPoint -1 / Sqr(3), 1 / Sqr(2 / 3) - 2 * Sqr(2 / 3), 1   ' D

    ' Connect the points.
    pline.AddNewSegment 1, 2  ' AB
    pline.AddNewSegment 1, 3  ' AC
    pline.AddNewSegment 1, 4  ' AD
    pline.AddNewSegment 2, 3  ' BC
    pline.AddNewSegment 3, 4  ' CD
    pline.AddNewSegment 4, 2  ' DB
End Sub
' Fill in the points and segments for a octahedron.
Public Sub GetOctahedron(ByRef pline As Polyline3d)
Dim i As Integer

    ' Allocate the polyline.
    Set pline = New Polyline3d

    ' Make the points.
    pline.AddNewPoint 0, 1, 0   ' A
    pline.AddNewPoint 1, 0, 0   ' B
    pline.AddNewPoint 0, 0, -1  ' C
    pline.AddNewPoint -1, 0, 0  ' D
    pline.AddNewPoint 0, 0, 1   ' E
    pline.AddNewPoint 0, -1, 0  ' F

    ' Connect the points.
    For i = 2 To 5
        pline.AddNewSegment 1, i
        pline.AddNewSegment 6, i
    Next i
    pline.AddNewSegment 2, 3
    pline.AddNewSegment 3, 4
    pline.AddNewSegment 4, 5
    pline.AddNewSegment 5, 2
End Sub
' Fill in the points and segments for a dodecahedron.
Public Sub GetDodecahedron(ByRef pline As Polyline3d)
Const S = 1 ' The side length.

Dim theta1 As Single
Dim theta2 As Single
Dim theta3 As Single
Dim theta4 As Single
Dim d1 As Single
Dim d2 As Single
Dim d3 As Single
Dim d4 As Single
Dim d5 As Single
Dim Fx As Single
Dim Fy As Single
Dim Ay As Single

    ' Get the basic geometric values.
    theta1 = 2 * PI / 5
    theta2 = PI / 10
    theta3 = 3 * PI / 10
    theta4 = PI / 5
    d1 = S / 2 / Sin(theta4)
    d2 = Cos(theta4) * d1
    d3 = Cos(theta2) * d1
    d4 = Sin(theta2) * d1
    Fx = (S ^ 2 - (2 * d3) ^ 2 - (d1 ^ 2 - d4 ^ 2 - d3 ^ 2)) / _
        2 / (d4 - d1)
    d5 = Sqr((S ^ 2 + (2 * d3) ^ 2 - (d1 - Fx) ^ 2 - (d4 - Fx) ^ 2 - d3 ^ 2) / 2)
    Fy = (Fx ^ 2 - d1 ^ 2 - d5 ^ 2) / (2 * d5)
    Ay = d5 + Fy

    ' Allocate the polyline.
    Set pline = New Polyline3d

    ' Make the points.
    pline.AddNewPoint d1, Ay, 0                                  ' A
    pline.AddNewPoint d4, Ay, d3                                 ' B
    pline.AddNewPoint -d2, Ay, S / 2                             ' C
    pline.AddNewPoint -d2, Ay, -S / 2                            ' D
    pline.AddNewPoint d4, Ay, -d3                                ' E
    pline.AddNewPoint Fx, Fy, 0                                  ' F
    pline.AddNewPoint Fx * Sin(theta2), Fy, Fx * Cos(theta2)     ' G
    pline.AddNewPoint -Fx * Sin(theta3), Fy, Fx * Cos(theta3)    ' H
    pline.AddNewPoint -Fx * Sin(theta3), Fy, -Fx * Cos(theta3)   ' I
    pline.AddNewPoint Fx * Sin(theta2), Fy, -Fx * Cos(theta2)    ' J
    pline.AddNewPoint Fx * Sin(theta3), -Fy, Fx * Cos(theta3)    ' K
    pline.AddNewPoint -Fx * Sin(theta2), -Fy, Fx * Cos(theta2)   ' L
    pline.AddNewPoint -Fx, -Fy, 0                                ' M
    pline.AddNewPoint -Fx * Sin(theta2), -Fy, -Fx * Cos(theta2)  ' N
    pline.AddNewPoint Fx * Sin(theta3), -Fy, -Fx * Cos(theta3)   ' O
    pline.AddNewPoint d2, -Ay, S / 2                             ' P
    pline.AddNewPoint -d4, -Ay, d3                               ' Q
    pline.AddNewPoint -d1, -Ay, 0                                ' R
    pline.AddNewPoint -d4, -Ay, -d3                              ' S
    pline.AddNewPoint d2, -Ay, -S / 2                            ' T

    ' Connect the points.
    pline.AddNewSegment 1, 2   ' AB
    pline.AddNewSegment 2, 3   ' BC
    pline.AddNewSegment 3, 4   ' CD
    pline.AddNewSegment 4, 5   ' DE
    pline.AddNewSegment 5, 1   ' EA
    pline.AddNewSegment 1, 6   ' AF
    pline.AddNewSegment 2, 7   ' BG
    pline.AddNewSegment 3, 8   ' CH
    pline.AddNewSegment 4, 9   ' DI
    pline.AddNewSegment 5, 10  ' EJ
    pline.AddNewSegment 6, 11  ' FK
    pline.AddNewSegment 11, 7  ' KG
    pline.AddNewSegment 7, 12  ' GL
    pline.AddNewSegment 12, 8  ' LH
    pline.AddNewSegment 8, 13  ' HM
    pline.AddNewSegment 13, 9  ' MI
    pline.AddNewSegment 9, 14  ' IN
    pline.AddNewSegment 14, 10 ' NJ
    pline.AddNewSegment 10, 15 ' JO
    pline.AddNewSegment 15, 6  ' OF
    pline.AddNewSegment 16, 11 ' PK
    pline.AddNewSegment 17, 12 ' QL
    pline.AddNewSegment 18, 13 ' RM
    pline.AddNewSegment 19, 14 ' SN
    pline.AddNewSegment 20, 15 ' TO
    pline.AddNewSegment 16, 17 ' PQ
    pline.AddNewSegment 17, 18 ' QR
    pline.AddNewSegment 18, 19 ' RS
    pline.AddNewSegment 19, 20 ' ST
    pline.AddNewSegment 20, 16 ' TP
End Sub
' Fill in the points and segments for an icosahedron.
Public Sub GetIcosahedron(ByRef pline As Polyline3d)
Const S = 1 ' The side length.

Dim theta1 As Single
Dim theta2 As Single
Dim theta3 As Single
Dim theta4 As Single
Dim d1 As Single
Dim d2 As Single
Dim d3 As Single
Dim d4 As Single
Dim d6 As Single
Dim d7 As Single
Dim d8 As Single
Dim Ay As Single
Dim By As Single

    ' Get the basic geometric values.
    theta1 = 2 * PI / 5
    theta2 = PI / 10
    theta3 = 3 * PI / 10
    theta4 = PI / 5
    d1 = S / 2 / Sin(theta4)
    d2 = Cos(theta4) * d1
    d3 = Cos(theta2) * d1
    d4 = Sin(theta2) * d1
    d6 = Sqr(S ^ 2 - d1 ^ 2)
    d7 = Sqr((d1 + d2) ^ 2 - d2 ^ 2)
    d8 = d7 - d6
    By = d8 / 2
    Ay = By + d6

    ' Allocate the polyline.
    Set pline = New Polyline3d

    ' Make the points.
    pline.AddNewPoint 0, Ay, 0         ' A
    pline.AddNewPoint d1, By, 0        ' B
    pline.AddNewPoint d4, By, d3       ' C
    pline.AddNewPoint -d2, By, S / 2   ' D
    pline.AddNewPoint -d2, By, -S / 2  ' E
    pline.AddNewPoint d4, By, -d3      ' F
    pline.AddNewPoint d2, -By, S / 2   ' G
    pline.AddNewPoint -d4, -By, d3     ' H
    pline.AddNewPoint -d1, -By, 0      ' I
    pline.AddNewPoint -d4, -By, -d3    ' J
    pline.AddNewPoint d2, -By, -S / 2  ' K
    pline.AddNewPoint 0, -Ay, 0        ' L

    ' Connect the points.
    pline.AddNewSegment 1, 2   ' AB
    pline.AddNewSegment 1, 3   ' AC
    pline.AddNewSegment 1, 4   ' AD
    pline.AddNewSegment 1, 5   ' AE
    pline.AddNewSegment 1, 6   ' AF
    pline.AddNewSegment 2, 3   ' BC
    pline.AddNewSegment 3, 4   ' CD
    pline.AddNewSegment 4, 5   ' DE
    pline.AddNewSegment 5, 6   ' EF
    pline.AddNewSegment 6, 2   ' FB
    pline.AddNewSegment 2, 7   ' BG
    pline.AddNewSegment 7, 3   ' GC
    pline.AddNewSegment 3, 8   ' CH
    pline.AddNewSegment 8, 4   ' HD
    pline.AddNewSegment 4, 9   ' DI
    pline.AddNewSegment 9, 5   ' IE
    pline.AddNewSegment 5, 10  ' EJ
    pline.AddNewSegment 10, 6  ' JF
    pline.AddNewSegment 6, 11  ' FK
    pline.AddNewSegment 11, 2  ' KB
    pline.AddNewSegment 7, 8   ' GH
    pline.AddNewSegment 8, 9   ' HI
    pline.AddNewSegment 9, 10  ' IJ
    pline.AddNewSegment 10, 11 ' JK
    pline.AddNewSegment 11, 7  ' KG
    pline.AddNewSegment 7, 12  ' GL
    pline.AddNewSegment 8, 12  ' HL
    pline.AddNewSegment 9, 12  ' IL
    pline.AddNewSegment 10, 12 ' JL
    pline.AddNewSegment 11, 12 ' KL
End Sub
' Verify that all the segments have the same
' length and all the points are the same
' distance from the origin.
Public Function SolidOk(ByVal num_points As Integer, Points() As Point3D, ByVal num_segments As Integer, Segments() As Segment3D) As Boolean
Const TINY = 0.0001
Dim i As Integer
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim dist_squared As Single

    ' Verify that all the segments have the
    ' same length.
    dx = Points(Segments(1).pt1).coord(1) - Points(Segments(1).pt2).coord(1)
    dy = Points(Segments(1).pt1).coord(2) - Points(Segments(1).pt2).coord(2)
    dz = Points(Segments(1).pt1).coord(3) - Points(Segments(1).pt2).coord(3)
    dist_squared = dx * dx + dy * dy + dz * dz
    For i = 2 To num_segments
        dx = Points(Segments(i).pt1).coord(1) - Points(Segments(i).pt2).coord(1)
        dy = Points(Segments(i).pt1).coord(2) - Points(Segments(i).pt2).coord(2)
        dz = Points(Segments(i).pt1).coord(3) - Points(Segments(i).pt2).coord(3)
        If Abs(dist_squared - (dx * dx + dy * dy + dz * dz)) _
            > TINY _
        Then
            SolidOk = False
            Exit Function
        End If
    Next i

    ' Verify that all the points are the same
    ' distance from the origin.
    dist_squared = _
        Points(1).coord(1) * Points(1).coord(1) + _
        Points(1).coord(2) * Points(1).coord(2) + _
        Points(1).coord(3) * Points(1).coord(3)
    For i = 2 To num_points
        If Abs(dist_squared - ( _
            Points(1).coord(1) * Points(1).coord(1) + _
            Points(1).coord(2) * Points(1).coord(2) + _
            Points(1).coord(3) * Points(1).coord(3))) _
                > TINY _
        Then
            SolidOk = False
            Exit Function
        End If
    Next i

    SolidOk = True
End Function
