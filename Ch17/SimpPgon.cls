VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimplePolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A simple polygon class.

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public ForeColor As Long
Public FillColor As Long

Public PointX As Collection
Public PointY As Collection
Public PointZ As Collection

' Normal vector.
Public Nx As Single
Public Ny As Single
Public Nz As Single

' Bounding box.
Public Xmin As Single
Public Xmax As Single
Public Ymin As Single
Public Ymax As Single
Public Zmin As Single
Public Zmax As Single

' Compare this polygon to pgon. If it is below us,
' return -1. If it is above us, return 1. If it is
' neither above nor below, return 0.
Public Function CompareToSimplePolygon(ByVal pgon As SimplePolygon) As Integer
Dim i As Integer
Dim Cx As Single
Dim Cy As Single
Dim Cz As Single
Dim Vx As Single
Dim Vy As Single
Dim Vz As Single
Dim old_sign As Integer
Dim new_sign As Integer
Dim dot_product As Single
Dim same_side As Boolean

    ' Get a point on our polygon.
    Cx = PointX(1)
    Cy = PointY(1)
    Cz = PointZ(1)

    ' See if pgon lies on one side or the other.
    With pgon
        ' Assume we will succeed.
        same_side = True

        old_sign = 0
        For i = 1 To pgon.PointX.Count
            ' Get the vector to this point.
            Vx = .PointX(i) - Cx
            Vy = .PointY(i) - Cy
            Vz = .PointZ(i) - Cz

            ' Get the dot product.
            dot_product = Vx * Nx + Vy * Ny + Vz * Nz

            ' See if the dot_product is too
            ' small to be useful.
            If Abs(dot_product) > 0.01 Then
                ' dot_product is big enough to use.
                ' Get the dot product's sign.
                new_sign = Sgn(dot_product)

                ' See if this matches the current sign.
                If old_sign = 0 Then
                    old_sign = new_sign
                ElseIf old_sign <> new_sign Then
                    same_side = False
                    Exit For
                End If
            End If
        Next i
    End With

    ' See if we got a result.
    If same_side Then
        ' We got a result. See which side pgon is on.
        If (old_sign < 0) Then
            ' It's below.
            CompareToSimplePolygon = -1
        Else
            ' It's above.
            CompareToSimplePolygon = 1
        End If
    Else
        CompareToSimplePolygon = 0
    End If
End Function


' Add a point to the polygon, skipping duplicates.
Public Sub AddPoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim num_points As Integer

    If PointX Is Nothing Then
        ' Allocate the coordinate collections.
        Set PointX = New Collection
        Set PointY = New Collection
        Set PointZ = New Collection
    Else
        ' If this is a duplicate point, skip it.
        num_points = PointX.Count
        If (Abs(PointX(num_points) - X) < 0.001) And _
           (Abs(PointY(num_points) - Y) < 0.001) And _
           (Abs(PointZ(num_points) - Z) < 0.001) _
                Then Exit Sub
    End If

    ' Add the new point.
    PointX.Add X
    PointY.Add Y
    PointZ.Add Z
End Sub
' Draw the polygon.
Public Sub DrawPolygon(ByVal pic As PictureBox)
Dim num_points As Long
Dim pts() As POINTAPI
Dim i As Integer
Dim light_source As LightSource

    ' Load the points array.
    num_points = PointX.Count
    ReDim pts(1 To num_points)
    For i = 1 To num_points
        With pts(i)
            .X = PointX(i)
            .Y = PointY(i)
        End With
    Next i

    pic.ForeColor = ForeColor
    pic.FillColor = FillColor

    ' Draw the polygon.
    Polygon pic.hdc, pts(1), num_points
End Sub
' Return True if this is a backface.
Public Function IsBackface() As Boolean
    ' After the transformation (which includes
    ' the projection transformation), the viewing
    ' vector is <0, 0, -EyeR>. Then N dot V is
    ' nx * 0 + ny * 0 + nz * (-EyeR) = -nz * EyeR.
    ' The face a backface if dot product >= 0.
    ' That happens is nz <= 0.
    IsBackface = (Nz <= 0)
End Function
' Prepare for sorting the polygon with others.
Public Sub Finish()
Dim i As Integer
Dim dist As Single

    ' Get the bounding box.
    SetBoundingBox

    ' Get the normal vector.
    If PointX.Count < 3 Then
        Nx = 0
        Ny = 0
        Nz = 0
    Else
        ' Get the normal.
        m3Cross Nx, Ny, Nz, _
            PointX(2) - PointX(1), PointY(2) - PointY(1), PointZ(2) - PointZ(1), _
            PointX(3) - PointX(2), PointY(3) - PointY(2), PointZ(3) - PointZ(2)
        dist = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
        Nx = Nx / dist
        Ny = Ny / dist
        Nz = Nz / dist
    End If
End Sub
' Set the bounding box.
Public Sub SetBoundingBox()
Dim i As Integer

    ' Get the bounding box.
    Xmin = PointX(1): Xmax = Xmin
    Ymin = PointY(1): Ymax = Ymin
    Zmin = PointZ(1): Zmax = Zmin
    For i = 2 To PointX.Count
        If Xmin > PointX(i) Then Xmin = PointX(i)
        If Xmax < PointX(i) Then Xmax = PointX(i)
        If Ymin > PointX(i) Then Ymin = PointX(i)
        If Ymax < PointX(i) Then Ymax = PointX(i)
        If Zmin > PointX(i) Then Zmin = PointX(i)
        If Zmax < PointX(i) Then Zmax = PointX(i)
    Next i
End Sub

' Return true if this polygon is completly above
' the plane containing pgon.
Public Function IsAbove(pgon As SimplePolygon) As Boolean
    IsAbove = (pgon.CompareToSimplePolygon(Me) > 0)
End Function
' Return true if this polygon is completly below
' the plane containing pgon.
Public Function IsBelow(pgon As SimplePolygon) As Boolean
    IsBelow = (pgon.CompareToSimplePolygon(Me) < 0)
End Function


' Return true if this polygon obscures pgon.
'
' 1. Check X and Y extents.
' 2. See if we are below the plane of pgon.
' 3. See if pgon is above our plane.
' 4. See where the projections of the edges of
'    the polygons cross. Where they cross, see
'    if one Z value is greater than the other.
' 5. See if one polygon contains the other.
Public Function Obscures(pgon As SimplePolygon) As Boolean
Dim num_i As Integer
Dim num_j As Integer
Dim i As Integer
Dim j As Integer
Dim xi1 As Single
Dim yi1 As Single
Dim zi1 As Single
Dim xi2 As Single
Dim yi2 As Single
Dim zi2 As Single
Dim xj1 As Single
Dim yj1 As Single
Dim zj1 As Single
Dim xj2 As Single
Dim yj2 As Single
Dim zj2 As Single
Dim X As Single
Dim Y As Single
Dim z1 As Single
Dim z2 As Single

    ' 1. Check X and Y extents.
    Obscures = False
    If Xmin >= pgon.Xmax Then Exit Function
    If Xmax <= pgon.Xmin Then Exit Function
    If Ymin >= pgon.Ymax Then Exit Function
    If Ymax <= pgon.Ymin Then Exit Function

    ' 2. See if we are below the plane of pgon.
    If IsBelow(pgon) Then Exit Function

    ' 3. See if pgon is above our plane.
    If pgon.IsAbove(Me) Then Exit Function

    ' 4. See where the projections of the edges of
    '    the polygons cross. Where they cross, see
    '    if one Z value is greater than the other.
    num_i = PointX.Count

    ' Check each edge in this polygon.
    xi1 = PointX(num_i)
    yi1 = PointY(num_i)
    zi1 = PointZ(num_i)
    For i = 1 To num_i
        xi2 = PointX(i)
        yi2 = PointY(i)
        zi2 = PointZ(i)

        ' Compare the i1-i2 edge with each edge
        ' in pgon.
        num_j = pgon.PointX.Count
        xj1 = pgon.PointX(num_j)
        yj1 = pgon.PointY(num_j)
        zj1 = pgon.PointZ(num_j)
        For j = 1 To num_j
            xj2 = pgon.PointX(j)
            yj2 = pgon.PointY(j)
            zj2 = pgon.PointZ(j)

            ' See if the segments cross.
            If FindCrossing( _
                xi1, yi1, zi1, _
                xi2, yi2, zi2, _
                xj1, yj1, zj1, _
                xj2, yj2, zj2, _
                X, Y, z1, z2) _
            Then
                If z1 - z2 > 0.01 Then
                    ' z1 > z2. We obscure pgon.
                    Obscures = True
                    Exit Function
                End If
                If z2 - z1 > 0.01 Then
                    ' z2 > z1. pgon obscures us.
                    Obscures = False
                    Exit Function
                End If
            End If

            xj1 = xj2
            yj1 = yj2
            zj1 = zj2
        Next j

        xi1 = xi2
        yi1 = yi2
        zi1 = zi2
    Next i

    ' No edges cross. See if one polygon contains
    ' the other.
    '
    ' If any points of one polygon are inside the
    ' other, then they must all be. Since the
    ' IsAbove tests were inconclusive, some points
    ' in one polygon are on the "bad" side of the
    ' other. In that case, we obecure pgon.
    '
    ' See if this polygon is inside the other.
    xi1 = PointX(1)
    yi1 = PointY(1)
    If pgon.PointInside(xi1, yi1) Then
        Obscures = True
        Exit Function
    End If

    ' See if the other polygon is inside this one.
    xi1 = pgon.PointX(1)
    yi1 = pgon.PointY(1)
    If PointInside(xi1, yi1) Then
        Obscures = True
        Exit Function
    End If

    Obscures = False
End Function
' Return true if the point's projection lies within
' this polygon's projection.
Function PointInside(ByVal X As Single, ByVal Y As Single) As Boolean
Dim i As Integer
Dim theta1 As Double
Dim theta2 As Double
Dim dtheta As Double
Dim dx As Double
Dim dy As Double
Dim angles As Double

    dx = PointX(PointX.Count) - X
    dy = PointY(PointY.Count) - Y
    theta1 = ATan2(CSng(dx), CSng(dy))
    If theta1 < 0 Then theta1 = theta1 + 2 * PI
    For i = 1 To PointX.Count
        dx = PointX(i) - X
        dy = PointY(i) - Y
        theta2 = ATan2(CSng(dx), CSng(dy))
        If theta2 < 0 Then theta2 = theta2 + 2 * PI
        dtheta = theta2 - theta1
        If dtheta > PI Then dtheta = dtheta - 2 * PI
        If dtheta < -PI Then dtheta = dtheta + 2 * PI
        angles = angles + dtheta
        theta1 = theta2
    Next i

    PointInside = (Abs(angles) > 0.001)
End Function
' See where the projections of two segments cross.
' Return true if the segments cross, false
' otherwise.
Function FindCrossing( _
    ax1 As Single, ay1 As Single, az1 As Single, _
    ax2 As Single, ay2 As Single, az2 As Single, _
    bx1 As Single, by1 As Single, bz1 As Single, _
    bx2 As Single, by2 As Single, bz2 As Single, _
    X As Single, Y As Single, z1 As Single, z2 As Single) _
        As Boolean
Dim dxa As Single
Dim dya As Single
Dim dza As Single
Dim dxb As Single
Dim dyb As Single
Dim dzb As Single
Dim t1 As Single
Dim t2 As Single
Dim denom As Single

    dxa = ax2 - ax1
    dya = ay2 - ay1
    dxb = bx2 - bx1
    dyb = by2 - by1
    
    FindCrossing = False
    
    denom = dxb * dya - dyb * dxa
    ' If the segments are parallel, stop.
    If denom < 0.01 And denom > -0.01 Then Exit Function

    t2 = (ax1 * dya - ay1 * dxa - bx1 * dya + by1 * dxa) / denom
    If t2 < 0 Or t2 > 1 Then Exit Function
    
    t1 = (ax1 * dyb - ay1 * dxb - bx1 * dyb + by1 * dxb) / denom
    If t1 < 0 Or t1 > 1 Then Exit Function

    ' Compute the points of overlap.
    X = ax1 + t1 * dxa
    Y = ay1 + t1 * dya
    dza = az2 - az1
    dzb = bz2 - bz1
    z1 = az1 + t1 * dza
    z2 = bz1 + t2 * dzb
    FindCrossing = True
End Function
