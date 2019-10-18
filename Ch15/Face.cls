VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Face3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Point3D is defined in module M3OPS.BAS as:
'    Type Point3D
'        coord(1 To 4) As Single
'        trans(1 To 4) As Single
'    End Type

Public NumPts As Long       ' Number of points.
Private Points() As Point3D ' Data points.

Public IsCulled As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

' Return True if this polygon partially obscures
' (has greater Z value than) polygon target.
'
' We assume one polygon may obscure the other, but
' they cannot obscure each other.
'
' This check is executed by seeing where the
' projections of the edges of the polygons cross.
' Where they cross, see if one Z value is greater
' than the other.
'
' If no edges cross, see if one polygon contains
' the other. If so, there is an overlap.
Public Function Obscures(ByVal target As Face3d) As Boolean
Dim num As Integer
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

    num = target.NumPts

    ' Check each edge in this polygon.
    GetTransformedPoint NumPts, xi1, yi1, zi1
    For i = 1 To NumPts
        GetTransformedPoint i, xi2, yi2, zi2

        ' Compare with each edge in the other.
        target.GetTransformedPoint num, xj1, yj1, zj1
        For j = 1 To num
            target.GetTransformedPoint j, xj2, yj2, zj2
            ' See if the segments cross.
            If FindCrossing( _
                xi1, yi1, zi1, _
                xi2, yi2, zi2, _
                xj1, yj1, zj1, _
                xj2, yj2, zj2, _
                X, Y, z1, z2) _
            Then
                If z1 - z2 > 0.01 Then
                    ' z1 > z2. We obscure it.
                    Obscures = True
                    Exit Function
                End If
                If z2 - z1 > 0.01 Then
                    ' z2 > z1. It obscures us.
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

    ' If any points of one polygon are inside the
    ' other, then they must all be. Since the
    ' IsAbove tests were inconclusive, some points
    ' in one polygon are on the "bad" side of the
    ' other. In that case there is an overlap.

    ' See if this polygon is inside the other.
    GetTransformedPoint 1, xi1, yi1, zi1
    If target.PointInside(xi1, yi1) Then
        Obscures = True
        Exit Function
    End If

    ' See if the other polygon is inside this one.
    target.GetTransformedPoint 1, xi1, yi1, zi1
    If PointInside(xi1, yi1) Then
        Obscures = True
        Exit Function
    End If

    Obscures = False
End Function
' See where the projections of two segments cross.
' Return true if the segments cross, false
' otherwise.
Private Function FindCrossing( _
    ByVal ax1 As Single, ByVal ay1 As Single, ByVal az1 As Single, _
    ByVal ax2 As Single, ByVal ay2 As Single, ByVal az2 As Single, _
    ByVal bx1 As Single, ByVal by1 As Single, ByVal bz1 As Single, _
    ByVal bx2 As Single, ByVal by2 As Single, ByVal bz2 As Single, _
    ByRef X As Single, ByRef Y As Single, ByRef z1 As Single, ByRef z2 As Single) _
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

' Return True if the point projection lies within
' this polygon's projection.
Public Function PointInside(ByVal X As Single, ByVal Y As Single) As Boolean
Dim i As Integer
Dim theta1 As Double
Dim theta2 As Double
Dim dtheta As Double
Dim dx As Double
Dim dy As Double
Dim angles As Double

    dx = Points(NumPts).trans(1) - X
    dy = Points(NumPts).trans(2) - Y
    theta1 = ATan2(CSng(dy), CSng(dx))
    If theta1 < 0 Then theta1 = theta1 + 2 * PI
    For i = 1 To NumPts
        dx = Points(i).trans(1) - X
        dy = Points(i).trans(2) - Y
        theta2 = ATan2(CSng(dy), CSng(dx))
        If theta2 < 0 Then theta2 = theta2 + 2 * PI
        dtheta = theta2 - theta1
        If dtheta > PI Then dtheta = dtheta - 2 * PI
        If dtheta < -PI Then dtheta = dtheta + 2 * PI
        angles = angles + dtheta
        theta1 = theta2
    Next i

    PointInside = (Abs(angles) > 0.001)
End Function

' Return True if this polygon is completly above
' the plane containing target.
Public Function IsAbove(ByVal target As Face3d) As Boolean
Dim nx As Single
Dim ny As Single
Dim nz As Single
Dim px As Single
Dim py As Single
Dim pz As Single
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim cx As Single
Dim cy As Single
Dim cz As Single
Dim i As Integer

    ' Compute an upward pointing normal to the plane.
    target.TransformedNormalVector nx, ny, nz
    If nz < 0 Then
        nx = -nx
        ny = -ny
        nz = -nz
    End If

    ' Get a point on the plane.
    target.GetTransformedPoint 1, px, py, pz

    ' See if the points in this polygon all lie
    ' above the plane containing target.
    For i = 1 To NumPts
        ' Get the vector from plane to point.
        dx = Points(i).trans(1) - px
        dy = Points(i).trans(2) - py
        dz = Points(i).trans(3) - pz

        ' If the dot product < 0, the point is
        ' below the plane.
        If dx * nx + dy * ny + dz * nz < -0.01 Then
            IsAbove = False
            Exit Function
        End If
    Next i
    IsAbove = True
End Function
' Return true if this polygon is completly below
' the plane containing target.
Public Function IsBelow(ByVal target As Face3d) As Boolean
Dim nx As Single
Dim ny As Single
Dim nz As Single
Dim px As Single
Dim py As Single
Dim pz As Single
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim cx As Single
Dim cy As Single
Dim cz As Single
Dim i As Integer

    ' Compute a downward pointing normal to the plane.
    target.TransformedNormalVector nx, ny, nz
    If nz > 0 Then
        nx = -nx
        ny = -ny
        nz = -nz
    End If

    ' Get a point on the plane.
    target.GetTransformedPoint 1, px, py, pz

    ' See if the points in this polygon all lie
    ' below the plane containing target.
    For i = 1 To NumPts
        ' Get the vector from plane to point.
        dx = Points(i).trans(1) - px
        dy = Points(i).trans(2) - py
        dz = Points(i).trans(3) - pz

        ' If the dot product < 0, the point is
        ' below the plane.
        If dx * nx + dy * ny + dz * nz < -0.01 Then
            IsBelow = False
            Exit Function
        End If
    Next i
    IsBelow = True
End Function
' Return the transformed coordinates of a point
' on the polygon.
Public Sub GetTransformedPoint(ByVal Index As Long, ByRef X As Single, ByRef Y As Single, ByRef z As Single)
    X = Points(Index).trans(1)
    Y = Points(Index).trans(2)
    z = Points(Index).trans(3)
End Sub
' Return the bounds of this polygon.
Public Sub GetExtent(ByRef Xmin As Single, ByRef xmax As Single, ByRef ymin As Single, ByRef ymax As Single, ByRef zmin As Single, ByRef zmax As Single)
Dim i As Integer

    If NumPts < 1 Then Exit Sub

    With Points(1)
        Xmin = .trans(1)
        xmax = Xmin
        ymin = .trans(2)
        ymax = ymin
        zmin = .trans(3)
        zmax = zmin
    End With

    For i = 2 To NumPts
        With Points(i)
            If Xmin > .trans(1) Then Xmin = .trans(1)
            If xmax < .trans(1) Then xmax = .trans(1)
            If ymin > .trans(2) Then ymin = .trans(2)
            If ymax < .trans(2) Then ymax = .trans(2)
            If zmin > .trans(3) Then zmin = .trans(3)
            If zmax < .trans(3) Then zmax = .trans(3)
        End With
    Next i
End Sub


' Compute a normal vector for this polygon.
Public Sub NormalVector(ByRef nx As Single, ByRef ny As Single, ByRef nz As Single)
Dim Ax As Single
Dim Ay As Single
Dim Az As Single
Dim Bx As Single
Dim By As Single
Dim Bz As Single

    Ax = Points(2).coord(1) - Points(1).coord(1)
    Ay = Points(2).coord(2) - Points(1).coord(2)
    Az = Points(2).coord(3) - Points(1).coord(3)
    Bx = Points(3).coord(1) - Points(2).coord(1)
    By = Points(3).coord(2) - Points(2).coord(2)
    Bz = Points(3).coord(3) - Points(2).coord(3)
    m3Cross nx, ny, nz, Ax, Ay, Az, Bx, By, Bz
End Sub
' Compute a transformed normal vector for this polygon.
Public Sub TransformedNormalVector(ByRef nx As Single, ByRef ny As Single, ByRef nz As Single)
Dim Ax As Single
Dim Ay As Single
Dim Az As Single
Dim Bx As Single
Dim By As Single
Dim Bz As Single

    Ax = Points(2).trans(1) - Points(1).trans(1)
    Ay = Points(2).trans(2) - Points(1).trans(2)
    Az = Points(2).trans(3) - Points(1).trans(3)
    Bx = Points(3).trans(1) - Points(2).trans(1)
    By = Points(3).trans(2) - Points(2).trans(2)
    Bz = Points(3).trans(3) - Points(2).trans(3)
    m3Cross nx, ny, nz, Ax, Ay, Az, Bx, By, Bz
End Sub


' Add one or more points to the polygon.
Public Sub AddPoints(ParamArray coord() As Variant)
Dim num_pts As Integer
Dim i As Integer
Dim pt As Integer

    num_pts = (UBound(coord) + 1) \ 3
    ReDim Preserve Points(1 To NumPts + num_pts)

    pt = 0
    For i = 1 To num_pts
        Points(NumPts + i).coord(1) = coord(pt)
        Points(NumPts + i).coord(2) = coord(pt + 1)
        Points(NumPts + i).coord(3) = coord(pt + 2)
        Points(NumPts + i).coord(4) = 1#
        pt = pt + 3
    Next i

    NumPts = NumPts + num_pts
End Sub
' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer

    ' Do nothing if we are culled.
    If IsCulled Then Exit Sub

    For i = 1 To NumPts
        m3ApplyFull Points(i).coord, M, Points(i).trans
    Next i
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer

    ' Do nothing if we are culled.
    If IsCulled Then Exit Sub

    For i = 1 To NumPts
        m3Apply Points(i).coord, M, Points(i).trans
    Next i
End Sub

' Draw the transformed points on a Form, Printer,
' or PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional r As Variant)
Dim pts() As POINTAPI
Dim i As Integer

    ' Do nothing if we are culled.
    If IsCulled Then Exit Sub
    If NumPts < 3 Then Exit Sub

    ReDim pts(1 To NumPts)
    For i = 1 To NumPts
        pts(i).X = Points(i).trans(1)
        pts(i).Y = Points(i).trans(2)
    Next i

    Polygon pic.hdc, pts(1), NumPts
End Sub
' Cull if any points are behind the center of
' projection.
Public Sub ClipEye(ByVal r As Single)
Dim pt As Integer

    ' Do nothing if we are already culled.
    If IsCulled Then Exit Sub

    For pt = 1 To NumPts
        If Points(pt).trans(3) >= r Then Exit For
    Next pt

    If pt <= NumPts Then IsCulled = True
End Sub
' Perform backface removal for the center
' of projection (X, Y, Z).
Public Sub Cull(ByVal X As Single, ByVal Y As Single, ByVal z As Single)
Dim Ax As Single
Dim Ay As Single
Dim Az As Single
Dim nx As Single
Dim ny As Single
Dim nz As Single

    ' Compute a normal to the face.
    NormalVector nx, ny, nz

    ' Compute a vector from the center of
    ' projection to the face.
    Ax = Points(1).coord(1) - X
    Ay = Points(1).coord(2) - Y
    Az = Points(1).coord(3) - z

    ' See if the vectors meet at an angle < 90.
    IsCulled = (Ax * nx + Ay * ny + Az * nz > -0.0001)
End Sub
' Return the largest transformed Z value for this face.
Public Function zmax() As Single
Dim i As Integer
Dim z_max As Single

    z_max = -1E+30
    If IsCulled Then Exit Function

    For i = 1 To NumPts
        If z_max < Points(i).trans(3) _
            Then z_max = Points(i).trans(3)
    Next i

    zmax = z_max
End Function


