VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Polyline3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Point3D and Segment3D are defined in module M3OPS.BAS as:
'    Type Point3D
'        coord(1 To 4) As Single
'        trans(1 To 4) As Single
'    End Type
'
'    Type Segment3D
'        pt1 As Integer
'        pt2 As Integer
'    End Type

Private NumPoints As Integer ' Number of points.
Private Points() As Point3D  ' Data points.

Private NumSegs As Integer   ' Number of segments.
Private Segs() As Segment3D  ' The segments.

' Create a pyramid with height L and base given
' by the points in the coord array. Add the
' segments that make up the pyramid to this
' polyline.
Public Sub Stellate(L As Single, ParamArray coord() As Variant)
Dim x0 As Single
Dim y0 As Single
Dim z0 As Single
Dim x1 As Single
Dim y1 As Single
Dim z1 As Single
Dim x2 As Single
Dim y2 As Single
Dim z2 As Single
Dim x3 As Single
Dim y3 As Single
Dim z3 As Single
Dim Ax As Single
Dim Ay As Single
Dim Az As Single
Dim Bx As Single
Dim By As Single
Dim Bz As Single
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim num As Integer
Dim i As Integer
Dim pt As Integer

    num = (UBound(coord) + 1) \ 3
    If num < 3 Then
        MsgBox "Must have at least 3 points to stellate.", , vbExclamation
        Exit Sub
    End If
    
    ' (x0, y0, z0) is the center of the polygon.
    x0 = 0
    y0 = 0
    z0 = 0
    pt = 0
    For i = 1 To num
        x0 = x0 + coord(pt)
        y0 = y0 + coord(pt + 1)
        z0 = z0 + coord(pt + 2)
        pt = pt + 3
    Next i
    x0 = x0 / num
    y0 = y0 / num
    z0 = z0 / num
    
    ' Find the normal.
    x1 = coord(0)
    y1 = coord(1)
    z1 = coord(2)
    x2 = coord(3)
    y2 = coord(4)
    z2 = coord(5)
    x3 = coord(6)
    y3 = coord(7)
    z3 = coord(8)
    Ax = x2 - x1
    Ay = y2 - y1
    Az = z2 - z1
    Bx = x3 - x2
    By = y3 - y2
    Bz = z3 - z2
    m3Cross Nx, Ny, Nz, Ax, Ay, Az, Bx, By, Bz
    
    ' Give the normal length L.
    m3SizeVector L, Nx, Ny, Nz
    
    ' The normal + <x0, y0, z0> gives the point.
    x0 = x0 + Nx
    y0 = y0 + Ny
    z0 = z0 + Nz

    ' Build the segments that make up the object.
    x1 = coord(3 * num - 3)
    y1 = coord(3 * num - 2)
    z1 = coord(3 * num - 1)
    pt = 0
    For i = 1 To num
        x2 = coord(pt)
        y2 = coord(pt + 1)
        z2 = coord(pt + 2)
        AddSegment x1, y1, z1, x2, y2, z2, x0, y0, z0
        x1 = x2
        y1 = y2
        z1 = z2
        pt = pt + 3
    Next i
End Sub

' Add one or more line segments to the polyline.
Public Sub AddSegment(ParamArray coord() As Variant)
Dim num_segs As Integer
Dim i As Integer
Dim last As Integer
Dim pt As Integer

    num_segs = (UBound(coord) + 1) \ 3 - 1
    ReDim Preserve Segs(1 To NumSegs + num_segs)

    last = AddPoint((coord(0)), (coord(1)), (coord(2)))
    pt = 0
    For i = 1 To num_segs
        Segs(NumSegs + i).pt1 = last
        pt = pt + 3
        last = AddPoint((coord(pt)), (coord(pt + 1)), (coord(pt + 2)))
        Segs(NumSegs + i).pt2 = last
    Next i

    NumSegs = NumSegs + num_segs
End Sub
' Add a point to the polyline or reuse a point.
' Return the point's index.
Private Function AddPoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Integer
Dim i As Integer

    ' See if the point is already here.
    For i = 1 To NumPoints
        If X = Points(i).coord(1) And _
           Y = Points(i).coord(2) And _
           Z = Points(i).coord(3) Then _
                Exit For
    Next i
    AddPoint = i
    
    ' If so, we're done.
    If i <= NumPoints Then Exit Function
    
    ' Otherwise create the new point.
    NumPoints = NumPoints + 1
    ReDim Preserve Points(1 To NumPoints)
    Points(i).coord(1) = X
    Points(i).coord(2) = Y
    Points(i).coord(3) = Z
    Points(i).coord(4) = 1#
End Function


' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer

    For i = 1 To NumPoints
        m3ApplyFull Points(i).coord, M, Points(i).trans
    Next i
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer

    For i = 1 To NumPoints
        m3Apply Points(i).coord, M, Points(i).trans
    Next i
End Sub


' Draw the transformed points on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional R As Variant)
Dim seg As Integer
Dim pt1 As Integer
Dim pt2 As Integer
Dim dist As Single

    On Error Resume Next
    If IsMissing(R) Then R = INFINITY
    dist = R
    For seg = 1 To NumSegs
        pt1 = Segs(seg).pt1
        pt2 = Segs(seg).pt2
        ' Don't draw if either point is farther
        ' from the focus point than the center of
        ' projection (which is distance dist away).
        If (Points(pt1).trans(3) < R) And (Points(pt2).trans(3) < R) Then _
            pic.Line _
                (Points(pt1).trans(1), Points(pt1).trans(2))- _
                (Points(pt2).trans(1), Points(pt2).trans(2))
    Next seg
End Sub
