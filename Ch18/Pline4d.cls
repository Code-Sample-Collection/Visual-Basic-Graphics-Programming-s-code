VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Polyline4d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Point4D and Segment4D are defined in module M4ops.bas as:
'    Type Point4D
'        coord(1 To 4) As Single
'        trans(1 To 4) As Single
'    End Type
'
'    Type Segment4D
'        pt1 As Integer
'        pt2 As Integer
'    End Type

Private NumPoints As Integer ' Number of points.
Private Points() As Point4D  ' Data points.

Private NumSegs As Integer   ' Number of segments.
Private Segs() As Segment4D  ' The segments.
' Add one or more line segments to the polyline.
Public Sub AddSegment(ParamArray coord() As Variant)
Dim num_segs As Integer
Dim i As Integer
Dim last As Integer
Dim pt As Integer

    num_segs = (UBound(coord) + 1) \ 4 - 1
    ReDim Preserve Segs(1 To NumSegs + num_segs)

    last = AddPoint((coord(0)), (coord(1)), (coord(2)), (coord(3)))
    pt = 0
    For i = 1 To num_segs
        Segs(NumSegs + i).pt1 = last
        pt = pt + 4
        last = AddPoint((coord(pt)), (coord(pt + 1)), (coord(pt + 2)), (coord(pt + 3)))
        Segs(NumSegs + i).pt2 = last
    Next i

    NumSegs = NumSegs + num_segs
End Sub
' Add a point to the polyline or reuse a point.
' Return the point's index.
Private Function AddPoint(ByVal X As Single, ByVal y As Single, ByVal z As Single, ByVal W As Single) As Integer
Dim i As Integer

    ' See if the point is already here.
    For i = 1 To NumPoints
        If X = Points(i).coord(1) And _
           y = Points(i).coord(2) And _
           z = Points(i).coord(3) And _
           W = Points(i).coord(4) Then _
                Exit For
    Next i
    AddPoint = i

    ' If so, we're done.
    If i <= NumPoints Then Exit Function

    ' Otherwise create the new point.
    NumPoints = NumPoints + 1
    ReDim Preserve Points(1 To NumPoints)
    Points(i).coord(1) = X
    Points(i).coord(2) = y
    Points(i).coord(3) = z
    Points(i).coord(4) = W
    Points(i).coord(5) = 1#
End Function


' Apply a transformation matrix which may not
' contain 0, 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer

    For i = 1 To NumPoints
        m4ApplyFull Points(i).coord, M, Points(i).trans
    Next i
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer

    For i = 1 To NumPoints
        m4Apply Points(i).coord, M, Points(i).trans
    Next i
End Sub


' Draw the transformed points on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional r As Variant)
Dim seg As Integer
Dim pt1 As Integer
Dim pt2 As Integer
Dim dist As Single

    On Error Resume Next
    If IsMissing(r) Then r = INFINITY
    dist = r
    For seg = 1 To NumSegs
        pt1 = Segs(seg).pt1
        pt2 = Segs(seg).pt2
        ' Don't draw if either point is farther
        ' from the focus point than the center of
        ' projection (which is distance dist away).
        If (Points(pt1).trans(4) < r) And (Points(pt2).trans(4) < r) Then _
            pic.Line _
                (Points(pt1).trans(1), Points(pt1).trans(2))- _
                (Points(pt2).trans(1), Points(pt2).trans(2))
    Next seg
End Sub
' Copy the transformed points into the data points.
Public Sub FixPoints()
Dim i As Integer
Dim j As Integer

    For i = 1 To NumPoints
        For j = 1 To 5
            Points(i).coord(j) = Points(i).trans(j)
        Next j
    Next i
End Sub


