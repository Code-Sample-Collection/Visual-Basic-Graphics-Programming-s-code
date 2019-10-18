VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Extrusion3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NumCurvePts As Integer
Private NumPathPts As Integer

Private CurvePoints() As Point3D
Private PathPoints() As Point3D

Private ThePolyline As Polyline3d
' Add a point to the generating path.
Public Sub AddPathPoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    NumPathPts = NumPathPts + 1
    ReDim Preserve PathPoints(1 To NumPathPts)

    With PathPoints(NumPathPts)
        .coord(1) = X
        .coord(2) = Y
        .coord(3) = Z
        .coord(4) = 1
    End With
End Sub

' Add a point to the base curve.
Public Sub AddCurvePoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    NumCurvePts = NumCurvePts + 1
    ReDim Preserve CurvePoints(1 To NumCurvePts)

    With CurvePoints(NumCurvePts)
        .coord(1) = X
        .coord(2) = Y
        .coord(3) = Z
        .coord(4) = 1
    End With
End Sub

' Create the display polylines.
Public Sub Extrude()
Dim i As Integer
Dim j As Integer
Dim xoff1 As Single
Dim yoff1 As Single
Dim zoff1 As Single
Dim xoff2 As Single
Dim yoff2 As Single
Dim zoff2 As Single
Dim x1 As Single
Dim y1 As Single
Dim z1 As Single
Dim x2 As Single
Dim y2 As Single
Dim z2 As Single

    Set ThePolyline = New Polyline3d

    ' Create the translated images of the curve.
    For i = 1 To NumPathPts
        ' Calculate offsets for this path point.
        xoff1 = PathPoints(i).coord(1) - PathPoints(1).coord(1)
        yoff1 = PathPoints(i).coord(2) - PathPoints(1).coord(2)
        zoff1 = PathPoints(i).coord(3) - PathPoints(1).coord(3)

        x1 = CurvePoints(1).coord(1) + xoff1
        y1 = CurvePoints(1).coord(2) + yoff1
        z1 = CurvePoints(1).coord(3) + zoff1
        For j = 2 To NumCurvePts
            x2 = CurvePoints(j).coord(1) + xoff1
            y2 = CurvePoints(j).coord(2) + yoff1
            z2 = CurvePoints(j).coord(3) + zoff1
            ThePolyline.AddSegment x1, y1, z1, x2, y2, z2
            x1 = x2
            y1 = y2
            z1 = z2
        Next j
    Next i

    ' Create the translated images of the path.
    xoff1 = PathPoints(1).coord(1) - PathPoints(1).coord(1)
    yoff1 = PathPoints(1).coord(2) - PathPoints(1).coord(2)
    zoff1 = PathPoints(1).coord(3) - PathPoints(1).coord(3)
    For i = 2 To NumPathPts
        ' Calculate offsets for this path point.
        xoff2 = PathPoints(i).coord(1) - PathPoints(1).coord(1)
        yoff2 = PathPoints(i).coord(2) - PathPoints(1).coord(2)
        zoff2 = PathPoints(i).coord(3) - PathPoints(1).coord(3)

        For j = 1 To NumCurvePts
            ThePolyline.AddSegment _
                CurvePoints(j).coord(1) + xoff1, _
                CurvePoints(j).coord(2) + yoff1, _
                CurvePoints(j).coord(3) + zoff1, _
                CurvePoints(j).coord(1) + xoff2, _
                CurvePoints(j).coord(2) + yoff2, _
                CurvePoints(j).coord(3) + zoff2
        Next j
        xoff1 = xoff2
        yoff1 = yoff2
        zoff1 = zoff2
    Next i
End Sub

' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer

    ' Transform the base curve.
    For i = 1 To NumCurvePts
        m3ApplyFull CurvePoints(i).coord, M, _
                    CurvePoints(i).trans
    Next i

    ' Transform the generating path.
    For i = 1 To NumPathPts
        m3ApplyFull PathPoints(i).coord, M, _
                    PathPoints(i).trans
    Next i

    ' Transform the display polyline if it exists.
    If Not (ThePolyline Is Nothing) Then ThePolyline.ApplyFull M
End Sub
' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer

    ' Transform the base curve.
    For i = 1 To NumCurvePts
        m3Apply CurvePoints(i).coord, M, _
                CurvePoints(i).trans
    Next i

    ' Transform the generating path.
    For i = 1 To NumPathPts
        m3Apply PathPoints(i).coord, M, _
                PathPoints(i).trans
    Next i

    ' Transform the display polyline if it exists.
    If Not (ThePolyline Is Nothing) Then ThePolyline.Apply M
End Sub
' Draw the extrusion on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional R As Variant)
    If Not ThePolyline Is Nothing Then _
        ThePolyline.Draw pic, R
End Sub
