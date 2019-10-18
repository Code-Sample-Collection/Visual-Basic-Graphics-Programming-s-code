VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Transformed3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NumCurvePts As Integer
Private CurvePoints() As Point3D

Private NumTransformations As Integer
Private Transformations() As Transformation

Private Faces As Collection
' Add a point to the curve.
Public Sub AddCurvePoint(ByVal X As Single, ByVal Y As Single, ByVal z As Single)
    NumCurvePts = NumCurvePts + 1
    ReDim Preserve CurvePoints(1 To NumCurvePts)
    With CurvePoints(NumCurvePts)
        .coord(1) = X
        .coord(2) = Y
        .coord(3) = z
        .coord(4) = 1
    End With
End Sub


' Set a transformation.
Public Sub SetTransformation(M() As Single)
    NumTransformations = NumTransformations + 1
    ReDim Preserve Transformations(1 To NumTransformations)
    m3MatCopy Transformations(NumTransformations).M, M
End Sub
' Create the display faces by applying the
' series of transformations in array M().
Public Sub Transform()
Dim i As Integer
Dim j As Integer
Dim x0 As Single
Dim y0 As Single
Dim z0 As Single
Dim x1 As Single
Dim y1 As Single
Dim z1 As Single
Dim x2 As Single
Dim y2 As Single
Dim z2 As Single

    Set ThePolyline = New Polyline3d

    ' Add the original curve to ThePolyline.
    x1 = CurvePoints(1).coord(1)
    y1 = CurvePoints(1).coord(2)
    z1 = CurvePoints(1).coord(3)
    For j = 2 To NumCurvePts
        x2 = CurvePoints(j).coord(1)
        y2 = CurvePoints(j).coord(2)
        z2 = CurvePoints(j).coord(3)
        ThePolyline.AddSegment x1, y1, z1, x2, y2, z2
        x1 = x2
        y1 = y2
        z1 = z2
    Next j

    ' Start with the transformed coordinates
    ' the same as the original coordinates.
    For j = 1 To NumCurvePts
        CurvePoints(j).trans(1) = CurvePoints(j).coord(1)
        CurvePoints(j).trans(2) = CurvePoints(j).coord(2)
        CurvePoints(j).trans(3) = CurvePoints(j).coord(3)
    Next j

    ' Create the transformed copies of the curve.
    For i = 1 To NumTransformations
        ' Place the first point.
        x1 = CurvePoints(1).trans(1)
        y1 = CurvePoints(1).trans(2)
        z1 = CurvePoints(1).trans(3)
        m3ApplyFull _
            CurvePoints(1).coord, _
            Transformations(i).M, _
            CurvePoints(1).trans
        x0 = CurvePoints(1).trans(1)
        y0 = CurvePoints(1).trans(2)
        z0 = CurvePoints(1).trans(3)
        ThePolyline.AddSegment x1, y1, z1, x0, y0, z0

        ' Add the rest of the points.
        For j = 2 To NumCurvePts
            x1 = CurvePoints(j).trans(1)
            y1 = CurvePoints(j).trans(2)
            z1 = CurvePoints(j).trans(3)
            m3ApplyFull _
                CurvePoints(j).coord, _
                Transformations(i).M, _
                CurvePoints(j).trans
            x2 = CurvePoints(j).trans(1)
            y2 = CurvePoints(j).trans(2)
            z2 = CurvePoints(j).trans(3)
            ' (x0, y0, z0) = previous point, new.
            ' (x1, y1, z1) = current point, old.
            ' (x2, y2, z2) = current point, new.
            ThePolyline.AddSegment x0, y0, z0, x2, y2, z2
            ThePolyline.AddSegment x1, y1, z1, x2, y2, z2
            x0 = x2
            y0 = y2
            z0 = z2
        Next j
    Next i
End Sub
' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer

    ' Transform the curve.
    For i = 1 To NumCurvePts
        m3ApplyFull CurvePoints(i).coord, M, _
                    CurvePoints(i).trans
    Next i
    
    ' Transform the display polyline if it exists.
    If Not ThePolyline Is Nothing Then ThePolyline.ApplyFull M
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer

    ' Transform the curve.
    For i = 1 To NumCurvePts
        m3Apply CurvePoints(i).coord, M, _
                CurvePoints(i).trans
    Next i
    
    ' Transform the display polyline if it exists.
    If Not ThePolyline Is Nothing Then ThePolyline.Apply M
End Sub


' Draw the extrusion on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional r As Variant)
    If Not ThePolyline Is Nothing Then _
        ThePolyline.Draw pic, r
End Sub
