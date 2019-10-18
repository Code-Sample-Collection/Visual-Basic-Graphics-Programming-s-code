VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rotated3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NumCurvePts As Integer
Private CurvePoints() As Point3D

Private ThePolyline As Polyline3d    ' The display polyline.
' Add a point to the curve.
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

' Create the display polyline by rotating around
' the Y axis.
Public Sub Rotate()
Dim i As Integer
Dim R As Single
Dim theta As Single
Dim dtheta As Single
Dim T As Single
Dim X As Single
Dim Z As Single
Dim x1 As Single
Dim y1 As Single
Dim z1 As Single
Dim x2 As Single
Dim y2 As Single
Dim z2 As Single

    Set ThePolyline = New Polyline3d
        
    ' Create the translated images of the curve.
    dtheta = PI / 8
    For theta = 0 To 2 * PI - dtheta + 0.01 Step dtheta
        X = CurvePoints(1).coord(1)
        Z = CurvePoints(1).coord(3)
        R = Sqr(X * X + Z * Z)
        T = ATan2(Z, X)
        x1 = R * Cos(T + theta)
        y1 = CurvePoints(1).coord(2)
        z1 = R * Sin(T + theta)
        For i = 2 To NumCurvePts
            X = CurvePoints(i).coord(1)
            Z = CurvePoints(i).coord(3)
            R = Sqr(X * X + Z * Z)
            T = ATan2(Z, X)
            x2 = R * Cos(T + theta)
            y2 = CurvePoints(i).coord(2)
            z2 = R * Sin(T + theta)
            
            ThePolyline.AddSegment x1, y1, z1, x2, y2, z2
            x1 = x2
            y1 = y2
            z1 = z2
        Next i
    Next theta

    ' Create the circles of rotation.
    For i = 1 To NumCurvePts
        X = CurvePoints(i).coord(1)
        Z = CurvePoints(i).coord(3)
        R = Sqr(X * X + Z * Z)
        T = ATan2(Z, X)
        x1 = R * Cos(T)
        y1 = CurvePoints(i).coord(2)
        z1 = R * Sin(T)
        For theta = dtheta To 2 * PI - dtheta + 0.01 Step dtheta
            x2 = R * Cos(T + theta)
            z2 = R * Sin(T + theta)
            ThePolyline.AddSegment x1, y1, z1, x2, y1, z2
            x1 = x2
            z1 = z2
        Next theta
        x2 = R * Cos(T)
        z2 = R * Sin(T)
        ThePolyline.AddSegment x1, y1, z1, x2, y1, z2
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

    ' Transform the display polyline if it exists.
    If Not ThePolyline Is Nothing Then ThePolyline.ApplyFull M
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer

    ' Transform the base curve.
    For i = 1 To NumCurvePts
        m3Apply CurvePoints(i).coord, M, _
                CurvePoints(i).trans
    Next i

    ' Transform the display polyline if it exists.
    If Not ThePolyline Is Nothing Then ThePolyline.Apply M
End Sub


' Draw the extrusion on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional R As Variant)
    If Not ThePolyline Is Nothing Then _
        ThePolyline.Draw pic, R
End Sub
