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

Private NumTrans As Integer
Private trans() As Transformation

Private solid As Solid3d   ' The display solid.
Public Property Get HideSurfaces() As Boolean
    If Not (solid Is Nothing) Then HideSurfaces = solid.HideSurfaces
End Property

Public Property Let HideSurfaces(ByVal new_value As Boolean)
    If Not (solid Is Nothing) Then solid.HideSurfaces = new_value
End Property

' Create the display solid by applying the
' series of transformations in array M().
Public Sub Transform(Optional cap_ends As Variant)
Dim face As Face3d
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
Dim x3 As Single
Dim y3 As Single
Dim z3 As Single

    If IsMissing(cap_ends) Then cap_ends = True

    Set solid = New Solid3d

    ' Add the base curve to solid assuming the
    ' curve is stored oriented towards the
    ' transformations.
    If cap_ends Then
        Set face = New Face3d
        solid.Faces.Add face
        For i = NumCurvePts - 1 To 1 Step -1
            face.AddPoints _
                CurvePoints(i).coord(1), _
                CurvePoints(i).coord(2), _
                CurvePoints(i).coord(3)
        Next i
    End If

    ' Start with the transformed coordinates
    ' the same as the original coordinates.
    For i = 1 To NumCurvePts
        CurvePoints(i).trans(1) = CurvePoints(i).coord(1)
        CurvePoints(i).trans(2) = CurvePoints(i).coord(2)
        CurvePoints(i).trans(3) = CurvePoints(i).coord(3)
    Next i

    ' Create the transformed copies of the curve.
    For i = 1 To NumTrans
        x0 = CurvePoints(1).trans(1)
        y0 = CurvePoints(1).trans(2)
        z0 = CurvePoints(1).trans(3)
        m3ApplyFull _
            CurvePoints(1).coord, trans(i).M, _
            CurvePoints(1).trans
        x1 = CurvePoints(1).trans(1)
        y1 = CurvePoints(1).trans(2)
        z1 = CurvePoints(1).trans(3)

        For j = 2 To NumCurvePts
            x2 = CurvePoints(j).trans(1)
            y2 = CurvePoints(j).trans(2)
            z2 = CurvePoints(j).trans(3)
            m3ApplyFull _
                CurvePoints(j).coord, trans(i).M, _
                CurvePoints(j).trans
            x3 = CurvePoints(j).trans(1)
            y3 = CurvePoints(j).trans(2)
            z3 = CurvePoints(j).trans(3)

            solid.AddFace _
                x0, y0, z0, _
                x2, y2, z2, _
                x1, y1, z1
            
            solid.AddFace _
                x2, y2, z2, _
                x3, y3, z3, _
                x1, y1, z1
            x0 = x2
            y0 = y2
            z0 = z2
            x1 = x3
            y1 = y3
            z1 = z3
        Next j
    Next i

    ' Add the final curve to solid assuming
    ' the curve is stored oriented towards the
    ' transformations.
    If cap_ends Then
        Set face = New Face3d
        solid.Faces.Add face
        For i = 2 To NumCurvePts
            face.AddPoints _
                CurvePoints(i).trans(1), _
                CurvePoints(i).trans(2), _
                CurvePoints(i).trans(3)
        Next i
    End If
End Sub

' Clip the display solid.
Public Sub ClipEye(ByVal r As Single)
    If Not solid Is Nothing Then solid.ClipEye r
End Sub
' Add a point to the curve.
Public Sub AddCurvePoint(X As Single, Y As Single, z As Single)
    NumCurvePts = NumCurvePts + 1
    ReDim Preserve CurvePoints(1 To NumCurvePts)
    CurvePoints(NumCurvePts).coord(1) = X
    CurvePoints(NumCurvePts).coord(2) = Y
    CurvePoints(NumCurvePts).coord(3) = z
    CurvePoints(NumCurvePts).coord(4) = 1
End Sub




' Set a transformation.
Public Sub SetTransformation(M() As Single)
    NumTrans = NumTrans + 1
    ReDim Preserve trans(1 To NumTrans)
    m3MatCopy trans(NumTrans).M, M
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
    
    ' Transform the display solid if it exists.
    If Not solid Is Nothing Then solid.ApplyFull M
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer

    ' Transform the curve.
    For i = 1 To NumCurvePts
        m3Apply CurvePoints(i).coord, M, _
                CurvePoints(i).trans
    Next i
    
    ' Transform the display solid if it exists.
    If Not solid Is Nothing Then solid.Apply M
End Sub


' Draw the extrusion on a Form, Printer, or
' PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional r As Variant)
    If Not solid Is Nothing Then _
        solid.Draw pic, r
End Sub


' Perform backface removal on the display solid.
Public Sub Cull(ByVal X As Single, ByVal Y As Single, ByVal z As Single)
    If Not solid Is Nothing Then solid.Cull X, Y, z
End Sub

' Set or clear the Culled property for the solid.
Property Let Culled(value As Boolean)
    If Not solid Is Nothing Then solid.Culled = value
End Property
