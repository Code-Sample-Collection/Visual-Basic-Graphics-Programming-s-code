VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bezier3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MaxU As Integer     ' Dimensions of control grid.
Private MaxV As Integer
Private Points() As Point3D ' Control points.

' Holds polylines containing the refined
' grid to display the surface.
Private Polylines As Collection

' u and v increment parameters.
Private GapU As Single
Private GapV As Single
Private Du As Single
Private Dv As Single

' Display flags.
Private ShowControls As Boolean ' Draw control points?
Private ShowGrid As Boolean     ' Draw control grid?
' Return the factorial of a number (n!).
Function Factorial(ByVal n As Single) As Single
Dim i As Integer
Dim tot As Single

    tot = 1
    For i = 2 To n
        tot = tot * i
    Next i
    Factorial = tot
End Function

' Create polylines to represent the surface.
Public Sub InitializeGrid(ByVal gap_x As Single, ByVal gap_z As Single, ByVal d_x As Single, ByVal d_z As Single)
Dim u As Single
Dim v As Single
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim x1 As Single
Dim y1 As Single
Dim z1 As Single
Dim pline As Polyline3d

    GapU = gap_x
    GapV = gap_z
    Du = d_x
    Dv = d_z

    Set Polylines = New Collection

    ' Create curves with constant u.
    For u = 0 To 1 + GapU / 10 Step GapU
        Set pline = New Polyline3d
        Polylines.Add pline

        SurfaceValue u, 0, x1, y1, z1
        
        For v = Dv To 1 + Dv / 10 Step Dv
            SurfaceValue u, v, X, Y, Z
            pline.AddSegment x1, y1, z1, X, Y, Z
            x1 = X
            y1 = Y
            z1 = Z
        Next v
    Next u

    ' Create curves with constant v.
    For v = 0 To 1 + GapV / 10 Step GapV
        Set pline = New Polyline3d
        Polylines.Add pline
        
        SurfaceValue 0, v, x1, y1, z1
        For u = Du To 1 + Du / 10 Step Du
            SurfaceValue u, v, X, Y, Z
            pline.AddSegment x1, y1, z1, X, Y, Z
            x1 = X
            y1 = Y
            z1 = Z
        Next u
    Next v
End Sub
' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer
Dim j As Integer
Dim pline As Polyline3d

    ' Apply the matrix to the grid if it exists.
    If Not Polylines Is Nothing Then
        For Each pline In Polylines
            pline.ApplyFull M
        Next pline
    End If

    ' Apply the matrix to the control points.
    For i = 0 To MaxU
        For j = 0 To MaxV
            m3ApplyFull Points(i, j).coord, M, Points(i, j).trans
        Next j
    Next i
End Sub
' Draw the transformed object on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional R As Variant)
Dim i As Integer
Dim j As Integer
Dim pline As Polyline3d

    ' Draw the grid if it exists.
    If Not Polylines Is Nothing Then
        For Each pline In Polylines
            pline.Draw pic, R
        Next pline
    End If

    ' Draw the control points if desired.
    If ShowControls Then
        On Error Resume Next
        For i = 0 To MaxU
            For j = 0 To MaxV
                pic.Line (Points(i, j).trans(1) - 2, Points(i, j).trans(2) - 2)-Step(4, 4), , BF
            Next j
        Next i
    End If

    ' Draw the control grid if desired.
    If ShowGrid Then
        On Error Resume Next
        For i = 0 To MaxU
            pic.CurrentX = Points(i, 0).trans(1)
            pic.CurrentY = Points(i, 0).trans(2)
            For j = 1 To MaxV
                pic.Line -(Points(i, j).trans(1), Points(i, j).trans(2))
            Next j
        Next i
        For j = 0 To MaxV
            pic.CurrentX = Points(0, j).trans(1)
            pic.CurrentY = Points(0, j).trans(2)
            For i = 1 To MaxU
                pic.Line -(Points(i, j).trans(1), Points(i, j).trans(2))
            Next i
        Next j
    End If
End Sub

' Return a value indicating whether we
' are drawing the control grid.
Property Get DrawGrid() As Boolean
    DrawGrid = ShowGrid
End Property

' Return a value indicating whether we
' are drawing the control points.
Property Get DrawControls() As Boolean
    DrawControls = ShowControls
End Property


' Set the value indicating whether we
' are drawing the control grid.
Property Let DrawGrid(ByVal new_value As Boolean)
    ShowGrid = new_value
End Property
' Set the value indicating whether we
' are drawing the control points.
Property Let DrawControls(ByVal new_value As Boolean)
    ShowControls = new_value
End Property




' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer
Dim j As Integer
Dim pline As Polyline3d

    ' Apply the matrix to the polylines.
    If Not Polylines Is Nothing Then
        For Each pline In Polylines
            pline.Apply M
        Next pline
    End If

    ' Apply the matrix to the control points.
    For i = 0 To MaxU
        For j = 0 To MaxV
            m3Apply Points(i, j).coord, M, Points(i, j).trans
        Next j
    Next i
End Sub





' Set MaxU and MaxV and allocate room for the
' control points.
Public Sub SetBounds(ByVal NumX As Integer, ByVal NumZ As Integer)
    MaxU = NumX - 1
    MaxV = NumZ - 1
    ReDim Points(0 To NumX, 0 To NumZ)
End Sub

' Set the value for a control point.
Public Sub SetControlPoint(ByVal i As Integer, ByVal j As Integer, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    Points(i - 1, j - 1).coord(1) = X
    Points(i - 1, j - 1).coord(2) = Y
    Points(i - 1, j - 1).coord(3) = Z
    Points(i - 1, j - 1).coord(4) = 1
End Sub
' Return the (X, Y, Z) coordinates of the
' Bezier surface for these u and v values.
Private Sub SurfaceValue(ByVal u As Single, ByVal v As Single, ByRef X As Single, ByRef Y As Single, ByRef Z As Single)
Dim P As Integer
Dim i As Integer
Dim j As Integer
Dim pt As Point3D
Dim Bix As Single
Dim Bjz As Single

    For i = 0 To MaxU
        ' Compute Bix.
        Bix = Factorial(MaxU) / Factorial(i) / _
            Factorial(MaxU - i) * _
            u ^ i * (1 - u) ^ (MaxU - i)
        
        For j = 0 To MaxV
            ' Compute Bjz.
            Bjz = Factorial(MaxV) / Factorial(j) / _
                Factorial(MaxV - j) * _
                v ^ j * (1 - v) ^ (MaxV - j)
            
            ' Add to the coordinates.
            For P = 1 To 3
                pt.coord(P) = pt.coord(P) + _
                    Points(i, j).coord(P) * _
                    Bix * Bjz
            Next P
        Next j
    Next i
    
    ' Prepare the output.
    X = pt.coord(1)
    Y = pt.coord(2)
    Z = pt.coord(3)
End Sub
