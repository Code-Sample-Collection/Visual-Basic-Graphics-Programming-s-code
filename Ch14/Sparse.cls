VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SparseGrid3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TheGrid As Grid3d   ' The display grid.

Private NumPts As Integer   ' # actual data values.
Private Data() As Point3D   ' Actual data values.

Public ShowData As Boolean  ' Draw the actual data?
' Return the index of the closest data point.
Private Function NearestPoint(ByVal X As Single, ByVal Z As Single) As Integer
Dim i As Integer
Dim best_i As Integer
Dim best_dist2 As Single
Dim diffx As Single
Dim diffz As Single
Dim dist2 As Single

    ' Start with no data point.
    best_i = 0
    best_dist2 = 1E+30

    ' See which points are closer.
    For i = 1 To NumPts
        ' See if this point is closer than the ones
        ' already chosen.
        diffx = X - Data(i).coord(1)
        diffz = Z - Data(i).coord(3)
        dist2 = diffx * diffx + diffz * diffz
        If dist2 < best_dist2 Then
            best_i = i
            best_dist2 = dist2
        End If
    Next i

    NearestPoint = best_i
End Function

' Create the grid values for display.
'
' Dx and Dz tell how far apart to make the grid
' lines.
Public Sub InitializeGrid(ByVal Dx As Single, ByVal Dz As Single)
Dim Xmin As Single
Dim Xmax As Single
Dim Zmin As Single
Dim Zmax As Single
Dim NumX As Integer
Dim NumZ As Integer
Dim wid As Single
Dim hgt As Single
Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim best_i As Integer

    ' Find the X and Z data bounds.
    Xmin = Data(1).coord(1)
    Xmax = Xmin
    Zmin = Data(1).coord(3)
    Zmax = Zmin
    For i = 2 To NumPts
        If Xmin > Data(i).coord(1) Then Xmin = Data(i).coord(1)
        If Xmax < Data(i).coord(1) Then Xmax = Data(i).coord(1)
        If Zmin > Data(i).coord(3) Then Zmin = Data(i).coord(3)
        If Zmax < Data(i).coord(3) Then Zmax = Data(i).coord(3)
    Next i

    ' Set the data boundaries.
    wid = Xmax - Xmin
    hgt = Zmax - Zmin
    NumX = wid / Dx + 1
    NumZ = hgt / Dz + 1
    X = (wid - NumX * Dx) / 2
    Z = (hgt - NumZ * Dz) / 2
    Xmin = Xmin - X
    Xmax = Xmax + X
    Zmin = Zmin - Z
    Zmax = Zmax + Z
    
    ' Create the new grid object.
    Set TheGrid = New Grid3d
    TheGrid.SetBounds Xmin, Dx, NumX, Zmin, Dz, NumZ

    ' Fill in data values.
    X = Xmin
    For i = 1 To NumX
        Z = Zmin
        For j = 1 To NumZ
            ' Find the closest data value.
            best_i = NearestPoint(X, Z)

            ' Add the value to the grid.
            TheGrid.SetValue X, Data(best_i).coord(2), Z
            Z = Z + Dz
        Next j
        X = X + Dx
    Next i
End Sub

' Set a data value.
Public Sub SetValue(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    NumPts = NumPts + 1
    ReDim Preserve Data(1 To NumPts)
    With Data(NumPts)
        .coord(1) = X
        .coord(2) = Y
        .coord(3) = Z
        .coord(4) = 1#
    End With
End Sub
' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer

    ' Apply the matrix to the grid if it exists.
    If Not TheGrid Is Nothing Then TheGrid.ApplyFull M

    ' Apply the matrix to the sparse data.
    For i = 1 To NumPts
        m3ApplyFull Data(i).coord, M, Data(i).trans
    Next i
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer
    
    ' Apply the matrix to the grid if it exists.
    If Not TheGrid Is Nothing Then TheGrid.Apply M

    ' Apply the matrix to the sparse data.
    For i = 1 To NumPts
        m3Apply Data(i).coord, M, Data(i).trans
    Next i
End Sub


' Draw the transformed points on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, Optional R As Variant)
Dim i As Integer
    
    ' Draw the grid if it exists.
    If Not TheGrid Is Nothing Then TheGrid.Draw pic

    ' Draw the original data points if desired.
    If ShowData Then
        pic.FillStyle = vbFSSolid
        pic.FillColor = vbRed
        On Error Resume Next
        For i = 1 To NumPts
            pic.Circle (Data(i).trans(1), Data(i).trans(2)), 3, vbRed
        Next i
    End If
End Sub
