VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeightedGrid3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TheGrid As Grid3d   ' The display grid.

Private NumPts As Integer   ' # actual data values.
Private Data() As Point3D   ' Actual data values.

Public ShowData As Boolean  ' Draw the actual data?
' Return the index of the nearest point in the
' indicated direction.
Private Function NearestPoint(ByVal X As Single, ByVal Z As Single, ByVal on_left As Boolean, ByVal on_top As Boolean) As Integer
Dim i As Integer
Dim best_i As Integer
Dim best_dist2 As Single
Dim diffx As Single
Dim diffz As Single
Dim dist2 As Single

    ' Start with the first data point.
    best_i = 0
    best_dist2 = 1E+30

    ' See which points are closer.
    For i = 1 To NumPts
        ' See if the point satisfies on_left/on_top.
        If CBool(X < Data(i).coord(1)) = on_left And _
           CBool(Z > Data(i).coord(3)) = on_top Then

            ' See if this point is closer than the
            ' best one so far.
            diffx = X - Data(i).coord(1)
            diffz = Z - Data(i).coord(3)
            dist2 = diffx * diffx + diffz * diffz
            If dist2 < best_dist2 Then
                best_i = i
                best_dist2 = dist2
            End If
        End If
    Next i

    NearestPoint = best_i
End Function
' Return a weighted average for this point's value.
Private Function WeightedAverage(ByVal X As Single, ByVal Z As Single, best_i() As Integer, ByVal num As Integer) As Single
Dim i As Integer
Dim j As Integer
Dim diffx As Single
Dim diffz As Single
Dim dist2(1 To 4) As Single
Dim wgt As Single
Dim tot As Single
Dim Y As Single

    ' Compute the distance squared to each point.
    For i = 1 To num
        diffx = X - Data(best_i(i)).coord(1)
        diffz = Z - Data(best_i(i)).coord(3)
        dist2(i) = diffx * diffx + diffz * diffz
        If dist2(i) = 0 Then
            Y = Data(best_i(i)).coord(2)
            Exit Function
        End If
    Next i

    ' Compute the contribution due to each point.
    Y = 0
    For i = 1 To num
        ' Compute the weight for point i.
        wgt = 1
        For j = 1 To num
            If j <> i Then
                wgt = wgt * dist2(j)
            End If
        Next j
        Y = Y + wgt * Data(best_i(i)).coord(2)
        tot = tot + wgt
    Next i

    WeightedAverage = Y / tot
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
Dim best_i(1 To 4) As Integer
Dim num_close As Integer

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
            ' Find close points to the upper left,
            ' upper right, lower left, and lower
            ' right. Average them.
            num_close = 1
            best_i(num_close) = NearestPoint( _
                X, Z, True, True)
            If best_i(num_close) > 0 Then num_close = num_close + 1

            best_i(num_close) = NearestPoint( _
                X, Z, True, False)
            If best_i(num_close) > 0 Then num_close = num_close + 1

            best_i(num_close) = NearestPoint( _
                X, Z, False, True)
            If best_i(num_close) > 0 Then num_close = num_close + 1

            best_i(num_close) = NearestPoint( _
                X, Z, False, False)
            If best_i(num_close) > 0 Then num_close = num_close + 1

            Y = WeightedAverage(X, Z, best_i, num_close - 1)

            ' Add the value to the grid.
            TheGrid.SetValue X, Y, Z
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
