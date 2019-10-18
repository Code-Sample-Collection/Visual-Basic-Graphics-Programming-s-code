VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValleyGrid3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Xmin As Single      ' Min X and Y values.
Private Zmin As Single
Private Dx As Single        ' Spacing between rows of data.
Private Dz As Single
Private NumX As Integer     ' Number of X and Y entries.
Private NumZ As Integer
Private points() As Point3D ' Data values.

Public RemoveHidden As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

' Generate the fractal surface.
Public Sub GenerateSurface(ByVal divisions As Integer, ByVal Dy As Single)
Dim oldpoints() As Point3D
Dim oldx As Integer
Dim oldz As Integer
Dim factor As Integer
Dim newx As Integer
Dim newz As Integer
Dim i As Integer
Dim j As Integer
Dim newi As Integer
Dim newj As Integer

    ' Make room for the new data.
    factor = 2 ^ divisions
    newx = (NumX - 1) * factor + 1
    newz = (NumZ - 1) * factor + 1

    ' Copy the original data.
    ReDim oldpoints(1 To NumX, 1 To NumZ)
    For i = 1 To NumX
        For j = 1 To NumZ
            oldpoints(i, j) = points(i, j)
        Next j
    Next i

    ' Resize and initialize the Points array.
    oldx = NumX
    oldz = NumZ
    SetBounds Xmin, Dx / factor, newx, _
              Zmin, Dz / factor, newz

    ' Move the old data to the new positions.
    newi = 1
    For i = 1 To oldx
        newj = 1
        For j = 1 To oldz
            points(newi, newj) = oldpoints(i, j)
            newj = newj + factor
        Next j
        newi = newi + factor
    Next i

    ' Subdivide each area in the data.
    newi = 1
    For i = 2 To oldx
        newj = 1
        For j = 2 To oldz
            Subdivide newi, newi + factor, _
                      newj, newj + factor, Dy
            newj = newj + factor
        Next j
        newi = newi + factor
    Next i
End Sub
' If a Y value is within distance range of the
' value target_y, then reduce that distance by
' a factor of smooth_factor.
Public Sub Flatten(ByVal target_y As Single, ByVal range As Single, ByVal smooth_factor As Single)
Dim i As Integer
Dim j As Integer
Dim diff As Single

    For i = 1 To NumX
        For j = 1 To NumZ
            With points(i, j)
                diff = .coord(2) - target_y
                If Abs(diff) < range Then
                    .coord(2) = target_y + smooth_factor * diff
                End If
            End With
        Next j
    Next i
End Sub

' Pick a color for this area.
Private Sub PickColor(ByVal pic As PictureBox, ByVal y1 As Single, ByVal y2 As Single, ByVal y3 As Single, ByVal y4 As Single)
Dim min_y As Single
Dim max_y As Single
Dim g As Integer
Dim b As Integer

    min_y = y1
    max_y = y1
    If min_y > y2 Then min_y = y2
    If max_y < y2 Then max_y = y2
    If min_y > y3 Then min_y = y3
    If max_y < y3 Then max_y = y3
    If min_y > y4 Then min_y = y4
    If max_y < y4 Then max_y = y4

    If max_y < -1.09 Then
        ' River.
        pic.FillColor = RGB(0, &H80, &HFF)
    ElseIf max_y > 0.7 Then
        ' Snow.
        pic.FillColor = vbWhite
    ElseIf Abs(min_y - max_y) < 0.3 Then
        ' Vegetation.
        g = &HFF - (min_y + 1) * 70
        b = &H80 - (min_y + 1) * 100
        pic.FillColor = RGB(0, g, b)
    Else
        ' Cliff.
        pic.FillColor = RGB(&HC0, &HC0, &HC0)
    End If
End Sub

' Recursively subdivide the indicated area.
Private Sub Subdivide(ByVal i1 As Integer, ByVal i2 As Integer, ByVal j1 As Integer, ByVal j2 As Integer, ByVal Dy As Single)
Dim y11 As Single
Dim y12 As Single
Dim y21 As Single
Dim y22 As Single
Dim imid As Integer
Dim jmid As Integer

    If (i2 - i1 <= 1) Or (j2 - j1 <= 1) Then Exit Sub

    ' Compute the midpoint locations.
    y11 = points(i1, j1).coord(2)
    y12 = points(i1, j2).coord(2)
    y21 = points(i2, j1).coord(2)
    y22 = points(i2, j2).coord(2)

    imid = (i1 + i2) \ 2
    jmid = (j1 + j2) \ 2
    points(i1, jmid).coord(2) = (y11 + y12) / 2 + 2 * Dy * Rnd - Dy
    points(i2, jmid).coord(2) = (y21 + y22) / 2 + 2 * Dy * Rnd - Dy
    points(imid, j1).coord(2) = (y11 + y21) / 2 + 2 * Dy * Rnd - Dy
    points(imid, j2).coord(2) = (y12 + y22) / 2 + 2 * Dy * Rnd - Dy
    points(imid, jmid).coord(2) = (y11 + y12 + y21 + y22) / 4 + 2 * Dy * Rnd - Dy

    ' Recursively subdivide the four new areas.
    Subdivide i1, imid, j1, jmid, Dy / 2
    Subdivide imid, i2, j1, jmid, Dy / 2
    Subdivide i1, imid, jmid, j2, Dy / 2
    Subdivide imid, i2, jmid, j2, Dy / 2
End Sub

' Create the Points array.
Public Sub SetBounds(ByVal x1 As Single, ByVal deltax As Single, ByVal xnum As Integer, ByVal z1 As Single, ByVal deltaz As Single, ByVal znum As Integer)
Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Z As Single

    Xmin = x1
    Zmin = z1
    Dx = deltax
    Dz = deltaz
    NumX = xnum
    NumZ = znum
    ReDim points(1 To NumX, 1 To NumZ)
    
    X = Xmin
    For i = 1 To NumX
        Z = Zmin
        For j = 1 To NumZ
            points(i, j).coord(1) = X
            points(i, j).coord(2) = 0
            points(i, j).coord(3) = Z
            points(i, j).coord(4) = 1#
            Z = Z + Dz
        Next j
        X = X + Dx
    Next i
End Sub
' Save the indicated data value.
Public Sub SetValue(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim i As Integer
Dim j As Integer

    i = (X - Xmin) / Dx + 1
    j = (Z - Zmin) / Dz + 1
    points(i, j).coord(2) = Y
End Sub

' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To NumX
        For j = 1 To NumZ
            m3ApplyFull points(i, j).coord, M, points(i, j).trans
        Next j
    Next i
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To NumX
        For j = 1 To NumZ
            m3Apply points(i, j).coord, M, points(i, j).trans
        Next j
    Next i
End Sub


' Draw the transformed points on a PictureBox.
Public Sub Draw(ByVal pic As Object)
Dim i As Integer
Dim j As Integer
Dim api_points(1 To 4) As POINTAPI

    On Error Resume Next

    ' See if we should fill the "rectangles."
    If RemoveHidden Then
        pic.FillStyle = vbFSSolid
        pic.FillColor = vbWhite
    Else
        pic.FillStyle = vbFSTransparent
    End If

    ' Draw the "rectangles."
    For i = 1 To NumX - 1
        For j = 1 To NumZ - 1
            ' Load the POINTAPI array.
            With api_points(1)
                .X = points(i, j).trans(1)
                .Y = points(i, j).trans(2)
            End With
            With api_points(2)
                .X = points(i + 1, j).trans(1)
                .Y = points(i + 1, j).trans(2)
            End With
            With api_points(3)
                .X = points(i + 1, j + 1).trans(1)
                .Y = points(i + 1, j + 1).trans(2)
            End With
            With api_points(4)
                .X = points(i, j + 1).trans(1)
                .Y = points(i, j + 1).trans(2)
            End With

            PickColor pic, _
                points(i, j).coord(2), _
                points(i + 1, j).coord(2), _
                points(i + 1, j + 1).coord(2), _
                points(i, j + 1).coord(2)

            Polygon pic.hdc, api_points(1), 4
        Next j
    Next i
End Sub
