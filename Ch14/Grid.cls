VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Grid3d"
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
Private Points() As Point3D ' Data values.
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
    ReDim Points(1 To NumX, 1 To NumZ)
    
    X = Xmin
    For i = 1 To NumX
        Z = Zmin
        For j = 1 To NumZ
            Points(i, j).coord(1) = X
            Points(i, j).coord(2) = 0
            Points(i, j).coord(3) = Z
            Points(i, j).coord(4) = 1#
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
    Points(i, j).coord(2) = Y
End Sub

' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To NumX
        For j = 1 To NumZ
            m3ApplyFull Points(i, j).coord, M, Points(i, j).trans
        Next j
    Next i
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To NumX
        For j = 1 To NumZ
            m3Apply Points(i, j).coord, M, Points(i, j).trans
        Next j
    Next i
End Sub


' Draw the transformed points on a PictureBox.
Public Sub Draw(ByVal pic As Object)
Dim i As Integer
Dim j As Integer

    On Error Resume Next

    ' Draw lines parallel to the X axis.
    For i = 1 To NumX
        pic.CurrentX = Points(i, 1).trans(1)
        pic.CurrentY = Points(i, 1).trans(2)
        For j = 2 To NumZ
            pic.Line -(Points(i, j).trans(1), _
                          Points(i, j).trans(2))
        Next j
    Next i
    
    ' Draw lines parallel to the Y axis.
    For j = 1 To NumZ
        pic.CurrentX = Points(1, j).trans(1)
        pic.CurrentY = Points(1, j).trans(2)
        For i = 2 To NumX
            pic.Line -(Points(i, j).trans(1), _
                          Points(i, j).trans(2))
        Next i
    Next j
End Sub
