VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZOrderGrid3d"
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

            Polygon pic.hdc, api_points(1), 4
        Next j
    Next i
End Sub
