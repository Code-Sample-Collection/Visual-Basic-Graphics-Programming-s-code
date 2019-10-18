VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GZOrderGrid3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xmin As Single      ' Min X and Y values.
Private zmin As Single
Private dx As Single        ' Spacing between rows of data.
Private dz As Single
Private NumX As Integer     ' Number of X and Y entries.
Private NumZ As Integer
Private Points() As Point3D     ' Data values.
Private Normals() As Point3D    ' Vertex normals.

Public RemoveHidden As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const ALTERNATE = 1
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Diffuse reflection coefficients.
Public DiffuseKr As Single
Public DiffuseKg As Single
Public DiffuseKb As Single

' Ambient light coefficients.
Public AmbientKr As Single
Public AmbientKg As Single
Public AmbientKb As Single

' Specular reflection coefficients.
Public SpecularK As Single
Public SpecularN As Single
' Create the Points array.
Public Sub SetBounds(ByVal x1 As Single, ByVal deltax As Single, ByVal xnum As Integer, ByVal z1 As Single, ByVal deltaz As Single, ByVal znum As Integer)
Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Z As Single

    xmin = x1
    zmin = z1
    dx = deltax
    dz = deltaz
    NumX = xnum
    NumZ = znum
    ReDim Points(1 To NumX, 1 To NumZ)
    ReDim Normals(1 To NumX, 1 To NumZ)

    X = xmin
    For i = 1 To NumX
        Z = zmin
        For j = 1 To NumZ
            With Points(i, j)
                .coord(1) = X
                .coord(2) = 0
                .coord(3) = Z
                .coord(4) = 1#
            End With
            With Normals(i, j)
                .coord(1) = 0
                .coord(2) = -1
                .coord(3) = 0
                .coord(4) = 1#
            End With
            Z = Z + dz
        Next j
        X = X + dx
    Next i
End Sub
' Save the indicated data value.
Public Sub SetValue(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim i As Integer
Dim j As Integer

    i = (X - xmin) / dx + 1
    j = (Z - zmin) / dz + 1
    Points(i, j).coord(2) = Y
End Sub
' Save the indicated vertex normal value.
Public Sub SetNormal(ByVal X As Single, ByVal Z As Single, ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single)
Dim i As Integer
Dim j As Integer
Dim Length As Single

    Length = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
    i = (X - xmin) / dx + 1
    j = (Z - zmin) / dz + 1
    With Normals(i, j)
        .coord(1) = Nx / Length
        .coord(2) = Ny / Length
        .coord(3) = Nz / Length
    End With
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
            m3ApplyFull Normals(i, j).coord, M, Normals(i, j).trans
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
            m3ApplyFull Normals(i, j).coord, M, Normals(i, j).trans
        Next j
    Next i
End Sub


' Draw the transformed points on a PictureBox.
Public Sub Draw(ByVal pic As Object, ByVal light_sources As Collection, ByVal ambient_light As Integer, ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single)
Dim i As Integer
Dim j As Integer

    ' Draw the "rectangles."
    For i = 1 To NumX - 1
        For j = 1 To NumZ - 1
            GouraudQuadrilateral pic, _
                light_sources, ambient_light, _
                eye_x, eye_y, eye_z, _
                DiffuseKr, DiffuseKg, DiffuseKb, _
                AmbientKr, AmbientKg, AmbientKb, _
                SpecularK, SpecularN, _
                Points(i, j).coord(1), Points(i, j).coord(2), Points(i, j).coord(3), _
                Points(i, j + 1).coord(1), Points(i, j + 1).coord(2), Points(i, j + 1).coord(3), _
                Points(i + 1, j + 1).coord(1), Points(i + 1, j + 1).coord(2), Points(i + 1, j + 1).coord(3), _
                Points(i + 1, j).coord(1), Points(i + 1, j).coord(2), Points(i + 1, j).coord(3), _
                Normals(i, j).coord(1), Normals(i, j).coord(2), Normals(i, j).coord(3), _
                Normals(i, j + 1).coord(1), Normals(i, j + 1).coord(2), Normals(i, j + 1).coord(3), _
                Normals(i + 1, j + 1).coord(1), Normals(i + 1, j + 1).coord(2), Normals(i + 1, j + 1).coord(3), _
                Normals(i + 1, j).coord(1), Normals(i + 1, j).coord(2), Normals(i + 1, j).coord(3), _
                Points(i, j).trans(1), Points(i, j).trans(2), _
                Points(i, j + 1).trans(1), Points(i, j + 1).trans(2), _
                Points(i + 1, j + 1).trans(1), Points(i + 1, j + 1).trans(2), _
                Points(i + 1, j).trans(1), Points(i + 1, j).trans(2)
        Next j
    Next i
End Sub
