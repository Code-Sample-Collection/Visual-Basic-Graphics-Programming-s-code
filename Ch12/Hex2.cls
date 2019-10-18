VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159265
Private Const NUM_SEGMENTS = 6

Public Cx As Single
Public Cy As Single
Public Radius As Single

Public Highlighted As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hrgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const ALTERNATE = 1

' Inidcates the object has been drawn lately.
Public Drawn As Boolean
' Bound this object.
Public Sub Bound(ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim dx As Single

    dx = Radius * Cos(PI / 6)
    xmin = Cx - dx
    xmax = Cx + dx
    ymin = Cy - Radius
    ymax = Cy + Radius
End Sub


' Draw on the PictureBox.
Public Sub Draw(ByVal pic As PictureBox)
Dim i As Integer
Dim theta As Single
Dim dtheta As Single
Dim clr As OLE_COLOR

    ' Do nothing if we're already drawn.
    If Drawn Then Exit Sub

    ' Pick an appropriate color.
    If Highlighted Then
        clr = vbRed
    Else
        clr = vbBlack
    End If

    ' Draw.
    theta = PI / 2
    pic.CurrentX = Cx + Radius * Cos(theta)
    pic.CurrentY = Cy + Radius * Sin(theta)

    dtheta = 2 * PI / NUM_SEGMENTS
    For i = 1 To NUM_SEGMENTS
        theta = theta + dtheta
        pic.Line -(Cx + Radius * Cos(theta), Cy + Radius * Sin(theta)), clr
    Next i

    ' Remember that we have been drawn.
    Drawn = True
End Sub
' Return True if the point is in the hex.
Public Function IsAt(ByVal X As Single, ByVal Y As Single) As Boolean
Dim points(1 To NUM_SEGMENTS) As POINTAPI
Dim hrgn As Long
Dim i As Integer
Dim theta As Single
Dim dtheta As Single

    theta = PI / 2
    dtheta = 2 * PI / NUM_SEGMENTS
    For i = 1 To NUM_SEGMENTS
        points(i).X = 100 * (Cx + Radius * Cos(theta))
        points(i).Y = 100 * (Cy + Radius * Sin(theta))
        theta = theta + dtheta
    Next i

    hrgn = CreatePolygonRgn(points(1), 5, ALTERNATE)
    IsAt = PtInRegion(hrgn, 100 * X, 100 * Y)
    DeleteObject hrgn
End Function


