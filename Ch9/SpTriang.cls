VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TriangleSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Moving triangle sprite.
Option Explicit

Implements Sprite

' The three corners are stored as the distance
' PtR(i) and angle PtT(i) from the center of the
' triangle (Cx, Cy) to the corner. This makes it
' easier to rotate the triangle. When rotated by
' angle Theta, the coordinates of corner i are:
'
'   x = Cx + PtR(i) * Cos(PtT(i) + Theta)
'   y = Cy + PtR(i) * Sin(PtT(i) + Theta)
'
Private PtR(1 To 3) As Integer
Private PtT(1 To 3) As Single
Private Cx As Integer       ' Position of center.
Private Cy As Integer
Private Vx As Integer       ' Velocity.
Private Vy As Integer
Private Theta As Single     ' Orientation.
Private Vtheta As Single    ' Angular velocity.
Private Color As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

' Draw the triangle on the indicated picture box.
Public Sub Sprite_DrawSprite(ByVal pic As PictureBox)
Dim i As Integer
Dim pts(1 To 3) As POINTAPI

    ' Compute the current corner locations.
    For i = 1 To 3
        pts(i).x = Cx + PtR(i) * Cos(PtT(i) + Theta)
        pts(i).y = Cy + PtR(i) * Sin(PtT(i) + Theta)
    Next i

    ' Draw the triangle.
    pic.FillColor = Color
    pic.ForeColor = Color
    Polygon pic.hdc, pts(1), 3
End Sub

' Initialize the rectangle.
Public Sub InitializeTriangle(ByVal new_cx As Integer, ByVal new_cy As Integer, ByVal new_vx As Integer, ByVal new_vy As Integer, ByVal r1 As Integer, ByVal t1 As Integer, ByVal r2 As Integer, ByVal t2 As Integer, ByVal r3 As Integer, ByVal t3 As Integer, ByVal new_theta As Single, ByVal new_vtheta As Single, ByVal new_color As Long)
    Cx = new_cx
    Cy = new_cy
    Vx = new_vx
    Vy = new_vy

    PtR(1) = r1
    PtT(1) = t1
    PtR(2) = r2
    PtT(2) = t2
    PtR(3) = r3
    PtT(3) = t3

    Theta = new_theta
    Vtheta = new_vtheta
    Color = new_color
End Sub
' Add the velocity components to the sprite's
' position components.
Public Sub Sprite_MoveSprite(ByVal xmin As Integer, ByVal xmax As Integer, ByVal ymin As Integer, ByVal ymax As Integer)
    Cx = Cx + Vx
    Cy = Cy + Vy

    ' Keep the object within the drawing area.
    If (Cx < xmin) Then
        Cx = 2 * xmin - Cx
        Vx = -Vx
    ElseIf (Cx > xmax) Then
        Cx = 2 * xmax - Cx
        Vx = -Vx
    End If
    If (Cy < ymin) Then
        Cy = 2 * ymin - Cy
        Vy = -Vy
    ElseIf (Cy > ymax) Then
        Cy = 2 * ymax - Cy
        Vy = -Vy
    End If

    Theta = Theta + Vtheta
End Sub


