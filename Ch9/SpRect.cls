VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectangleSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Moving rectangle sprite.
Option Explicit

Implements Sprite

Private Wid As Integer
Private Hgt As Integer
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
' Draw the rectangle on the indicated picture box.
Public Sub Sprite_DrawSprite(ByVal pic As PictureBox)
Const PI = 3.14159265
Const PI_OVER_2 = PI / 2

Dim wx As Single
Dim wy As Single
Dim hx As Single
Dim hy As Single
Dim pts(1 To 4) As POINTAPI
Dim status As Long
Dim newpen As Long
Dim oldpen As Long
Dim newbrush As Long
Dim oldbrush As Long

    ' Compute vectors parallel to the axes.
    wx = Wid * Cos(Theta)
    wy = Wid * Sin(Theta)
    hx = Hgt * Cos(Theta + PI_OVER_2)
    hy = Hgt * Sin(Theta + PI_OVER_2)
    
    pts(1).x = Cx + (wx + hx) / 2
    pts(1).y = Cy + (wy + hy) / 2
    pts(2).x = pts(1).x - hx
    pts(2).y = pts(1).y - hy
    pts(3).x = pts(2).x - wx
    pts(3).y = pts(2).y - wy
    pts(4).x = pts(3).x + hx
    pts(4).y = pts(3).y + hy

    ' Draw the rectangle.
    pic.FillColor = Color
    pic.ForeColor = Color
    Polygon pic.hdc, pts(1), 4
End Sub

' Initialize the rectangle.
Public Sub InitializeRectangle(ByVal new_wid As Integer, ByVal new_hgt As Integer, ByVal new_cx As Integer, ByVal new_cy As Integer, ByVal new_vx As Integer, ByVal new_vy As Integer, ByVal new_theta As Single, ByVal new_vtheta As Single, ByVal new_color As Long)
    Wid = new_wid
    Hgt = new_hgt
    Cx = new_cx
    Cy = new_cy
    Vx = new_vx
    Vy = new_vy
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


