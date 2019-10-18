VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BallSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Bouncing ball sprite.
Option Explicit

Implements Sprite

Private Radius As Integer
Private Cx As Integer       ' Position of center.
Private Cy As Integer
Private Vx As Integer       ' Velocity.
Private Vy As Integer
Private Color As Long
' Draw the circle on the indicated picture box.
Public Sub Sprite_DrawSprite(ByVal pic As PictureBox)
    pic.FillColor = Color
    pic.Circle (Cx, Cy), Radius, Color
End Sub

' Initialize the ball.
Public Sub InitializeBall(ByVal new_radius As Integer, ByVal new_cx As Integer, ByVal new_cy As Integer, ByVal new_vx As Integer, ByVal new_vy As Integer, ByVal new_color As Long)
    Radius = new_radius
    Cx = new_cx
    Cy = new_cy
    Vx = new_vx
    Vy = new_vy
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
End Sub
