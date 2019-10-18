VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' A fisheye transformation.

Implements Transformation

' The center of transformation.
Public Cx As Single
Public Cy As Single
Public Radius As Single

' Transform the point (X, Y).
Private Sub Transformation_Transform(X As Single, Y As Single)
Dim r As Single
Dim new_r As Single
Dim dx As Single
Dim dy As Single

    dx = X - Cx
    dy = Y - Cy
    r = Sqr(dx * dx + dy * dy)
    If r < 1 Then Exit Sub

    new_r = Radius * (1 - 1 / (r / Radius * 2 + 1))

    X = Cx + dx / r * new_r
    Y = Cy + dy / r * new_r
End Sub
