VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransTwist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' A twisting transformation.

Implements Transformation

' The center about which to twist.
Public Cx As Single
Public Cy As Single

' The amount by which to twist at different
' distances from the center.
Public TwistSpeed As Single
' Transform the point (X, Y).
Private Sub Transformation_Transform(X As Single, Y As Single)
Dim R As Single
Dim dx As Single
Dim dy As Single

    dx = X - Cx
    dy = Y - Cy
    R = Sqr(dx * dx + dy * dy) / TwistSpeed / 5

    X = Cx + dx * Sin(R) - dy * Cos(R)
    Y = Cy + dx * Cos(R) + dy * Sin(R)
End Sub
