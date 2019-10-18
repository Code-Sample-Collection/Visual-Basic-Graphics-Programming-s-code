VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DistortCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Distortion

Public period As Single
Public amplitude As Single
' Transform the point.
Private Sub Distortion_Transform(X As Single, Y As Single, Z As Single)
Dim D As Single

    D = Sqr(X * X + Z * Z)
    Y = Y + amplitude * Sin(D * 6.28 / period)
End Sub


