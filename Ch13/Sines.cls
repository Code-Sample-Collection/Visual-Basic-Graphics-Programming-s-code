VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DistortSines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Distortion

' The center about which to twist.
Public amplitude As Single
Public period As Single

' Transform the point.
Private Sub Distortion_Transform(X As Single, Y As Single, Z As Single)
Dim A As Single
Dim sy As Single
Dim sz As Single

    A = amplitude / 2
    sy = Sin(Y * 6.28 / period)
    sz = Sin(Z * 6.28 / period)
    X = X + A * (sy + sz)
End Sub


