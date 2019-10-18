VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DistortTwist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Distortion

' The center about which to twist.
Public cx As Single
Public cz As Single
Public period As Single
Public offset As Single

' Transform the point.
Private Sub Distortion_Transform(X As Single, Y As Single, Z As Single)
Dim theta As Single
    
    theta = (offset - Y) * 3.14 / period
    
    X = X * Cos(theta) - Z * Sin(theta)
    Z = X * Sin(theta) + Z * Cos(theta)
End Sub


