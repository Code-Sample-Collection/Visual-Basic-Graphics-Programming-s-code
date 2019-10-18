VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' A wavy transformation.

Implements Transformation

' Transformation parameters.
Public Amplitude As Single
Public Period As Single
' Transform the point (X, Y).
Private Sub Transformation_Transform(X As Single, Y As Single)
Dim new_x As Single

    new_x = X + Amplitude * Sin(Y * 6.28 / Period)
    Y = Y + Amplitude * Sin(X * 6.28 / Period)
    X = new_x
End Sub
