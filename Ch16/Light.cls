VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LightSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X As Single
Public Y As Single
Public Z As Single
Public Ir As Single
Public Ig As Single
Public Ib As Single
Public Kdist As Single
Public Rmin As Single
' Initialize the light source's values.
Public Sub Initialize(ByVal new_x As Single, ByVal new_y As Single, ByVal new_z As Single, ByVal i_r As Integer, ByVal i_g As Integer, ByVal i_b As Integer)
    X = new_x
    Y = new_y
    Z = new_z
    Ir = i_r
    Ig = i_g
    Ib = i_b
End Sub
