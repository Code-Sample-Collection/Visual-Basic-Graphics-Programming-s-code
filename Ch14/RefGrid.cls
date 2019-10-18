VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RefinedGrid3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The collection of Polyline3d objects.
Public Polylines As Collection
' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim pline As Polyline3d

    For Each pline In Polylines
        pline.ApplyFull M
    Next pline
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim pline As Polyline3d

    For Each pline In Polylines
        pline.Apply M
    Next pline
End Sub
' Draw the transformed points on a PictureBox.
Public Sub Draw(ByVal pic As Object)
Dim pline As Polyline3d

    For Each pline In Polylines
        pline.Draw pic
    Next pline
End Sub

' Create the empty Polylines collection.
Private Sub Class_Initialize()
    Set Polylines = New Collection
End Sub

