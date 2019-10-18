VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Two-dimensional object parent class.

' The object's textual serialization.
Public Serialization As String

' Drawing properties.
Public DrawWidth As Integer
Public DrawStyle As DrawStyleConstants
Public ForeColor As OLE_COLOR
Public FillColor As OLE_COLOR
Public FillStyle As FillStyleConstants

' Indicates the object is Selected.
Public Selected As Boolean

' The form on which the object is created.
Public Canvas As PictureBox
' Clear the object's transformation.
Public Sub ClearTransformation()

End Sub
' Add this transformation to the current one.
Public Sub AddTransformation(M() As Single)

End Sub

' Draw the object on the canvas.
Public Sub Draw(ByVal Canvas As Object)

End Sub
' Return this object's bounds.
Public Sub Bound(ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)

End Sub
' Draw the object in a metafile.
Public Sub DrawInMetafile(ByVal mf_dc As Long)

End Sub
' Return True if the object is at this location.
Public Function IsAt(ByVal X As Single, ByVal Y As Single) As Boolean

End Function
