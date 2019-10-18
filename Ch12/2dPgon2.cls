VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwoDPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

' The object's points.
Private m_NumPoints As Long
Private m_Points() As POINTAPI

' Invalid property-array index
Private Const INVALID_INDEX = 381
' Return the number of points.
Public Property Get NumPoints() As Integer
    NumPoints = m_NumPoints
End Property
' Set the number of points.
Public Property Let NumPoints(ByVal new_value As Integer)
    m_NumPoints = new_value
    If m_NumPoints < 1 Then
        Erase m_Points
    Else
        ReDim m_Points(1 To NumPoints)
    End If
End Property
' Return an X coordinate.
Property Get X(ByVal Index As Integer) As Single
    If (Index < 1) Or (Index > NumPoints) Then
        Err.Raise INVALID_INDEX, "TwoDPolygon.X"
    End If

    X = m_Points(Index).X
End Property
' Return a Y coordinate.
Property Get Y(ByVal Index As Integer) As Single
    If (Index < 1) Or (Index > NumPoints) Then
        Err.Raise INVALID_INDEX, "TwoDPolygon.X"
    End If

    Y = m_Points(Index).Y
End Property
' Set an X coordinate.
Property Let X(ByVal Index As Integer, ByVal new_value As Single)
    If (Index < 1) Or (Index > NumPoints) Then
        Err.Raise INVALID_INDEX, "TwoDPolygon.X"
    End If

    m_Points(Index).X = new_value
End Property
' Set a Y coordinate.
Property Let Y(ByVal Index As Integer, ByVal new_value As Single)
    If (Index < 1) Or (Index > NumPoints) Then
        Err.Raise INVALID_INDEX, "TwoDPolygon.X"
    End If

    m_Points(Index).Y = new_value
End Property

' Draw the object on the canvas.
Public Sub Draw(ByVal canvas As Object)
    ' Make sure we have at least 2 points.
    If NumPoints < 2 Then Exit Sub

    ' Draw the polygon.
    Polygon canvas.hdc, m_Points(1), NumPoints
End Sub
' Transform the object using a two-dimensional
' transformation matrix.
Public Sub Transform(M() As Single)
Dim i As Integer
Dim new_x As Single
Dim new_y As Single

    For i = 1 To m_NumPoints
        With m_Points(i)
            new_x = .X * M(1, 1) + .Y * M(2, 1) + M(3, 1)
            new_y = .X * M(1, 2) + .Y * M(2, 2) + M(3, 2)
            .X = new_x
            .Y = new_y
        End With
    Next i
End Sub
