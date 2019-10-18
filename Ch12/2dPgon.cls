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
' Two-dimensional polygon object.

Implements TwoDObject

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

' Drawing properties.
Private m_DrawWidth As Integer
Private m_DrawStyle As DrawStyleConstants
Private m_ForeColor As OLE_COLOR
Private m_FillColor As OLE_COLOR
Private m_FillStyle As FillStyleConstants

' Draw the object in a metafile.
Private Sub TwoDObject_DrawInMetafile(ByVal mf_dc As Long)
    ' Make sure we have at least 2 points.
    If NumPoints < 2 Then Exit Sub

    SetMetafileDrawingParameters Me, mf_dc

    ' Draw the polygon.
    Polygon mf_dc, m_Points(1), NumPoints

    RestoreMetafileDrawingParameters mf_dc
End Sub
' Return the object's DrawWidth.
Public Property Get TwoDObject_DrawWidth() As Integer
    TwoDObject_DrawWidth = m_DrawWidth
End Property
' Set the object's DrawWidth.
Public Property Let TwoDObject_DrawWidth(ByVal new_value As Integer)
    m_DrawWidth = new_value
End Property

' Return the object's DrawStyle.
Public Property Get TwoDObject_DrawStyle() As DrawStyleConstants
    TwoDObject_DrawStyle = m_DrawStyle
End Property
' Set the object's DrawStyle.
Public Property Let TwoDObject_DrawStyle(ByVal new_value As DrawStyleConstants)
    m_DrawStyle = new_value
End Property

' Return the object's ForeColor.
Public Property Get TwoDObject_ForeColor() As OLE_COLOR
    TwoDObject_ForeColor = m_ForeColor
End Property
' Set the object's ForeColor.
Public Property Let TwoDObject_ForeColor(ByVal new_value As OLE_COLOR)
    m_ForeColor = new_value
End Property

' Return the object's FillColor.
Public Property Get TwoDObject_FillColor() As OLE_COLOR
    TwoDObject_FillColor = m_FillColor
End Property
' Set the object's FillColor.
Public Property Let TwoDObject_FillColor(ByVal new_value As OLE_COLOR)
    m_FillColor = new_value
End Property

' Return the object's FillStyle.
Public Property Get TwoDObject_FillStyle() As FillStyleConstants
    TwoDObject_FillStyle = m_FillStyle
End Property
' Set the object's FillStyle.
Public Property Let TwoDObject_FillStyle(ByVal new_value As FillStyleConstants)
    m_FillStyle = new_value
End Property

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

' Return this object's bounds.
Public Sub TwoDObject_Bound(ByRef xmin As Single, ByRef xmax As Single, ByRef ymin As Single, ByRef ymax As Single)
Dim i As Integer

    If NumPoints < 1 Then
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
    Else
        With m_Points(1)
            xmin = .X
            xmax = xmin
            ymin = .Y
            ymax = ymin
        End With

        For i = 2 To NumPoints
            With m_Points(i)
                If xmin > .X Then xmin = .X
                If xmax < .X Then xmax = .X
                If ymin > .Y Then ymin = .Y
                If ymax < .Y Then ymax = .Y
            End With
        Next i
    End If
End Sub
' Draw the object on the canvas.
Public Sub TwoDObject_Draw(ByVal canvas As Object)
    ' Make sure we have at least 2 points.
    If NumPoints < 2 Then Exit Sub

    SetCanvasDrawingParameters Me, canvas

    ' Draw the polygon.
    Polygon canvas.hdc, m_Points(1), NumPoints
End Sub
' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let TwoDObject_Serialization(ByVal RHS As String)
Dim token_name As String
Dim token_value As String
Dim next_x As Integer
Dim next_y As Integer

    InitializeDrawingProperties Me

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "NumPoints"
                ' This allocates the m_X and m_Y arrays.
                NumPoints = CSng(token_value)
                next_x = 1
                next_y = 1
            Case "X"
                X(next_x) = CSng(token_value)
                next_x = next_x + 1
            Case "Y"
                Y(next_y) = CSng(token_value)
                next_y = next_y + 1
            Case Else
                ReadDrawingPropertySerialization Me, token_name, token_value
        End Select
    Loop
End Property
' Return a serialization string for the object.
Public Property Get TwoDObject_Serialization() As String
Dim txt As String
Dim i As Integer

    txt = DrawingPropertySerialization(Me)
    txt = txt & " NumPoints(" & Format$(NumPoints) & ")"
    For i = 1 To NumPoints
        With m_Points(i)
            txt = txt & vbCrLf & "    X(" & Format$(.X) & ")"
            txt = txt & " Y(" & Format$(.Y) & ")"
        End With
    Next i

    TwoDObject_Serialization = "TwoDPolygon(" & txt & ")"
End Property


