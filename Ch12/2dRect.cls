VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwoDRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Two-dimensional rectangle object.

Implements TwoDObject

Public X1 As Single
Public Y1 As Single
Public X2 As Single
Public Y2 As Single

' Drawing properties.
Private m_DrawWidth As Integer
Private m_DrawStyle As DrawStyleConstants
Private m_ForeColor As OLE_COLOR
Private m_FillColor As OLE_COLOR
Private m_FillStyle As FillStyleConstants

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' Draw the object in a metafile.
Private Sub TwoDObject_DrawInMetafile(ByVal mf_dc As Long)
    SetMetafileDrawingParameters Me, mf_dc

    Rectangle mf_dc, X1, Y1, X2, Y2

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

' Return this object's bounds.
Public Sub TwoDObject_Bound(ByRef xmin As Single, ByRef xmax As Single, ByRef ymin As Single, ByRef ymax As Single)
    If X1 < X2 Then
        xmin = X1
        xmax = X2
    Else
        xmin = X2
        xmax = X1
    End If
    If Y1 < Y2 Then
        ymin = Y1
        ymax = Y2
    Else
        ymin = Y2
        ymax = Y1
    End If
End Sub
' Draw the object on the canvas.
Public Sub TwoDObject_Draw(ByVal canvas As Object)
    SetCanvasDrawingParameters Me, canvas
    canvas.Line (X1, Y1)-(X2, Y2), , B
End Sub
' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let TwoDObject_Serialization(ByVal RHS As String)
Dim token_name As String
Dim token_value As String

    InitializeDrawingProperties Me

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "X1"
                X1 = CSng(token_value)
            Case "Y1"
                Y1 = CSng(token_value)
            Case "X2"
                X2 = CSng(token_value)
            Case "Y2"
                Y2 = CSng(token_value)
            Case Else
                ReadDrawingPropertySerialization Me, token_name, token_value
        End Select
    Loop
End Property

' Return a serialization string for the object.
Public Property Get TwoDObject_Serialization() As String
Dim txt As String

    txt = DrawingPropertySerialization(Me)
    txt = txt & " X1(" & Format$(X1) & ")"
    txt = txt & " Y1(" & Format$(Y1) & ")"
    txt = txt & " X2(" & Format$(X2) & ")"
    txt = txt & " Y2(" & Format$(Y2) & ")"
    TwoDObject_Serialization = "TwoDRectangle(" & txt & ")"
End Property


