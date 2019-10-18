VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdScribble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw Line/Rectangle object.

Implements vbdObject

' Indicates a box rather than a line.
Public IsBox As Boolean

' The surface on which the user is clicking
' to define the object. This is set only during
' creation of this object.
Public WithEvents m_Canvas As PictureBox
Attribute m_Canvas.VB_VarHelpID = -1
Private m_DrawingStarted As Boolean

' Constituent vbdPolygon object.
Private m_Polygon As vbdPolygon
Private m_Object As vbdObject

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single

' Start the scribble.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Start drawing the scribble.
    m_DrawingStarted = True

    ' Create the vbdPolygon that represents us.
    Set m_Polygon = New vbdPolygon
    Set m_Object = m_Polygon
    m_Polygon.IsClosed = False

    ' Save this point.
    m_Polygon.NumPoints = 1
    m_Polygon.X(m_Polygon.NumPoints) = X
    m_Polygon.Y(m_Polygon.NumPoints) = Y

    ' Draw the line.
    m_Canvas.CurrentX = X
    m_Canvas.CurrentY = Y

    ' Remember where we are.
    m_LastX = X
    m_LastY = Y
End Sub

' Continue the scribble.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_DrawingStarted Then Exit Sub

    ' Make sure we really moved.
    If m_LastX = X And m_LastY = Y Then Exit Sub

    ' Save this point.
    m_Polygon.NumPoints = m_Polygon.NumPoints + 1
    m_Polygon.X(m_Polygon.NumPoints) = X
    m_Polygon.Y(m_Polygon.NumPoints) = Y

    ' Draw the line.
    m_Canvas.Line -(X, Y)

    ' Remember where we are.
    m_LastX = X
    m_LastY = Y
End Sub
' Finish the scribble.
Private Sub m_Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_DrawingStarted Then Exit Sub
    m_DrawingStarted = False

    ' Stop receiving events from the canvas.
    Set m_Canvas = Nothing

    ' Tell the form to save us.
    frmVbDraw.AddObject Me
End Sub
' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(M() As Single)
    m_Object.AddTransformation M
End Sub

Private Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    Set m_Canvas = RHS
End Property

Private Property Get vbdObject_Canvas() As PictureBox
    Set vbdObject_Canvas = m_Canvas
End Property

' Clear the object's transformation.
Private Sub vbdObject_ClearTransformation()
    m_Object.ClearTransformation
End Sub

' Draw the object in a metafile.
Private Sub vbdObject_DrawInMetafile(ByVal mf_dc As Long)
    m_Object.DrawInMetafile mf_dc
End Sub
' Return the object's DrawWidth.
Public Property Get vbdObject_DrawWidth() As Integer
    vbdObject_DrawWidth = m_Object.DrawWidth
End Property
' Set the object's DrawWidth.
Public Property Let vbdObject_DrawWidth(ByVal new_value As Integer)
    m_Object.DrawWidth = new_value
End Property

' Return the object's DrawStyle.
Public Property Get vbdObject_DrawStyle() As DrawStyleConstants
    vbdObject_DrawStyle = m_Object.DrawStyle
End Property
' Set the object's DrawStyle.
Public Property Let vbdObject_DrawStyle(ByVal new_value As DrawStyleConstants)
    m_Object.DrawStyle = new_value
End Property

' Return the object's ForeColor.
Public Property Get vbdObject_ForeColor() As OLE_COLOR
    vbdObject_ForeColor = m_Object.ForeColor
End Property
' Set the object's ForeColor.
Public Property Let vbdObject_ForeColor(ByVal new_value As OLE_COLOR)
    m_Object.ForeColor = new_value
End Property

' Return the object's FillColor.
Public Property Get vbdObject_FillColor() As OLE_COLOR
    vbdObject_FillColor = m_Object.FillColor
End Property
' Set the object's FillColor.
Public Property Let vbdObject_FillColor(ByVal new_value As OLE_COLOR)
    m_Object.FillColor = new_value
End Property

' Return the object's FillStyle.
Public Property Get vbdObject_FillStyle() As FillStyleConstants
    vbdObject_FillStyle = m_Object.FillStyle
End Property
' Set the object's FillStyle.
Public Property Let vbdObject_FillStyle(ByVal new_value As FillStyleConstants)
    m_Object.FillStyle = new_value
End Property

' Return this object's bounds.
Public Sub vbdObject_Bound(ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
    m_Object.Bound xmin, ymin, xmax, ymax
End Sub
' Draw the object on the canvas.
Public Sub vbdObject_Draw(ByVal pic As Object)
    m_Object.Draw pic
End Sub

' Set the object's Selected status.
Private Property Let vbdObject_Selected(ByVal RHS As Boolean)
    m_Object.Selected = RHS
End Property

' Return the object's Selected status.
Private Property Get vbdObject_Selected() As Boolean
    vbdObject_Selected = m_Object.Selected
End Property

' Return True if the object is at this location.
Private Function vbdObject_IsAt(ByVal X As Single, ByVal Y As Single) As Boolean
    vbdObject_IsAt = m_Object.IsAt(X, Y)
End Function


' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let vbdObject_Serialization(ByVal RHS As String)
Dim token_name As String
Dim token_value As String
Dim next_x As Integer
Dim next_y As Integer

    ' Start with a new polygon.
    Set m_Polygon = New vbdPolygon
    Set m_Object = m_Polygon

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "IsBox"
                IsBox = CBool(token_value)
            Case "vbdPolygon"
                m_Object.Serialization = token_value
        End Select
    Loop
End Property
' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String

    txt = txt & "  IsBox(" & Format$(IsBox) & ") "
    txt = txt & m_Object.Serialization

    vbdObject_Serialization = "vbdScribble(" & txt & ")"
End Property

