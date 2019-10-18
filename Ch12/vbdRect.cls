VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw Rectangle object.

Implements vbdObject

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
' Start drawing a rubberband box.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_DrawingStarted = True

    ' Start using dotted vbInvert mode.
    m_Canvas.DrawMode = vbInvert
    m_Canvas.DrawStyle = vbDot

    ' Start the first rubberband box.
    m_StartX = X
    m_StartY = Y
    m_LastX = X
    m_LastY = Y
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
End Sub

' Continue drawing the rubberband box.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_DrawingStarted Then Exit Sub

    ' Erase the old box.
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B

    ' Update the point.
    m_LastX = X
    m_LastY = Y

    ' Draw the new box.
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
End Sub


' Finish drawing the rubberband box.
Private Sub m_Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_DrawingStarted Then Exit Sub

    ' Erase the old box.
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B

    ' Go back to vbCopyPen drawing mode.
    m_Canvas.DrawMode = vbCopyPen

    ' Stop receiving events from the canvas.
    Set m_Canvas = Nothing

    ' Create the vbdPolygon that represents us.
    Set m_Polygon = New vbdPolygon
    Set m_Object = m_Polygon
    With m_Polygon
        .NumPoints = 4
        .X(1) = m_StartX
        .X(2) = m_LastX
        .X(3) = m_LastX
        .X(4) = m_StartX
        .Y(1) = m_StartY
        .Y(2) = m_StartY
        .Y(3) = m_LastY
        .Y(4) = m_LastY
        .Closed = True
    End With

    ' Tell the form to save us.
    frmVbDraw.AddObject Me

    ' Select the arrow tool.
    frmVbDraw.tbrTools.Buttons("Arrow").Value = tbrPressed
End Sub
Private Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    Set m_Canvas = RHS
End Property

Private Property Get vbdObject_Canvas() As PictureBox
    Set vbdObject_Canvas = m_Canvas
End Property

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
Public Sub vbdObject_Bound(ByRef xmin As Single, ByRef xmax As Single, ByRef ymin As Single, ByRef ymax As Single)
    m_Object.Bound xmin, ymin, xmax, ymax
End Sub
' Draw the object on the canvas.
Public Sub vbdObject_Draw(ByVal pic As Object)
    m_Object.Draw pic
End Sub
' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let vbdObject_Serialization(ByVal RHS As String)
'Dim token_name As String
'Dim token_value As String
'Dim next_x As Integer
'Dim next_y As Integer
'
'    InitializeDrawingProperties Me
'
'    ' Read tokens until there are no more.
'    Do While Len(RHS) > 0
'        ' Read a token.
'        GetNamedToken RHS, token_name, token_value
'        Select Case token_name
'            Case "NumPoints"
'                ' This allocates the m_X and m_Y arrays.
'                m_NumPoints = CSng(token_value)
'                next_x = 1
'                next_y = 1
'            Case "X"
'                x(next_x) = CSng(token_value)
'                next_x = next_x + 1
'            Case "Y"
'                y(next_y) = CSng(token_value)
'                next_y = next_y + 1
'            Case Else
'                ReadDrawingPropertySerialization Me, token_name, token_value
'        End Select
'    Loop
End Property
' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
'Dim txt As String
'Dim i As Integer
'
'    txt = DrawingPropertySerialization(Me)
'    txt = txt & " NumPoints(" & Format$(NumPoints) & ")"
'    For i = 1 To NumPoints
'        With m_OriginalPoints(i)
'            txt = txt & vbCrLf & "    X(" & Format$(.x) & ")"
'            txt = txt & " Y(" & Format$(.y) & ")"
'        End With
'    Next i
'
'    vbdObject_Serialization = "TwoDPolygon(" & txt & ")"
End Property

