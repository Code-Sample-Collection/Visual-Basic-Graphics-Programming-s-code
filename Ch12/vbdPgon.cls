VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw Polygon/Polyline object.

Implements vbdObject

Private Enum vbdPolygonErrors
    errInvalidIndex = 381   ' Invalid property array index.
End Enum

' Indicates a closed polygon rather than a polyline.
Public IsClosed As Boolean

' The surface on which the user is clicking
' to define the object. This is set only during
' creation of this object.
Public WithEvents m_Canvas As PictureBox
Attribute m_Canvas.VB_VarHelpID = -1
Private m_DrawingStarted As Boolean

' Drawing properties.
Private m_DrawWidth As Integer
Private m_DrawStyle As DrawStyleConstants
Private m_ForeColor As OLE_COLOR
Private m_FillColor As OLE_COLOR
Private m_FillStyle As FillStyleConstants
Private m_Selected As Boolean

' Data variables.
Private m_NumPoints As Long
Private m_OriginalPoints() As POINTAPI
Private m_TransformedPoints() As POINTAPI
Private m_M(1 To 3, 1 To 3) As Single

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single
' Return the number of points.
Public Property Get NumPoints() As Integer
    NumPoints = m_NumPoints
End Property

' Set the number of points.
Public Property Let NumPoints(ByVal new_value As Integer)
    m_NumPoints = new_value
    If m_NumPoints < 1 Then
        Erase m_OriginalPoints
    Else
        ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
    End If
End Property
' Return an X coordinate.
Property Get X(ByVal Index As Integer) As Single
    If (Index < 1) Or (Index > m_NumPoints) Then
        Err.Raise errInvalidIndex, "vbdPolygon.X"
    End If

    X = m_OriginalPoints(Index).X
End Property
' Set an X coordinate.
Property Let X(ByVal Index As Integer, ByVal new_value As Single)
    If (Index < 1) Or (Index > NumPoints) Then
        Err.Raise errInvalidIndex, "vbdPolygon.X"
    End If

    m_OriginalPoints(Index).X = new_value
End Property

' Return a Y coordinate.
Property Get Y(ByVal Index As Integer) As Single
    If (Index < 1) Or (Index > m_NumPoints) Then
        Err.Raise errInvalidIndex, "vbdPolygon.Y"
    End If

    Y = m_OriginalPoints(Index).Y
End Property
' Set a Y coordinate.
Property Let Y(ByVal Index As Integer, ByVal new_value As Single)
    If (Index < 1) Or (Index > NumPoints) Then
        Err.Raise errInvalidIndex, "vbdPolygon.Y"
    End If

    m_OriginalPoints(Index).Y = new_value
End Property

' Start with an identity transformation.
Private Sub Class_Initialize()
    ' Initialize the drawing parameters.
    InitializeDrawingProperties Me
    m2Identity m_M
End Sub

' The user has selected a point.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are drawing a rubberband line,
    ' erase it.
    If m_NumPoints > 0 Then m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)

    ' If this is the first point, start using
    ' dotted vbInvert mode.
    If m_NumPoints = 0 Then
        m_Canvas.DrawMode = vbInvert
        m_Canvas.DrawStyle = vbDot
        m_DrawingStarted = True
    End If

    ' See if this is the left or right button.
    If Button And vbLeftButton Then
        ' It's the left button.
        ' Add this point to the polygon.

        ' Add the new point.
        m_NumPoints = m_NumPoints + 1
        ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
        m_OriginalPoints(m_NumPoints).X = X
        m_OriginalPoints(m_NumPoints).Y = Y

        ' Draw the line permanently.
        m_Canvas.DrawMode = vbCopyPen
        m_Canvas.DrawStyle = vbSolid
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
        m_Canvas.DrawMode = vbInvert
        m_Canvas.DrawStyle = vbDot

        ' Start the next rubberband line.
        m_StartX = X
        m_StartY = Y
        m_LastX = X
        m_LastY = Y
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
    Else
        ' It's the right button.
        ' Stop building the polygon.

        ' Go back to vbCopyPen drawing mode.
        m_Canvas.DrawMode = vbCopyPen

        ' Stop receiving events from the canvas.
        Set m_Canvas = Nothing

        ' If we have at least 3 points, tell the
        ' form to save us.
        If m_NumPoints >= 3 Then
            ' We have at least 3 points. Tell the
            ' form to save us.
            frmVbDraw.AddObject Me
        Else
            ' We do not have 3 points. Tell the
            ' form to cancel us.
            frmVbDraw.CancelObject
        End If
    End If
End Sub

' Continue drawing the rubberband line.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_DrawingStarted Then Exit Sub

    ' Erase the old line.
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)

    ' Update the point.
    m_LastX = X
    m_LastY = Y

    ' Draw the new line.
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
End Sub


Private Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    Set m_Canvas = RHS
End Property

Private Property Get vbdObject_Canvas() As PictureBox
    Set vbdObject_Canvas = m_Canvas
End Property

' Clear the object's transformation.
Private Sub vbdObject_ClearTransformation()
    m2Identity m_M
End Sub
' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(M() As Single)
Dim T(1 To 3, 1 To 3) As Single

    m2MatMultiply T, m_M, M
    m2MatCopy m_M, T
End Sub


' Draw the object in a metafile.
Private Sub vbdObject_DrawInMetafile(ByVal mf_dc As Long)
    ' Make sure we have at least 2 points.
    If m_NumPoints < 2 Then Exit Sub

    SetMetafileDrawingParameters Me, mf_dc

    ' Draw the polygon.
    TransformPoints
    If IsClosed Then
        Polygon mf_dc, m_TransformedPoints(1), m_NumPoints
    Else
        Polyline mf_dc, m_TransformedPoints(1), m_NumPoints
    End If

    RestoreMetafileDrawingParameters mf_dc
End Sub
' Return the object's DrawWidth.
Public Property Get vbdObject_DrawWidth() As Integer
    vbdObject_DrawWidth = m_DrawWidth
End Property
' Set the object's DrawWidth.
Public Property Let vbdObject_DrawWidth(ByVal new_value As Integer)
    m_DrawWidth = new_value
End Property

' Return the object's DrawStyle.
Public Property Get vbdObject_DrawStyle() As DrawStyleConstants
    vbdObject_DrawStyle = m_DrawStyle
End Property
' Set the object's DrawStyle.
Public Property Let vbdObject_DrawStyle(ByVal new_value As DrawStyleConstants)
    m_DrawStyle = new_value
End Property

' Return the object's ForeColor.
Public Property Get vbdObject_ForeColor() As OLE_COLOR
    vbdObject_ForeColor = m_ForeColor
End Property
' Set the object's ForeColor.
Public Property Let vbdObject_ForeColor(ByVal new_value As OLE_COLOR)
    m_ForeColor = new_value
End Property

' Return the object's FillColor.
Public Property Get vbdObject_FillColor() As OLE_COLOR
    vbdObject_FillColor = m_FillColor
End Property
' Set the object's FillColor.
Public Property Let vbdObject_FillColor(ByVal new_value As OLE_COLOR)
    m_FillColor = new_value
End Property

' Return the object's FillStyle.
Public Property Get vbdObject_FillStyle() As FillStyleConstants
    vbdObject_FillStyle = m_FillStyle
End Property
' Set the object's FillStyle.
Public Property Let vbdObject_FillStyle(ByVal new_value As FillStyleConstants)
    m_FillStyle = new_value
End Property

' Return this object's bounds.
Public Sub vbdObject_Bound(ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim i As Integer

    If m_NumPoints < 1 Then
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
    Else
        With m_TransformedPoints(1)
            xmin = .X
            xmax = xmin
            ymin = .Y
            ymax = ymin
        End With

        For i = 2 To m_NumPoints
            With m_TransformedPoints(i)
                If xmin > .X Then xmin = .X
                If xmax < .X Then xmax = .X
                If ymin > .Y Then ymin = .Y
                If ymax < .Y Then ymax = .Y
            End With
        Next i
    End If
End Sub
' Draw the object on the canvas.
Public Sub vbdObject_Draw(ByVal pic As Object)
Const GAP = 4

Dim xmin As Single
Dim xmax As Single
Dim ymin As Single
Dim ymax As Single

    ' Make sure we have at least 2 points.
    If m_NumPoints < 2 Then Exit Sub

    SetCanvasDrawingParameters Me, pic

    ' Draw the polygon.
    TransformPoints
    If IsClosed Then
        Polygon pic.hdc, m_TransformedPoints(1), m_NumPoints
    Else
        Polyline pic.hdc, m_TransformedPoints(1), m_NumPoints
    End If

    ' Highlight if necessary.
    If m_Selected Then
        vbdObject_Bound xmin, ymin, xmax, ymax
        pic.DrawWidth = 1
        pic.DrawStyle = vbSolid
        pic.FillStyle = vbFSSolid
        pic.FillColor = vbWhite
        pic.Line (xmin, ymin)-Step(-GAP, -GAP), vbBlack, B
        pic.Line (xmin, ymax)-Step(-GAP, GAP), vbBlack, B
        pic.Line (xmax, ymin)-Step(GAP, -GAP), vbBlack, B
        pic.Line (xmax, ymax)-Step(GAP, GAP), vbBlack, B
    End If
End Sub
' Apply the transformation matrix.
Private Sub TransformPoints()
Dim i As Integer

    ReDim m_TransformedPoints(1 To m_NumPoints)
    For i = 1 To m_NumPoints
        With m_OriginalPoints(i)
            m_TransformedPoints(i).X = .X * m_M(1, 1) + .Y * m_M(2, 1) + m_M(3, 1)
            m_TransformedPoints(i).Y = .X * m_M(1, 2) + .Y * m_M(2, 2) + m_M(3, 2)
        End With
    Next i
End Sub

' Set the object's Selected status.
Private Property Let vbdObject_Selected(ByVal RHS As Boolean)
    m_Selected = RHS
End Property
' Return the object's Selected status.
Private Property Get vbdObject_Selected() As Boolean
    vbdObject_Selected = m_Selected
End Property

' Return True if the object is at this location.
Private Function vbdObject_IsAt(ByVal X As Single, ByVal Y As Single) As Boolean
Dim is_at As Boolean

    is_at = PolygonIsAt(IsClosed, X, Y, m_TransformedPoints)
    If (Not is_at) And IsClosed And _
        (m_FillStyle <> vbFSTransparent) _
    Then
        is_at = PointIsInPolygon(X, Y, m_TransformedPoints)
    End If

    vbdObject_IsAt = is_at
End Function

' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let vbdObject_Serialization(ByVal RHS As String)
Dim token_name As String
Dim token_value As String
Dim next_x As Integer
Dim next_y As Integer

    InitializeDrawingProperties Me
    m2Identity m_M

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "IsClosed"
                IsClosed = CBool(token_value)
            Case "NumPoints"
                ' This allocates the m_X and m_Y arrays.
                NumPoints = CLng(token_value)
                next_x = 1
                next_y = 1
            Case "X"
                X(next_x) = CSng(token_value)
                next_x = next_x + 1
            Case "Y"
                Y(next_y) = CSng(token_value)
                next_y = next_y + 1
            Case "Transformation"
                SetTransformationSerialization token_value, m_M
            Case Else
                ReadDrawingPropertySerialization Me, token_name, token_value
        End Select
    Loop
End Property
' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String
Dim i As Integer

    txt = DrawingPropertySerialization(Me)
    txt = txt & TransformationSerialization(m_M)
    txt = txt & " IsClosed(" & Format$(IsClosed) & ")"
    txt = txt & " NumPoints(" & Format$(NumPoints) & ")"
    For i = 1 To NumPoints
        With m_OriginalPoints(i)
            txt = txt & vbCrLf & "    X(" & Format$(.X) & ")"
            txt = txt & " Y(" & Format$(.Y) & ")"
        End With
    Next i

    vbdObject_Serialization = "vbdPolygon(" & txt & ")"
End Property

