VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw scene object.

Implements vbdObject

' The objects in the scene.
Public SceneObjects As Collection

' Drawing properties.
Private m_DrawWidth As Integer
Private m_DrawStyle As DrawStyleConstants
Private m_ForeColor As OLE_COLOR
Private m_FillColor As OLE_COLOR
Private m_FillStyle As FillStyleConstants

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' Move these objects to the beginning of the
' SceneObjects collection so they are drawn
' first.
Public Sub MoveToBack(ByVal targets As Collection)
Dim target As vbdObject

    ' Remove the objects from SceneObjects.
    RemoveObjects targets

    ' Re-add the objects at the beginning.
    For Each target In targets
        If SceneObjects.Count < 1 Then
            SceneObjects.Add target
        Else
            SceneObjects.Add target, , 1
        End If
    Next target
End Sub
' Move these objects to the end of the
' SceneObjects collection so they are drawn
' last.
Public Sub MoveToFront(ByVal targets As Collection)
Dim target As vbdObject

    ' Remove the objects from SceneObjects.
    RemoveObjects targets

    ' Re-add the objects at the end.
    For Each target In targets
        SceneObjects.Add target
    Next target
End Sub
' Remove these objects from SceneObjects.
Public Sub RemoveObjects(ByVal targets As Collection)
Dim target As vbdObject
Dim obj As vbdObject
Dim i As Integer

    ' Remove the objects from SceneObjects.
    For Each target In targets
        ' Find this target.
        i = 1
        For Each obj In SceneObjects
            If obj Is target Then
                SceneObjects.Remove i
                Exit For
            End If
            i = i + 1
        Next obj
    Next target
End Sub

' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(M() As Single)
Dim obj As vbdObject

    For Each obj In SceneObjects
        obj.AddTransformation M
    Next obj
End Sub

Private Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    ' Do nothing. This object is not directly
    ' creatable by the user.
End Property

Private Property Get vbdObject_Canvas() As PictureBox
    ' Do nothing. This object is not directly
    ' creatable by the user.
    Set vbdObject_Canvas = Nothing
End Property

' Clear the object's transformation.
Private Sub vbdObject_ClearTransformation()
Dim obj As vbdObject

    For Each obj In SceneObjects
        obj.ClearTransformation
    Next obj
End Sub

' Draw the object in a metafile.
Private Sub vbdObject_DrawInMetafile(ByVal mf_dc As Long)
Dim obj As vbdObject

    For Each obj In SceneObjects
        obj.DrawInMetafile mf_dc
    Next obj
End Sub
Private Sub Class_Initialize()
    Set SceneObjects = New Collection
End Sub


' Return this object's bounds.
Private Sub vbdObject_Bound(xmin As Single, ymin As Single, xmax As Single, ymax As Single)
    BoundObjects SceneObjects, xmin, ymin, xmax, ymax
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

' Draw the object on the canvas.
Private Sub vbdObject_Draw(ByVal Canvas As Object)
Dim obj As vbdObject

    For Each obj In SceneObjects
        obj.Draw Canvas
    Next obj
End Sub

' Set the objects' Selected statuses.
Private Property Let vbdObject_Selected(ByVal RHS As Boolean)
Dim obj As vbdObject

    For Each obj In SceneObjects
        obj.Selected = RHS
    Next obj
End Property
' Return the objects' Selected status.
Private Property Get vbdObject_Selected() As Boolean
Dim obj As vbdObject

    If SceneObjects.Count = 0 Then
        vbdObject_Selected = False
    Else
        Set obj = SceneObjects(1)
        vbdObject_Selected = obj.Selected
    End If
End Property

' Find the object at this position.
Public Function FindObjectAt(ByVal X As Single, ByVal Y As Single) As vbdObject
Dim obj As vbdObject
Dim i As Integer

    Set FindObjectAt = Nothing

    ' Search for the object starting with
    ' the objects on top.
    For i = SceneObjects.Count To 1 Step -1
        Set obj = SceneObjects(i)
        If obj.IsAt(X, Y) Then
            Set FindObjectAt = obj
            Exit For
        End If
    Next i
End Function

' Return True if the object is at this location.
Private Function vbdObject_IsAt(ByVal X As Single, ByVal Y As Single) As Boolean

End Function


' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let vbdObject_Serialization(ByVal RHS As String)
Dim obj As vbdObject
Dim token_name As String
Dim token_value As String

    ' Remove non-printable characters from the
    ' serialization.
    RHS = RemoveNonPrintables(RHS)

    ' Start with no objects.
    Set SceneObjects = New Collection

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "vbdLine"
                Set obj = New vbdLine
            Case "vbdPolygon"
                Set obj = New vbdPolygon
            Case "vbdScribble"
                Set obj = New vbdScribble
            Case "vbdScene"
                Set obj = New vbdScene
            Case Else
                Set obj = Nothing
        End Select

        ' Initialize the object.
        If Not obj Is Nothing Then
            obj.Serialization = token_value
            SceneObjects.Add obj
            Set obj = Nothing
        End If
    Loop
End Property

' Return a serialization string for the object.
Private Property Get vbdObject_Serialization() As String
Dim txt As String
Dim obj As vbdObject

    ' Don't bother with this object's
    ' drawing properties.


    ' Get the sub-objects' serializations.
    For Each obj In SceneObjects
        txt = txt & vbCrLf & "  " & _
            obj.Serialization
    Next obj

    vbdObject_Serialization = _
        "vbdScene(" & txt & vbCrLf & _
        "  )"
End Property
