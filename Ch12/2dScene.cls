VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwoDScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Two-dimensional scene object.

Implements TwoDObject

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

' Draw the object in a metafile.
Private Sub TwoDObject_DrawInMetafile(ByVal mf_dc As Long)
Dim obj As TwoDObject

    For Each obj In SceneObjects
        obj.DrawInMetafile mf_dc
    Next obj
End Sub
Private Sub Class_Initialize()
    Set SceneObjects = New Collection
End Sub


' Return this object's bounds.
Private Sub TwoDObject_Bound(xmin As Single, xmax As Single, ymin As Single, ymax As Single)
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim obj As TwoDObject

    If SceneObjects.Count < 1 Then
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
    Else
        Set obj = SceneObjects(1)
        obj.Bound xmin, ymin, xmax, ymax

        For Each obj In SceneObjects
            obj.Bound X1, Y1, X2, Y2
            If xmin > X1 Then xmin = X1
            If xmax < X2 Then xmax = X2
            If ymin > Y1 Then ymin = Y1
            If ymax < Y2 Then ymax = Y2
        Next obj
    End If
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

' Draw the object on the canvas.
Private Sub TwoDObject_Draw(ByVal canvas As Object)
Dim obj As TwoDObject

    For Each obj In SceneObjects
        obj.Draw canvas
    Next obj
End Sub

' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let TwoDObject_Serialization(ByVal RHS As String)
Dim obj As TwoDObject
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
            Case "TwoDCircle"
                Set obj = New TwoDCircle
            Case "TwoDEllipse"
                Set obj = New TwoDEllipse
            Case "TwoDLine"
                Set obj = New TwoDLine
            Case "TwoDPolygon"
                Set obj = New TwoDPolygon
            Case "TwoDRectangle"
                Set obj = New TwoDRectangle
            Case "TwoDSquare"
                Set obj = New TwoDSquare
            Case "TwoDScene"
                Set obj = New TwoDScene
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
Private Property Get TwoDObject_Serialization() As String
Dim txt As String
Dim obj As TwoDObject

    ' Don't bother with this object's
    ' drawing properties.

    
    ' Get the sub-objects' serializations.
    For Each obj In SceneObjects
        txt = txt & vbCrLf & "  " & _
            obj.Serialization
    Next obj

    TwoDObject_Serialization = _
        "TwoDScene(" & txt & vbCrLf & _
        "  )"
End Property
