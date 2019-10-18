Attribute VB_Name = "vbdStuff"
Option Explicit

Private m_OldPen As Long
Private m_OldBrush As Long
Private m_NewBrush As Long
Private m_NewPen As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const HS_BDIAGONAL = 3
Private Const HS_CROSS = 4
Private Const HS_DIAGCROSS = 5
Private Const HS_FDIAGONAL = 2
Private Const HS_HORIZONTAL = 0
Private Const HS_VERTICAL = 1

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

' Bound the objects in the collection.
Public Sub BoundObjects(ByVal the_objects As Collection, ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single
Dim obj As vbdObject

    Set obj = the_objects(1)
    obj.Bound xmin, ymin, xmax, ymax

    For Each obj In the_objects
        obj.Bound x1, y1, x2, y2
        If xmin > x1 Then xmin = x1
        If xmax < x2 Then xmax = x2
        If ymin > y1 Then ymin = y1
        If ymax < y2 Then ymax = y2
    Next obj
End Sub

' Initialize default drawing properties.
Public Sub InitializeDrawingProperties(ByVal obj As vbdObject)
    obj.DrawWidth = 1
    obj.DrawStyle = vbSolid
    obj.ForeColor = vbBlack
    obj.FillColor = vbBlack
    obj.FillStyle = vbFSTransparent
End Sub
' Return the drawing property serialization
' for this object.
Public Function DrawingPropertySerialization(ByVal obj As vbdObject) As String
Dim txt As String

    txt = txt & " DrawWidth(" & Format$(obj.DrawWidth) & ")"
    txt = txt & " DrawStyle(" & Format$(obj.DrawStyle) & ")"
    txt = txt & " ForeColor(" & Format$(obj.ForeColor) & ")"
    txt = txt & " FillColor(" & Format$(obj.FillColor) & ")"
    txt = txt & " FillStyle(" & Format$(obj.FillStyle) & ")"

    DrawingPropertySerialization = txt & vbCrLf & "    "
End Function

' Read the token name and value and to see
' if it is drawing property information.
Public Sub ReadDrawingPropertySerialization(ByVal obj As vbdObject, ByVal token_name As String, ByVal token_value As String)
    Select Case token_name
        Case "DrawWidth"
            obj.DrawWidth = CInt(token_value)
        Case "DrawStyle"
            obj.DrawStyle = CInt(token_value)
        Case "ForeColor"
            obj.ForeColor = CLng(token_value)
        Case "FillColor"
            obj.FillColor = CLng(token_value)
        Case "FillStyle"
            obj.FillStyle = CInt(token_value)
    End Select
End Sub


' Set the drawing properties for the canvas.
Public Sub SetCanvasDrawingParameters(ByVal obj As vbdObject, ByVal Canvas As Object)
    Canvas.DrawWidth = obj.DrawWidth
    Canvas.DrawStyle = obj.DrawStyle
    Canvas.ForeColor = obj.ForeColor
    Canvas.FillColor = obj.FillColor
    Canvas.FillStyle = obj.FillStyle
End Sub
' Set the drawing properties for the metafile.
Public Sub SetMetafileDrawingParameters(ByVal obj As vbdObject, ByVal mf_dc As Long)
Dim log_brush As LOGBRUSH
Dim new_brush As Long
Dim new_pen As Long

    With log_brush
        If obj.FillStyle = vbFSTransparent Then
            .lbStyle = BS_HOLLOW
        ElseIf obj.FillStyle = vbFSSolid Then
            .lbStyle = BS_SOLID
        Else
            .lbStyle = BS_HATCHED
            Select Case obj.FillStyle
                Case vbCross
                    .lbHatch = HS_CROSS
                Case vbDiagonalCross
                    .lbHatch = HS_DIAGCROSS
                Case vbDownwardDiagonal
                    .lbHatch = HS_BDIAGONAL
                Case vbHorizontalLine
                    .lbHatch = HS_HORIZONTAL
                Case vbUpwardDiagonal
                    .lbHatch = HS_FDIAGONAL
                Case vbVerticalLine
                    .lbHatch = HS_VERTICAL
            End Select
        End If
        .lbColor = obj.FillColor
    End With

    m_NewPen = CreatePen(obj.DrawStyle, obj.DrawWidth, obj.ForeColor)
    m_NewBrush = CreateBrushIndirect(log_brush)
    m_OldPen = SelectObject(mf_dc, m_NewPen)
    m_OldBrush = SelectObject(mf_dc, m_NewBrush)
End Sub
' Restore the drawing properties for the metafile.
Public Sub RestoreMetafileDrawingParameters(ByVal mf_dc As Long)
    SelectObject mf_dc, m_OldBrush
    SelectObject mf_dc, m_OldPen
    DeleteObject m_NewBrush
    DeleteObject m_NewPen
End Sub
' Return the serialization for this
' transformation matrix.
Public Function TransformationSerialization(M() As Single) As String
Dim i As Integer
Dim j As Integer
Dim txt As String

    For i = 1 To 3
        For j = 1 To 3
            txt = txt & Format$(M(i, j)) & " "
        Next j
    Next i

    TransformationSerialization = _
        "Transformation(" & txt & ")"
End Function
' initialize the transformation matrix
' using this serialization.
Public Sub SetTransformationSerialization(ByVal txt As String, M() As Single)
Dim i As Integer
Dim j As Integer
Dim token As String

    For i = 1 To 3
        For j = 1 To 3
            token = GetDelimitedToken(txt, " ")
            M(i, j) = CSng(token)
        Next j
    Next i
End Sub
