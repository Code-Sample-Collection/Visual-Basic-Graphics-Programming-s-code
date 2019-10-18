Attribute VB_Name = "Seg3D"
Option Explicit

Public Type Segment
    ' The points to connect.
    fr_pt(1 To 4) As Single
    to_pt(1 To 4) As Single
    
    ' The transformed points to connect.
    fr_tr(1 To 4) As Single
    to_tr(1 To 4) As Single
End Type

Public Type Transformation
    M(1 To 4, 1 To 4) As Single
End Type

Public NumSegments As Integer
Public Segments() As Segment
' Check that all of the segments in this object
' have the same length. Return true if the
' segments all have the same length.
Public Function SameSideLengths(ByVal pt1 As Integer, ByVal pt2 As Integer) As Boolean
Dim A As Single
Dim B As Single
Dim C As Single
Dim S As Single
Dim i As Integer

    A = Segments(pt1).fr_pt(1) - Segments(pt1).to_pt(1)
    B = Segments(pt1).fr_pt(2) - Segments(pt1).to_pt(2)
    C = Segments(pt1).fr_pt(3) - Segments(pt1).to_pt(3)
    S = Sqr(A * A + B * B + C * C)
    
    SameSideLengths = False
    For i = pt1 + 1 To pt2
        A = Segments(i).fr_pt(1) - Segments(i).to_pt(1)
        B = Segments(i).fr_pt(2) - Segments(i).to_pt(2)
        C = Segments(i).fr_pt(3) - Segments(i).to_pt(3)
        If Abs(S - Sqr(A * A + B * B + C * C)) > 0.001 Then Exit Function
    Next i
    
    SameSideLengths = True
End Function

' Apply the translation matrix to all the
' segments using m3ApplyFull. The transformation
' may not have 0, 0, 0, 1 in its last column.
Public Sub TransformAllDataFull(M() As Single)
    TransformDataFull M, 1, NumSegments
End Sub

' Apply the translation matrix to the indicated
' segments using m3ApplyFull. The transformation
' may not have 0, 0, 0, 1 in its last column.
Public Sub TransformDataFull(M() As Single, ByVal seg1 As Integer, ByVal seg2 As Integer)
Dim i As Integer
    
    For i = seg1 To seg2
        m3ApplyFull Segments(i).fr_pt, M, Segments(i).fr_tr
        m3ApplyFull Segments(i).to_pt, M, Segments(i).to_tr
    Next i
End Sub


' Apply the translation matrix to all of the
' segments using m3Apply. This transformation
' must have 0, 0, 0, 1 in its last column.
Public Sub TransformAllData(M() As Single)
    TransformData M, 1, NumSegments
End Sub




' Apply the translation matrix to all the
' indicated segments using m3Apply. This
' transformation must have 0, 0, 0, 1 in its last
' column.
Public Sub TransformData(M() As Single, ByVal seg1 As Integer, ByVal seg2 As Integer)
Dim i As Integer
    
    For i = seg1 To seg2
        m3Apply Segments(i).fr_pt, M, Segments(i).fr_tr
        m3Apply Segments(i).to_pt, M, Segments(i).to_tr
    Next i
End Sub

' Set the point data to the transformed point data.
Public Sub SetPoints(ByVal seg1 As Integer, ByVal seg2 As Integer)
Dim i As Integer
Dim j As Integer

    For i = seg1 To seg2
        For j = 1 To 3
            Segments(i).fr_pt(j) = Segments(i).fr_tr(j)
            Segments(i).to_pt(j) = Segments(i).to_tr(j)
        Next j
    Next i
End Sub

' Draw the transformed segments.
Public Sub DrawAllData(ByVal pic As PictureBox, ByVal color As Long, ByVal clear As Boolean)
    DrawSomeData pic, 1, NumSegments, color, clear
End Sub

' Draw the indicated transformed segments.
Public Sub DrawSomeData(ByVal pic As PictureBox, ByVal first_seg As Integer, ByVal last_seg As Integer, ByVal color As Long, ByVal clear As Boolean)
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single

    If clear Then pic.Cls
    
    pic.ForeColor = color
    For i = first_seg To last_seg
        x1 = Segments(i).fr_tr(1)
        y1 = Segments(i).fr_tr(2)
        x2 = Segments(i).to_tr(1)
        y2 = Segments(i).to_tr(2)
        pic.Line (x1, y1)-(x2, y2)
    Next i
End Sub


' Create a segment.
Public Sub MakeSegment(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single)
    NumSegments = NumSegments + 1
    ReDim Preserve Segments(1 To NumSegments)
    Segments(NumSegments).fr_pt(1) = x1
    Segments(NumSegments).fr_pt(2) = y1
    Segments(NumSegments).fr_pt(3) = z1
    Segments(NumSegments).fr_pt(4) = 1
    Segments(NumSegments).to_pt(1) = x2
    Segments(NumSegments).to_pt(2) = y2
    Segments(NumSegments).to_pt(3) = z2
    Segments(NumSegments).to_pt(4) = 1
End Sub
