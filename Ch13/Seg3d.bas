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

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
Private Type SIZE
    Cx As Long
    Cy As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

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
Public Sub DrawAllData(ByVal pic As Object, ByVal color As Long, ByVal clear As Boolean)
    DrawSomeData pic, 1, NumSegments, color, clear
End Sub

' Draw the indicated transformed segments.
Public Sub DrawSomeData(ByVal pic As Object, ByVal first_seg As Integer, ByVal last_seg As Integer, ByVal color As Long, ByVal clear As Boolean)
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
' Draw the indicated transformed segments into
' a metafile.
Public Sub DrawSomeDataToMetafile(ByVal file_name As String, ByVal wid As Single, ByVal hgt As Single, ByVal first_seg As Integer, ByVal last_seg As Integer)
Dim mDC As Long
Dim old_size As SIZE
Dim hmf As Long
Dim i As Integer
Dim x1 As Long
Dim y1 As Long
Dim x2 As Long
Dim y2 As Long

    ' Create the metafile.
    mDC = CreateMetaFile(ByVal file_name)
    If mDC = 0 Then
        MsgBox "Error creating the metafile.", vbExclamation
        Exit Sub
    End If

    ' Set the metafile's size to something reasonable.
    SetWindowExtEx mDC, wid, hgt, old_size

    ' Draw in the metafile.
    For i = first_seg To last_seg
        x1 = Segments(i).fr_tr(1) * 1000
        y1 = Segments(i).fr_tr(2) * 1000
        x2 = Segments(i).to_tr(1) * 1000
        y2 = Segments(i).to_tr(2) * 1000
        MoveToEx mDC, x1, y1, vbNullString
        LineTo mDC, x2, y2
    Next i

    ' Close the metafile.
    hmf = CloseMetaFile(mDC)
    If hmf = 0 Then
        MsgBox "Error closing the metafile.", vbExclamation
    End If

    ' Delete the metafile to free resources.
    If DeleteMetaFile(hmf) = 0 Then
        MsgBox "Error deleting the metafile.", vbExclamation
    End If
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
