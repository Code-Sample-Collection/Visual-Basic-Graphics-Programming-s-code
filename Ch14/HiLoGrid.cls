VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HiLoGrid3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Xmin As Single      ' Min X and Y values.
Private Zmin As Single
Private Dx As Single        ' Spacing between rows of data.
Private Dz As Single
Private NumX As Integer     ' Number of X and Y entries.
Private NumZ As Integer
Private Points() As Point3D ' Data values.

Public RemoveHidden As Boolean
' Draw a line between the points. Set the hi and
' lo values for the line.
Private Sub DrawAndSetLine(ByVal pic As PictureBox, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, hi() As Integer, lo() As Integer)
Dim tmp As Single
Dim ix As Integer
Dim iy As Integer
Dim Y As Single
Dim dy As Single

    ' Deal only with integers.
    x1 = CInt(x1)
    y1 = CInt(y1)
    x2 = CInt(x2)
    y2 = CInt(y2)

    ' Make x1 < x2.
    If x2 < x1 Then
        tmp = x1
        x1 = x2
        x2 = tmp
        tmp = y1
        y1 = y2
        y2 = tmp
    End If

    ' Draw the line.
    pic.Line (x1, y1)-(x2, y2)

    ' Deal with vertical lines separately.
    If x1 = x2 Then
        If y1 < y2 Then
            lo(x1) = y1
            hi(x1) = y2
        Else
            lo(x1) = y2
            hi(x1) = y1
        End If
        Exit Sub
    End If

    ' Deal with non-vertical lines.
    dy = (y2 - y1) / CInt(x2 - x1)
    Y = y1
    For ix = x1 To x2
        iy = CInt(Y)

        lo(ix) = iy
        hi(ix) = iy

        Y = Y + dy
    Next ix
End Sub
' Draw a line between the points using and
' updating the hi and lo arrays.
Private Sub DrawLine(ByVal pic As PictureBox, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, hi() As Integer, lo() As Integer)
Dim tmp As Single
Dim ix As Integer
Dim iy As Integer
Dim Y As Single
Dim dy As Single
Dim firstx As Integer
Dim firsty As Integer
Dim skipping As Boolean
Dim above As Boolean

    ' Deal only with integers.
    x1 = CInt(x1)
    y1 = CInt(y1)
    x2 = CInt(x2)
    y2 = CInt(y2)

    ' Make x1 < x2.
    If x2 < x1 Then
        tmp = x1
        x1 = x2
        x2 = tmp
        tmp = y1
        y1 = y2
        y2 = tmp
    End If

    ' Deal with vertical lines separately.
    If x1 = x2 Then
        ' Make y1 < y2.
        If y2 < y1 Then
            tmp = y1
            y1 = y2
            y2 = tmp
        End If
        If y1 <= lo(x1) Then
            If y2 <= lo(x1) Then
                pic.Line (x1, y1)-(x2, y2)
            Else
                pic.Line (x1, y1)-(x2, lo(x2))
            End If
            lo(x1) = y1
        End If
        If y2 >= hi(x2) Then
            If y1 >= hi(x2) Then
                pic.Line (x1, y1)-(x2, y2)
            Else
                pic.Line (x1, hi(x1))-(x2, y2)
            End If
            hi(x2) = y2
        End If
        Exit Sub
    End If

    ' Deal with non-vertical lines.
    dy = (y2 - y1) / CInt(x2 - x1)
    Y = y1

    ' Find the first visible point.
    skipping = True
    For ix = x1 To x2
        iy = CInt(Y)
        ' See if this point is visible.
        If iy <= lo(ix) Then
            If skipping Then
                ' Start a new line below.
                skipping = False
                above = False
                firstx = ix
                firsty = lo(ix)
            End If
        ElseIf iy >= hi(ix) Then
            If skipping Then
                ' Start a new line above.
                skipping = False
                above = True
                firstx = ix
                firsty = hi(ix)
            End If
        Else
            ' This point is not visible.
            If Not skipping Then
                ' Draw the previous visible line.
                If above Then
                    ' The line is coming from
                    ' above. Connect it to hi(ix).
                    pic.Line (firstx, firsty)-(ix, hi(ix))
                Else
                    ' The line is coming from
                    ' below. Connect it to lo(ix).
                    pic.Line (firstx, firsty)-(ix, lo(ix))
                End If
                
                skipping = True
            End If
        End If

        If iy < lo(ix) Then lo(ix) = iy
        If iy > hi(ix) Then hi(ix) = iy

        Y = Y + dy
    Next ix

    ' Draw to the last point if necessary.
    If Not skipping Then _
        pic.Line (firstx, firsty)-(x2, y2)
End Sub
' Draw the grid including hidden surfaces.
Public Sub DrawWithHidden(ByVal pic As PictureBox, Optional R As Variant)
Dim i As Integer
Dim j As Integer

    On Error Resume Next
        
    ' Draw lines parallel to the X axis.
    For i = 1 To NumX
        pic.CurrentX = Points(i, 1).trans(1)
        pic.CurrentY = Points(i, 1).trans(2)
        For j = 2 To NumZ
            pic.Line -(Points(i, j).trans(1), _
                          Points(i, j).trans(2))
        Next j
    Next i

    ' Draw lines parallel to the Y axis.
    For j = 1 To NumZ
        pic.CurrentX = Points(1, j).trans(1)
        pic.CurrentY = Points(1, j).trans(2)
        For i = 2 To NumX
            pic.Line -(Points(i, j).trans(1), _
                          Points(i, j).trans(2))
        Next i
    Next j
End Sub
' Draw the grid without hidden surfaces using the
' Hi-Lo algorithm.
Public Sub DrawWithoutHidden(ByVal pic As Object, Optional R As Variant)
Dim Xmin As Integer
Dim Xmax As Integer
Dim lo() As Integer
Dim hi() As Integer
Dim ix As Integer
Dim i As Integer
Dim j As Integer

    ' Bound the X values.
    Xmin = Points(1, 1).trans(1)
    Xmax = Xmin
    For i = 1 To NumX
        For j = 1 To NumZ
            ix = CInt(Points(i, j).trans(1))
            If Xmin > ix Then Xmin = ix
            If Xmax < ix Then Xmax = ix
        Next j
    Next i

    ' Create the hi and lo arrays.
    ReDim lo(Xmin To Xmax)
    ReDim hi(Xmin To Xmax)

    ' Draw the X and Z front edges.
    For i = 2 To NumX
        ' Draw the edge between
        ' Points(i - 1, NumZ) and Points(i, NumZ)
        ' and set hi and lo for its values.
        DrawAndSetLine pic, _
            Points(i - 1, NumZ).trans(1), _
            Points(i - 1, NumZ).trans(2), _
            Points(i, NumZ).trans(1), _
            Points(i, NumZ).trans(2), _
            hi, lo
    Next i
    For i = 2 To NumZ
        ' Draw the edge between
        ' Points(NumX, i - 1) and Points(NumX, i)
        ' and set hi and lo for its values.
        DrawAndSetLine pic, _
            Points(NumX, i - 1).trans(1), _
            Points(NumX, i - 1).trans(2), _
            Points(NumX, i).trans(1), _
            Points(NumX, i).trans(2), _
            hi, lo
    Next i

    ' Draw the "rectangles."
    For i = NumX - 1 To 1 Step -1
        For j = NumZ - 1 To 1 Step -1
            ' Draw the edges between:
            '   Points(i, j) and Points(i + 1, j)
            '   Points(i, j) and Points(i, j + 1)
            
            ' If the right side of the "rectangle"
            ' leans over the top like this:
            '    +_
            '    | \_
            '    |   \_
            '    +     \_
            '     \      \
            '      +------+
            ' draw the top first so the right side
            ' doesn't make hi() too bit and stop
            ' the top from being drawn.
            '
            ' This only happens with perspective
            ' projection.
            If Points(i + 1, j).trans(1) >= Points(i, j).trans(1) Then
                DrawLine pic, _
                    Points(i, j).trans(1), _
                    Points(i, j).trans(2), _
                    Points(i, j + 1).trans(1), _
                    Points(i, j + 1).trans(2), _
                    hi, lo
                DrawLine pic, _
                    Points(i, j).trans(1), _
                    Points(i, j).trans(2), _
                    Points(i + 1, j).trans(1), _
                    Points(i + 1, j).trans(2), _
                    hi, lo
            Else
                DrawLine pic, _
                    Points(i, j).trans(1), _
                    Points(i, j).trans(2), _
                    Points(i + 1, j).trans(1), _
                    Points(i + 1, j).trans(2), _
                    hi, lo
                DrawLine pic, _
                    Points(i, j).trans(1), _
                    Points(i, j).trans(2), _
                    Points(i, j + 1).trans(1), _
                    Points(i, j + 1).trans(2), _
                    hi, lo
            End If
        Next j
    Next i
End Sub
' Create the Points array.
Public Sub SetBounds(ByVal x1 As Single, ByVal deltax As Single, ByVal xnum As Integer, ByVal z1 As Single, ByVal deltaz As Single, ByVal znum As Integer)
Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Z As Single

    Xmin = x1
    Zmin = z1
    Dx = deltax
    Dz = deltaz
    NumX = xnum
    NumZ = znum
    ReDim Points(1 To NumX, 1 To NumZ)
    
    X = Xmin
    For i = 1 To NumX
        Z = Zmin
        For j = 1 To NumZ
            Points(i, j).coord(1) = X
            Points(i, j).coord(2) = 0
            Points(i, j).coord(3) = Z
            Points(i, j).coord(4) = 1#
            Z = Z + Dz
        Next j
        X = X + Dx
    Next i
End Sub
' Save the indicated data value.
Public Sub SetValue(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim i As Integer
Dim j As Integer

    i = (X - Xmin) / Dx + 1
    j = (Z - Zmin) / Dz + 1
    Points(i, j).coord(2) = Y
End Sub

' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To NumX
        For j = 1 To NumZ
            m3ApplyFull Points(i, j).coord, M, Points(i, j).trans
        Next j
    Next i
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To NumX
        For j = 1 To NumZ
            m3Apply Points(i, j).coord, M, Points(i, j).trans
        Next j
    Next i
End Sub


' Draw the transformed points on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox)
    If RemoveHidden Then
        DrawWithoutHidden pic
    Else
        DrawWithHidden pic
    End If
End Sub
