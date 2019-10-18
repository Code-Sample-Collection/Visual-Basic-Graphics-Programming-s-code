Attribute VB_Name = "Wipes"
Option Explicit

Private ActiveImage As Integer
Private Wiping As Boolean

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Tile the pic_old with pic_new in a spiral
' from the outside in.
Public Sub TileSpiralIn(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal divisions_per_side As Integer)
Dim chunk_row() As Integer
Dim chunk_col() As Integer

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    ' Place the chunk rows and columns in the
    ' arrays in their correct order.
    PrepareTilesForSpiralIn divisions_per_side, chunk_row, chunk_col

    ' Display the tiles.
    DisplayTiles pic_new, pic_old, ms_per_frame, divisions_per_side, chunk_row, chunk_col

    Wiping = False
End Sub
' Tile the pic_old with pic_new randomly.
Public Sub TileRandom(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal divisions_per_side As Integer)
Dim chunk_row() As Integer
Dim chunk_col() As Integer
Dim num_chunks As Integer
Dim chunk As Integer
Dim i As Integer
Dim j As Integer
Dim tmp As Integer

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    ' Allocate the chunk_row and chunk_col arrays.
    num_chunks = divisions_per_side * divisions_per_side
    ReDim chunk_row(1 To num_chunks)
    ReDim chunk_col(1 To num_chunks)

    ' Put the row and column numbers in the
    ' chunk_row and chunk_col arrays.
    chunk = 1
    For i = 1 To divisions_per_side
        For j = 1 To divisions_per_side
            chunk_row(chunk) = i - 1
            chunk_col(chunk) = j - 1
            chunk = chunk + 1
        Next j
    Next i

    ' Randomize the chunks.
    For i = 1 To num_chunks - 1
        ' Pick a random entry between i and divisions_per_side.
        j = Int((num_chunks - i + 1) * Rnd + i)

        ' Swap that entry with the one in position i.
        If i <> j Then
            tmp = chunk_row(i)
            chunk_row(i) = chunk_row(j)
            chunk_row(j) = tmp
            tmp = chunk_col(i)
            chunk_col(i) = chunk_col(j)
            chunk_col(j) = tmp
        End If
    Next i

    ' Display the tiles.
    DisplayTiles pic_new, pic_old, ms_per_frame, divisions_per_side, chunk_row, chunk_col

    Wiping = False
End Sub

' Tile the pic_old with pic_new in a spiral
' from the inside out.
Public Sub TileSpiralOut(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal divisions_per_side As Integer)
Dim chunk_row_in() As Integer
Dim chunk_col_in() As Integer
Dim chunk_row() As Integer
Dim chunk_col() As Integer
Dim num_chunks As Integer
Dim chunk As Integer

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    ' Place the chunk rows and columns in the
    ' arrays in their correct order for a spiral in.
    PrepareTilesForSpiralIn divisions_per_side, chunk_row_in, chunk_col_in

    ' Allocate space for the outward spiral info.
    num_chunks = UBound(chunk_row_in)
    ReDim chunk_row(1 To num_chunks)
    ReDim chunk_col(1 To num_chunks)

    ' Reverse the tiles so they spiral out.
    For chunk = 1 To num_chunks
        chunk_row(chunk) = chunk_row_in(num_chunks - chunk + 1)
        chunk_col(chunk) = chunk_col_in(num_chunks - chunk + 1)
    Next chunk

    ' Display the tiles.
    DisplayTiles pic_new, pic_old, ms_per_frame, divisions_per_side, chunk_row, chunk_col

    Wiping = False
End Sub

' Place the chunk rows and columns in the
' arrays in their correct order.
Private Sub PrepareTilesForSpiralIn(ByVal divisions_per_side As Integer, ByRef chunk_row() As Integer, ByRef chunk_col() As Integer)
Dim num_chunks As Integer
Dim chunk As Integer
Dim r As Integer
Dim c As Integer
Dim dr As Integer
Dim dc As Integer
Dim rmin As Integer
Dim rmax As Integer
Dim cmin As Integer
Dim cmax As Integer

    ' Allocate arrays to hold the chunk rows and
    ' columns in the correct order.
    num_chunks = divisions_per_side * divisions_per_side
    ReDim chunk_row(1 To num_chunks)
    ReDim chunk_col(1 To num_chunks)

    ' Place the chunk rows and columns in the
    ' arrays in their correct order.
    rmin = 0
    cmin = 0
    rmax = divisions_per_side - 1
    cmax = divisions_per_side - 1
    chunk = 1
    Do
        ' Top.
        For c = cmin To cmax
            chunk_row(chunk) = rmin
            chunk_col(chunk) = c
            chunk = chunk + 1
        Next c
        If chunk > num_chunks Then Exit Do
        rmin = rmin + 1

        ' Right.
        For r = rmin To rmax
            chunk_row(chunk) = r
            chunk_col(chunk) = cmax
            chunk = chunk + 1
        Next r
        If chunk > num_chunks Then Exit Do
        cmax = cmax - 1

        ' Bottom.
        For c = cmax To cmin Step -1
            chunk_row(chunk) = rmax
            chunk_col(chunk) = c
            chunk = chunk + 1
        Next c
        If chunk > num_chunks Then Exit Do
        rmax = rmax - 1

        ' Left.
        For r = rmax To rmin Step -1
            chunk_row(chunk) = r
            chunk_col(chunk) = cmin
            chunk = chunk + 1
        Next r
        If chunk > num_chunks Then Exit Do
        cmin = cmin + 1
    Loop
End Sub

' Display the tiles in the indicated order.
Private Sub DisplayTiles(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal divisions_per_side As Integer, chunk_row() As Integer, chunk_col() As Integer)
Dim num_chunks As Integer
Dim chunk As Integer
Dim next_time As Long
Dim wid As Single
Dim hgt As Single

    wid = pic_old.ScaleWidth / divisions_per_side
    hgt = pic_old.ScaleHeight / divisions_per_side
    num_chunks = divisions_per_side * divisions_per_side

    ' Start displaying the tiles.
    next_time = GetTickCount()
    For chunk = 1 To num_chunks
        ' Copy the tile area.
        BitBlt pic_old.hDC, _
            wid * chunk_col(chunk), _
            hgt * chunk_row(chunk), _
            wid, hgt, _
            pic_new.hDC, _
            wid * chunk_col(chunk), _
            hgt * chunk_row(chunk), _
            vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time
    Next chunk

    ' Finish up.
    pic_old.Picture = pic_new.Picture
End Sub

' Wipe pic_new onto pic_old from the bottom up.
Public Sub WipeBottomToTop(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim Y As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight

    ' Start moving the image.
    Y = 0
    next_time = GetTickCount()
    Do While Y <= hgt
        ' Copy the area.
        BitBlt pic_old.hDC, 0, hgt - Y, wid, Y, _
            pic_new.hDC, 0, hgt - Y, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        Y = Y + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Push pic_new onto pic_old from the bottom up.
Public Sub PushBottomToTop(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim Y As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight

    ' Start moving the image.
    Y = 0
    next_time = GetTickCount()
    Do While Y <= hgt
        ' Move the existing area.
        BitBlt pic_old.hDC, 0, 0, wid, hgt - Y, _
            pic_old.hDC, 0, pixels_per_frame, vbSrcCopy

        ' Copy the area.
        BitBlt pic_old.hDC, 0, hgt - Y, wid, Y, _
            pic_new.hDC, 0, 0, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        Y = Y + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Push pic_new onto pic_old from the top down.
Public Sub PushTopToBottom(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim Y As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight

    ' Start moving the image.
    Y = 0
    next_time = GetTickCount()
    Do While Y <= hgt
        ' Move the existing area.
        BitBlt pic_old.hDC, 0, Y + pixels_per_frame, wid, hgt - Y, _
            pic_old.hDC, 0, Y, vbSrcCopy

        ' Copy the area.
        BitBlt pic_old.hDC, 0, 0, wid, Y, _
            pic_new.hDC, 0, hgt - Y, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        Y = Y + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub


' Wipe pic_new onto pic_old from left to right.
Public Sub WipeLeftToRight(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim X As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight

    ' Start moving the image.
    X = 0
    next_time = GetTickCount()
    Do While X <= wid
        ' Copy the area.
        BitBlt pic_old.hDC, 0, 0, X, hgt, _
            pic_new.hDC, 0, 0, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        X = X + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Push pic_new onto pic_old from left to right.
Public Sub PushLeftToRight(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim X As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight

    ' Start moving the image.
    X = 0
    next_time = GetTickCount()
    Do While X <= wid
        ' Move the existing area.
        BitBlt pic_old.hDC, X, 0, wid - X, hgt, _
            pic_old.hDC, X - pixels_per_frame, 0, vbSrcCopy

        ' Copy the area.
        BitBlt pic_old.hDC, 0, 0, X, hgt, _
            pic_new.hDC, wid - X, 0, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        X = X + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Push pic_new onto pic_old from right to left.
Public Sub PushRightToLeft(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim X As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight

    ' Start moving the image.
    X = 0
    next_time = GetTickCount()
    Do While X <= wid
        ' Move the existing area.
        BitBlt pic_old.hDC, 0, 0, wid - X, hgt, _
            pic_old.hDC, pixels_per_frame, 0, vbSrcCopy

        ' Copy the area.
        BitBlt pic_old.hDC, wid - X, 0, X, hgt, _
            pic_new.hDC, 0, 0, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        X = X + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub


' Wipe pic_new onto pic_old from right to left.
Public Sub WipeRightToLeft(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim X As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight

    ' Start moving the image.
    X = 0
    next_time = GetTickCount()
    Do While X <= wid
        ' Copy the area.
        BitBlt pic_old.hDC, wid - X, 0, X, hgt, _
            pic_new.hDC, wid - X, 0, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        X = X + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub


' Wipe pic_new onto pic_old from the upper right
' to the lower left.
Public Sub WipeURtoLL(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim imax As Single
Dim i As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight
    If wid > hgt Then
        imax = wid
    Else
        imax = hgt
    End If

    ' Start moving the image.
    i = 0
    next_time = GetTickCount()
    Do While i <= imax
        ' Copy the area.
        BitBlt pic_old.hDC, imax - i, 0, wid, i, _
            pic_new.hDC, imax - i, 0, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        i = i + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Wipe pic_new onto pic_old from the lower right
' to the upper left.
Public Sub WipeLRtoUL(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim imax As Single
Dim i As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight
    If wid > hgt Then
        imax = wid
    Else
        imax = hgt
    End If

    ' Start moving the image.
    i = 0
    next_time = GetTickCount()
    Do While i <= imax
        ' Copy the area.
        BitBlt pic_old.hDC, 0, imax - i, i, i, _
            pic_new.hDC, 0, imax - i, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        i = i + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Wipe pic_new onto pic_old from the center outward.
Public Sub WipeCenterOut(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim imax As Single
Dim i As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight
    If wid > hgt Then
        imax = wid / 2
    Else
        imax = hgt / 2
    End If

    ' Start moving the image.
    i = 0
    next_time = GetTickCount()
    Do While i <= imax
        ' Copy the area.
        BitBlt pic_old.hDC, wid / 2 - i, hgt / 2 - i, 2 * i, 2 * i, _
            pic_new.hDC, wid / 2 - i, hgt / 2 - i, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        i = i + pixels_per_frame / 2
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Wipe pic_new onto pic_old from the lower left
' to the upper right.
Public Sub WipeLLtoUR(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim imax As Single
Dim i As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight
    If wid > hgt Then
        imax = wid
    Else
        imax = hgt
    End If

    ' Start moving the image.
    i = 0
    next_time = GetTickCount()
    Do While i <= imax
        ' Copy the area.
        BitBlt pic_old.hDC, imax - i, imax - i, i, i, _
            pic_new.hDC, imax - i, imax - i, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        i = i + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Wipe pic_new onto pic_old from the upper left
' to the lower right.
Public Sub WipeULtoLR(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim imax As Single
Dim i As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight
    If wid > hgt Then
        imax = wid
    Else
        imax = hgt
    End If

    ' Start moving the image.
    i = 0
    next_time = GetTickCount()
    Do While i <= imax
        ' Copy the area.
        BitBlt pic_old.hDC, 0, 0, i, i, _
            pic_new.hDC, 0, 0, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        i = i + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
' Wipe pic_new onto pic_old from the top down.
Public Sub WipeTopToBottom(ByVal pic_new As PictureBox, ByVal pic_old As PictureBox, ByVal ms_per_frame As Long, ByVal pixels_per_frame As Integer)
Dim next_time As Long
Dim wid As Single
Dim hgt As Single
Dim Y As Single

    ' Prevent more than one wipe at a time.
    If Wiping Then Exit Sub
    Wiping = True

    wid = pic_old.ScaleWidth
    hgt = pic_old.ScaleHeight

    ' Start moving the image.
    Y = 0
    next_time = GetTickCount()
    Do While Y <= hgt
        ' Copy the area.
        BitBlt pic_old.hDC, 0, 0, wid, Y, _
            pic_new.hDC, 0, 0, vbSrcCopy
        pic_old.Refresh

        ' Wait for the next frame's time.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        Y = Y + pixels_per_frame
    Loop

    ' Finish up.
    pic_old.Picture = pic_new.Picture
    Wiping = False
End Sub
