Attribute VB_Name = "DIBHelper"
Option Explicit

' ------------------------
' Bitmap Array Information
' ------------------------
Public Type RGBTriplet
    rgbRed As Byte
    rgbGreen As Byte
    rgbBlue As Byte
End Type

' ------------------
' Bitmap Information
' ------------------
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

' BITMAPINFO structure with room for up to
' 256 colors.
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

' Error codes.
Public Enum dibhErrors
    dibhInvalidBitsPerPixel = vbObjectError + 1001
    dibhCreateDCFailed
    dibhCreateBitmapFailed
    dibhSelectPaletteFailed
    dibhBitBltFailed
    dibhDeselectBitmapFailed
    dibhGetDIBitsFailed
    dibhStretchDIBitsFailed
End Enum

' API functions.
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

' API constants.
Private Const SRCCOPY = &HCC0020
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private Const GDI_ERROR = &HFFFF

' Return a binary representation of the byte.
' This helper function is useful for understanding
' byte values.
Public Function BinaryByte(ByVal value As Byte) As String
Dim i As Integer
Dim txt As String

    For i = 1 To 8
        If value And 1 Then
            txt = "1" & txt
        Else
            txt = "0" & txt
        End If
        value = value \ 2
    Next i

    BinaryByte = txt
End Function

' Find the closest color in the Colors array.
Public Function FindColorIndex(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, ByVal num_colors As Integer, Colors() As RGBQUAD) As Byte
Dim i As Integer
Dim best_i As Integer
Dim best_dist2 As Long
Dim dr As Long
Dim dg As Long
Dim db As Long
Dim dist2 As Long

    best_i = 0
    best_dist2 = CLng(3) * 256 * 256
    For i = 0 To num_colors - 1
        With Colors(i)
            dr = r - .rgbRed
            dg = g - .rgbGreen
            db = b - .rgbBlue
        End With
        dist2 = dr * dr + dg * dg + db * db
        If best_dist2 > dist2 Then
            best_dist2 = dist2
            best_i = i
        End If
    Next i

    FindColorIndex = best_i
End Function

' Load the bits from this PictureBox with a color
' depth of 24 bits into a two-dimensional array of
' RGB values.
'
' Note that the pixels are flipped vertically in
' the DIB structure. This routine flips them
' so the upper left corner is at pixels(0, 0).
Public Sub GetDIBPixels24Bit(ByVal pic As PictureBox, ByRef bitmap_info As BITMAPINFO, ByRef pixels() As RGBTriplet)
Dim memory_dc As Long
Dim new_bmp As Long
Dim old_bmp As Long
Dim wid As Long
Dim hgt As Long
Dim old_hpal As Long
Dim bytes_per_row As Long
Dim bytes() As Byte
Dim X As Integer
Dim Y As Integer
Dim bits_per_pixel As Long

    bits_per_pixel = 24

    ' Get the image's dimensions.
    wid = pic.ScaleX(pic.Image.Width, vbHimetric, vbPixels)
    hgt = pic.ScaleY(pic.Image.Height, vbHimetric, vbPixels)

    ' Create a memory device context (DC).
    memory_dc = CreateCompatibleDC(pic.hdc)
    If memory_dc = 0 Then
        Err.Raise dibhCreateDCFailed, _
            "DIBHelper.GetDIBPixels24Bit", _
            "Error creating compatible device context"
    End If

    ' Create a compatible bitmap.
    new_bmp = CreateCompatibleBitmap(pic.hdc, wid, hgt)
    If new_bmp = 0 Then
        Err.Raise dibhCreateBitmapFailed, _
            "DIBHelper.GetDIBPixels24Bit", _
            "Error creating compatible bitmap"
    End If

    ' Select the bitmap into the DC, saving
    ' its old bitmap handle.
    old_bmp = SelectObject(memory_dc, new_bmp)

    ' Make sure the PictureBox has an Image.
    pic.AutoRedraw = True

    ' If the picture has a palette,
    ' get and realize a copy.
    If pic.Image.hPal <> 0 Then
        ' Select the palette.
        old_hpal = SelectPalette(memory_dc, pic.Image.hPal, False)
        If old_hpal = 0 Then
            Err.Raise dibhSelectPaletteFailed, _
                "DIBHelper.GetDIBPixels24Bit", _
                "Error selecting palette into compatible bitmap"
        End If
        
        ' Realize the palette.
        RealizePalette memory_dc
    End If

    ' Copy the image from the PictureBox
    ' into the DC in memory.
    If BitBlt(memory_dc, 0, 0, wid, hgt, _
        pic.hdc, 0, 0, SRCCOPY) = 0 _
    Then
        Err.Raise dibhBitBltFailed, _
            "DIBHelper.GetDIBPixels24Bit", _
            "Error copying image from PictureBox into compatible bitmap"
    End If

    ' Deselect the compatible bitmap. GetDIBits
    ' requires that the bitmap from which it loads
    ' data is not selected by any DC.
    new_bmp = SelectObject(memory_dc, old_bmp)
    If new_bmp = 0 Then
        Err.Raise dibhDeselectBitmapFailed, _
            "DIBHelper.GetDIBPixels24Bit", _
            "Error deselecting the compatible bitmap"
    End If

    ' Initialize important bitmap info header fields.
    With bitmap_info.bmiHeader
        .biSize = Len(bitmap_info.bmiHeader)
        .biWidth = wid
        .biHeight = hgt
        .biPlanes = 1
        .biBitCount = bits_per_pixel
        .biCompression = BI_RGB
    End With

    ' Calculate the number of bytes per row needed
    ' to store the bitmap data. This is rounded up
    ' to the next multiple of 32 pixels (4 bytes).
    bytes_per_row = _
        ((wid * bits_per_pixel + 31) \ 32) * 4

    ' Allocate the array for the pixel data.
    ReDim bytes(0 To bytes_per_row - 1, 0 To hgt - 1)

    ' Get the bitmap data.
    If GetDIBits(memory_dc, new_bmp, 0, hgt, _
        bytes(0, 0), bitmap_info, DIB_RGB_COLORS) = 0 _
    Then
        Err.Raise dibhGetDIBitsFailed, _
            "DIBHelper.GetDIBPixels24Bit", _
            "Error using GetDIBits"
    End If

    ' Delete the objects we created.
    DeleteObject new_bmp
    DeleteObject memory_dc

    ' Create the pixels array.
    ReDim pixels(0 To wid - 1, 0 To hgt - 1)

    ' Copy the color values into the pixels array.
    For Y = 0 To hgt - 1
        For X = 0 To wid - 1
            With pixels(X, hgt - 1 - Y)
                .rgbBlue = bytes(X * 3, Y)
                .rgbGreen = bytes(X * 3 + 1, Y)
                .rgbRed = bytes(X * 3 + 2, Y)
            End With
        Next X
    Next Y
End Sub
' Load the bits from this PictureBox with a color
' depth of 1, 4, or 8 bits into a two-dimensional
' array of color index values.
'
' Note that the pixels are flipped vertically in
' the DIB structure. This routine flips them
' so the upper left corner is at pixels(0, 0).
Public Sub GetDIBPixelsWithPalette(ByVal pic As PictureBox, ByRef bitmap_info As BITMAPINFO, color_index() As Byte, ByVal bits_per_pixel As Long)
Dim memory_dc As Long
Dim new_bmp As Long
Dim old_bmp As Long
Dim wid As Long
Dim hgt As Long
Dim old_hpal As Long
Dim bytes_per_row As Long
Dim bytes() As Byte
Dim X As Integer
Dim Y As Integer
Dim shift_value As Integer
Dim i As Integer

    ' Verify that bits_per_pixel is 1, 4, or 8.
    If (bits_per_pixel <> 1) And _
       (bits_per_pixel <> 4) And _
       (bits_per_pixel <> 8) _
    Then
        Err.Raise dibhInvalidBitsPerPixel, _
            "DIBHelper.GetDIBPixelsWithPalette", _
            "The number of bits per pixel must be 1, 4, or 8"
    End If

    ' Get the image's dimensions.
    wid = pic.ScaleX(pic.Image.Width, vbHimetric, vbPixels)
    hgt = pic.ScaleY(pic.Image.Height, vbHimetric, vbPixels)

    ' Create a memory device context (DC).
    memory_dc = CreateCompatibleDC(pic.hdc)
    If memory_dc = 0 Then
        Err.Raise dibhCreateDCFailed, _
            "DIBHelper.GetDIBPixelsWithPalette", _
            "Error creating compatible device context"
    End If

    ' Create a compatible bitmap.
    new_bmp = CreateCompatibleBitmap(pic.hdc, wid, hgt)
    If new_bmp = 0 Then
        Err.Raise dibhCreateBitmapFailed, _
            "DIBHelper.GetDIBPixelsWithPalette", _
            "Error creating compatible bitmap"
    End If

    ' Select the bitmap into the DC, saving
    ' its old bitmap handle.
    old_bmp = SelectObject(memory_dc, new_bmp)

    ' Make sure the PictureBox has an Image.
    pic.AutoRedraw = True

    ' If the picture has a palette,
    ' get and realize a copy.
    If pic.Image.hPal <> 0 Then
        ' Select the palette.
        old_hpal = SelectPalette(memory_dc, pic.Image.hPal, False)
        If old_hpal = 0 Then
            Err.Raise dibhSelectPaletteFailed, _
                "DIBHelper.GetDIBPixelsWithPalette", _
                "Error selecting palette into compatible bitmap"
        End If
        
        ' Realize the palette.
        RealizePalette memory_dc
    End If

    ' Copy the image from the PictureBox
    ' into the DC in memory.
    If BitBlt(memory_dc, 0, 0, wid, hgt, _
        pic.hdc, 0, 0, SRCCOPY) = 0 _
    Then
        Err.Raise dibhBitBltFailed, _
            "DIBHelper.GetDIBPixelsWithPalette", _
            "Error copying image from PictureBox into compatible bitmap"
    End If

    ' Deselect the compatible bitmap. GetDIBits
    ' requires that the bitmap from which it loads
    ' data is not selected by any DC.
    new_bmp = SelectObject(memory_dc, old_bmp)
    If new_bmp = 0 Then
        Err.Raise dibhDeselectBitmapFailed, _
            "DIBHelper.GetDIBPixelsWithPalette", _
            "Error deselecting the compatible bitmap"
    End If

    ' Initialize important bitmap info header fields.
    With bitmap_info.bmiHeader
        .biSize = Len(bitmap_info.bmiHeader)
        .biWidth = wid
        .biHeight = hgt
        .biPlanes = 1
        .biBitCount = bits_per_pixel
        .biCompression = BI_RGB
    End With

    ' Calculate the number of bytes per row needed
    ' to store the bitmap data. This is rounded up
    ' to the next multiple of 32 pixels (4 bytes).
    bytes_per_row = _
        ((wid * bits_per_pixel + 31) \ 32) * 4

    ' Allocate the array for bitmap data.
    ReDim bytes(0 To bytes_per_row - 1, 0 To hgt - 1)

    ' Get the bitmap data.
    If GetDIBits(memory_dc, new_bmp, 0, hgt, _
        bytes(0, 0), bitmap_info, DIB_RGB_COLORS) = 0 _
    Then
        Err.Raise dibhGetDIBitsFailed, _
            "DIBHelper.GetDIBPixelsWithPalette", _
            "Error using GetDIBits"
    End If

    ' Delete the objects we created.
    DeleteObject new_bmp
    DeleteObject memory_dc

    ' Fill the color_index array.
    Select Case bits_per_pixel
        Case 1
            ' Allow room for all of the bytes array
            ' entries, even though some of those
            ' were added to make each row contain
            ' a multiple of 4 bytes.
            ReDim color_index(0 To (8 * bytes_per_row) - 1, 0 To hgt - 1)

            ' Copy the color index data.
            For Y = 0 To hgt - 1
                For X = 0 To bytes_per_row - 1
                    shift_value = 128
                    For i = 0 To 7
                        If bytes(X, Y) And shift_value Then
                            color_index(8 * X + i, hgt - 1 - Y) = 1
                        Else
                            color_index(8 * X + i, hgt - 1 - Y) = 0
                        End If
                        shift_value = shift_value \ 2
                    Next i
                Next X
            Next Y

        Case 4
            ' Allow room for all of the bytes array
            ' entries, even though some of those
            ' were added to make each row contain
            ' a multiple of 4 bytes.
            ReDim color_index(0 To (2 * bytes_per_row) - 1, 0 To hgt - 1)

            ' Copy the color index data.
            ' new_x gives the first index of the
            ' next entry in color_index.
            For Y = 0 To hgt - 1
                For X = 0 To bytes_per_row - 1
                    color_index(2 * X, hgt - 1 - Y) = _
                        bytes(X, Y) \ 16
                    color_index(2 * X + 1, hgt - 1 - Y) = _
                        bytes(X, Y) Mod 16
                Next X
            Next Y

        Case 8
            ' Allocate the color index array.
            ReDim color_index(0 To wid - 1, 0 To hgt - 1)

            ' Fill the color_index array.
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    color_index(X, hgt - 1 - Y) = bytes(X, Y)
                Next X
            Next Y

    End Select
End Sub

' Set the bits in this PictureBox using a
' two-dimensional array of RGB values.
'
' Note that the pixels are flipped vertically in
' the DIB structure. This routine flips them back
' so the upper left corner is at pixels(0, 0).
Public Sub SetDIBPixels24Bit(ByVal pic As PictureBox, bitmap_info As BITMAPINFO, pixels() As RGBTriplet)
Dim wid As Integer
Dim hgt As Integer
Dim bytes_per_row As Integer
Dim bits_per_pixel As Integer
Dim bytes() As Byte
Dim clr As Byte
Dim X As Integer
Dim Y As Integer

    ' See how big the image is.
    wid = bitmap_info.bmiHeader.biWidth
    hgt = bitmap_info.bmiHeader.biHeight
    bits_per_pixel = bitmap_info.bmiHeader.biBitCount

    ' Calculate the number of bytes per row needed
    ' to store the bitmap data. This is rounded up
    ' to the next multiple of 32 pixels (4 bytes).
    bytes_per_row = _
        ((wid * bits_per_pixel + 31) \ 32) * 4

    ' Allocate the bytes array.
    ReDim bytes(0 To bytes_per_row - 1, 0 To hgt - 1)

    ' Copy the pixel information into the bytes array.
    For Y = 0 To hgt - 1
        For X = 0 To wid - 1
            With pixels(X, hgt - 1 - Y)
                bytes(X * 3, Y) = .rgbBlue
                bytes(X * 3 + 1, Y) = .rgbGreen
                bytes(X * 3 + 2, Y) = .rgbRed
            End With
        Next X
    Next Y

    ' Copy the DIB information into the picture.
    If StretchDIBits( _
        pic.hdc, 0, 0, wid, hgt, _
        0, 0, wid, hgt, bytes(0, 0), bitmap_info, _
        DIB_RGB_COLORS, SRCCOPY) = GDI_ERROR _
    Then
        Err.Raise dibhStretchDIBitsFailed, _
            "DIBHelper.SetDIBPixels24Bit", _
            "Error using StretchDIBits"
    End If

    ' Make the changes visible.
    pic.Picture = pic.Image
End Sub
' Set the bits in this PictureBox using a 0-based
' two-dimensional array of color indexes.
'
' Note that the pixels are flipped vertically in
' the DIB structure. This routine flips them back
' so the upper left corner is at pixels(0, 0).
Public Sub SetDIBPixelsWithPalette(ByVal pic As PictureBox, bitmap_info As BITMAPINFO, color_index() As Byte)
Dim wid As Integer
Dim hgt As Integer
Dim bytes_per_row As Integer
Dim bits_per_pixel As Integer
Dim bytes() As Byte
Dim clr As Byte
Dim X As Integer
Dim Y As Integer
Dim shift_value As Integer
Dim i As Integer
Dim byte_value As Integer

    ' See how big the image is.
    wid = bitmap_info.bmiHeader.biWidth
    hgt = bitmap_info.bmiHeader.biHeight
    bits_per_pixel = bitmap_info.bmiHeader.biBitCount

    ' Calculate the number of bytes per row needed
    ' to store the bitmap data. This is rounded up
    ' to the next multiple of 32 pixels (4 bytes).
    bytes_per_row = _
        ((wid * bits_per_pixel + 31) \ 32) * 4

    ' Allocate the bytes array.
    ReDim bytes(0 To bytes_per_row - 1, 0 To hgt - 1)

    ' Copy the pixel information into the bytes array.
    Select Case bits_per_pixel
        Case 1
            ' Define the color data.
            For Y = 0 To hgt - 1
                For X = 0 To bytes_per_row - 1
                    shift_value = 128
                    byte_value = 0
                    For i = 0 To 7
                        If color_index(8 * X + i, hgt - 1 - Y) = 1 Then
                            byte_value = byte_value Or shift_value
                        End If
                        shift_value = shift_value \ 2
                    Next i
                    bytes(X, Y) = byte_value And &HFF&
                Next X
            Next Y

        Case 4
            ' Define the color data.
            For Y = 0 To hgt - 1
                For X = 0 To bytes_per_row - 1
                    bytes(X, Y) = _
                        16 * color_index(2 * X, hgt - 1 - Y) + _
                        color_index(2 * X + 1, hgt - 1 - Y)
                Next X
            Next Y

        Case 8
            ' Define the color data.
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    bytes(X, hgt - 1 - Y) = color_index(X, Y)
                Next X
            Next Y

    End Select

    ' Copy the DIB information into the picture.
    If StretchDIBits( _
        pic.hdc, 0, 0, wid, hgt, _
        0, 0, wid, hgt, bytes(0, 0), bitmap_info, _
        DIB_RGB_COLORS, SRCCOPY) = GDI_ERROR _
    Then
        Err.Raise dibhStretchDIBitsFailed, _
            "DIBHelper.SetDIBPixels24Bit", _
            "Error using StretchDIBits"
    End If

    ' Make the changes visible.
    pic.Picture = pic.Image
End Sub
