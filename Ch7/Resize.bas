Attribute VB_Name = "Resize"
Option Explicit

' Shrink or enlarge the picture.
Public Sub ResizePicture(ByVal pic_from As PictureBox, ByVal pic_to As PictureBox, ByVal from_xmin As Single, ByVal from_ymin As Single, ByVal from_wid As Single, ByVal from_hgt As Single, ByVal to_xmin As Single, ByVal to_ymin As Single, ByVal to_wid As Single, ByVal to_hgt As Single)
Dim x_scale As Single
Dim y_scale As Single

    ' If either scale is less than 1, use ShrinkPicture
    If (to_wid / from_wid < 1#) Or _
       (to_hgt / from_hgt < 1#) _
    Then
        ' Shrink the picture.
        ShrinkPicture pic_from, pic_to, _
            from_xmin, from_ymin, _
            from_wid, from_hgt, _
            to_xmin, to_ymin, _
            to_wid, to_hgt
    Else
        ' Enlarge the picture.
        EnlargePicture pic_from, pic_to, _
            from_xmin, from_ymin, _
            from_wid, from_hgt, _
            to_xmin, to_ymin, _
            to_wid, to_hgt
    End If
End Sub
' Shrink the image.
Private Sub ShrinkPicture(ByVal pic_from As PictureBox, ByVal pic_to As PictureBox, ByVal from_xmin As Single, ByVal from_ymin As Single, ByVal from_wid As Single, ByVal from_hgt As Single, ByVal to_xmin As Single, ByVal to_ymin As Single, ByVal to_wid As Single, ByVal to_hgt As Single)
Dim x_scale As Single
Dim y_scale As Single
Dim white_pixel As RGBTriplet
Dim input_pixels() As RGBTriplet
Dim result_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim ix_max As Single
Dim iy_max As Single
Dim x_in As Single
Dim y_in As Single
Dim ix_out As Integer
Dim iy_out As Integer
Dim ix_in As Integer
Dim iy_in As Integer
Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single
Dim X As Integer
Dim Y As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim num_pixels As Integer

    ' Set the white pixel's value.
    With white_pixel
        .rgbRed = 255
        .rgbGreen = 255
        .rgbBlue = 255
    End With

    ' Get the pixels from pic_from.
    GetBitmapPixels pic_from, input_pixels, bits_per_pixel

    ' Get the pixels from pic_to.
    GetBitmapPixels pic_to, result_pixels, bits_per_pixel

    ' Get the original image's bounds.
    ix_max = pic_from.ScaleWidth - 2
    iy_max = pic_from.ScaleHeight - 2

    ' Calulate the mapping values.
    from_xmin = from_xmin
    from_ymin = from_ymin
    to_xmin = to_xmin
    to_ymin = to_ymin
    x_scale = to_wid / (from_wid - 1)
    y_scale = to_hgt / (from_hgt - 1)

    ' Calculate the output pixel values.
    For iy_out = 0 To pic_to.ScaleHeight - 1
        For ix_out = 0 To pic_to.ScaleWidth - 1
            ' Map the pixel value from
            ' (ix_out, iy_out) to (x_in, y_in).
            x1 = Int(from_xmin + (ix_out - to_xmin) / x_scale)
            x2 = Int(from_xmin + (ix_out + 1 - to_xmin) / x_scale) - 1
            y1 = Int(from_ymin + (iy_out - to_ymin) / y_scale)
            y2 = Int(from_ymin + (iy_out + 1 - to_ymin) / y_scale) - 1

            ' Average the pixels in this area.
            r = 0
            g = 0
            b = 0
            For X = x1 To x2
                For Y = y1 To y2
                    With input_pixels(X, Y)
                        r = r + .rgbRed
                        g = g + .rgbGreen
                        b = b + .rgbBlue
                    End With
                Next Y
            Next X

            ' Save the result.
            num_pixels = (x2 - x1 + 1) * (y2 - y1 + 1)
            With result_pixels(ix_out, iy_out)
                .rgbRed = r / num_pixels
                .rgbGreen = g / num_pixels
                .rgbBlue = b / num_pixels
            End With
        Next ix_out
    Next iy_out

    ' Set pic_to's pixels.
    SetBitmapPixels pic_to, bits_per_pixel, result_pixels
    pic_to.Picture = pic_to.Image
End Sub
' Enlarge the image.
Private Sub EnlargePicture(ByVal pic_from As PictureBox, ByVal pic_to As PictureBox, ByVal from_xmin As Single, ByVal from_ymin As Single, ByVal from_wid As Single, ByVal from_hgt As Single, ByVal to_xmin As Single, ByVal to_ymin As Single, ByVal to_wid As Single, ByVal to_hgt As Single)
Dim x_scale As Single
Dim y_scale As Single
Dim white_pixel As RGBTriplet
Dim input_pixels() As RGBTriplet
Dim result_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim ix_max As Single
Dim iy_max As Single
Dim x_in As Single
Dim y_in As Single
Dim ix_out As Integer
Dim iy_out As Integer
Dim ix_in As Integer
Dim iy_in As Integer
Dim dx As Single
Dim dy As Single
Dim dx1 As Single
Dim dx2 As Single
Dim dy1 As Single
Dim dy2 As Single
Dim v11 As Integer
Dim v12 As Integer
Dim v21 As Integer
Dim v22 As Integer

    ' Set the white pixel's value.
    With white_pixel
        .rgbRed = 255
        .rgbGreen = 255
        .rgbBlue = 255
    End With

    ' Get the pixels from pic_from.
    GetBitmapPixels pic_from, input_pixels, bits_per_pixel

    ' Get the pixels from pic_to.
    GetBitmapPixels pic_to, result_pixels, bits_per_pixel

    ' Get the original image's bounds.
    ix_max = pic_from.ScaleWidth - 2
    iy_max = pic_from.ScaleHeight - 2

    ' Calulate the mapping values.
    from_xmin = from_xmin
    from_ymin = from_ymin
    to_xmin = to_xmin
    to_ymin = to_ymin
    x_scale = to_wid / (from_wid - 1)
    y_scale = to_hgt / (from_hgt - 1)

    ' Calculate the output pixel values.
    For iy_out = 0 To pic_to.ScaleHeight - 1
        For ix_out = 0 To pic_to.ScaleWidth - 1
            ' Map the pixel value from
            ' (ix_out, iy_out) to (x_in, y_in).
            x_in = from_xmin + (ix_out - to_xmin) / x_scale
            y_in = from_ymin + (iy_out - to_ymin) / y_scale

            ' Interpolate to find the pixel's value.
            ' Find the nearest integral position.
            ix_in = Int(x_in)
            iy_in = Int(y_in)

            ' See if this is out of bounds.
            If (ix_in < 0) Or (ix_in > ix_max) Or _
               (iy_in < 0) Or (iy_in > iy_max) _
            Then
                ' The point is outside the image.
                ' Use white.
                result_pixels(ix_out, iy_out) = white_pixel
            Else
                ' The point lies within the image.
                ' Calculate its value.
                dx1 = x_in - ix_in
                dy1 = y_in - iy_in
                dx2 = 1# - dx1
                dy2 = 1# - dy1

                With result_pixels(ix_out, iy_out)
                    ' Calculate the red value.
                    v11 = input_pixels(ix_in, iy_in).rgbRed
                    v12 = input_pixels(ix_in, iy_in + 1).rgbRed
                    v21 = input_pixels(ix_in + 1, iy_in).rgbRed
                    v22 = input_pixels(ix_in + 1, iy_in + 1).rgbRed
                    .rgbRed = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
        
                    ' Calculate the green value.
                    v11 = input_pixels(ix_in, iy_in).rgbGreen
                    v12 = input_pixels(ix_in, iy_in + 1).rgbGreen
                    v21 = input_pixels(ix_in + 1, iy_in).rgbGreen
                    v22 = input_pixels(ix_in + 1, iy_in + 1).rgbGreen
                    .rgbGreen = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1

                    ' Calculate the blue value.
                    v11 = input_pixels(ix_in, iy_in).rgbBlue
                    v12 = input_pixels(ix_in, iy_in + 1).rgbBlue
                    v21 = input_pixels(ix_in + 1, iy_in).rgbBlue
                    v22 = input_pixels(ix_in + 1, iy_in + 1).rgbBlue
                    .rgbBlue = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
                End With
            End If
        Next ix_out
    Next iy_out

    ' Set pic_to's pixels.
    SetBitmapPixels pic_to, bits_per_pixel, result_pixels
    pic_to.Picture = pic_to.Image
End Sub

