VERSION 5.00
Begin VB.Form frmStretch2 
   Caption         =   "Stretch2"
   ClientHeight    =   4965
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   120
      Picture         =   "Stretch2.frx":0000
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   120
      Width           =   2310
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      Height          =   2310
      Index           =   0
      Left            =   2520
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   120
      Width           =   2310
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      Height          =   2310
      Index           =   1
      Left            =   120
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      Top             =   2520
      Width           =   2310
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      Height          =   2310
      Index           =   2
      Left            =   2520
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   0
      Top             =   2520
      Width           =   2310
   End
End
Attribute VB_Name = "frmStretch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FromXmin As Single
Private FromYmin As Single
Private ToXmin As Single
Private ToYmin As Single
Private XScale As Single
Private YScale As Single

' Map the output pixel (ix_out, iy_out) to the input
' pixel (x_in, y_in).
Private Sub MapPixel(ByVal ix_out As Single, ByVal iy_out As Single, ByRef x_in As Single, ByRef y_in As Single)
    x_in = FromXmin + (ix_out - ToXmin) / XScale
    y_in = FromYmin + (iy_out - ToYmin) / YScale
End Sub
' Copy the picture.
Private Sub StretchPicture(ByVal pic_from As PictureBox, ByVal pic_to As PictureBox, ByVal from_xmin As Single, ByVal from_ymin As Single, ByVal from_wid As Single, ByVal from_hgt As Single, ByVal to_xmin As Single, ByVal to_ymin As Single, ByVal to_wid As Single, ByVal to_hgt As Single)
    ' Save mapping values.
    FromXmin = from_xmin
    FromYmin = from_ymin
    ToXmin = to_xmin
    ToYmin = to_ymin
    XScale = to_wid / (from_wid - 1)
    YScale = to_hgt / (from_hgt - 1)

    ' Transform the image.
    TransformImage pic_from, pic_to
End Sub
' Transform the image.
Private Sub TransformImage(ByVal pic_from As PictureBox, ByVal pic_to As PictureBox)
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

    ' Calculate the output pixel values.
    For iy_out = 0 To pic_to.ScaleHeight - 1
        For ix_out = 0 To pic_to.ScaleWidth - 1
            ' Map the pixel value from
            ' (ix_out, iy_out) to (x_in, y_in).
            MapPixel ix_out, iy_out, x_in, y_in

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
' Arrange the controls.
Private Sub ArrangeControls(ByVal the_scale As Single)
Dim new_wid As Single
Dim new_hgt As Single
Dim old_wid As Single
Dim old_hgt As Single

    ' Calculate the result's size.
    new_wid = (picOriginal.ScaleWidth - 1) * the_scale
    new_hgt = (picOriginal.ScaleHeight - 1) * the_scale
    new_wid = ScaleX(new_wid, vbPixels, ScaleMode) + picOriginal.Width - ScaleX(picOriginal.ScaleWidth, vbPixels, ScaleMode)
    new_hgt = ScaleY(new_hgt, vbPixels, ScaleMode) + picOriginal.Height - ScaleY(picOriginal.ScaleHeight, vbPixels, ScaleMode)

    ' Position the result PictureBox.
    picResult.Move _
        picOriginal.Left + picOriginal.Width + 120, _
        picOriginal.Top, new_wid, new_hgt
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    picResult.Picture = picResult.Image
    picResult.Visible = True

    ' This makes the image resize itself to
    ' fit the picture.
    picResult.Picture = picResult.Image

    ' Make the form big enough.
    new_wid = picResult.Left + picResult.Width
    If new_wid < cmdEnlarge.Left + cmdEnlarge.Width _
        Then new_wid = cmdEnlarge.Left + cmdEnlarge.Width
    new_hgt = picResult.Top + picResult.Height
    Move Left, Top, new_wid + 237, new_hgt + 816

    DoEvents
End Sub

' Start in the current directory.
Private Sub Form_Load()
Dim i As Integer
Dim scale_factor As Single
Dim X As Single
Dim Y As Single
Dim orig_wid As Single
Dim orig_hgt As Single
Dim wid As Single
Dim hgt As Single

    Show
    Screen.MousePointer = vbHourglass

    orig_wid = picOriginal.ScaleWidth
    orig_hgt = picOriginal.ScaleHeight

    scale_factor = 4
    For i = 0 To 2
        wid = orig_wid / scale_factor
        hgt = orig_hgt / scale_factor
        X = (orig_wid - wid) / 2
        Y = (orig_hgt - hgt) / 2

        DoEvents
        StretchPicture picOriginal, _
            picResult(i), _
            X, Y, wid, hgt, _
            0, 0, picResult(i).ScaleWidth, picResult(i).ScaleHeight

        scale_factor = scale_factor * 2
    Next i

    Screen.MousePointer = vbDefault
End Sub
