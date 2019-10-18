VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrans 
   Caption         =   "Trans []"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   ScaleHeight     =   2895
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTransform 
      Caption         =   "Transform"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox picResult 
      Height          =   2295
      Left            =   840
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdRotate 
      Caption         =   "Rotate"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "30"
      Top             =   60
      Width           =   495
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   120
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Angle"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Phi As Single
Private CxIn As Single
Private CyIn As Single
Private CxOut As Single
Private CyOut As Single
' Map the output pixel (ix_out, iy_out) to the input
' pixel (x_in, y_in).
Private Sub MapPixel(ByVal ix_out As Single, ByVal iy_out As Single, ByRef x_in As Single, ByRef y_in As Single)
Dim dx As Single
Dim dy As Single
Dim R As Single
Dim theta As Single

    dx = ix_out - CxOut
    dy = iy_out - CyOut
    R = Sqr(dx * dx + dy * dy)
    theta = ATan2(dy, dx)
    x_in = CxIn + R * Cos(theta + Phi)
    y_in = CyIn + R * Sin(theta + Phi)
End Sub
' Rotate the image.
Private Sub RotateImage(ByVal pic_from As PictureBox, ByVal pic_to As PictureBox, ByVal angle As Single)
Dim white_pixel As RGBTriplet
Dim input_pixels() As RGBTriplet
Dim result_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim xmax_in As Integer
Dim ymax_in As Integer
Dim CxIn As Single
Dim CyIn As Single
Dim CxOut As Single
Dim CyOut As Single
Dim x_in As Single
Dim y_in As Single
Dim ix_in As Integer
Dim iy_in As Integer
Dim ix_out As Integer
Dim iy_out As Integer
Dim dx As Single
Dim dy As Single
Dim radius As Single
Dim theta As Single
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

    ' Get the centers of both images.
    CxIn = pic_from.ScaleWidth / 2
    CyIn = pic_from.ScaleHeight / 2
    CxOut = pic_to.ScaleWidth / 2
    CyOut = pic_to.ScaleHeight / 2

    ' Get the size of the original image.
    xmax_in = pic_from.ScaleWidth - 1
    ymax_in = pic_from.ScaleHeight - 1

    ' Calculate the output pixel values.
    For iy_out = 0 To pic_to.ScaleHeight - 1
        For ix_out = 0 To pic_to.ScaleWidth - 1
            ' Map the pixel value from
            ' (ix_out, iy_out) to (x_in, y_in).
            dx = ix_out - CxOut
            dy = iy_out - CyOut
            radius = Sqr(dx * dx + dy * dy)
            theta = ATan2(dy, dx)
            x_in = CxIn + radius * Cos(theta + angle)
            y_in = CyIn + radius * Sin(theta + angle)

            ' Find the nearest integral position.
            ix_in = Int(x_in)
            iy_in = Int(y_in)

            ' See if this is in bounds.
            If (ix_in >= 0) And (ix_in < xmax_in) And _
               (iy_in >= 0) And (iy_in < ymax_in) _
            Then
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
            Else
                ' The point is outside the image.
                ' Use white.
                result_pixels(ix_out, iy_out) = white_pixel
            End If
        Next ix_out
    Next iy_out

    ' Set pic_to's pixels.
    SetBitmapPixels pic_to, bits_per_pixel, result_pixels
    pic_to.Picture = pic_to.Image
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
Private Sub ArrangeControls(ByVal angle As Single)
Dim new_wid As Single
Dim new_hgt As Single
Dim old_wid As Single
Dim old_hgt As Single

    ' Calculate the result's size.
    old_wid = picOriginal.ScaleWidth
    old_hgt = picOriginal.ScaleHeight
    new_wid = Abs(old_wid * Cos(angle)) + Abs(old_hgt * Sin(angle))
    new_hgt = Abs(old_wid * Sin(angle)) + Abs(old_hgt * Cos(angle))
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
    If new_wid < cmdTransform.Left + cmdTransform.Width _
        Then new_wid = cmdTransform.Left + cmdTransform.Width
    new_hgt = picResult.Top + picResult.Height
    Move Left, Top, new_wid + 237, new_hgt + 816

    DoEvents
End Sub

' Rotate the picture.
Private Sub cmdRotate_Click()
Const PI = 3.14159265

Dim angle As Single
Dim start_time As Single

    ' Do nothing if no picture is loaded.
    If picOriginal.Picture = 0 Then Exit Sub

    ' Get the angle of rotation in radians.
    On Error GoTo AngleError
    angle = CSng(txtAngle.Text) * PI / 180
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents

    ' Arrange picResult.
    ArrangeControls angle

    ' Rotate the image.
    start_time = Timer
    RotateImage picOriginal, picResult, angle
    MsgBox Format$(Timer - start_time, "0.00") & " seconds"

    Screen.MousePointer = vbDefault
    Exit Sub

AngleError:
    MsgBox "Invalid angle"
    txtAngle.SetFocus
End Sub
' Transform the picture.
Private Sub cmdTransform_Click()
Const PI = 3.14159265

Dim start_time As Single

    ' Do nothing if no picture is loaded.
    If picOriginal.Picture = 0 Then Exit Sub

    ' Get the angle of rotation in radians.
    On Error GoTo TransformAngleError
    Phi = CSng(txtAngle.Text) * PI / 180
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents

    ' Arrange picResult.
    ArrangeControls Phi

    ' Get the centers of both images. These are needed
    ' by MapPixel.
    CxIn = picOriginal.ScaleWidth / 2
    CyIn = picOriginal.ScaleHeight / 2
    CxOut = picResult.ScaleWidth / 2
    CyOut = picResult.ScaleHeight / 2

    ' Transform the image.
    start_time = Timer
    TransformImage picOriginal, picResult
    MsgBox Format$(Timer - start_time, "0.00") & " seconds"

    Screen.MousePointer = vbDefault
    Exit Sub

TransformAngleError:
    MsgBox "Invalid angle"
    txtAngle.SetFocus
End Sub

' Start in the current directory.
Private Sub Form_Load()
    picOriginal.AutoSize = True
    picOriginal.ScaleMode = vbPixels
    picOriginal.AutoRedraw = True
    picResult.ScaleMode = vbPixels
    picResult.AutoRedraw = True

    dlgOpenFile.CancelError = True
    dlgOpenFile.InitDir = App.Path
    dlgOpenFile.Filter = _
        "Bitmaps (*.bmp)|*.bmp|" & _
        "GIFs (*.gif)|*.gif|" & _
        "JPEGs (*.jpg)|*.jpg;*.jpeg|" & _
        "Icons (*.ico)|*.ico|" & _
        "Cursors (*.cur)|*.cur|" & _
        "Run-Length Encoded (*.rle)|*.rle|" & _
        "Metafiles (*.wmf)|*.wmf|" & _
        "Enhanced Metafiles (*.emf)|*.emf|" & _
        "Graphic Files|*.bmp;*.gif;*.jpg;*.jpeg;*.ico;*.cur;*.rle;*.wmf;*.emf|" & _
        "All Files (*.*)|*.*"

    Width = cmdTransform.Left + cmdTransform.Width + 120 + Width - ScaleWidth
    Height = picOriginal.Top + picOriginal.Height + 120 + Height - ScaleHeight
End Sub
' Load the indicated file.
Private Sub mnuFileOpen_Click()
Dim file_name As String

    ' Let the user select a file.
    On Error Resume Next
    dlgOpenFile.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    dlgOpenFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
    DoEvents

    file_name = Trim$(dlgOpenFile.FileName)
    dlgOpenFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgOpenFile.FileTitle) - 1)
    Caption = "Trans [" & dlgOpenFile.FileTitle & "]"

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal.Picture = LoadPicture(file_name)
    On Error GoTo 0

    ' Hide picResult.
    picResult.Visible = False
    If cmdTransform.Left + cmdTransform.Width > picOriginal.Left + picOriginal.Width Then
        Width = cmdTransform.Left + cmdTransform.Width + 120 + Width - ScaleWidth
    Else
        Width = picOriginal.Left + picOriginal.Width + 120 + Width - ScaleWidth
    End If
    Height = picOriginal.Top + picOriginal.Height + 120 + Height - ScaleHeight

    Screen.MousePointer = vbDefault
    Exit Sub

LoadError:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Format$(Err.Number) & _
        " opening file '" & file_name & "'" & vbCrLf & _
        Err.Description
End Sub

' Save the transformed image.
Private Sub mnuFileSaveAs_Click()
Dim file_name As String

    ' Let the user select a file.
    On Error Resume Next
    dlgOpenFile.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    dlgOpenFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
    DoEvents

    file_name = Trim$(dlgOpenFile.FileName)
    dlgOpenFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgOpenFile.FileTitle) - 1)
    Caption = "Trans [" & dlgOpenFile.FileTitle & "]"

    ' Save the transformed image into the file.
    On Error GoTo SaveError
    SavePicture picResult.Picture, file_name
    On Error GoTo 0

    Screen.MousePointer = vbDefault
    Exit Sub

SaveError:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file '" & file_name & "'" & vbCrLf & _
        Err.Description
End Sub

