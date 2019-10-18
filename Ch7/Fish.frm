VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFish 
   Caption         =   "Fish []"
   ClientHeight    =   2760
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   ScaleHeight     =   2760
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResult 
      Height          =   2295
      Left            =   2640
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   120
      Width           =   2415
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
      Top             =   120
      Width           =   2415
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
Attribute VB_Name = "frmFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RMax As Single
Private XHit As Single
Private YHit As Single
Private Xmax As Single
Private Ymax As Single

' Map the output pixel (ix_out, iy_out) to the input
' pixel (x_in, y_in).
Private Sub MapPixel(ByVal ix_out As Single, ByVal iy_out As Single, ByRef x_in As Single, ByRef y_in As Single)
Dim dx As Single
Dim dy As Single
Dim r As Single

    ' Find the distance (XHit, YHit)-(ix_out, iy_out).
    dx = ix_out - XHit
    dy = iy_out - YHit
    r = Sqr(dx * dx + dy * dy)

    ' If this is the hit point, do nothing.
    If r < 1 Then
        x_in = ix_out
        y_in = iy_out
    Else
        ' Normalize <dx, dy>.
        dx = dx / r
        dy = dy / r

        ' Convert the distance into a fraction of RMax.
        r = r / RMax

        ' Transform this distance.
        r = r * r

        ' Find the input position.
        x_in = XHit + dx * RMax * r
        y_in = YHit + dy * RMax * r
    End If
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
Dim ix_out As Long
Dim iy_out As Long
Dim ix_in As Long
Dim iy_in As Long
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

    Width = picResult.Left + picResult.Width + 120 + Width - ScaleWidth
    Height = picResult.Top + picResult.Height + 120 + Height - ScaleHeight
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
    Caption = "Fish [" & dlgOpenFile.FileTitle & "]"

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal.Picture = LoadPicture(file_name)
    On Error GoTo 0
    picOriginal.Picture = picOriginal.Image

    ' Arrange the controls.
    picResult.Move _
        picOriginal.Left + picOriginal.Width + 120, _
        picOriginal.Top, picOriginal.Width, _
        picOriginal.Height
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    picResult.Picture = picResult.Image
    Width = picResult.Left + picResult.Width + 120 + Width - ScaleWidth
    Height = picResult.Top + picResult.Height + 120 + Height - ScaleHeight

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
    Caption = "Fish [" & dlgOpenFile.FileTitle & "]"

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

' Transform the image.
Private Sub picOriginal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dx As Single
Dim dy As Single

    ' Do nothing if no picture is loaded.
    If picOriginal.Picture = 0 Then Exit Sub

    ' Prepare for the transformation.
    XHit = X
    YHit = Y
    Xmax = picResult.ScaleWidth
    Ymax = picResult.ScaleHeight

    If X > Xmax / 2 Then
        dx = X
    Else
        dx = Xmax - X
    End If
    If Y > Ymax / 2 Then
        dy = Y
    Else
        dy = Ymax - Y
    End If
    RMax = Sqr(dx * dx + dy * dy)

    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents

    ' Transform the image.
    TransformImage picOriginal, picResult

    Screen.MousePointer = vbDefault
End Sub
