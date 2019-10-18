VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReduce 
   Caption         =   "Reduce []"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   ScaleHeight     =   2895
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
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
   Begin VB.CommandButton cmdReduce 
      Caption         =   "Reduce"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "1.0"
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
      Caption         =   "Scale"
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
Attribute VB_Name = "frmReduce"
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

' Copy the picture.
Private Sub ShrinkPicture(ByVal pic_from As PictureBox, ByVal pic_to As PictureBox, ByVal from_xmin As Single, ByVal from_ymin As Single, ByVal from_wid As Single, ByVal from_hgt As Single, ByVal to_xmin As Single, ByVal to_ymin As Single, ByVal to_wid As Single, ByVal to_hgt As Single)
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
Dim X As Integer
Dim Y As Integer
Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
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

    ' Calculate the output pixel values.
    For iy_out = 0 To pic_to.ScaleHeight - 1
        For ix_out = 0 To pic_to.ScaleWidth - 1
            ' Map the pixel value from
            ' (ix_out, iy_out) to (x_in, y_in).
            x1 = Int(FromXmin + (ix_out - ToXmin) / XScale)
            x2 = Int(FromXmin + (ix_out + 1 - ToXmin) / XScale) - 1
            y1 = Int(FromYmin + (iy_out - ToYmin) / YScale)
            y2 = Int(FromYmin + (iy_out + 1 - ToYmin) / YScale) - 1

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
' Arrange the controls.
Private Sub ArrangeControls(ByVal scale_factor As Single)
Dim new_wid As Single
Dim new_hgt As Single

    ' Calculate the result's size.
    new_wid = picOriginal.ScaleWidth * scale_factor
    new_hgt = picOriginal.ScaleHeight * scale_factor
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
    If new_wid < cmdReduce.Left + cmdReduce.Width _
        Then new_wid = cmdReduce.Left + cmdReduce.Width
    Move Left, Top, new_wid + 237

    DoEvents
End Sub

' Transform the picture.
Private Sub cmdReduce_Click()
Dim scale_factor As Single

    ' Do nothing if no picture is loaded.
    If picOriginal.Picture = 0 Then Exit Sub

    ' Get the scale.
    On Error GoTo ScaleError
    scale_factor = CSng(txtScale.Text)
    On Error GoTo 0

    ' Make sure the scale is no more than 1.
    If scale_factor > 1# Then
        MsgBox "Scale must be no more than 1.0"
        txtScale.SetFocus
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents

    ' Arrange picResult.
    ArrangeControls scale_factor

    ' Transform the image.
    ShrinkPicture picOriginal, picResult, _
        0, 0, _
        picOriginal.ScaleWidth, picOriginal.ScaleHeight, _
        0, 0, _
        picResult.ScaleWidth, picResult.ScaleHeight

    Screen.MousePointer = vbDefault
    Exit Sub

ScaleError:
    MsgBox "Invalid scale"
    txtScale.SetFocus
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
    Caption = "Reduce [" & dlgOpenFile.FileTitle & "]"

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal.Picture = LoadPicture(file_name)
    On Error GoTo 0

    ' Hide picResult.
    picResult.Visible = False
    If cmdReduce.Left + cmdReduce.Width > picOriginal.Left + picOriginal.Width Then
        Width = cmdReduce.Left + cmdReduce.Width + 120 + Width - ScaleWidth
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
    Caption = "Reduce [" & dlgOpenFile.FileTitle & "]"

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

