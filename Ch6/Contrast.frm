VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContrast 
   Caption         =   "Contrast []"
   ClientHeight    =   4800
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHistogram 
      Height          =   1455
      Index           =   2
      Left            =   6120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   7
      Top             =   0
      Width           =   2880
   End
   Begin VB.PictureBox picHistogram 
      Height          =   1455
      Index           =   1
      Left            =   3120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   6
      Top             =   0
      Width           =   2880
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picHistogram 
      Height          =   1455
      Index           =   0
      Left            =   120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   5
      Top             =   0
      Width           =   2880
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "Adjust"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   1500
      Width           =   855
   End
   Begin VB.HScrollBar hbarBrightness 
      Height          =   255
      Left            =   120
      Max             =   1000
      TabIndex        =   2
      Top             =   1560
      Width           =   4335
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2775
      Left            =   120
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.PictureBox picResult 
      Height          =   2775
      Left            =   2640
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblBrighhtness 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   1560
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
Attribute VB_Name = "frmContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MinIndex(0 To 2) As Integer
Private MaxIndex(0 To 2) As Integer
' Arrange the controls.
Private Sub ArrangeControls()
Dim wid As Single

    ' Position the result PictureBox.
    picResult.Move _
        picOriginal.Left + picOriginal.Width + 120, _
        picOriginal.Top, _
        picOriginal.Width, _
        picOriginal.Height
    picResult.Cls

    ' This makes the image resize itself to
    ' fit the picture.
    picResult.Picture = picResult.Image

    ' Make the form big enough.
    wid = picResult.Left + picResult.Width
    If wid < picHistogram(2).Left + picHistogram(2).Width Then _
        wid = picHistogram(2).Left + picHistogram(2).Width
    Width = wid + Width - ScaleWidth + 120
    Height = picResult.Top + picResult.Height + _
        Height - ScaleHeight + 120
    DoEvents
End Sub

' Transform the image.
Private Sub TransformImage()
Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim r_mid As Integer
Dim g_mid As Integer
Dim b_mid As Integer
Dim r_scale As Single
Dim g_scale As Single
Dim b_scale As Single
Dim r_diff As Integer
Dim g_diff As Integer
Dim b_diff As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim X As Integer
Dim Y As Integer

    ' Get the pixels from picOriginal.
    GetBitmapPixels picOriginal, pixels, bits_per_pixel

    ' Get the middle values for the components.
    r_mid = (MaxIndex(0) + MinIndex(0)) / 2
    g_mid = (MaxIndex(1) + MinIndex(1)) / 2
    b_mid = (MaxIndex(2) + MinIndex(2)) / 2

    ' Calculate the scale factors needed to resize
    ' the color values.
    r_scale = hbarBrightness.value / (MaxIndex(0) - MinIndex(0))
    g_scale = hbarBrightness.value / (MaxIndex(1) - MinIndex(1))
    b_scale = hbarBrightness.value / (MaxIndex(2) - MinIndex(2))

    ' Set the colors for each component separately.
    For Y = 0 To picOriginal.ScaleHeight - 1
        For X = 0 To picOriginal.ScaleWidth - 1
            With pixels(X, Y)
                r_diff = .rgbRed - r_mid
                r_diff = r_diff * r_scale
                r = 127 + r_diff
                If r < 0 Then r = 0
                If r > 255 Then r = 255
                .rgbRed = r

                g_diff = .rgbGreen - g_mid
                g_diff = g_diff * g_scale
                g = 127 + g_diff
                If g < 0 Then g = 0
                If g > 255 Then g = 255
                .rgbGreen = g

                b_diff = .rgbBlue - b_mid
                b_diff = b_diff * b_scale
                b = 127 + b_diff
                If b < 0 Then b = 0
                If b > 255 Then b = 255
                .rgbBlue = b
            End With
        Next X
    Next Y

    ' Set picResult's pixels.
    SetBitmapPixels picResult, bits_per_pixel, pixels
    picResult.Picture = picResult.Image

    ' Show the new brightness histogram.
    ShowHistograms picResult, False
End Sub
' Show the component histograms.
Private Sub ShowHistograms(ByVal picImage As PictureBox, ByVal save_min_max As Boolean)
Dim counts(0 To 2, 0 To 255) As Long
Dim max_count As Long
Dim brightness As Integer
Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim j As Integer

    ' Clear the previous results.
    For i = 0 To 2
        picHistogram(i).Cls
        picHistogram(i).Refresh
    Next i

    ' Get the pixels from picImage.
    GetBitmapPixels picImage, pixels, bits_per_pixel

    ' Count the brightness values.
    For Y = 0 To picImage.ScaleHeight - 1
        For X = 0 To picImage.ScaleWidth - 1
            With pixels(X, Y)
                counts(0, .rgbRed) = counts(0, .rgbRed) + 1
                counts(1, .rgbGreen) = counts(1, .rgbGreen) + 1
                counts(2, .rgbBlue) = counts(2, .rgbBlue) + 1
            End With
        Next X
    Next Y

    ' Find the largest count value.
    For i = 0 To 2
        ' Skip value 0. There tend to be a lot of
        ' them and they dominate things.
        For j = 1 To 255
            If max_count < counts(i, j) _
                Then max_count = counts(i, j)
        Next j
    Next i

    ' Display the brightness histograms.
    For i = 0 To 2
        picHistogram(i).ScaleTop = 1.1 * max_count
        picHistogram(i).ScaleHeight = -1.2 * max_count
        picHistogram(i).ScaleLeft = -1
        picHistogram(i).ScaleWidth = 258
        For brightness = 0 To 255
            If counts(i, brightness) > 0 Then _
                picHistogram(i).Line (brightness, 0)-(brightness + 1, counts(i, brightness)), , BF
        Next brightness
    Next i

    ' Find the largest and smallest non-zero counts.
    If save_min_max Then
        For i = 0 To 2
            MinIndex(i) = 255
            For brightness = 0 To 255
                If counts(i, brightness) > 0 Then
                    MinIndex(i) = brightness
                    Exit For
                End If
            Next brightness

            MaxIndex(i) = 0
            For brightness = 255 To 0 Step -1
                If counts(i, brightness) > 0 Then
                    MaxIndex(i) = brightness
                    Exit For
                End If
            Next brightness
        Next i
    End If
End Sub
' Transform the image.
Private Sub cmdAdjust_Click()
    If picResult.Picture <> 0 Then
        Screen.MousePointer = vbHourglass
        DoEvents

        TransformImage

        Screen.MousePointer = vbDefault
    End If
End Sub
' Start in the current directory.
Private Sub Form_Load()
Dim i As Integer

    picOriginal.AutoSize = True
    picOriginal.ScaleMode = vbPixels
    picOriginal.AutoRedraw = True
    picResult.ScaleMode = vbPixels
    picResult.AutoRedraw = True
    For i = 0 To 2
        picHistogram(i).AutoRedraw = True
    Next i

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
End Sub

' Display the brightness value selected.
Private Sub hbarBrightness_Change()
    lblBrighhtness.Caption = Format$(hbarBrightness.value)
End Sub

' Display the brightness value selected.
Private Sub hbarBrightness_Scroll()
    lblBrighhtness.Caption = Format$(hbarBrightness.value)
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
    Caption = "Contrast [" & dlgOpenFile.FileTitle & "]"

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal.Picture = LoadPicture(file_name)
    On Error GoTo 0

    ' Make picResult the same size and position it.
    ArrangeControls

    ' Make picResult show the same image.
    picResult.Picture = picOriginal.Picture
    DoEvents

    ' Display the brightness histogram.
    ShowHistograms picOriginal, True

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
    Caption = "Contrast [" & dlgOpenFile.FileTitle & "]"

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

