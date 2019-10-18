VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBinCont2 
   Caption         =   "BinCont2 []"
   ClientHeight    =   4425
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   ScaleHeight     =   4425
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHistogram 
      Height          =   1455
      Index           =   2
      Left            =   6120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   0
      Width           =   2880
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2775
      Left            =   120
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.PictureBox picResult 
      Height          =   2775
      Left            =   2640
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   1560
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
Attribute VB_Name = "frmBinCont2"
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
Private Sub TransformImage(ByVal Index As Integer, ByVal cutoff As Single)
Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim test_value As Byte
Dim X As Integer
Dim Y As Integer

    ' Get the pixels from picOriginal.
    GetBitmapPixels picOriginal, pixels, bits_per_pixel

    ' Set the pixel color values.
    For Y = 0 To picOriginal.ScaleHeight - 1
        For X = 0 To picOriginal.ScaleWidth - 1
            With pixels(X, Y)
                Select Case Index
                    Case 0
                        test_value = .rgbRed
                    Case 1
                        test_value = .rgbGreen
                    Case 2
                        test_value = .rgbBlue
                End Select
                If test_value >= cutoff Then
                    .rgbRed = 255
                    .rgbGreen = 255
                    .rgbBlue = 255
                Else
                    .rgbRed = 0
                    .rgbGreen = 0
                    .rgbBlue = 0
                End If
            End With
        Next X
    Next Y

    ' Set picResult's pixels.
    SetBitmapPixels picResult, bits_per_pixel, pixels
    picResult.Picture = picResult.Image
End Sub

' Show the component histograms.
Private Sub ShowHistograms(ByVal picImage As PictureBox)
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
        picHistogram(i).Line _
            (picHistogram(i).ScaleLeft, picHistogram(i).ScaleTop)- _
            Step(picHistogram(i).ScaleWidth, picHistogram(i).ScaleHeight), _
            picHistogram(i).BackColor, BF
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

        picHistogram(i).Picture = picHistogram(i).Image
    Next i
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
    Caption = "BinCont2 [" & dlgOpenFile.FileTitle & "]"

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
    ShowHistograms picOriginal

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
    Caption = "BinCont2 [" & dlgOpenFile.FileTitle & "]"

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

' Set the binary contrast enhancement level.
Private Sub picHistogram_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picOriginal.Picture <> 0 Then
        picHistogram(0).Cls
        picHistogram(1).Cls
        picHistogram(2).Cls
        picHistogram(Index).Line _
            (X, picHistogram(Index).ScaleTop)- _
            Step(0, picHistogram(Index).ScaleHeight), vbRed
        Screen.MousePointer = vbHourglass
        DoEvents

        TransformImage Index, X

        Screen.MousePointer = vbDefault
    End If
End Sub

