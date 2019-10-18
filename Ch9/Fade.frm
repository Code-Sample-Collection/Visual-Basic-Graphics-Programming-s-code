VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFade 
   Caption         =   "Fade"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   3885
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumFrames 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "10"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdFade 
      Caption         =   "Fade"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtBaseName 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCanvas 
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   120
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblFrameNumber 
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Num Frames"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Base Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmFade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Make the fade frames.
Private Sub cmdFade_Click()
Dim num_frames As Integer
Dim base_name As String
Dim old_pixels() As RGBTriplet
Dim new_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim fraction As Single

    If Not IsNumeric(txtNumFrames.Text) Then txtNumFrames.Text = "10"
    num_frames = CInt(txtNumFrames.Text)
    base_name = txtBaseName.Text

    ' Get the input pixels.
    GetBitmapPixels picCanvas, old_pixels, bits_per_pixel

    ' Make room for the output pixels.
    ReDim new_pixels(0 To UBound(old_pixels, 1), 0 To UBound(old_pixels, 2))

    ' Build the frames.
    For i = 1 To num_frames
        lblFrameNumber.Caption = Format$(i)
        DoEvents

        fraction = (num_frames - i) / num_frames
        For X = 0 To picCanvas.ScaleWidth - 1
            For Y = 0 To picCanvas.ScaleHeight - 1
                With new_pixels(X, Y)
                    .rgbRed = fraction * old_pixels(X, Y).rgbRed
                    .rgbGreen = fraction * old_pixels(X, Y).rgbGreen
                    .rgbBlue = fraction * old_pixels(X, Y).rgbBlue
                End With
            Next Y
        Next X

        ' Update the image.
        SetBitmapPixels picCanvas, bits_per_pixel, new_pixels
        picCanvas.Picture = picCanvas.Image

        ' Save the results.
        SavePicture picCanvas.Picture, base_name & Format$(i) & ".bmp"
    Next i

    ' Restore the original image.
    SetBitmapPixels picCanvas, bits_per_pixel, old_pixels
    picCanvas.Picture = picCanvas.Image
    lblFrameNumber.Caption = ""
End Sub

' Start in the current directory.
Private Sub Form_Load()
Dim base_name As String

    base_name = App.Path
    If Right$(base_name, 1) <> "\" Then base_name = base_name & "\"
    txtBaseName = base_name & "Fade_"

    picCanvas.AutoSize = True
    picCanvas.ScaleMode = vbPixels
    picCanvas.AutoRedraw = True

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

    ' Open the file.
    On Error GoTo LoadError
    picCanvas.Picture = LoadPicture(file_name)
    On Error GoTo 0
    picCanvas.Picture = picCanvas.Image

    Screen.MousePointer = vbDefault
    Exit Sub

LoadError:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Format$(Err.Number) & _
        " opening file '" & file_name & "'" & vbCrLf & _
        Err.Description
End Sub
