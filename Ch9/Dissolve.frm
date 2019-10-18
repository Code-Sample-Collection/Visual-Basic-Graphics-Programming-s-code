VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDissolve 
   Caption         =   "Dissolve"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5940
   LinkTopic       =   "Form2"
   ScaleHeight     =   3885
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCanvas 
      AutoSize        =   -1  'True
      Height          =   2295
      Index           =   1
      Left            =   2640
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtNumFrames 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "10"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdDissolve 
      Caption         =   "Dissolve"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   855
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
      Index           =   0
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
         Caption         =   "Open &From Image..."
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open &To Image..."
         Index           =   1
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmDissolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Make the fade frames.
Private Sub cmdDissolve_Click()
Dim num_frames As Integer
Dim base_name As String
Dim pic0_pixels() As RGBTriplet
Dim pic1_pixels() As RGBTriplet
Dim new_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim f0 As Single
Dim f1 As Single

    If Not IsNumeric(txtNumFrames.Text) Then txtNumFrames.Text = "10"
    num_frames = CInt(txtNumFrames.Text)
    base_name = txtBaseName.Text

    ' Get the input pixels.
    GetBitmapPixels picCanvas(0), pic0_pixels, bits_per_pixel
    GetBitmapPixels picCanvas(1), pic1_pixels, bits_per_pixel

    ' Make room for the output pixels.
    ReDim new_pixels(0 To UBound(pic0_pixels, 1), 0 To UBound(pic0_pixels, 2))

    ' Build the frames.
    For i = 1 To num_frames
        lblFrameNumber.Caption = Format$(i)
        DoEvents

        f1 = i / num_frames
        f0 = 1 - f1
        For X = 0 To picCanvas(0).ScaleWidth - 1
            For Y = 0 To picCanvas(0).ScaleHeight - 1
                With new_pixels(X, Y)
                    .rgbRed = f0 * pic0_pixels(X, Y).rgbRed + f1 * pic1_pixels(X, Y).rgbRed
                    .rgbGreen = f0 * pic0_pixels(X, Y).rgbGreen + f1 * pic1_pixels(X, Y).rgbGreen
                    .rgbBlue = f0 * pic0_pixels(X, Y).rgbBlue + f1 * pic1_pixels(X, Y).rgbBlue
                End With
            Next Y
        Next X

        ' Update the image.
        SetBitmapPixels picCanvas(0), bits_per_pixel, new_pixels
        picCanvas(0).Picture = picCanvas(0).Image

        ' Save the results.
        SavePicture picCanvas(0).Picture, base_name & Format$(i) & ".bmp"
    Next i

    ' Restore the original image.
    SetBitmapPixels picCanvas(0), bits_per_pixel, pic0_pixels
    picCanvas(0).Picture = picCanvas(0).Image
    lblFrameNumber.Caption = ""
End Sub

' Start in the current directory.
Private Sub Form_Load()
Dim base_name As String
Dim i As Integer

    base_name = App.Path
    If Right$(base_name, 1) <> "\" Then base_name = base_name & "\"
    txtBaseName = base_name & "Diss_"

    For i = 0 To 1
        picCanvas(i).AutoSize = True
        picCanvas(i).ScaleMode = vbPixels
        picCanvas(i).AutoRedraw = True
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
Private Sub mnuFileOpen_Click(Index As Integer)
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
    picCanvas(Index).Picture = LoadPicture(file_name)
    On Error GoTo 0
    picCanvas(Index).Picture = picCanvas(Index).Image

    ' Arrange the controls.
    picCanvas(1).Left = picCanvas(0).Left + picCanvas(0).Width + 120

    Screen.MousePointer = vbDefault
    Exit Sub

LoadError:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Format$(Err.Number) & _
        " opening file '" & file_name & "'" & vbCrLf & _
        Err.Description
End Sub
