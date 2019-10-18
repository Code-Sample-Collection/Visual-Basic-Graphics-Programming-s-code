VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBright 
   Caption         =   "Bright []"
   ClientHeight    =   3375
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5160
   LinkTopic       =   "Form2"
   ScaleHeight     =   3375
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hbarBrightness 
      Height          =   255
      Left            =   720
      Max             =   100
      Min             =   -100
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   4560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2775
      Left            =   120
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.PictureBox picResult 
      Height          =   2775
      Left            =   2640
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblBrighhtness 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Darker"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Brighter"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   615
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
Attribute VB_Name = "frmBright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Arrange the controls.
Private Sub ArrangeControls()
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
    Width = picResult.Left + picResult.Width + _
        Width - ScaleWidth + 120
    Height = picResult.Top + picResult.Height + _
        Height - ScaleHeight + 120
    DoEvents
End Sub

' Transform the image.
Private Sub TransformImage()
Dim factor As Single
Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer

    ' Get the selected brightness value.
    factor = hbarBrightness.value / 100#

    ' Get the pixels from picOriginal.
    GetBitmapPixels picOriginal, pixels, bits_per_pixel

    ' Set the pixel colors.
    For Y = 0 To picOriginal.ScaleHeight - 1
        For X = 0 To picOriginal.ScaleWidth - 1
            With pixels(X, Y)
                If factor < 0 Then
                    ' Make the color darker.
                    .rgbRed = (1 + factor) * .rgbRed
                    .rgbGreen = (1 + factor) * .rgbGreen
                    .rgbBlue = (1 + factor) * .rgbBlue
                Else
                    ' Make the color brighter.
                    .rgbRed = .rgbRed + factor * (255 - .rgbRed)
                    .rgbGreen = .rgbGreen + factor * (255 - .rgbGreen)
                    .rgbBlue = .rgbBlue + factor * (255 - .rgbBlue)
                End If
            End With
        Next X
    Next Y

    ' Set picResult's pixels.
    SetBitmapPixels picResult, bits_per_pixel, pixels
    picResult.Picture = picResult.Image
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
End Sub

' Transform the image.
Private Sub hbarBrightness_Change()
    lblBrighhtness.Caption = Format$(hbarBrightness.value)

    ' If an image is loaded, transform it.
    If picOriginal.Picture <> 0 Then
        Screen.MousePointer = vbHourglass
        DoEvents
        TransformImage
        Screen.MousePointer = vbDefault
    End If
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
    Caption = "Bright [" & dlgOpenFile.FileTitle & "]"

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal.Picture = LoadPicture(file_name)
    On Error GoTo 0

    ' Make picResult the same size and position it.
    ArrangeControls

    ' Transform the image.
    TransformImage

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
    Caption = "Bright [" & dlgOpenFile.FileTitle & "]"

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

