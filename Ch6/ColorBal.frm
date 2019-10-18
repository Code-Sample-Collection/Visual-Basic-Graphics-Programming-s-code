VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBright 
   Caption         =   "ColorBal []"
   ClientHeight    =   3960
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5160
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.HScrollBar hbarBrightness 
      Height          =   255
      Index           =   2
      Left            =   720
      Max             =   100
      Min             =   -100
      TabIndex        =   8
      Top             =   720
      Width           =   2415
   End
   Begin VB.HScrollBar hbarBrightness 
      Height          =   255
      Index           =   1
      Left            =   720
      Max             =   100
      Min             =   -100
      TabIndex        =   5
      Top             =   420
      Width           =   2415
   End
   Begin VB.HScrollBar hbarBrightness 
      Height          =   255
      Index           =   0
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
      Top             =   1080
      Width           =   2415
   End
   Begin VB.PictureBox picResult 
      Height          =   2775
      Left            =   2640
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Blue"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblBrighhtness 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Green"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   420
      Width           =   495
   End
   Begin VB.Label lblBrighhtness 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   420
      Width           =   495
   End
   Begin VB.Label lblBrighhtness 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
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
Dim r_factor As Single
Dim g_factor As Single
Dim b_factor As Single
Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer

    ' Get the selected color values.
    r_factor = hbarBrightness(0).value / 100#
    g_factor = hbarBrightness(1).value / 100#
    b_factor = hbarBrightness(2).value / 100#

    ' Get the pixels from picOriginal.
    GetBitmapPixels picOriginal, pixels, bits_per_pixel

    ' Set the pixel colors.
    For Y = 0 To picOriginal.ScaleHeight - 1
        For X = 0 To picOriginal.ScaleWidth - 1
            With pixels(X, Y)
                If r_factor < 0 Then
                    .rgbRed = (1 + r_factor) * .rgbRed
                Else
                    .rgbRed = .rgbRed + r_factor * (255 - .rgbRed)
                End If
                If g_factor < 0 Then
                    .rgbGreen = (1 + g_factor) * .rgbGreen
                Else
                    .rgbGreen = .rgbGreen + g_factor * (255 - .rgbGreen)
                End If
                If b_factor < 0 Then
                    .rgbBlue = (1 + b_factor) * .rgbBlue
                Else
                    .rgbBlue = .rgbBlue + b_factor * (255 - .rgbBlue)
                End If
            End With
        Next X
    Next Y

    ' Set picResult's pixels.
    SetBitmapPixels picResult, bits_per_pixel, pixels
    picResult.Picture = picResult.Image
End Sub

' Transform the image.
Private Sub cmdRefresh_Click()
    If picResult.Picture <> 0 Then
        Screen.MousePointer = vbHourglass
        DoEvents

        TransformImage

        Screen.MousePointer = vbDefault
    End If
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

' Display the brightness value selected.
Private Sub hbarBrightness_Change(Index As Integer)
    lblBrighhtness(Index).Caption = Format$(hbarBrightness(Index).value)
End Sub

' Display the brightness value selected.
Private Sub hbarBrightness_Scroll(Index As Integer)
    lblBrighhtness(Index).Caption = Format$(hbarBrightness(Index).value)
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
    Caption = "ColorBal [" & dlgOpenFile.FileTitle & "]"

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
    Caption = "ColorBal [" & dlgOpenFile.FileTitle & "]"

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

