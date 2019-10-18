VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCompose3 
   Caption         =   "Compose3 []"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBackground 
      AutoSize        =   -1  'True
      Height          =   3360
      Left            =   120
      Picture         =   "Compose3.frx":0000
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   274
      TabIndex        =   2
      Top             =   0
      Width           =   4170
   End
   Begin VB.PictureBox picForeground 
      AutoSize        =   -1  'True
      Height          =   3360
      Left            =   4320
      Picture         =   "Compose3.frx":2C462
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   274
      TabIndex        =   1
      Top             =   0
      Width           =   4170
   End
   Begin VB.PictureBox picResult 
      Height          =   3360
      Left            =   2220
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   274
      TabIndex        =   0
      Top             =   3360
      Width           =   4170
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmCompose3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Make a mask from the foreground picture.
Private Sub ComposeImages()
Dim background_pixels() As RGBTriplet
Dim foreground_pixels() As RGBTriplet
Dim result_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim transparent_r As Byte
Dim transparent_g As Byte
Dim transparent_b As Byte
Dim is_transparent As Boolean
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim j As Integer
Dim num_transparent As Integer
Dim foreground_fraction As Single
Dim background_fraction As Single

    ' Get the pixels from the images.
    GetBitmapPixels picBackground, background_pixels, bits_per_pixel
    GetBitmapPixels picForeground, foreground_pixels, bits_per_pixel

    ' Allocate the result pixels.
    ReDim result_pixels( _
        LBound(foreground_pixels, 1) To UBound(foreground_pixels, 1), _
        LBound(foreground_pixels, 2) To UBound(foreground_pixels, 2))

    ' See what the upper left pixel's color is.
    ' We will convert this value into white and other
    ' values into black.
    With foreground_pixels(0, 0)
        transparent_r = .rgbRed
        transparent_g = .rgbGreen
        transparent_b = .rgbBlue
    End With

    ' Set the result color values. Skip the edges so
    ' we can look at adjacent pixels.
    For Y = 1 To picForeground.ScaleHeight - 2
        For X = 1 To picForeground.ScaleWidth - 2
            ' See if the mask pixel is transparent.
            With foreground_pixels(X, Y)
                is_transparent = ( _
                    (.rgbRed = transparent_r) And _
                    (.rgbGreen = transparent_g) And _
                    (.rgbBlue = transparent_b))
            End With
            If is_transparent Then
                ' The foreground pixel's transparent.
                ' Just use the background pixel color.
                With result_pixels(X, Y)
                    .rgbRed = background_pixels(X, Y).rgbRed
                    .rgbGreen = background_pixels(X, Y).rgbGreen
                    .rgbBlue = background_pixels(X, Y).rgbBlue
                End With
            Else
                ' Use a weighted average of the
                ' foreground and background pixels.

                ' See how many adjacent pixels are transparent.
                num_transparent = 0
                For i = -1 To 1
                    For j = -1 To 1
                        With foreground_pixels(X + i, Y + j)
                            If (.rgbRed = transparent_r) And _
                               (.rgbGreen = transparent_g) And _
                               (.rgbBlue = transparent_b) _
                            Then
                                num_transparent = num_transparent + 1
                            End If
                        End With
                    Next j
                Next i

                ' Pick the color.
                background_fraction = num_transparent / 9#
                foreground_fraction = 1# - background_fraction
                With result_pixels(X, Y)
                    .rgbRed = foreground_fraction * foreground_pixels(X, Y).rgbRed + _
                              background_fraction * background_pixels(X, Y).rgbRed
                    .rgbGreen = foreground_fraction * foreground_pixels(X, Y).rgbGreen + _
                                background_fraction * background_pixels(X, Y).rgbGreen
                    .rgbBlue = foreground_fraction * foreground_pixels(X, Y).rgbBlue + _
                               background_fraction * background_pixels(X, Y).rgbBlue
                End With
            End If
        Next X
    Next Y

    ' Copy the background edge pixels.
    For Y = 0 To picForeground.ScaleHeight - 1
        result_pixels(0, Y) = background_pixels(0, Y)
        result_pixels(picForeground.ScaleWidth - 1, Y) = background_pixels(picForeground.ScaleWidth - 1, Y)
    Next Y
    For X = 0 To picForeground.ScaleWidth - 1
        result_pixels(X, 0) = background_pixels(X, 0)
        result_pixels(X, picForeground.ScaleHeight - 1) = background_pixels(X, picForeground.ScaleHeight - 1)
    Next X

    ' Set picResult's pixels.
    SetBitmapPixels picResult, bits_per_pixel, result_pixels
    picResult.Picture = picResult.Image
End Sub
' Start in the current directory.
Private Sub Form_Load()
Dim ctl As Control

    For Each ctl In Controls
        If TypeOf ctl Is PictureBox Then
            ctl.ScaleMode = vbPixels
            ctl.AutoRedraw = True
        End If
    Next ctl
    picBackground.AutoSize = True
    picForeground.AutoSize = True

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

    ' Make the form appear.
    Show
    Screen.MousePointer = vbHourglass
    DoEvents

    ' Compose the images.
    ComposeImages

    Screen.MousePointer = vbDefault
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
    Caption = "Compose [" & dlgOpenFile.FileTitle & "]"

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
