VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAverage 
   Caption         =   "Average []"
   ClientHeight    =   4800
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8400
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResult 
      Height          =   2265
      Left            =   5640
      ScaleHeight     =   147
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2265
      Index           =   3
      Left            =   2880
      Picture         =   "Average.frx":0000
      ScaleHeight     =   147
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   4
      Top             =   2400
      Width           =   2655
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2265
      Index           =   2
      Left            =   120
      Picture         =   "Average.frx":12ADA
      ScaleHeight     =   147
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2265
      Index           =   1
      Left            =   2880
      Picture         =   "Average.frx":255B4
      ScaleHeight     =   147
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2265
      Index           =   0
      Left            =   120
      Picture         =   "Average.frx":3808E
      ScaleHeight     =   147
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Transform the image.
Private Sub TransformImage()
Dim pixels0() As RGBTriplet
Dim pixels1() As RGBTriplet
Dim pixels2() As RGBTriplet
Dim pixels3() As RGBTriplet
Dim new_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim X As Integer
Dim Y As Integer

    ' Get the pixels from picOriginal images.
    GetBitmapPixels picOriginal(0), pixels0, bits_per_pixel
    GetBitmapPixels picOriginal(1), pixels1, bits_per_pixel
    GetBitmapPixels picOriginal(2), pixels2, bits_per_pixel
    GetBitmapPixels picOriginal(3), pixels3, bits_per_pixel

    ' Allocate the new_pixels array.
    ReDim new_pixels( _
        LBound(pixels0, 1) To UBound(pixels0, 1), _
        LBound(pixels0, 2) To UBound(pixels0, 2))

    ' Set the pixel color values.
    For Y = 1 To picOriginal(0).ScaleHeight - 2
        For X = 1 To picOriginal(0).ScaleWidth - 2
            r = 0
            g = 0
            b = 0
            With pixels0(X, Y)
                r = r + .rgbRed
                g = g + .rgbGreen
                b = b + .rgbBlue
            End With
            With pixels1(X, Y)
                r = r + .rgbRed
                g = g + .rgbGreen
                b = b + .rgbBlue
            End With
            With pixels2(X, Y)
                r = r + .rgbRed
                g = g + .rgbGreen
                b = b + .rgbBlue
            End With
            With pixels3(X, Y)
                r = r + .rgbRed
                g = g + .rgbGreen
                b = b + .rgbBlue
            End With
            With new_pixels(X, Y)
                .rgbRed = r / 4
                .rgbGreen = g / 4
                .rgbBlue = b / 4
            End With
        Next X
    Next Y

    ' Set picResult's pixels.
    SetBitmapPixels picResult, bits_per_pixel, new_pixels
    picResult.Picture = picResult.Image
End Sub
' Start in the current directory.
Private Sub Form_Load()
Dim i As Integer

    For i = 0 To 3
        picOriginal(i).AutoSize = True
        picOriginal(i).ScaleMode = vbPixels
        picOriginal(i).AutoRedraw = True
    Next i
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

    TransformImage
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
    Caption = "Average [" & dlgOpenFile.FileTitle & "]"

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
