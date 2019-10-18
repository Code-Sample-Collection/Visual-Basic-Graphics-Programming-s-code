VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUnsharp 
   Caption         =   "Unsharp []"
   ClientHeight    =   2865
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2865
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   2760
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
      Top             =   0
      Width           =   2415
   End
   Begin VB.PictureBox picResult 
      Height          =   2775
      Left            =   2640
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   0
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
Attribute VB_Name = "frmUnsharp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Unsharp mask the image.
Private Sub ApplyFilter()
Const BOUND = 1

Dim kernel() As Single
Dim input_pixels() As RGBTriplet
Dim result_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer

    ' Get the pixels from picOriginal.
    GetBitmapPixels picOriginal, input_pixels, bits_per_pixel

    ' Allocate space for the result pixels.
    ReDim result_pixels( _
        LBound(input_pixels, 1) To UBound(input_pixels, 1), _
        LBound(input_pixels, 2) To UBound(input_pixels, 2))

    ' Apply a low pass filter.
    ReDim kernel(-BOUND To BOUND, -BOUND To BOUND)
    For i = -BOUND To BOUND
        For j = -BOUND To BOUND
            kernel(i, j) = 1 / 9
        Next j
    Next i

    ' Set the pixel colors. Note that we
    ' must skip the edges because some of
    ' the kernel values would correspond
    ' to pixels off the image.
    For Y = BOUND To picOriginal.ScaleHeight - 1 - BOUND
        For X = BOUND To picOriginal.ScaleWidth - 1 - BOUND
            ' Start with no color.
            r = 0
            g = 0
            b = 0
            ' Apply the kernel values to
            ' the nearby pixels.
            For i = -BOUND To BOUND
                For j = -BOUND To BOUND
                    With input_pixels(X + i, Y + j)
                        r = r + .rgbRed * kernel(i, j)
                        g = g + .rgbGreen * kernel(i, j)
                        b = b + .rgbBlue * kernel(i, j)
                    End With
                Next j
            Next i

            ' Make sure the values are
            ' between 0 and 255.
            If r < 0 Then r = 0
            If r > 255 Then r = 255
            If g < 0 Then g = 0
            If g > 255 Then g = 255
            If b < 0 Then b = 0
            If b > 255 Then b = 255

            ' Set the output pixel value.
            With result_pixels(X, Y)
                .rgbRed = r
                .rgbGreen = g
                .rgbBlue = b
            End With
        Next X
    Next Y

    ' Subtract from twice the original values.
    For Y = BOUND To picOriginal.ScaleHeight - 1 - BOUND
        For X = BOUND To picOriginal.ScaleWidth - 1 - BOUND
            With result_pixels(X, Y)
                r = 2 * input_pixels(X, Y).rgbRed - .rgbRed
                g = 2 * input_pixels(X, Y).rgbGreen - .rgbGreen
                b = 2 * input_pixels(X, Y).rgbBlue - .rgbBlue
    
                ' Make sure the values are
                ' between 0 and 255.
                If r < 0 Then r = 0
                If r > 255 Then r = 255
                If g < 0 Then g = 0
                If g > 255 Then g = 255
                If b < 0 Then b = 0
                If b > 255 Then b = 255

                ' Set the output pixel value.
                .rgbRed = r
                .rgbGreen = g
                .rgbBlue = b
            End With
        Next X
    Next Y

    ' Set picResult's pixels.
    SetBitmapPixels picResult, bits_per_pixel, result_pixels
    picResult.Picture = picResult.Image
End Sub
' Manage the mouse and apply the image.
Private Sub ApplyTheFilter()
    ' Do nothing if no picture is loaded.
    If picOriginal.Picture = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents

    ' Apply the filter.
    ApplyFilter

    Screen.MousePointer = vbDefault
End Sub

' Arrange the controls.
Private Sub ArrangeControls()
    ' Position the result PictureBox.
    picResult.Move _
        picOriginal.Left + picOriginal.Width + 120, _
        picOriginal.Top, _
        picOriginal.Width, _
        picOriginal.Height
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF

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
    Caption = "Unsharp [" & dlgOpenFile.FileTitle & "]"

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal.Picture = LoadPicture(file_name)
    On Error GoTo 0

    ' Make picResult the same size and position it.
    ArrangeControls

    ' Apply the filter.
    ApplyTheFilter

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
    Caption = "Unsharp [" & dlgOpenFile.FileTitle & "]"

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

