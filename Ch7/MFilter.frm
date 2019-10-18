VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMFilter 
   Caption         =   "MFilter []"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11610
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   11610
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
      Height          =   7035
      Left            =   120
      Picture         =   "MFilter.frx":0000
      ScaleHeight     =   465
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.PictureBox picResult 
      Height          =   7035
      Left            =   7800
      ScaleHeight     =   465
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.PictureBox picMask 
      AutoSize        =   -1  'True
      Height          =   7035
      Left            =   3960
      Picture         =   "MFilter.frx":53922
      ScaleHeight     =   465
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   3
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblFilterType 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   7800
      TabIndex        =   2
      Top             =   0
      Width           =   3735
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
   Begin VB.Menu mnuFilter 
      Caption         =   "Fil&ter"
      Begin VB.Menu mnuFilterIdentity 
         Caption         =   "&Identity"
      End
      Begin VB.Menu mnuFilterLowPass 
         Caption         =   "&Low Pass"
         Begin VB.Menu mnuLowPass 
            Caption         =   "&3x3 Uniform"
            Index           =   3
         End
         Begin VB.Menu mnuLowPass 
            Caption         =   "&5x5 Uniform"
            Index           =   5
         End
         Begin VB.Menu mnuLowPass 
            Caption         =   "&7x7 Uniform"
            Index           =   7
         End
         Begin VB.Menu mnuLowPassSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFilterLowPassPeaked 
            Caption         =   "3x3 Peaked"
            Index           =   3
         End
         Begin VB.Menu mnuFilterLowPassPeaked 
            Caption         =   "5x5 Peaked"
            Index           =   5
         End
         Begin VB.Menu mnuFilterLowPassPeaked 
            Caption         =   "7x7 Peaked"
            Index           =   7
         End
         Begin VB.Menu mnuLowPassSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLowPassStrongPeak 
            Caption         =   "&Strongly Peaked"
         End
      End
      Begin VB.Menu mnuFilterHighPass 
         Caption         =   "&High Pass"
         Begin VB.Menu mnuHighPassVeryWeak 
            Caption         =   "3x3 Very Weak"
         End
         Begin VB.Menu mnuHighPassWeak 
            Caption         =   "3x3 &Weak"
         End
         Begin VB.Menu mnuHighPassStrong 
            Caption         =   "3x3 &Strong"
         End
         Begin VB.Menu mnuHighPassVeryStrong 
            Caption         =   "3x3 &Very Strong"
         End
      End
      Begin VB.Menu mnuPrewittGradient 
         Caption         =   "&Prewitt Gradient Edge Detection"
         Begin VB.Menu mnuPrewitt 
            Caption         =   "NW to SE"
            Index           =   0
         End
         Begin VB.Menu mnuPrewitt 
            Caption         =   "N to S"
            Index           =   1
         End
         Begin VB.Menu mnuPrewitt 
            Caption         =   "NE to SW"
            Index           =   2
         End
         Begin VB.Menu mnuPrewitt 
            Caption         =   "E to W"
            Index           =   3
         End
         Begin VB.Menu mnuPrewitt 
            Caption         =   "SE to NW"
            Index           =   4
         End
         Begin VB.Menu mnuPrewitt 
            Caption         =   "S to N"
            Index           =   5
         End
         Begin VB.Menu mnuPrewitt 
            Caption         =   "SW to NE"
            Index           =   6
         End
         Begin VB.Menu mnuPrewitt 
            Caption         =   "W to E"
            Index           =   7
         End
      End
      Begin VB.Menu mnuLaplacianEdgeDetection 
         Caption         =   "&Laplacian Edge Detection"
         Begin VB.Menu mnuLaplacianWeak 
            Caption         =   "&Weak"
         End
         Begin VB.Menu mnuLaplacianStrong 
            Caption         =   "&Strong"
         End
         Begin VB.Menu mnuLaplacianVeryStrong 
            Caption         =   "&Very Strong"
         End
      End
      Begin VB.Menu mnuFilterSep 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFilterShowFilter 
         Caption         =   "&Show Filter"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilterCustom 
         Caption         =   "&Define Custom Filter"
      End
   End
End
Attribute VB_Name = "frmMFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TheKernel() As Single
' Manage the mouse and apply the image.
Private Sub ApplyTheFilter()
    ' Do nothing if no picture is loaded.
    If picOriginal.Picture = 0 Then Exit Sub

    ' Do nothing if no filter is loaded.
    If Len(lblFilterType.Caption) = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents

    ' Apply the filter.
    ApplyFilter TheKernel

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
    lblFilterType.Move picResult.Left, _
        0, picResult.Width

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

' Apply a filter to an image.
Private Sub ApplyFilter(kernel() As Single)
Dim bound As Integer
Dim input_pixels() As RGBTriplet
Dim mask_pixels() As RGBTriplet
Dim result_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim new_fraction As Single

    ' Get the kernel's bounds.
    bound = UBound(kernel, 1)

    ' Get the pixels from picOriginal.
    GetBitmapPixels picOriginal, input_pixels, bits_per_pixel

    ' Get the mask pixels.
    GetBitmapPixels picMask, mask_pixels, bits_per_pixel

    ' Allocate space for the result pixels.
    ReDim result_pixels( _
        LBound(input_pixels, 1) To UBound(input_pixels, 1), _
        LBound(input_pixels, 2) To UBound(input_pixels, 2))

    ' Set the pixel colors. Note that we
    ' must skip the edges because some of
    ' the kernel values would correspond
    ' to pixels off the image.
    For Y = bound To picOriginal.ScaleHeight - 1 - bound
        For X = bound To picOriginal.ScaleWidth - 1 - bound
            ' See what fraction of the result
            ' should be due to the new value.
            new_fraction = (255 - mask_pixels(X, Y).rgbRed) / 255

            If new_fraction < 0.001 Then
                ' Don't bother to apply the filter.
                ' Set the output pixel equal to
                ' the input pixel.
                result_pixels(X, Y) = input_pixels(X, Y)
            Else
                ' Start with no color.
                r = 0
                g = 0
                b = 0
                ' Apply the kernel values to
                ' the nearby pixels.
                For i = -bound To bound
                    For j = -bound To bound
                        With input_pixels(X + i, Y + j)
                            r = r + .rgbRed * kernel(j, i)
                            g = g + .rgbGreen * kernel(j, i)
                            b = b + .rgbBlue * kernel(j, i)
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
                    .rgbRed = new_fraction * r + (1 - new_fraction) * input_pixels(X, Y).rgbRed
                    .rgbGreen = new_fraction * g + (1 - new_fraction) * input_pixels(X, Y).rgbGreen
                    .rgbBlue = new_fraction * b + (1 - new_fraction) * input_pixels(X, Y).rgbBlue
                End With
            End If
        Next X
    Next Y

    ' Set picResult's pixels.
    SetBitmapPixels picResult, bits_per_pixel, result_pixels
    picResult.Picture = picResult.Image
End Sub

' Copy kernel entries from a variant array of
' variant arrays into a normal array.
Private Sub VariantToArray(ByVal var As Variant, ByRef arr() As Single)
Dim bound As Integer
Dim i As Integer
Dim j As Integer

    bound = UBound(var) \ 2
    ReDim arr(-bound To bound, -bound To bound)
    For i = -bound To bound
        For j = -bound To bound
            arr(i, j) = var(i + bound)(j + bound)
        Next j
    Next i
End Sub

' Start in the current directory.
Private Sub Form_Load()
    picOriginal.AutoSize = True
    picOriginal.ScaleMode = vbPixels
    picOriginal.AutoRedraw = True
    picMask.AutoSize = True
    picMask.ScaleMode = vbPixels
    picMask.AutoRedraw = True
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
    Caption = "MFilter [" & dlgOpenFile.FileTitle & "]"

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
    Caption = "MFilter [" & dlgOpenFile.FileTitle & "]"

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

' Let the user define a custom filter.
Private Sub mnuFilterCustom_Click()
Dim bound As Integer
Dim i As Integer
Dim j As Integer
Dim idx As Integer

    frmCustom.Show vbModal

    If Not frmCustom.Canceled Then
        bound = frmCustom.CustomBound
        ReDim TheKernel(-bound To bound, -bound To bound)
        idx = 0
        For i = -bound To bound
            For j = -bound To bound
                TheKernel(i, j) = CSng(frmCustom.txtCoefficient(idx))
                idx = idx + 1
            Next j
        Next i

        mnuFilterShowFilter.Enabled = True
        lblFilterType.Caption = "Custom " & _
            Format$(bound) & "x" & Format$(bound)
    End If

    Unload frmCustom
End Sub

Private Sub mnuFilterIdentity_Click()
    ' Create an identity kernel.
    ReDim TheKernel(0 To 0, 0 To 0)
    TheKernel(0, 0) = 1#

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Identity"

    ' Apply the filter.
    ApplyTheFilter
End Sub

' Display the filter coefficients.
Private Sub mnuFilterShowFilter_Click()
    frmShowFilter.PrepareForm TheKernel
    frmShowFilter.Show vbModal
End Sub

' Apply a strong high pass filter.
Private Sub mnuHighPassStrong_Click()
    ' Build the kernel.
    VariantToArray Array( _
        Array(0, -1, 0), _
        Array(-1, 5, -1), _
        Array(0, -1, 0)), _
        TheKernel

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Strong High Pass 3x3"
    ApplyTheFilter
End Sub
' Apply a very strong high pass filter.
Private Sub mnuHighPassVeryStrong_Click()
    ' Build the kernel.
    VariantToArray Array( _
        Array(-1, -1, -1), _
        Array(-1, 9, -1), _
        Array(-1, -1, -1)), _
        TheKernel

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Very Strong High Pass 3x3"
    ApplyTheFilter
End Sub

' Apply a very weak high pass filter.
Private Sub mnuHighPassVeryWeak_Click()
    ' Build the kernel.
    VariantToArray Array( _
        Array(-1 / 12, -1 / 12, -1 / 12), _
        Array(-1 / 12, 20 / 12, -1 / 12), _
        Array(-1 / 12, -1 / 12, -1 / 12)), _
        TheKernel

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Weak High Pass 3x3"
    ApplyTheFilter

End Sub

' Apply a weak high pass filter.
Private Sub mnuHighPassWeak_Click()
    ' Build the kernel.
    VariantToArray Array( _
        Array(-1 / 4, -1 / 4, -1 / 4), _
        Array(-1 / 4, 12 / 4, -1 / 4), _
        Array(-1 / 4, -1 / 4, -1 / 4)), _
        TheKernel

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Weak High Pass 3x3"
    ApplyTheFilter
End Sub

' Apply a weak Laplacian edge detection filter.
Private Sub mnuLaplacianWeak_Click()
    ' Build the kernel.
    VariantToArray Array( _
        Array(0, -1, 0), _
        Array(-1, 4, -1), _
        Array(0, -1, 0)), _
        TheKernel

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Weak Laplacian 3x3"
    ApplyTheFilter
End Sub
' Apply a strong Laplacian edge detection filter.
Private Sub mnuLaplacianStrong_Click()
    ' Build the kernel.
    VariantToArray Array( _
        Array(-1, -1, -1), _
        Array(-1, 8, -1), _
        Array(-1, -1, -1)), _
        TheKernel

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Strong Laplacian 3x3"
    ApplyTheFilter
End Sub
' Apply a very strong Laplacian edge detection filter.
Private Sub mnuLaplacianVeryStrong_Click()
    ' Build the kernel.
    VariantToArray Array( _
        Array(-1, -2, -1), _
        Array(-2, 12, -2), _
        Array(-1, -2, -1)), _
        TheKernel

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Very Strong Laplacian 3x3"
    ApplyTheFilter
End Sub

' Apply a low pass filter.
Private Sub mnuLowPass_Click(Index As Integer)
Dim bound As Integer
Dim i As Integer
Dim j As Integer

    ' Build the kernel.
    bound = (Index - 1) \ 2
    ReDim TheKernel(-bound To bound, -bound To bound)
    For i = -bound To bound
        For j = -bound To bound
            TheKernel(i, j) = 1 / (Index * Index)
        Next j
    Next i

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Identity"

    ' Apply the filter.
    lblFilterType.Caption = "Low Pass " & _
        Format$(Index) & "x" & _
        Format$(Index)
    ApplyTheFilter
End Sub
' Apply a peaked low pass filter.
Private Sub mnuFilterLowPassPeaked_Click(Index As Integer)
Dim bound As Integer
Dim i As Integer
Dim j As Integer
Dim total_weight As Integer

    ' Build the kernel.
    bound = (Index - 1) \ 2
    ReDim TheKernel(-bound To bound, -bound To bound)
    For i = -bound To bound
        For j = -bound To bound
            TheKernel(i, j) = 2 * bound + 1 - Abs(i) - Abs(j)
            total_weight = total_weight + TheKernel(i, j)
        Next j
    Next i

    ' Adjust the kernel so the sum of the
    ' coefficients is 1.
    For i = -bound To bound
        For j = -bound To bound
            TheKernel(i, j) = TheKernel(i, j) / total_weight
        Next j
    Next i

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Low Pass Peaked " & _
        Format$(Index) & "x" & _
        Format$(Index)
    ApplyTheFilter
End Sub


' Apply a stongly peaked low pass filter.
Private Sub mnuLowPassStrongPeak_Click()
Dim i As Integer
Dim j As Integer

    ' Build the kernel.
    ReDim TheKernel(-1 To 1, -1 To 1)
    For i = -1 To 1
        For j = -1 To 1
            TheKernel(i, j) = 1 / 20
        Next j
    Next i
    TheKernel(0, 0) = 12 / 20

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Strongly Peaked 3x3"
    ApplyTheFilter
End Sub

' Apply a Prewitt edge detector.
Private Sub mnuPrewitt_Click(Index As Integer)
Dim i As Integer
Dim j As Integer

    ' Build the kernel.
    Select Case Index
        Case 0  ' NW to SE
            VariantToArray Array( _
                Array(1, 1, 1), _
                Array(1, -2, -1), _
                Array(1, -1, -1)), _
                TheKernel
        Case 1  ' N to S
            VariantToArray Array( _
                Array(1, 1, 1), _
                Array(1, -2, 1), _
                Array(-1, -1, -1)), _
                TheKernel
        Case 2  ' NE to SW
            VariantToArray Array( _
                Array(1, 1, 1), _
                Array(-1, -2, 1), _
                Array(-1, -1, 1)), _
                TheKernel
        Case 3  ' E to W
            VariantToArray Array( _
                Array(-1, 1, 1), _
                Array(-1, -2, 1), _
                Array(-1, 1, 1)), _
                TheKernel
        Case 4  ' SE to NW
            VariantToArray Array( _
                Array(-1, -1, 1), _
                Array(-1, -2, 1), _
                Array(1, 1, 1)), _
                TheKernel
        Case 5  ' S to N
            VariantToArray Array( _
                Array(-1, -1, -1), _
                Array(1, -2, 1), _
                Array(1, 1, 1)), _
                TheKernel
        Case 6  ' SW to NE
            VariantToArray Array( _
                Array(1, -1, -1), _
                Array(1, -2, -1), _
                Array(1, 1, 1)), _
                TheKernel
        Case 7  ' W to E
            VariantToArray Array( _
                Array(1, 1, -1), _
                Array(1, -2, -1), _
                Array(1, 1, -1)), _
                TheKernel
    End Select

    ' Prepare some controls.
    mnuFilterShowFilter.Enabled = True
    lblFilterType.Caption = "Prewitt " & _
        mnuPrewitt(Index).Caption
    ApplyTheFilter
End Sub


