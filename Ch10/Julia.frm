VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmJulia 
   Caption         =   "Julia"
   ClientHeight    =   3810
   ClientLeft      =   2370
   ClientTop       =   1320
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   254
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3810
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   0
      Width           =   3810
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save As..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuScaleMnu 
      Caption         =   "&Scale"
      Begin VB.Menu mnuScale 
         Caption         =   "x&2"
         Index           =   2
      End
      Begin VB.Menu mnuScale 
         Caption         =   "x&4"
         Index           =   4
      End
      Begin VB.Menu mnuScale 
         Caption         =   "x&8"
         Index           =   8
      End
      Begin VB.Menu mnuScaleFull 
         Caption         =   "&Full Scale"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptOptions 
         Caption         =   "&Set Options"
      End
      Begin VB.Menu mnuOptSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptMandelbrotSet 
         Caption         =   "&Mandelbrot Set"
         Checked         =   -1  'True
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuOptJuliaSet 
         Caption         =   "&Julia Set"
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu mnuMovie 
      Caption         =   "&Movie"
      Begin VB.Menu mnuMovieCreate 
         Caption         =   "&Create Movie..."
      End
   End
End
Attribute VB_Name = "frmJulia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_DrawingBox As Boolean
Private m_StartX As Single
Private m_StartY As Single
Private m_CurX As Single
Private m_CurY As Single

Private m_Xmin As Single
Private m_Xmax As Single
Private m_Ymin As Single
Private m_Ymax As Single

Public MaxMandelbrotIterations As Integer
Public MaxJuliaIterations As Integer

Public numcolors As Integer
Private m_Colors() As Long

Private Const MIN_X = -2.2
Private Const MAX_X = 1
Private Const MIN_Y = -1.2
Private Const MAX_Y = 1.2

' 0 = Mandelbrot set
' 1 = Julia set
Private Enum FractalTypes
    fractal_Mandelbrot = 0
    fractal_Julia = 1
End Enum
Private m_SelectedFractal As FractalTypes

Private m_Mandelbrot_Xmin As Single
Private m_Mandelbrot_Xmax As Single
Private m_Mandelbrot_Ymin As Single
Private m_Mandelbrot_Ymax As Single
Private m_Julia_ReaC As Single
Private m_Julia_ImaC As Single

' Draw the appropriate fractal.
Private Sub DrawFractal()
    If m_SelectedFractal = fractal_Mandelbrot Then
        DrawMandelbrot
    Else
        DrawJulia
    End If
End Sub


' Return this color's value.
Property Get color(ByVal Index As Integer) As Long
    color = m_Colors(Index)
End Property

' Add this color to the list.
Public Sub AddColor(ByVal new_color As Long)
    numcolors = numcolors + 1
    ReDim Preserve m_Colors(1 To numcolors)
    m_Colors(numcolors) = new_color
End Sub
' Adjust the aspect ratio of the selected
' coordinates so they fit the window properly.
Private Sub AdjustAspect()
Dim want_aspect As Single
Dim picCanvas_aspect As Single
Dim hgt As Single
Dim wid As Single
Dim mid As Single

    want_aspect = (m_Ymax - m_Ymin) / (m_Xmax - m_Xmin)
    picCanvas_aspect = picCanvas.ScaleHeight / picCanvas.ScaleWidth
    If want_aspect > picCanvas_aspect Then
        ' The selected area is too tall and thin.
        ' Make it wider.
        wid = (m_Ymax - m_Ymin) / picCanvas_aspect
        mid = (m_Xmin + m_Xmax) / 2
        m_Xmin = mid - wid / 2
        m_Xmax = mid + wid / 2
    Else
        ' The selected area is too short and wide.
        ' Make it taller.
        hgt = (m_Xmax - m_Xmin) * picCanvas_aspect
        mid = (m_Ymin + m_Ymax) / 2
        m_Ymin = mid - hgt / 2
        m_Ymax = mid + hgt / 2
    End If
End Sub


' Draw the Mandelbrot set.
Private Sub DrawMandelbrot()
' Work until the magnitude squared > 4.
Const MAX_MAG_SQUARED = 4

Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim wid As Long
Dim hgt As Long
Dim clr As Integer
Dim color As Long
Dim i As Integer
Dim j As Integer
Dim ReaC As Double
Dim ImaC As Double
Dim dReaC As Double
Dim dImaC As Double
Dim ReaZ As Double
Dim ImaZ As Double
Dim ReaZ2 As Double
Dim ImaZ2 As Double
Dim r As Integer
Dim b As Integer
Dim g As Integer

    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), vbBlack, BF
    DoEvents

    ' Get the image's pixels.
    GetBitmapPixels picCanvas, pixels, bits_per_pixel

    ' Adjust the coordinate bounds to fit picCanvas.
    AdjustAspect

    ' dReaC is the change in the real part
    ' (X value) for C. dImaC is the change in the
    ' imaginary part (Y value).
    wid = picCanvas.ScaleWidth
    hgt = picCanvas.ScaleHeight
    dReaC = (m_Xmax - m_Xmin) / (wid - 1)
    dImaC = (m_Ymax - m_Ymin) / (hgt - 1)

    ' Calculate the values.
    ReaC = m_Xmin
    For i = 0 To wid - 1
        ImaC = m_Ymin
        For j = 0 To hgt - 1
            ReaZ = 0
            ImaZ = 0
            ReaZ2 = 0
            ImaZ2 = 0
            clr = 1
            Do While clr < MaxMandelbrotIterations And _
                    ReaZ2 + ImaZ2 < MAX_MAG_SQUARED
                ' Calculate Z(clr).
                ReaZ2 = ReaZ * ReaZ
                ImaZ2 = ImaZ * ImaZ
                ImaZ = 2 * ImaZ * ReaZ + ImaC
                ReaZ = ReaZ2 - ImaZ2 + ReaC
                clr = clr + 1
            Loop

            color = m_Colors(1 + clr Mod numcolors)
            With pixels(i, j)
                .rgbRed = color And &HFF&
                .rgbGreen = (color And &HFF00&) \ &H100&
                .rgbBlue = (color And &HFF0000) \ &H10000
            End With

            ImaC = ImaC + dImaC
        Next j
        ReaC = ReaC + dReaC

        ' Let the user know we're not dead.
        If i Mod 10 = 0 Then
            picCanvas.Line (0, 0)-(wid, i), vbWhite, BF
            picCanvas.Refresh
        End If
    Next i

    ' Update the image.
    SetBitmapPixels picCanvas, bits_per_pixel, pixels
    picCanvas.Refresh
    picCanvas.Picture = picCanvas.Image

    Caption = "Julia (" & Format$(m_Xmin) & ", " & _
        Format$(m_Ymin) & ")-(" & _
        Format$(m_Xmax) & ", " & _
        Format$(m_Ymax) & ")"
End Sub
' Draw the Mandelbrot set.
Private Sub DrawJulia()
' Work until the magnitude squared > 4.
Const MAX_MAG_SQUARED = 4

Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim wid As Long
Dim hgt As Long
Dim clr As Long
Dim color As Long
Dim i As Integer
Dim j As Integer
Dim dReaZ0 As Double
Dim dImaZ0 As Double
Dim ReaZ0 As Double
Dim ImaZ0 As Double
Dim ReaZ As Double
Dim ImaZ As Double
Dim ReaZ2 As Double
Dim ImaZ2 As Double
Dim r As Integer
Dim b As Integer
Dim g As Integer

    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), vbBlack, BF
    DoEvents

    ' Get the image's pixels.
    GetBitmapPixels picCanvas, pixels, bits_per_pixel

    ' Adjust the coordinate bounds to fit picCanvas.
    AdjustAspect

    ' dReaZ0 is the change in the real part
    ' (X value) for Z0. dImaZ0 is the change in the
    ' imaginary part (Y value).
    wid = picCanvas.ScaleWidth
    hgt = picCanvas.ScaleHeight
    dReaZ0 = (m_Xmax - m_Xmin) / (wid - 1)
    dImaZ0 = (m_Ymax - m_Ymin) / (hgt - 1)

    ' Calculate the values.
    ReaZ0 = m_Xmin
    For i = 0 To wid - 1
        ImaZ0 = m_Ymin
        For j = 0 To hgt - 1
            ReaZ = ReaZ0
            ImaZ = ImaZ0
            ReaZ2 = ReaZ * ReaZ
            ImaZ2 = ImaZ * ImaZ
            clr = 1
            Do While clr < MaxJuliaIterations And _
                    ReaZ2 + ImaZ2 < MAX_MAG_SQUARED
                ' Calculate Z(clr).
                ReaZ2 = ReaZ * ReaZ
                ImaZ2 = ImaZ * ImaZ
                ImaZ = 2 * ImaZ * ReaZ + m_Julia_ImaC
                ReaZ = ReaZ2 - ImaZ2 + m_Julia_ReaC
                clr = clr + 1
            Loop

            If clr >= MaxJuliaIterations Then
                ' Use a non-background color.
                color = m_Colors(((ReaZ2 + ImaZ2) * _
                    (numcolors - 1)) Mod _
                    (numcolors - 1) + 1)
            Else
                ' Use the background color.
                color = m_Colors(1)
            End If
            With pixels(i, j)
                .rgbRed = color And &HFF&
                .rgbGreen = (color And &HFF00&) \ &H100&
                .rgbBlue = (color And &HFF0000) \ &H10000
            End With

            ImaZ0 = ImaZ0 + dImaZ0
        Next j
        ReaZ0 = ReaZ0 + dReaZ0

        ' Let the user know we're not dead.
        If i Mod 10 = 0 Then
            picCanvas.Line (0, 0)-(wid, i), vbWhite, BF
            picCanvas.Refresh
        End If
    Next i

    ' Update the image.
    SetBitmapPixels picCanvas, bits_per_pixel, pixels
    picCanvas.Refresh
    picCanvas.Picture = picCanvas.Image

    Caption = "Julia (" & Format$(m_Xmin) & ", " & _
        Format$(m_Ymin) & ")-(" & _
        Format$(m_Xmax) & ", " & _
        Format$(m_Ymax) & ")"
End Sub

' Reset the number of colors to 0.
Public Sub ResetColors()
    numcolors = 0
    Erase m_Colors
End Sub

' Display the Julia set.
Private Sub mnuOptJuliaSet_Click()
    If m_SelectedFractal = fractal_Julia Then Exit Sub

    ' Save the current Mandelbrot position.
    m_Mandelbrot_Xmin = m_Xmin
    m_Mandelbrot_Xmax = m_Xmax
    m_Mandelbrot_Ymin = m_Ymin
    m_Mandelbrot_Ymax = m_Ymax

    ' Use the center as C for the Julia set.
    m_Julia_ReaC = (m_Xmin + m_Xmax) / 2
    m_Julia_ImaC = (m_Ymin + m_Ymax) / 2

    mnuOptJuliaSet.Checked = True
    mnuOptMandelbrotSet.Checked = False
    m_SelectedFractal = fractal_Julia

    ' Zoom out.
    mnuScaleFull_Click
End Sub
' Select this kind of fractal.
Private Sub mnuOptMandelbrotSet_Click()
    If m_SelectedFractal = fractal_Mandelbrot Then Exit Sub

    ' Restore the Mandelbrot position.
    m_Xmin = m_Mandelbrot_Xmin
    m_Xmax = m_Mandelbrot_Xmax
    m_Ymin = m_Mandelbrot_Ymin
    m_Ymax = m_Mandelbrot_Ymax

    mnuOptJuliaSet.Checked = False
    mnuOptMandelbrotSet.Checked = True
    m_SelectedFractal = fractal_Mandelbrot

    ' Redraw.
    Screen.MousePointer = vbHourglass
    DrawFractal
    Screen.MousePointer = vbDefault
End Sub

' Start a rubberband box to select a zoom area.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_DrawingBox = True
    m_StartX = X
    m_StartY = Y
    m_CurX = X
    m_CurY = Y
    picCanvas.DrawMode = vbInvert
    picCanvas.Line (m_StartX, m_StartY)-(m_CurX, m_CurY), , B
End Sub


' Continue the zoom area rubberband box.
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_DrawingBox Then Exit Sub
    picCanvas.Line (m_StartX, m_StartY)-(m_CurX, m_CurY), , B
    m_CurX = X
    m_CurY = Y
    picCanvas.Line (m_StartX, m_StartY)-(m_CurX, m_CurY), , B
End Sub


' Zoom in on the selected area.
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single
Dim factor As Single

    If Not m_DrawingBox Then Exit Sub
    m_DrawingBox = False

    picCanvas.Line (m_StartX, m_StartY)-(m_CurX, m_CurY), , B
    picCanvas.DrawMode = vbCopyPen
    m_CurX = X
    m_CurY = Y
    
    ' Put the coordinates in proper order.
    If m_CurX < m_StartX Then
        x1 = m_CurX
        x2 = m_StartX
    Else
        x1 = m_StartX
        x2 = m_CurX
    End If
    If x1 = x2 Then x2 = x1 + 1
    If m_CurY < m_StartY Then
        y1 = m_CurY
        y2 = m_StartY
    Else
        y1 = m_StartY
        y2 = m_CurY
    End If
    If y1 = y2 Then y2 = y1 + 1

    ' Convert screen coords into drawing coords.
    factor = (m_Xmax - m_Xmin) / picCanvas.ScaleWidth
    m_Xmax = m_Xmin + x2 * factor
    m_Xmin = m_Xmin + x1 * factor

    factor = (m_Ymax - m_Ymin) / picCanvas.ScaleHeight
    m_Ymax = m_Ymin + y2 * factor
    m_Ymin = m_Ymin + y1 * factor

    Screen.MousePointer = vbHourglass
    DrawFractal
    Screen.MousePointer = vbDefault
End Sub



' Force Visual Basic to resize the bitmap.
Private Sub picCanvas_Resize()
    picCanvas.Cls
End Sub


' Save the picture.
Private Sub mnuFileSaveAs_Click()
Dim file_name As String

    ' Allow the user to pick a file.
    On Error Resume Next

    dlgFile.DialogTitle = "Save As File"
    dlgFile.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    file_name = Trim$(dlgFile.FileName)
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    ' Save the picture.
    SavePicture picCanvas.Image, file_name
End Sub

' Draw the initial Mandelbrot set.
Private Sub Form_Load()
Dim i As Integer

    Me.Show
    DoEvents

    MaxMandelbrotIterations = 64
    MaxJuliaIterations = 16

    ' Create some default colors.
    ResetColors
    AddColor frmConfig.picColor(40).BackColor
    For i = 17 To 23
        AddColor frmConfig.picColor(i).BackColor
    Next i
    Unload frmConfig

    dlgFile.Filter = "Bitmap Files (*.bmp)|*.bmp|" & _
        "All Files (*.*)|*.*"
    dlgFile.InitDir = App.Path
    dlgFile.CancelError = True

    ' Display the first Mandelbrot set.
    mnuScaleFull_Click
End Sub

Private Sub Form_Resize()
    picCanvas.Move 0, 0, ScaleWidth, ScaleHeight
End Sub



' Let the user set program options.
Private Sub mnuOptOptions_Click()
    frmConfig.Initialize Me
    frmConfig.Show vbModal
End Sub

' Zoom out to full scale.
Private Sub mnuScaleFull_Click()
    m_Xmin = MIN_X
    m_Xmax = MAX_X
    m_Ymin = MIN_Y
    m_Ymax = MAX_Y

    Screen.MousePointer = vbHourglass
    DrawFractal
    Screen.MousePointer = vbDefault
End Sub

' Make a series of images.
Private Sub MakeMovie(file_name As String)
Dim num_frames As Integer
Dim frame As Integer
Dim fraction As Single  ' Amount to reduce image.
Dim xmid As Single      ' Center of image.
Dim ymid As Single
Dim wid1 As Single      ' Starting dimensions.
Dim hgt1 As Single
Dim wid2 As Single      ' Finishing dimensions.
Dim hgt2 As Single
Dim wid As Single       ' Current dimensions.
Dim hgt As Single

Dim start_time As Single
Dim stop_time As Single
Dim max_time As Single
Dim min_time As Single

Dim txt As String
Dim value As Integer

    ' See how may frames the user wants.
    txt = InputBox("Number of frames:", _
        "Frames", "20")
    If txt = "" Then Exit Sub
    If IsNumeric(txt) Then num_frames = CInt(txt)
    If num_frames < 1 Then num_frames = 20

    Screen.MousePointer = vbHourglass
    max_time = 0
    min_time = 100000

    ' Set the center of focus and dimensions.
    xmid = (m_Xmin + m_Xmax) / 2
    ymid = (m_Ymin + m_Ymax) / 2
    wid1 = MAX_X - MIN_X
    wid2 = m_Xmax - m_Xmin

    ' Compute start and finish heights.
    hgt1 = wid1 * picCanvas.ScaleHeight / picCanvas.ScaleWidth
    hgt2 = wid2 * picCanvas.ScaleHeight / picCanvas.ScaleWidth

    ' Compute the amount to reduce the image for
    ' each frame.
    fraction = Exp(Log(wid2 / wid1) / (num_frames - 1))

    ' Start cranking out frames.
    wid = wid1
    hgt = hgt1
    For frame = 0 To num_frames - 1
        Caption = "Julia " & Str$(frame) & _
            "/" & Format$(num_frames - 1)
        m_Xmin = xmid - wid / 2
        m_Xmax = xmid + wid / 2
        m_Ymin = ymid - hgt / 2
        m_Ymax = ymid + hgt / 2

        start_time = Timer
        DrawFractal
        stop_time = Timer

        If min_time > stop_time - start_time Then min_time = stop_time - start_time
        If max_time < stop_time - start_time Then max_time = stop_time - start_time

        SavePicture picCanvas.Image, _
            file_name & Format$(frame) & ".bmp"
        Beep
        DoEvents

        wid = wid * fraction
        hgt = hgt * fraction
    Next frame

    Screen.MousePointer = vbDefault

    MsgBox _
        "Longest:  " & Format$(max_time, "0.00") & _
            " seconds." & vbCrLf & _
        "Shortest: " & Format$(min_time, "0.00") & _
            " seconds." & vbCrLf
End Sub
' Make a series of images.
Private Sub mnuMovieCreate_Click()
Dim old_file_name As String
Dim file_name As String
Dim pos As Integer

    ' Allow the user to pick a file.
    On Error Resume Next
    old_file_name = dlgFile.FileName
    dlgFile.DialogTitle = "Select base file name (no number)"
    dlgFile.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly

    pos = InStr(old_file_name, ".")
    If pos > 0 Then old_file_name = Left$(old_file_name, pos - 1)
    dlgFile.FileName = old_file_name

    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        dlgFile.FileName = old_file_name
        Exit Sub
    ElseIf Err.Number <> 0 Then
        dlgFile.FileName = old_file_name
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    file_name = Trim$(dlgFile.FileName)
    dlgFile.FileName = old_file_name
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    ' Trim off the extension if any.
    pos = InStr(file_name, ".")
    If pos > 0 Then file_name = Left$(file_name, pos - 1)
    
    ' Add a trailing underscore if needed.
    If Right$(file_name, 1) <> "_" Then _
        file_name = file_name & "_"
    
    ' Make the movie.
    MakeMovie file_name
End Sub
' Increase the area shown by a factor of Index.
Private Sub mnuScale_Click(Index As Integer)
Dim size As Single
Dim mid As Single

    size = Index * (m_Xmax - m_Xmin)
    If size > 3.2 Then
        mnuScaleFull_Click
        Exit Sub
    End If
    mid = (m_Xmin + m_Xmax) / 2
    m_Xmin = mid - size / 2
    m_Xmax = mid + size / 2
    
    size = Index * (m_Ymax - m_Ymin)
    If size > 2.4 Then
        mnuScaleFull_Click
        Exit Sub
    End If
    mid = (m_Ymin + m_Ymax) / 2
    m_Ymin = mid - size / 2
    m_Ymax = mid + size / 2
    
    Screen.MousePointer = vbHourglass
    DrawFractal
    Screen.MousePointer = vbDefault
End Sub

