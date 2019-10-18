VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMorph 
   Caption         =   "Morph [ -> ]"
   ClientHeight    =   3120
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   ScaleHeight     =   3120
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFrames 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "10"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtBaseName 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   5055
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2010
      Index           =   1
      Left            =   2280
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   3
      Top             =   960
      Width           =   2010
   End
   Begin VB.CommandButton cmdMorph 
      Caption         =   "Morph"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox picResult 
      Height          =   2010
      Left            =   4440
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   1
      Top             =   960
      Width           =   2010
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2010
      Index           =   0
      Left            =   120
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   0
      Top             =   960
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Frames"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Base Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open File &1..."
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open File &2..."
         Index           =   1
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFileGridSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLoadCoordinates 
         Caption         =   "&Load Grid Coordinates..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSaveCoordinates 
         Caption         =   "&Save Grid Coordinates..."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Grids"
      Begin VB.Menu mnuGridReset 
         Caption         =   "&Reset"
         Begin VB.Menu mnuGridResetGrid 
            Caption         =   "&Left"
            Index           =   0
         End
         Begin VB.Menu mnuGridResetGrid 
            Caption         =   "&Right"
            Index           =   1
         End
      End
      Begin VB.Menu mnuGridSwap 
         Caption         =   "&Swap"
      End
      Begin VB.Menu mnuGridCopy 
         Caption         =   "&Copy Grid"
         Begin VB.Menu mnuGridCopyRightToLeft 
            Caption         =   "-->"
         End
         Begin VB.Menu mnuGridCopyLeftToRight 
            Caption         =   "<--"
         End
      End
      Begin VB.Menu mnuGridSym 
         Caption         =   "&Symmetry"
         Begin VB.Menu mnuGridSymLeft 
            Caption         =   "&Left Grid"
            Begin VB.Menu mnuGridSymLeftLeftToRight 
               Caption         =   "-->"
            End
            Begin VB.Menu mnuGridSymLeftRightToLeft 
               Caption         =   "<--"
            End
         End
         Begin VB.Menu mnuGridSymRight 
            Caption         =   "&Right Grid"
            Begin VB.Menu mnuGridSymRightLeftToRight 
               Caption         =   "-->"
            End
            Begin VB.Menu mnuGridSymRightRightToLeft 
               Caption         =   "<--"
            End
         End
      End
   End
End
Attribute VB_Name = "frmMorph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_POINTS = 7
Private PointX(0 To 1, 0 To NUM_POINTS + 1, 0 To NUM_POINTS + 1) As Single
Private PointY(0 To 1, 0 To NUM_POINTS + 1, 0 To NUM_POINTS + 1) As Single
Private GridDx As Single
Private GridDy As Single

Private MorphGridX(0 To NUM_POINTS + 1, 0 To NUM_POINTS + 1) As Single
Private MorphGridY(0 To NUM_POINTS + 1, 0 To NUM_POINTS + 1) As Single

Private Dragging As Boolean
Private DragR As Integer
Private DragC As Integer

Private FileTitle(0 To 1) As String

Private Morphing As Boolean
' Make the grid symmetric by copying its left
' half to its right half.
Private Sub MakeLeftToRightSymmetric(ByVal Index As Integer)
Dim row As Integer
Dim col As Integer
Dim mid_col As Integer
Dim mid_x As Single

    mid_x = picOriginal(Index).ScaleWidth / 2
    mid_col = (NUM_POINTS + 1) \ 2
    For row = 0 To NUM_POINTS + 1
        For col = 0 To mid_col
            PointX(Index, row, NUM_POINTS + 1 - col) = _
                mid_x + mid_x - PointX(Index, row, col)
            PointY(Index, row, NUM_POINTS + 1 - col) = _
                PointY(Index, row, col)
        Next col
    Next row

    ' Redraw the grid.
    DrawGrid Index
End Sub

' Make the grid symmetric by copying its right
' half to its left half.
Private Sub MakeRightToLeftSymmetric(ByVal Index As Integer)
Dim row As Integer
Dim col As Integer
Dim mid_col As Integer
Dim mid_x As Single

    mid_x = picOriginal(Index).ScaleWidth / 2
    mid_col = (NUM_POINTS + 1) \ 2
    For row = 0 To NUM_POINTS + 1
        For col = 0 To mid_col
            PointX(Index, row, col) = _
                mid_x + mid_x - PointX(Index, row, NUM_POINTS + 1 - col)
            PointY(Index, row, col) = _
                PointY(Index, row, NUM_POINTS + 1 - col)
        Next col
    Next row

    ' Redraw the grid.
    DrawGrid Index
End Sub


' Using s and t values, return the coordinates of a
' point in a quadrilateral.
Private Sub STToPoints(ByRef X As Single, ByRef Y As Single, ByVal s As Single, ByVal t As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single)
Dim xa As Single
Dim ya As Single
Dim xb As Single
Dim yb As Single

    xa = x1 + t * (x2 - x1)
    ya = y1 + t * (y2 - y1)
    xb = x3 + t * (x4 - x3)
    yb = y3 + t * (y4 - y3)
    X = xa + s * (xb - xa)
    Y = ya + s * (yb - ya)
End Sub

' Find S and T for the point (X, Y) in the
' quadrilateral with points (x1, y1), (x2, y2),
' (x3, y3), and (x4, y4). Return True if the point
' lies within the quadrilateral and False otherwise.
Private Function PointsToST(ByVal X As Single, ByVal Y As Single, ByRef s As Single, ByRef t As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As Boolean
Dim Ax As Single
Dim Bx As Single
Dim Cx As Single
Dim Dx As Single
Dim Ex As Single
Dim Ay As Single
Dim By As Single
Dim Cy As Single
Dim Dy As Single
Dim Ey As Single
Dim a As Single
Dim b As Single
Dim c As Single
Dim det As Single
Dim denom As Single

    Ax = x2 - x1: Ay = y2 - y1
    Bx = x4 - x3: By = y4 - y3
    Cx = x3 - x1: Cy = y3 - y1
    Dx = X - x1: Dy = Y - y1
    Ex = Bx - Ax: Ey = By - Ay

    a = -Ax * Ey + Ay * Ex
    b = Ey * Dx - Dy * Ex + Ay * Cx - Ax * Cy
    c = Dx * Cy - Dy * Cx

    det = b * b - 4 * a * c
    If det >= 0 Then
        If Abs(a) < 0.001 Then
            t = -c / b
        Else
            t = (-b - Sqr(det)) / (2 * a)
        End If
        denom = (Cx + Ex * t)
        If Abs(denom) > 0.001 Then
            s = (Dx - Ax * t) / denom
        Else
            denom = (Cy + Ey * t)
            If Abs(denom) > 0.001 Then
                s = (Dy - Ay * t) / denom
            Else
                s = -1
            End If
        End If

        PointsToST = _
            (t >= -0.00001 And t <= 1.00001 And _
             s >= -0.00001 And s <= 1.00001)
    Else
        PointsToST = False
    End If
End Function

' Arrange the controls.
Private Sub ArrangeControls()
    picOriginal(1).Left = picOriginal(0).Left + picOriginal(0).Width + 60
    picResult.Move picOriginal(1).Left + picOriginal(1).Width + 60, _
        picOriginal(1).Top, picOriginal(0).Width, picOriginal(0).Height
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    picResult.Picture = picResult.Image
    Width = picResult.Left + picResult.Width + 120 + Width - ScaleWidth
    Height = picResult.Top + picResult.Height + 120 + Height - ScaleHeight
    txtBaseName.Width = ScaleWidth - txtBaseName.Left - 120
End Sub

' Copy points from picture fr_index to picture to_index.
Private Sub CopyPoints(ByVal fr_index As Integer, ByVal to_index As Integer)
Dim r As Integer
Dim c As Integer

    ' Copy the points.
    For r = 0 To NUM_POINTS + 1
        For c = 0 To NUM_POINTS + 1
            PointX(to_index, r, c) = PointX(fr_index, r, c)
            PointY(to_index, r, c) = PointY(fr_index, r, c)
        Next c
    Next r

    ' Redraw the grids.
    DrawGrid 0
    DrawGrid 1
End Sub

' Set the file dialog's filters for graphic files.
Private Sub SetFiltersGraphic()
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

' Set the file dialog's filters for text files.
Private Sub SetFiltersText()
    dlgOpenFile.Filter = _
        "Morph Grid Files (*.mor)|*.mor|" & _
        "Text Files (*.txt)|*.txt|" & _
        "All Files (*.*)|*.*"
End Sub


' Create the morph frames.
Private Sub cmdMorph_Click()
Dim num_frames As Integer
Dim frame As Integer
Dim base_name As String
Dim Dx As Single
Dim Dy As Single
Dim start_time As Single
Dim stop_time As Single
Dim minutes As Integer

    ' Do nothing if the pictures are not loaded.
    If (picOriginal(0).Picture = 0) Or _
       (picOriginal(1).Picture = 0) _
    Then
        MsgBox "You must load pictures before morphing."
        Exit Sub
    End If

    On Error Resume Next
    num_frames = CInt(txtFrames.Text)
    If Err.Number <> 0 Then num_frames = 10
    On Error GoTo 0

    base_name = txtBaseName.Text

    ' Prepare for the transformation.
    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents
    start_time = Timer
    Morphing = True

    ' Restore the original images.
    picOriginal(0).Cls
    picOriginal(1).Cls

    ' Save frame 0.
    SavePicture picOriginal(0).Picture, base_name & "0.bmp"
    picResult.Picture = picOriginal(0).Image

    ' Make the frames.
    For frame = 1 To num_frames
        txtFrames.Text = Format$(frame)
        DoEvents

        ' Create the frame.
        picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
            picResult.BackColor, BF
        MorphImage frame / (num_frames + 1)
        picResult.Picture = picResult.Image

        ' Save the result.
        SavePicture picResult.Picture, base_name & Format$(frame) & ".bmp"
    Next frame

    ' Save the last frame.
    SavePicture picOriginal(1).Picture, base_name & Format$(num_frames) & ".bmp"
    picResult.Picture = picOriginal(1).Image
    txtFrames.Text = Format$(num_frames + 1)

    ' Redraw the grids.
    DrawGrid 0
    DrawGrid 1

    stop_time = Timer
    minutes = (stop_time - start_time) \ 60
    MsgBox "Ellapsed time: " & _
        Format$(minutes) & ":" & _
        Format$((stop_time - start_time) - minutes * 60, "00")
    Screen.MousePointer = vbDefault
    Morphing = False
End Sub
' Draw the positioning grid.
Private Sub DrawGrid(ByVal Index As Integer)
Dim r As Integer
Dim c As Integer

    picOriginal(Index).Cls

    ' Draw the lines.
    For r = 0 To NUM_POINTS
        For c = 0 To NUM_POINTS
            If r > 0 Then
                picOriginal(Index).Line _
                    (PointX(Index, r, c), PointY(Index, r, c))- _
                    (PointX(Index, r, c + 1), PointY(Index, r, c + 1))
            End If
            If c > 0 Then
                picOriginal(Index).Line _
                    (PointX(Index, r, c), PointY(Index, r, c))- _
                    (PointX(Index, r + 1, c), PointY(Index, r + 1, c))
            End If
        Next c
    Next r

    ' Draw the control points.
    For r = 0 To NUM_POINTS + 1
        For c = 0 To NUM_POINTS + 1
            picOriginal(Index).Line _
                (PointX(Index, r, c) - 1, PointY(Index, r, c) - 1)- _
                Step(3, 3), , BF
        Next c
    Next r
End Sub
' Find the control point at this mouse position.
Private Sub FindControlPoint(ByVal Index As Integer, ByVal X As Single, ByVal Y As Single, ByRef r As Integer, ByRef c As Integer)
Dim Dx As Single
Dim Dy As Single

    For r = 0 To NUM_POINTS + 1
        For c = 0 To NUM_POINTS + 1
            Dx = Abs(PointX(Index, r, c) - X)
            Dy = Abs(PointY(Index, r, c) - Y)
            If (Dx < 2) And (Dy < 2) Then Exit Sub
        Next c
    Next r

    ' The mouse is not over a control point.
    r = -1
    c = -1
End Sub
' Initialize the positioning grid for this picture.
Private Sub InitializeGrid(ByVal Index As Integer)
Dim X As Single
Dim Y As Single
Dim r As Integer
Dim c As Integer

    GridDx = picOriginal(Index).ScaleWidth / (NUM_POINTS + 1)
    GridDy = picOriginal(Index).ScaleHeight / (NUM_POINTS + 1)
    Y = 0
    For r = 0 To NUM_POINTS + 1
        X = 0
        For c = 0 To NUM_POINTS + 1
            PointX(Index, r, c) = X
            PointY(Index, r, c) = Y
            X = X + GridDx
        Next c
        Y = Y + GridDy
    Next r
End Sub

' Create one frame in the animation.
Private Sub MorphImage(ByVal fraction As Single)
Dim input_pixels0() As RGBTriplet
Dim input_pixels1() As RGBTriplet
Dim result_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim row As Integer
Dim col As Integer
Dim ix_max As Single
Dim iy_max As Single
Dim x_in As Single
Dim y_in As Single
Dim ix_out As Long
Dim iy_out As Long
Dim ix_in As Long
Dim iy_in As Long
Dim Dx As Single
Dim Dy As Single
Dim dx1 As Single
Dim dx2 As Single
Dim dy1 As Single
Dim dy2 As Single
Dim v11 As Integer
Dim v12 As Integer
Dim v21 As Integer
Dim v22 As Integer
Dim r0 As Integer
Dim g0 As Integer
Dim b0 As Integer
Dim r1 As Integer
Dim g1 As Integer
Dim b1 As Integer
Dim s As Single
Dim t As Single
Dim found_grid As Boolean

    ' Get the input pixels.
    GetBitmapPixels picOriginal(0), input_pixels0, bits_per_pixel
    GetBitmapPixels picOriginal(1), input_pixels1, bits_per_pixel

    ' Get the pixels from pic_to.
    GetBitmapPixels picResult, result_pixels, bits_per_pixel

    ' Get the original image's bounds.
    ix_max = picOriginal(0).ScaleWidth - 2
    iy_max = picOriginal(0).ScaleHeight - 2

    ' See where the grid points should be this fraction
    ' of the way between the start and end grids.
    For row = 0 To NUM_POINTS + 1
        For col = 0 To NUM_POINTS + 1
            MorphGridX(row, col) = PointX(0, row, col) * (1# - fraction) + PointX(1, row, col) * fraction
            MorphGridY(row, col) = PointY(0, row, col) * (1# - fraction) + PointY(1, row, col) * fraction
        Next col
    Next row

    ' Calculate the output pixel values.
    For iy_out = 0 To picOriginal(0).ScaleHeight - 1
        For ix_out = 0 To picOriginal(0).ScaleWidth - 1
            ' Find the row and column in the current
            ' grid that contains this point.
            found_grid = False
            For row = 0 To NUM_POINTS
                For col = 0 To NUM_POINTS
                    If PointsToST(ix_out, iy_out, s, t, _
                        MorphGridX(row, col), MorphGridY(row, col), _
                        MorphGridX(row, col + 1), MorphGridY(row, col + 1), _
                        MorphGridX(row + 1, col), MorphGridY(row + 1, col), _
                        MorphGridX(row + 1, col + 1), MorphGridY(row + 1, col + 1)) _
                    Then
                        ' The point is in this grid.
                        found_grid = True
                        Exit For
                    End If
                Next col
                If found_grid Then Exit For
            Next row
            If found_grid Then
                ' picOriginal(0)
                ' Find the corresponding points
                ' in picOriginal(i).
                STToPoints x_in, y_in, s, t, _
                    PointX(0, row, col), PointY(0, row, col), _
                    PointX(0, row, col + 1), PointY(0, row, col + 1), _
                    PointX(0, row + 1, col), PointY(0, row + 1, col), _
                    PointX(0, row + 1, col + 1), PointY(0, row + 1, col + 1)
    
                ' Interpolate to find the pixel's value.
                ' Find the nearest integral position.
                ix_in = Int(x_in)
                iy_in = Int(y_in)
    
                ' See if this is out of bounds.
                If (ix_in < 0) Or (ix_in > ix_max) Or _
                   (iy_in < 0) Or (iy_in > iy_max) _
                Then
                    ' The point is outside the image.
                    ' Use black.
                    r0 = 0
                    g0 = 0
                    b0 = 0
                Else
                    ' The point lies within the image.
                    ' Calculate its value.
                    dx1 = x_in - ix_in
                    dy1 = y_in - iy_in
                    dx2 = 1# - dx1
                    dy2 = 1# - dy1
    
                    ' Calculate the red value.
                    v11 = input_pixels0(ix_in, iy_in).rgbRed
                    v12 = input_pixels0(ix_in, iy_in + 1).rgbRed
                    v21 = input_pixels0(ix_in + 1, iy_in).rgbRed
                    v22 = input_pixels0(ix_in + 1, iy_in + 1).rgbRed
                    r0 = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
    
                    ' Calculate the green value.
                    v11 = input_pixels0(ix_in, iy_in).rgbGreen
                    v12 = input_pixels0(ix_in, iy_in + 1).rgbGreen
                    v21 = input_pixels0(ix_in + 1, iy_in).rgbGreen
                    v22 = input_pixels0(ix_in + 1, iy_in + 1).rgbGreen
                    g0 = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
    
                    ' Calculate the blue value.
                    v11 = input_pixels0(ix_in, iy_in).rgbBlue
                    v12 = input_pixels0(ix_in, iy_in + 1).rgbBlue
                    v21 = input_pixels0(ix_in + 1, iy_in).rgbBlue
                    v22 = input_pixels0(ix_in + 1, iy_in + 1).rgbBlue
                    b0 = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
                End If
    
                ' picOriginal(1)
                ' Find the corresponding points
                ' in picOriginal(i).
                STToPoints x_in, y_in, s, t, _
                    PointX(1, row, col), PointY(1, row, col), _
                    PointX(1, row, col + 1), PointY(1, row, col + 1), _
                    PointX(1, row + 1, col), PointY(1, row + 1, col), _
                    PointX(1, row + 1, col + 1), PointY(1, row + 1, col + 1)
    
                ' Interpolate to find the pixel's value.
                ' Find the nearest integral position.
                ix_in = Int(x_in)
                iy_in = Int(y_in)
    
                ' See if this is out of bounds.
                If (ix_in < 0) Or (ix_in > ix_max) Or _
                   (iy_in < 0) Or (iy_in > iy_max) _
                Then
                    ' The point is outside the image.
                    ' Use black.
                    r1 = 0
                    g1 = 0
                    b1 = 0
                Else
                    ' The point lies within the image.
                    ' Calculate its value.
                    dx1 = x_in - ix_in
                    dy1 = y_in - iy_in
                    dx2 = 1# - dx1
                    dy2 = 1# - dy1
    
                    ' Calculate the red value.
                    v11 = input_pixels1(ix_in, iy_in).rgbRed
                    v12 = input_pixels1(ix_in, iy_in + 1).rgbRed
                    v21 = input_pixels1(ix_in + 1, iy_in).rgbRed
                    v22 = input_pixels1(ix_in + 1, iy_in + 1).rgbRed
                    r1 = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
    
                    ' Calculate the green value.
                    v11 = input_pixels1(ix_in, iy_in).rgbGreen
                    v12 = input_pixels1(ix_in, iy_in + 1).rgbGreen
                    v21 = input_pixels1(ix_in + 1, iy_in).rgbGreen
                    v22 = input_pixels1(ix_in + 1, iy_in + 1).rgbGreen
                    g1 = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
    
                    ' Calculate the blue value.
                    v11 = input_pixels1(ix_in, iy_in).rgbBlue
                    v12 = input_pixels1(ix_in, iy_in + 1).rgbBlue
                    v21 = input_pixels1(ix_in + 1, iy_in).rgbBlue
                    v22 = input_pixels1(ix_in + 1, iy_in + 1).rgbBlue
                    b1 = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
                End If
    
                ' Combine the values of the two colors.
                With result_pixels(ix_out, iy_out)
                    .rgbRed = r0 * (1# - fraction) + r1 * fraction
                    .rgbGreen = g0 * (1# - fraction) + g1 * fraction
                    .rgbBlue = b0 * (1# - fraction) + b1 * fraction
                End With
            Else
                With result_pixels(ix_out, iy_out)
                    .rgbRed = 255
                    .rgbGreen = 0
                    .rgbBlue = 0
                End With
            End If ' End if found_grid ...
        Next ix_out
    Next iy_out

    ' Set pic_to's pixels.
    SetBitmapPixels picResult, bits_per_pixel, result_pixels
    picResult.Picture = picResult.Image
End Sub

' Start in the current directory.
Private Sub Form_Load()
Dim i As Integer
Dim file_name As String

    For i = 0 To 1
        picOriginal(i).AutoSize = True
        picOriginal(i).ScaleMode = vbPixels
        picOriginal(i).AutoRedraw = True
        picOriginal(i).ForeColor = vbWhite
    Next i
    picResult.ScaleMode = vbPixels
    picResult.AutoRedraw = True

    dlgOpenFile.CancelError = True
    dlgOpenFile.InitDir = App.Path

    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "morph_"
    txtBaseName.Text = file_name

    ArrangeControls
End Sub
Private Sub mnuFileLoadCoordinates_Click()
Dim file_name As String
Dim fnum As Integer
Dim i As Integer
Dim r As Integer
Dim c As Integer

    ' Let the user select a file.
    SetFiltersText
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

    ' Load the data.
    fnum = FreeFile
    Open file_name For Input As fnum
    For i = 0 To 1
        For r = 0 To NUM_POINTS + 1
            For c = 0 To NUM_POINTS + 1
                Input #fnum, PointX(i, r, c), PointY(i, r, c)
            Next c
        Next r
    Next i
    Close fnum

    ' Redraw the positioning grid.
    For i = 0 To 1
        DrawGrid i
    Next i

    Screen.MousePointer = vbDefault
End Sub
' Load the indicated file.
Private Sub mnuFileOpen_Click(Index As Integer)
Dim file_name As String

    ' Let the user select a file.
    SetFiltersGraphic
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
    FileTitle(Index) = dlgOpenFile.FileTitle
    Caption = "Morph [" & FileTitle(0) & " -> " & FileTitle(1) & "]"

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal(Index).Picture = LoadPicture(file_name)
    On Error GoTo 0
    picOriginal(Index).Picture = picOriginal(Index).Image

    ' Draw the positioning grid.
    InitializeGrid Index
    DrawGrid Index

    ' Arrange the controls.
    If Index = 0 Then
        picOriginal(1).Width = picOriginal(0).Width
        picOriginal(1).Height = picOriginal(0).Height
    Else
        picOriginal(0).Width = picOriginal(1).Width
        picOriginal(0).Height = picOriginal(1).Height
    End If
    ArrangeControls

    Screen.MousePointer = vbDefault
    Exit Sub

LoadError:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Format$(Err.Number) & _
        " opening file '" & file_name & "'" & vbCrLf & _
        Err.Description
End Sub


Private Sub mnuFileSaveCoordinates_Click()
Dim file_name As String
Dim fnum As Integer
Dim i As Integer
Dim r As Integer
Dim c As Integer

    ' Let the user select a file.
    SetFiltersText
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

    ' Save the grid data into the file.
    fnum = FreeFile
    Open file_name For Output As fnum
    For i = 0 To 1
        For r = 0 To NUM_POINTS + 1
            For c = 0 To NUM_POINTS + 1
                Write #fnum, PointX(i, r, c), PointY(i, r, c)
            Next c
        Next r
    Next i
    Close fnum

    Screen.MousePointer = vbDefault
End Sub

' Copy the control points from the right
' picture to the left.
Private Sub mnuGridCopyLeftToRight_Click()
    CopyPoints 1, 0
End Sub

' Copy the control points from the left
' picture to the right.
Private Sub mnuGridCopyRightToLeft_Click()
    CopyPoints 0, 1
End Sub

' Reset the grid.
Private Sub mnuGridResetGrid_Click(Index As Integer)
    InitializeGrid Index
    DrawGrid Index
End Sub

' Swap the left and right grids.
Private Sub mnuGridSwap_Click()
Dim row As Integer
Dim col As Integer
Dim tmp As Single

    For row = 0 To NUM_POINTS + 1
        For col = 0 To NUM_POINTS + 1
            tmp = PointX(0, row, col)
            PointX(0, row, col) = PointX(1, row, col)
            PointX(1, row, col) = tmp
            tmp = PointY(0, row, col)
            PointY(0, row, col) = PointY(1, row, col)
            PointY(1, row, col) = tmp
        Next col
    Next row
    DrawGrid 0
    DrawGrid 1
End Sub

' Make the left grid symmetric by copying its left
' half to its right half.
Private Sub mnuGridSymLeftLeftToRight_Click()
    MakeLeftToRightSymmetric 0
End Sub

' Make the left grid symmetric by copying its right
' half to its left half.
Private Sub mnuGridSymLeftRightToLeft_Click()
    MakeRightToLeftSymmetric 0
End Sub

' Make the right grid symmetric by copying its left
' half to its right half.
Private Sub mnuGridSymRightLeftToRight_Click()
    MakeLeftToRightSymmetric 1
End Sub


' Make the right grid symmetric by copying its right
' half to its left half.
Private Sub mnuGridSymRightRightToLeft_Click()
    MakeRightToLeftSymmetric 1
End Sub


' Start dragging a control point.
Private Sub picOriginal_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Morphing Then Exit Sub

    ' See if the mouse is over a control point.
    FindControlPoint Index, X, Y, DragR, DragC
End Sub

' Move a control point.
Private Sub picOriginal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim row As Integer
Dim col As Integer

    If Morphing Then Exit Sub

    ' Do nothing if we are not dragging.
    If DragR < 0 Then
        ' No drag is in progress.
        ' See if the mouse is over a control point.
        FindControlPoint Index, X, Y, row, col
        If row >= 0 Then
            ' We're over a control point. Display
            ' the crosshair cursor.
            If picOriginal(Index).MousePointer <> vbCrosshair Then
                picOriginal(Index).MousePointer = vbCrosshair
            End If
        Else
            ' We're not over a control point. Display
            ' the default cursor.
            If picOriginal(Index).MousePointer <> vbDefault Then
                picOriginal(Index).MousePointer = vbDefault
            End If
        End If
    Else
        ' A drag is in progress.
        ' Make sure the point stays in bounds.
        If X < 1 Then X = 1
        If X > picOriginal(Index).ScaleWidth Then X = picOriginal(Index).ScaleWidth
        If Y < 1 Then Y = 1
        If Y > picOriginal(Index).ScaleHeight Then Y = picOriginal(Index).ScaleHeight

        ' Make sure edge points stay on the edge.
        If DragC = 0 Then X = 0
        If DragC = NUM_POINTS + 1 Then X = picOriginal(Index).ScaleWidth
        If DragR = 0 Then Y = 0
        If DragR = NUM_POINTS + 1 Then Y = picOriginal(Index).ScaleHeight

        ' Move the control point.
        PointX(Index, DragR, DragC) = X
        PointY(Index, DragR, DragC) = Y

        ' Redraw the control grid.
        DrawGrid Index
    End If
End Sub


' Finish moving a control point.
Private Sub picOriginal_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragR = -1
End Sub
