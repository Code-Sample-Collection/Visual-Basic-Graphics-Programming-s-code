VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGrid 
   Caption         =   "Grid []"
   ClientHeight    =   2760
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   ScaleHeight     =   2760
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTransform 
      Caption         =   "Transform"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox picResult 
      Height          =   2295
      Left            =   2640
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   120
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   480
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
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_POINTS = 3
Private PointX(0 To NUM_POINTS + 1, 0 To NUM_POINTS + 1) As Single
Private PointY(0 To NUM_POINTS + 1, 0 To NUM_POINTS + 1) As Single
Private GridDx As Single
Private GridDy As Single

Private Dragging As Boolean
Private DragR As Integer
Private DragC As Integer

' Transform the image.
Private Sub cmdTransform_Click()
Dim Dx As Single
Dim Dy As Single

    ' Do nothing if no picture is loaded.
    If picOriginal.Picture = 0 Then Exit Sub

    ' Prepare for the transformation.
'    Xmax = picResult.ScaleWidth
'    Ymax = picResult.ScaleHeight

    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents

    ' Restore the original image.
    picOriginal.Cls

    ' Transform the image.
    TransformImage picOriginal, picResult

    ' Redraw the grid.
    DrawGrid

    Screen.MousePointer = vbDefault
End Sub

' Draw the positioning grid.
Private Sub DrawGrid()
Dim r As Integer
Dim c As Integer

    picOriginal.Cls

    ' Draw the lines.
    For r = 0 To NUM_POINTS
        For c = 0 To NUM_POINTS
            If r > 0 Then
                picOriginal.Line _
                    (PointX(r, c), PointY(r, c))- _
                    (PointX(r, c + 1), PointY(r, c + 1))
            End If
            If c > 0 Then
                picOriginal.Line _
                    (PointX(r, c), PointY(r, c))- _
                    (PointX(r + 1, c), PointY(r + 1, c))
            End If
        Next c
    Next r

    ' Draw the control points.
    For r = 1 To NUM_POINTS
        For c = 1 To NUM_POINTS
            picOriginal.Line _
                (PointX(r, c) - 1, PointY(r, c) - 1)- _
                Step(3, 3), , BF
        Next c
    Next r
End Sub
' Find the control point at this mouse position.
Private Sub FindControlPoint(ByVal X As Single, ByVal Y As Single, ByRef r As Integer, ByRef c As Integer)
Dim Dx As Single
Dim Dy As Single

    For r = 0 To NUM_POINTS + 1
        For c = 0 To NUM_POINTS + 1
            Dx = Abs(PointX(r, c) - X)
            Dy = Abs(PointY(r, c) - Y)
            If (Dx < 2) And (Dy < 2) Then Exit Sub
        Next c
    Next r

    ' The mouse is not over a control point.
    r = -1
    c = -1
End Sub

' Initialize the positioning grid for this picture.
Private Sub InitializeGrid()
Dim X As Single
Dim Y As Single
Dim r As Integer
Dim c As Integer

    GridDx = picOriginal.ScaleWidth / (NUM_POINTS + 1)
    GridDy = picOriginal.ScaleHeight / (NUM_POINTS + 1)
    Y = 0
    For r = 0 To NUM_POINTS + 1
        X = 0
        For c = 0 To NUM_POINTS + 1
            PointX(r, c) = X
            PointY(r, c) = Y
            X = X + GridDx
        Next c
        Y = Y + GridDy
    Next r
End Sub

' Map the output pixel (ix_out, iy_out) to the input
' pixel (x_in, y_in).
Private Sub MapPixel(ByVal ix_out As Single, ByVal iy_out As Single, ByRef x_in As Single, ByRef y_in As Single)
Dim r As Integer
Dim c As Integer
Dim x0 As Single
Dim y0 As Single
Dim dx1 As Single
Dim dy1 As Single
Dim dx2 As Single
Dim dy2 As Single
Dim v11 As Integer
Dim v12 As Integer
Dim v21 As Integer
Dim v22 As Integer

    ' See in which rectangle the point lies.
    c = Int(ix_out / GridDx)
    r = Int(iy_out / GridDy)

    ' Find the area's upper left corner.
    x0 = c * GridDx
    y0 = r * GridDy

    ' Map to a point in the corresponding quadrilateral
    ' using bilinear interpolation.
    dx1 = (ix_out - x0) / GridDx
    dy1 = (iy_out - y0) / GridDy
    dx2 = 1# - dx1
    dy2 = 1# - dy1

    ' Calculate the X value.
    v11 = PointX(r, c)
    v21 = PointX(r, c + 1)
    v12 = PointX(r + 1, c)
    v22 = PointX(r + 1, c + 1)
    x_in = _
        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
        v21 * dx1 * dy2 + v22 * dx1 * dy1

    ' Calculate the Y value.
    v11 = PointY(r, c)
    v21 = PointY(r, c + 1)
    v12 = PointY(r + 1, c)
    v22 = PointY(r + 1, c + 1)
    y_in = _
        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
        v21 * dx1 * dy2 + v22 * dx1 * dy1
End Sub
' Transform the image.
Private Sub TransformImage(ByVal pic_from As PictureBox, ByVal pic_to As PictureBox)
Dim white_pixel As RGBTriplet
Dim input_pixels() As RGBTriplet
Dim result_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
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

    ' Set the white pixel's value.
    With white_pixel
        .rgbRed = 255
        .rgbGreen = 255
        .rgbBlue = 255
    End With

    ' Get the pixels from pic_from.
    GetBitmapPixels pic_from, input_pixels, bits_per_pixel

    ' Get the pixels from pic_to.
    GetBitmapPixels pic_to, result_pixels, bits_per_pixel

    ' Get the original image's bounds.
    ix_max = pic_from.ScaleWidth - 2
    iy_max = pic_from.ScaleHeight - 2

    ' Calculate the output pixel values.
    For iy_out = 0 To pic_to.ScaleHeight - 1
        For ix_out = 0 To pic_to.ScaleWidth - 1
            ' Map the pixel value from
            ' (ix_out, iy_out) to (x_in, y_in).
            MapPixel ix_out, iy_out, x_in, y_in

            ' Interpolate to find the pixel's value.
            ' Find the nearest integral position.
            ix_in = Int(x_in)
            iy_in = Int(y_in)

            ' See if this is out of bounds.
            If (ix_in < 0) Or (ix_in > ix_max) Or _
               (iy_in < 0) Or (iy_in > iy_max) _
            Then
                ' The point is outside the image.
                ' Use white.
                result_pixels(ix_out, iy_out) = white_pixel
            Else
                ' The point lies within the image.
                ' Calculate its value.
                dx1 = x_in - ix_in
                dy1 = y_in - iy_in
                dx2 = 1# - dx1
                dy2 = 1# - dy1

                With result_pixels(ix_out, iy_out)
                    ' Calculate the red value.
                    v11 = input_pixels(ix_in, iy_in).rgbRed
                    v12 = input_pixels(ix_in, iy_in + 1).rgbRed
                    v21 = input_pixels(ix_in + 1, iy_in).rgbRed
                    v22 = input_pixels(ix_in + 1, iy_in + 1).rgbRed
                    .rgbRed = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
        
                    ' Calculate the green value.
                    v11 = input_pixels(ix_in, iy_in).rgbGreen
                    v12 = input_pixels(ix_in, iy_in + 1).rgbGreen
                    v21 = input_pixels(ix_in + 1, iy_in).rgbGreen
                    v22 = input_pixels(ix_in + 1, iy_in + 1).rgbGreen
                    .rgbGreen = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1

                    ' Calculate the blue value.
                    v11 = input_pixels(ix_in, iy_in).rgbBlue
                    v12 = input_pixels(ix_in, iy_in + 1).rgbBlue
                    v21 = input_pixels(ix_in + 1, iy_in).rgbBlue
                    v22 = input_pixels(ix_in + 1, iy_in + 1).rgbBlue
                    .rgbBlue = _
                        v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
                        v21 * dx1 * dy2 + v22 * dx1 * dy1
                End With
            End If
        Next ix_out
    Next iy_out

    ' Set pic_to's pixels.
    SetBitmapPixels pic_to, bits_per_pixel, result_pixels
    pic_to.Picture = pic_to.Image
End Sub
' Start in the current directory.
Private Sub Form_Load()
    picOriginal.AutoSize = True
    picOriginal.ScaleMode = vbPixels
    picOriginal.AutoRedraw = True
    picOriginal.ForeColor = vbWhite
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

    Width = picResult.Left + picResult.Width + 120 + Width - ScaleWidth
    Height = picResult.Top + picResult.Height + 120 + Height - ScaleHeight
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
    Caption = "Grid [" & dlgOpenFile.FileTitle & "]"

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal.Picture = LoadPicture(file_name)
    On Error GoTo 0
    picOriginal.Picture = picOriginal.Image

    ' Draw the positioning grid.
    InitializeGrid
    DrawGrid

    ' Arrange the controls.
    picResult.Move _
        picOriginal.Left + picOriginal.Width + 120, _
        picOriginal.Top, picOriginal.Width, _
        picOriginal.Height
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    picResult.Picture = picResult.Image
    Width = picResult.Left + picResult.Width + 120 + Width - ScaleWidth
    Height = picResult.Top + picResult.Height + 120 + Height - ScaleHeight

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
    Caption = "Grid [" & dlgOpenFile.FileTitle & "]"

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

' Start dragging a control point.
Private Sub picOriginal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' See if the mouse is over a control point.
    FindControlPoint X, Y, DragR, DragC
End Sub

' Move a control point.
Private Sub picOriginal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim row As Integer
Dim col As Integer

    ' Do nothing if we are not dragging.
    If DragR < 0 Then
        ' No drag is in progress.
        ' See if the mouse is over a control point.
        FindControlPoint X, Y, row, col
        If row >= 0 Then
            ' We're over a control point. Display
            ' the crosshair cursor.
            If picOriginal.MousePointer <> vbCrosshair Then
                picOriginal.MousePointer = vbCrosshair
            End If
        Else
            ' We're not over a control point. Display
            ' the default cursor.
            If picOriginal.MousePointer <> vbDefault Then
                picOriginal.MousePointer = vbDefault
            End If
        End If
    Else
        ' A drag is in progress.
        ' Make sure the point stays in bounds.
        If X < 1 Then X = 1
        If X > picOriginal.ScaleWidth Then X = picOriginal.ScaleWidth
        If Y < 1 Then Y = 1
        If Y > picOriginal.ScaleHeight Then Y = picOriginal.ScaleHeight

        ' Make sure edge points stay on the edge.
        If DragC = 0 Then X = 0
        If DragC = NUM_POINTS + 1 Then X = picOriginal.ScaleWidth
        If DragR = 0 Then Y = 0
        If DragR = NUM_POINTS + 1 Then Y = picOriginal.ScaleHeight

        ' Move the control point.
        PointX(DragR, DragC) = X
        PointY(DragR, DragC) = Y

        ' Redraw the control grid.
        DrawGrid
    End If
End Sub


' Finish moving a control point.
Private Sub picOriginal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragR = -1
End Sub


