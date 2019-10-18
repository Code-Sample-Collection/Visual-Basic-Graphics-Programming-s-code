VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmResize2 
   Caption         =   "Resize2 []"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   ScaleHeight     =   2895
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResult 
      Height          =   2295
      Left            =   840
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdResize 
      Caption         =   "Resize"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "1.0"
      Top             =   60
      Width           =   495
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
   Begin VB.Label Label1 
      Caption         =   "Scale"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   60
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
Attribute VB_Name = "frmResize2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FileName As String

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const LR_LOADFROMFILE = &H10
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
' Arrange the controls.
Private Sub ArrangeControls(ByVal scale_factor As Single)
Dim new_wid As Single
Dim new_hgt As Single

    ' Calculate the result's size.
    new_wid = picOriginal.ScaleWidth * scale_factor
    new_hgt = picOriginal.ScaleHeight * scale_factor
    new_wid = ScaleX(new_wid, vbPixels, ScaleMode) + picOriginal.Width - ScaleX(picOriginal.ScaleWidth, vbPixels, ScaleMode)
    new_hgt = ScaleY(new_hgt, vbPixels, ScaleMode) + picOriginal.Height - ScaleY(picOriginal.ScaleHeight, vbPixels, ScaleMode)

    ' Position the result PictureBox.
    picResult.Move _
        picOriginal.Left + picOriginal.Width + 120, _
        picOriginal.Top, new_wid, new_hgt
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    picResult.Picture = picResult.Image
    picResult.Visible = True

    ' This makes the image resize itself to
    ' fit the picture.
    picResult.Picture = picResult.Image

    ' Make the form big enough.
    new_wid = picResult.Left + picResult.Width
    If new_wid < cmdResize.Left + cmdResize.Width _
        Then new_wid = cmdResize.Left + cmdResize.Width
    new_hgt = picResult.Top + picResult.Height
    If new_hgt < picOriginal.Top + picOriginal.Height _
        Then new_hgt = picOriginal.Top + picOriginal.Height
    Move Left, Top, new_wid + 237, new_hgt + 816

    DoEvents
End Sub

' Use the LoadImage API function to load a picture
' from a file into a PictureBox, filling the PictureBox.
Private Sub LoadImageFromFile(ByVal file_name As String, ByVal pic As PictureBox)
Dim wid As Long
Dim hgt As Long
Dim mem_dc As Long
Dim hbmp As Long

    ' Get the desired size in pixels.
    wid = pic.ScaleX(pic.ScaleWidth, pic.ScaleMode, vbPixels)
    hgt = pic.ScaleY(pic.ScaleHeight, pic.ScaleMode, vbPixels)

    ' Get the bitmap handle from the file.
    hbmp = LoadImage(ByVal 0&, file_name, 0, _
        wid, hgt, LR_LOADFROMFILE)

    ' Create a device context to hold the image.
    mem_dc = CreateCompatibleDC(0)

    ' Select the bitmap into the device context.
    SelectObject mem_dc, hbmp

    ' Copy the bitmap into picResult.
    BitBlt pic.hdc, 0, 0, wid, hgt, _
        mem_dc, 0, 0, SRCCOPY
    pic.Refresh

    ' Delete the device context and bitmap.
    DeleteDC mem_dc
    DeleteObject hbmp
End Sub

' Transform the picture.
Private Sub cmdResize_Click()
Dim scale_factor As Single

    ' Do nothing if no picture is loaded.
    If picOriginal.Picture = 0 Then Exit Sub

    ' Get the scale.
    On Error GoTo ScaleError
    scale_factor = CSng(txtScale.Text)
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
    picResult.Line (0, 0)-(picResult.ScaleWidth, picResult.ScaleHeight), _
        picResult.BackColor, BF
    DoEvents

    ' Arrange picResult.
    ArrangeControls scale_factor

    ' Reload the picture using LoadImage.
    LoadImageFromFile FileName, picResult

    Screen.MousePointer = vbDefault
    Exit Sub

ScaleError:
    MsgBox "Invalid scale"
    txtScale.SetFocus
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

    Width = picResult.Left + picResult.Width + 120 + Width - ScaleWidth
    Height = picOriginal.Top + picOriginal.Height + 120 + Height - ScaleHeight
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
    Caption = "Resize [" & dlgOpenFile.FileTitle & "]"

    ' Save the file name for use with LoadImage.
    FileName = file_name

    ' Open the original file.
    On Error GoTo LoadError
    picOriginal.Picture = LoadPicture(file_name)
    On Error GoTo 0

    ' Hide picResult.
    picResult.Visible = False
    If cmdResize.Left + cmdResize.Width > picOriginal.Left + picOriginal.Width Then
        Width = cmdResize.Left + cmdResize.Width + 120 + Width - ScaleWidth
    Else
        Width = picOriginal.Left + picOriginal.Width + 120 + Width - ScaleWidth
    End If
    Height = picOriginal.Top + picOriginal.Height + 120 + Height - ScaleHeight

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
    Caption = "Resize [" & dlgOpenFile.FileTitle & "]"

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

