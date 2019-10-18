VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLoadImg 
   Caption         =   "LoadImg []"
   ClientHeight    =   3915
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   2760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtScale 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "0.5"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picStretched 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   840
      ScaleHeight     =   1815
      ScaleWidth      =   1695
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picAntiAliased 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   3360
      ScaleHeight     =   2400
      ScaleWidth      =   2400
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Scale"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgOriginal 
      Height          =   600
      Left            =   120
      Top             =   480
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmLoadImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const LR_LOADFROMFILE = &H10&

Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2

Private FileName As String
' Arrange the controls.
Private Sub ArrangeControls(ByVal picture_scale As Single)
Dim wid As Single
Dim hgt As Single

    ' Position the image controls.
    picStretched.Move _
        imgOriginal.Left + imgOriginal.Width + 120, _
        imgOriginal.Top, _
        imgOriginal.Width * picture_scale, _
        imgOriginal.Height * picture_scale
    picAntiAliased.Move _
        picStretched.Left + picStretched.Width + 120, _
        imgOriginal.Top, _
        imgOriginal.Width * picture_scale, _
        imgOriginal.Height * picture_scale

    ' Make the form big enough.
    wid = picAntiAliased.Left + picAntiAliased.Width
    Width = wid + Width - ScaleWidth + 120

    hgt = picAntiAliased.Top + picAntiAliased.Height
    If hgt < imgOriginal.Top + imgOriginal.Height _
        Then hgt = imgOriginal.Top + imgOriginal.Height
    Height = hgt + Height - ScaleHeight + 120
End Sub

' Load a picture into a PictureBox using LoadImage.
Private Sub LoadImageFile(ByVal pic As PictureBox, ByVal file_name As String)
Dim wid As Long
Dim hgt As Long
Dim hbmp As Long
Dim image_hdc As Long

    ' Get the PictureBox's dimensions in pixels.
    wid = pic.ScaleX(pic.ScaleWidth, pic.ScaleMode, vbPixels)
    hgt = pic.ScaleY(pic.ScaleHeight, pic.ScaleMode, vbPixels)

    ' Load the bitmap.
    hbmp = LoadImage(0, file_name, IMAGE_BITMAP, _
        wid, hgt, LR_LOADFROMFILE)

    ' Make the picture box display the image.
    SelectObject pic.hdc, hbmp

    ' Destroy the bitmap to free its resources.
    DeleteObject hbmp

    ' Refresh the image.
    pic.Refresh
End Sub


' Display the images.
Private Sub DisplayImages(ByVal file_name As String)
Dim picture_scale As Single

    ' Do nothing if no picture is loaded.
    If Len(FileName) = 0 Then Exit Sub

    ' Get the scale.
    On Error Resume Next
    picture_scale = CSng(txtScale.Text)
    If Err.Number <> 0 Then picture_scale = 1
    On Error GoTo LoadError

    ' Load the file at normal scale using LoadPicture.
    imgOriginal.Picture = LoadPicture(file_name)

    ' Arrange the controls.
    ArrangeControls picture_scale

    ' Stretch the image using PaintPicture.
    picStretched.Cls
    picStretched.PaintPicture imgOriginal.Picture, _
        0, 0, picStretched.Width, picStretched.Height

    ' Load the file using LoadImage.
    picAntiAliased.Cls
    LoadImageFile picAntiAliased, file_name

    imgOriginal.Visible = True
    picStretched.Visible = True
    picAntiAliased.Visible = True
    Exit Sub

LoadError:
    MsgBox "Error " & Format$(Err.Number) & vbCrLf & _
        " opening file '" & FileName & "'" & vbCrLf & _
        Err.Description
End Sub

' Redisplay the images.
Private Sub cmdRefresh_Click()
    DisplayImages FileName
End Sub

' Start in the current directory.
Private Sub Form_Load()
    dlgOpenFile.CancelError = True
    dlgOpenFile.InitDir = App.Path
    dlgOpenFile.Filter = _
        "Bitmaps (*.bmp)|*.bmp|" & _
        "Icons (*.ico)|*.ico|" & _
        "Cursors (*.cur)|*.cur|" & _
        "Graphic Files|*.bmp;*.ico;*.cur|" & _
        "All Files (*.*)|*.*"
End Sub
' Load the indicated file.
Private Sub mnuFileOpen_Click()
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

    FileName = Trim$(dlgOpenFile.FileName)
    dlgOpenFile.InitDir = Left$(FileName, Len(FileName) _
        - Len(dlgOpenFile.FileTitle) - 1)
    Caption = "LoadImg [" & dlgOpenFile.FileTitle & "]"

    ' Display the images.
    DisplayImages FileName
End Sub
