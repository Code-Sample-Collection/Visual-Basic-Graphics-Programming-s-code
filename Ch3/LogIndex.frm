VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogIndex 
   AutoRedraw      =   -1  'True
   Caption         =   "LogIndex"
   ClientHeight    =   5085
   ClientLeft      =   1395
   ClientTop       =   1005
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   Palette         =   "LogIndex.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   Begin VB.PictureBox picBitmap 
      AutoRedraw      =   -1  'True
      Height          =   5055
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   4320
      TabIndex        =   0
      Top             =   0
      Width           =   4380
   End
   Begin VB.PictureBox picPalette 
      AutoRedraw      =   -1  'True
      Height          =   2460
      Left            =   4440
      ScaleHeight     =   2400
      ScaleWidth      =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2460
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   4440
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmLogIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PALETTE_INDEX = &H1000000
' Make the controls as large as possible.
Private Sub Form_Resize()
Dim wid As Single

    picPalette.Left = ScaleWidth - picPalette.Width

    wid = picPalette.Left - 3
    If wid < 10 Then wid = 10

    picBitmap.Move picBitmap.Left, picBitmap.Top, wid, ScaleHeight
End Sub

' Load an image.
Private Sub mnuFileOpen_Click()
Dim fname As String

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgOpenFile.FileName = "*.BMP;*.WMF;*.DIB;*.JPG;*.GIF"
    dlgOpenFile.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    dlgOpenFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo LoadError
    
    fname = Trim$(dlgOpenFile.FileName)
    dlgOpenFile.InitDir = Left$(fname, Len(fname) _
        - Len(dlgOpenFile.FileTitle) - 1)
    Caption = "ShowLog [" & fname & "]"
    
    ' Load the picture.
    picBitmap.Picture = LoadPicture(fname)
    picBitmap.Refresh

    ' Make picPalette use the same logical palette.
    picPalette.Picture = picPalette.Image
    picPalette.Picture.hPal = picBitmap.Picture.hPal

    ' Display the logical palette colors.
    FillPicture
    Exit Sub
    
LoadError:
    Beep
    MsgBox "Error loading picture " & fname & _
        "." & vbCrLf & Error$, vbExclamation
End Sub

' Fill picture box Pal with its logical palette
' colors using palette indexes.
Private Sub FillPicture()
Dim i As Integer
Dim j As Integer
Dim dx As Single
Dim dy As Single
Dim clr As Long

    dx = picPalette.ScaleWidth / 16
    dy = picPalette.ScaleHeight / 16
    clr = 0
    For i = 0 To 16
        For j = 0 To 16
            picPalette.Line (j * dx, i * dy)-Step(dx, dy), _
                clr + PALETTE_INDEX, BF
            clr = clr + 1
        Next j
    Next i
End Sub

Private Sub Form_Load()
    ' Start the file selection dialog in the
    ' current directory.
    dlgOpenFile.InitDir = App.Path

    ' Fill in the initial palette.
    FillPicture
End Sub


