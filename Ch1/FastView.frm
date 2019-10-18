VERSION 5.00
Begin VB.Form frmFastView 
   Caption         =   "FastView"
   ClientHeight    =   5685
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   8715
   Begin VB.ComboBox cboPattern 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.DirListBox dirDirectories 
      Height          =   1155
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.FileListBox filFiles 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image imgView 
      Height          =   615
      Left            =   2280
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmFastView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Make the FileList look in this directory.
Private Sub dirDirectories_Change()
    filFiles.Path = dirDirectories.Path
End Sub
' Make the DirectoryList look at this drive.
Private Sub drvDrives_Change()
    'On Error GoTo DriveError
    dirDirectories.Path = drvDrives.Drive
    Exit Sub

DriveError:
    drvDrives.Drive = dirDirectories.Path
    Exit Sub
End Sub


' Display the file.
Private Sub filFiles_Click()
Dim fname As String

    On Error GoTo LoadPictureError

    fname = filFiles.Path & "\" & filFiles.FileName
    Caption = "FastView [" & fname & "]"

    MousePointer = vbHourglass
    DoEvents
    imgView.Picture = LoadPicture(fname)
    MousePointer = vbDefault

    Exit Sub

LoadPictureError:
    Beep
    MousePointer = vbDefault
    Caption = "Viewer [Invalid picture]"
    Set imgView.Picture = Nothing
    Exit Sub
End Sub

' Initialize the file patterns.
Private Sub Form_Load()
    cboPattern.AddItem "Bitmaps (*.bmp)"
    cboPattern.AddItem "GIF (*.gif)"
    cboPattern.AddItem "JPEG (*.jpg)"
    cboPattern.AddItem "Icons (*.ico)"
    cboPattern.AddItem "Matafiles (*.wmf)"
    cboPattern.AddItem "DIBs (*.dib)"
    cboPattern.AddItem "Graphic (*.gif;*.jpg;*.ico;*.bmp;*.wmf;*.dib)"
    cboPattern.AddItem "All Files (*.*)"

    ' Select all graphic files.
    cboPattern.ListIndex = 6
End Sub

' Make the controls as large as possible.
Private Sub Form_Resize()
Dim wid As Integer
Dim hgt As Integer

    If WindowState = vbMinimized Then Exit Sub

    wid = drvDrives.Width
    drvDrives.Move 0, 0, wid

    hgt = ScaleHeight - cboPattern.Height
    If hgt < 120 Then hgt = 120
    cboPattern.Move 0, hgt, wid

    hgt = (cboPattern.Top - drvDrives.Top - drvDrives.Height) / 2
    If hgt < 120 Then hgt = 120
    dirDirectories.Move 0, drvDrives.Top + drvDrives.Height + 0, wid, hgt
    filFiles.Move 0, dirDirectories.Top + dirDirectories.Height + 0, wid, hgt

    wid = ScaleWidth - drvDrives.Width
    If wid < 120 Then wid = 120
    imgView.Move drvDrives.Width, 0, wid, ScaleHeight
End Sub





' Set the filFiles control's Pattern property to
' the selected pattern.
Private Sub cboPattern_Click()
Dim pat As String
Dim p1 As Integer
Dim p2 As Integer

    pat = cboPattern.List(cboPattern.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    filFiles.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub
