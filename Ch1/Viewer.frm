VERSION 5.00
Begin VB.Form frmViewer 
   Caption         =   "Viewer []"
   ClientHeight    =   5685
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   8715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   8715
   Begin VB.VScrollBar vbar 
      Height          =   1335
      Left            =   3960
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar hbar 
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.PictureBox picScroller 
      Height          =   1215
      Left            =   2280
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   0
      Width           =   1575
      Begin VB.PictureBox picImage 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1335
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.ComboBox cboPatterns 
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
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' The user selected a directory. Tell the FileList
' control so it can list the files in it.
Private Sub dirDirectories_Change()
    filFiles.Path = dirDirectories.Path
End Sub

' The user selected a drive. Tell the DirectoryList
' control so it can list the files in it.
Private Sub drvDrives_Change()
    On Error GoTo DriveError
    dirDirectories.Path = drvDrives.Drive
    Exit Sub

DriveError:
    drvDrives.Drive = dirDirectories.Path
    Exit Sub
End Sub


' The user selected a file. Display it if possible.
Private Sub filFiles_Click()
Dim fname As String

    On Error GoTo LoadPictureError

    fname = filFiles.Path & "\" & filFiles.FileName
    Caption = "Viewer [" & fname & "]"
    
    MousePointer = vbHourglass
    DoEvents
    picImage.Picture = LoadPicture(fname)
    ArrangeScrollbars
    MousePointer = vbDefault

    Exit Sub

LoadPictureError:
    Beep
    MousePointer = vbDefault
    Caption = "Viewer [Invalid picture]"
    Exit Sub
End Sub

' Create the list of file patterns.
Private Sub Form_Load()
    dirDirectories.Path = App.Path

    cboPatterns.AddItem "Bitmaps (*.bmp)"
    cboPatterns.AddItem "GIFs (*.gif)"
    cboPatterns.AddItem "JPEGs (*.jpg)"
    cboPatterns.AddItem "Icons (*.ico)"
    cboPatterns.AddItem "Cursors (*.cur)"
    cboPatterns.AddItem "Run-Length Encoded (*.rle)"
    cboPatterns.AddItem "Metafiles (*.wmf)"
    cboPatterns.AddItem "Enhanced Metafiles (*.emf)"
    cboPatterns.AddItem "Graphic Files (*.bmp;*.gif;*.jpg;*.jpeg;*.ico;*.cur;*.rle;*.wmf;*.emf)"
    cboPatterns.AddItem "All Files (*.*)"

    cboPatterns.ListIndex = 0
End Sub

' Arrange the scrolling controls.
Private Sub ArrangeScrollbars()
Dim need_hgt As Single
Dim need_wid As Single
Dim got_hgt As Single
Dim got_wid As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    ' See which scroll bars we need.
    need_wid = picImage.Width
    need_hgt = picImage.Height
    got_wid = ScaleWidth - picScroller.Left
    got_hgt = ScaleHeight - picScroller.Top

    ' See if we need the horizontal scroll bar.
    need_hbar = (got_wid < need_wid)
    If need_hbar Then
        got_hgt = got_hgt - hbar.Height
    End If

    ' See if we need the vertical scroll bar.
    need_vbar = (got_hgt < need_hgt)
    If need_vbar Then
        got_wid = got_wid - vbar.Width

        ' See if we did not need the horizontal
        ' scroll bar but we now do.
        If Not need_hbar Then
            need_hbar = (got_wid < need_wid)
            If need_hbar Then
                got_hgt = got_hgt - hbar.Height
            End If
        End If
    End If
    If got_hgt < 120 Then got_hgt = 120
    If got_wid < 120 Then got_wid = 120

    ' Display the needed scroll bars.
    If need_hbar Then
        hbar.Move picScroller.Left, got_hgt, got_wid
        hbar.Min = 0
        hbar.Max = got_wid - need_wid
        hbar.SmallChange = got_wid / 3
        hbar.LargeChange = _
            (hbar.Max - hbar.Min) _
                * need_wid / _
                (got_wid - need_wid)
        hbar.Visible = True
    Else
        hbar.Value = 0
        hbar.Visible = False
    End If
    If need_vbar Then
        vbar.Move picScroller.Left + got_wid, 0, vbar.Width, got_hgt
        vbar.Min = 0
        vbar.Max = got_hgt - need_hgt
        vbar.SmallChange = got_hgt / 3
        vbar.LargeChange = _
            (vbar.Max - vbar.Min) _
                * need_hgt / _
                (got_hgt - need_hgt)
        vbar.Visible = True
    Else
        vbar.Value = 0
        vbar.Visible = False
    End If

    ' Arrange the window.
    picScroller.Move picScroller.Left, 0, got_wid, got_hgt
End Sub
' Make the controls fill the form.
Private Sub Form_Resize()
Const GAP = 60

Dim wid As Integer
Dim hgt As Integer

    If WindowState = vbMinimized Then Exit Sub

    wid = drvDrives.Width
    drvDrives.Move GAP, GAP, wid
    cboPatterns.Move GAP, ScaleHeight - cboPatterns.Height, wid
    
    hgt = (cboPatterns.Top - drvDrives.Top - drvDrives.Height - 3 * GAP) / 2
    If hgt < 100 Then hgt = 100
    dirDirectories.Move GAP, drvDrives.Top + drvDrives.Height + GAP, wid, hgt
    filFiles.Move GAP, dirDirectories.Top + dirDirectories.Height + GAP, wid, hgt

    ArrangeScrollbars
End Sub

Private Sub hbar_Change()
    picImage.Left = hbar.Value
End Sub

Private Sub hbar_Scroll()
    picImage.Left = hbar.Value
End Sub


' The user has selected a file pattern. Apply it
' to the FileList box.
Private Sub cboPatterns_Click()
Dim pat As String
Dim p1 As Integer
Dim p2 As Integer

    pat = cboPatterns.List(cboPatterns.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    filFiles.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub


Private Sub vbar_Change()
    picImage.Top = vbar.Value
End Sub


Private Sub vbar_Scroll()
    picImage.Top = vbar.Value
End Sub
