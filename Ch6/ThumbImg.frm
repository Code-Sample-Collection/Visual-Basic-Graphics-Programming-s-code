VERSION 5.00
Begin VB.Form frmThumbImg 
   Caption         =   "ThumbImg"
   ClientHeight    =   5685
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   581
   Begin VB.FileListBox filFiles 
      Height          =   1065
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ComboBox cboPatterns 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Text            =   "PatternCombo"
      Top             =   3240
      Width           =   2175
   End
   Begin VB.PictureBox picHidden 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   4200
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picThumb 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1560
      Index           =   0
      Left            =   2235
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.DirListBox dirDirectories 
      Height          =   1155
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblThumb 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2235
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuThumbs 
      Caption         =   "&Thumbs"
      Begin VB.Menu mnuThumbsShow 
         Caption         =   "&Show"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuThumbsSize 
         Caption         =   "S&ize"
         Begin VB.Menu mnuThumbsSetSize 
            Caption         =   "&Small"
            Index           =   50
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuThumbsSetSize 
            Caption         =   "&Medium"
            Index           =   100
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuThumbsSetSize 
            Caption         =   "&Large"
            Index           =   200
            Shortcut        =   ^L
         End
      End
   End
End
Attribute VB_Name = "frmThumbImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Running As Boolean
Private DirName As String
Private MaxFileNum As Integer
Private SelectedThumb As Integer
Private ThumbSize As Single

' API stuff for moving files to the wastebasket.
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10

' API stuff for LoadImage.
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const LR_LOADFROMFILE = &H10&

Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2

' Load a bitmap file into a PictureBox using
' LoadImage.
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

' Move the file into the wastebasket.
Private Sub DeleteFile(ByVal Index As Integer)
Dim op As SHFILEOPSTRUCT
Dim file_name As String

    file_name = DirName & lblThumb(Index).Caption

    file_name = DirName & lblThumb(Index).Caption
    With op
        .wFunc = FO_DELETE
        .pFrom = file_name
        .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
    End With
    SHFileOperation op

    If Not op.fAnyOperationsAborted Then
        ' Mark the file as deleted.
        lblThumb(Index).Caption = ""
        picThumb(Index).Line (0, 0)- _
            (picThumb(Index).ScaleWidth, _
             picThumb(Index).ScaleHeight)
        picThumb(Index).Line _
            (picThumb(Index).ScaleWidth, 0)- _
            (0, picThumb(Index).ScaleHeight)
    End If
End Sub

' Display thumbnails for this directory.
Private Sub ShowThumbs()
Const GAP = 2

Dim i As Integer
Dim new_name As String
Dim new_ext As String
Dim wid As Single
Dim hgt As Single
Dim thumb_left As Single
Dim thumb_top As Single

    MaxFileNum = 0
    SelectedThumb = -1

    ' Get the directory name.
    DirName = dirDirectories.Path
    If Right$(DirName, 1) <> "\" Then
        DirName = DirName & "\"
    End If

    ' Hide the thumbnail pictures.
    For i = 0 To picThumb.UBound
        picThumb(i).Visible = False
        lblThumb(i).Visible = False
    Next i

    ' See where the first thumb goes.
    thumb_left = drvDrives.Left + drvDrives.Width + GAP
    thumb_top = 0

    ' Get the file names.
    For i = 0 To filFiles.ListCount - 1
        new_name = filFiles.List(i)
        new_ext = LCase$(Right$(new_name, 3))

        ' Load the file.
        On Error Resume Next
        ' Load the picture using LoadPicture.
        picHidden.Picture = LoadPicture(DirName & new_name)

        If Err.Number = 0 Then
            ' We loaded the picture successfully.
            ' Display its thumbnail.
            On Error GoTo 0

            ' Calculate the thumbnail size.
            wid = picHidden.ScaleWidth
            hgt = picHidden.ScaleHeight
            If wid > ThumbSize Then
                hgt = hgt * ThumbSize / wid
                wid = ThumbSize
            End If
            If hgt > ThumbSize Then
                wid = wid * ThumbSize / hgt
                hgt = ThumbSize
            End If

            ' Load the thumbnail picture.
            If MaxFileNum > picThumb.UBound Then
                Load picThumb(MaxFileNum)
                Load lblThumb(MaxFileNum)
            End If

            ' Display the thumbnail.
            picThumb(MaxFileNum).BorderStyle = vbBSNone
            
            ' See if this is a bitmap.
            If (new_ext = "bmp") Then
                ' Load the picture using LoadImage.
                ' Make the thumbnail the right shape.
                picThumb(MaxFileNum).Move _
                    thumb_left + (ThumbSize - wid) / 2, _
                    thumb_top + (ThumbSize - hgt) / 2, _
                    wid, hgt
                picThumb(MaxFileNum).Picture = picThumb(MaxFileNum).Image

                ' Display the image.
                LoadImageFile picThumb(MaxFileNum), DirName & new_name
            Else
                ' Copy the picture using PaintPicture.
                ' Make the thumbnail fill its area.
                picThumb(MaxFileNum).Move _
                    thumb_left, thumb_top, _
                    ThumbSize, ThumbSize

                ' Clear the thumbnail.
                picThumb(MaxFileNum).Line (0, 0)-(picThumb(MaxFileNum).ScaleWidth, picThumb(MaxFileNum).ScaleHeight), vbWhite, BF

                ' Copy the image reduced.
                picThumb(MaxFileNum).PaintPicture _
                    picHidden.Picture, _
                    (ThumbSize - wid) / 2, _
                    (ThumbSize - hgt) / 2, wid, hgt, _
                    0, 0, picHidden.ScaleWidth, picHidden.ScaleHeight
            End If
            picThumb(MaxFileNum).Visible = True

            lblThumb(MaxFileNum).Move _
                thumb_left, thumb_top + ThumbSize, _
                ThumbSize
            lblThumb(MaxFileNum).Caption = new_name
            lblThumb(MaxFileNum).Visible = True

            MaxFileNum = MaxFileNum + 1

            ' See where the next thumb goes.
            thumb_left = thumb_left + ThumbSize + GAP
            If thumb_left + ThumbSize > ScaleWidth Then
                thumb_left = drvDrives.Left + drvDrives.Width + GAP
                thumb_top = thumb_top + ThumbSize + _
                    lblThumb(0).Height + 3 * GAP
                If thumb_top + ThumbSize > ScaleHeight Then Exit For
            End If

            DoEvents
            If Not Running Then Exit Sub
        End If ' End if we got no error loading the picture.
    Next i
End Sub
' The user selected a directory. Let the filFiles
' control know so it can update its list.
Private Sub dirDirectories_Change()
    filFiles.Path = dirDirectories.Path
End Sub

' The user selected a drive. Let the dirDirectories
' control know so it can update its list.
Private Sub drvDrives_Change()
    'On Error GoTo DriveError
    dirDirectories.Path = drvDrives.Drive
    Exit Sub

DriveError:
    drvDrives.Drive = dirDirectories.Path
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

    cboPatterns.ListIndex = 8

    mnuThumbsSetSize_Click 100
End Sub
' Make the controls fill the form.
Private Sub Form_Resize()
Const GAP = 2

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
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Set the thumbnail size.
Private Sub mnuThumbsSetSize_Click(Index As Integer)
    mnuThumbsSetSize(50).Checked = False
    mnuThumbsSetSize(100).Checked = False
    mnuThumbsSetSize(200).Checked = False
    mnuThumbsSetSize(Index).Checked = True

    ThumbSize = Index

    mnuThumbsShow_Click
End Sub
' Start or stop displaying thumbnails.
Private Sub mnuThumbsShow_Click()
    If Running Then
        ' Stop.
        mnuThumbsShow.Enabled = False
        mnuThumbsShow.Caption = "Stopping"
        Running = False
        DoEvents
    Else
        ' Start.
        mnuThumbsShow.Caption = "Stop"
        Running = True
        MousePointer = vbHourglass
        DoEvents

        ShowThumbs

        Running = False
        mnuThumbsShow.Caption = "Show"
        mnuThumbsShow.Enabled = True
        MousePointer = vbDefault
    End If
End Sub
' The user selected a pattern. Let the filFiles
' control know so it can filter its list.
Private Sub cboPatterns_Click()
Dim pat As String
Dim p1 As Integer
Dim p2 As Integer

    pat = cboPatterns.List(cboPatterns.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    filFiles.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub

' The user clicked on a thumbnail. Select it.
Private Sub picThumb_Click(Index As Integer)
    If SelectedThumb >= 0 Then
        picThumb(SelectedThumb).BorderStyle = vbBSNone
    End If

    SelectedThumb = Index
    picThumb(SelectedThumb).BorderStyle = vbFixedSingle

    Caption = "Thumbs - " & lblThumb(SelectedThumb).Caption
End Sub


' The user pressed a key while a thumbnail had
' the focus. If it is the delete key, move the
' file into the waste basket.
Private Sub picThumb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDelete) And _
       (Len(lblThumb(Index).Caption) > 0) _
    Then
        ' Move the file into the wastebasket.
        DeleteFile Index
    End If
End Sub
