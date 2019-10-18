VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGetEntry 
   Caption         =   "GetEntry"
   ClientHeight    =   3495
   ClientLeft      =   1500
   ClientTop       =   1260
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   5910
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   3240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtColors 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   3495
      Left            =   0
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmGetEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long

Private Const RASTERCAPS = 38    ' Raster device capabilities.
Private Const RC_PALETTE = &H100 ' Has palettes.

' Display a list of the colors in the logical palette.
Private Sub ShowEntries()
Dim num_entries As Integer
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim txt As String

    If picCanvas.Picture = 0 Then
        txtColors.Text = "No picture loaded."
        Exit Sub
    ElseIf picCanvas.Picture.hPal = 0 Then
        txtColors.Text = "Default palette."
        Exit Sub
    End If
    
    num_entries = GetPaletteEntries(picCanvas.Picture.hPal, 0, 256, palentry(0))
    
    txt = "  #  Red Green Blue" & vbCrLf
    For i = 0 To num_entries - 1
        txt = txt & _
            Format$(i, "@@@") & ":" & _
            Format$(palentry(i).peRed, "@@@@") & _
            Format$(palentry(i).peGreen, "@@@@@@") & _
            Format$(palentry(i).peBlue, "@@@@@") & _
            vbCrLf
    Next i

    txtColors.Text = txt
End Sub
Private Sub Form_Load()
    ' Make sure the screen supports palettes.
    If Not GetDeviceCaps(hdc, RASTERCAPS) And RC_PALETTE Then
        Beep
        MsgBox "This monitor does not support palettes.", _
            vbCritical
        End
    End If

    ' Start in the current directory.
    dlgOpenFile.InitDir = App.Path

    ShowEntries
End Sub



Private Sub Form_Resize()
Dim wid As Single

    txtColors.Move ScaleWidth - txtColors.Width, _
        0, txtColors.Width, ScaleHeight

    wid = txtColors.Left - 20
    If wid < 100 Then wid = 100
    picCanvas.Move 0, 0, wid, ScaleHeight
End Sub

' Load a picture.
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
    On Error GoTo 0
    
    MousePointer = vbHourglass
    DoEvents
    
    fname = Trim$(dlgOpenFile.FileName)
    dlgOpenFile.InitDir = Left$(fname, Len(fname) _
        - Len(dlgOpenFile.FileTitle) - 1)

    ' Load the picture.
    picCanvas.Picture = LoadPicture(fname)

    Caption = "GetEntry [" & fname & "]"

    ' Update the list of colors.
    ShowEntries
    MousePointer = vbDefault
End Sub
