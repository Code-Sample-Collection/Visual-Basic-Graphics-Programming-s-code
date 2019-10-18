VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PalDrawForm 
   Caption         =   "PalDraw"
   ClientHeight    =   4260
   ClientLeft      =   1455
   ClientTop       =   1440
   ClientWidth     =   7200
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4260
   ScaleWidth      =   7200
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4238
      Left            =   2700
      ScaleHeight     =   4185
      ScaleWidth      =   4440
      TabIndex        =   0
      Top             =   0
      Width           =   4500
   End
   Begin VB.PictureBox picForeColor 
      AutoRedraw      =   -1  'True
      Height          =   500
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   1440
      Width           =   500
   End
   Begin VB.PictureBox picFillColor 
      AutoRedraw      =   -1  'True
      Height          =   500
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   2040
      Width           =   500
   End
   Begin VB.ComboBox cboFill 
      Height          =   315
      ItemData        =   "PalDraw.frx":0000
      Left            =   840
      List            =   "PalDraw.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ComboBox cboDraw 
      Height          =   315
      ItemData        =   "PalDraw.frx":008F
      Left            =   840
      List            =   "PalDraw.frx":00A8
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox cboObject 
      Height          =   315
      ItemData        =   "PalDraw.frx":00E7
      Left            =   840
      List            =   "PalDraw.frx":00F7
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   840
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   360
      Width           =   375
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "FillColor"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ForeColor"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "FillStyle"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "DrawStyle"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "DrawWidth"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Object"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "PalDrawForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLOSEST_IN_PALETTE = &H2000000

' The object types.
Private Enum ObjectTypes
    objLine
    objBox
    objEllipse
    objPoint
End Enum

' The type of object we should draw.
Private SelectedObject As ObjectTypes

Private Rubberbanding As Boolean
Private OldStyle As Integer
Private FirstX As Single
Private FirstY As Single
Private LastX As Single
Private LastY As Single

Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long

' Draw the object.
Private Sub DrawObject()
    ' Draw the object.
    Select Case SelectedObject
        Case objLine
            picCanvas.Line (FirstX, FirstY)-(LastX, LastY)
        Case objBox
            picCanvas.Line (FirstX, FirstY)-(LastX, LastY), , B
        Case objEllipse
            DrawEllipse FirstX, FirstY, LastX, LastY
        Case objPoint
            picCanvas.PSet (LastX, LastY)
    End Select
End Sub

' Draw an ellipse specified by a bounding box.
Private Sub DrawEllipse( _
    ByVal xmin As Single, ByVal ymin As Single, _
    ByVal xmax As Single, ByVal ymax As Single)
Dim cx As Single
Dim cy As Single
Dim wid As Single
Dim hgt As Single
Dim aspect As Single
Dim radius As Single

    ' Find the center.
    cx = (xmin + xmax) / 2
    cy = (ymin + ymax) / 2

    ' Get the ellipse's size.
    wid = xmax - xmin
    hgt = ymax - ymin
    If wid = 0 Or hgt = 0 Then Exit Sub
    aspect = hgt / wid

    ' See which dimension is larger and
    ' calculate the radius.
    If wid > hgt Then
        ' The major axis is horizontal.
        ' Use a horizontal radius.
        radius = wid / 2
    Else
        ' The major axis is vertical.
        ' Use a vertical radius.
        radius = aspect * wid / 2
    End If

    ' Draw the circle.
    picCanvas.Circle (cx, cy), radius, , , , aspect
End Sub

' Erase the image using the current fill color.
Private Sub cmdClear_Click()
    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), picCanvas.FillColor, BF
End Sub

' Start a rubberbanding of some sort.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Let MouseMove know we are rubberbanding.
    Rubberbanding = True

    ' Save values so we can restore them later.
    OldStyle = picCanvas.DrawStyle
    picCanvas.DrawMode = vbInvert
    If SelectedObject = objLine Then
        picCanvas.DrawStyle = vbSolid
    Else
        picCanvas.DrawStyle = vbDot
    End If

    ' Save the starting coordinates.
    FirstX = X
    FirstY = Y

    ' Save the ending coordinates.
    LastX = X
    LastY = Y

    ' Draw the appropriate rubberband object.
    DrawObject
End Sub


' Continue rubberbanding.
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub
    
    ' Erase the previous rubberband object.
    DrawObject

    ' Save the new ending coordinates.
    LastX = X
    LastY = Y
    
    ' Draw the new rubberband object.
    DrawObject
End Sub

' Finish rubberbanding and draw the object.
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' We are no longer rubberbanding.
    Rubberbanding = False

    ' Erase the previous rubberband object.
    DrawObject

    ' Restore the original DrawMode and DrawStyle.
    picCanvas.DrawMode = vbCopyPen
    picCanvas.DrawStyle = OldStyle

    ' Draw the final object.
    DrawObject
End Sub



' Select the draw style.
Private Sub cboDraw_Click()
    picCanvas.DrawStyle = cboDraw.ListIndex
End Sub

' Select the fill style.
Private Sub cboFill_Click()
    picCanvas.FillStyle = cboFill.ListIndex
End Sub


' Allow the user to select a new foreground color.
Private Sub picForeColor_Click()
Dim popup As New PalettePopup
Dim clr As Long

    ' Load the picture to get its palette.
    popup.Picture = picCanvas.Picture
    
    ' Fill the popup with palette colors.
    popup.Fill
        
    ' Select the current foreground color.
    popup.SelectedColor = picCanvas.ForeColor
    
    ' Let the user select a color.
    popup.Show vbModal
    
    ' Set the selected color using the palete
    ' relative RGB value.
    clr = popup.SelectedColor + CLOSEST_IN_PALETTE
    picCanvas.ForeColor = clr
    picForeColor.Line (0, 0)-(picForeColor.ScaleWidth, picForeColor.ScaleHeight), clr, BF

    Unload popup
End Sub
' Allow the user to select a new fill color.
Private Sub picFillColor_Click()
Dim popup As New PalettePopup
Dim clr As Long

    ' Load the picture to get its palette.
    popup.Picture = picCanvas.Picture
    
    ' Fill the popup with palette colors.
    popup.Fill
        
    ' Select the current background color.
    popup.SelectedColor = picCanvas.FillColor
    
    ' Let the user select a color.
    popup.Show vbModal
    
    ' Set the selected color using the palete
    ' relative RGB value.
    clr = popup.SelectedColor + CLOSEST_IN_PALETTE
    picCanvas.FillColor = clr
    picFillColor.Line (0, 0)-(picFillColor.ScaleWidth, picFillColor.ScaleHeight), clr, BF

    Unload popup
End Sub


Private Sub Form_Load()
    ' Select the default options.
    cboDraw.ListIndex = picCanvas.DrawStyle
    cboFill.ListIndex = picCanvas.FillStyle
    cboObject.ListIndex = picCanvas.FillStyle
    txtWidth.Text = Format$(picCanvas.DrawWidth)

    ' Fill the color swatches.
    ResetSwatches
End Sub
' Set the colors in the swatches.
Private Sub ResetSwatches()
Dim clr As Long

    picCanvas.Refresh

    ' Make the swatches use the same logical
    ' palette as the picCanvas.
    picForeColor.Picture = picCanvas.Picture
    picFillColor.Picture = picCanvas.Picture

    ' Start with black again.
    picCanvas.ForeColor = vbBlack
    picCanvas.FillColor = vbBlack
    picForeColor.Line (0, 0)-(picForeColor.ScaleWidth, picForeColor.ScaleHeight), vbBlack, BF
    picFillColor.Line (0, 0)-(picFillColor.ScaleWidth, picFillColor.ScaleHeight), vbBlack, BF
End Sub

' Make the controls as larger as possible.
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - cboObject.Left - cboObject.Width - 30
    If wid < 100 Then wid = 100

    picCanvas.Move ScaleWidth - wid, 0, wid, ScaleHeight
End Sub

Private Sub mnuFileOpen_Click()
Dim fname As String

    ' Allow the user to pick a file.
    On Error Resume Next
    FileDialog.FileName = "*.BMP;*.ICO;*.DIB;*.JPG;*.GIF"
    FileDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    FileDialog.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo LoadError
    
    fname = Trim$(FileDialog.FileName)
    FileDialog.InitDir = Left$(fname, Len(fname) _
        - Len(FileDialog.FileTitle) - 1)
    Caption = "PalDraw [" & fname & "]"
    
    ' Load the picture.
    picCanvas.Picture = LoadPicture(fname)
    RealizePalette picCanvas.hdc
    ResetSwatches
    Exit Sub
    
LoadError:
    Beep
    MsgBox "Error loading picture " & fname & _
        "." & vbCrLf & Error$, vbExclamation
End Sub

Private Sub cboObject_Click()
    SelectedObject = cboObject.ListIndex
End Sub


' Change set DrawWidth.
Private Sub txtWidth_Change()
Dim wid As Integer

    If Not IsNumeric(txtWidth.Text) Then Exit Sub
    
    wid = CInt(txtWidth.Text)
    If wid < 1 Then Exit Sub
    
    picCanvas.DrawWidth = wid
End Sub

' Only allow 1 through 9.
Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc(" ") Or _
       KeyAscii > Asc("~") Then Exit Sub
    If KeyAscii >= Asc("1") And _
       KeyAscii <= Asc("9") Then Exit Sub
    Beep
    KeyAscii = 0
End Sub
