VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPalEdit 
   Caption         =   "PalEdit"
   ClientHeight    =   5805
   ClientLeft      =   1305
   ClientTop       =   780
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   388
   ScaleMode       =   0  'User
   ScaleWidth      =   468
   Begin VB.PictureBox picVisible 
      AutoRedraw      =   -1  'True
      Height          =   4515
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "PalEdit.frx":0000
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   1
      Top             =   0
      Width           =   4245
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   4200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   8.37851e-39
   End
   Begin VB.PictureBox picSwatch 
      AutoRedraw      =   -1  'True
      Height          =   2280
      Left            =   4560
      Picture         =   "PalEdit.frx":0446
      ScaleHeight     =   2220
      ScaleWidth      =   2400
      TabIndex        =   15
      Top             =   2505
      Width           =   2460
   End
   Begin VB.PictureBox picSystemColors 
      AutoRedraw      =   -1  'True
      Height          =   2460
      Left            =   4560
      Picture         =   "PalEdit.frx":088C
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   14
      Top             =   0
      Width           =   2460
   End
   Begin VB.PictureBox picColors 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   468
      TabIndex        =   4
      Top             =   4830
      Width           =   7020
      Begin VB.HScrollBar hbarBlue 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   16
         Left            =   885
         Max             =   255
         TabIndex        =   7
         Top             =   720
         Width           =   6090
      End
      Begin VB.HScrollBar hbarGreen 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   16
         Left            =   885
         Max             =   255
         TabIndex        =   6
         Top             =   360
         Width           =   6090
      End
      Begin VB.HScrollBar hbarRed 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   16
         Left            =   885
         Max             =   255
         TabIndex        =   5
         Top             =   0
         Width           =   6090
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Red"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Green"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Blue"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.PictureBox picHidden 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   3720
      Picture         =   "PalEdit.frx":0CD2
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar HBar 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4530
      Width           =   4245
   End
   Begin VB.VScrollBar VBar 
      Height          =   4515
      Left            =   4260
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "&Revert"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuScale 
      Caption         =   "&Scale"
      Begin VB.Menu mnuScaleZoomIn 
         Caption         =   "Zoom &In"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuScaleFull 
         Caption         =   "&Full Scale"
      End
      Begin VB.Menu mnuScaleZoomOut 
         Caption         =   "Zoom &Out"
      End
   End
   Begin VB.Menu mnuColor 
      Caption         =   "&Color"
      Begin VB.Menu mnuNear 
         Caption         =   "&Nearest"
         Begin VB.Menu mnuNearRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuNearGreen 
            Caption         =   "&Green"
         End
         Begin VB.Menu mnuNearBlue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu mnuNearGray 
            Caption         =   "Gray"
         End
      End
      Begin VB.Menu mnuGrad 
         Caption         =   "&Gradient"
         Begin VB.Menu mnuGradRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuGradGreen 
            Caption         =   "&Green"
         End
         Begin VB.Menu mnuGradBlue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu mnuGradGray 
            Caption         =   "Gray"
         End
         Begin VB.Menu mnuGradRainbow 
            Caption         =   "Rainbow"
         End
      End
   End
End
Attribute VB_Name = "frmPalEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PALETTE_RELATIVE = &H2000000

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ResizePalette Lib "gdi32" (ByVal hPalette As Long, ByVal nNumEntries As Long) As Long
Private Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const PC_EXPLICIT = &H2      ' Match to system palette index.
Private Const PC_NOCOLLAPSE = &H4    ' Do not match color existing entries.

' GetDeviceCaps constants.
Private Const RASTERCAPS = 38    ' Raster device capabilities.
Private Const RC_PALETTE = &H100 ' Has palettes.
Private Const NUMRESERVED = 106  ' # reserved entries in palette.
Private Const SIZEPALETTE = 104  ' Size of system palette.

Private Const NO_COLOR = -1

Private LogicalPalette As Long
Private SystemPalette As Long

Private SysPalSize As Integer
Private NumStaticColors As Integer
Private StaticColor1 As Integer
Private StaticColor2 As Integer

Private SelectedI As Integer
Private SelectedJ As Integer
Private SelectedColor As Integer
Private SelectedR As Integer
Private SelectedG As Integer
Private SelectedB As Integer

Private Dx As Integer
Private Dy As Integer
Private SWid As Single
Private SHgt As Single
Private IWid As Single
Private IHgt As Single
Private ImageScale As Single

Private SettingColor As Boolean
Private DataChanged As Boolean
Private FileLoaded As String
' If the data has been modified, allow the user
' to save the changes or cancel the operation.
' Return True if:
'
'   - The image data has not been changed since
'       it was loaded.
'   - The user saves the changes.
'   - The user says not to save.
'
' Return False otherwise.
Function DataSafe() As Boolean
    DataSafe = True
    
    ' This is done in a while loop in case the
    ' user starts a save and then cancels.
    Do While DataChanged
        Select Case MsgBox("The data has been modified. Do you want to save the changes?", vbQuestion + vbYesNoCancel, "Data Modified")
            Case vbYes
                If FileLoaded <> "" Then
                    mnuFileSave_Click
                Else
                    mnuFileSaveAs_Click
                End If
                DataSafe = Not DataChanged
            
            Case vbNo
                DataSafe = True
                Exit Do

            Case vbCancel
                DataSafe = False
                Exit Do
        End Select
    Loop
End Function
' Copy the image from picHidden to picVisible at
' the correct scale.
Private Sub DrawImage()
Dim image_wid As Single
Dim image_hgt As Single
Dim hidden_wid As Single
Dim hidden_hgt As Single

    If Not Visible Then Exit Sub

    ' Fill it with white. Cls would redisplay the
    ' Picture which is bad if ImageScale < 1.
    picVisible.Line (0, 0)-(IWid, IHgt), vbWhite, BF

    ' Copy the picture at the correct scale.
    image_wid = picVisible.ScaleWidth
    image_hgt = picVisible.ScaleHeight
    hidden_wid = image_wid / ImageScale
    hidden_hgt = image_hgt / ImageScale
    picVisible.PaintPicture _
        picHidden.Picture, 0, 0, _
        image_wid, image_hgt, _
        HBar.Value, VBar.Value, _
        hidden_wid, hidden_hgt
End Sub

' Load the indicated file and prepare to work
' with its palette.
Private Sub LoadImage(fname As String)
    On Error GoTo LoadFileError
    picHidden.Picture = LoadPicture(fname)
    ImageScale = 1#
    ResetScrollBars

    On Error GoTo LoadPalError
    LoadLogicalPalette

    FileLoaded = fname
    Caption = "PalEdit [" & fname & "]"
    mnuFileSave.Enabled = True
    mnuFileRevert.Enabled = True
    DataChanged = False
    Exit Sub
    
LoadFileError:
    Beep
    MsgBox "Error loading file " & fname & "." & _
        vbCrLf & Error$
    Exit Sub

LoadPalError:
    Beep
    MsgBox "Error loading logical palette." & _
        vbCrLf & Error$
    Exit Sub
End Sub

' Set the Max and LargeChange properties for the
' image scroll bars.
Private Sub ResetScrollBars()
Dim change As Single

    change = picVisible.ScaleWidth / ImageScale
    If picHidden.ScaleWidth <= change Then
        HBar.Value = 0
        HBar.Enabled = False
    Else
        HBar.Max = picHidden.ScaleWidth - change
        HBar.LargeChange = change
        HBar.Enabled = True
    End If
    
    change = picVisible.ScaleHeight / ImageScale
    If picHidden.ScaleHeight <= change Then
        VBar.Value = 0
        VBar.Enabled = False
    Else
        VBar.Max = picHidden.ScaleHeight - change
        VBar.LargeChange = change
        VBar.Enabled = True
    End If
End Sub

' Select the color with the indicated index.
Private Sub SelectColorIndex(ByVal index As Integer)
Dim i As Integer
Dim j As Integer

    i = index \ 16
    j = index Mod 16
    SelectColor i, j
End Sub
' Load the picHidden palette so its entries
' match the system entries.
Private Sub LoadLogicalPalette()
Dim palentry(0 To 255) As PALETTEENTRY
Dim blanked(0 To 255) As PALETTEENTRY
Dim i As Integer

    ' Make picVisible and picSwatch use the same
    ' palette as picHidden.
    picVisible.Picture = picHidden.Picture
    picSwatch.Picture = picHidden.Picture
    LogicalPalette = picHidden.Picture.hPal

    ' Draw the image at the correct scale.
    DrawImage

    ' Make sure picVisible has the foreground palette.
    RealizePalette picVisible.hdc

    ' Give the system a chance to catch up.
    DoEvents

    ' Make the logical palette as big as possible.
    If ResizePalette(LogicalPalette, SysPalSize) = 0 Then
        MsgBox "Error resizing logical palette."
        Exit Sub
    End If

    ' Get the system palette entries.
    GetSystemPaletteEntries picHidden.hdc, 0, SysPalSize, palentry(0)

    ' Blank the non-static colors.
    For i = 0 To StaticColor1
        blanked(i) = palentry(i)
        blanked(i).peFlags = PC_NOCOLLAPSE
    Next i
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With blanked(i)
            .peRed = i
            .peGreen = 0
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
    Next i
    For i = StaticColor2 To 255
        blanked(i) = palentry(i)
        blanked(i).peFlags = PC_NOCOLLAPSE
    Next i
    SetPaletteEntries LogicalPalette, 0, SysPalSize, blanked(0)

    ' Insert the non-static colors.
    For i = StaticColor1 + 1 To StaticColor2 - 1
        palentry(i).peFlags = PC_NOCOLLAPSE
    Next i
    SetPaletteEntries LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)

    ' Realize the new palette values.
    RealizePalette picVisible.hdc

    ' Select the color that was selected before.
    SelectColor SelectedI, SelectedJ
End Sub

' Load the picSystemColors palette with PC_EXPLICIT
' entries so they match the system palette.
Private Sub LoadSystemPalette()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer

    ' Make the logical palette as big as possible.
    SystemPalette = picSystemColors.Picture.hPal
    If ResizePalette(SystemPalette, SysPalSize) = 0 Then
        Beep
        MsgBox "Error resizing system palette.", _
            vbExclamation
        Exit Sub
    End If
    
    ' Flag all palette entries as PC_EXPLICIT.
    ' Set peRed to the system palette indexes.
    For i = 0 To SysPalSize - 1
        palentry(i).peRed = i
        palentry(i).peFlags = PC_EXPLICIT
    Next i
    
    ' Update the palette (ignore return value).
    i = SetPaletteEntries(SystemPalette, 0, SysPalSize, palentry(0))
End Sub


' Fill the system picture with all the palette
' colors, hatching the static colors.
Private Sub ShowpicSystemColors()
Dim i As Integer
Dim j As Integer
Dim clr As Integer
Dim oldfill As Integer
Dim olddraw As Integer

    picSystemColors.Cls
    
    ' Display the colors using palette indexing.
    Dx = picSystemColors.ScaleWidth / 16
    Dy = picSystemColors.ScaleHeight / 16
    clr = 0
    For i = 0 To 15
        For j = 0 To 15
            picSystemColors.Line _
                (j * Dx, i * Dy)-Step(Dx, Dy), _
                clr + &H1000000, BF
            clr = clr + 1
        Next j
    Next i
    
    ' Hatch the static colors.
    oldfill = picSystemColors.FillStyle
    olddraw = picSystemColors.DrawMode
    picSystemColors.FillStyle = vbDownwardDiagonal
    picSystemColors.DrawMode = vbInvisible
    
    picSystemColors.Line (0, 0)-Step((NumStaticColors \ 2) * Dx - 1, Dy - 1), , B
    picSystemColors.Line (16 * Dx, 16 * Dy)-Step(-(NumStaticColors \ 2) * Dx, -Dy), , B
    
    picSystemColors.FillStyle = oldfill
    picSystemColors.DrawMode = olddraw

    ' Highlight color (0, 0).
    SelectedColor = NO_COLOR
    SelectColor 0, 0
End Sub

' Select the color at the indicated position.
Private Sub SelectColor(ByVal i As Integer, ByVal j As Integer)
Const GAP1 = 1
Const GAP2 = 2
Const DRAW_WID = 2

Dim oldmode As Integer
Dim oldwid As Integer

    oldmode = picSystemColors.DrawMode
    oldwid = picSystemColors.DrawWidth
    picSystemColors.DrawMode = vbInvert
    picSystemColors.DrawWidth = DRAW_WID
    
    ' Unhighlight the previously selected color.
    If SelectedColor <> NO_COLOR Then _
        picSystemColors.Line (SelectedJ * Dx + GAP1, SelectedI * Dx + GAP1)-Step(Dx - GAP2, Dx - GAP2), , B
    
    ' Record the new color.
    SelectedI = i
    SelectedJ = j
    SelectedColor = i * 16 + j

    ' Highlight the new color.
    picSystemColors.Line (SelectedJ * Dx + GAP1, SelectedI * Dx + GAP1)-Step(Dx - GAP2, Dx - GAP2), , B
    picSystemColors.DrawMode = oldmode
    picSystemColors.DrawWidth = oldwid

    ' Display the color's components.
    ShowColorValue
End Sub


' Display the selected color's components in the
' colors labels and scroll bars.
Private Sub ShowColorValue()
Dim palentry As PALETTEENTRY
Dim status As Integer

    If SelectedColor = NO_COLOR Then Exit Sub
    
    status = GetSystemPaletteEntries(picSystemColors.hdc, SelectedColor, 1, palentry)
    
    ' Update the labels.
    lblRed.Caption = Format$(palentry.peRed)
    lblGreen.Caption = Format$(palentry.peGreen)
    lblBlue.Caption = Format$(palentry.peBlue)
    
    ' Update the color swatch.
    picSwatch.Line (0, 0)-(SWid, SHgt), RGB(palentry.peRed, palentry.peGreen, palentry.peBlue) + PALETTE_RELATIVE, BF

    ' Update the scroll bars.
    If SelectedColor > StaticColor1 And SelectedColor < StaticColor2 Then
        SettingColor = True
        hbarRed.Value = palentry.peRed
        hbarGreen.Value = palentry.peGreen
        hbarBlue.Value = palentry.peBlue
        SettingColor = False
        hbarRed.Enabled = True
        hbarGreen.Enabled = True
        hbarBlue.Enabled = True
    Else
        hbarRed.Enabled = False
        hbarGreen.Enabled = False
        hbarBlue.Enabled = False
    End If
End Sub


' Update the selected color's value.
Private Sub UpdatePalette()
Dim pe As PALETTEENTRY
Dim i As Integer

    pe.peRed = hbarRed.Value
    pe.peGreen = hbarGreen.Value
    pe.peBlue = hbarBlue.Value
    pe.peFlags = PC_NOCOLLAPSE

    ' Update the hidden picture's palette.
    SetPaletteEntries LogicalPalette, SelectedColor, 1, pe
    RealizePalette picHidden.hdc
    picHidden.Picture = picHidden.Image

'@
'picVisible.Picture = picHidden.Picture
'picSwatch.Picture = picHidden.Picture
'LogicalPalette = picHidden.Picture.hPal
'DrawImage
''@
'Palette = picVisible.Picture
'PaletteMode = vbPaletteModeCustom
    SetPaletteEntries picVisible.Picture.hPal, SelectedColor, 1, pe
    RealizePalette picVisible.hdc
    picVisible.Picture = picVisible.Image

    SetPaletteEntries picSwatch.Picture.hPal, SelectedColor, 1, pe
    RealizePalette picSwatch.hdc
    picSwatch.Picture = picSwatch.Image
'@

    picSwatch.Line (0, 0)-(SWid, SHgt), RGB(pe.peRed, pe.peGreen, pe.peBlue) + PALETTE_RELATIVE, BF

    DataChanged = True
End Sub




' Update the selected color's value.
Private Sub hbarBlue_Change()
    If SettingColor Then Exit Sub
    lblBlue.Caption = Format$(hbarBlue.Value)
    UpdatePalette
End Sub


' Update the selected color's value.
Private Sub hbarBlue_Scroll()
    If SettingColor Then Exit Sub
    lblBlue.Caption = Format$(hbarBlue.Value)
    UpdatePalette
End Sub


' Make the scroll bars as big as possible within picColors.
Private Sub picColors_Resize()
Dim wid As Single

    wid = picColors.ScaleWidth - lblRed.Left - lblRed.Width - 2
    If wid < 10 Then wid = 10
    hbarRed.Width = wid
    hbarGreen.Width = wid
    hbarBlue.Width = wid
End Sub


' 1. Make sure we can handle palettes.
' 2. Find out how big the system palette is and how
' many static colors there are.
' 3. Load and display the system palette.
Private Sub Form_Load()
    ' Make sure the screen supports palettes.
    If Not GetDeviceCaps(hdc, RASTERCAPS) And RC_PALETTE Then
        MsgBox "This system is not using palettes."
        End
    End If

    ' Get system palette size and # static colors.
    SysPalSize = GetDeviceCaps(hdc, SIZEPALETTE)
    NumStaticColors = GetDeviceCaps(hdc, NUMRESERVED)
    StaticColor1 = NumStaticColors \ 2 - 1
    StaticColor2 = SysPalSize - NumStaticColors \ 2

    picHidden.AutoSize = True
    ImageScale = 1#

    ' Load the system palette.
    LoadSystemPalette

    ' Display the system palette.
    ShowpicSystemColors

    ' Load the logical palette.
    LoadLogicalPalette

    ' Start in the current directory.
    dlgOpenFile.InitDir = App.Path
End Sub

' Refuse to unload if there are unsaved changes.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not DataSafe()
End Sub


' Make the picture as large as possible.
Private Sub Form_Resize()
Dim L As Single
Dim T As Single
Dim wid As Single
Dim hgt As Single

    If WindowState = vbMinimized Then Exit Sub
    
    ' Keep system colors in the upper right corner.
    picSystemColors.Move ScaleWidth - picSystemColors.Width
    
    ' Keep color box stretched across the bottom.
    picColors.Move 0, ScaleHeight - picColors.Height, ScaleWidth
    
    ' Put color swatch under system colors.
    hgt = picColors.Top - picSystemColors.Height - 6
    If hgt < 10 Then hgt = 10
    picSwatch.Move picSystemColors.Left, picSystemColors.Height + 3, picSwatch.Width, hgt
    SWid = picSwatch.ScaleWidth - 1
    SHgt = picSwatch.ScaleHeight - 1
    
    ' Place the vertical scroll bar.
    L = picSystemColors.Left - VBar.Width - 3
    hgt = picColors.Top - HBar.Height - 4
    If hgt < 10 Then hgt = 10
    VBar.Move L, 0, VBar.Width, hgt
    
    ' Place the horizontal scroll bar.
    T = picColors.Top - HBar.Height - 3
    wid = picSystemColors.Left - VBar.Width - 4
    If wid < 10 Then wid = 10
    HBar.Move 0, T, wid
        
    ' Place picVisible inside the scroll bars.
    picVisible.Move 0, 0, wid, hgt
    IWid = picVisible.ScaleWidth - 1
    IHgt = picVisible.ScaleHeight - 1

    ' Set the scroll bar limits.
    ResetScrollBars
    
    ' Redraw the image in case we've grown.
    DrawImage
    
    ' Refill picSwatch (it may have grown).
    ShowColorValue
End Sub


' Update the selected color's value.
Private Sub hbarGreen_Change()
    If SettingColor Then Exit Sub
    lblGreen.Caption = Format$(hbarGreen.Value)
    UpdatePalette
End Sub

' Update the selected color's value.
Private Sub hbarGreen_Scroll()
    If SettingColor Then Exit Sub
    lblGreen.Caption = Format$(hbarGreen.Value)
    UpdatePalette
End Sub

' Redraw the image scrolled appropriately.
Private Sub HBar_Change()
    DrawImage
End Sub

' Redraw the image scrolled appropriately.
Private Sub HBar_Scroll()
    DrawImage
End Sub


' Select the color the user clicked on.
Private Sub picVisible_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bm As BITMAP
Dim hbm As Long
Dim status As Long
Dim bytes() As Byte
Dim wid As Long
Dim hgt As Long

    ' Get a handle to the bitmap.
    hbm = picVisible.Image
    
    ' See how big it is.
    status = GetObjectAPI(hbm, Len(bm), bm)
    wid = bm.bmWidthBytes
    hgt = bm.bmHeight
    
    ' If the mouse is out of bounds, bail out.
    If X >= wid Or Y >= hgt Then
        Beep
        Exit Sub
    End If
    
    ' Get the bits.
    ReDim bytes(0 To wid - 1, 0 To hgt - 1)
    status = GetBitmapBits(hbm, wid * hgt, bytes(0, 0))
    
    ' Select the color of this pixel.
    SelectColorIndex bytes(CInt(X), CInt(Y))
End Sub


' Load a new image file.
Private Sub mnuFileOpen_Click()
Dim fname As String

    ' Make sure any changes have been saved.
    If Not DataSafe() Then Exit Sub
    
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
    
    fname = Trim$(dlgOpenFile.FileName)
    dlgOpenFile.InitDir = Left$(fname, Len(fname) _
        - Len(dlgOpenFile.FileTitle) - 1)
    
    ' Load the picture.
    Screen.MousePointer = vbHourglass
    DoEvents

    LoadImage fname

    Screen.MousePointer = vbDefault
End Sub

' Reload the file.
Private Sub mnuFileRevert_Click()
    ' If the data has changed, get confirmation.
    If DataChanged Then
        If MsgBox("The data has been modified. Are you sure you want to remove the changes?", _
            vbQuestion + vbYesNo) = vbNo Then _
                Exit Sub
    End If

    ' Reload the picture.
    Screen.MousePointer = vbHourglass
    DoEvents

    LoadImage FileLoaded

    Screen.MousePointer = vbDefault
End Sub


' Save the image in the file from which it was
' loaded.
Private Sub mnuFileSave_Click()
    Screen.MousePointer = vbHourglass
    DoEvents

    SaveImage FileLoaded

    Screen.MousePointer = vbDefault
End Sub


' Save the image in a new file.
Private Sub mnuFileSaveAs_Click()
Dim fname As String

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgOpenFile.FileName = "*.BMP;*.ICO;*.RLE;*.WMF;*.DIB"
    dlgOpenFile.Flags = cdlOFNOverwritePrompt + _
        cdlOFNHideReadOnly + cdlOFNPathMustExist
    dlgOpenFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    fname = Trim$(dlgOpenFile.FileName)
    dlgOpenFile.InitDir = Left$(fname, Len(fname) _
        - Len(dlgOpenFile.FileTitle) - 1)

    ' Save the picture.
    Screen.MousePointer = vbHourglass
    DoEvents

    SaveImage fname

    Screen.MousePointer = vbDefault
End Sub


' Save the picture in the indicated file.
Private Sub SaveImage(fname As String)
    On Error GoTo SaveError
    picVisible.Picture = picVisible.Image
    SavePicture picVisible.Picture, fname

    Caption = "PalEdit [" & fname & "]"
    FileLoaded = fname
    DataChanged = False
    Exit Sub

SaveError:
    Beep
    MsgBox "Error saving picture in file " & _
        fname & "." & vbCrLf & vbCrLf & _
        Error$, , vbExclamation
    Exit Sub

End Sub



' Replace colors with a green gradient.
Private Sub mnuGradGreen_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim g As Single
Dim Dg As Single

    Dg = 255 / (StaticColor2 - StaticColor1)
    g = Dg
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With palentry(i)
            .peRed = 0
            .peGreen = g
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
        g = g + Dg
    Next i
    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub


' Replace colors with red, green, and blue
' gradients.
Private Sub mnuGradRainbow_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim num_each As Integer
Dim clr As Integer
Dim c As Single
Dim Dc As Single

    num_each = (StaticColor2 - StaticColor1) / 3
    Dc = 255 / num_each
    clr = StaticColor1 + 1
    
    ' Red shades.
    c = Dc
    For i = 1 To num_each
        With palentry(clr)
            .peRed = c
            .peGreen = 0
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
        c = c + Dc
        clr = clr + 1
    Next i
    
    ' Green shades.
    c = Dc
    For i = 1 To num_each
        With palentry(clr)
            .peRed = 0
            .peGreen = c
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
        c = c + Dc
        clr = clr + 1
    Next i
    
    ' Blue shades.
    c = Dc
    For i = clr To StaticColor2 - 1
        With palentry(clr)
            .peRed = 0
            .peGreen = 0
            .peBlue = c
            .peFlags = PC_NOCOLLAPSE
        End With
        c = c + Dc
        clr = clr + 1
    Next i

    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub


' Replace colors with a red gradient.
Private Sub mnuGradRed_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim r As Single
Dim Dr As Single

    Dr = 255 / (StaticColor2 - StaticColor1)
    r = Dr
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With palentry(i)
            .peRed = r
            .peGreen = 0
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
        r = r + Dr
    Next i
    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub

' Replace colors with appropriate greens.
Private Sub mnuNearGreen_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim clr As Integer

    ' Get the current color values.
    i = GetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1))

    ' Fill in the nearest shades.
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With palentry(i)
            clr = (CInt(.peRed) + .peGreen + .peBlue) / 3
            .peRed = 0
            .peGreen = clr
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
    Next i
    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub



' Replace colors with appropriate reds.
Private Sub mnuNearRed_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim clr As Integer

    ' Get the current color values.
    i = GetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1))

    ' Fill in the nearest shades.
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With palentry(i)
            clr = (CInt(.peRed) + .peGreen + .peBlue) / 3
            .peRed = clr
            .peGreen = 0
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
    Next i
    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub




' Replace colors with appropriate grays.
Private Sub mnuNearGray_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim clr As Integer

    ' Get the current color values.
    i = GetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1))

    ' Fill in the nearest shades.
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With palentry(i)
            clr = (CInt(.peRed) + .peGreen + .peBlue) / 3
            .peRed = clr
            .peGreen = clr
            .peBlue = clr
            .peFlags = PC_NOCOLLAPSE
        End With
    Next i
    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub




' Replace colors with appropriate blues.
Private Sub mnuNearBlue_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim clr As Integer

    ' Get the current color values.
    i = GetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1))

    ' Fill in the nearest shades.
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With palentry(i)
            clr = (CInt(.peRed) + .peGreen + .peBlue) / 3
            .peRed = 0
            .peGreen = 0
            .peBlue = clr
            .peFlags = PC_NOCOLLAPSE
        End With
    Next i
    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub




' Replace colors with a gray gradient.
Private Sub mnuGradGray_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim g As Single
Dim Dg As Single

    Dg = 255 / (StaticColor2 - StaticColor1)
    g = Dg
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With palentry(i)
            .peRed = g
            .peGreen = g
            .peBlue = g
            .peFlags = PC_NOCOLLAPSE
        End With
        g = g + Dg
    Next i
    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub



' Replace colors with a blue gradient.
Private Sub mnuGradBlue_Click()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer
Dim b As Single
Dim Db As Single

    Db = 255 / (StaticColor2 - StaticColor1)
    b = Db
    For i = StaticColor1 + 1 To StaticColor2 - 1
        With palentry(i)
            .peRed = 0
            .peGreen = 0
            .peBlue = b
            .peFlags = PC_NOCOLLAPSE
        End With
        b = b + Db
    Next i
    If SetPaletteEntries(LogicalPalette, StaticColor1 + 1, StaticColor2 - StaticColor1 - 1, palentry(StaticColor1 + 1)) = 0 Then
        Beep
        MsgBox "Error resetting colors.", , vbExclamation
        Exit Sub
    End If
    i = RealizePalette(picVisible.hdc)
    DataChanged = True
End Sub



' Set ImageScale = 1 and redraw the image.
Private Sub mnuScaleFull_Click()
    ImageScale = 1#
    ResetScrollBars
    DrawImage
End Sub

' Increase ImageScale and redraw the image.
Private Sub mnuScaleZoomIn_Click()
    ImageScale = ImageScale * 2#
    ResetScrollBars
    DrawImage
End Sub



' Decrease ImageScale and redraw the image.
Private Sub mnuScaleZoomOut_Click()
    ImageScale = ImageScale / 2#
    ResetScrollBars
    DrawImage
End Sub


' Update the selected color's value.
Private Sub hbarRed_Change()
    If SettingColor Then Exit Sub
    lblRed.Caption = Format$(hbarRed.Value)
    UpdatePalette
End Sub


' Update the selected color's value.
Private Sub hbarRed_Scroll()
    If SettingColor Then Exit Sub
    lblRed.Caption = Format$(hbarRed.Value)
    UpdatePalette
End Sub

' Select the color the user clicked on.
Private Sub picSystemColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim j As Integer

    i = Y \ Dx
    j = X \ Dy
    SelectColor i, j
End Sub

' End the application. (See also the QueryUnload event.)
Private Sub mnuFileExit_Click()
    Unload Me
End Sub


' Allow the user to select a new color with the arrow keys.
Private Sub picSystemColors_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim j As Integer

    i = SelectedI
    j = SelectedJ

    Select Case KeyCode
        Case vbKeyDown
            i = i + 1
            If i * 16 + j >= SysPalSize Then i = 0
        
        Case vbKeyUp
            i = i - 1
            If i < 0 Then
                i = (SysPalSize - 1) \ 16
                If i * 16 + j >= SysPalSize Then _
                    i = i - 1
            End If
        
        Case vbKeyLeft
            j = j - 1
            If j < 0 Then
                j = 15
                If i * 16 + j >= SysPalSize Then _
                    j = SysPalSize - 1 - i * 16
            End If
        
        Case vbKeyRight
            j = j + 1
            If j > 15 Or _
                i * 16 + j >= SysPalSize Then _
                    j = 0
        
    End Select
    
    SelectColor i, j
End Sub



' Redraw the image scrolled appropriately.
Private Sub VBar_Change()
    DrawImage
End Sub


' Redraw the image scrolled appropriately.
Private Sub VBar_Scroll()
    DrawImage
End Sub
