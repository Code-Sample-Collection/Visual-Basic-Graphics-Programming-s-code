VERSION 5.00
Begin VB.Form PalWatchForm 
   Caption         =   "PalWatch"
   ClientHeight    =   2460
   ClientLeft      =   6810
   ClientTop       =   975
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   164
   Begin VB.Timer ColorTimer 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   0
      Picture         =   "PalWatch.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Menu mnuColor 
      Caption         =   "(0, 0, 0)"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "PalWatchForm"
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

Private Const PC_EXPLICIT = &H2
Private Const RASTERCAPS = 38
Private Const RC_PALETTE = &H100
Private Const NUMRESERVED = 106
Private Const SIZEPALETTE = 104

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ResizePalette Lib "gdi32" (ByVal hPalette As Long, ByVal nNumEntries As Long) As Long
Private Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long

Private Const PALETTE_INDEX = &H1000000
Private Const NO_COLOR = -1

Private LogicalPalette As Long

Private SysPalSize As Integer
Private NumStaticColors As Integer

Private SelectedI As Integer
Private SelectedJ As Integer
Private SelectedColor As Integer
Private SelectedR As Integer
Private SelectedG As Integer
Private SelectedB As Integer

Private dx As Integer
Private dy As Integer
' Load the Pict palette with PC_EXPLICIT entries
' so they match the system palette.
Private Sub LoadSystemPalette()
Dim palentry(0 To 255) As PALETTEENTRY
Dim i As Integer

    ' Make the logical palette as big as possible.
    LogicalPalette = picCanvas.Picture.hPal
    If ResizePalette(LogicalPalette, SysPalSize) = 0 Then
        MsgBox "Error resizing the palette."
        End
    End If

    ' Flag all palette entries as PC_EXPLICIT.
    ' Set peRed to the system palette indexes.
    For i = 0 To SysPalSize - 1
        palentry(i).peRed = i
        palentry(i).peFlags = PC_EXPLICIT
    Next i
    
    ' Update the palette (ignore return value).
    i = SetPaletteEntries(LogicalPalette, 0, SysPalSize, palentry(0))
End Sub

' Fill the system picture with all the palette
' colors, hatching the static colors.
Private Sub ShowColors()
Dim i As Integer
Dim j As Integer
Dim clr As Integer
Dim oldfill As Integer
Dim olddraw As Integer

    picCanvas.Cls
    
    ' Display the colors using palette indexing.
    dx = picCanvas.ScaleWidth / 16
    dy = picCanvas.ScaleHeight / 16
    clr = 0
    For i = 0 To 15
        For j = 0 To 15
            picCanvas.Line (j * dx, i * dy)-Step(dx, dy), _
                clr + PALETTE_INDEX, BF
            clr = clr + 1
        Next j
    Next i

    ' Hatch the static colors.
    oldfill = picCanvas.FillStyle
    olddraw = picCanvas.DrawMode
    picCanvas.FillStyle = vbDownwardDiagonal
    picCanvas.DrawMode = vbInvisible

    picCanvas.Line (0, 0)-Step((NumStaticColors \ 2) * dx - 1, dy - 1), , B
    picCanvas.Line (j * dx, i * dy)-Step(-(NumStaticColors \ 2) * dx, -dy), , B

    picCanvas.FillStyle = oldfill
    picCanvas.DrawMode = olddraw

    ' Highlight the previously selected color.
    SelectedColor = NO_COLOR
    SelectColor SelectedI, SelectedJ
End Sub
' Select the color at the indicated position.
Private Sub SelectColor(ByVal i As Integer, ByVal j As Integer)
Const GAP1 = 1
Const GAP2 = 2
Const DRAW_WID = 2

Dim oldmode As Integer
Dim oldwid As Integer

    oldmode = picCanvas.DrawMode
    oldwid = picCanvas.DrawWidth
    picCanvas.DrawMode = vbInvert
    picCanvas.DrawWidth = DRAW_WID
    
    ' Unhighlight the previously selected color.
    If SelectedColor <> NO_COLOR Then _
        picCanvas.Line (SelectedJ * dx + GAP1, SelectedI * dx + GAP1)-Step(dx - GAP2, dx - GAP2), , B

    ' Record the new color.
    SelectedI = i
    SelectedJ = j
    SelectedColor = i * 16 + j

    ' Highlight the new color.
    picCanvas.Line (SelectedJ * dx + GAP1, SelectedI * dx + GAP1)-Step(dx - GAP2, dx - GAP2), , B
    picCanvas.DrawMode = oldmode
    picCanvas.DrawWidth = oldwid

    ' Display the color's components in mnuColor.
    ShowColorValue
End Sub


' If the selected color's components have
' changed, display the new values in mnuColor.
Private Sub ShowColorValue()
Dim palentry As PALETTEENTRY
Dim status As Integer

    status = GetSystemPaletteEntries(picCanvas.hdc, SelectedColor, 1, palentry)
    If palentry.peRed <> SelectedR Or _
       palentry.peGreen <> SelectedG Or _
       palentry.peBlue <> SelectedB Then
            mnuColor.Caption = "(" & _
                Format$(palentry.peRed) & "," & _
                Str$(palentry.peGreen) & "," & _
                Str$(palentry.peBlue) & ")"
    End If
End Sub

' Make sure the selected color's components are up to date.
Private Sub ColorTimer_Timer()
    ShowColorValue
End Sub
' Get basic palette information.
Private Sub Form_Load()
    ' Make sure the screen supports palettes.
    If Not GetDeviceCaps(hdc, RASTERCAPS) And RC_PALETTE Then
        MsgBox "This system is not using palettes."
        End
    End If

    ' See how big the system palette is.
    SysPalSize = GetDeviceCaps(hdc, SIZEPALETTE)

    ' See how many colors are reserved.
    NumStaticColors = GetDeviceCaps(hdc, NUMRESERVED)

    ' Load the system palette.
    LoadSystemPalette
End Sub

' Make the picture as large as possible.
Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single

    If WindowState = vbMinimized Then Exit Sub

    wid = ScaleWidth - 2 * picCanvas.Left
    If wid < 10 Then wid = 10
    hgt = ScaleHeight - 2 * picCanvas.Top
    If hgt < 10 Then hgt = 10
    picCanvas.Move picCanvas.Left, picCanvas.Top, wid, hgt

    ' Display the colors.
    ShowColors
End Sub
' Select the color the user clicked on.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim j As Integer

    i = Y \ dx
    j = X \ dy
    SelectColor i, j
End Sub
' Allow the user to select a new color with the
' arrow keys.
Private Sub picCanvas_KeyDown(KeyCode As Integer, Shift As Integer)
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
