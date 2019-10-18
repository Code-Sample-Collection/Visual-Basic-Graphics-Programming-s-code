VERSION 5.00
Begin VB.Form frmCustom 
   Caption         =   "Custom"
   ClientHeight    =   3270
   ClientLeft      =   2055
   ClientTop       =   1320
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   Begin VB.HScrollBar hbarRed 
      Height          =   255
      LargeChange     =   16
      Left            =   1020
      Max             =   255
      TabIndex        =   4
      Top             =   2160
      Width           =   4275
   End
   Begin VB.HScrollBar hbarGreen 
      Height          =   255
      LargeChange     =   16
      Left            =   1020
      Max             =   255
      TabIndex        =   3
      Top             =   2520
      Width           =   4275
   End
   Begin VB.HScrollBar hbarBlue 
      Height          =   255
      LargeChange     =   16
      Left            =   1020
      Max             =   255
      TabIndex        =   2
      Top             =   2880
      Width           =   4275
   End
   Begin VB.PictureBox picCustom 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   2760
      Picture         =   "Custom.frx":0000
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox picDefault 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Green"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Blue"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Default Palette"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblRed 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblGreen 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblBlue 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Custom Palette"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "frmCustom"
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
Private Declare Function ResizePalette Lib "gdi32" (ByVal hPalette As Long, ByVal nNumEntries As Long) As Long
Private Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long

Private Const RASTERCAPS = 38
Private Const RC_PALETTE = &H100

' Resize picCustom's palette so it has only one
' entry. We will use that entry to display the
' color selected by the scroll bars.
Private Sub ShrinkPalette()
    If ResizePalette(picCustom.Picture.hPal, 1) = 0 Then
        MsgBox "Error resizing palette."
    End If
End Sub


' Display the selected RGB value in all picture
' boxes.
Private Sub UpdateColors()
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim palentry As PALETTEENTRY

    r = hbarRed.Value
    g = hbarGreen.Value
    b = hbarBlue.Value

    ' Update the numeric labels.
    lblRed.Caption = Format$(r)
    lblGreen.Caption = Format$(g)
    lblBlue.Caption = Format$(b)

    ' Display the color in the default picture.
    picDefault.Line (0, 0)-Step(picDefault.ScaleWidth, picDefault.ScaleHeight), RGB(r, g, b), BF

    ' Put the new color in the custom palette.
    palentry.peRed = r
    palentry.peGreen = g
    palentry.peBlue = b
    If SetPaletteEntries(picCustom.Picture.hPal, 0, 1, palentry) = 0 Then
        MsgBox "Error updating palette entry."
    End If
    
    ' Make the change take effect.
    RealizePalette picCustom.hdc
    
    ' Fill the custom palette picture.
    picCustom.Line (0, 0)-Step(picCustom.ScaleWidth, picCustom.ScaleHeight), RGB(r, g, b) + &H2000000, BF
End Sub
Private Sub hbarBlue_Change()
    UpdateColors
End Sub

Private Sub hbarBlue_Scroll()
    UpdateColors
End Sub


Private Sub Form_Load()
    ' Make sure the screen supports palettes.
    If Not GetDeviceCaps(hdc, RASTERCAPS) And RC_PALETTE Then
        MsgBox "This system is not using palettes.", vbCritical
        End
    End If

    ' Load the system palette.
    ShrinkPalette

    ' Display the initial color (black).
    UpdateColors
End Sub


Private Sub hbarGreen_Change()
    UpdateColors
End Sub

Private Sub hbarGreen_Scroll()
    UpdateColors
End Sub



Private Sub hbarRed_Change()
    UpdateColors
End Sub


Private Sub hbarRed_Scroll()
    UpdateColors
End Sub


