VERSION 5.00
Begin VB.Form frmRainbow 
   Caption         =   "Rainbow"
   ClientHeight    =   3255
   ClientLeft      =   2055
   ClientTop       =   1320
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   Begin VB.PictureBox picDefault 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   2475
      TabIndex        =   12
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox picRainbow 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   2760
      Picture         =   "Rainbow.frx":0000
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.HScrollBar hbarBlue 
      Height          =   255
      LargeChange     =   16
      Left            =   1020
      Max             =   255
      TabIndex        =   6
      Top             =   2880
      Width           =   4275
   End
   Begin VB.HScrollBar hbarGreen 
      Height          =   255
      LargeChange     =   16
      Left            =   1020
      Max             =   255
      TabIndex        =   5
      Top             =   2520
      Width           =   4275
   End
   Begin VB.HScrollBar hbarRed 
      Height          =   255
      LargeChange     =   16
      Left            =   1020
      Max             =   255
      TabIndex        =   4
      Top             =   2160
      Width           =   4275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rainbow Palette"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblBlue 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblGreen 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblRed 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Default Palette"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Blue"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Green"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   495
   End
End
Attribute VB_Name = "frmRainbow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const RASTERCAPS = 38
Private Const RC_PALETTE = &H100
' Display the selected RGB value in all picture
' boxes.
Private Sub UpdateColors()
Dim r As Integer
Dim g As Integer
Dim b As Integer

    r = hbarRed.Value
    g = hbarGreen.Value
    b = hbarBlue.Value

    ' Update the numeric labels.
    lblRed.Caption = Format$(r)
    lblGreen.Caption = Format$(g)
    lblBlue.Caption = Format$(b)

    ' Display the color in the default picture.
    picDefault.Line (0, 0)-Step(picDefault.ScaleWidth, picDefault.ScaleHeight), RGB(r, g, b), BF

    ' Display the color in the rainbow picture.
    picRainbow.Line (0, 0)-Step(picRainbow.ScaleWidth, picRainbow.ScaleHeight), RGB(r, g, b), BF
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


