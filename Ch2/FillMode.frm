VERSION 5.00
Begin VB.Form frmFillMode 
   Caption         =   "FillMode"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBothAlternateCW 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   4560
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   6
      Top             =   480
      Width           =   2145
   End
   Begin VB.PictureBox picBothWindingCW 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   4560
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   5
      Top             =   2520
      Width           =   2145
   End
   Begin VB.PictureBox picBothAlternate 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   2400
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   3
      Top             =   480
      Width           =   2145
   End
   Begin VB.PictureBox picBothWinding 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   2400
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   2
      Top             =   2520
      Width           =   2145
   End
   Begin VB.PictureBox picStarWinding 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   240
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   1
      Top             =   2520
      Width           =   2145
   End
   Begin VB.PictureBox picStarAlternate 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   240
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   0
      Top             =   480
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Star Counterclockwise"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   9
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Star Counterclockwise"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rectangle Clockwise"
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Star Counterclockwise"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmFillMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long

Private Const ALTERNATE = 1
Private Const WINDING = 2

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W2 As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const FW_BOLD = 700
Private Const ANSI_CHARSET = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_LH_ANGLES = 16
Private Const PROOF_QUALITY = 2
Private Const VARIABLE_PITCH = 2
Private Const FF_ROMAN = 16

Private Sub DrawLabels()
Dim new_font As Long
Dim old_font As Long
Dim wid As Single
Dim hgt As Single

    AutoRedraw = True
    ScaleMode = vbPixels

    ' Create the rotated font.
    new_font = CreateFont(12, 0, 900, 900, _
        FW_BOLD, False, False, False, _
        ANSI_CHARSET, OUT_DEFAULT_PRECIS, _
        CLIP_LH_ANGLES Or CLIP_CHARACTER_PRECIS, _
        PROOF_QUALITY, 0, "Arial")
    SelectObject hdc, new_font

    hgt = TextWidth("ALTERNATE")
    CurrentX = 0
    CurrentY = picStarAlternate.Top + _
        (picStarAlternate.Height + hgt) / 2
    Print "ALTERNATE"

    hgt = TextWidth("WINDING")
    CurrentX = 0
    CurrentY = picStarWinding.Top + _
        (picStarWinding.Height + hgt) / 2
    Print "WINDING"

    SelectObject hdc, old_font
    DeleteObject new_font
End Sub

' Initialize the point data.
Private Sub InitializeData(pt() As POINTAPI, num_pts() As Long)
    ' Counterclockwise star.
    pt(1).x = 72:  pt(1).y = 3
    pt(2).x = 8:   pt(2).y = 119
    pt(3).x = 133: pt(3).y = 24
    pt(4).x = 16:  pt(4).y = 37
    pt(5).x = 120: pt(5).y = 108
    num_pts(1) = 5

    ' Counterclockwise rectangle.
    pt(6).x = 43:  pt(6).y = 110
    pt(7).x = 108: pt(7).y = 110
    pt(8).x = 108: pt(8).y = 16
    pt(9).x = 43:  pt(9).y = 16
    num_pts(2) = 4
End Sub
' Make the rectangle's points clockwise.
Private Sub ReverseRectangle(pt() As POINTAPI, num_pts() As Long)
    ' Clockwise rectangle.
    pt(6).x = 43:  pt(6).y = 16
    pt(7).x = 108: pt(7).y = 16
    pt(8).x = 108: pt(8).y = 110
    pt(9).x = 43:  pt(9).y = 110
    num_pts(2) = 4
End Sub

Private Sub Form_Load()
Dim pt(1 To 100) As POINTAPI
Dim num_pts(1 To 2) As Long

    ' Initialize the data.
    InitializeData pt, num_pts

    ' Draw just the star with fill style ALTERNATE.
    SetPolyFillMode picStarAlternate.hdc, ALTERNATE
    PolyPolygon picStarAlternate.hdc, pt(1), num_pts(1), 1

    ' Draw just the star with fill style WINDING.
    SetPolyFillMode picStarWinding.hdc, WINDING
    PolyPolygon picStarWinding.hdc, pt(1), num_pts(1), 1

    ' Draw both shapes with fill style ALTERNATE.
    SetPolyFillMode picBothAlternate.hdc, ALTERNATE
    PolyPolygon picBothAlternate.hdc, pt(1), num_pts(1), 2

    ' Draw both shapes with fill style WINDING.
    SetPolyFillMode picBothWinding.hdc, WINDING
    PolyPolygon picBothWinding.hdc, pt(1), num_pts(1), 2

    ' Make the rectangle clockwise.
    ReverseRectangle pt, num_pts

    ' Draw both shapes with fill style ALTERNATE.
    SetPolyFillMode picBothAlternateCW.hdc, ALTERNATE
    PolyPolygon picBothAlternateCW.hdc, pt(1), num_pts(1), 2

    ' Draw both shapes with fill style WINDING.
    SetPolyFillMode picBothWindingCW.hdc, WINDING
    PolyPolygon picBothWindingCW.hdc, pt(1), num_pts(1), 2

    ' Draw the ALTERNATE and WINDING labels.
    DrawLabels
End Sub
