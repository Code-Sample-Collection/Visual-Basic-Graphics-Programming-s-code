VERSION 5.00
Begin VB.Form frmCurves 
   Caption         =   "Curves"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPie 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   2640
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   6
      Top             =   2160
      Width           =   2535
   End
   Begin VB.PictureBox picArc 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   2640
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox picChord 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   0
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.PictureBox picEllipse 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   0
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pie"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Arc"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Chord"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ellipse"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmCurves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Private Sub Form_Load()
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim X3 As Single
Dim Y3 As Single
Dim X4 As Single
Dim Y4 As Single

    picEllipse.FillStyle = vbDiagonalCross
    picChord.FillStyle = vbDiagonalCross
    PicPie.FillStyle = vbDiagonalCross

    ' Get coordinates to use for all curves.
    X1 = 5
    Y1 = 5
    X2 = picEllipse.ScaleWidth - 10
    Y2 = picEllipse.ScaleHeight - 10
    X3 = picEllipse.ScaleWidth - 10
    Y3 = picEllipse.ScaleHeight - 10
    X4 = picEllipse.ScaleWidth * 0.45
    Y4 = picEllipse.ScaleHeight * 0.25

    ' Draw an ellipse.
    picEllipse.DrawWidth = 3
    Ellipse picEllipse.hdc, X1, Y1, X2, Y2

    ' Draw an arc.
    picArc.DrawWidth = 3
    Arc picArc.hdc, X1, Y1, X2, Y2, X3, Y3, X4, Y4
    picArc.DrawWidth = 1
    picArc.DrawStyle = vbDot
    picArc.Line (X3, Y3)-((X1 + X2) / 2, (Y1 + Y2) / 2)
    picArc.Line -(X4, Y4)

    ' Draw a chord.
    picChord.DrawWidth = 3
    Chord picChord.hdc, X1, Y1, X2, Y2, X3, Y3, X4, Y4
    picChord.DrawWidth = 1
    picChord.DrawStyle = vbDot
    picChord.Line (X3, Y3)-((X1 + X2) / 2, (Y1 + Y2) / 2)
    picChord.Line -(X4, Y4)

    ' Draw a pie slice.
    PicPie.DrawWidth = 3
    Pie PicPie.hdc, X1, Y1, X2, Y2, X3, Y3, X4, Y4
    PicPie.DrawWidth = 1
    PicPie.DrawStyle = vbDot
    PicPie.Line (X3, Y3)-((X1 + X2) / 2, (Y1 + Y2) / 2)
    PicPie.Line -(X4, Y4)
End Sub
