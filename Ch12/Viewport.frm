VERSION 5.00
Begin VB.Form frmViewport 
   Caption         =   "Viewport"
   ClientHeight    =   2910
   ClientLeft      =   2550
   ClientTop       =   1515
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   2910
   Begin VB.PictureBox picViewport 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Draw a smiley face in the viewport centered
' around the point (5, 5).
Private Sub DrawSmiley(ByVal pic As PictureBox)
Const PI = 3.14159265

    ' Head.
    pic.FillColor = vbYellow
    pic.FillStyle = vbSolid
    pic.Circle (5, 5), 4

    ' Nose.
    pic.FillColor = RGB(0, &HFF, &H80)
    pic.Circle (5, 4.5), 1, , , , 1.5

    ' Eye whites.
    pic.FillColor = vbWhite
    pic.Circle (3.5, 6), 0.75, , , , 1.25
    pic.Circle (6.5, 6), 0.75, , , , 1.25

    ' Pupils.
    pic.FillColor = vbBlack
    pic.Circle (3.7, 6), 0.5, , , , 1.25
    pic.Circle (6.7, 6), 0.5, , , , 1.25

    ' Smile.
    pic.Circle (5, 5), 2.75, , 1.15 * PI, 1.8 * PI
End Sub
Private Sub Form_Load()
Dim X As Single
Dim Y As Single
Dim border_wid As Single
Dim border_hgt As Single
Dim wid As Single
Dim hgt As Single

    ' Find the PictureBox's border sizes.
    border_wid = picViewport.Width - picViewport.ScaleWidth
    border_hgt = picViewport.Height - picViewport.ScaleHeight
    wid = 2 * 1440 + border_wid
    hgt = 2 * 1440 + border_hgt

    ' Make the viewport 2 inches square.
    X = picViewport.Left
    Y = picViewport.Top
    picViewport.Move X, Y, wid, hgt

    ' Scale the world window.
    picViewport.ScaleLeft = 0
    picViewport.ScaleTop = 10
    picViewport.ScaleWidth = (10 - 0)
    picViewport.ScaleHeight = (0 - 10)
End Sub


Private Sub Form_Resize()
    picViewport.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


Private Sub picViewport_Paint()
    DrawSmiley picViewport
End Sub


