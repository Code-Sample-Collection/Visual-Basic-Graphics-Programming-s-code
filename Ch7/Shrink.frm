VERSION 5.00
Begin VB.Form frmShrink 
   Caption         =   "Shrink"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   303
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSmall 
      Height          =   1410
      Left            =   3000
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   1
      Top             =   840
      Width           =   1410
   End
   Begin VB.PictureBox picOriginal 
      Height          =   2760
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2760
   End
End
Attribute VB_Name = "frmShrink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Const PI = 3.14159265
Const NUM_LINES = 40

Dim xmid As Single
Dim ymid As Single
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim theta As Single
Dim dtheta As Single
Dim r1 As Single
Dim r2 As Single
Dim i  As Integer

    picOriginal.AutoRedraw = True
    xmid = picOriginal.ScaleWidth / 2
    ymid = picOriginal.ScaleHeight / 2
    r1 = picOriginal.ScaleWidth * 0.45
    r2 = picOriginal.ScaleWidth * 0.1
    dtheta = 2 * PI / NUM_LINES
    theta = 0
    For i = 1 To NUM_LINES
        x1 = xmid + r1 * Cos(theta)
        y1 = xmid + r1 * Sin(theta)
        x2 = xmid + r2 * Cos(theta)
        y2 = xmid + r2 * Sin(theta)
        picOriginal.Line (x1, y1)-(x2, y2)
        theta = theta + dtheta
    Next i

    picOriginal.Picture = picOriginal.Image

    picSmall.AutoRedraw = True
    picSmall.PaintPicture _
        picOriginal.Picture, 0, 0, _
        picSmall.ScaleWidth, _
        picSmall.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight
End Sub
