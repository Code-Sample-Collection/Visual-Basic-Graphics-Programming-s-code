VERSION 5.00
Begin VB.Form frmTile 
   Caption         =   "Tile"
   ClientHeight    =   5490
   ClientLeft      =   1740
   ClientTop       =   900
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   6870
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   3000
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image imgTile 
      Height          =   945
      Index           =   30
      Left            =   1560
      Picture         =   "Tile.frx":0000
      Top             =   4440
      Width           =   945
   End
   Begin VB.Image imgTile 
      Height          =   945
      Index           =   28
      Left            =   120
      Picture         =   "Tile.frx":0242
      Top             =   4440
      Width           =   945
   End
   Begin VB.Image imgTile 
      Height          =   600
      Index           =   27
      Left            =   2280
      Picture         =   "Tile.frx":31C4
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image imgTile 
      Height          =   600
      Index           =   26
      Left            =   1560
      Picture         =   "Tile.frx":334E
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image imgTile 
      Height          =   600
      Index           =   25
      Left            =   840
      Picture         =   "Tile.frx":34D8
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image imgTile 
      Height          =   600
      Index           =   24
      Left            =   120
      Picture         =   "Tile.frx":47DA
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image imgTile 
      Height          =   600
      Index           =   23
      Left            =   2280
      Picture         =   "Tile.frx":4B7C
      Top             =   3000
      Width           =   600
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   22
      Left            =   1560
      Picture         =   "Tile.frx":4D06
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   600
      Index           =   21
      Left            =   840
      Picture         =   "Tile.frx":4DD0
      Top             =   3000
      Width           =   600
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   20
      Left            =   120
      Picture         =   "Tile.frx":60D2
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   19
      Left            =   2280
      Picture         =   "Tile.frx":6D14
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   18
      Left            =   1560
      Picture         =   "Tile.frx":6DDE
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   17
      Left            =   840
      Picture         =   "Tile.frx":6EA8
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   16
      Left            =   120
      Picture         =   "Tile.frx":7AEA
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   420
      Index           =   15
      Left            =   2280
      Picture         =   "Tile.frx":872C
      Top             =   120
      Width           =   420
   End
   Begin VB.Image imgTile 
      Height          =   420
      Index           =   14
      Left            =   1560
      Picture         =   "Tile.frx":87E6
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgTile 
      Height          =   420
      Index           =   13
      Left            =   840
      Picture         =   "Tile.frx":88A0
      Top             =   120
      Width           =   420
   End
   Begin VB.Image imgTile 
      Height          =   420
      Index           =   12
      Left            =   120
      Picture         =   "Tile.frx":9212
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgTile 
      Height          =   450
      Index           =   11
      Left            =   2280
      Picture         =   "Tile.frx":9A34
      Top             =   1800
      Width           =   450
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   10
      Left            =   1560
      Picture         =   "Tile.frx":9AF6
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   450
      Index           =   9
      Left            =   840
      Picture         =   "Tile.frx":9BC0
      Top             =   1800
      Width           =   450
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   8
      Left            =   120
      Picture         =   "Tile.frx":A6CA
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   7
      Left            =   2280
      Picture         =   "Tile.frx":B30C
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   6
      Left            =   1560
      Picture         =   "Tile.frx":B3D6
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   5
      Left            =   840
      Picture         =   "Tile.frx":B4A0
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   4
      Left            =   120
      Picture         =   "Tile.frx":C0E2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "Tile.frx":CD24
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   2
      Left            =   1560
      Picture         =   "Tile.frx":CDEE
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   1
      Left            =   840
      Picture         =   "Tile.frx":CEB8
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "Tile.frx":DAFA
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TileChoice As Integer
' Tile a PictureBox with the picture in an Image control.
Private Sub TilePicture(ByVal pic As PictureBox, ByVal tile_image As Image)
Dim wid As Integer
Dim hgt As Integer
Dim rows As Integer
Dim cols As Integer
Dim r As Integer
Dim c As Integer
Dim x As Integer
Dim y As Integer

    pic.Cls     ' Clear the picture box.
    wid = ScaleX(tile_image.Width, tile_image.Parent.ScaleMode, pic.ScaleMode)
    hgt = ScaleY(tile_image.Height, tile_image.Parent.ScaleMode, pic.ScaleMode)

    ' See how many rows and columns we will need.
    cols = Int(pic.ScaleWidth / wid + 1)
    rows = Int(pic.ScaleHeight / hgt + 1)
    
    ' Copy the tile.
    y = 0
    For r = 1 To rows
        x = 0
        For c = 1 To cols
            pic.PaintPicture tile_image.Picture, x, y
            x = x + wid
        Next c
        y = y + hgt
    Next r
End Sub

' Tile the form.
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
    TilePicture picCanvas, imgTile(TileChoice)
End Sub


Private Sub imgTile_Click(Index As Integer)
    TileChoice = Index
    TilePicture picCanvas, imgTile(Index)
End Sub
