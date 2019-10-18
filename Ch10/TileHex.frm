VERSION 5.00
Begin VB.Form frmTileHex 
   Caption         =   "TileHex"
   ClientHeight    =   4740
   ClientLeft      =   1740
   ClientTop       =   900
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4740
   ScaleWidth      =   5535
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   1200
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image imgTile 
      Height          =   975
      Index           =   5
      Left            =   0
      Picture         =   "TileHex.frx":0000
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Image imgMask 
      Height          =   975
      Index           =   5
      Left            =   840
      Picture         =   "TileHex.frx":3922
      Top             =   1080
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image imgTile 
      Height          =   405
      Index           =   4
      Left            =   0
      Picture         =   "TileHex.frx":3C78
      Top             =   4320
      Width           =   465
   End
   Begin VB.Image imgMask 
      Height          =   405
      Index           =   4
      Left            =   600
      Picture         =   "TileHex.frx":46DA
      Top             =   4320
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgMask 
      Height          =   405
      Index           =   3
      Left            =   600
      Picture         =   "TileHex.frx":4E7C
      Top             =   3840
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgTile 
      Height          =   405
      Index           =   3
      Left            =   0
      Picture         =   "TileHex.frx":561E
      Top             =   3840
      Width           =   465
   End
   Begin VB.Image imgTile 
      Height          =   735
      Index           =   2
      Left            =   0
      Picture         =   "TileHex.frx":5DC0
      Top             =   3000
      Width           =   735
   End
   Begin VB.Image imgMask 
      Height          =   735
      Index           =   2
      Left            =   840
      Picture         =   "TileHex.frx":6BF6
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgTile 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "TileHex.frx":7A2C
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgMask 
      Height          =   735
      Index           =   1
      Left            =   840
      Picture         =   "TileHex.frx":8862
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgMask 
      Height          =   975
      Index           =   0
      Left            =   840
      Picture         =   "TileHex.frx":9698
      Top             =   0
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image imgTile 
      Height          =   975
      Index           =   0
      Left            =   0
      Picture         =   "TileHex.frx":99EE
      Top             =   0
      Width           =   1110
   End
End
Attribute VB_Name = "frmTileHex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TileChoice As Integer

Private ColDx() As Integer
Private RowDx() As Integer
Private RowDy() As Integer
' Tile the PictureBox with the image in the image control.
Private Sub TilePicture(ByVal pic As PictureBox, ByVal tile_image As Image, ByVal mask_image As Image, ByVal cdx As Integer, ByVal rdx As Integer, ByVal rdy As Integer)
Dim wid As Single
Dim hgt As Single
Dim x As Integer
Dim y As Integer
Dim x1 As Integer
Dim x2 As Integer
Dim startx As Integer

    pic.Cls     ' Clear the picture box.

    ' Start above and to the left of the drawing area.
    wid = tile_image.Parent.ScaleX(tile_image.Width, tile_image.Parent.ScaleMode, pic.ScaleMode)
    hgt = tile_image.Parent.ScaleY(tile_image.Height, tile_image.Parent.ScaleMode, pic.ScaleMode)
    y = -hgt
    x1 = -wid
    x2 = x1 + rdx
    startx = x1

    ' Copy the tile until we're to the right and
    ' below the drawing area.
    Do While y <= pic.ScaleHeight
        x = startx
        Do While x <= pic.ScaleWidth
            ' Copy the mask with vbMergePaint.
            pic.PaintPicture mask_image.Picture, x, y, , , , , , , vbMergePaint

            ' Copy the mask with vbSrcAnd.
            pic.PaintPicture tile_image.Picture, x, y, , , , , , , vbSrcAnd
            x = x + cdx
        Loop

        If startx = x1 Then
            startx = x2
        Else
            startx = x1
        End If
        y = y + rdy
    Loop
End Sub
' Initialize row and column offsets.
Private Sub Form_Load()
    ReDim ColDx(imgTile.LBound To imgTile.UBound)
    ReDim RowDx(imgTile.LBound To imgTile.UBound)
    ReDim RowDy(imgTile.LBound To imgTile.UBound)

    ColDx(0) = 108
    RowDx(0) = 54
    RowDy(0) = 31

    ColDx(1) = 72
    RowDx(1) = 35
    RowDy(1) = 20

    ColDx(2) = ColDx(1)
    RowDx(2) = RowDx(1)
    RowDy(2) = RowDy(1)

    ColDx(3) = ColDx(1)
    RowDx(3) = RowDx(1)
    RowDy(3) = RowDy(1)

    ColDx(4) = 46
    RowDx(4) = 23
    RowDy(4) = 13

    ColDx(5) = 108
    RowDx(5) = 54
    RowDy(5) = 31
End Sub
' Tile the form.
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight

    TilePicture picCanvas, imgTile(TileChoice), imgMask(TileChoice), ColDx(TileChoice), RowDx(TileChoice), RowDy(TileChoice)
End Sub

Private Sub imgTile_Click(Index As Integer)
    TileChoice = Index
    TilePicture picCanvas, imgTile(TileChoice), imgMask(TileChoice), ColDx(TileChoice), RowDx(TileChoice), RowDy(TileChoice)
End Sub
