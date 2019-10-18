VERSION 5.00
Begin VB.Form frmImagePic 
   Caption         =   "ImagePic"
   ClientHeight    =   3630
   ClientLeft      =   1605
   ClientTop       =   1140
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   6150
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   495
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy ==>"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.PictureBox picScribble 
      Height          =   3015
      Index           =   1
      Left            =   3120
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.PictureBox picScribble 
      AutoRedraw      =   -1  'True
      Height          =   3015
      Index           =   0
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmImagePic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DrawingIndex As Integer
Private LastX As Single
Private LastY As Single
' Clear the corresponding picture.
Private Sub CmdClear_Click(Index As Integer)
    picScribble(Index).Cls
End Sub


' Copy picScribble(0)'s current display to
' picScribble(1)'s permanent background.
Private Sub CmdCopy_Click()
    picScribble(1).Picture = picScribble(0).Image
End Sub



Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DrawingIndex = -1
End Sub

Private Sub picScribble_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawingIndex = Index
    picScribble(Index).CurrentX = X
    picScribble(Index).CurrentY = Y
    LastX = X
    LastY = Y
End Sub


Private Sub picScribble_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawingIndex <> Index Then Exit Sub
    picScribble(Index).Line -(X, Y)
    LastX = X
    LastY = Y
End Sub


Private Sub picScribble_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawingIndex <> Index Then Exit Sub
    DrawingIndex = -1
    picScribble(Index).Line -(X, Y)
End Sub
