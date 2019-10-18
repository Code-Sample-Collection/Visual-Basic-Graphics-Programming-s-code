VERSION 5.00
Begin VB.Form frmAuto 
   Caption         =   "Auto"
   ClientHeight    =   4365
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   7095
   Begin VB.PictureBox PaintPict 
      Height          =   4095
      Left            =   3600
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.PictureBox AutoPict 
      AutoRedraw      =   -1  'True
      Height          =   4095
      Left            =   0
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Paint Events"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "AutoRedraw"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Draw squiggly lines.
Private Sub DrawPict(ByVal pic As PictureBox)
Const Amp = 3
Const PI = 3.14159
Const Per = 4 * PI

Dim i As Single
Dim j As Single
Dim hgt As Single
Dim wid As Single
    
    pic.ScaleMode = vbPixels
    pic.Cls     ' Clear the picture box.

    For i = 0 To pic.ScaleHeight Step 4
        pic.CurrentX = 0
        pic.CurrentY = i
        For j = 0 To pic.ScaleWidth
            pic.Line -(j, i + Amp * Sin(j / Per))
        Next j
    Next i
    For i = 1 To hgt Step 2
        pic.Line (0, i)-(wid, i)
    Next i
End Sub
' Draw on the PictureBox with AutoRedraw = True.
Private Sub Form_Load()
    DrawPict AutoPict
End Sub

' Draw on the PictureBox with AutoRedraw = False.
Private Sub PaintPict_Paint()
    ' This makes the other PictureBox refresh
    ' itself before we start hogging the CPU.
    Refresh

    DrawPict PaintPict
End Sub
