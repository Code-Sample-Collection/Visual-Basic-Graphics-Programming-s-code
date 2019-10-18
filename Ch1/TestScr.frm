VERSION 5.00
Object = "*\AScrWin.vbp"
Begin VB.Form frmTestScr 
   Caption         =   "ScrollingWindow"
   ClientHeight    =   4020
   ClientLeft      =   4110
   ClientTop       =   2295
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   3315
   Begin Project1.ScrolledWindow swinLake 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _extentx        =   5741
      _extenty        =   7011
      Begin VB.Image imgLake 
         Height          =   5295
         Left            =   0
         Picture         =   "TestScr.frx":0000
         Top             =   0
         Width           =   3720
      End
   End
End
Attribute VB_Name = "frmTestScr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    swinLake.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

