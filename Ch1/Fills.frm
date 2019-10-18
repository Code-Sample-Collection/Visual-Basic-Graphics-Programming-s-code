VERSION 5.00
Begin VB.Form frmFills 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Fills"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbDiagonalCross"
      Height          =   255
      Index           =   7
      Left            =   2820
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbCross"
      Height          =   255
      Index           =   6
      Left            =   1020
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbDownwardDiagonal"
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbUpwardDiagonal"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbVerticalLine"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbHorizntalLine"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbFSTransparent"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbFSSolid"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmFills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Draw boxes using different DrawStyle values.
Private Sub Form_Load()
Const GAP = 120
Dim i As Single
Dim wid As Single
Dim hgt As Single
Dim X As Single
Dim Y As Single

    ' Make changes permanent.
    AutoRedraw = True

    ' Draw the boxes.
    wid = lblStyle(0).Width
    hgt = 5 * 120
    For i = 0 To lblStyle.UBound
        FillStyle = i
        X = lblStyle(i).Left
        Y = lblStyle(i).Top + lblStyle(i).Height
        Line (X, Y)-Step(wid, hgt), , B
    Next i
End Sub
