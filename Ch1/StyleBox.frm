VERSION 5.00
Begin VB.Form frmStyleBox 
   BackColor       =   &H00FFFFFF&
   Caption         =   "StyleBox"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbInsideSolid"
      Height          =   255
      Index           =   6
      Left            =   2580
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbInvisible"
      Height          =   255
      Index           =   5
      Left            =   1020
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbDashDotDot"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbDashDot"
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbDot"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbDash"
      Height          =   255
      Index           =   1
      Left            =   2580
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vbSolid"
      Height          =   255
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmStyleBox"
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
    wid = lblStyle(0).Width + 2 * GAP
    hgt = lblStyle(0).Height + 2 * GAP
    For i = 0 To lblStyle.UBound
        DrawStyle = i
        X = lblStyle(i).Left - GAP
        Y = lblStyle(i).Top - GAP
        Line (X, Y)-Step(wid, hgt), , B
    Next i
End Sub
