VERSION 5.00
Begin VB.Form frmBkMode 
   Caption         =   "BkMode"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OPAQUE"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TRANSPARENT"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmBkMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
Private Sub Form_Load()
Dim i As Integer
Dim j As Integer
Dim start_j As Long
Dim gray As Long

    AutoRedraw = True
    FillStyle = vbDiagonalCross
    ScaleMode = vbPixels

    ' Create a checkerboard background.
    gray = RGB(128, 128, 128)
    For i = Label1(0).Height To ScaleHeight - 1 Step 20
        For j = start_j To ScaleWidth - 1 Step 40
            Line (j, i)-Step(20, 20), gray, BF
        Next j
        start_j = 20 - start_j
    Next i

    ' Draw an ellipse with BkMode = TRANSPARENT.
    SetBkMode hdc, TRANSPARENT
    Ellipse hdc, 10, 25, ScaleWidth / 2 - 5, ScaleHeight - 20

    ' Draw an ellipse with BkMode = OPAQUE.
    SetBkMode hdc, OPAQUE
    SetBkColor hdc, vbWhite
    Ellipse hdc, ScaleWidth / 2 + 5, 25, ScaleWidth - 10, ScaleHeight - 20
End Sub

