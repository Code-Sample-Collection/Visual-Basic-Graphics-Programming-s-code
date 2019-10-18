VERSION 5.00
Begin VB.Form frmTransmit 
   Caption         =   "Transmit"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtN2 
      Height          =   285
      Left            =   4380
      TabIndex        =   4
      Text            =   "1.5"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtN1 
      Height          =   285
      Left            =   4380
      TabIndex        =   2
      Text            =   "1.0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "n2/n1"
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   12
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Sin1/Sin2"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   11
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblSin1Sin2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4380
      TabIndex        =   10
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblN2N1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4380
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblSin2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4380
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblSin1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4380
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Sin1"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Sin2"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "n2"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "n1"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmTransmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LINE_SCALE = 100
Private Ymid As Single
Private Xmid As Single

Private Sub Form_Load()
    picCanvas.ScaleMode = vbPixels
    Ymid = picCanvas.ScaleHeight / 2
    Xmid = picCanvas.ScaleWidth / 2
    picCanvas.Line (0, Ymid)-Step(picCanvas.ScaleWidth, 0)
    picCanvas.Line (Xmid, Ymid)-Step(0, -LINE_SCALE)
    picCanvas.Picture = picCanvas.Image
End Sub


Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const PI = 3.14159265

Dim nx As Single
Dim ny As Single
Dim n1 As Single
Dim n2 As Single

Dim lx As Single
Dim ly As Single
Dim ltx As Single
Dim lty As Single
Dim dist As Single

Dim LdotN As Single
Dim cos1 As Single
Dim cos2 As Single
Dim n1_over_n2 As Single
Dim normal_factor As Single

Dim theta1 As Single
Dim theta2 As Single

    picCanvas.Cls
    n1 = CSng(txtN1.Text)
    n2 = CSng(txtN2.Text)

    ' Get the normal vector.
    nx = 0
    ny = -1

    ' Get the light vector.
    lx = X - Xmid
    ly = Y - Ymid
    dist = Sqr(lx * lx + ly * ly)
    lx = lx / dist
    ly = ly / dist

    picCanvas.Line (Xmid, Ymid)-Step(LINE_SCALE * lx, LINE_SCALE * ly), vbBlue

    ' Calculate the transmission vector LT.
    LdotN = lx * nx + ly * ny
    cos1 = Abs(LdotN)
    n1_over_n2 = n1 / n2

    cos2 = 1 - (1 - cos1 * cos1) * n1_over_n2 * n1_over_n2
    If cos2 < 0 Then
        ' Reflection.
    Else
        cos2 = Sqr(cos2)
        ' Note that the incident vector I = -L.
        normal_factor = cos2 - n1_over_n2 * cos1
        ltx = -n1_over_n2 * lx - normal_factor * nx
        lty = -n1_over_n2 * ly - normal_factor * ny
    End If

    picCanvas.Line (Xmid, Ymid)-Step(LINE_SCALE * ltx, LINE_SCALE * lty), vbYellow

    theta1 = PI / 2 - Atn(ly / lx)
    lblSin1.Caption = Format$(Sin(theta1))
    theta2 = PI / 2 - Atn(lty / ltx)
    lblSin2.Caption = Format$(Sin(theta2))
    lblSin1Sin2.Caption = Format$(Sin(theta1) / Sin(theta2))
    lblN2N1.Caption = Format$(n2 / n1)
End Sub
