VERSION 5.00
Begin VB.Form frmPickover 
   Caption         =   "Pickover"
   ClientHeight    =   5430
   ClientLeft      =   1800
   ClientTop       =   705
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5430
   ScaleWidth      =   6375
   Begin VB.Frame Frame1 
      Caption         =   "Projection"
      Height          =   1335
      Left            =   0
      TabIndex        =   19
      Top             =   3120
      Width           =   930
      Begin VB.OptionButton optPlane 
         Caption         =   "YZ"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton optPlane 
         Caption         =   "XZ"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optPlane 
         Caption         =   "XY"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.TextBox txtZ0 
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Text            =   "0"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtY0 
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Text            =   "0"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtX0 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtE 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Text            =   "1.0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Text            =   "-2.5"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtC 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "-0.6"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtB 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "0.5"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtA 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "2.0"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   615
   End
   Begin VB.PictureBox picCanvas 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   5415
      Left            =   960
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Z0"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   17
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Y0"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   15
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X0"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmPickover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PlaneTypes
    plane_XY = 0
    plane_XZ = 1
    plane_YZ = 2
End Enum
Private SelectedPlane As PlaneTypes

Private Running As Boolean

Private A As Single
Private B As Single
Private C As Single
Private D As Single
Private E As Single
Private X0 As Single
Private Y0 As Single
Private Z0 As Single
' Draw the curve.
Private Sub DrawCurve()
Const XMIN = -2.1
Const XMAX = 2.1
Const YMIN = -2.1
Const YMAX = 2.1
Const ZMIN = -1.2
Const ZMAX = 1.2

Dim wid As Single
Dim hgt As Single
Dim xoff As Single
Dim yoff As Single
Dim zoff As Single
Dim xscale As Single
Dim yscale As Single
Dim zscale As Single
Dim x As Single
Dim y As Single
Dim z As Single
Dim x2 As Single
Dim y2 As Single
Dim i As Integer

    ' See how much room we have.
    wid = picCanvas.ScaleWidth
    hgt = picCanvas.ScaleHeight

    Select Case SelectedPlane
        Case plane_XY
            xoff = wid / 2
            yoff = hgt / 2
            xscale = wid / (XMAX - XMIN)
            yscale = hgt / (YMAX - YMIN)
        Case plane_XZ
            xoff = wid / 2
            zoff = hgt / 2
            xscale = wid / (XMAX - XMIN)
            zscale = hgt / (ZMAX - ZMIN)
        Case plane_YZ
            yoff = wid / 2
            zoff = hgt / 2
            yscale = wid / (YMAX - YMIN)
            zscale = hgt / (ZMAX - ZMIN)
    End Select

    ' Get the drawing parameters.
    GetParameters

    ' Compute the values.
    x = X0
    y = Y0
    z = Z0
    i = 0
    Do While Running
        ' Move to the next point.
        x2 = Sin(A * y) - z * Cos(B * x)
        y2 = z * Sin(C * x) - Cos(D * y)
        z = Sin(x)
        x = x2
        y = y2

        ' Plot the point.
        Select Case SelectedPlane
            Case plane_XY
                picCanvas.PSet (x * xscale + xoff, y * yscale + yoff)
            Case plane_XZ
                picCanvas.PSet (x * xscale + xoff, z * zscale + zoff)
            Case plane_YZ
                picCanvas.PSet (y * yscale + yoff, z * zscale + zoff)
        End Select

        ' To make things faster, only DoEvents
        ' every 100 times.
        i = i + 1
        If i > 100 Then
            i = 0
            DoEvents
        End If
    Loop
End Sub

' Get the curve parameters.
Private Sub GetParameters()
    If Not IsNumeric(txtA.Text) Then txtA.Text = "2"
    If Not IsNumeric(txtB.Text) Then txtB.Text = ".5"
    If Not IsNumeric(txtC.Text) Then txtC.Text = "-.6"
    If Not IsNumeric(txtD.Text) Then txtD.Text = "-2.5"
    If Not IsNumeric(txtE.Text) Then txtE.Text = "1"
    If Not IsNumeric(txtX0.Text) Then txtX0.Text = "0"
    If Not IsNumeric(txtY0.Text) Then txtY0.Text = "0"
    If Not IsNumeric(txtZ0.Text) Then txtZ0.Text = "0"

    A = CSng(txtA.Text)
    B = CSng(txtB.Text)
    C = CSng(txtC.Text)
    D = CSng(txtD.Text)
    E = CSng(txtE.Text)
    X0 = CSng(txtX0.Text)
    Y0 = CSng(txtY0.Text)
    Z0 = CSng(txtZ0.Text)
End Sub
' Erase the picCanvas.
Private Sub CmdClear_Click()
    picCanvas.Cls
End Sub

Private Sub cmdGo_Click()
    If Running Then
        Running = False
        cmdGo.Caption = "Stopped"
    Else
        Running = True
        cmdGo.Caption = "Stop"
        DrawCurve
        cmdGo.Caption = "Go"
    End If
End Sub


Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight - 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub optPlane_Click(Index As Integer)
    SelectedPlane = Index
End Sub


