VERSION 5.00
Begin VB.Form frmDistance 
   Caption         =   "Distance"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLineToLine 
      Caption         =   "Line P1/V1 to Line P2/V2"
      Height          =   495
      Left            =   2880
      TabIndex        =   24
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdP2ToPlane 
      Caption         =   "P2 to Plane P1/V1"
      Height          =   495
      Left            =   2880
      TabIndex        =   22
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdP2toLine 
      Caption         =   "P2 to Line P1/V1"
      Height          =   495
      Left            =   2880
      TabIndex        =   20
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdP1toP2 
      Caption         =   "P1 to P2"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vectors"
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
      Begin VB.TextBox txtV2z 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtV2y 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtV2x 
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtV1z 
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtV1y 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtV1x 
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "V2"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "V1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Points"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.TextBox txtP2z 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtP2y 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtP2x 
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtP1z 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtP1y 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtP1x 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "P2"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "P1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Label lblLineToLine 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblP2ToPlane 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblP2ToLine 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblP1toP2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmDistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Shoe the distance between two lines.
Private Sub cmdLineToLine_Click()
Dim px1 As Single
Dim py1 As Single
Dim pz1 As Single
Dim px2 As Single
Dim py2 As Single
Dim pz2 As Single
Dim vx1 As Single
Dim vy1 As Single
Dim vz1 As Single
Dim vx2 As Single
Dim vy2 As Single
Dim vz2 As Single

    On Error GoTo LineToLineError
    px1 = CSng(txtP1x.Text)
    py1 = CSng(txtP1y.Text)
    pz1 = CSng(txtP1z.Text)
    px2 = CSng(txtP2x.Text)
    py2 = CSng(txtP2y.Text)
    pz2 = CSng(txtP2z.Text)
    vx1 = CSng(txtV1x.Text)
    vy1 = CSng(txtV1y.Text)
    vz1 = CSng(txtV1z.Text)
    vx2 = CSng(txtV2x.Text)
    vy2 = CSng(txtV2y.Text)
    vz2 = CSng(txtV2z.Text)

    lblLineToLine.Caption = Format$( _
        DistanceLineToLine( _
            px1, py1, pz1, _
            px2, py2, pz2, _
            vx1, vy1, vz1, _
            vx2, vy2, vz2), _
            "0.0000")
    Exit Sub

LineToLineError:
    MsgBox Err.Description
End Sub

' Show the distance between point P2 and the
' line through P1 in direction V1.
Private Sub cmdP2toLine_Click()
Dim px1 As Single
Dim py1 As Single
Dim pz1 As Single
Dim px2 As Single
Dim py2 As Single
Dim pz2 As Single
Dim vx1 As Single
Dim vy1 As Single
Dim vz1 As Single

    On Error GoTo P2ToLineError
    px1 = CSng(txtP1x.Text)
    py1 = CSng(txtP1y.Text)
    pz1 = CSng(txtP1z.Text)
    px2 = CSng(txtP2x.Text)
    py2 = CSng(txtP2y.Text)
    pz2 = CSng(txtP2z.Text)
    vx1 = CSng(txtV1x.Text)
    vy1 = CSng(txtV1y.Text)
    vz1 = CSng(txtV1z.Text)

    lblP2ToLine.Caption = Format$( _
        DistancePointToLine( _
            px2, py2, pz2, _
            px1, py1, pz1, _
            vx1, vy1, vz1), _
            "0.0000")
    Exit Sub

P2ToLineError:
    MsgBox Err.Description
End Sub

' Show the distance between P1 and P2.
Private Sub cmdP1toP2_Click()
Dim px1 As Single
Dim py1 As Single
Dim pz1 As Single
Dim px2 As Single
Dim py2 As Single
Dim pz2 As Single

    On Error GoTo P1toP2Error
    px1 = CSng(txtP1x.Text)
    py1 = CSng(txtP1y.Text)
    pz1 = CSng(txtP1z.Text)
    px2 = CSng(txtP2x.Text)
    py2 = CSng(txtP2y.Text)
    pz2 = CSng(txtP2z.Text)

    lblP1toP2.Caption = Format$( _
        DistancePointToPoint( _
            px1, py1, pz1, _
            px2, py2, pz2), _
            "0.0000")
    Exit Sub

P1toP2Error:
    MsgBox Err.Description
End Sub

' Show the distance between a point and a plane.
Private Sub cmdP2ToPlane_Click()
Dim px1 As Single
Dim py1 As Single
Dim pz1 As Single
Dim px2 As Single
Dim py2 As Single
Dim pz2 As Single
Dim vx1 As Single
Dim vy1 As Single
Dim vz1 As Single

    On Error GoTo P2ToPlaneError
    px1 = CSng(txtP1x.Text)
    py1 = CSng(txtP1y.Text)
    pz1 = CSng(txtP1z.Text)
    px2 = CSng(txtP2x.Text)
    py2 = CSng(txtP2y.Text)
    pz2 = CSng(txtP2z.Text)
    vx1 = CSng(txtV1x.Text)
    vy1 = CSng(txtV1y.Text)
    vz1 = CSng(txtV1z.Text)

    lblP2ToPlane.Caption = Format$( _
        DistancePointToPlane( _
            px2, py2, pz2, _
            px1, py1, pz1, _
            vx1, vy1, vz1), _
            "0.0000")
    Exit Sub

P2ToPlaneError:
    MsgBox Err.Description
End Sub


