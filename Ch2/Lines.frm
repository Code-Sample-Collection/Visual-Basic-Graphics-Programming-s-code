VERSION 5.00
Begin VB.Form frmLines 
   Caption         =   "Lines"
   ClientHeight    =   4125
   ClientLeft      =   1140
   ClientTop       =   1530
   ClientWidth     =   6855
   LinkTopic       =   "LineForm"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   6855
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3720
      Width           =   615
   End
   Begin VB.PictureBox picLine 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   0
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.PictureBox picPolyline 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   3480
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Line"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Polyline"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label lblLine 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblPolyline 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "frmLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Points() As POINTAPI
Private NumPoints As Integer
' Draw the lines 100 times.
Private Sub cmdDraw_Click()
Const NUM_TRIALS = 100

Dim start_time As Single
Dim stop_time As Single
Dim i As Integer
Dim trial As Integer

    picPolyline.Cls
    picLine.Cls
    lblPolyline.Caption = ""
    lblLine.Caption = ""
    MousePointer = vbHourglass
    DoEvents
    
    start_time = Timer()
    For trial = 1 To NUM_TRIALS
        picLine.CurrentX = Points(1).x
        picLine.CurrentY = Points(1).y
        For i = 2 To NumPoints
            picLine.Line -(Points(i).x, Points(i).y)
        Next i
    Next trial
    stop_time = Timer()
    picLine.Refresh
    lblLine.Caption = Format$(stop_time - start_time, "0.0000")
    DoEvents

    start_time = Timer()
    For trial = 1 To NUM_TRIALS
        If Polyline(picPolyline.hdc, Points(1), NumPoints) = 0 Then Exit Sub
    Next trial
    stop_time = Timer()
    picPolyline.Refresh
    lblPolyline.Caption = Format$(stop_time - start_time, "0.0000")

    MousePointer = vbDefault
End Sub

' Create the points for the lines.
Private Sub Form_Load()
Dim i As Integer
Dim small As Boolean
Dim half As Integer
Dim hgt As Integer
Dim wid As Integer

    NumPoints = picPolyline.ScaleWidth
    ReDim Points(1 To NumPoints)
    half = NumPoints \ 2
    
    hgt = picPolyline.ScaleHeight
    For i = 1 To half
        Points(i).x = 2 * i
        If small Then
            Points(i).y = 0
        Else
            Points(i).y = hgt
        End If
        small = Not small
    Next i
    
    wid = picPolyline.ScaleWidth
    For i = half + 1 To NumPoints
        Points(i).y = 2 * (i - half)
        If small Then
            Points(i).x = 0
        Else
            Points(i).x = wid
        End If
        small = Not small
    Next i
End Sub
