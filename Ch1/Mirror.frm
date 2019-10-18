VERSION 5.00
Begin VB.Form MirrorForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Mirror"
   ClientHeight    =   3705
   ClientLeft      =   3705
   ClientTop       =   450
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   3420
   Begin VB.CommandButton cmdMirror 
      Caption         =   "Mirror"
      Default         =   -1  'True
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picPSetDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   0
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   12
      Top             =   2280
      Width           =   1020
   End
   Begin VB.PictureBox picPaintDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   2400
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   11
      Top             =   2280
      Width           =   1020
   End
   Begin VB.PictureBox picRefreshDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   1200
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   10
      Top             =   2280
      Width           =   1020
   End
   Begin VB.PictureBox picRefreshSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   1200
      Picture         =   "Mirror.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   1200
      Width           =   1020
   End
   Begin VB.PictureBox picPaintSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   2400
      Picture         =   "Mirror.frx":0882
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   1200
      Width           =   1020
   End
   Begin VB.PictureBox picPSetSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   0
      Picture         =   "Mirror.frx":1104
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Point/PSet W/refresh"
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   9
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label lblRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   3360
      Width           =   1020
   End
   Begin VB.Label lblPaint 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   3360
      Width           =   1020
   End
   Begin VB.Label lblPSet 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3360
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PaintPicture"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Point/PSet"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1020
   End
End
Attribute VB_Name = "MirrorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Mirror the images.
Private Sub cmdMirror_Click()
Dim X As Single
Dim Y As Single
Dim wid As Single
Dim hgt As Single
Dim start_time As Single
Dim stop_time As Single

    ' Prepare.
    cmdMirror.Enabled = False
    MousePointer = vbHourglass
    DoEvents
    
    ' Clear any previous results.
    picPSetDest.Line (0, 0)-(picPSetDest.ScaleWidth, picPSetDest.ScaleHeight), picPSetDest.BackColor, BF
    picRefreshDest.Line (0, 0)-(picPSetDest.ScaleWidth, picPSetDest.ScaleHeight), picPSetDest.BackColor, BF
    picPaintDest.Line (0, 0)-(picPSetDest.ScaleWidth, picPSetDest.ScaleHeight), picPSetDest.BackColor, BF
    lblPSet.Caption = ""
    lblRefresh.Caption = ""
    lblPaint.Caption = ""
    Refresh

    ' Mirror using PSet/Paint.
    start_time = Timer
    wid = picPSetSource.ScaleWidth
    For X = 0 To wid
        For Y = 0 To picPSetSource.ScaleHeight
            picPSetDest.PSet (wid - X, Y), picPSetSource.Point(X, Y)
        Next Y
    Next X
    stop_time = Timer
    lblPSet.Caption = Format$(stop_time - start_time, "0.000") & " sec"
    Refresh
    
    ' Mirror using PSet/Paint with refresh.
    start_time = Timer
    wid = picRefreshSource.ScaleWidth
    For X = 0 To wid
        For Y = 0 To picRefreshSource.ScaleHeight
            picRefreshDest.PSet (wid - X, Y), picRefreshSource.Point(X, Y)
        Next Y
        picRefreshDest.Refresh
    Next X
    stop_time = Timer
    lblRefresh.Caption = Format$(stop_time - start_time, "0.000") & " sec"
    Refresh

    ' Mirror using PaintPicture.
    start_time = Timer
    wid = picPaintSource.ScaleWidth
    hgt = picPaintSource.ScaleHeight
    picPaintDest.PaintPicture picPaintSource.Picture, _
        0, 0, wid + 1, hgt, _
        wid + 1, 0, -(wid + 1), hgt, vbSrcCopy
    stop_time = Timer
    lblPaint.Caption = Format$(stop_time - start_time, "0.000") & " sec"

    MousePointer = vbDefault
    cmdMirror.Enabled = True
End Sub
' End if the program is in the middle of drawing.
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
