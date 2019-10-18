VERSION 5.00
Begin VB.Form frmLineTo 
   Caption         =   "LineTo"
   ClientHeight    =   3585
   ClientLeft      =   1620
   ClientTop       =   1425
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3585
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
   Begin VB.PictureBox picLineTo 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   3000
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox picLine 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   120
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "MoveToEx and LineTo"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Line Method"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmLineTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' We declare the last argument to MoveToEx as Any so
' we can pass vbNullString to not get a return value.
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

' Draw the lines.
Private Sub cmdDraw_Click()
Dim start_time As Single
Dim stop_time As Single
Dim x As Long
Dim y As Long
Dim hdc As Long

    picLine.Cls
    picLineTo.Cls
    MousePointer = vbHourglass
    DoEvents

    start_time = Timer()
    With picLine
        For y = 0 To .ScaleHeight Step 4
            .CurrentX = .ScaleLeft
            .CurrentY = .ScaleTop + y - 1
            For x = 0 To .ScaleWidth Step 4
                picLine.Line -Step(2, 0)
                picLine.Line -Step(0, 2)
                picLine.Line -Step(2, 0)
                picLine.Line -Step(0, -2)
            Next x
        Next y
    End With
    picLine.Refresh
    stop_time = Timer()
    Label1.Caption = Format$(stop_time - start_time, "0.0000")
    DoEvents

    start_time = Timer()
    For y = 0 To picLineTo.ScaleHeight Step 4
        MoveToEx picLineTo.hdc, 0, y, vbNullString
        For x = 0 To picLineTo.ScaleWidth Step 4
            LineTo picLineTo.hdc, x + 2, y
            LineTo picLineTo.hdc, x + 2, y + 2
            LineTo picLineTo.hdc, x + 4, y + 2
            LineTo picLineTo.hdc, x + 4, y
        Next x
    Next y
    picLineTo.Refresh
    stop_time = Timer()
    Label2.Caption = Format$(stop_time - start_time, "0.0000")

    MousePointer = vbDefault
End Sub
