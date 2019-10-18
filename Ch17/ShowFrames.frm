VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Scale"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.OptionButton optScale 
         Caption         =   "Small"
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optScale 
         Caption         =   "Medium"
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optScale 
         Caption         =   "Large"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5280
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   720
      Width           =   7635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_FileBase As String
Private Sub Form_Load()
    optScale(0).Value = True
End Sub


Private Sub optScale_Click(Index As Integer)
    m_FileBase = "C:\Temp\VBGamer" & Index & "_"
End Sub


Private Sub Timer1_Timer()
Static frame As Integer

    Picture1.Picture = LoadPicture( _
        m_FileBase & Format$(frame, "00") & ".bmp")

    frame = frame + 1
    If frame > 49 Then frame = 0
End Sub


