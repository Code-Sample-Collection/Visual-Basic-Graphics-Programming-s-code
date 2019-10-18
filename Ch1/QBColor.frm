VERSION 5.00
Begin VB.Form frmQBColor 
   Caption         =   "QBColor"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   15
      Left            =   4080
      TabIndex        =   31
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   30
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   29
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   14
      Left            =   4440
      TabIndex        =   28
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   27
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   13
      Left            =   4440
      TabIndex        =   26
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   25
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   12
      Left            =   4440
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   23
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   3120
      TabIndex        =   22
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   21
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   20
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   19
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   18
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   15
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   14
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   13
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmQBColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Display the colors and their names.
Private Sub Form_Load()
Dim i As Integer

    For i = 0 To 15
        lblColor(i).BackColor = QBColor(i)
        lblName(i).Caption = Format$(i)
    Next i
End Sub
