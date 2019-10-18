VERSION 5.00
Begin VB.Form frmClipCtls 
   Caption         =   "ClipCtls"
   ClientHeight    =   4515
   ClientLeft      =   600
   ClientTop       =   1365
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4515
   ScaleWidth      =   8205
   Begin VB.PictureBox PaintPict 
      ClipControls    =   0   'False
      Height          =   3855
      Index           =   2
      Left            =   5520
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   6
      Top             =   240
      Width           =   2655
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   240
         TabIndex        =   28
         Text            =   "TextBox"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   1440
         TabIndex        =   27
         Text            =   "TextBox"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   240
         TabIndex        =   26
         Text            =   "TextBox"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   1440
         TabIndex        =   25
         Text            =   "TextBox"
         Top             =   2040
         Width           =   975
      End
      Begin VB.Image imgObscures 
         Height          =   960
         Index           =   4
         Left            =   240
         Picture         =   "ClipCtls.frx":0000
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   1440
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   1440
         TabIndex        =   29
         Top             =   840
         Width           =   975
      End
      Begin VB.Image imgObscures 
         Height          =   960
         Index           =   5
         Left            =   1440
         Picture         =   "ClipCtls.frx":0882
         Top             =   2640
         Width           =   960
      End
   End
   Begin VB.PictureBox PaintPict 
      ClipControls    =   0   'False
      Height          =   3855
      Index           =   1
      Left            =   2760
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   4
      Top             =   240
      Width           =   2655
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   1440
         TabIndex        =   24
         Text            =   "TextBox"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Text            =   "TextBox"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   1440
         TabIndex        =   22
         Text            =   "TextBox"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Text            =   "TextBox"
         Top             =   1440
         Width           =   975
      End
      Begin VB.Image imgObscures 
         Height          =   960
         Index           =   3
         Left            =   1440
         Picture         =   "ClipCtls.frx":1104
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   1440
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgObscures 
         Height          =   960
         Index           =   2
         Left            =   240
         Picture         =   "ClipCtls.frx":1986
         Top             =   2640
         Width           =   960
      End
   End
   Begin VB.PictureBox PaintPict 
      Height          =   3855
      Index           =   0
      Left            =   0
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   1440
         TabIndex        =   16
         Text            =   "TextBox"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Text            =   "TextBox"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   14
         Text            =   "TextBox"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtObscures 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Text            =   "TextBox"
         Top             =   1440
         Width           =   975
      End
      Begin VB.Image imgObscures 
         Height          =   960
         Index           =   1
         Left            =   1440
         Picture         =   "ClipCtls.frx":2208
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   1440
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgObscures 
         Height          =   960
         Index           =   0
         Left            =   240
         Picture         =   "ClipCtls.frx":2A8A
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label lblObscuring 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   10
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ClipControls = False (manual refresh)"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ClipControls = False"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ClipControls = True"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmClipCtls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Draw a bunch of squiggly lines.
Private Sub DrawPict(pic As PictureBox)
Const Amp = 3
Const PI = 3.14159
Const Per = 4 * PI

Dim i As Single
Dim j As Single
Dim hgt As Single
Dim wid As Single
    
    pic.ScaleMode = 3   ' Pixel.
    
    pic.Cls     ' Clear the picture box.
    
    For i = 0 To pic.ScaleHeight Step 4
        pic.CurrentX = 0
        pic.CurrentY = i
        For j = 0 To pic.ScaleWidth
            pic.Line -(j, i + Amp * Sin(j / Per))
        Next j
    Next i
    For i = 1 To hgt Step 2
        pic.Line (0, i)-(wid, i)
    Next i
End Sub

' Redraw this PictureBox.
Private Sub PaintPict_Paint(Index As Integer)
Dim start_time As Single
Dim i As Integer

    start_time = Timer
    DrawPict PaintPict(Index)

    ' Manually refresh txtObscures(8) through txtObscures(11).
    If Index = 2 Then
        For i = 8 To 11
            txtObscures(i).Refresh
        Next i
    End If

    lblTime(Index).Caption = _
        Format$(Timer - start_time, "0.00") & _
        " seconds"
End Sub
