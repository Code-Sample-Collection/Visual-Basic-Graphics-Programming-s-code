VERSION 5.00
Begin VB.Form frmHiRes 
   AutoRedraw      =   -1  'True
   Caption         =   "HiRes"
   ClientHeight    =   4920
   ClientLeft      =   1035
   ClientTop       =   1050
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4920
   ScaleWidth      =   5865
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "HiRes.frx":0000
      Left            =   2400
      List            =   "HiRes.frx":000D
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Wrapped Caption"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3720
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2640
      TabIndex        =   21
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "A long caption"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      TabIndex        =   15
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "A long caption"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "A long caption"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "HiRes.frx":0028
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   8
      Text            =   "HiRes.frx":003D
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "HiRes.frx":006B
      Top             =   2220
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "HiRes.frx":0099
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Text            =   "The quick brown fox jumped over the lazy dog."
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      ItemData        =   "HiRes.frx":00C7
      Left            =   120
      List            =   "HiRes.frx":00EC
      MultiSelect     =   1  'Simple
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   1575
      Begin VB.CheckBox Check1 
         Caption         =   "A long caption"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   2040
      Picture         =   "HiRes.frx":013F
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Right justified"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   23
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   1
      Left            =   4200
      Picture         =   "HiRes.frx":16C5
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   4800
      X2              =   5760
      Y1              =   3720
      Y2              =   3240
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Index           =   5
      Left            =   240
      Shape           =   5  'Rounded Square
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Index           =   4
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   735
      Index           =   2
      Left            =   4800
      Shape           =   2  'Oval
      Top             =   2400
      Width           =   975
   End
   Begin VB.Shape Shape1 
      FillStyle       =   7  'Diagonal Cross
      Height          =   1935
      Index           =   1
      Left            =   4920
      Shape           =   1  'Square
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   2175
      Index           =   0
      Left            =   4800
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   0
      Left            =   3120
      Picture         =   "HiRes.frx":2C4B
      Top             =   3840
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4800
      X2              =   5760
      Y1              =   3240
      Y2              =   3720
   End
   Begin VB.Label Label1 
      Caption         =   "Left justified"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePreview 
         Caption         =   "Print Preview..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilePrintForm 
         Caption         =   "&PrintForm"
      End
      Begin VB.Menu mnuFileHiResPrint 
         Caption         =   "&HiResPrint"
      End
      Begin VB.Menu mnuFileLargePrint 
         Caption         =   "HiResPrint &Large Scale"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmHiRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Make the form have a picture.
    Picture = Image

    ' Set default selections.
    Combo1.ListIndex = 1
    List1.Selected(1) = True
    List1.Selected(3) = True
    
    ' Give the form the same palette as the pictures.
    Picture.hPal = Picture1(0).Picture.hPal
End Sub

' ************************************************
' Unload the form.
' ************************************************
Private Sub mnuFileExit_Click()
    Unload Me
End Sub


' ************************************************
' Print using the HiResPrint subroutine.
' ************************************************
Private Sub mnuFileHiResPrint_Click()
    MousePointer = vbHourglass
    DoEvents
    HiResPrint Me, Printer, hires_Normal
    Printer.EndDoc
    MousePointer = vbDefault
End Sub
' ************************************************
' Print at large scale using HiResPrint.
' ************************************************
Private Sub mnuFileLargePrint_Click()
    MousePointer = vbHourglass
    DoEvents
    HiResPrint Me, Printer, hires_StretchToFit
    Printer.EndDoc
    MousePointer = vbDefault
End Sub
' Display a print preview.
Private Sub mnuFilePreview_Click()
    frmHiResPrv.ShowPreview Me
End Sub

' ************************************************
' Print using the PrintForm method.
' ************************************************
Private Sub mnuFilePrintForm_Click()
    PrintForm
End Sub

