VERSION 5.00
Begin VB.Form frmShowForm 
   Caption         =   "ShowFont"
   ClientHeight    =   4020
   ClientLeft      =   2115
   ClientTop       =   1215
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   6525
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Text            =   "12.0"
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox chkUnderline 
      Caption         =   "Underline"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CheckBox chkStrikeThrough 
      Caption         =   "StrikeThrough"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtSample 
      Height          =   3735
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.ListBox lstScreenFonts 
      Height          =   3765
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Size"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sample"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fonts"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmShowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Change the sample text's font.
Private Sub ShowSample()
    With txtSample.Font
        .Name = lstScreenFonts.List(lstScreenFonts.ListIndex)
        .Bold = (chkBold.Value = vbChecked)
        .Italic = (chkItalic.Value = vbChecked)
        .Strikethrough = (chkStrikeThrough.Value = vbChecked)
        .Underline = (chkUnderline.Value = vbChecked)
        .Size = CSng(txtSize.Text)
    End With
End Sub

Private Sub chkBold_Click()
    ShowSample
End Sub

Private Sub chkItalic_Click()
    ShowSample
End Sub


Private Sub chkStrikeThrough_Click()
    ShowSample
End Sub


Private Sub chkUnderline_Click()
    ShowSample
End Sub


Private Sub Form_Load()
Dim i As Integer

    ' Fill the font list with font names.
    For i = 0 To Screen.FontCount - 1
        lstScreenFonts.AddItem Screen.Fonts(i)
    Next i
    lstScreenFonts.ListIndex = 0
    lstScreenFonts.Selected(0) = True

    ' Fill in the sample text.
    txtSample.Text = "ABCDE" & vbCrLf & _
        "FGHIJ" & vbCrLf & "KLMNO" & vbCrLf & _
        "PQRST" & vbCrLf & "UVWXYZ" & vbCrLf & _
        "abcde" & vbCrLf & "fghij" & vbCrLf & _
        "klmno" & vbCrLf & "pqrst" & vbCrLf & _
        "uvwxyz" & vbCrLf & "12345" & vbCrLf & _
        "67890"
End Sub


' Change the sample label's font.
Private Sub lstScreenFonts_Click()
    ShowSample
End Sub

Private Sub txtSize_Change()
    If IsNumeric(txtSize.Text) Then ShowSample
End Sub


