VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmRich 
   Caption         =   "Rich"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCombine 
      Caption         =   "Combine"
      Height          =   495
      Left            =   3900
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Index           =   1
      Left            =   4560
      TabIndex        =   6
      Top             =   -60
      Width           =   4455
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "Italic"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Underline"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkStrikethru 
         Caption         =   "Strikethru"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin RichTextLib.RichTextBox rchInput 
         Height          =   1455
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
         _Version        =   393217
         TextRTF         =   $"Rich.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   4455
      Begin VB.CheckBox chkStrikethru 
         Caption         =   "&Strikethru"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "&Underline"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "&Italic"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "&Bold"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin RichTextLib.RichTextBox rchInput 
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
         _Version        =   393217
         TextRTF         =   $"Rich.frx":0133
      End
   End
   Begin RichTextLib.RichTextBox rchOutput 
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"Rich.frx":0266
   End
   Begin RichTextLib.RichTextBox rchOutput 
      Height          =   1455
      Index           =   1
      Left            =   4680
      TabIndex        =   14
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"Rich.frx":032F
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SelRTF Properties"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   16
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SelText Properties"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   4215
   End
End
Attribute VB_Name = "frmRich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IgnoreChanges As Boolean

' Set a check box's tri-state value.
Private Sub SetCheck(ByVal chk As CheckBox, ByVal value As Variant)
    IgnoreChanges = True
    If IsNull(value) Then
        chk.value = vbGrayed
    ElseIf value Then
        chk.value = vbChecked
    Else
        chk.value = vbUnchecked
    End If
    IgnoreChanges = False
End Sub

Private Sub chkBold_Click(Index As Integer)
    If IgnoreChanges Then Exit Sub

    rchInput(Index).SelBold = (chkBold(Index).value = vbChecked)
End Sub


' Set the selection's style.
Private Sub chkItalic_Click(Index As Integer)
    If IgnoreChanges Then Exit Sub

    rchInput(Index).SelItalic = (chkItalic(Index).value = vbChecked)
    rchInput(Index).SetFocus
End Sub

Private Sub chkStrikethru_Click(Index As Integer)
    If IgnoreChanges Then Exit Sub

    rchInput(Index).SelStrikeThru = (chkStrikethru(Index).value = vbChecked)
End Sub


Private Sub chkunderline_Click(Index As Integer)
    If IgnoreChanges Then Exit Sub

    rchInput(Index).SelUnderline = (chkUnderline(Index).value = vbChecked)
End Sub


' Combine the values in the rchInput controls.
Private Sub cmdCombine_Click()
Dim sel_start(0 To 1) As Integer
Dim sel_length(0 To 1) As Integer
Dim i As Integer

    ' Combine text only.
    rchOutput(0).Text = rchInput(0).Text & _
        vbCrLf & rchInput(1).Text

    ' Save the current SelStart and SelLength values.
    For i = 0 To 1
        sel_start(i) = rchInput(i).SelStart
        sel_length(i) = rchInput(i).SelLength
        rchInput(i).SelStart = 0
        rchInput(i).SelLength = Len(rchInput(i).Text)
    Next i

    ' Combine the text only.
    rchOutput(0).Text = rchInput(0).Text & _
        vbCrLf & rchInput(1).Text

    ' Combine the rich text values.
    ' Copy the first control's text with RTF codes.
    rchOutput(1).SelStart = 0
    rchOutput(1).SelLength = Len(rchOutput(1).Text)
    rchOutput(1).SelRTF = rchInput(0).SelRTF

    ' Add vbCrLf to the end.
    rchOutput(1).SelStart = Len(rchOutput(1).Text)
    rchOutput(1).SelLength = 0
    rchOutput(1).SelText = vbCrLf

    ' Add the second control's text with RTF codes.
    rchOutput(1).SelStart = Len(rchOutput(1).Text)
    rchOutput(1).SelLength = 0
    rchOutput(1).SelRTF = rchInput(1).SelRTF

    ' Restore the SetStart and SelLength values.
    For i = 0 To 1
        rchInput(i).SelStart = sel_start(i)
        rchInput(i).SelLength = sel_length(i)
    Next i
    rchOutput(1).SelLength = 0
End Sub

' Prepare the program.
Private Sub Form_Load()
    ' Set the form's width.
    Width = Frame1(1).Left + Frame1(1).Width + Width - ScaleWidth

    ' Set some text properties.
    rchInput(0).SelStart = 0
    rchInput(0).SelLength = Len(rchInput(0).Text)
    rchInput(0).SelBold = True
    rchInput(0).SelLength = 0

    rchInput(1).SelStart = 0
    rchInput(1).SelLength = Len(rchInput(1).Text)
    rchInput(1).SelItalic = True
    rchInput(1).SelLength = 0
End Sub

Private Sub rchInput_KeyPress(Index As Integer, KeyAscii As Integer)
Const CTRL_B = 2
Const CTRL_I = 9
Const CTRL_U = 21
Const CTRL_S = 19

    If KeyAscii = CTRL_B Then
        rchInput(Index).SelBold = Not rchInput(Index).SelBold
        KeyAscii = 0
    End If
    If KeyAscii = CTRL_I Then
        rchInput(Index).SelItalic = Not rchInput(Index).SelItalic
        KeyAscii = 0
    End If
    If KeyAscii = CTRL_U Then
        rchInput(Index).SelUnderline = Not rchInput(Index).SelUnderline
        KeyAscii = 0
    End If
    If KeyAscii = CTRL_S Then
        rchInput(Index).SelStrikeThru = Not rchInput(Index).SelStrikeThru
        KeyAscii = 0
    End If

    ' Recheck the check box values.
    If KeyAscii = 0 Then rchInput_SelChange Index
End Sub
' Set the proper check boxes for the selected text.
Private Sub rchInput_SelChange(Index As Integer)
    SetCheck chkBold(Index), rchInput(Index).SelBold
    SetCheck chkItalic(Index), rchInput(Index).SelItalic
    SetCheck chkUnderline(Index), rchInput(Index).SelUnderline
    SetCheck chkStrikethru(Index), rchInput(Index).SelStrikeThru
End Sub
