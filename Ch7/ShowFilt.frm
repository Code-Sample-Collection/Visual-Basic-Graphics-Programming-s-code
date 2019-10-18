VERSION 5.00
Begin VB.Form frmShowFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter Coefficients"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2730
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtFilter 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmShowFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' Prepare the form to display the kernel.
Public Sub PrepareForm(kernel() As Single)
Dim bound As Integer
Dim i As Integer
Dim j As Integer
Dim txt As String
Dim wid As Single
Dim hgt As Single

    bound = UBound(kernel, 1)

    ' Make a string holding the coefficients.
    For i = -bound To bound
        For j = -bound To bound
            txt = txt & " " & Format$(kernel(i, j), "0.000000") & " "
        Next j
        txt = txt & vbCrLf
    Next i
    txt = Left$(txt, Len(txt) - Len(vbCrLf))

    ' Size the TextBox.
    txtFilter.Text = txt
    wid = TextWidth(txt) + 120
    hgt = TextHeight(txt) + 120
    txtFilter.Move 120, 120, wid, hgt

    ' Size the form and position the Close button.
    If txtFilter.Width > cmdClose.Width Then
        Width = txtFilter.Width + 2 * 120 + Width - ScaleWidth
    Else
        Width = cmdClose.Width + 2 * 120 + Width - ScaleWidth
    End If
    Height = txtFilter.Height + cmdClose.Height + 3 * 120 + Height - ScaleHeight
    cmdClose.Move _
        (ScaleWidth - cmdClose.Width) / 2, _
        txtFilter.Height + 2 * 120
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
