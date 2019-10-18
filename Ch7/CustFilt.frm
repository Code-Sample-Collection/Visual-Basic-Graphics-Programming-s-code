VERSION 5.00
Begin VB.Form frmCustom 
   Caption         =   "Custom Filter"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBound 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtCoefficient 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "0.000000"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Bound"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean
Public CustomBound As Integer
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    Canceled = False
    Me.Hide
End Sub

Private Sub Form_Load()
    ' Assume we will cancel.
    Canceled = True
    txtBound.Text = "1"
End Sub

' Set a new filter size.
Private Sub txtBound_Change()
Dim new_bound As Integer
Dim i As Integer
Dim j As Integer
Dim idx As Integer
Dim X As Single
Dim Y As Single
Dim wid As Single

    ' Get the new bound.
    On Error Resume Next
    new_bound = CInt(txtBound.Text)
    If Err.Number > 0 Then Exit Sub
    If new_bound < 0 Then Exit Sub

    CustomBound = new_bound

    ' Position the controls.
    idx = 0
    Y = txtCoefficient(0).Top
    For i = -CustomBound To CustomBound
        X = txtCoefficient(0).Left
        For j = -CustomBound To CustomBound
            ' See if we need a new TextBox.
            If idx > txtCoefficient.UBound Then
                Load txtCoefficient(idx)
            End If

            ' Position the control.
            txtCoefficient(idx).Move X, Y
            txtCoefficient(idx).Visible = True
            X = X + txtCoefficient(0).Width + 120
            idx = idx + 1
        Next j
        Y = Y + txtCoefficient(0).Height + 120
    Next i

    ' Size the form and position the buttons.
    Height = txtCoefficient(idx - 1).Top + _
        txtCoefficient(idx - 1).Height + _
        cmdOk.Height + 2 * 120 + _
        Height - ScaleHeight
    wid = txtCoefficient(idx - 1).Left + _
        txtCoefficient(idx - 1).Width + 120 + _
        Width - ScaleWidth
    If wid < 2 * cmdOk.Width + 3 * 120 _
        Then wid = 2 * cmdOk.Width + 3 * 120
    Width = wid

    ' Position the buttons.
    cmdOk.Move ScaleWidth / 2 - cmdOk.Width - 60, _
        ScaleHeight - cmdOk.Height - 120
    cmdCancel.Move ScaleWidth / 2 + 60, cmdOk.Top

    ' Hide unneeded controls.
    For idx = idx To txtCoefficient.UBound
        txtCoefficient(idx).Visible = False
    Next idx
End Sub
