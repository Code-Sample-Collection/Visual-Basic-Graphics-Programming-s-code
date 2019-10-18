VERSION 5.00
Begin VB.Form dlgScale 
   Caption         =   "Scale"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtScaleY 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "100.0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtScaleX 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "100.0"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Vertical"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Horizontal"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "dlgScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Canceled As Boolean
' Display the form. Return True if the user cancels.
Public Function ShowForm(ByRef x_scale As Single, ByRef y_scale As Single) As Boolean
    ' Assume we will cancel.
    Canceled = True

    ' Display the form.
    Show vbModal

    ShowForm = Canceled
    If Not Canceled Then
        On Error Resume Next
        x_scale = CSng(txtScaleX.Text) / 100#
        y_scale = CSng(txtScaleY.Text) / 100#
        On Error GoTo 0
    End If
End Function

Private Sub cmdCancel_Click()
    Hide
End Sub


Private Sub cmdOk_Click()
    Canceled = False
    Hide
End Sub
