VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "ScrSaver Configuration"
   ClientHeight    =   1545
   ClientLeft      =   3555
   ClientTop       =   2850
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1545
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNumBalls 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of balls:"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub


' Save the new configuration values.
Private Sub cmdOk_Click()
    ' Get the new values.
    On Error Resume Next
    NumBalls = CInt(txtNumBalls)
    On Error GoTo 0

    ' Save the new values.
    SaveConfig

    ' Unload this form.
    Unload Me
End Sub

' Fill in current values.
Private Sub Form_Load()
    ' Load the current configuration information.
    LoadConfig
    
    txtNumBalls = Format$(NumBalls)
End Sub
