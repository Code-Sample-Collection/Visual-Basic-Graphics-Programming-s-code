VERSION 5.00
Begin VB.Form frmIconAnim 
   Caption         =   "IconAnim"
   ClientHeight    =   1410
   ClientLeft      =   2925
   ClientTop       =   1875
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1410
   ScaleWidth      =   2175
   Begin VB.OptionButton optIcon 
      Caption         =   "Signal"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "Flame"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "Circle"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Timer tmrIcon 
      Interval        =   100
      Left            =   240
      Top             =   360
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   5
      Left            =   1680
      Picture         =   "IconAnim.frx":0000
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   6
      Left            =   2160
      Picture         =   "IconAnim.frx":030A
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   7
      Left            =   2640
      Picture         =   "IconAnim.frx":0614
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   4
      Left            =   1200
      Picture         =   "IconAnim.frx":091E
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   3
      Left            =   2640
      Picture         =   "IconAnim.frx":0C28
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   2
      Left            =   2160
      Picture         =   "IconAnim.frx":0F32
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   1
      Left            =   1680
      Picture         =   "IconAnim.frx":123C
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   0
      Left            =   1200
      Picture         =   "IconAnim.frx":1546
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSignal 
      Height          =   480
      Index           =   2
      Left            =   2160
      Picture         =   "IconAnim.frx":1850
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSignal 
      Height          =   480
      Index           =   1
      Left            =   1680
      Picture         =   "IconAnim.frx":1B5A
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSignal 
      Height          =   480
      Index           =   0
      Left            =   1200
      Picture         =   "IconAnim.frx":1E64
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRing 
      Height          =   480
      Index           =   7
      Left            =   2640
      Picture         =   "IconAnim.frx":216E
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRing 
      Height          =   480
      Index           =   6
      Left            =   2160
      Picture         =   "IconAnim.frx":2478
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRing 
      Height          =   480
      Index           =   5
      Left            =   1680
      Picture         =   "IconAnim.frx":2782
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRing 
      Height          =   480
      Index           =   4
      Left            =   1200
      Picture         =   "IconAnim.frx":2A8C
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRing 
      Height          =   480
      Index           =   3
      Left            =   2640
      Picture         =   "IconAnim.frx":2D96
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRing 
      Height          =   480
      Index           =   2
      Left            =   2160
      Picture         =   "IconAnim.frx":30A0
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRing 
      Height          =   480
      Index           =   1
      Left            =   1680
      Picture         =   "IconAnim.frx":33AA
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRing 
      Height          =   480
      Index           =   0
      Left            =   1200
      Picture         =   "IconAnim.frx":36B4
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmIconAnim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const icon_Ring = 0
Private Const icon_FLAME = 1
Private Const icon_SIGNAL = 2

Private IconType As Integer
Private IconNumber As Integer

Private Sub optIcon_Click(Index As Integer)
    IconType = Index
    IconNumber = 0
    Select Case IconType
        Case icon_Ring
            tmrIcon.Interval = 100
        Case icon_FLAME
            tmrIcon.Interval = 100
        Case icon_SIGNAL
            tmrIcon.Interval = 1000
    End Select
End Sub



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


' Display the next icon.
Private Sub tmrIcon_Timer()
    Select Case IconType
        Case icon_Ring
            Icon = imgRing(IconNumber).Picture
            IconNumber = (IconNumber + 1) Mod 8
        Case icon_FLAME
            Icon = imgFlame(Int(8 * Rnd)).Picture
        Case icon_SIGNAL
            Icon = imgSignal(IconNumber).Picture
            If IconNumber = 1 Then
                tmrIcon.Interval = 1000
            Else
                tmrIcon.Interval = 2000
            End If
            IconNumber = (IconNumber + 1) Mod 3
    End Select
End Sub


