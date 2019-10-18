VERSION 5.00
Begin VB.Form frmBoxSize 
   Caption         =   "BoxSize"
   ClientHeight    =   3450
   ClientLeft      =   3210
   ClientTop       =   1425
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1000
   ScaleLeft       =   -10
   ScaleMode       =   0  'User
   ScaleTop        =   100
   ScaleWidth      =   10
   Begin VB.PictureBox RightPict 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox WrongPict 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "frmBoxSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Size the picture boxes and draw a diamond in each.
Private Sub Form_Load()
Dim wid As Single
Dim hgt As Single
Dim extra_wid As Single
Dim extra_hgt As Single

    ' Convert the desired width and height from
    ' twips into the form's custom coordinates.
    wid = Me.ScaleX(2500, 1, 0)
    hgt = Me.ScaleY(1500, 1, 0)
    
    WrongPict.Width = wid
    WrongPict.Height = hgt
    WrongPict.Line (0, 750)-(1250, 0)
    WrongPict.Line -(2500, 750)
    WrongPict.Line -(1250, 1500)
    WrongPict.Line -(0, 750)
    
    With RightPict
        extra_wid = .Width - Me.ScaleX( _
            .ScaleWidth, .ScaleMode, Me.ScaleMode)
        extra_hgt = .Height - Me.ScaleY( _
            .ScaleHeight, .ScaleMode, Me.ScaleMode)
        .Width = wid + extra_wid
        .Height = hgt + extra_hgt
    End With
    RightPict.Line (0, 750)-(1250, 0)
    RightPict.Line -(2500, 750)
    RightPict.Line -(1250, 1500)
    RightPict.Line -(0, 750)
End Sub

