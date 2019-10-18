VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Print"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Const X = 1440
Const Y = 720

    AutoRedraw = True
    Line (X - 100, Y - 100)-Step(100, 100)
    Line (X + 100, Y - 100)-Step(-100, 100)

    CurrentX = X
    CurrentY = Y

    Print "Line 1"
    Print "Line 2", "Zone 2", "Zone 3";
    Print "Line 3"
End Sub


