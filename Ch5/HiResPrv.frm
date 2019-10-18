VERSION 5.00
Begin VB.Form frmHiResPrv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HiResPrv"
   ClientHeight    =   2250
   ClientLeft      =   510
   ClientTop       =   1290
   ClientWidth     =   2115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2250
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox HiddenPict 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   600
      ScaleHeight     =   1695
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PreviewPict 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuScale 
      Caption         =   "&Scale"
      Begin VB.Menu mnuSetScale 
         Caption         =   "&Large"
         Index           =   0
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSetScale 
         Caption         =   "&Normal"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSetScale 
         Caption         =   "&Small"
         Index           =   2
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmHiResPrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************
' Display a preview for this form.
' ************************************************
Public Sub ShowPreview(frm As Form)
    ' Make the PictureBoxes have Picture properties
    ' and give them useful color palettes.
    HiddenPict.Picture = HiddenPict.Image
    PreviewPict.Picture = PreviewPict.Image
    HiddenPict.Picture.hPal = frm.Picture.hPal
    PreviewPict.Picture.hPal = frm.Picture.hPal

    HiResPrint frm, HiddenPict, hires_ResizePrinter
    mnuSetScale_Click 1 ' Start at normal scale.
    Show vbModal
End Sub

' ************************************************
' Unload if the user presses escape.
' ************************************************
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then mnuFileClose_Click
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

' ************************************************
' Copy the hidden picture to PreviewPict at
' the appropriate scale.
' ************************************************
Private Sub mnuSetScale_Click(Index As Integer)
Dim i As Integer
Dim s As Single

    ' Check the selected menu item.
    For i = 0 To 2
        mnuSetScale(i).Checked = (i = Index)
    Next i

    ' Make PreviewPict the right size.
    Select Case Index
        Case 0  ' Large scale.
            s = 1.5
        Case 1  ' Normal scale.
            s = 1#
        Case 2  ' Small scale.
            s = 0.75
    End Select
    PreviewPict.Move 0, 0, _
        s * HiddenPict.Width, s * HiddenPict.Height
    Width = PreviewPict.Width + Width - ScaleWidth
    Height = PreviewPict.Height + Height - ScaleHeight

    ' Copy the image.
    HiddenPict.Picture = HiddenPict.Image
    PreviewPict.Picture = HiddenPict.Picture
    PreviewPict.PaintPicture HiddenPict.Image, _
        0, 0, _
        PreviewPict.ScaleWidth, _
        PreviewPict.ScaleHeight, _
        0, 0, _
        HiddenPict.ScaleWidth, _
        HiddenPict.ScaleHeight, vbSrcCopy
End Sub

Private Sub PreviewPict_Resize()
    Width = PreviewPict.Width + Width - ScaleWidth
    Height = PreviewPict.Height + Height - ScaleHeight
End Sub

