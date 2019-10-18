VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Preview"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hbarMarble 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   2295
   End
   Begin VB.VScrollBar vbarMarble 
      Height          =   2295
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picOuter 
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.PictureBox picInner 
         AutoSize        =   -1  'True
         Height          =   6045
         Left            =   360
         ScaleHeight     =   5985
         ScaleWidth      =   6165
         TabIndex        =   1
         Top             =   240
         Width           =   6225
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Arrange the scroll bars.
Private Sub ArrangeControls()
Dim border_width As Single
Dim got_wid As Single
Dim got_hgt As Single
Dim need_wid As Single
Dim need_hgt As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    ' See how much room we have and need.
    border_width = picOuter.Width - picOuter.ScaleWidth
    got_wid = ScaleWidth - border_width
    got_hgt = ScaleHeight - border_width
    need_wid = picInner.Width
    need_hgt = picInner.Height

    ' See if we need the horizontal scroll bar.
    If need_wid > got_wid Then
        need_hbar = True
        got_hgt = got_hgt - hbarMarble.Height
    End If

    ' See if we need the vertical scroll bar.
    If need_hgt > got_hgt Then
        need_vbar = True
        got_wid = got_wid - vbarMarble.Width

        ' See if we now need the horizontal scroll bar.
        If (Not need_hbar) And need_wid > got_wid Then
            need_hbar = True
            got_hgt = got_hgt - hbarMarble.Height
        End If
    End If

    ' Arrange the controls.
    picOuter.Move 0, 0, got_wid + border_width, got_hgt + border_width
    If need_hbar Then
        hbarMarble.Move 0, got_hgt + border_width, got_wid + border_width
        hbarMarble.Min = 0
        hbarMarble.Max = picInner.ScaleWidth - got_wid
        hbarMarble.SmallChange = got_wid / 5
        hbarMarble.LargeChange = got_wid
        hbarMarble.Visible = True
    Else
        hbarMarble.Value = 0
        hbarMarble.Visible = False
    End If
    If need_vbar Then
        vbarMarble.Move got_wid + border_width, 0, vbarMarble.Width, got_hgt + border_width
        vbarMarble.Min = 0
        vbarMarble.Max = picInner.ScaleHeight - got_hgt
        vbarMarble.SmallChange = got_hgt / 5
        vbarMarble.LargeChange = got_hgt
        vbarMarble.Visible = True
    Else
        vbarMarble.Value = 0
        vbarMarble.Visible = False
    End If
End Sub
Private Sub Form_Load()
    picInner.AutoSize = True
    picInner.Move 0, 0
    picInner.BorderStyle = vbBSNone
    picInner.BackColor = vbWhite
    picInner.AutoRedraw = True
End Sub

Private Sub Form_Resize()
    ArrangeControls
End Sub


' Reposition picInner.
Private Sub hbarMarble_Change()
    picInner.Left = -hbarMarble.Value
End Sub


' Reposition picInner.
Private Sub hbarMarble_Scroll()
    picInner.Left = -hbarMarble.Value
End Sub


' Reposition picInner.
Private Sub vbarMarble_Change()
    picInner.Top = -vbarMarble.Value
End Sub


' Reposition picInner.
Private Sub vbarMarble_Scroll()
    picInner.Top = -vbarMarble.Value
End Sub


