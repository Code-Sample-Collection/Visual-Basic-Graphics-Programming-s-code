VERSION 5.00
Begin VB.UserControl ScrolledWindow 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ControlContainer=   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   2310
   ToolboxBitmap   =   "Swin.ctx":0000
   Begin VB.HScrollBar hbar 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.VScrollBar vbar 
      Height          =   1335
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picOuter 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.PictureBox picInner 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   855
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "ScrolledWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

' Reposition picInner.
Private Sub hbar_Change()
    picInner.Left = -hbar.Value
End Sub

' Reposition picInner.
Private Sub hbar_Scroll()
    picInner.Left = -hbar.Value
End Sub

' Reparent the contained controls into picInner
' and see how much room they need.
Private Sub ReparentControls()
Dim ctl As Control
Dim xmax As Single
Dim ymax As Single
Dim need_wid As Single
Dim need_hgt As Single

    ' Do nothing if no controls have been loaded.
    If ContainedControls.Count = 0 Then Exit Sub

    For Each ctl In ContainedControls
        SetParent ctl.hWnd, picInner.hWnd

        xmax = ctl.Left + ctl.Width
        ymax = ctl.Top + ctl.Height
        If need_wid < xmax Then need_wid = xmax
        If need_hgt < ymax Then need_hgt = ymax
    Next ctl

    ' Make picInner big enough to hold the controls.
    picInner.Move 0, 0, need_wid, need_hgt

    ' Hide the borders on picInner and picOuter.
    picOuter.BorderStyle = vbBSNone
    picInner.BorderStyle = vbBSNone
End Sub

Private Sub UserControl_Resize()
    ' Hide the controls at design time.
    If Not Ambient.UserMode Then
        vbar.Visible = False
        hbar.Visible = False
        picInner.Visible = False
        Exit Sub
    End If

    ' Arrange the controls.
    ArrangeControls
End Sub
' Arrange the scroll bars.
Private Sub ArrangeControls()
Dim border_width As Single
Dim got_wid As Single
Dim got_hgt As Single
Dim need_wid As Single
Dim need_hgt As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    ' Reparent the controls.
    ReparentControls

    ' See how much room we have and need.
    border_width = picOuter.Width - picOuter.ScaleWidth
    got_wid = ScaleWidth - border_width
    got_hgt = ScaleHeight - border_width
    need_wid = picInner.Width
    need_hgt = picInner.Height

    ' See if we need the horizontal scroll bar.
    If need_wid > got_wid Then
        need_hbar = True
        got_hgt = got_hgt - hbar.Height
    End If

    ' See if we need the vertical scroll bar.
    If need_hgt > got_hgt Then
        need_vbar = True
        got_wid = got_wid - vbar.Width

        ' See if we now need the horizontal scroll bar.
        If (Not need_hbar) And need_wid > got_wid Then
            need_hbar = True
            got_hgt = got_hgt - hbar.Height
        End If
    End If

    ' Arrange the controls.
    picOuter.Move 0, 0, got_wid + border_width, got_hgt + border_width
    If need_hbar Then
        hbar.Move 0, got_hgt + border_width, got_wid + border_width
        hbar.Min = 0
        hbar.Max = picInner.ScaleWidth - got_wid
        hbar.SmallChange = got_wid / 5
        hbar.LargeChange = got_wid
        hbar.Visible = True
    Else
        hbar.Value = 0
        hbar.Visible = False
    End If
    If need_vbar Then
        vbar.Move got_wid + border_width, 0, vbar.Width, got_hgt + border_width
        vbar.Min = 0
        vbar.Max = picInner.ScaleHeight - got_hgt
        vbar.SmallChange = got_hgt / 5
        vbar.LargeChange = got_hgt
        vbar.Visible = True
    Else
        vbar.Value = 0
        vbar.Visible = False
    End If
End Sub

' Reposition picInner.
Private Sub vbar_Change()
    picInner.Top = -vbar.Value
End Sub

' Reposition picInner.
Private Sub vbar_Scroll()
    picInner.Top = -vbar.Value
End Sub
