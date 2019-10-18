VERSION 5.00
Begin VB.UserControl ScrolledWindow 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ScrWin.ctx":0000
   Begin VB.PictureBox Plug 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   2880
      ScaleHeight     =   1335
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.VScrollBar VBar 
      Height          =   3375
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HBar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   4575
   End
End
Attribute VB_Name = "ScrolledWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CCC Scrolled window container control"
Option Explicit

' Current contained control offsets.
Private xoff As Single
Private yoff As Single

' Border styles.
Public Enum swBorderStyle
    None_swBorderStyle = 0
    Fixed_Single_swBorderStyle = 1
End Enum

Private Const dflt_BorderStyle = Fixed_Single_swBorderStyle
' Move the contained controls appropriately.
Private Sub HBar_Change()
    ArrangeControls
End Sub
' Move the contained controls appropriately.
Private Sub VBar_Change()
    ArrangeControls
End Sub
' Move the contained controls appropriately.
Private Sub HBar_Scroll()
    ArrangeControls
End Sub
' Move the contained controls appropriately.
Private Sub VBar_Scroll()
    ArrangeControls
End Sub
' Arrange the scroll bars.
Private Sub UserControl_Resize()
Dim ctl As Object
Dim is_visible As Boolean
Dim xmax As Single
Dim ymax As Single
Dim x1 As Single
Dim y1 As Single
Dim got_wid As Single
Dim got_hgt As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    ' Do nothing at design time.
    If Not Ambient.UserMode Then
        VBar.Visible = False
        HBar.Visible = False
        Plug.Visible = False
        Exit Sub
    End If

    ' Start with really wild values.
    xmax = 0
    ymax = 0
    
    ' In the following, if a control does not
    ' have a given property, the value remains
    ' unchanged. For example, in x1 = ctl.Left
    ' if ctl has no Left property, the value of
    ' x1 is whatever it was before the
    ' operation.
    '
    ' The following are safe values for now.
    ' Once a real value is set, the real value
    ' will be safe for subsequent controls.
    x1 = 0
    y1 = 0

    ' Guard against controls with no Visible
    ' Top, Left, Width, and Height properties.
    On Error Resume Next

    ' Find bounds for the visible controls
    ' contained within.
    For Each ctl In ContainedControls
        is_visible = False
        is_visible = ctl.Visible
        If is_visible Then
            x1 = ctl.Left + ctl.Width
            y1 = ctl.Top + ctl.Height
            If xmax < x1 Then xmax = x1
            If ymax < y1 Then ymax = y1
        End If
    Next ctl
    
    ' See which scroll bars we need.
    got_wid = ScaleWidth
    got_hgt = ScaleHeight
    
    ' See if we need the horizontal scroll bar.
    If xmax > got_wid Then
        ' We do. Leave room for it.
        need_hbar = True
        got_hgt = got_hgt - HBar.Height
    Else
        need_hbar = False
    End If
    ' See if we need the vertical scroll bar.
    If ymax > got_hgt Then
        ' We do. Leave room for it.
        need_vbar = True
        got_wid = got_wid - VBar.Width
    
        ' See if we now need the horizontal
        ' scroll bar.
        If (Not need_hbar) And (xmax > got_wid) Then
            ' We do. Leave room for it.
            need_hbar = True
            got_hgt = got_hgt - HBar.Height
        End If
    Else
        need_vbar = False
    End If

    ' Arrange the controls.
    If need_hbar Then
        HBar.Move 0, got_hgt, got_wid
        HBar.Max = xmax - got_wid
        HBar.SmallChange = (xmax - got_wid) / 5
        HBar.LargeChange = (HBar.Max - HBar.Min) * _
            got_wid / (xmax - got_wid)
        HBar.Visible = True
    Else
        HBar.Value = 0
        HBar.Visible = False
    End If
    If need_vbar Then
        VBar.Move got_wid, 0, VBar.Width, got_hgt
        VBar.Max = ymax - got_hgt
        VBar.SmallChange = (ymax - got_hgt) / 5
        VBar.LargeChange = (VBar.Max - VBar.Min) * _
            got_hgt / (ymax - got_hgt)
        VBar.Visible = True
    Else
        VBar.Value = 0
        VBar.Visible = False
    End If
    If need_hbar And need_vbar Then
        Plug.Move got_wid, got_hgt, VBar.Width, HBar.Height
        Plug.Visible = True
    Else
        Plug.Visible = False
    End If

    ' Make sure these are on top.
    HBar.ZOrder
    VBar.ZOrder
    Plug.ZOrder
    
    ' Place the contained controls.
    ArrangeControls
End Sub
' Position the contained controls.
Private Sub ArrangeControls()
Attribute ArrangeControls.VB_Description = "Arranges the controls for the current scroll bar values."
Dim dx As Single
Dim dy As Single
Dim ctl As Object
Dim is_visible As Boolean

    ' See where the controls should be.
    dx = -HBar.Value - xoff
    dy = -VBar.Value - yoff
    xoff = dx + xoff
    yoff = dy + yoff
    
    ' Guard against controls with no Visible,
    ' Left, and Top properties.
    On Error Resume Next

    ' Position the controls.
    For Each ctl In ContainedControls
        is_visible = False
        is_visible = ctl.Visible
        If is_visible Then
            ctl.Move ctl.Left + dx, ctl.Top + dy
        End If
    Next ctl
End Sub
' Set default property values.
Private Sub UserControl_InitProperties()
    BorderStyle = dflt_BorderStyle
End Sub
' Read the property values.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BorderStyle = PropBag.ReadProperty("BorderStyle", dflt_BorderStyle)
End Sub

' Save the property values.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BorderStyle", BorderStyle, dflt_BorderStyle
End Sub

' Delegate BorderStyle to UserControl.
Property Let BorderStyle(Style As swBorderStyle)
Attribute BorderStyle.VB_Description = "Sets the control's border style."
    UserControl.BorderStyle = Style
End Property
' Delegate BorderStyle to UserControl.
Property Get BorderStyle() As swBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property
