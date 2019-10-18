VERSION 5.00
Begin VB.Form frmValley 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Valley"
   ClientHeight    =   5295
   ClientLeft      =   300
   ClientTop       =   570
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5295
   ScaleWidth      =   9135
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtDy 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "0.25"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "3"
      Top             =   360
      Width           =   495
   End
   Begin VB.CheckBox chkRemoveHidden 
      Caption         =   "Remove Hidden"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   2160
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "Dy"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Level"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmValley"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

Private Const Dtheta = PI / 20
Private Const Dphi = PI / 20
Private Const Dr = 1

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

Private TheGrid As ValleyGrid3d

Private Const Xmin = -5
Private Const Zmin = -5
' Return the Y coordinate for these X and
' Z coordinates.
Private Function YValue(ByVal X As Single, ByVal Z As Single)
Dim Y As Single

    Y = -2 * Cos(2 * PI / 10 * Z) * (5 - Abs(Z)) / 5 + 0.25 * Sin(2 * X) + 0.25 * Sin(1# * X) + 0.5 * Rnd
    If Y < -1 Then Y = -1

    YValue = Y
End Function

' Project and display the data.
Private Sub DrawData(pic As Object)
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim S(1 To 4, 1 To 4) As Single
Dim T(1 To 4, 1 To 4) As Single
Dim ST(1 To 4, 1 To 4) As Single
Dim PST(1 To 4, 1 To 4) As Single

    MousePointer = vbHourglass
    DoEvents

    ' Make the data.
    CreateData

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 35, -35, 1
    m3Translate T, 230, 175, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the points.
    TheGrid.ApplyFull PST

    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next

    ' Display the data.
    pic.Cls
    TheGrid.RemoveHidden = (chkRemoveHidden.value = vbChecked)
    TheGrid.Draw pic
    pic.Refresh

    MousePointer = vbDefault
End Sub

Private Sub cmdDraw_Click()
    DrawData picCanvas
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            EyeTheta = EyeTheta - Dtheta
        
        Case vbKeyRight
            EyeTheta = EyeTheta + Dtheta
        
        Case vbKeyUp
            EyePhi = EyePhi - Dphi
        
        Case vbKeyDown
            EyePhi = EyePhi + Dphi
                
        Case Else
            Exit Sub
    End Select

    m3PProject Projector, m3Parallel, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    DrawData picCanvas
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("+")
            EyeR = EyeR + Dr
        
        Case Asc("-")
            EyeR = EyeR - Dr
        
        Case Else
            Exit Sub
    End Select

    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    DrawData picCanvas
End Sub

Private Sub Form_Load()
    Randomize

    ' Initialize the eye position.
    EyeR = 10
    EyeTheta = PI * 0.2
    EyePhi = PI * 0.1

    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Project and draw the data.
    Me.Show
    DrawData picCanvas
End Sub

' Create the surface.
Private Sub CreateData()
Const Dx = 1
Const Dz = 1
Const NumX = -2 * Xmin / Dx
Const NumZ = -2 * Zmin / Dz

Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim level As Integer
Dim Dy As Single
Dim small_dx As Single
Dim small_dz As Single
Dim min_z As Single
Dim max_z As Single
Dim river_width As Single
Dim period1 As Single
Dim period2 As Single
Dim period3 As Single

    Set TheGrid = New ValleyGrid3d
    TheGrid.SetBounds Xmin, Dx, NumX, Zmin, Dz, NumZ

    X = Xmin
    For i = 1 To NumX
        Z = Zmin
        For j = 1 To NumZ
            Y = YValue(X, Z)
            TheGrid.SetValue X, Y, Z
            Z = Z + Dz
        Next j
        X = X + Dx
    Next i

    On Error Resume Next
    level = CInt(txtLevel.Text)
    If Err.Number <> 0 Then
        txtLevel.Text = "3"
        level = 3
    End If

    Dy = CSng(txtDy.Text)
    If Err.Number <> 0 Then
        txtDy.Text = "0.25"
        Dy = 0.25
    End If

    TheGrid.GenerateSurface level, Dy

    ' Flatten the bottom.
    TheGrid.Flatten -1, 0.25, 0.25

    ' Make a river bed in the bottom.
    period1 = 0.5 + Rnd * 1
    period2 = 0.5 + Rnd * 1
    period3 = 0.5 + Rnd * 1
    small_dx = Dx / (2 ^ level)
    small_dz = Dz / (2 ^ level)
    X = Xmin
    For i = 1 To NumX * (2 ^ level) - (2 ^ level - 1)
        river_width = Abs(Sin(X / 3) * Rnd / 2 + Sin(X / 2.5) * Rnd / 3 - Sin(X) * Rnd / 4) + 1 / 8
        If river_width < 2 * small_dx Then river_width = 2 * small_dx
        min_z = Sin(X / period1) / 2 + Sin(X / period2) / 4 - Sin(X / period3) / 8
        max_z = min_z + river_width
        For Z = min_z To max_z Step small_dz
            TheGrid.SetValue X, -1.1, Z
        Next Z
        X = X + small_dx
    Next i
End Sub
