VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBspline2 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Bspline2 []"
   ClientHeight    =   6750
   ClientLeft      =   300
   ClientTop       =   840
   ClientWidth     =   9645
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
   ScaleHeight     =   6750
   ScaleWidth      =   9645
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3600
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picZY 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   6480
      ScaleHeight     =   3075
      ScaleWidth      =   3075
      TabIndex        =   15
      Top             =   3600
      Width           =   3135
   End
   Begin VB.PictureBox picXY 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   3240
      ScaleHeight     =   3075
      ScaleWidth      =   3075
      TabIndex        =   12
      Top             =   3600
      Width           =   3135
   End
   Begin VB.PictureBox picXZ 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   3075
      TabIndex        =   11
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   3855
      Begin VB.CheckBox chkShowControlGrid 
         Caption         =   "Show Control Grid"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Draw"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkShowControlPoints 
         Caption         =   "Show Control Points"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Points"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton cmdInitialize 
         Caption         =   "Initialize"
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtNumZ 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "4"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtNumX 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "4"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "NumZ"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "NumX"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   3960
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Z-Y (side view)"
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   16
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "X-Y (side view)"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   14
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "X-Z (top view)"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileExitSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmBspline2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GAP = 0.2

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

Private TheSurface As Bspline3d

Private NumX As Integer
Private NumZ As Integer
Private PtX() As Single
Private PtY() As Single
Private PtZ() As Single

Private DragPicture As Integer
Private DragI As Integer
Private DragJ As Integer
Private DragX As Single
Private DragY As Single

' Load data from this file.
Private Sub LoadBsplineData(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim i As Integer
Dim j As Integer

    On Error GoTo LoadError
    fnum = FreeFile
    Open file_name For Input As fnum

    ' Get the number of points.
    Input #fnum, NumX, NumZ

    ' Initialize the data.
    txtNumX.Text = Format$(NumX)
    txtNumZ.Text = Format$(NumZ)

    ' Prepare to save the data.
    cmdInitialize_Click

    ' Read the control point locations.
    For i = 1 To NumX
        For j = 1 To NumZ
            Input #fnum, PtX(i, j), PtY(i, j), PtZ(i, j)
        Next j
    Next i

    Close fnum

    ' Draw the control points.
    DrawSideViews
    picCanvas.Cls

    Caption = "Bspline2 [" & file_title & "]"
    Exit Sub

LoadError:
    MsgBox "Error " & Format$(Err.Number) & _
        " loading data." & vbCrLf & _
        Err.Description
    Exit Sub
End Sub
' Save data into this file.
Private Sub SaveBsplineData(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim i As Integer
Dim j As Integer

    On Error GoTo SaveError
    fnum = FreeFile
    Open file_name For Output As fnum

    ' Save the number of points.
    Write #fnum, NumX, NumZ

    ' Save the control point locations.
    For i = 1 To NumX
        For j = 1 To NumZ
            Write #fnum, PtX(i, j), PtY(i, j), PtZ(i, j)
        Next j
    Next i

    Close fnum
    Caption = "Bspline2 [" & file_title & "]"

    Exit Sub

SaveError:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving data file." & vbCrLf & _
        Err.Description
    Exit Sub
End Sub

Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Width = wid
End Sub

' Load a Bspline surface data file.
Private Sub mnuFileOpen_Click()
Dim file_name As String

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Flags = _
        cdlOFNExplorer Or _
        cdlOFNLongNames Or _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    file_name = Trim$(dlgFile.FileName)
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    ' Load the information.
    LoadBsplineData file_name, dlgFile.FileTitle
End Sub
' Save a Bspline surface data file.
Private Sub mnuFileSave_Click()
Dim file_name As String

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Flags = _
        cdlOFNExplorer Or _
        cdlOFNLongNames Or _
        cdlOFNOverwritePrompt Or _
        cdlOFNHideReadOnly
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    file_name = Trim$(dlgFile.FileName)
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    ' Save the information.
    SaveBsplineData file_name, dlgFile.FileTitle
End Sub

' Display the surface.
Private Sub DrawData(pic As Object)
Dim S(1 To 4, 1 To 4) As Single
Dim T(1 To 4, 1 To 4) As Single
Dim ST(1 To 4, 1 To 4) As Single
Dim PST(1 To 4, 1 To 4) As Single

    If TheSurface Is Nothing Then Exit Sub

    MousePointer = vbHourglass
    Refresh

    TheSurface.DrawControls = (chkShowControlPoints.value = vbChecked)
    TheSurface.DrawGrid = (chkShowControlGrid.value = vbChecked)

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 30, -30, 1
    m3Translate T, 200, 100, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the points.
    TheSurface.ApplyFull PST

    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next

    ' Display the data.
    pic.Cls
    TheSurface.Draw pic, EyeR

    picCanvas.SetFocus
    MousePointer = vbDefault
End Sub
' Draw the X-Z plane projection of the
' control points.
Private Sub DrawXZ()
Dim i As Integer
Dim j As Integer

    For i = 1 To NumX
        For j = 1 To NumZ
            picXZ.Line (PtX(i, j) - GAP / 2, PtZ(i, j) - GAP / 2)-Step(GAP, GAP), vbBlack, BF
        Next j
    Next i
End Sub
' Draw the X-Y plane projection of the
' control points.
Private Sub DrawSideViews()
Dim i As Integer
Dim j As Integer

    picXZ.Cls
    picXY.Cls
    picZY.Cls

    ' Draw the points.
    For i = 1 To NumX
        For j = 1 To NumZ
            picXZ.Line (PtX(i, j) - GAP / 2, PtZ(i, j) - GAP / 2)-Step(GAP, GAP), vbBlack, BF
            picXY.Line (PtX(i, j) - GAP / 2, PtY(i, j) - GAP / 2)-Step(GAP, GAP), vbBlack, BF
            picZY.Line (PtZ(i, j) - GAP / 2, PtY(i, j) - GAP / 2)-Step(GAP, GAP), vbBlack, BF
        Next j
    Next i

    ' Draw the lines.
    For i = 1 To NumX
        picXY.CurrentX = PtX(i, 1)
        picXY.CurrentY = PtY(i, 1)
        picXZ.CurrentX = PtX(i, 1)
        picXZ.CurrentY = PtZ(i, 1)
        picZY.CurrentX = PtZ(i, 1)
        picZY.CurrentY = PtY(i, 1)
        For j = 2 To NumZ
            picXY.Line -(PtX(i, j), PtY(i, j))
            picXZ.Line -(PtX(i, j), PtZ(i, j))
            picZY.Line -(PtZ(i, j), PtY(i, j))
        Next j
    Next i

    For j = 1 To NumZ
        picXY.CurrentX = PtX(1, j)
        picXY.CurrentY = PtY(1, j)
        picXZ.CurrentX = PtX(1, j)
        picXZ.CurrentY = PtZ(1, j)
        picZY.CurrentX = PtZ(1, j)
        picZY.CurrentY = PtY(1, j)
        For i = 2 To NumX
            picXY.Line -(PtX(i, j), PtY(i, j))
            picXZ.Line -(PtX(i, j), PtZ(i, j))
            picZY.Line -(PtZ(i, j), PtY(i, j))
        Next i
    Next j
End Sub
' Draw the spline.
Private Sub cmdDraw_Click()
    CreateData
    DrawData picCanvas
End Sub

' Initialize the points array.
Private Sub cmdInitialize_Click()
Dim i As Integer
Dim j As Integer
Dim xmin As Single
Dim ymin As Single
Dim zmin As Single
Dim dx As Single
Dim dz As Single
Dim dy As Single

    NumX = CInt(txtNumX.Text)
    NumZ = CInt(txtNumZ.Text)
    ReDim PtX(1 To NumX, 1 To NumZ)
    ReDim PtY(1 To NumX, 1 To NumZ)
    ReDim PtZ(1 To NumX, 1 To NumZ)

    ' Spread the points around slightly.
    xmin = picXZ.ScaleLeft
    dx = picXZ.ScaleWidth / (NumX + 2)
    zmin = picXZ.ScaleTop
    dz = picXZ.ScaleHeight / (NumZ + 2)
    ymin = 0
    dy = GAP * 1.5
    For i = 1 To NumX
        For j = 1 To NumZ
            PtX(i, j) = xmin + i * dx
            PtY(i, j) = ymin + (i + j) * dy
            PtZ(i, j) = zmin + j * dz
        Next j
    Next i

    DrawSideViews
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

    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
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
    ' Initialize the file dialog.
    dlgFile.InitDir = App.Path
    dlgFile.CancelError = True
    dlgFile.Filter = _
        "Bspline Files (*.bsp)|*.bsp|" & _
        "All Files (*.*)|*.*"

    ' Initialize the eye position.
    EyeR = 10
    EyeTheta = PI * 1.2
    EyePhi = PI * 0.1

    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Set some useful scales.
    picXZ.ScaleLeft = -5
    picXZ.ScaleWidth = 10
    picXZ.ScaleTop = 5
    picXZ.ScaleHeight = -10
    picXY.ScaleLeft = -5
    picXY.ScaleWidth = 10
    picXY.ScaleTop = 5
    picXY.ScaleHeight = -10
    picZY.ScaleLeft = -5
    picZY.ScaleWidth = 10
    picZY.ScaleTop = 5
    picZY.ScaleHeight = -10

    ' Start with some uninitialized data.
    cmdInitialize_Click
    cmdDraw_Click
End Sub
' Create the surface.
Private Sub CreateData()
Const GapU = 0.25
Const GapV = 0.25
Const Du = GapU / 1
Const Dv = GapV / 1

Dim i As Integer
Dim j As Integer

    MousePointer = vbHourglass
    Refresh

    Set TheSurface = New Bspline3d

    TheSurface.DrawControls = (chkShowControlPoints.value = vbChecked)
    TheSurface.DrawGrid = (chkShowControlGrid.value = vbChecked)

    ' Initialize the control points.
    TheSurface.SetBounds NumX, NumZ
    For i = 1 To NumX
        For j = 1 To NumZ
            TheSurface.SetControlPoint i, j, _
                PtX(i, j), PtY(i, j), PtZ(i, j)
        Next j
    Next i

    ' Initialize the B-spline.
    TheSurface.InitializeGrid 3, 3, _
        GapU, GapV, Du, Dv
End Sub
Private Sub chkShowControlPoints_Click()
    DrawData picCanvas
End Sub
Private Sub chkshowcontrolgrid_Click()
    DrawData picCanvas
End Sub
' Try to drag a node.
Private Sub picXZ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim j As Integer

    ' Find the node.
    DragPicture = 0
    For i = 1 To NumX
        For j = 1 To NumZ
            If (Abs(PtX(i, j) - X) < GAP / 2) And _
               (Abs(PtZ(i, j) - Y) < GAP / 2) _
            Then
                ' This is the node.
                DragI = i
                DragJ = j
                DragPicture = 1
                Exit For
            End If
        Next j
    Next i

    ' See if we found a node.
    If DragPicture < 1 Then Exit Sub

    ' Start the drag.
    picXZ.DrawMode = vbInvert
    DragX = X
    DragY = Y
    picXZ.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
End Sub
' Try to drag a node.
Private Sub picXY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim j As Integer

    ' Find the node.
    DragPicture = 0
    For i = 1 To NumX
        For j = 1 To NumZ
            If (Abs(PtX(i, j) - X) < GAP / 2) And _
               (Abs(PtY(i, j) - Y) < GAP / 2) _
            Then
                ' This is the node.
                DragI = i
                DragJ = j
                DragPicture = 2
                Exit For
            End If
        Next j
    Next i

    ' See if we found a node.
    If DragPicture < 1 Then Exit Sub

    ' Start the drag.
    picXY.DrawMode = vbInvert
    DragX = X
    DragY = Y
    picXY.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
End Sub
' Try to drag a node.
Private Sub picZY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim j As Integer

    ' Find the node.
    DragPicture = 0
    For i = 1 To NumX
        For j = 1 To NumZ
            If (Abs(PtZ(i, j) - X) < GAP / 2) And _
               (Abs(PtY(i, j) - Y) < GAP / 2) _
            Then
                ' This is the node.
                DragI = i
                DragJ = j
                DragPicture = 3
                Exit For
            End If
        Next j
    Next i

    ' See if we found a node.
    If DragPicture < 1 Then Exit Sub

    ' Start the drag.
    picZY.DrawMode = vbInvert
    DragX = X
    DragY = Y
    picZY.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
End Sub

' Continue dragging a node.
Private Sub picXZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragPicture <> 1 Then Exit Sub

    picXZ.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
    DragX = X
    DragY = Y
    picXZ.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
End Sub
' Continue dragging a node.
Private Sub picXY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragPicture <> 2 Then Exit Sub

    picXY.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
    DragX = X
    DragY = Y
    picXY.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
End Sub

' Finish dragging a node.
Private Sub picXZ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragPicture <> 1 Then Exit Sub
    DragPicture = 0

    ' Update the node's position and redraw.
    PtX(DragI, DragJ) = X
    PtZ(DragI, DragJ) = Y
    picXZ.DrawMode = vbCopyPen

    DrawSideViews
End Sub
' Finish dragging a node.
Private Sub picXY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragPicture <> 2 Then Exit Sub
    DragPicture = 0

    ' Update the node's position and redraw.
    PtX(DragI, DragJ) = X
    PtY(DragI, DragJ) = Y
    picXY.DrawMode = vbCopyPen

    DrawSideViews
End Sub

' Finish dragging a node.
Private Sub picZY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragPicture <> 3 Then Exit Sub
    DragPicture = 0

    ' Update the node's position and redraw.
    PtZ(DragI, DragJ) = X
    PtY(DragI, DragJ) = Y
    picZY.DrawMode = vbCopyPen

    DrawSideViews
End Sub


' Continue dragging a node.
Private Sub picZY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragPicture <> 3 Then Exit Sub

    picZY.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
    DragX = X
    DragY = Y
    picZY.Line (DragX - GAP / 2, DragY - GAP / 2)-Step(GAP, GAP), , BF
End Sub


