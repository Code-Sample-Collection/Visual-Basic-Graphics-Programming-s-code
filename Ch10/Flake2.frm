VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFlake2 
   Caption         =   "Flake2"
   ClientHeight    =   4785
   ClientLeft      =   2280
   ClientTop       =   1185
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4785
   ScaleWidth      =   5820
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   480
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtDepth 
      Height          =   285
      Left            =   600
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "3"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   1080
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   3
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Depth"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open File..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmFlake2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159

' Coordinates of the points in the initiator.
Private NumInitiatorPoints As Integer
Private InitiatorX() As Single
Private InitiatorY() As Single

' Angles and distances for the generator.
Private NumGeneratorAngles As Integer
Private ScaleFactor As Single
Private GeneratorDTheta() As Single
' Draw the complete snowflake.
Private Sub DrawFlake(ByVal depth As Integer, ByVal length As Single)
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim dx As Single
Dim dy As Single
Dim theta As Single

    picCanvas.Cls

    ' Draw the snowflake.
    For i = 1 To NumInitiatorPoints
        x1 = InitiatorX(i - 1)
        y1 = InitiatorY(i - 1)
        x2 = InitiatorX(i)
        y2 = InitiatorY(i)
        dx = x2 - x1
        dy = y2 - y1
        theta = ATan2(dy, dx)
        DrawFlakeEdge depth, x1, y1, _
            theta, length
    Next i
End Sub

' Recursively draw a snowflake edge starting at
' (x1, y1) in direction theta and distance dist.
' Leave the coordinates of the endpoint in
' (x1, y1).
Private Sub DrawFlakeEdge(ByVal depth As Integer, ByRef x1 As Single, ByRef y1 As Single, ByVal theta As Single, ByVal dist As Single)
Dim status As Integer
Dim i As Integer
Dim x2 As Single
Dim y2 As Single

    If depth <= 0 Then
        x2 = x1 + dist * Cos(theta)
        y2 = y1 + dist * Sin(theta)
        picCanvas.Line (x1, y1)-(x2, y2)
        x1 = x2
        y1 = y2
        Exit Sub
    End If

    ' Recursively draw the edge.
    dist = dist * ScaleFactor
    For i = 1 To NumGeneratorAngles
        theta = theta + GeneratorDTheta(i)
        DrawFlakeEdge depth - 1, x1, y1, theta, dist
    Next i
End Sub
Private Sub CmdGo_Click()
Dim depth As Integer
Dim dx As Single
Dim dy As Single
Dim length As Single

    picCanvas.Cls
    MousePointer = vbHourglass
    DoEvents

    ' Get the parameters.
    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    depth = CInt(txtDepth.Text)

    ' Find the distance between initiator points.
    dx = InitiatorX(2) - InitiatorX(1)
    dy = InitiatorY(2) - InitiatorY(1)
    length = Sqr(dx * dx + dy * dy)

    ' Draw the snowflake.
    DrawFlake depth, length

    MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    dlgFile.Filter = "Snowflake Files (*.sno)|*.sno"
    dlgFile.InitDir = App.Path
End Sub

Private Sub Form_Resize()
Dim wid As Single

    ' Make the picCanvas as big as possible.
    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub

' Load a snowflake definition file with format:
'
'   # Initiator points.
'   (x1, y1)
'   (x2, y2)
'       :
'   scalefactor
'   # Generator angles.
'   theta1
'   theta2
'       :
Private Sub mnuFileOpen_Click()
Dim file_name As String
Dim fnum As Integer
Dim theta As Single
Dim i As Integer

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    file_name = Trim$(dlgFile.FileName)
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    ' Open the file.
    fnum = FreeFile
    Open file_name For Input Access Read As #fnum

    ' Read the initiator.
    Input #fnum, NumInitiatorPoints
    ReDim InitiatorX(0 To NumInitiatorPoints)
    ReDim InitiatorY(0 To NumInitiatorPoints)
    For i = 1 To NumInitiatorPoints
        Input #fnum, InitiatorX(i), InitiatorY(i)
    Next i
    InitiatorX(0) = InitiatorX(NumInitiatorPoints)
    InitiatorY(0) = InitiatorY(NumInitiatorPoints)

    ' Read the generator information.
    Input #fnum, ScaleFactor, NumGeneratorAngles
    ReDim GeneratorDTheta(1 To NumGeneratorAngles)
    For i = 1 To NumGeneratorAngles
        Input #fnum, theta
        GeneratorDTheta(i) = theta * PI / 180
    Next i

    Close #fnum

    Caption = "Flake2 [" & dlgFile.FileTitle & "]"
    cmdGo.Enabled = True
End Sub


