VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChaos 
   Caption         =   "Chaos"
   ClientHeight    =   4335
   ClientLeft      =   2280
   ClientTop       =   1185
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   5310
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   960
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmChaos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Running As Boolean
Private NumAnchors As Integer
Private AnchorX() As Single
Private AnchorY() As Single

' Draw the anchor points.
Private Sub DrawAnchors()
Const GAP = 2

Dim i As Integer
Dim wid As Single
Dim hgt As Single

    wid = picCanvas.ScaleWidth
    hgt = picCanvas.ScaleHeight

    picCanvas.Cls
    For i = 1 To NumAnchors
        picCanvas.Line _
            (wid * AnchorX(i) - GAP, hgt * AnchorY(i) - GAP)- _
            Step(2 * GAP, 2 * GAP), , BF
    Next i
End Sub
' Load anchor point data.
Private Sub LoadChaosData(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim i As Integer

    fnum = FreeFile
    Open file_name For Input Access Read As #fnum

    Input #fnum, NumAnchors
    ReDim AnchorX(1 To NumAnchors)
    ReDim AnchorY(1 To NumAnchors)
    For i = 1 To NumAnchors
        Input #fnum, AnchorX(i), AnchorY(i)
    Next i

    Close #fnum

    DrawAnchors
    Caption = "Chaos [" & file_title & "]"
    cmdGo.Enabled = True
End Sub
' This routine prints chaos game coordinates for
' a regular polygon. It is not used in the program,
' but I am leaving it here because you may find
' it useful.
Private Sub PrintPolygonPoints(ByVal num_sides As Integer)
Const PI = 3.14159265

Dim theta As Single
Dim dtheta As Single
Dim i As Integer
Dim X As Single
Dim Y As Single

    theta = -PI / 2
    dtheta = 2 * PI / num_sides
    For i = 1 To num_sides
        X = 0.5 + 0.45 * Cos(theta)
        Y = 0.5 + 0.45 * Sin(theta)
        Debug.Print Format$(X, "0.00") & ", " & _
            Format$(Y, "0.00")
        theta = theta + dtheta
    Next i
End Sub

Private Sub CmdGo_Click()
    If Running Then
        Running = False
        cmdGo.Enabled = False
        cmdGo.Caption = "Stopped"
    Else
        Running = True
        cmdGo.Caption = "Stop"
        DrawAnchors
        PlayGame
        cmdGo.Enabled = True
        cmdGo.Caption = "Go"
    End If
End Sub

' Play the chaos game.
Private Sub PlayGame()
Dim wid As Single
Dim hgt As Single
Dim X As Single
Dim Y As Single
Dim anchor As Integer
Dim i As Integer

    ' See how much room we have.
    wid = picCanvas.ScaleWidth
    hgt = picCanvas.ScaleHeight

    ' Pick a random starting point.
    X = wid * Rnd
    Y = hgt * Rnd

    ' Start the game.
    i = 0
    Do While Running
        ' Pick a random anchor point.
        anchor = Int(NumAnchors * Rnd + 1)

        ' Move halfway there.
        X = (X + wid * AnchorX(anchor)) / 2
        Y = (Y + hgt * AnchorY(anchor)) / 2
        picCanvas.PSet (X, Y)

        ' To make things faster, only DoEvents
        ' every 100 times.
        i = i + 1
        If i > 100 Then
            i = 0
            DoEvents
        End If
    Loop
End Sub
Private Sub Form_Load()
    Randomize
    dlgFile.Filter = "Chaos Files (*.cha)|*.cha"
    dlgFile.CancelError = True
    dlgFile.InitDir = App.Path
End Sub

Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, _
        wid, ScaleHeight
End Sub

' Load a chaos data file.
Private Sub mnuFileOpen_Click()
Dim file_name As String

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
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
    LoadChaosData file_name, dlgFile.FileTitle
End Sub


