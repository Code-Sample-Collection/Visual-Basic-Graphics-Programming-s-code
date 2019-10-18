VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PlanetForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planets"
   ClientHeight    =   5775
   ClientLeft      =   1575
   ClientTop       =   720
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   Begin VB.TextBox txtFramesPerSecond 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "20"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   5280
      Width           =   855
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   5250
      Left            =   0
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3120
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Frames per second:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "PlanetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Playing As Boolean

Private NumPlanets As Integer

Private Cx() As Double          ' Position.
Private Cy() As Double
Private Vx() As Double          ' Velocity.
Private Vy() As Double
Private M() As Double           ' Mass.
Private R() As Double           ' Radius.
Private Clr() As Long           ' Colors.

Private BitmapWid As Long
Private BitmapHgt As Long
Private BitmapNumBytes As Long
Private Bytes() As Byte

' Draw some random rectangles on the bacground.
Private Sub DrawBackground()
Dim X As Single
Dim Y As Single
Dim xmax As Single
Dim ymax As Single
Dim i As Integer

    ' Start with a clean slate.
    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), vbBlack, BF

    ' Draw some "stars."
    xmax = picCanvas.ScaleWidth
    ymax = picCanvas.ScaleHeight
    For i = 1 To 100
        X = Rnd * xmax
        Y = Rnd * ymax
        picCanvas.PSet (X, Y), vbWhite
    Next i

    ' Make the background permanent.
    picCanvas.Picture = picCanvas.Image
End Sub
' Load the data in a planet file.
Private Sub LoadPlanets(file_name As String)
Dim fnum As Integer
Dim i As Integer
Dim old_style As Integer
Dim bm As BITMAP

    ' Make a random background.
    DrawBackground

    ' Save the background bitmap data.
    GetObject picCanvas.Image, Len(bm), bm
    BitmapWid = bm.bmWidthBytes
    BitmapHgt = bm.bmHeight
    BitmapNumBytes = BitmapWid * BitmapHgt
    ReDim Bytes(1 To bm.bmWidthBytes, 1 To bm.bmHeight)
    GetBitmapBits picCanvas.Image, BitmapNumBytes, Bytes(1, 1)

    ' Load the data.
    fnum = FreeFile
    Open file_name For Input As #fnum
        
    Input #fnum, NumPlanets
    ReDim Cx(1 To NumPlanets)
    ReDim Cy(1 To NumPlanets)
    ReDim Vx(1 To NumPlanets)
    ReDim Vy(1 To NumPlanets)
    ReDim M(1 To NumPlanets)
    ReDim R(1 To NumPlanets)
    ReDim Clr(1 To NumPlanets)
    
    For i = 1 To NumPlanets
        Input #fnum, _
            Cx(i), Cy(i), Vx(i), Vy(i), M(i), Clr(i)
        R(i) = Sqr(M(i)) + 1
    Next i
        
    Close #fnum
    
    ' Draw the planets.
    old_style = picCanvas.FillStyle
    picCanvas.FillStyle = vbSolid
    picCanvas.Cls
    For i = 1 To NumPlanets
        picCanvas.FillColor = Clr(i)
        picCanvas.Circle (Cx(i), Cy(i)), R(i), Clr(i)
    Next i
    picCanvas.FillStyle = old_style

    Caption = "Planets [" & file_name & "]"
    cmdRun.Enabled = True
End Sub
' Make the planets move until Playing is false.
Private Sub RunSimulation()
Dim ms_per_frame As Long

    ' See how fast we should go.
    If Not IsNumeric(txtFramesPerSecond.Text) Then _
        txtFramesPerSecond.Text = "10"
    ms_per_frame = 1000 \ CLng(txtFramesPerSecond.Text)

    PlayImages ms_per_frame
End Sub

' Make the planets move until Playing is false.
Private Sub PlayImages(ByVal ms_per_frame As Long)
Const F_SCALE = 1000

Dim next_time As Long
Dim i As Integer
Dim j As Integer
Dim dx As Double
Dim dy As Double
Dim d2 As Double
Dim d As Double
Dim f As Double
Dim a_d As Double
    
    ' Start the animation.
    next_time = GetTickCount()
    Do While Playing
        ' Calculate the forces on the planets.
        For i = 1 To NumPlanets - 1
            For j = i + 1 To NumPlanets
                ' Calculate the force between planets
                ' i and j. Translate the force into a
                ' change in velocity.
                dx = Cx(i) - Cx(j)
                dy = Cy(i) - Cy(j)
                d2 = dx * dx + dy * dy
                f = F_SCALE * M(i) * M(j) / d2
                d = Sqr(d2)
                            
                a_d = f / M(i) / d
                Vx(i) = Vx(i) - a_d * dx
                Vy(i) = Vy(i) - a_d * dy
            
                a_d = f / M(j) / d
                Vx(j) = Vx(j) + a_d * dx
                Vy(j) = Vy(j) + a_d * dy
            Next j
        Next i
        
        ' Move all the planets.
        For i = 1 To NumPlanets
            Cx(i) = Cx(i) + Vx(i)
            Cy(i) = Cy(i) + Vy(i)
        Next i
           
        ' Restore the background.
        SetBitmapBits picCanvas.Image, BitmapNumBytes, Bytes(1, 1)

        ' Redraw the planets.
        For i = 1 To NumPlanets
            picCanvas.FillColor = Clr(i)
            picCanvas.Circle (Cx(i), Cy(i)), R(i), Clr(i)
        Next i

        ' Wait until it's time for the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time
    Loop
End Sub


' Start a new simulation.
Private Sub cmdRun_Click()
    If Playing Then
        cmdRun.Caption = "Stopped"
        cmdRun.Enabled = False
        Playing = False
    Else
        Playing = True
        cmdRun.Caption = "Stop"
        RunSimulation
        cmdRun.Caption = "Run"
        cmdRun.Enabled = True
        Playing = False
    End If
End Sub



Private Sub Form_Load()
    picCanvas.FillStyle = vbSolid
    dlgFile.InitDir = App.Path
    dlgFile.Filter = _
        "Planet Files (*.pla)|*.pla|" & _
        "All Files (*.*)|*.*"
End Sub

Private Sub Form_Resize()
Const GAP = 3

Dim hgt As Double

    hgt = ScaleHeight - cmdRun.Height - 2 * GAP
    picCanvas.Move 0, 0, ScaleWidth, hgt
    cmdRun.Move (ScaleWidth - cmdRun.Width) / 2, _
        picCanvas.Height + GAP
    Label1.Top = cmdRun.Top
    txtFramesPerSecond.Top = cmdRun.Top
End Sub
' Load a new data file.
Private Sub mnuFileOpen_Click()
Dim fname As String

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
    
    fname = Trim$(dlgFile.FileName)
    dlgFile.InitDir = Left$(fname, Len(fname) _
        - Len(dlgFile.FileTitle) - 1)

    ' Load the data.
    LoadPlanets fname
End Sub
