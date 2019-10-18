VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlayWait 
   Caption         =   "PlayWait"
   ClientHeight    =   3825
   ClientLeft      =   1680
   ClientTop       =   975
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   Begin VB.TextBox txtNumFrames 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "100"
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optRunType 
      Caption         =   "Looping"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton optRunType 
      Caption         =   "Reversing"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton optRunType 
      Caption         =   "One time"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtFramesPerSecond 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "20"
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1560
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox picCanvas 
      Height          =   3810
      Left            =   2040
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   0
      Width           =   3810
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   1560
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Frames to load:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Frames per second:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblResults 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmPlayWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumImages As Integer
Private MaxImage As Integer
Private Playing As Boolean
Private NumPlayed As Long
' Load the images.
Private Sub LoadImages(file_name As String)
Dim base As String
Dim i As Integer

    ' Get the base file name.
    base = Left$(file_name, Len(file_name) - 5)

    ' See how many frames the user wants to load.
    If Not IsNumeric(txtNumFrames.Text) Then _
        txtNumFrames.Text = Format$(10)
    NumImages = CInt(txtNumFrames.Text)

    ' Create any needed picture boxes.
    For i = MaxImage + 1 To NumImages - 1
        Load picFrame(i)
    Next i

    ' Get rid of any that are no longer needed.
    For i = NumImages To MaxImage
        Unload picFrame(i)
    Next i
    MaxImage = NumImages - 1
    
    ' Load the images.
    On Error GoTo LoadPictureError
    i = 0
    Do While i < NumImages
        lblResults.Caption = Format$(i + 1)
        lblResults.Refresh
        picFrame(i).Picture = LoadPicture(base & Format$(i) & ".bmp")
        i = i + 1
    Loop

    picCanvas.AutoSize = True
    picCanvas.Picture = picFrame(0).Image
    picCanvas.AutoSize = False
    lblResults.Caption = ""
    txtNumFrames.Text = Format$(NumImages)
    Exit Sub
    
LoadPictureError:
    ' We ran out of images early.
    NumImages = i
    txtNumFrames.Text = Format$(NumImages)
    Resume Next
End Sub

' Run the animation until Playing is false.
Private Sub PlayImages()
Dim ms_per_frame As Integer
Dim start_time As Long
Dim stop_time As Long

    ' See how long it should be between frames.
    If Not IsNumeric(txtFramesPerSecond.Text) Then _
        txtFramesPerSecond.Text = "20"
    ms_per_frame = 1000 / CInt(txtFramesPerSecond.Text)

    ' Start the appropriate animation.
    NumPlayed = 0
    start_time = GetTickCount
    If optRunType(0).Value Then
        PlayImagesOnce ms_per_frame
    ElseIf optRunType(1).Value Then
        PlayImagesBackAndForth ms_per_frame
    Else
        PlayImagesLooping ms_per_frame
    End If

    ' Display results.
    stop_time = GetTickCount
    lblResults.Caption = _
        Format$(NumPlayed) & " frames/" & _
        Format$((stop_time - start_time) / 1000#, "0.00") & _
        " sec" & vbCrLf & vbCrLf & _
        Format$(CSng(NumPlayed) / ((stop_time - start_time) / 1000#), "0.00") & _
        " frames/sec"
End Sub
' Run the animation forward and backward until
' Playing is False.
Private Sub PlayImagesBackAndForth(ByVal ms_per_frame As Integer)
    ' Start the animation.
    Do While Playing
        PlayImagesOnce ms_per_frame
        If Not Playing Then Exit Do
        PlayImagesReversed ms_per_frame
    Loop
End Sub
' Run the animation until Playing is false.
Private Sub PlayImagesLooping(ByVal ms_per_frame As Integer)
    ' Start the animation.
    Do While Playing
        PlayImagesOnce ms_per_frame
    Loop
End Sub
' Run the animation once or until Playing is False.
Private Sub PlayImagesOnce(ByVal ms_per_frame As Integer)
Dim i As Integer
Dim next_time As Long

    ' Get the current time.
    next_time = GetTickCount

    ' Start the animation.
    For i = 0 To NumImages - 1
        ' Display the next frame.
        picCanvas.Picture = picFrame(i).Picture
        NumPlayed = NumPlayed + 1

        ' Wait till we should display the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        If Not Playing Then Exit For
    Next i
End Sub
' Run the animation reversed once or until Playing
' is False.
Private Sub PlayImagesReversed(ByVal ms_per_frame As Integer)
Dim i As Integer
Dim next_time As Long

    ' Get the current time.
    next_time = GetTickCount

    ' Start the animation.
    For i = NumImages - 1 To 0 Step -1
        ' Display the next frame.
        picCanvas.Picture = picFrame(i).Picture
        NumPlayed = NumPlayed + 1

        ' Wait till we should display the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        If Not Playing Then Exit For
    Next i
End Sub

' Start or stop playing.
Private Sub CmdStart_Click()
    If Playing Then
        Playing = False
        cmdStart.Caption = "Stopped"
        cmdStart.Enabled = False
    Else
        cmdStart.Caption = "Stop"
        lblResults.Caption = ""
        DoEvents
        Playing = True
        PlayImages
        Playing = False
        cmdStart.Caption = "Start"
        cmdStart.Enabled = True
    End If
End Sub






Private Sub Form_Load()
    dlgOpenFile.InitDir = App.Path
End Sub

' Load new image files.
Private Sub mnuFileOpen_Click()
Dim file_name As String

    ' Let the user select a file.
    On Error Resume Next
    dlgOpenFile.FileName = "*_0.BMP"
    dlgOpenFile.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    dlgOpenFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
    DoEvents

    file_name = Trim$(dlgOpenFile.FileName)
    dlgOpenFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgOpenFile.FileTitle) - 1)
    Caption = "PlayWait [" & dlgOpenFile.FileTitle & "]"

    ' Load the pictures.
    On Error GoTo LoadError
    LoadImages file_name
    On Error GoTo 0

    cmdStart.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub

LoadError:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Format$(Err.Number) & _
        " opening file '" & file_name & "'" & vbCrLf & _
        Err.Description
End Sub
