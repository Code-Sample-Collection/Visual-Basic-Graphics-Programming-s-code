VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRobot 
   Caption         =   "Robot []"
   ClientHeight    =   4590
   ClientLeft      =   2040
   ClientTop       =   1035
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   Begin VB.TextBox txtFramesPerSecond 
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Text            =   "20"
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.OptionButton optPlay 
      Caption         =   "Reversing"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.OptionButton optPlay 
      Caption         =   "Looping"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.OptionButton optPlay 
      Caption         =   "Once"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.HScrollBar sbarFrame 
      Height          =   255
      Left            =   0
      Max             =   1
      Min             =   1
      TabIndex        =   1
      Top             =   3960
      Value           =   1
      Width           =   3255
   End
   Begin VB.PictureBox picCanvas 
      Height          =   3975
      Left            =   0
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3360
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Frames per Second"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblFrame 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1/1"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Frame:"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   4320
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFileFrames 
      Caption         =   "Frames"
      Begin VB.Menu mnuFrameAfter 
         Caption         =   "Insert &After"
      End
      Begin VB.Menu mnuFrameBefore 
         Caption         =   "Insert &Before"
      End
      Begin VB.Menu mnuFrameSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFrameDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmRobot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumFrames As Integer
Private Frames() As New Robot
Private SelectedFrame As Integer
Private SelectingFrame As Boolean
Private FileName As String
Private FileTitle As String
Private DataModified As Boolean
Private Playing As Boolean
Private NumPlayed As Long

Private Dragging As Boolean
Private DragPoint As Integer
Private DragX As Integer
Private DragY As Integer
Private AnchorX As Integer
Private AnchorY As Integer
' Convert (X, Y) into the point in the direction
' of (X, Y) that is the correct distance from the
' anchor point. For example, when dragging an
' elbow, the point should be UpperArmLength distance
' from the shoulders.
Private Sub AdjustPoint(x As Single, y As Single)
Dim dist As Single
Dim factor As Single
Dim dx As Single
Dim dy As Single

    ' Heads have no anchor point.
    If DragPoint = part_Head Then
        DragX = x
        DragY = y
        Exit Sub
    End If

    dx = x - AnchorX
    dy = y - AnchorY
    dist = Sqr(dx * dx + dy * dy)

    Select Case DragPoint
        Case part_Lelbow, part_RElbow
            factor = Frames(1).UpperArmLength / dist
        Case part_LHand, part_RHand
            factor = Frames(1).LowerArmLength / dist
        Case part_LKnee, part_RKnee
            factor = Frames(1).UpperLegLength / dist
        Case part_LFoot, part_RFoot
            factor = Frames(1).LowerLegLength / dist
    End Select

    DragX = AnchorX + dx * factor
    DragY = AnchorY + dy * factor
End Sub

' Return true if the data has not been modified,
' or the user has saved the changes, or the user
' wants to lose the changes.
Private Function DataSafe() As Boolean
Dim ans As Integer

    Do While DataModified
        ans = MsgBox("The data has been modified." & _
            " Do you want to save the changes?", _
            vbYesNoCancel)
        If ans = vbCancel Then Exit Do
        If ans = vbNo Then
            DataSafe = True
            Exit Function
        End If
            
        ' Otherwise save the data.
        If FileName <> "" Then
            mnuFileSave_Click
        Else
            mnuFileSaveAs_Click
        End If
    Loop
    
    DataSafe = Not DataModified
End Function

' Draw the highlight fot the drag.
Private Sub DrawDrag()
    If DragPoint = part_Head Then
        picCanvas.Line (DragX - Near, DragY - Near)-Step(Near2, Near2), , BF
    Else
        picCanvas.Line (AnchorX, AnchorY)-(DragX, DragY)
    End If
End Sub

' Draw the selected configuration.
Private Sub DrawSelected()
    picCanvas.Cls
    Frames(SelectedFrame).Draw picCanvas, True
End Sub



' Save a robot script into the file.
Private Sub SaveScript(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim i As Integer

    On Error GoTo SaveScriptError
    ' Open the file.
    fnum = FreeFile
    Open file_name For Output As fnum
    
    ' Write the number of frames.
    Write #fnum, NumFrames

    ' Write the parameters for each frame.
    For i = 1 To NumFrames
        Frames(i).FileWrite fnum
    Next i
    Close fnum
    
    FileName = file_name
    FileTitle = file_title
    DataModified = False
    Caption = "Robot [" & FileTitle & "]"
    Exit Sub
    
SaveScriptError:
    Beep
    MsgBox "Error saving file " & file_name & "." & _
        vbCrLf & Format$(Err.Number) & " : " & _
        Err.Description
    Exit Sub
End Sub

' Load a robot script from the file.
Private Sub LoadScript(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim i As Integer

    On Error GoTo LoadScriptError
    ' Open the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    
    ' Read the number of frames.
    Input #fnum, NumFrames
    ReDim Frames(1 To NumFrames)
    sbarFrame.Max = NumFrames
    
    ' Read the parameters for each frame.
    For i = 1 To NumFrames
        Frames(i).FileInput fnum
    Next i
    Close fnum
    
    SelectFrame 1
    
    mnuFrameDelete.Enabled = (NumFrames > 1)
    FileName = file_name
    FileTitle = file_title
    DataModified = False
    Caption = "Robot [" & FileTitle & "]"
    Exit Sub
    
LoadScriptError:
    Beep
    MsgBox "Error loading file " & file_name & "." & _
        vbCrLf & Format$(Err.Number) & " : " & _
        Err.Description
    Exit Sub
End Sub
' Select and display the indicated frame.
Private Sub SelectFrame(index As Integer)
    SelectedFrame = index
    
    SelectingFrame = True
    sbarFrame.Value = index
    SelectingFrame = False
    
    lblFrame.Caption = Format$(index) & _
        "/" & Format$(NumFrames)
    DrawSelected
End Sub


' Set the point that anchors the selected control
' point. For example, when moving a hand the
' corresponding elbow is the control point.
Private Sub SetAnchor()
    Select Case DragPoint
        Case part_Head  ' The head has no anchor.
            AnchorX = -1
        Case part_Lelbow, part_RElbow
            Frames(SelectedFrame).Position _
                part_Shoulders, AnchorX, AnchorY
        Case part_LHand
            Frames(SelectedFrame).Position _
                part_Lelbow, AnchorX, AnchorY
        Case part_RHand
            Frames(SelectedFrame).Position _
                part_RElbow, AnchorX, AnchorY
        Case part_LKnee, part_RKnee
            Frames(SelectedFrame).Position _
                part_Hips, AnchorX, AnchorY
        Case part_LFoot
            Frames(SelectedFrame).Position _
                part_LKnee, AnchorX, AnchorY
        Case part_RFoot
            Frames(SelectedFrame).Position _
                part_RKnee, AnchorX, AnchorY
    End Select
End Sub

' Grab the nearest control point within distance
' Near of the mouse.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
Dim best_i As Integer
Dim best_dist As Long
Dim dx As Long
Dim dy As Long
Dim dist As Long
Dim fx As Integer
Dim fy As Integer

    ' Find the closest control point.
    best_dist = Near + 1
    For i = part_MinPart To part_MaxControlPart
        Frames(SelectedFrame).Position i, fx, fy
        dx = x - fx
        dy = y - fy
        dist = Sqr(dx * dx + dy * dy)
        If best_dist > dist Then
            best_dist = dist
            best_i = i
        End If
    Next i
    
    ' If nothing is close enough, leave.
    If best_dist > Near Then
        Beep
        Exit Sub
    End If
    
    ' Begin moving the control point.
    Dragging = True
    DragPoint = best_i
    picCanvas.DrawMode = vbInvert
    SetAnchor
    DragX = x
    DragY = y
    DrawDrag
End Sub

' Continue dragging a control point.
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Dragging Then Exit Sub
    
    ' Erase the old highlight.
    DrawDrag
    
    ' Draw the new highlight.
    AdjustPoint x, y
    DrawDrag
End Sub


' Finish dragging the control point.
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Dragging Then Exit Sub
    Dragging = False

    ' Erase the old highlight.
    DrawDrag
    picCanvas.DrawMode = vbCopyPen

    ' Adjust the control point.
    AdjustPoint x, y
    Frames(SelectedFrame).MoveControlPoint _
        DragPoint, AnchorX, AnchorY, DragX, DragY
    DrawSelected
    DataModified = True
    Caption = "Robot*[" & FileTitle & "]"
End Sub

' Play the animation.
Private Sub cmdPlay_Click()
    If Playing Then
        Playing = False
        cmdPlay.Caption = "Stopped"
        cmdPlay.Enabled = False
    Else
        Playing = True
        cmdPlay.Caption = "Stop"
        PlayData
        cmdPlay.Caption = "Play"
        Playing = False
        cmdPlay.Enabled = True
        DrawSelected
    End If
End Sub

' Play the animation.
Private Sub PlayData()
Dim ms_per_frame As Long
Dim start_time As Single
Dim stop_time As Single

    ' See how fast we should go.
    If Not IsNumeric(txtFramesPerSecond.Text) Then _
        txtFramesPerSecond.Text = "10"
    ms_per_frame = 1000 \ CLng(txtFramesPerSecond.Text)

    ' See what kind of animation this should be.
    NumPlayed = 0
    start_time = Timer
    If optPlay(0).Value Then
        PlayDataOnce ms_per_frame
    ElseIf optPlay(1).Value Then
        PlayDataLooping ms_per_frame
    ElseIf optPlay(2).Value Then
        PlayDataBackAndForth ms_per_frame
    End If

    stop_time = Timer
    MsgBox "Displayed" & Str$(NumPlayed) & _
        " frames in " & _
        Format$(stop_time - start_time, "0.00") & _
        " seconds (" & _
        Format$(NumPlayed / (stop_time - start_time), "0.00") & _
        " FPS)."
End Sub
' Play the animation once.
Private Sub PlayDataOnce(ByVal ms_per_frame As Long)
Dim next_time As Long
Dim frame As Integer

    ' Start the animation.
    next_time = GetTickCount()

    ' Show the frames once.
    For frame = 1 To NumFrames
        If Not Playing Then Exit For
        NumPlayed = NumPlayed + 1

        ' Draw the frame.
        picCanvas.Cls
        Frames(frame).Draw picCanvas, False

        ' Wait until it's time for the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time
    Next frame
End Sub
' Play the animation in a loop.
Private Sub PlayDataLooping(ByVal ms_per_frame As Long)
    Do While Playing
        PlayDataOnce ms_per_frame
    Loop
End Sub
' Play the animation backward and forward.
Private Sub PlayDataBackAndForth(ByVal ms_per_frame As Long)
    Do While Playing
        PlayDataOnce ms_per_frame
        If Not Playing Then Exit Do
        PlayDataBackwards ms_per_frame
    Loop
End Sub

' Play the animation once backwards.
Private Sub PlayDataBackwards(ByVal ms_per_frame As Long)
Dim next_time As Long
Dim frame As Integer

    ' Start the animation.
    next_time = GetTickCount()

    ' Show the frames once.
    For frame = NumFrames To 1 Step -1
        If Not Playing Then Exit For
        NumPlayed = NumPlayed + 1

        ' Draw the frame.
        picCanvas.Cls
        Frames(frame).Draw picCanvas, False

        ' Wait until it's time for the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time
    Next frame
End Sub
Private Sub Form_Load()
    dlgFile.Filter = _
        "Robot Files (*.rob)|*.rob|" & _
        "All Files (*.*)|*.*"
        dlgFile.InitDir = App.Path

    ' Create a single default frame.
    NumFrames = 1
    ReDim Frames(1 To NumFrames)
    
    With Frames(1)
        .SetParameters _
            picCanvas.ScaleWidth / 2, _
            (picCanvas.ScaleHeight - .MaxHeight) / 2 + _
                .HeadRoom, _
            210, -30, 150, 30, 240, -60, 255, -75
    End With
    
    ' Position the scroll bar.
    sbarFrame.Top = picCanvas.Top + picCanvas.Height + 1
    
    SelectFrame 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not DataSafe()
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


' Load a robot script file.
Private Sub mnuFileOpen_Click()
Dim file_name As String

    If Not DataSafe() Then Exit Sub

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
    
    ' Load the robot script file.
    LoadScript file_name, dlgFile.FileTitle
End Sub

' Save the robot script file.
Private Sub mnuFileSave_Click()
    If FileName = "" Then
        mnuFileSaveAs_Click
        Exit Sub
    End If
    
    SaveScript FileName, FileTitle
End Sub

' Save the robot script file with a new name.
Private Sub mnuFileSaveAs_Click()
Dim file_name As String
    
    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    dlgFile.ShowSave
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

    ' Save the robot script file.
    SaveScript file_name, dlgFile.FileTitle
End Sub


' Insert a frame next to the selected one.
Private Sub AddFrame()
Dim i As Integer

    NumFrames = NumFrames + 1
    ReDim Preserve Frames(1 To NumFrames)
    For i = NumFrames - 1 To SelectedFrame Step -1
        Frames(i + 1).CopyFrame Frames(i)
    Next i

    sbarFrame.Max = NumFrames

    mnuFrameDelete.Enabled = (NumFrames > 1)
End Sub



' Insert a frame after the selected one.
Private Sub mnuFrameAfter_Click()
    AddFrame
    SelectFrame SelectedFrame + 1
End Sub


' Insert a frame before the selected one.
Private Sub mnuFrameBefore_Click()
    AddFrame
End Sub


' Delete the selected frame.
Private Sub mnuFrameDelete_Click()
Dim i As Integer
    
    For i = SelectedFrame To NumFrames - 1
        Frames(i).CopyFrame Frames(i + 1)
    Next i

    NumFrames = NumFrames - 1
    ReDim Preserve Frames(1 To NumFrames)

    sbarFrame.Max = NumFrames

    If SelectedFrame > NumFrames Then _
       SelectedFrame = NumFrames
    SelectFrame SelectedFrame

    mnuFrameDelete.Enabled = (NumFrames > 1)
End Sub

' Repaint the current frame.
Private Sub picCanvas_Paint()
    If SelectingFrame Then Exit Sub
    SelectFrame sbarFrame.Value
End Sub

' Select a new frame.
Private Sub sbarFrame_Change()
    If SelectingFrame Then Exit Sub
    SelectFrame sbarFrame.Value
End Sub


' Select a new frame.
Private Sub sbarFrame_Scroll()
    sbarFrame_Change
End Sub
