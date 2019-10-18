VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTweenSmo 
   Caption         =   "TweenSmo"
   ClientHeight    =   4590
   ClientLeft      =   2040
   ClientTop       =   1035
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   Begin VB.CommandButton cmdTween 
      Caption         =   "Tween"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtNumTweens 
      Height          =   285
      Left            =   4200
      TabIndex        =   10
      Text            =   "4"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtFramesPerSecond 
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Text            =   "20"
      Top             =   1770
      Width           =   375
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.OptionButton optPlay 
      Caption         =   "Reversing"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton optPlay 
      Caption         =   "Looping"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton optPlay 
      Caption         =   "Once"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
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
      Left            =   2640
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Tweens:"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   11
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Frames per Second"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblFrame 
      Alignment       =   2  'Center
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
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFrame 
      Caption         =   "Frame"
      Begin VB.Menu mnuFrameAfter 
         Caption         =   "Insert &After"
      End
      Begin VB.Menu mnuFrameBefore 
         Caption         =   "Insert &Before"
      End
      Begin VB.Menu mnuFrameSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFrameClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuFrameDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmTweenSmo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumFrames As Integer
Private Frames() As PolylineFrame
Private FileName As String
Private FileTitle As String
Private DataModified As Boolean
Private Playing As Boolean
Private NumPlayed As Long
Private SelectedFrame As Integer
Private SelectingFrame As Boolean

Private Drawing As Boolean
Private StartX As Integer
Private StartY As Integer
Private LastX As Integer
Private LastY As Integer

Private Type Polyline
    NumPoints As Integer
    X() As Integer
    Y() As Integer
End Type

Private Type PolylineFrame
    NumPolylines As Integer
    Poly() As Polyline
End Type
' Insert a frame next to the selected one.
Private Sub AddFrame()
Dim i As Integer

    NumFrames = NumFrames + 1
    ReDim Preserve Frames(1 To NumFrames)
    For i = NumFrames - 1 To SelectedFrame Step -1
        CopyFrame i, i + 1
    Next i

    sbarFrame.Max = NumFrames

    mnuFrameDelete.Enabled = (NumFrames > 1)
    DataModified = True
    Caption = "TweenSmo*[" & FileTitle & "]"
End Sub


' Copy a polyline from frame1 to frame2.
Private Sub CopyFrame(frame1 As Integer, frame2 As Integer)
Dim pline As Integer
Dim point As Integer

    Frames(frame2).NumPolylines = Frames(frame1).NumPolylines
    If Frames(frame2).NumPolylines < 1 Then
        Erase Frames(frame2).Poly
    Else
        ReDim Frames(frame2).Poly(1 To Frames(frame2).NumPolylines)
    End If
    For pline = 1 To Frames(frame2).NumPolylines
        With Frames(frame2).Poly(pline)
            .NumPoints = Frames(frame1).Poly(pline).NumPoints
            If .NumPoints < 1 Then
                Erase .X
                Erase .Y
            Else
                ReDim .X(1 To .NumPoints)
                ReDim .Y(1 To .NumPoints)
            End If
            For point = 1 To .NumPoints
                .X(point) = Frames(frame1).Poly(pline).X(point)
                .Y(point) = Frames(frame1).Poly(pline).Y(point)
            Next point
        End With
    Next pline
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


' Draw the indicated frame.
Private Sub DrawFrame(frame As Integer)
Dim pline As Integer
Dim point As Integer

    picCanvas.Cls

    For pline = 1 To Frames(frame).NumPolylines
        With Frames(frame).Poly(pline)
            If .NumPoints >= 2 Then
                picCanvas.Line (.X(1), .Y(1))-(.X(2), .Y(2))
                For point = 3 To .NumPoints
                    picCanvas.Line -(.X(point), .Y(point))
                Next point
            End If
        End With
    Next pline
End Sub


' Save the data.
Private Sub SaveData(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim frame As Integer
Dim pline As Integer
Dim point As Integer

    On Error GoTo SaveDataError
    ' Open the file.
    fnum = FreeFile
    Open file_name For Output As fnum
    
    ' Save the number of frames.
    Write #fnum, NumFrames
    
    ' Save each frame.
    For frame = 1 To NumFrames
        With Frames(frame)
            ' Save the number of polylines.
            Write #fnum, .NumPolylines
                    
            ' Save each polyline.
            For pline = 1 To .NumPolylines
                With .Poly(pline)
                    ' Save the number of points.
                    Write #fnum, .NumPoints
                    For point = 1 To .NumPoints
                        Write #fnum, .X(point), .Y(point)
                    Next point
                End With
            Next pline
        End With
    Next frame
    Close fnum

    FileName = file_name
    FileTitle = file_title
    Caption = "TweenSmo [" & FileTitle & "]"
    DataModified = False
    Exit Sub
    
SaveDataError:
    Beep
    MsgBox "Error saving file " & file_name & "." & _
        vbCrLf & Format$(Err.Number) & " : " & _
        Err.Description
    Exit Sub
End Sub

' Load polyline frames from the file.
Private Sub LoadData(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim frame As Integer
Dim pline As Integer
Dim point As Integer

    On Error GoTo SaveDataError
    ' Open the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    
    ' Read the number of frames.
    Input #fnum, NumFrames
    ReDim Frames(1 To NumFrames)
    sbarFrame.Max = NumFrames
    
    ' Read each frame.
    For frame = 1 To NumFrames
        With Frames(frame)
            ' Read the number of polylines.
            Input #fnum, .NumPolylines
            ReDim .Poly(1 To .NumPolylines)
                    
            ' Read each polyline.
            For pline = 1 To .NumPolylines
                With .Poly(pline)
                    ' Read the number of points.
                    Input #fnum, .NumPoints
                    ReDim .X(1 To .NumPoints)
                    ReDim .Y(1 To .NumPoints)
                    For point = 1 To .NumPoints
                        Input #fnum, .X(point), .Y(point)
                    Next point
                End With
            Next pline
        End With
    Next frame
    Close fnum
    
    SelectFrame 1
    
    FileName = file_name
    FileTitle = file_title
    Caption = "TweenSmo [" & FileTitle & "]"
    DataModified = False
    Exit Sub
    
SaveDataError:
    Beep
    MsgBox "Error loading file " & file_name & "." & _
        vbCrLf & Format$(Err.Number) & " : " & _
        Err.Description
    Exit Sub
End Sub
' Select and display the indicated frame.
Private Sub SelectFrame(num As Integer)
    SelectedFrame = num
    
    ' If we're drawing, stop drawing.
    If Drawing Then
        picCanvas.DrawMode = vbCopyPen
        Drawing = False
    End If
    
    DrawFrame SelectedFrame
    
    lblFrame.Caption = Format$(SelectedFrame) _
         & "/" & Format$(NumFrames)
    
    SelectingFrame = True
    sbarFrame.Value = SelectedFrame
    SelectingFrame = False
End Sub


' Create the tweens between two key frames using
' Hermite curves.
Private Sub MakeTweens(ByVal key2 As Integer, ByVal key3 As Integer)
Dim tween As Integer
Dim pline As Integer
Dim point As Integer
Dim key1 As Integer
Dim key4 As Integer
Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
Dim x3 As Integer
Dim y3 As Integer
Dim x4 As Integer
Dim y4 As Integer
Dim dx1 As Integer
Dim dy1 As Integer
Dim dx2 As Integer
Dim dy2 As Integer
Dim t As Single
Dim t2 As Single
Dim t3 As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single

    ' Make room for the points.
    For tween = key2 + 1 To key3 - 1
        Frames(tween).NumPolylines = Frames(key2).NumPolylines
        ReDim Frames(tween).Poly(1 To Frames(tween).NumPolylines)
        For pline = 1 To Frames(tween).NumPolylines
            With Frames(tween).Poly(pline)
                .NumPoints = Frames(key2).Poly(pline).NumPoints
                ReDim .X(1 To .NumPoints)
                ReDim .Y(1 To .NumPoints)
            End With
        Next pline
    Next tween
    
    ' For each endpoint, create the tween endpoints.
    For pline = 1 To Frames(key2).NumPolylines
        With Frames(key2).Poly(pline)
            For point = 1 To .NumPoints
                ' Pick slopes for the start & end.
                If key2 > 1 Then
                    key1 = key2 - (key3 - key2)
                Else
                    key1 = key2
                End If
                x1 = Frames(key1).Poly(pline).X(point)
                y1 = Frames(key1).Poly(pline).Y(point)
                x2 = .X(point)
                y2 = .Y(point)
                x3 = Frames(key3).Poly(pline).X(point)
                y3 = Frames(key3).Poly(pline).Y(point)
                If key3 < NumFrames Then
                    key4 = key3 + (key3 - key2)
                Else
                    key4 = key3
                End If
                x4 = Frames(key4).Poly(pline).X(point)
                y4 = Frames(key4).Poly(pline).Y(point)
                dx1 = x3 - x1
                dy1 = y3 - y1
                dx2 = x4 - x2
                dy2 = y4 - y2
                ' Compute the Hermite values.
                For tween = key2 + 1 To key3 - 1
                    t = (tween - key2) / (key3 - key2)
                    t2 = t * t
                    t3 = t * t2
                    A = 2 * t3 - 3 * t2 + 1
                    B = -2 * t3 + 3 * t2
                    C = t3 - 2 * t2 + t
                    D = t3 - t2
                    Frames(tween).Poly(pline).X(point) = x2 * A + x3 * B + dx1 * C + dx2 * D
                    Frames(tween).Poly(pline).Y(point) = y2 * A + y3 * B + dy1 * C + dy2 * D
                Next tween
            Next point
        End With
    Next pline
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Drawing And Button = vbRightButton Then
        ' End the previous polyline.
        picCanvas.Line (StartX, StartY)-(LastX, LastY)
        picCanvas.DrawMode = vbCopyPen
        Drawing = False
        Exit Sub
    End If
    
    ' See if this is the start of a new polyline.
    If Drawing Then
        ' Nope. Erase the previous line.
        picCanvas.Line (StartX, StartY)-(LastX, LastY)
    Else
        ' Start a new polyline.
        With Frames(SelectedFrame)
            .NumPolylines = .NumPolylines + 1
            ReDim Preserve .Poly(1 To .NumPolylines)
            With .Poly(.NumPolylines)
                .NumPoints = 1
                ReDim .X(1 To 1)
                ReDim .Y(1 To 1)
                .X(1) = X
                .Y(1) = Y
            End With
        End With
        picCanvas.DrawMode = vbInvert
        Drawing = True
        DataModified = True
        Caption = "TweenSmo*[" & FileTitle & "]"
        StartX = X
        StartY = Y
    End If
    
    LastX = X
    LastY = Y
    picCanvas.Line (StartX, StartY)-(LastX, LastY)
End Sub
' Repaint the current frame.
Private Sub picCanvas_Paint()
    If SelectingFrame Then Exit Sub
    SelectFrame sbarFrame.Value
End Sub
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Drawing Then Exit Sub
    
    picCanvas.Line (StartX, StartY)-(LastX, LastY)
    LastX = X
    LastY = Y
    picCanvas.Line (StartX, StartY)-(LastX, LastY)
End Sub


Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Drawing Then Exit Sub
    
    picCanvas.Line (StartX, StartY)-(LastX, LastY)
    picCanvas.DrawMode = vbCopyPen
    picCanvas.Line (StartX, StartY)-(X, Y)
    picCanvas.DrawMode = vbInvert

    With Frames(SelectedFrame)
        With .Poly(.NumPolylines)
            .NumPoints = .NumPoints + 1
            ReDim Preserve .X(1 To .NumPoints)
            ReDim Preserve .Y(1 To .NumPoints)
            .X(.NumPoints) = X
            .Y(.NumPoints) = Y
        End With
    End With

    DataModified = True
    Caption = "TweenSmo*[" & FileTitle & "]"
    StartX = X
    StartY = Y
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
        DrawFrame SelectedFrame
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
Dim frame As Integer
Dim next_time As Long

    ' Start the animation.
    next_time = GetTickCount()
    For frame = 1 To NumFrames
        If Not Playing Then Exit For
        NumPlayed = NumPlayed + 1

        ' Draw the frame.
        DrawFrame frame

        ' Wait until it's time for the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time
    Next frame
End Sub
' Play the animation backwards.
Private Sub PlayDataBackward(ByVal ms_per_frame As Long)
Dim frame As Integer
Dim next_time As Long

    ' Start the animation.
    next_time = GetTickCount()
    For frame = NumFrames To 1 Step -1
        If Not Playing Then Exit For
        NumPlayed = NumPlayed + 1

        ' Draw the frame.
        DrawFrame frame

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
' Play the animation back and forth.
Private Sub PlayDataBackAndForth(ByVal ms_per_frame As Long)
    Do While Playing
        PlayDataOnce ms_per_frame
        If Not Playing Then Exit Do
        PlayDataBackward ms_per_frame
    Loop
End Sub

' Make the tweens.
Private Sub cmdTween_Click()
Dim num_tweens As Integer
Dim old_frames As Integer
Dim frame1 As Integer
Dim frame2 As Integer
Dim frame As Integer

    ' See how many tweens to make.
    If Not IsNumeric(txtNumTweens.Text) Then _
        txtNumTweens.Text = "4"
    num_tweens = txtNumTweens.Text
    If num_tweens < 1 Then num_tweens = 1
    
    ' Make room for the new frames.
    old_frames = NumFrames
    NumFrames = num_tweens * (NumFrames - 1) + NumFrames
    ReDim Preserve Frames(1 To NumFrames)
    
    ' Spread the original frames out.
    For frame = old_frames To 2 Step -1
        CopyFrame frame, _
            num_tweens * (frame - 1) + frame
    Next frame

    ' Make the tweens.
    For frame = 1 To old_frames - 1
        frame1 = num_tweens * (frame - 1) + frame
        frame2 = frame1 + num_tweens + 1
        MakeTweens frame1, frame2
    Next frame

    sbarFrame.Max = NumFrames
    SelectFrame num_tweens * (SelectedFrame - 1) + _
        SelectedFrame
    DataModified = True
    Caption = "TweenSmo*[" & FileTitle & "]"
End Sub



Private Sub Form_Load()
    ' Position the scroll bar.
    sbarFrame.Top = picCanvas.Top + picCanvas.Height + 1

    ' Create an empty frame.
    mnuFileNew_Click

    dlgFile.InitDir = App.Path
    dlgFile.Filter = _
        "Tween Files (*.twe)|*.twe|" & _
        "All Files (*.*)|*.*"
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


' Load a data file.
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

    ' Load the data file.
    LoadData file_name, dlgFile.FileTitle

    lblFrame.Caption = Format$(SelectedFrame) _
         & "/" & Format$(NumFrames)
End Sub

' Clear out all the data.
Private Sub mnuFileNew_Click()
    If Not DataSafe() Then Exit Sub
    
    NumFrames = 1
    ReDim Frames(1 To NumFrames)
    Frames(1).NumPolylines = 0
    sbarFrame.Max = NumFrames
    SelectFrame 1
End Sub

' Save the data file.
Private Sub mnuFileSave_Click()
    If FileName = "" Then
        mnuFileSaveAs_Click
        Exit Sub
    End If
    
    SaveData FileName, FileTitle
End Sub

' Save the data file with a new name.
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

    ' Save the script file.
    SaveData file_name, dlgFile.FileTitle
End Sub





' Insert a frame after the selected one.
Private Sub mnuFrameAfter_Click()
    AddFrame
    SelectFrame SelectedFrame + 1
End Sub
' Insert a frame before the selected one.
Private Sub mnuFrameBefore_Click()
    AddFrame
    lblFrame.Caption = Format$(SelectedFrame) & "/" & Format$(NumFrames)
End Sub

' Remove the polylines from the selected frame.
Private Sub mnuFrameClear_Click()
Dim i As Integer
    
    With Frames(SelectedFrame)
        .NumPolylines = 0
        Erase .Poly
    End With
    
    SelectFrame SelectedFrame

    DataModified = True
    Caption = "TweenSmo*[" & FileTitle & "]"
End Sub

' Delete the selected frame.
Private Sub mnuFrameDelete_Click()
Dim i As Integer

    For i = SelectedFrame To NumFrames - 1
        CopyFrame i + 1, i
    Next i

    NumFrames = NumFrames - 1
    ReDim Preserve Frames(1 To NumFrames)

    sbarFrame.Max = NumFrames

    If SelectedFrame > NumFrames Then _
       SelectedFrame = NumFrames
    SelectFrame SelectedFrame

    mnuFrameDelete.Enabled = (NumFrames > 1)
    DataModified = True
    Caption = "TweenSmo*[" & FileTitle & "]"
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
