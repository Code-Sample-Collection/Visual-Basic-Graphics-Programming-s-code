VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDrawStyleSample 
      Height          =   375
      Index           =   0
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   720
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picDrawStyle 
      Height          =   375
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   720
      TabIndex        =   10
      Top             =   2880
      Width           =   780
   End
   Begin VB.PictureBox picFillStyleSample 
      Height          =   375
      Index           =   0
      Left            =   3360
      ScaleHeight     =   315
      ScaleWidth      =   720
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picFillStyle 
      Height          =   375
      Left            =   3360
      ScaleHeight     =   315
      ScaleWidth      =   720
      TabIndex        =   8
      Top             =   2880
      Width           =   780
   End
   Begin VB.PictureBox picDrawWidthSample 
      Height          =   375
      Index           =   0
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   720
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picDrawWidth 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   780
   End
   Begin VB.PictureBox picColorSamples 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   2520
      Width           =   615
      Begin VB.PictureBox picForeColorSample 
         AutoRedraw      =   -1  'True
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picFillColorSample 
         AutoRedraw      =   -1  'True
         Height          =   255
         Left            =   240
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.PictureBox picSwatch 
      Height          =   255
      Index           =   0
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin ComctlLib.Toolbar tbrButtons 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin ComctlLib.ImageList imlButtons 
      Left            =   960
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BitEdit.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BitEdit.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BitEdit.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BitEdit.frx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BitEdit.frx":0448
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BitEdit.frx":055A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BitEdit.frx":066C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BitEdit.frx":077E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Bitmap Files (*.bmp)|*.bmp"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuDrawWidth 
      Caption         =   "Draw&Width"
      Begin VB.Menu mnuDrawWidthSet 
         Caption         =   "1"
         Index           =   1
      End
      Begin VB.Menu mnuDrawWidthSet 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu mnuDrawWidthSet 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu mnuDrawWidthSet 
         Caption         =   "4"
         Index           =   4
      End
      Begin VB.Menu mnuDrawWidthSet 
         Caption         =   "5"
         Index           =   5
      End
   End
   Begin VB.Menu mnuDrawStyle 
      Caption         =   "Draw&Style"
      Begin VB.Menu mnuDrawStyleSet 
         Caption         =   "0"
         Index           =   0
      End
      Begin VB.Menu mnuDrawStyleSet 
         Caption         =   "1"
         Index           =   1
      End
      Begin VB.Menu mnuDrawStyleSet 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu mnuDrawStyleSet 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu mnuDrawStyleSet 
         Caption         =   "4"
         Index           =   4
      End
      Begin VB.Menu mnuDrawStyleSet 
         Caption         =   "5"
         Index           =   5
      End
   End
   Begin VB.Menu mnuFillStyle 
      Caption         =   "&FillStyle"
      Begin VB.Menu mnuFillStyleSet 
         Caption         =   "0"
         Index           =   0
      End
      Begin VB.Menu mnuFillStyleSet 
         Caption         =   "1"
         Index           =   1
      End
      Begin VB.Menu mnuFillStyleSet 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu mnuFillStyleSet 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu mnuFillStyleSet 
         Caption         =   "4"
         Index           =   4
      End
      Begin VB.Menu mnuFillStyleSet 
         Caption         =   "5"
         Index           =   5
      End
      Begin VB.Menu mnuFillStyleSet 
         Caption         =   "6"
         Index           =   6
      End
      Begin VB.Menu mnuFillStyleSet 
         Caption         =   "7"
         Index           =   7
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Tool variables.
Private Enum ToolTypes
    tool_Point = 1
    tool_Line
    tool_Rectangle
    tool_Ellipse
    tool_Scribble
    tool_Polyline
    tool_Undo
    tool_Redo
End Enum
Private SelectedTool As Integer

' Undo/redo variables.
Private Const NUM_UNDOS = 10
Private LastCheckpoint As Integer
Private Checkpoints As Collection

' File variables.
Private FileName As String
Private FileTitle As String

' Drawing variables.
Private Drawing As Boolean
Private FirstX As Single
Private FirstY As Single
Private LastX As Single
Private LastY As Single

Private DataModified As Boolean

' API stuff for putting bitmaps in menus.
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10

' See if it is safe to discard the data.
Private Function DataSafe() As Boolean
    If Not DataModified Then
        ' The data has not been modified. It's safe.
        DataSafe = True
    Else
        ' Ask the user if we should save changes.
        Select Case MsgBox("The data has been modified. Do you want to save the changes?", vbYesNoCancel)
            Case vbYes
                ' Save the data.
                'mnuFileSave_Click
                DataSafe = Not DataModified
            Case vbNo
                DataSafe = True
            Case vbCancel
                DataSafe = False
        End Select
    End If
End Function

' Draw a color sample.
Private Sub DrawSample()
    picFillColorSample.Line (0, 0)-(1000, 1000), picCanvas.FillColor, BF
    picForeColorSample.Line (0, 0)-(1000, 1000), picCanvas.ForeColor, BF
End Sub

' Draw the shape for the selected tool.
Private Sub DrawShape()
Dim cx As Single
Dim cy As Single
Dim wid As Single
Dim hgt As Single

    Select Case SelectedTool
        Case tool_Point
            picCanvas.PSet (LastX, LastY)
        Case tool_Line
            picCanvas.Line (FirstX, FirstY)-(LastX, LastY)
        Case tool_Rectangle
            picCanvas.Line (FirstX, FirstY)-(LastX, LastY), , B
        Case tool_Ellipse
            wid = Abs(LastX - FirstX)
            hgt = Abs(LastY - FirstY)
            If wid = 0 Or hgt = 0 Then Exit Sub
            cx = (FirstX + LastX) / 2
            cy = (FirstY + LastY) / 2
            If wid > hgt Then
                picCanvas.Circle (cx, cy), wid / 2, , , , hgt / wid
            Else
                picCanvas.Circle (cx, cy), hgt / 2, , , , hgt / wid
            End If
        Case tool_Scribble
            picCanvas.Line -(LastX, LastY)
        Case tool_Polyline
            picCanvas.Line (FirstX, FirstY)-(LastX, LastY)
    End Select
End Sub
' Set DataModified = True to indicate the data has
' been changed. Save the changes for undo/redo.
Private Sub SetModified()
    ' Update the caption if necessary.
    If Not DataModified Then Caption = "BitEdit*[" & FileTitle & "]"
    DataModified = True
End Sub
' Save the picture for undo/redo.
Private Sub SaveCheckpoint()
Dim new_picture As StdPicture
Dim i As Integer

    ' Get the next checkpoint index.
    LastCheckpoint = LastCheckpoint + 1

    ' Remove any checkpoints after the current one.
    Do While Checkpoints.Count >= LastCheckpoint
        Checkpoints.Remove Checkpoints.Count
    Loop

    ' See if we have too many stored.
    If LastCheckpoint > NUM_UNDOS Then
        ' Too many. Drop the oldest image.
        Checkpoints.Remove 1
        LastCheckpoint = LastCheckpoint - 1
    End If

    ' Save the current image.
    picCanvas.Picture = picCanvas.Image
    Set new_picture = New StdPicture
    Set new_picture = picCanvas.Picture
    Checkpoints.Add new_picture

    ' Enable and disable the undo buttons.
    SetUndoButtons
End Sub
' Enable the appropriate undo buttons.
Private Sub SetUndoButtons()
Dim enable_undo As Boolean
Dim enable_redo As Boolean

    enable_undo = (LastCheckpoint > 1)
    enable_redo = (LastCheckpoint < Checkpoints.Count)

    If enable_undo <> mnuEditUndo.Enabled Then
        tbrButtons.Buttons("Undo").Enabled = enable_undo
        mnuEditUndo.Enabled = enable_undo
    End If

    If enable_redo <> mnuEditRedo.Enabled Then
        tbrButtons.Buttons("Redo").Enabled = enable_redo
        mnuEditRedo.Enabled = enable_redo
    End If
End Sub

Private Sub Form_Load()
Dim btn As Button
Dim i As Integer
Dim tips(tool_Point To tool_Redo) As String
Dim pos As Single
Dim main_menu As Long
Dim sub_menu As Long
Dim menu_info As MENUITEMINFO

    dlgFile.InitDir = App.Path

    ' Load the tool tips.
    tips(tool_Point) = "Point"
    tips(tool_Line) = "Line"
    tips(tool_Rectangle) = "Rectangle"
    tips(tool_Ellipse) = "Ellipse"
    tips(tool_Scribble) = "Scribble"
    tips(tool_Polyline) = "Polyline"
    tips(tool_Undo) = "Undo"
    tips(tool_Redo) = "Redo"

    ' Load the tool buttons.
    tbrButtons.ImageList = imlButtons
    For i = tool_Point To tool_Redo
        Set btn = tbrButtons.Buttons.Add(, , , , i)
        btn.ToolTipText = tips(i)
        btn.Key = tips(i)
    Next i

    ' Create color swatches.
    For i = 0 To 15
        If i > 0 Then
            Load picSwatch(i)
            picSwatch(i).Visible = True
        End If
        picSwatch(i).BackColor = QBColor(i)
    Next i
    picColorSamples.Height = 2 * picSwatch(0).Height + 30
    picColorSamples.Width = picColorSamples.Height
    pos = picColorSamples.ScaleWidth * 0.1
    picForeColorSample.Move pos, pos
    pos = picColorSamples.ScaleWidth * 0.9 - picFillColorSample.Width
    picFillColorSample.Move pos, pos

    ' Create the DrawWidth menu.
    main_menu = GetMenu(hwnd)
    sub_menu = GetSubMenu(main_menu, 2)
    For i = 1 To 5
        Load picDrawWidthSample(i)
        picDrawWidthSample(i).AutoRedraw = True
        picDrawWidthSample(i).DrawWidth = i
        picDrawWidthSample(i).Line (-1000, picDrawWidthSample(0).ScaleHeight / 2)-Step(2000, 0)
        picDrawWidthSample(i).Picture = picDrawWidthSample(i).Image
        With menu_info
            .cbSize = Len(menu_info)
            .fMask = MIIM_TYPE
            .fType = MFT_BITMAP
            .dwTypeData = picDrawWidthSample(i).Picture
        End With
        SetMenuItemInfo sub_menu, i - 1, True, menu_info
    Next i
    ' Start with DrawWidth = 1.
    mnuDrawWidthSet_Click 1

    ' Create the DrawStyle menu.
    main_menu = GetMenu(hwnd)
    sub_menu = GetSubMenu(main_menu, 3)
    For i = 0 To 5
        If i > 0 Then Load picDrawStyleSample(i)
        picDrawStyleSample(i).AutoRedraw = True
        picDrawStyleSample(i).Line (0, 0)-(2000, 2000), picDrawStyleSample(0).BackColor, BF
        picDrawStyleSample(i).DrawStyle = i
        picDrawStyleSample(i).Line (-1000, picDrawStyleSample(0).ScaleHeight / 2)-Step(2000, 0)
        picDrawStyleSample(i).Picture = picDrawStyleSample(i).Image
        With menu_info
            .cbSize = Len(menu_info)
            .fMask = MIIM_TYPE
            .fType = MFT_BITMAP
            .dwTypeData = picDrawStyleSample(i).Picture
        End With
        SetMenuItemInfo sub_menu, i, True, menu_info
    Next i
    ' Start with Drawstyle = vbSolid.
    mnuDrawStyleSet_Click vbSolid

    ' Create the fillstyle menu.
    main_menu = GetMenu(hwnd)
    sub_menu = GetSubMenu(main_menu, 4)
    For i = 0 To 7
        If i > 0 Then Load picFillStyleSample(i)
        picFillStyleSample(i).AutoRedraw = True
        picFillStyleSample(i).FillStyle = vbFSSolid
        picFillStyleSample(i).Line (-1000, -1000)-(2000, 2000), picFillStyleSample(0).BackColor, BF
        picFillStyleSample(i).FillStyle = i
        picFillStyleSample(i).Line (-1000, -1000)-(2000, 2000), , B
        picFillStyleSample(i).Picture = picFillStyleSample(i).Image
        With menu_info
            .cbSize = Len(menu_info)
            .fMask = MIIM_TYPE
            .fType = MFT_BITMAP
            .dwTypeData = picFillStyleSample(i).Picture
        End With
        SetMenuItemInfo sub_menu, i, True, menu_info
    Next i
    ' Start with fillstyle = vbFSTransparent.
    mnuFillStyleSet_Click vbFSTransparent

    ' Start a new project.
    mnuFileNew_Click

    ' Draw the initial sample.
    DrawSample

    ' Select the point tool.
    tbrButtons_ButtonClick tbrButtons.Buttons(tool_Point)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not DataSafe
End Sub

Private Sub Form_Resize()
Dim hgt As Single
Dim t As Single
Dim i As Integer

    If WindowState = vbMinimized Then Exit Sub

    t = ScaleHeight - picColorSamples.Height
    picColorSamples.Top = t

    picSwatch(0).Move picColorSamples.Width + 30, picColorSamples.Top
    For i = 0 To 15
        picSwatch(i).Visible = False
    Next i
    For i = 1 To 15
        If i = 8 Then
            picSwatch(i).Left = picSwatch(0).Left
            picSwatch(i).Top = picSwatch(0).Top + picSwatch(0).Height + 30
        Else
            picSwatch(i).Left = picSwatch(i - 1).Left + picSwatch(i - 1).Width + 30
            picSwatch(i).Top = picSwatch(i - 1).Top
        End If
    Next i
    For i = 0 To 15
        picSwatch(i).Visible = True
    Next i

    hgt = picColorSamples.Top - tbrButtons.Height - 30
    If hgt <= 0 Then Exit Sub
    picCanvas.Move 0, tbrButtons.Height, ScaleWidth, hgt

    picDrawWidth.Move picSwatch(7).Left + _
        picSwatch(7).Width + 120, _
        picSwatch(7).Top
    picDrawStyle.Move picDrawWidth.Left + _
        picDrawWidth.Width + 120, _
        picDrawWidth.Top
    picFillStyle.Move picDrawStyle.Left + _
        picDrawStyle.Width + 120, _
        picDrawStyle.Top
End Sub

' Set the DrawStyle.
Private Sub mnuDrawStyleSet_Click(Index As Integer)
Dim i As Integer

    ' Check the selected style.
    For i = 0 To 5
        mnuDrawStyleSet(i).Checked = False
    Next i
    mnuDrawStyleSet(Index).Checked = True

    ' Display the selected style.
    picDrawStyle.Picture = picDrawStyleSample(Index).Picture

    ' Select the DrawStyle.
    picCanvas.DrawStyle = Index
End Sub

' Redo the previously undone command.
Private Sub mnuEditRedo_Click()
    LastCheckpoint = LastCheckpoint + 1
    picCanvas.Picture = Checkpoints(LastCheckpoint)
    SetUndoButtons

    ' Flag the data as modified.
    SetModified
End Sub

' Undo the previous command.
Private Sub mnuEditUndo_Click()
    LastCheckpoint = LastCheckpoint - 1
    picCanvas.Picture = Checkpoints(LastCheckpoint)

    ' Enable and disable the undo buttons.
    SetUndoButtons

    ' Flag the data as modified.
    SetModified
End Sub

' Unload the form. The QueryUnload event handler
' will make sure it's safe to do so.
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Start a new project.
Private Sub mnuFileNew_Click()
    ' Make sure the data is safe.
    If Not DataSafe() Then Exit Sub

    ' Start a new project.
    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), picCanvas.BackColor, BF

    ' Start a new Checkpoints collection.
    Set Checkpoints = New Collection
    LastCheckpoint = 0

    ' Checkpoint the blank project.
    SaveCheckpoint

    DataModified = False
    Caption = "BitEdit []"
    FileName = ""
    FileTitle = ""
End Sub

' Open a file.
Private Sub mnuFileOpen_Click()
    ' Make sure the data is safe.
    If Not DataSafe() Then Exit Sub

    ' Let the user select a file name.
    dlgFile.Flags = _
        cdlOFNExplorer + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNFileMustExist
    dlgFile.CancelError = True
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number > 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting the file." & _
            vbCrLf & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' Open the file.
    On Error GoTo OpenErr
    picCanvas.Picture = LoadPicture(dlgFile.FileName)

    ' Start a new Checkpoints collection.
    Set Checkpoints = New Collection
    LastCheckpoint = 0

    ' Checkpoint the new file.
    SaveCheckpoint

    ' Update the file name and title.
    FileName = dlgFile.FileName
    FileTitle = dlgFile.FileTitle
    Caption = "BitEdit [" & FileTitle & "]"
    DataModified = False
    Exit Sub

OpenErr:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file '" & dlgFile.FileName & "'." & _
        vbCrLf & Err.Description
    Exit Sub


    ' Update the file name and title.
    FileName = dlgFile.FileName
    FileTitle = dlgFile.FileTitle
    Caption = "BitEdit [" & FileTitle & "]"
    DataModified = False

End Sub

' Save the file.
Private Sub mnuFileSave_Click()
    ' If there is no file name, treat as Save As.
    If Len(FileName) = 0 Then
        mnuFileSaveAs_Click
        Exit Sub
    End If

    ' Save the file.
    On Error GoTo SaveErr
    SavePicture picCanvas.Picture, FileName

    ' Update the file name and title.
    FileName = dlgFile.FileName
    FileTitle = dlgFile.FileTitle
    Caption = "BitEdit [" & FileTitle & "]"
    DataModified = False
    Exit Sub

SaveErr:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file '" & FileName & "'." & _
        vbCrLf & Err.Description
    Exit Sub
End Sub
' Save the file with a new name.
Private Sub mnuFileSaveAs_Click()
    ' Let the user select a file name.
    dlgFile.Flags = _
        cdlOFNExplorer + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNOverwritePrompt + _
        cdlOFNPathMustExist
    dlgFile.CancelError = True
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number > 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting the file." & _
            vbCrLf & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' Save the file.
    On Error GoTo SaveAsErr
    SavePicture picCanvas.Picture, dlgFile.FileName

    ' Update the file name and title.
    FileName = dlgFile.FileName
    FileTitle = dlgFile.FileTitle
    Caption = "BitEdit [" & FileTitle & "]"
    DataModified = False
    Exit Sub

SaveAsErr:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file '" & dlgFile.FileName & "'." & _
        vbCrLf & Err.Description
    Exit Sub
End Sub

' Set the DrawWidth.
Private Sub mnuDrawWidthSet_Click(Index As Integer)
Dim i As Integer

    ' Check the selected width.
    For i = 1 To 5
        mnuDrawWidthSet(i).Checked = False
    Next i
    mnuDrawWidthSet(Index).Checked = True

    ' Display the selected width.
    picDrawWidth.Picture = picDrawWidthSample(Index).Picture

    ' Select the DrawWidth.
    picCanvas.DrawWidth = Index
End Sub

' Set the FillStyle.
Private Sub mnuFillStyleSet_Click(Index As Integer)
Dim i As Integer

    ' Check the selected style.
    For i = 0 To 7
        mnuFillStyleSet(i).Checked = False
    Next i
    mnuFillStyleSet(Index).Checked = True

    ' Display the selected style.
    picFillStyle.Picture = picFillStyleSample(Index).Picture

    ' Select the fillstyle.
    picCanvas.FillStyle = Index
End Sub

' Start doing something.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' See if we are ending a polyline.
    If (SelectedTool = tool_Polyline) And _
       (Button = vbRightButton) And Drawing _
    Then
        ' End the polyline.
        Drawing = False

        ' Erase the last segment.
        DrawShape

        ' Mark the data and save a checkpoint.
        SetModified
        SaveCheckpoint
        Exit Sub
    End If

    ' See if we are drawing a polyline.
    If (SelectedTool = tool_Polyline) And Drawing Then
        ' Finalize the segment.
        picCanvas.DrawMode = vbCopyPen
        DrawShape
    End If

    ' Deal with other situations.
    ' Save the coordinates.
    FirstX = X
    FirstY = Y
    LastX = X
    LastY = Y
    
    ' Prepare to draw in invert mode.
    If SelectedTool = tool_Scribble Then
        picCanvas.CurrentX = X
        picCanvas.CurrentY = Y
    ElseIf SelectedTool = tool_Polyline Then
        ' See if we are not already drawing.
        If Not Drawing Then
            ' Start the first segment here.
            picCanvas.CurrentX = X
            picCanvas.CurrentY = Y
        End If
        picCanvas.DrawMode = vbInvert
    Else
        picCanvas.DrawMode = vbInvert
    End If
    Drawing = True

    ' Draw the initial shape.
    DrawShape
End Sub
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Drawing Then Exit Sub

    ' Erase the previous shape.
    DrawShape

    LastX = X
    LastY = Y

    ' Draw the new shape.
    DrawShape
End Sub


Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Drawing Then Exit Sub

    ' Do nothing if we are drawing a polyline.
    ' All the interesting stuff happens in the
    ' MouseDown and MouseMove event handlers.
    If SelectedTool <> tool_Polyline Then
        Drawing = False

        ' Erase the previous shape.
        DrawShape

        LastX = X
        LastY = Y

        ' Draw the final shape.
        picCanvas.DrawMode = vbCopyPen
        DrawShape

        ' Mark the data and save a checkpoint.
        SetModified
        SaveCheckpoint
    End If
End Sub

' Display the DrawStyle popup.
Private Sub picDrawStyle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PopupMenu mnuDrawStyle
End Sub


' Display the DrawWidth popup.
Private Sub picDrawWidth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PopupMenu mnuDrawWidth
End Sub

' Display the FillStyle popup.
Private Sub picFillStyle_Click()
    PopupMenu mnuFillStyle
End Sub

' Select the new color.
Private Sub picSwatch_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        picCanvas.ForeColor = QBColor(Index)
    Else
        picCanvas.FillColor = QBColor(Index)
    End If

    ' Draw a new color sample.
    DrawSample
End Sub


' Process a toolbar button click.
Private Sub tbrButtons_ButtonClick(ByVal Button As ComctlLib.Button)
    ' See what kind of button this is.
    If Button.Index <= tool_Polyline Then
        ' This is a toggle button.
        ' Deselect the previously selected tool.
        If SelectedTool > 0 Then tbrButtons.Buttons(SelectedTool).Value = tbrUnpressed

        ' Select the new tool.
        SelectedTool = Button.Index
        tbrButtons.Buttons(SelectedTool).Value = tbrPressed
        tbrButtons.Refresh
    ElseIf Button.Index = tool_Undo Then
        ' Undo the previous command.
        mnuEditUndo_Click
    ElseIf Button.Index = tool_Redo Then
        ' Redo the previously undone command.
        mnuEditRedo_Click
    End If
End Sub
