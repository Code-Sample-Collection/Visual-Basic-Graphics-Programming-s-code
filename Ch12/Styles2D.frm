VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStyles2D 
   Caption         =   "Styles2D"
   ClientHeight    =   5190
   ClientLeft      =   825
   ClientTop       =   1740
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5190
   ScaleWidth      =   8685
   Begin VB.TextBox txtYStretch 
      Height          =   285
      Left            =   2520
      TabIndex        =   50
      Text            =   "1.0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtXStretch 
      Height          =   285
      Left            =   2520
      TabIndex        =   48
      Text            =   "1.0"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtRotation 
      Height          =   285
      Left            =   840
      TabIndex        =   46
      Text            =   "0.0"
      Top             =   480
      Width           =   615
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "ForeColor"
      Height          =   1575
      Index           =   1
      Left            =   0
      TabIndex        =   30
      Top             =   1200
      Width           =   2295
      Begin VB.OptionButton optForeColor 
         Caption         =   "Cyan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   40
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Yellow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Orange"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   38
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Lt Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   37
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Lt Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "FillColor"
      Height          =   1575
      Index           =   0
      Left            =   2400
      TabIndex        =   24
      Top             =   1200
      Width           =   2295
      Begin VB.OptionButton optFillColor 
         Caption         =   "Lt Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   45
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Lt Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   44
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Orange"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   43
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Yellow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Cyan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   41
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FillStyle"
      Height          =   2295
      Index           =   2
      Left            =   2400
      TabIndex        =   15
      Top             =   2880
      Width           =   2295
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbDiagonalCross"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbFSSolid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbFSTransparent"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbHorizontalLine"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbVerticalLine"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbUpwardDiagonal"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbCross"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbDownwardDiagonal"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1910
      End
   End
   Begin VB.TextBox txtDrawWidth 
      Height          =   285
      Left            =   840
      MaxLength       =   1
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "DrawStyle"
      Height          =   2295
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbInsideSolid"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbTransparent"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbDashDotDot"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbDashDot"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbDot"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbDash"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbSolid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Object"
      Height          =   1095
      Index           =   0
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      Begin VB.OptionButton optObject 
         Caption         =   "Box"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optObject 
         Caption         =   "Line"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optObject 
         Caption         =   "Circle"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.PictureBox picCanvas 
      Height          =   5055
      Left            =   4800
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Y Stretch"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   51
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "X Stretch"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   49
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Rotation"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   47
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "DrawWidth"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   150
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave2D 
         Caption         =   "Save &2D File..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveMetafile 
         Caption         =   "Save &Metafile..."
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
End
Attribute VB_Name = "frmStyles2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TheScene As TwoDObject

Private Enum ObjectTypes
    objLine = 0
    objBox = 1
    objCircle = 2
End Enum

Private ObjectType As ObjectTypes
Private Rubberbanding As Boolean
Private OldMode As Integer
Private OldStyle As Integer
Private FirstX As Single
Private FirstY As Single
Private LastX As Single
Private LastY As Single

Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
Private Type SIZE
    cx As Long
    cy As Long
End Type

' Currently selected drawing properties.
Private CurrentDrawWidth As Integer
Private CurrentDrawStyle As DrawStyleConstants
Private CurrentForeColor As OLE_COLOR
Private CurrentFillColor As OLE_COLOR
Private CurrentFillStyle As FillStyleConstants

' Undo variables.
Private Const MAX_UNDO = 50
Private Snapshots As Collection
Private CurrentSnapshot As Integer
' Save a snapshot for undo.
Private Sub SaveSnapshot()
    ' Remove any previously undone snapshots.
    Do While Snapshots.Count > CurrentSnapshot
        Snapshots.Remove Snapshots.Count
    Loop

    ' Save the current snapshot.
    Snapshots.Add TheScene.Serialization
    If m_Snapshots.Count > MAX_UNDO + 1 Then
        Snapshots.Remove 1
    End If
    CurrentSnapshot = Snapshots.Count

    ' Enable/disable the undo and redo menus.
    SetUndoMenus
End Sub
' Enable or disable the undo and redo menus.
Private Sub SetUndoMenus()
    mnuEditUndo.Enabled = (CurrentSnapshot > 1)
    mnuEditRedo.Enabled = (CurrentSnapshot < Snapshots.Count)
End Sub

' Restore the previous snapshot.
Private Sub Undo()
    If CurrentSnapshot <= 1 Then Exit Sub

    ' Restore the previous snapshot.
    CurrentSnapshot = CurrentSnapshot - 1
    TheScene.Serialization = Snapshots(CurrentSnapshot)

    ' Display the scene.
    picCanvas.Refresh

    ' Enable/disable the undo and redo menus.
    SetUndoMenus
End Sub
' Reapply a previously undone snapshot.
Private Sub Redo()
    If CurrentSnapshot >= Snapshots.Count Then Exit Sub

    ' Restore the previous snapshot.
    CurrentSnapshot = CurrentSnapshot + 1
    TheScene.Serialization = Snapshots(CurrentSnapshot)

    ' Display the scene.
    picCanvas.Refresh

    ' Enable/disable the undo and redo menus.
    SetUndoMenus
End Sub

Private Sub mnuEditRedo_Click()
    Redo
End Sub

Private Sub mnuEditUndo_Click()
    Undo
End Sub


Private Sub mnuFileNew_Click()
    Set TheScene = New TwoDScene

    ' Display the scene.
    picCanvas.Refresh
End Sub

Private Sub mnuFileOpen_Click()
Dim file_name As String
Dim fnum As Integer
Dim the_serialization As String
Dim token_name As String
Dim token_value As String

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Filter = "2D Files (*.2d)|*.2d|" & _
        "All Files (*.*)|*.*"
    dlgFile.Flags = cdlOFNExplorer Or _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Unload dlgFile
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Unload dlgFile
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Read the picture's serialization.
    file_name = dlgFile.FileName
    fnum = FreeFile
    Open file_name For Input As #fnum
    the_serialization = RemoveNonPrintables(Input$(LOF(fnum), fnum))
    Close fnum

    ' Make sure this is a TwoDScene serialization.
    GetNamedToken the_serialization, token_name, token_value
    If token_name <> "TwoDScene" Then
        ' This is not a valid serialization.
        MsgBox "This is not a valid TwoDScene serialization."
    Else
        Caption = "Show2D [" & dlgFile.FileTitle & "]"
        dlgFile.InitDir = Left$(file_name, Len(file_name) _
            - Len(dlgFile.FileTitle) - 1)

        ' Initialize the new scene.
        Set TheScene = New TwoDScene
        TheScene.Serialization = token_value
    End If

    ' Display the scene.
    picCanvas.Refresh
End Sub

' Save the object serialization.
Private Sub mnuFileSave2D_Click()
Dim file_name As String
Dim fnum As Integer

    If TheScene Is Nothing Then
        MsgBox "No scene is loaded."
        Exit Sub
    End If

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Filter = _
        "2D Files (*.2d)|*.2d|" & _
        "All Files (*.*)|*.*"
    dlgFile.Flags = _
        cdlOFNOverwritePrompt Or _
        cdlOFNPathMustExist Or _
        cdlOFNHideReadOnly
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        ' The user canceled.
        Unload dlgFile
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' Unknown error.
        Unload dlgFile
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo Save2DFileError

    ' Get the file name.
    file_name = dlgFile.FileName
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)
    Caption = "Show2D [" & dlgFile.FileTitle & "]"

    ' Open the file.
    fnum = FreeFile
    Open file_name For Output As fnum

    ' Write the serialization into the file.
    Print #fnum, TheScene.Serialization

    ' Close the file.
    Close fnum
    Exit Sub

Save2DFileError:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file." & vbCrLf & _
        Err.Description, vbExclamation
    Exit Sub
End Sub


' Draw an ellipse bounded by a rectangle.
Private Sub DrawEllipse(ByVal obj As Object, ByVal xmin As Single, ByVal ymin As Single, ByVal xmax As Single, ByVal ymax As Single)
Dim cx As Single
Dim cy As Single
Dim wid As Single
Dim hgt As Single
Dim aspect As Single
Dim Radius As Single

    ' Find the center.
    cx = (xmin + xmax) / 2
    cy = (ymin + ymax) / 2

    ' Get the ellipse's size.
    wid = Abs(xmax - xmin)
    hgt = Abs(ymax - ymin)

    ' Do nothing if the width or height is zero.
    If (wid = 0) Or (hgt = 0) Then Exit Sub

    aspect = hgt / wid

    ' See which dimension is larger.
    If wid > hgt Then
        ' The major axis is horizontal.
        ' Get the radius in custom coordinates.
        Radius = wid / 2
    Else
        ' The major axis is vertical.
        ' Get the radius in custom coordinates.
        Radius = hgt / 2
    End If

    ' Draw the circle.
    obj.Circle (cx, cy), Radius, , , , aspect
End Sub


' Draw the appropriate object.
Private Sub DrawObject(ByVal xmin As Single, ByVal ymin As Single, ByVal xmax As Single, ByVal ymax As Single)
    Select Case ObjectType
        Case objLine
            picCanvas.Line (xmin, ymin)-(xmax, ymax)
        Case objBox
            picCanvas.Line (xmin, ymin)-(xmax, ymax), , B
        Case objCircle
            DrawEllipse picCanvas, xmin, ymin, xmax, ymax
    End Select
End Sub
' Create the appropriate object and redraw.
Private Sub Create2DObject(ByVal xmin As Single, ByVal ymin As Single, ByVal xmax As Single, ByVal ymax As Single)
Const PI = 3.14159265

Dim obj As TwoDObject
Dim obj_line As TwoDLine
Dim obj_rectangle As TwoDRectangle
Dim obj_ellipse As TwoDEllipse
Dim obj_scene As TwoDScene
Dim M(1 To 3, 1 To 3) As Single

    ' Create the new object.
    Select Case ObjectType
        Case objLine
            Set obj = New TwoDLine
            Set obj_line = obj
            obj_line.X1 = xmin
            obj_line.X2 = xmax
            obj_line.Y1 = ymin
            obj_line.Y2 = ymax
        Case objBox
            Set obj = New TwoDRectangle
            Set obj_rectangle = obj
            obj_rectangle.X1 = xmin
            obj_rectangle.X2 = xmax
            obj_rectangle.Y1 = ymin
            obj_rectangle.Y2 = ymax
        Case objCircle
            Set obj = New TwoDEllipse
            Set obj_ellipse = obj
            obj_ellipse.X1 = xmin
            obj_ellipse.X2 = xmax
            obj_ellipse.Y1 = ymin
            obj_ellipse.Y2 = ymax
            obj_ellipse.X1 = xmin
    End Select

    ' Set the new object's drawing properties.
    obj.DrawWidth = CurrentDrawWidth
    obj.DrawStyle = CurrentDrawStyle
    obj.ForeColor = CurrentForeColor
    obj.FillColor = CurrentFillColor
    obj.FillStyle = CurrentFillStyle

    ' Add the object to the scene.
    Set obj_scene = TheScene
    obj_scene.SceneObjects.Add obj

    ' Save the current scene.
    SaveSnapshot

    ' Display the scene.
    picCanvas.Refresh
End Sub
' Make the canvas as big as possible.
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub

Private Sub mnuFileSaveMetafile_Click()
Dim file_name As String
Dim mf_dc As Long
Dim hmf As Long
Dim old_size As SIZE

    If TheScene Is Nothing Then
        MsgBox "No scene is loaded."
        Exit Sub
    End If

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Filter = _
        "Metafiles (*.wmf)|*.wmf|" & _
        "All Files (*.*)|*.*"
    dlgFile.Flags = _
        cdlOFNOverwritePrompt Or _
        cdlOFNPathMustExist Or _
        cdlOFNHideReadOnly
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        ' The user canceled.
        Unload dlgFile
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' Unknown error.
        Unload dlgFile
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo SaveMetafileError

    ' Get the file name.
    file_name = dlgFile.FileName
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)
    Caption = "Show2D [" & dlgFile.FileTitle & "]"

    ' Create the metafile.
    mf_dc = CreateMetaFile(ByVal file_name)
    If mf_dc = 0 Then
        MsgBox "Error creating the metafile.", vbExclamation
        Exit Sub
    End If

    ' Set the metafile's size to something reasonable.
    SetWindowExtEx mf_dc, picCanvas.ScaleWidth, _
        picCanvas.ScaleHeight, old_size

    ' Draw in the metafile.
    TheScene.DrawInMetafile mf_dc

    ' Close the metafile.
    hmf = CloseMetaFile(mf_dc)
    If hmf = 0 Then
        MsgBox "Error closing the metafile.", vbExclamation
    End If

    ' Delete the metafile to free resources.
    If DeleteMetaFile(hmf) = 0 Then
        MsgBox "Error deleting the metafile.", vbExclamation
    End If
    Exit Sub

SaveMetafileError:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file." & vbCrLf & _
        Err.Description, vbExclamation
    Exit Sub
End Sub

' Set the DrawStyle.
Private Sub optDrawStyle_Click(Index As Integer)
    CurrentDrawStyle = Index
End Sub

' Set the FillColor.
Private Sub optFillColor_Click(Index As Integer)
    CurrentFillColor = optFillColor(Index).ForeColor
End Sub

' Set the FillStyle.
Private Sub optFillStyle_Click(Index As Integer)
    CurrentFillStyle = Index
End Sub


' Start a rubberbanding of some sort.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Let MouseMove know we are rubberbanding.
    Rubberbanding = True

    ' Save values so we can restore them later.
    OldMode = picCanvas.DrawMode
    OldStyle = picCanvas.DrawStyle
    picCanvas.DrawMode = vbInvert
    If ObjectType = objLine Then
        picCanvas.DrawStyle = vbSolid
    Else
        picCanvas.DrawStyle = vbDot
    End If

    ' Save the starting coordinates.
    FirstX = X
    FirstY = Y

    ' Save the ending coordinates.
    LastX = X
    LastY = Y

    ' Draw the appropriate rubberband object.
    DrawObject FirstX, FirstY, LastX, LastY
End Sub
' Continue rubberbanding.
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' Erase the previous rubberband object.
    DrawObject FirstX, FirstY, LastX, LastY

    ' Save the new ending coordinates.
    LastX = X
    LastY = Y

    ' Draw the new rubberband object.
    DrawObject FirstX, FirstY, LastX, LastY
End Sub
' Finish rubberbanding and draw the object.
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' We are no longer rubberbanding.
    Rubberbanding = False

    ' Erase the previous rubberband object.
    DrawObject FirstX, FirstY, LastX, LastY

    ' Restore the original DrawMode and DrawStyle.
    picCanvas.DrawMode = OldMode
    picCanvas.DrawStyle = OldStyle

    ' Create the final object.
    Create2DObject FirstX, FirstY, LastX, LastY
End Sub
' Select the default options.
Private Sub Form_Load()
    optForeColor(0).Value = True
    optFillColor(0).Value = True
    optDrawStyle(vbSolid).Value = True
    optFillStyle(vbFSTransparent).Value = True
    txtDrawWidth.Text = Format$(picCanvas.DrawWidth)
    optObject(ObjectType).Value = True

    ' Initialize the common dialog.
    dlgFile.InitDir = App.Path
    dlgFile.CancelError = True

    ' Create an empty scene.
    Set TheScene = New TwoDScene

    ' Save the initial, empty snapshot.
    Set Snapshots = New Collection
    SaveSnapshot
End Sub
' Record the kind of object to draw next.
Private Sub optObject_Click(Index As Integer)
    ObjectType = Index
End Sub


' Set the ForeColor.
Private Sub optForeColor_Click(Index As Integer)
    CurrentForeColor = optForeColor(Index).ForeColor
End Sub

Private Sub picCanvas_Paint()
    picCanvas.Cls
    If Not TheScene Is Nothing Then TheScene.Draw picCanvas
End Sub

' Change set DrawWidth.
Private Sub txtDrawWidth_Change()
Dim wid As Integer

    If Not IsNumeric(txtDrawWidth.Text) Then Exit Sub

    wid = CInt(txtDrawWidth.Text)
    If wid < 1 Then Exit Sub

    CurrentDrawWidth = wid
End Sub

' Only allow 1 through 9.
Private Sub txtDrawWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc(" ") Or _
       KeyAscii > Asc("~") Then Exit Sub
    If KeyAscii >= Asc("1") And _
       KeyAscii <= Asc("9") Then Exit Sub
    Beep
    KeyAscii = 0
End Sub


