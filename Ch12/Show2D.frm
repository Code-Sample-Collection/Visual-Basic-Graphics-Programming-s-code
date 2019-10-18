VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmShow2D 
   Caption         =   "Show2D"
   ClientHeight    =   4365
   ClientLeft      =   2415
   ClientTop       =   1650
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   5355
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picCanvas 
      Height          =   2415
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
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
         Shortcut        =   ^M
      End
   End
End
Attribute VB_Name = "frmShow2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The scene that contains all other objects.
Private TheScene As TwoDObject

Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
Private Type SIZE
    cx As Long
    cy As Long
End Type

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
Private Sub picCanvas_Paint()
    If Not TheScene Is Nothing Then TheScene.Draw picCanvas
End Sub
Private Sub Form_Load()
    dlgFile.InitDir = App.Path
    dlgFile.Filter = "TwoD Files (*.2d)|*.2d|" & _
        "All Files (*.*)|*.*"
    dlgFile.CancelError = True
End Sub

Private Sub Form_Resize()
    picCanvas.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


Private Sub mnuFileOpen_Click()
Dim file_name As String
Dim fnum As Integer
Dim the_serialization As String
Dim token_name As String
Dim token_value As String

    ' Allow the user to pick a file.
    On Error Resume Next
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
    picCanvas.Cls
    TheScene.Draw picCanvas
    picCanvas.Refresh
End Sub
