VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMeta 
   AutoRedraw      =   -1  'True
   Caption         =   "Meta"
   ClientHeight    =   3405
   ClientLeft      =   1950
   ClientTop       =   1110
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   Begin MSComDlg.CommonDialog dlgMetafile 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save As..."
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClear 
         Caption         =   "&Clear"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Drawing As Boolean
Private MetafileLoaded As Boolean
Private PointX() As Single
Private PointY() As Single
Private NumPoints As Integer

Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
Private Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

' Start in the current directory.
Private Sub Form_Load()
    dlgMetafile.InitDir = App.Path
End Sub

' Start drawing.
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Drawing = True
    AddPoint -x, y
End Sub
' Add a point to the list of points.
Private Sub AddPoint(ByVal x As Single, ByVal y As Single)
    ' Start over if a metafile is currently displayed.
    If MetafileLoaded Then
        Cls
        MetafileLoaded = False
        NumPoints = 0
    End If

    ' Add the new point.
    NumPoints = NumPoints + 1
    ReDim Preserve PointX(1 To NumPoints)
    ReDim Preserve PointY(1 To NumPoints)
    PointX(NumPoints) = x
    PointY(NumPoints) = y

    ' This represents the start of a new segment.
    If x < 0 Then
        CurrentX = -x
        CurrentY = y
    Else
        Line -(x, y)
    End If

    mnuFileSaveAs.Enabled = True
End Sub

' Continue drawing.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Do nothing if we are not drawing.
    If Not Drawing Then Exit Sub

    ' Add the point.
    AddPoint x, y
End Sub

' Stop drawing.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Drawing = False
End Sub
' Clear the form.
Private Sub mnuFileClear_Click()
    Cls
    NumPoints = 0
    mnuFileSaveAs.Enabled = False
End Sub


' Load a metafile.
Private Sub mnuFileOpen_Click()
Dim fname As String
Dim hMF As Long

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgMetafile.FileName = "*.wmf"
    dlgMetafile.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    dlgMetafile.ShowOpen
    If Err.Number = cdlCancel Then
        ' The user clicked Cancel.
        Unload dlgMetafile
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' Unknown error.
        Unload dlgMetafile
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo LoadErr

    ' Get the file name.
    fname = dlgMetafile.FileName
    dlgMetafile.InitDir = Left$(fname, Len(fname) _
        - Len(dlgMetafile.FileTitle) - 1)

    ' Load the metafile.
    hMF = GetMetaFile(fname)
    If hMF = 0 Then
        MsgBox "Unable to load the metafile.", vbExclamation
        Exit Sub
    End If

    ' Play the metafile.
    Cls
    If PlayMetaFile(hdc, hMF) = 0 Then
        MsgBox "Error playing the metafile.", vbExclamation
    End If
    
    ' Delete the metafile to free resources.
    If DeleteMetaFile(hMF) = 0 Then
        MsgBox "Error deleting metafile " & _
            fname & ".", vbExclamation
    End If

    Refresh
    MetafileLoaded = True
    mnuFileSaveAs.Enabled = False
    Exit Sub

LoadErr:
    MsgBox "Error " & Format$(Err.Number) & _
        " loading the metafile." & vbCrLf & _
        Err.Description, vbExclamation
    Exit Sub
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Save the drawing in a metafile.
Private Sub mnuFileSaveAs_Click()
Dim fname As String
Dim i As Integer
Dim mDC As Long
Dim hMF As Long
Dim x As Single
Dim y As Single
Dim old_size As SIZE

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgMetafile.FileName = "*.wmf"
    dlgMetafile.Flags = cdlOFNOverwritePrompt + _
        cdlOFNPathMustExist + cdlOFNHideReadOnly
    dlgMetafile.ShowSave
    If Err.Number = cdlCancel Then
        ' The user canceled.
        Unload dlgMetafile
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' Unknown error.
        Unload dlgMetafile
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo SaveErr

    ' Get the file name.
    fname = dlgMetafile.FileName
    dlgMetafile.InitDir = Left$(fname, Len(fname) _
        - Len(dlgMetafile.FileTitle) - 1)

    ' Create the metafile.
    mDC = CreateMetaFile(ByVal fname)
    If mDC = 0 Then
        MsgBox "Error creating the metafile.", vbExclamation
        Exit Sub
    End If

    ' Set the metafile's size to something reasonable.
    SetWindowExtEx mDC, ScaleWidth, _
        ScaleHeight, old_size

    ' Draw in the metafile.
    For i = 1 To NumPoints
        x = PointX(i)
        y = PointY(i)
        If x < 0 Then
            MoveToEx mDC, -x, y, vbNullString
        Else
            LineTo mDC, x, y
        End If
    Next i

    ' Close the metafile.
    hMF = CloseMetaFile(mDC)
    If hMF = 0 Then
        MsgBox "Error closing the metafile.", vbExclamation
    End If

    ' Delete the metafile to free resources.
    If DeleteMetaFile(hMF) = 0 Then
        MsgBox "Error deleting the metafile.", vbExclamation
    End If
    Exit Sub

SaveErr:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file." & vbCrLf & _
        Err.Description, vbExclamation
    Exit Sub
End Sub
