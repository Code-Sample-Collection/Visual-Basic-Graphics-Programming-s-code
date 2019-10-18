VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDistort 
   Caption         =   "Distort"
   ClientHeight    =   4920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCanvas 
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   480
      Width           =   4335
   End
   Begin VB.OptionButton optTransformation 
      Caption         =   "Fish Eye"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton optTransformation 
      Caption         =   "Twist"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton optTransformation 
      Caption         =   "Wave"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton optTransformation 
      Caption         =   "None"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveToMetafile 
         Caption         =   "&Save to Metafile..."
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmDistort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const POINTS_PER_ROW = 20
Private PointX(1 To POINTS_PER_ROW, 1 To POINTS_PER_ROW) As Single
Private PointY(1 To POINTS_PER_ROW, 1 To POINTS_PER_ROW) As Single

' Matefile API functions.
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
Private Type SIZE
    Cx As Long
    Cy As Long
End Type
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' Draw the transformed points.
Private Sub DrawPoints(ByVal pic As PictureBox)
Dim i As Integer
Dim j As Integer

    ' Draw the horizontal lines.
    For i = 1 To POINTS_PER_ROW
        pic.CurrentX = PointX(i, 1)
        pic.CurrentY = PointY(i, 1)
        For j = 2 To POINTS_PER_ROW
            pic.Line -(PointX(i, j), PointY(i, j))
        Next j
    Next i

    ' Draw the vertical lines.
    For j = 1 To POINTS_PER_ROW
        pic.CurrentX = PointX(1, j)
        pic.CurrentY = PointY(1, j)
        For i = 2 To POINTS_PER_ROW
            pic.Line -(PointX(i, j), PointY(i, j))
        Next i
    Next j
End Sub

' Draw the transformed points into a metafile.
Private Sub DrawPointsIntoMetafile(ByVal pic As PictureBox, ByVal mf_dc As Long)
Dim i As Integer
Dim j As Integer

    ' Draw the horizontal lines.
    For i = 1 To POINTS_PER_ROW
        MoveTo mf_dc, PointX(i, 1), PointY(i, 1), ByVal 0&
        For j = 2 To POINTS_PER_ROW
            LineTo mf_dc, PointX(i, j), PointY(i, j)
        Next j
    Next i

    ' Draw the vertical lines.
    For j = 1 To POINTS_PER_ROW
        MoveTo mf_dc, PointX(1, j), PointY(1, j), ByVal 0&
        For i = 2 To POINTS_PER_ROW
            LineTo mf_dc, PointX(i, j), PointY(i, j)
        Next i
    Next j
End Sub


' Create the data points.
Private Sub MakeData(ByVal pic As PictureBox)
Const SQUARES_MARGIN = 2
Dim dx As Single
Dim dy As Single
Dim X As Single
Dim Y As Single
Dim i As Integer
Dim j As Integer

    dx = pic.ScaleWidth / (POINTS_PER_ROW + 2 * SQUARES_MARGIN - 1)
    dy = pic.ScaleHeight / (POINTS_PER_ROW + 2 * SQUARES_MARGIN - 1)
    Y = pic.ScaleTop + dy * SQUARES_MARGIN
    For i = 1 To POINTS_PER_ROW
        X = pic.ScaleLeft + dx * SQUARES_MARGIN
        For j = 1 To POINTS_PER_ROW
            PointX(i, j) = X
            PointY(i, j) = Y
            X = X + dx
        Next j
        Y = Y + dy
    Next i
End Sub

Private Sub Form_Load()
    picCanvas.ScaleMode = vbPixels

    dlgFile.InitDir = App.Path
    dlgFile.Filter = "Metafiles (*.wmf)|*.wmf|" & _
        "All Files (*.*)|*.*"
    dlgFile.CancelError = True
    dlgFile.Flags = cdlOFNExplorer Or _
        cdlOFNLongNames Or _
        cdlOFNHideReadOnly Or _
        cdlOFNOverwritePrompt

    optTransformation(0).Value = True
End Sub

Private Sub mnuFileSaveToMetafile_Click()
Dim file_name As String
Dim mf_dc As Long
Dim hmf As Long
Dim old_size As SIZE

    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
                Err.Description
    End If
    On Error GoTo 0

    ' Get the file name.
    file_name = dlgFile.FileName
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    ' Create the metafile.
    mf_dc = CreateMetaFile(ByVal file_name)
    If mf_dc = 0 Then
        MsgBox "Error creating the metafile.", vbExclamation
        Exit Sub
    End If

    ' Set the metafile's size to something reasonable.
    SetWindowExtEx mf_dc, picCanvas.ScaleWidth, picCanvas.ScaleHeight, old_size

    ' Draw into the metafile.
    DrawPointsIntoMetafile picCanvas, mf_dc

    ' Close the metafile.
    hmf = CloseMetaFile(mf_dc)
    If hmf = 0 Then
        MsgBox "Error closing the metafile.", vbExclamation
    End If

    ' Delete the metafile to free resources.
    If DeleteMetaFile(hmf) = 0 Then
        MsgBox "Error deleting the metafile.", vbExclamation
    End If
End Sub

Private Sub optTransformation_Click(Index As Integer)
Dim obj As Transformation
Dim twist As TransTwist
Dim wave As TransWave
Dim fish As TransFish
Dim i As Integer
Dim j As Integer

    ' Make the data.
    MakeData picCanvas

    ' Get the transformation object.
    If optTransformation(1).Value Then
        Set wave = New TransWave
        wave.Amplitude = 8
        wave.Period = 20 * 8
        Set obj = wave
    ElseIf optTransformation(2).Value Then
        Set twist = New TransTwist
        twist.Cx = picCanvas.ScaleWidth / 2
        twist.Cy = picCanvas.ScaleHeight / 2
        twist.TwistSpeed = 20
        Set obj = twist
    ElseIf optTransformation(3).Value Then
        Set fish = New TransFish
        fish.Cx = picCanvas.ScaleWidth / 2
        fish.Cy = picCanvas.ScaleHeight / 2
        fish.Radius = picCanvas.ScaleWidth
        Set obj = fish
    Else
        Set obj = New TransIdentity
    End If

    ' Transform the points.
    For i = 1 To POINTS_PER_ROW
        For j = 1 To POINTS_PER_ROW
            obj.Transform PointX(i, j), PointY(i, j)
        Next j
    Next i

    ' Redraw.
    picCanvas.Refresh
End Sub
Private Sub picCanvas_Paint()
    DrawPoints picCanvas
End Sub
