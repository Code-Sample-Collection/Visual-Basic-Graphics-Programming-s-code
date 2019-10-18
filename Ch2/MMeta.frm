VERSION 5.00
Begin VB.Form frmMMeta 
   AutoRedraw      =   -1  'True
   Caption         =   "MMeta"
   ClientHeight    =   3495
   ClientLeft      =   1950
   ClientTop       =   825
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   5295
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox picCopy 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Index           =   2
      Left            =   3600
      ScaleHeight     =   1635
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox picCopy 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Index           =   1
      Left            =   1800
      ScaleHeight     =   1635
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox picCopy 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Index           =   0
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   1800
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmMMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Drawing As Boolean
Private PointX() As Single
Private PointY() As Single
Private NumPoints As Integer

Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As Any) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

' Create a memory metafile and play it back into
' the destination picture boxes.
Private Sub cmdCopy_Click()
Dim i As Integer
Dim mDC As Long
Dim hMF As Long
Dim status As Long
Dim x As Single
Dim y As Single

    ' Create the memory metafile.
    mDC = CreateMetaFile(ByVal 0&)
    If mDC = 0 Then
        MsgBox "Error creating the metafile.", vbExclamation
        Exit Sub
    End If

    ' Draw in the metafile.
    For i = 1 To NumPoints
        x = PointX(i)
        y = PointY(i)
        If x < 0 Then
            MoveToEx mDC, -x, y, ByVal 0&
        Else
            LineTo mDC, x, y
        End If
    Next i

    ' Close the metafile.
    hMF = CloseMetaFile(mDC)
    If hMF = 0 Then
        MsgBox "Error closing the metafile.", vbExclamation
    Else
        ' Play the metafile.
        For i = 0 To 2
            picCopy(i).Cls
            If PlayMetaFile(picCopy(i).hdc, hMF) = 0 Then
                MsgBox "Error playing the metafile.", vbExclamation
                Exit For
            End If
            picCopy(i).Refresh
        Next i
    End If

    ' Delete the metafile.
    If DeleteMetaFile(hMF) = 0 Then
        MsgBox "Error deleting the metafile.", vbExclamation
    End If
End Sub
' Start drawing.
Private Sub picSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Drawing = True
    AddPoint -x, y
End Sub
' Add a point to the list of points.
Private Sub AddPoint(ByVal x As Single, ByVal y As Single)
    ' Add the new point.
    NumPoints = NumPoints + 1
    ReDim Preserve PointX(1 To NumPoints)
    ReDim Preserve PointY(1 To NumPoints)
    PointX(NumPoints) = x
    PointY(NumPoints) = y

    ' This represents the start of a new segment.
    If x < 0 Then
        picSource.CurrentX = -x
        picSource.CurrentY = y
    Else
        picSource.Line -(x, y)
    End If
End Sub

' Continue drawing.
Private Sub picSource_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Do nothing if we are not drawing.
    If Not Drawing Then Exit Sub

    ' Add the point.
    AddPoint x, y
End Sub

' Stop drawing.
Private Sub picSource_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Drawing = False
End Sub
' Clear the form.
Private Sub cmdClear_Click()
    picSource.Cls
    NumPoints = 0
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub
