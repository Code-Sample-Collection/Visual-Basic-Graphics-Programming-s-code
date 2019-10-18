VERSION 5.00
Begin VB.Form frmPaintPic 
   Caption         =   "PaintPic"
   ClientHeight    =   4455
   ClientLeft      =   2085
   ClientTop       =   1200
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7320
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   0
      Picture         =   "PaintPic.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   12
      Top             =   1320
      Width           =   960
   End
   Begin VB.CommandButton cmdClearResult 
      Caption         =   "Clear Result"
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   3
      Left            =   0
      Picture         =   "PaintPic.frx":0882
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   3480
      Width           =   960
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy ==>"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   2
      Left            =   0
      Picture         =   "PaintPic.frx":1CC4
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   2400
      Width           =   960
   End
   Begin VB.ListBox lstOpcode 
      Height          =   4155
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2520
      Index           =   2
      Left            =   4800
      ScaleHeight     =   164
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   3
      Top             =   1920
      Width           =   2520
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   510
      Index           =   0
      Left            =   4800
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   240
      Width           =   510
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Index           =   1
      Left            =   4800
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   840
      Width           =   1020
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   0
      Picture         =   "PaintPic.frx":3106
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Result"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Opcode"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Source"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmPaintPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelectedSource As Integer
' ***********************************************
' Add an opcode's name and value to the list of
' choices.
' ***********************************************
Sub AddOpcode(name As String, value As Long)
    lstOpcode.AddItem name
    lstOpcode.ItemData(lstOpcode.NewIndex) = value
End Sub




' Clear the result.
Private Sub cmdClearResult_Click()
Dim dest As Integer

    ' Clear all destination PictureBoxes.
    For dest = 0 To picDestination.UBound
        picDestination(dest).Cls
    Next dest
End Sub

' Load the opcode choices.
Private Sub Form_Load()
Dim copy_index As Integer

    AddOpcode "vbBlackness", vbBlackness
    AddOpcode "vbDstInvert", vbDstInvert
    AddOpcode "vbMergeCopy", vbMergeCopy
    AddOpcode "vbMergePaint", vbMergePaint
    AddOpcode "vbNotSrcCopy", vbNotSrcCopy
    AddOpcode "vbSrcErase", vbSrcErase
    AddOpcode "vbPatCopy", vbPatCopy
    AddOpcode "vbPatInvert", vbPatInvert
    AddOpcode "vbPatPaint", vbPatPaint
    AddOpcode "vbSrcAnd", vbSrcAnd
    AddOpcode "vbSrcCopy", vbSrcCopy
    copy_index = lstOpcode.NewIndex
    AddOpcode "vbSrcErase", vbSrcErase
    AddOpcode "vbSrcInvert", vbSrcInvert
    AddOpcode "vbSrcPaint", vbSrcPaint
    AddOpcode "vbWhiteness", vbWhiteness

    ' Start with vbSrcCopy.
    lstOpcode.ListIndex = copy_index

    ' Select the first source image.
    picSource_Click (0)
End Sub
' Copy the image.
Private Sub cmdCopy_Click()
Dim source_wid As Single
Dim source_hgt As Single
Dim dest_wid As Single
Dim dest_hgt As Single
Dim opcode As Long
Dim dest_num As Integer

    ' Get the source image's dimenstions.
    source_wid = picSource(SelectedSource).ScaleWidth
    source_hgt = picSource(SelectedSource).ScaleHeight

    ' Get the selected opcode.
    opcode = lstOpcode.ItemData(lstOpcode.ListIndex)

    ' Copy the image into the destination images.
    For dest_num = 0 To picDestination.UBound
        With picDestination(dest_num)
            ' Get the destination dimensions.
            dest_wid = picDestination(dest_num).ScaleWidth
            dest_hgt = picDestination(dest_num).ScaleHeight

            ' Copy the image.
            picDestination(dest_num).PaintPicture _
                picSource(SelectedSource).Picture, _
                0, 0, dest_wid, dest_hgt, _
                0, 0, source_wid, source_hgt, _
                opcode
        End With
    Next dest_num
End Sub
' Save the index of the selected source.
Private Sub picSource_Click(Index As Integer)
    picSource(SelectedSource).BorderStyle = vbBSNone
    SelectedSource = Index
    picSource(SelectedSource).BorderStyle = vbFixedSingle
End Sub


