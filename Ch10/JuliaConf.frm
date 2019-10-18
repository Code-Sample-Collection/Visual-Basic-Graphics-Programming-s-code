VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMaxJuliaIterations 
      Height          =   285
      Left            =   2760
      TabIndex        =   71
      Text            =   "100"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   480
      TabIndex        =   70
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdRow 
      Height          =   255
      Index           =   5
      Left            =   120
      Picture         =   "JuliaConf.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton cmdRow 
      Height          =   255
      Index           =   4
      Left            =   120
      Picture         =   "JuliaConf.frx":00DA
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton cmdRow 
      Height          =   255
      Index           =   3
      Left            =   120
      Picture         =   "JuliaConf.frx":01B4
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton cmdRow 
      Height          =   255
      Index           =   2
      Left            =   120
      Picture         =   "JuliaConf.frx":028E
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton cmdRow 
      Height          =   255
      Index           =   1
      Left            =   120
      Picture         =   "JuliaConf.frx":0368
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdRow 
      Height          =   255
      Index           =   0
      Left            =   120
      Picture         =   "JuliaConf.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdColumn 
      Height          =   255
      Index           =   7
      Left            =   3000
      Picture         =   "JuliaConf.frx":051C
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmdColumn 
      Height          =   255
      Index           =   6
      Left            =   2640
      Picture         =   "JuliaConf.frx":05F6
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmdColumn 
      Height          =   255
      Index           =   5
      Left            =   2280
      Picture         =   "JuliaConf.frx":06D0
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmdColumn 
      Height          =   255
      Index           =   4
      Left            =   1920
      Picture         =   "JuliaConf.frx":07AA
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmdColumn 
      Height          =   255
      Index           =   3
      Left            =   1560
      Picture         =   "JuliaConf.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmdColumn 
      Height          =   255
      Index           =   2
      Left            =   1200
      Picture         =   "JuliaConf.frx":095E
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmdColumn 
      Height          =   255
      Index           =   1
      Left            =   840
      Picture         =   "JuliaConf.frx":0A38
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00400040&
      Height          =   255
      Index           =   47
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   56
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00400000&
      Height          =   255
      Index           =   46
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   55
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00404000&
      Height          =   255
      Index           =   45
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   54
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00004000&
      Height          =   255
      Index           =   44
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   53
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00004040&
      Height          =   255
      Index           =   43
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00404080&
      Height          =   255
      Index           =   42
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   51
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00000040&
      Height          =   255
      Index           =   41
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   50
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   40
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   49
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton cmdColumn 
      Height          =   255
      Index           =   0
      Left            =   480
      Picture         =   "JuliaConf.frx":0B12
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00800080&
      Height          =   255
      Index           =   39
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   47
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00800000&
      Height          =   255
      Index           =   38
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   46
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00808000&
      Height          =   255
      Index           =   37
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   45
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00008000&
      Height          =   255
      Index           =   36
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   44
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00008080&
      Height          =   255
      Index           =   35
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   43
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00004080&
      Height          =   255
      Index           =   34
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   42
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00000080&
      Height          =   255
      Index           =   33
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   41
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   32
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   40
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C000C0&
      Height          =   255
      Index           =   31
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   39
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C00000&
      Height          =   255
      Index           =   30
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   38
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0C000&
      Height          =   255
      Index           =   29
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   37
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   28
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   36
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000C0C0&
      Height          =   255
      Index           =   27
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   35
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000040C0&
      Height          =   255
      Index           =   26
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   34
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000000C0&
      Height          =   255
      Index           =   25
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   33
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   24
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   32
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Index           =   23
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   31
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   22
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   30
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   21
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   29
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   20
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   28
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   19
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000080FF&
      Height          =   255
      Index           =   18
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   26
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   17
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   25
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   16
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   24
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Index           =   15
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   14
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   22
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Index           =   13
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0080FF80&
      Height          =   255
      Index           =   12
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Index           =   11
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Index           =   10
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H008080FF&
      Height          =   255
      Index           =   9
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   8
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   7
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   6
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   5
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   2
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   1
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox ColorCheck 
      BackColor       =   &H00F0CAA6&
      Caption         =   "Sky Blue"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox ColorCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "White"
      Height          =   255
      Index           =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "None"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtMaxMandelbrotIterations 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "100"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Maximum Julia Iterations"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   72
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Maximum Mandelbrot Iterations"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The form this one is configuring.
Private FractalForm As Form
Private Sub cmdAll_Click()
Dim i As Integer

    For i = picColor.LBound To picColor.UBound
        picColor(i).BorderStyle = vbFixedSingle
    Next i
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


' Select the colors in this column.
Private Sub cmdColumn_Click(Index As Integer)
Dim r As Integer
Dim c As Integer

    For r = 0 To 5
        For c = 0 To 7
            If c = Index Then
                picColor(r * 8 + c).BorderStyle = vbFixedSingle
            Else
                picColor(r * 8 + c).BorderStyle = vbBSNone
            End If
        Next c
    Next r
End Sub

' Select some default colors.
Private Sub cmdDefault_Click()
    cmdRow_Click 2
End Sub

Private Sub cmdNone_Click()
Dim i As Integer

    For i = picColor.LBound To picColor.UBound
        picColor(i).BorderStyle = vbBSNone
    Next i
End Sub

' Save the changes.
Private Sub cmdOk_Click()
Dim new_iter As Integer
Dim i As Integer

    ' Get the number of Mandelbrot iterations.
    new_iter = FractalForm.MaxMandelbrotIterations
    On Error Resume Next
    new_iter = CInt(txtMaxMandelbrotIterations.Text)
    On Error GoTo 0
    If new_iter < 1 Then new_iter = 1
    FractalForm.MaxMandelbrotIterations = new_iter

    ' Get the number of Julia iterations.
    new_iter = FractalForm.MaxJuliaIterations
    On Error Resume Next
    new_iter = CInt(txtMaxJuliaIterations.Text)
    On Error GoTo 0
    If new_iter < 1 Then new_iter = 1
    FractalForm.MaxJuliaIterations = new_iter

    ' Save the selected colors.
    FractalForm.ResetColors
    For i = picColor.LBound To picColor.UBound
        If picColor(i).BorderStyle = vbFixedSingle _
            Then FractalForm.AddColor picColor(i).BackColor
    Next i

    Unload Me
End Sub


Private Sub cmdRow_Click(Index As Integer)
Dim r As Integer
Dim c As Integer

    For r = 0 To 5
        For c = 0 To 7
            If r = Index Then
                picColor(r * 8 + c).BorderStyle = vbFixedSingle
            Else
                picColor(r * 8 + c).BorderStyle = vbBSNone
            End If
        Next c
    Next r
End Sub

' Initialize the options.
Public Sub Initialize(ByVal frm As Form)
Dim i As Integer
Dim j As Integer

    Set FractalForm = frm
    txtMaxMandelbrotIterations.Text = Format$(FractalForm.MaxMandelbrotIterations)
    txtMaxJuliaIterations.Text = Format$(FractalForm.MaxJuliaIterations)

    For i = picColor.LBound To picColor.UBound
        ' See if this color is selected.
        picColor(i).BorderStyle = vbBSNone
        For j = 1 To FractalForm.numcolors
            If FractalForm.color(j) = picColor(i).BackColor Then
                ' The color is selected.
                picColor(i).BorderStyle = vbFixedSingle
                Exit For
            End If
        Next j
    Next i
End Sub
Private Sub picColor_Click(Index As Integer)
    If picColor(Index).BorderStyle = vbBSNone Then
        picColor(Index).BorderStyle = vbFixedSingle
    Else
        picColor(Index).BorderStyle = vbBSNone
    End If
End Sub


