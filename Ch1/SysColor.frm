VERSION 5.00
Begin VB.Form frmSysColor 
   Caption         =   "SysColor"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmSysColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RowHgt As Single
Private MaxIndex As Integer

' Display one color.
Private Sub ShowColor(ByVal color_name As String, ByVal color_value As Long)
    ' Create the labels for this color.
    MaxIndex = MaxIndex + 1
    Load lblName(MaxIndex)
    Load lblColor(MaxIndex)

    If MaxIndex = 14 Then
        lblName(MaxIndex).Top = lblName(1).Top
        lblName(MaxIndex).Left = lblColor(1).Left + lblColor(1).Width + 240
        lblColor(MaxIndex).Top = lblColor(1).Top
        lblColor(MaxIndex).Left = lblName(MaxIndex).Left + lblColor(1).Left - lblName(1).Left
    Else
        lblName(MaxIndex).Top = lblName(MaxIndex - 1).Top + RowHgt
        lblName(MaxIndex).Left = lblName(MaxIndex - 1).Left
        lblColor(MaxIndex).Top = lblColor(MaxIndex - 1).Top + RowHgt
        lblColor(MaxIndex).Left = lblColor(MaxIndex - 1).Left
    End If

    ' Display the color and name.
    lblName(MaxIndex).Caption = color_name
    lblColor(MaxIndex).BackColor = color_value

    ' Make the controls visible.
    lblColor(MaxIndex).Visible = True
    lblName(MaxIndex).Visible = True
End Sub

' Display the colors and their names.
Private Sub Form_Load()
    ' Calculate the row spacing.
    RowHgt = lblColor(0).Height + 30

    ' Position the first controls.
    lblColor(0).Top = 30 - RowHgt
    lblName(0).Top = lblColor(0).Top + (lblColor(0).Height - lblName(0).Height) / 2

    ' Display the colors.
    ShowColor "vbScrollBars", vbScrollBars
    ShowColor "Desktop", vbDesktop
    ShowColor "ActiveTitleBar", vbActiveTitleBar
    ShowColor "InactiveTitleBar", vbInactiveTitleBar
    ShowColor "MenuBar", vbMenuBar
    ShowColor "WindowBackground", vbWindowBackground
    ShowColor "WindowFrame", vbWindowFrame
    ShowColor "MenuText", vbMenuText
    ShowColor "WindowText", vbWindowText
    ShowColor "TitleBarText", vbTitleBarText
    ShowColor "ActiveBorder", vbActiveBorder
    ShowColor "InactiveBorder", vbInactiveBorder
    ShowColor "ApplicationWorkspace", vbApplicationWorkspace
    ShowColor "Highlight", vbHighlight
    ShowColor "HighlightText", vbHighlightText
    ShowColor "ButtonFace", vbButtonFace
    ShowColor "ButtonShadow", vbButtonShadow
    ShowColor "GrayText", vbGrayText
    ShowColor "ButtonText", vbButtonText
    ShowColor "InactiveCaptionText", vbInactiveCaptionText
    ShowColor "3DHighlight", vb3DHighlight
    ShowColor "3DDKShadow", vb3DDKShadow
    ShowColor "3DLight", vb3DLight
    ShowColor "InfoText", vbInfoText
    ShowColor "InfoBackground", vbInfoBackground

    ' Resize the form.
    Height = lblColor(13).Top + lblColor(13).Height + 30 + Height - ScaleHeight
    Width = lblColor(MaxIndex).Left + lblColor(MaxIndex).Width + 30 + Width - ScaleWidth
End Sub
