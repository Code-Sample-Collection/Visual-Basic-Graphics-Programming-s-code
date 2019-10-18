VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRichEdit 
   Caption         =   "RichEdit"
   ClientHeight    =   4440
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   4800
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "RTF Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
      FilterIndex     =   1
   End
   Begin ComctlLib.Toolbar tbrButtons 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveAs"
            Object.ToolTipText     =   "Save As"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Strikethru"
            Object.ToolTipText     =   "Strikethru"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AlignLeft"
            Object.ToolTipText     =   "Align Left"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AlignCenter"
            Object.ToolTipText     =   "Align Center"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AlignRight"
            Object.ToolTipText     =   "Align Right"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Hanging"
            Object.ToolTipText     =   "Hanging Indent"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BulletedList"
            Object.ToolTipText     =   "Bulleted List"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IncreaseIndentation"
            Object.ToolTipText     =   "Increase Indentation"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DecreaseIndentation"
            Object.ToolTipText     =   "Decrease Indentation"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Black"
            Object.ToolTipText     =   "Black"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Red"
            Object.ToolTipText     =   "Red"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Green"
            Object.ToolTipText     =   "Green"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Blue"
            Object.ToolTipText     =   "Blue"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlButtons 
      Left            =   4800
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0448
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":055A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":066C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0890
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":09A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":0EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":100E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":1120
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":1232
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RichEdit.frx":1344
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rchText 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
      _Version        =   393217
      TextRTF         =   $"RichEdit.frx":1456
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
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent File"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent File"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent File"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent File"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "F&ont"
      Begin VB.Menu mnuFontBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFontItalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFontUnderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFontStrikethru 
         Caption         =   "&Strikethru"
      End
   End
   Begin VB.Menu mnuParagraph 
      Caption         =   "&Paragraph"
      Begin VB.Menu mnuParagraphAlignLeft 
         Caption         =   "Align &Left"
      End
      Begin VB.Menu mnuParagraphAlignCenter 
         Caption         =   "Align &Center"
      End
      Begin VB.Menu mnuParagraphAlignRight 
         Caption         =   "Align &Right"
      End
      Begin VB.Menu mnuParagraphSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParagraphBulletedList 
         Caption         =   "B&ulleted List"
      End
      Begin VB.Menu mnuParagraphSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParagraphIncreaseIndentation 
         Caption         =   "&Increase Indentation"
      End
      Begin VB.Menu mnuParagraphDecreaseIndentation 
         Caption         =   "&Decrease Indentation"
      End
      Begin VB.Menu mnuParagraphHangingIndent 
         Caption         =   "&Hanging Indent"
      End
   End
   Begin VB.Menu mnuColor 
      Caption         =   "&Color"
      Begin VB.Menu mnuSetColor 
         Caption         =   "&Black"
         Index           =   0
      End
      Begin VB.Menu mnuSetColor 
         Caption         =   "&Red"
         Index           =   1
      End
      Begin VB.Menu mnuSetColor 
         Caption         =   "&Green"
         Index           =   2
      End
      Begin VB.Menu mnuSetColor 
         Caption         =   "B&lue"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmRichEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Indicates whether the data has been modified since
' saved or loaded.
Private DataModified As Boolean

' The currently loaded file.
Private FileName As String
Private FileTitle As String
' Return true if it is safe to discard the data.
Private Function DataIsSafe() As Boolean
    If DataModified Then
        Select Case MsgBox("The data has been modified. Do you want to save the changes?", vbYesNoCancel)
            Case vbYes
                mnuFileSave_Click
                DataIsSafe = Not DataModified
            Case vbNo
                DataIsSafe = True
            Case vbCancel
                DataIsSafe = False
        End Select
    Else
        DataIsSafe = True
    End If
End Function
Private Sub rchText_KeyPress(KeyAscii As Integer)
Const CTRL_B = 2
Const CTRL_I = 9
Const CTRL_U = 21
Const CTRL_S = 19

    If KeyAscii = CTRL_B Then
        rchText.SelBold = Not rchText.SelBold
        KeyAscii = 0
    End If
    If KeyAscii = CTRL_I Then
        rchText.SelItalic = Not rchText.SelItalic
        KeyAscii = 0
    End If
    If KeyAscii = CTRL_U Then
        rchText.SelUnderline = Not rchText.SelUnderline
        KeyAscii = 0
    End If
    If KeyAscii = CTRL_S Then
        rchText.SelStrikeThru = Not rchText.SelStrikeThru
        KeyAscii = 0
    End If

    ' Recheck the check box values.
    If KeyAscii = 0 Then rchText_SelChange
End Sub
' Set the menu item and button states for the
' selected text.
Private Sub rchText_SelChange()
Dim i As Integer

    If rchText.SelBold Then
        tbrButtons.Buttons("Bold").value = tbrPressed
        mnuFontBold.Checked = True
    Else
        tbrButtons.Buttons("Bold").value = tbrUnpressed
        mnuFontBold.Checked = False
    End If

    If rchText.SelItalic Then
        tbrButtons.Buttons("Italic").value = tbrPressed
        mnuFontItalic.Checked = True
    Else
        tbrButtons.Buttons("Italic").value = tbrUnpressed
        mnuFontItalic.Checked = False
    End If

    If rchText.SelUnderline Then
        tbrButtons.Buttons("Underline").value = tbrPressed
        mnuFontUnderline.Checked = True
    Else
        tbrButtons.Buttons("Underline").value = tbrUnpressed
        mnuFontUnderline.Checked = False
    End If
    If rchText.SelStrikeThru Then
        tbrButtons.Buttons("Strikethru").value = tbrPressed
        mnuFontStrikethru.Checked = True
    Else
        tbrButtons.Buttons("Strikethru").value = tbrUnpressed
        mnuFontStrikethru.Checked = False
    End If
    If rchText.SelBullet Then
        tbrButtons.Buttons("BulletedList").value = tbrPressed
        mnuParagraphBulletedList.Checked = True
    Else
        tbrButtons.Buttons("BulletedList").value = tbrUnpressed
        mnuParagraphBulletedList.Checked = False
    End If
    If rchText.SelHangingIndent > 0 Then
        tbrButtons.Buttons("Hanging").value = tbrPressed
        mnuParagraphHangingIndent.Checked = True
    Else
        tbrButtons.Buttons("Hanging").value = tbrUnpressed
        mnuParagraphHangingIndent.Checked = False
    End If

    tbrButtons.Buttons("AlignLeft").value = tbrUnpressed
    tbrButtons.Buttons("AlignCenter").value = tbrUnpressed
    tbrButtons.Buttons("AlignRight").value = tbrUnpressed
    mnuParagraphAlignLeft.Checked = False
    mnuParagraphAlignCenter.Checked = False
    mnuParagraphAlignRight.Checked = False
    Select Case rchText.SelAlignment
        Case rtfLeft
            tbrButtons.Buttons("AlignLeft").value = tbrPressed
            mnuParagraphAlignLeft.Checked = True
        Case rtfCenter
            tbrButtons.Buttons("AlignCenter").value = tbrPressed
            mnuParagraphAlignCenter.Checked = True
        Case rtfRight
            tbrButtons.Buttons("AlignRight").value = tbrPressed
            mnuParagraphAlignRight.Checked = True
    End Select

    tbrButtons.Buttons("Black").value = tbrUnpressed
    tbrButtons.Buttons("Red").value = tbrUnpressed
    tbrButtons.Buttons("Green").value = tbrUnpressed
    tbrButtons.Buttons("Blue").value = tbrUnpressed
    For i = 0 To 3
        mnuSetColor(i).Checked = False
    Next i
    Select Case rchText.SelColor
        Case vbBlack
            tbrButtons.Buttons("Black").value = tbrPressed
            mnuSetColor(0).Checked = True
        Case vbRed
            tbrButtons.Buttons("Red").value = tbrPressed
            mnuSetColor(1).Checked = True
        Case vbGreen
            tbrButtons.Buttons("Green").value = tbrPressed
            mnuSetColor(2).Checked = True
        Case vbBlue
            tbrButtons.Buttons("Blue").value = tbrPressed
            mnuSetColor(3).Checked = True
    End Select

    tbrButtons.Refresh
End Sub
' Mark the data as mofified.
Private Sub SetModified()
    ' Do nothing if the data is already modified.
    If DataModified Then Exit Sub

    DataModified = True
    Caption = "RichEdit*[" & FileTitle & "]"
End Sub

Private Sub Form_Load()
Dim btn As Integer
Dim pic As Integer

    dlgFile.InitDir = App.Path

    ' Prepare the toolbar buttons. This is done here
    ' because it is a hassle to do during design time.
    ' To change the images, you need to disassociate
    ' the toolbar from the image list and then the
    ' toolbar loses this information.
    tbrButtons.ImageList = imlButtons
    pic = 1
    For btn = 1 To tbrButtons.Buttons.Count
        With tbrButtons.Buttons(btn)
            ' See if this button is a button.
            If .Style = tbrDefault Then
                ' Give this button the next picture.
                .Image = pic
                pic = pic + 1
            End If
        End With
    Next btn
End Sub
' See if it is safe to unload.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not DataIsSafe()
End Sub

' Make the RichTextBox as large as possible.
Private Sub Form_Resize()
Dim hgt As Single

    If WindowState = vbMinimized Then Exit Sub
    hgt = ScaleHeight - tbrButtons.Height
    If hgt < 120 Then hgt = 120
    rchText.Move 0, tbrButtons.Height, ScaleWidth, hgt
End Sub


' Unload. If the data is not safe, the QueryUnload
' event handler will stop the unload.
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Start a new file.
Private Sub mnuFileNew_Click()
    ' Make sure the data is safe.
    If Not DataIsSafe() Then Exit Sub

    ' Start over.
    FileTitle = ""
    FileName = ""
    rchText.Text = ""
    Caption = "RichEdit []"
    DataModified = False
End Sub
' Open a file.
Private Sub mnuFileOpen_Click()
    ' Make sure the data is safe.
    If Not DataIsSafe() Then Exit Sub

    ' Get the file.
    dlgFile.Flags = _
        cdlOFNExplorer + _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number > 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file" & vbCrLf & Err.Description
        Exit Sub
    End If

    ' Read the file.
    On Error GoTo ReadErr
    rchText.LoadFile dlgFile.FileName
    On Error GoTo 0

    FileName = dlgFile.FileName
    FileTitle = dlgFile.FileTitle
    Caption = "RichEdit [" & FileTitle & "]"
    DataModified = False

    Exit Sub

ReadErr:
    MsgBox "Error " & Format$(Err.Number) & _
        " reading file '" & dlgFile.FileName & _
        "'" & vbCrLf & Err.Description
    Exit Sub
End Sub

' Save the file.
Private Sub mnuFileSave_Click()
    ' See if we have a file name.
    If Len(FileTitle) = 0 Then
        mnuFileSaveAs_Click
        Exit Sub
    End If

    ' Save the file.
    On Error GoTo SaveErr
    rchText.SaveFile FileName
    On Error GoTo 0

    Caption = "RichEdit [" & FileTitle & "]"
    DataModified = False

    Exit Sub

SaveErr:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file '" & FileName & _
        "'" & vbCrLf & Err.Description
    Exit Sub
End Sub
' Save the file with a new name.
Private Sub mnuFileSaveAs_Click()
    ' Get the file name.
    dlgFile.Flags = _
        cdlOFNExplorer + _
        cdlOFNPathMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNOverwritePrompt
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number > 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file" & vbCrLf & Err.Description
        Exit Sub
    End If

    ' Save the file.
    On Error GoTo SaveAsErr
    rchText.SaveFile dlgFile.FileName
    On Error GoTo 0

    FileName = dlgFile.FileName
    FileTitle = dlgFile.FileTitle
    Caption = "RichEdit [" & FileTitle & "]"
    DataModified = False

    Exit Sub

SaveAsErr:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file '" & dlgFile.FileName & _
        "'" & vbCrLf & Err.Description
    Exit Sub
End Sub

Private Sub mnuFontBold_Click()
    rchText.SelBold = Not rchText.SelBold
    rchText_SelChange
End Sub

Private Sub mnuFontItalic_Click()
    rchText.SelItalic = Not rchText.SelItalic
    rchText_SelChange
End Sub


Private Sub mnuFontStrikethru_Click()
    rchText.SelStrikeThru = Not rchText.SelStrikeThru
    rchText_SelChange
End Sub


Private Sub mnuFontUnderline_Click()
    rchText.SelUnderline = Not rchText.SelUnderline
    rchText_SelChange
End Sub
Private Sub mnuParagraphAlignCenter_Click()
    rchText.SelAlignment = rtfCenter
    rchText_SelChange
End Sub

Private Sub mnuParagraphAlignLeft_Click()
    rchText.SelAlignment = rtfLeft
    rchText_SelChange
End Sub

Private Sub mnuParagraphAlignRight_Click()
    rchText.SelAlignment = rtfRight
    rchText_SelChange
End Sub


Private Sub mnuParagraphBulletedList_Click()
    If IsNull(rchText.SelBullet) Then
        rchText.SelBullet = True
    Else
        rchText.SelBullet = Not rchText.SelBullet
    End If
    rchText_SelChange
End Sub

Private Sub mnuParagraphDecreaseIndentation_Click()
    rchText.SelIndent = rchText.SelIndent - 240
End Sub

Private Sub mnuParagraphHangingIndent_Click()
    If IsNull(rchText.SelHangingIndent) Then
        rchText.SelHangingIndent = 240
    Else
        rchText.SelHangingIndent = 240 - rchText.SelHangingIndent
    End If
    rchText_SelChange
End Sub
Private Sub mnuParagraphIncreaseIndentation_Click()
    rchText.SelIndent = rchText.SelIndent + 240
End Sub


Private Sub mnuSetColor_Click(Index As Integer)
    Select Case Index
        Case 0
            rchText.SelColor = vbBlack
        Case 1
            rchText.SelColor = vbRed
        Case 2
            rchText.SelColor = vbGreen
        Case 3
            rchText.SelColor = vbBlue
    End Select
    rchText_SelChange
End Sub

' Mark the data as modified.
Private Sub rchText_Change()
    SetModified
End Sub

' Execute a command.
Private Sub tbrButtons_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "SaveAs"
            mnuFileSaveAs_Click
        Case "Bold"
            mnuFontBold_Click
        Case "Italic"
            mnuFontItalic_Click
        Case "Underline"
            mnuFontUnderline_Click
        Case "Strikethru"
            mnuFontStrikethru_Click
        Case "AlignLeft"
            mnuParagraphAlignLeft_Click
        Case "AlignCenter"
            mnuParagraphAlignCenter_Click
        Case "AlignRight"
            mnuParagraphAlignRight_Click
        Case "Hanging"
            mnuParagraphHangingIndent_Click
        Case "BulletedList"
            mnuParagraphBulletedList_Click
        Case "IncreaseIndentation"
            mnuParagraphIncreaseIndentation_Click
        Case "DecreaseIndentation"
            mnuParagraphDecreaseIndentation_Click
        Case "Black"
            mnuSetColor_Click 0
        Case "Red"
            mnuSetColor_Click 1
        Case "Green"
            mnuSetColor_Click 2
        Case "Blue"
            mnuSetColor_Click 3
        Case Else
            MsgBox "Unknown button '" & Button.Key & _
                "'"
            Stop
    End Select
End Sub
