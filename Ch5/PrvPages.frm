VERSION 5.00
Begin VB.Form frmPrvPages 
   Caption         =   "PrvPages"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "1"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtLongText 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Page"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   495
   End
End
Attribute VB_Name = "frmPrvPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Print a string on a Printer or PictureBox,
' wrapped within the margins.
Private Sub PrintWrappedText(ByVal ptr As Object, ByVal target_page As Integer, ByVal txt As String, ByVal indent As Single, ByVal left_margin As Single, ByVal top_margin As Single, ByVal right_margin As Single, ByVal bottom_margin As Single)
Dim next_paragraph As String
Dim next_word As String
Dim pos As Integer
Dim current_page As Integer

    ' Start at the top of the page.
    Printer.CurrentY = top_margin

    ' Keep track of the page we are on.
    current_page = 1

    ' Repeat until the text is all printed.
    Do While Len(txt) > 0
        ' Get the next paragraph.
        pos = InStr(txt, vbCrLf)
        If pos = 0 Then
            ' Use the rest of the text.
            next_paragraph = Trim$(txt)
            txt = ""
        Else
            ' Get the paragraph.
            next_paragraph = Trim$(Left$(txt, pos - 1))
            txt = Mid$(txt, pos + Len(vbCrLf))
        End If

        ' Indent the paragraph.
        Printer.CurrentX = left_margin + indent

        ' Print the paragraph.
        Do While Len(next_paragraph) > 0
            ' Get the next word.
            pos = InStr(next_paragraph, " ")
            If pos = 0 Then
                ' Use the rest of the paragraph.
                next_word = next_paragraph
                next_paragraph = ""
            Else
                ' Get the word.
                next_word = Left$(next_paragraph, pos - 1)
                next_paragraph = Trim$(Mid$(next_paragraph, pos + 1))
            End If

            ' See if there is room for this word.
            If Printer.CurrentX + Printer.TextWidth(next_word) _
                > right_margin _
            Then
                ' It won't fit. Start a new line.
                Printer.Print
                Printer.CurrentX = left_margin

                ' See if we have room for a new line.
                If Printer.CurrentY + Printer.TextHeight(next_word) _
                    > bottom_margin _
                Then
                    ' Start a new page.
                    current_page = current_page + 1
                    If current_page > target_page Then Exit Sub
                    Printer.CurrentX = left_margin
                    Printer.CurrentY = top_margin
                End If
            End If

            ' Now print the word. The ; makes the
            ' Printer not move to the next line.
            If current_page = target_page Then
                ' Print the word.
                If ptr Is Printer Then
                    ' This is the printer. Print.
                    ptr.Print next_word & " ";
                Else
                    ' This is not the printer. Go to
                    ' the printer's position and print.
                    ptr.CurrentX = Printer.CurrentX
                    ptr.CurrentY = Printer.CurrentY
                    ptr.Print next_word & " ";

                    ' Skip space for the word.
                    Printer.CurrentX = Printer.CurrentX + _
                        Printer.TextWidth(next_word & " ")
                End If
            Else
                ' Skip space for the word.
                Printer.CurrentX = Printer.CurrentX + _
                    Printer.TextWidth(next_word & " ")
            End If
        Loop

        ' Finish the paragraph by ending the line.
        Printer.Print
    Loop

    ' See if we got to the desired page yet.
    If current_page < target_page Then
        MsgBox "This page does not exist."
    End If
End Sub
' Preview the text.
Private Sub cmdPreview_Click()
    MousePointer = vbHourglass

    ' Select a big font so we have more than one page.
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 20

    ' Make the preview PictureBox use the same font.
    With frmPreview.picInner.Font
        .Name = Printer.Font.Name
        .Bold = Printer.Font.Bold
        .Italic = Printer.Font.Italic
        .Strikethrough = Printer.Font.Strikethrough
        .Underline = Printer.Font.Underline
        .Size = Printer.Font.Size
    End With
    frmPreview.picInner.ScaleMode = Printer.ScaleMode
    frmPreview.picInner.Move 0, 0, Printer.ScaleWidth, Printer.ScaleHeight

    ' Print the text.
    PrintWrappedText _
        frmPreview.picInner, CInt(txtPage.Text), _
        txtLongText.Text, 720, 1440, 1440, _
        Printer.ScaleWidth - 1440, _
        Printer.ScaleHeight - 1440
    Printer.KillDoc

    ' Make the preview form tell what page it shows.
    frmPreview.Caption = "Preview Page " & Format$(txtPage.Text)

    ' Display the preview.
    frmPreview.Show vbModal

    MousePointer = vbDefault
End Sub
' Print the text.
Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass

    ' Select a big font so we have more than one page.
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 20

    ' Print the text.
    PrintWrappedText _
        Printer, CInt(txtPage.Text), _
        txtLongText.Text, 720, 1440, 1440, _
        Printer.ScaleWidth - 1440, _
        Printer.ScaleHeight - 1440
    Printer.EndDoc

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim fname As String
Dim fnum As Integer
Dim txt As String

    fname = App.Path
    If Right$(fname, 1) <> "\" Then fname = fname & "\"
    fname = fname & "longtext.txt"

    On Error Resume Next
    fnum = FreeFile
    Open fname For Input As fnum
    txt = Input$(LOF(fnum), fnum)
    Close fnum

    txtLongText.Text = txt
End Sub


' Make the TextBox as big as possible.
Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single

    wid = ScaleWidth - 2 * 120
    If wid < 120 Then wid = 120
    hgt = ScaleHeight - cmdPrint.Height - 3 * 120
    If hgt < 120 Then hgt = 120

    txtLongText.Move 120, 120, wid, hgt
    cmdPreview.Top = ScaleHeight - cmdPrint.Height - 120
    cmdPrint.Top = cmdPreview.Top
    Label1.Top = cmdPreview.Top
    txtPage.Top = cmdPreview.Top
End Sub
