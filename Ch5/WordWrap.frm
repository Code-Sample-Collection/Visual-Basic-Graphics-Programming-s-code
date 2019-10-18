VERSION 5.00
Begin VB.Form frmWordWrap 
   Caption         =   "WordWrap"
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
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
End
Attribute VB_Name = "frmWordWrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Print a string on a Printer or PictureBox,
' wrapped within the margins.
Private Sub PrintWrappedText(ByVal txt As String, ByVal indent As Single, ByVal left_margin As Single, ByVal top_margin As Single, ByVal right_margin As Single, ByVal bottom_margin As Single)
Dim next_paragraph As String
Dim next_word As String
Dim pos As Integer

    ' Start at the top of the page.
    Printer.CurrentY = top_margin

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
                    Printer.NewPage
                    Printer.CurrentX = left_margin
                    Printer.CurrentY = top_margin
                End If
            End If

            ' Now print the word. The ; makes the
            ' Printer not move to the next line.
            Printer.Print next_word & " ";
        Loop

        ' Finish the paragraph by ending the line.
        Printer.Print
    Loop
End Sub
' Print the text.
Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass

    ' Select a big font so we have more than one page.
    Printer.Font.Size = 20

    ' Print the text.
    PrintWrappedText txtLongText.Text, _
        720, 1440, 1440, _
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
    cmdPrint.Move _
        (ScaleWidth - cmdPrint.Width) / 2, _
        ScaleHeight - cmdPrint.Height - 120
End Sub


