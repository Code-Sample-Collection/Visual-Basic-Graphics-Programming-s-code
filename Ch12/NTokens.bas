Attribute VB_Name = "NamedTokens"
Option Explicit

' Return a named token from the string txt.
' Tokens have the form TokenName(TokenValue).
Public Sub GetNamedToken(ByRef txt As String, ByRef token_name As String, ByRef token_value As String)
Dim pos1 As Integer
Dim pos2 As Integer
Dim open_parens As Integer
Dim ch As String

    ' Find the "(".
    pos1 = InStr(txt, "(")
    If pos1 = 0 Then
        ' No "(" found. Return the rest as the token name.
        token_name = Trim$(txt)
        token_value = ""
        txt = ""
        Exit Sub
    End If

    ' Find the corresponding ")". Note that
    ' parentheses may be nested.
    open_parens = 1
    pos2 = pos1 + 1
    Do While pos2 <= Len(txt)
        ch = Mid$(txt, pos2, 1)
        If ch = "(" Then
            open_parens = open_parens + 1
        ElseIf ch = ")" Then
            open_parens = open_parens - 1
            If open_parens = 0 Then
                ' This is the corresponding ")".
                Exit Do
            End If
        End If
        pos2 = pos2 + 1
    Loop

    ' At this point, pos1 points to the ( and
    ' pos2 points to the ).
    token_name = Trim$(Left$(txt, pos1 - 1))
    token_value = Trim$(Mid$(txt, pos1 + 1, pos2 - pos1 - 1))
    txt = Trim$(Mid$(txt, pos2 + 1))
End Sub
' Replace non-printable characters with spaces.
Public Function RemoveNonPrintables(ByVal txt As String) As String
Dim pos As Integer
Dim ch As String

    For pos = 1 To Len(txt)
        ch = Mid$(txt, pos, 1)
        If (ch < " ") Or (ch > "~") Then Mid$(txt, pos, 1) = " "
    Next pos

    RemoveNonPrintables = txt
End Function


