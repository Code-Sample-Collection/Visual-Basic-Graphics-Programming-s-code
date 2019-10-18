Attribute VB_Name = "Tokens"
Option Explicit

' Return the next delimited token from txt.
' Trim blanks.
Public Function GetDelimitedToken(ByRef txt As String, ByVal delimiter As String) As String
Dim pos As Integer

    pos = InStr(txt, delimiter)
    If pos < 1 Then
        ' The delimiter was not found. Return
        ' the rest of txt.
        GetDelimitedToken = Trim$(txt)
        txt = ""
    Else
        ' We found the delimiter. Return the token.
        GetDelimitedToken = Trim$(Left$(txt, pos - 1))
        txt = Trim$(Mid$(txt, pos + Len(delimiter)))
    End If
End Function
' Replace non-printable characters in txt with
' spaces.
Public Function NonPrintingToSpace(ByVal txt As String) As String
Dim i As Integer
Dim ch As String

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If (ch < " ") Or (ch > "~") Then Mid$(txt, i, 1) = " "
    Next i
    NonPrintingToSpace = txt
End Function


' Remove comments starting with ' from the
' end of lines.
Public Function RemoveComments(ByVal txt As String) As String
Dim pos As Integer
Dim new_txt As String

    Do While Len(txt) > 0
        ' Find the next '.
        pos = InStr(txt, "'")
        If pos = 0 Then
            new_txt = new_txt & txt
            Exit Do
        End If

        ' Add this part to the result.
        new_txt = new_txt & Left$(txt, pos - 1)

        ' Find the end of the line.
        pos = InStr(pos + 1, txt, vbCrLf)
        If pos = 0 Then
            ' There was no vbCrLf.
            ' Remove the rest of the text.
            txt = ""
        Else
            txt = Mid$(txt, pos + Len(vbCrLf))
        End If
    Loop

    RemoveComments = new_txt
End Function


