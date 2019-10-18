VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPickPtr 
   Caption         =   "PickPtr"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   1935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPickPrinter 
      Caption         =   "Pick Printer"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   1800
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPickPtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Let the user select a printer.
Private Sub cmdPickPrinter_Click()
    dlgPrinter.CancelError = True

    On Error Resume Next
    dlgPrinter.ShowPrinter
    If Err.Number = cdlCancel Then
        ' The user canceled. Do nothing.
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' Unexpected error. Report it.
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting printer." & vbCrLf & _
            Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' Print to the newly selected Printer object.
    Printer.Print "Selected printer: " & Printer.DeviceName
    Printer.EndDoc
End Sub
