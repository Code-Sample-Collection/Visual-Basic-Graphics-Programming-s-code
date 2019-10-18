VERSION 5.00
Begin VB.Form frmPrinters 
   Caption         =   "Printers"
   ClientHeight    =   4140
   ClientLeft      =   2325
   ClientTop       =   1260
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   4485
   Begin VB.TextBox txtPrinters 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Load the list of printers.
Private Sub Form_Load()
Dim pr As Printer
Dim txt As String
Dim device_name As String
Dim device_fmt As String
Dim device_len As Integer
Dim port_name As String
Dim port_fmt As String
Dim port_len As Integer
Dim driver_name As String

    ' See how long each field is.
    For Each pr In Printers
        If device_len < Len(pr.DeviceName) Then device_len = Len(pr.DeviceName)
        If port_len < Len(pr.Port) Then port_len = Len(pr.Port)
    Next pr

    ' Build some formatting strings.
    device_fmt = Left$("!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@", device_len + 3)
    port_fmt = Left$("!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@", port_len + 3)

    ' Build the output string.
    txt = Format$("Device", device_fmt) & _
          Format$("Port", port_fmt) & _
          "Driver" & vbCrLf
    txt = txt & _
          Format$("------", device_fmt) & _
          Format$("----", port_fmt) & _
          "------" & vbCrLf
    For Each pr In Printers
        txt = txt & _
              Format$(pr.DeviceName, device_fmt) & _
              Format$(pr.Port, port_fmt) & _
              pr.DriverName & vbCrLf
    Next pr

    txtPrinters.Text = txt
End Sub

Private Sub Form_Resize()
    txtPrinters.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
