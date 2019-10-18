VERSION 5.00
Begin VB.Form frmShapes 
   Caption         =   "Shapes"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape shpShape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   3000
      Shape           =   1  'Square
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape shpShape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   480
      Top             =   240
      Width           =   1335
   End
   Begin VB.Shape shpShape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   5
      Left            =   3000
      Shape           =   5  'Rounded Square
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape shpShape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   4
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape shpShape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   3
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   735
   End
   Begin VB.Shape shpShape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   480
      Shape           =   2  'Oval
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblShape 
      Alignment       =   2  'Center
      Caption         =   "vbShapeRoundedSquare"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblShape 
      Alignment       =   2  'Center
      Caption         =   "vbShapeRoundedRectangle"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblShape 
      Alignment       =   2  'Center
      Caption         =   "vbShapeCircle"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblShape 
      Alignment       =   2  'Center
      Caption         =   "vbShapeOval"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblShape 
      Alignment       =   2  'Center
      Caption         =   "vbShapeSquare"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblShape 
      Alignment       =   2  'Center
      Caption         =   "vbShapeRectangle"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




