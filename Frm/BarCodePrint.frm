VERSION 5.00
Begin VB.Form BarCodePrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   16845
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   297.127
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   186.796
   ShowInTaskbar   =   0   'False
   Begin VB.Label DNLbl 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   525
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label UpLbl 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   15
      TabIndex        =   1
      Top             =   165
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image IMG 
      Height          =   227
      Index           =   0
      Left            =   30
      Stretch         =   -1  'True
      Top             =   165
      Width           =   345
   End
End
Attribute VB_Name = "BarCodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

