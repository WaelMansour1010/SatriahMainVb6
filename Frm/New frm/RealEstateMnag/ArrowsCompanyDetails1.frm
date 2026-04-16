VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form ArrowsCompanyDetails1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8790
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   13335
   Icon            =   "ArrowsCompanyDetails1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13335
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton Cmd 
      Caption         =   "«⁄«œ…  Õ„Ì· «·’›Õ…"
      Height          =   315
      Index           =   0
      Left            =   11400
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   13335
      ExtentX         =   23521
      ExtentY         =   13996
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label LblName 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblSymbol 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·‘—ﬂ…"
      Height          =   255
      Index           =   0
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—„“ «·‘—ﬂÂ"
      Height          =   255
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "ArrowsCompanyDetails1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pathH As String
Dim NameH As String
Dim SymbolH As String

Dim NEW_interface As Boolean

Private Sub Cmd_Click(Index As Integer)
    LoadPage pathH, SymbolH, NameH
End Sub

Function LoadPage(path As String, Optional Symbol As String, Optional name As String)
    WebBrowser1.Navigate2 path
    pathH = path
    SymbolH = Symbol
    NameH = name
    lblSymbol = SymbolH
    LblName = NameH

End Function

Private Sub Form_Load()
    Resize_Form Me
End Sub
 
