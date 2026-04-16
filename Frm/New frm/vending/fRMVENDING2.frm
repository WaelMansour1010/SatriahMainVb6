VERSION 5.00
Begin VB.Form fRMVENDING2 
   Caption         =   "Convert Unix Timestamps"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   13755
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8355
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "fRMVENDING2.frx":0000
      Top             =   5280
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1455
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtIn 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5520
      Width           =   3135
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   9720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3120
      Width           =   2955
   End
End
Attribute VB_Name = "fRMVENDING2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim F As Integer
    Dim JSON As String
    Dim Candles As JsonBag
    Dim I As Long
    Dim DateValue As Date
    
  '  F = FreeFile(0)
  '  Open App.Path & "\sample.txt" For Input As #F
  '  JSON = Input$(LOF(F), #F)
  '  Close #F
  JSON = Text1.Text
    With New JsonBag
        .DecimalMode = True
        .JSON = JSON
        
 '       .Whitespace = True
        txtIn.Text = .JSON
        
        Set Candles = .Item("result")
        For I = 1 To Candles.Count
        MsgBox Candles(I).Item("transaction_id")
        MsgBox Candles(I).Item("machine_id")
        Next
        
        txtOut.Text = .JSON
    End With
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        txtIn.Move 0, 0, ScaleWidth / 2, ScaleHeight
        txtOut.Move ScaleWidth / 2, 0, ScaleWidth / 2, ScaleHeight
    End If
End Sub
