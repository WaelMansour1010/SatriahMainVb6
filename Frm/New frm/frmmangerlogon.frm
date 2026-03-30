VERSION 5.00
Begin VB.Form frmmangerlogon 
   BackColor       =   &H00000000&
   Caption         =   "œŒÊ· „œÌ— «·‰ﬁÿ…"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6435
   Icon            =   "frmmangerlogon.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "«·€«¡"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Height          =   375
      Left            =   4680
      Picture         =   "frmmangerlogon.frx":000C
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«œŒ· «·—ﬁ„ «·”—Ì/ «·»’„…"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2850
      Left            =   120
      Picture         =   "frmmangerlogon.frx":1C77
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmmangerlogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
