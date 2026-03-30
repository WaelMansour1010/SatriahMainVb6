VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ADD_PICTURE_FROM 
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "«Œ Ì«— «·’Ê—…"
      Height          =   1095
      Left            =   6360
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5400
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "«œ—«Ã"
      Height          =   1095
      Left            =   -360
      TabIndex        =   1
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«œ—«Ã"
      Height          =   1095
      Left            =   4440
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "„‰ ›÷·ﬂ «Œ «— «·’Ê—…"
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   480
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "ADD_PICTURE_FROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MEMBERS.Text5 = Text1
MEMBERS.Adodc1.Recordset.Update
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName

Image1.Picture = LoadPicture(Text1.Text)

End Sub

Private Sub Drive1_Change()
File1.FileName = Drive1.Drive
End Sub

