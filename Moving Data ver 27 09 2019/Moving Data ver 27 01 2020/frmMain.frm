VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "«Œ—Ï"
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Top             =   870
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ÕœÌÀ «·„»Ì⁄« "
      Height          =   495
      Left            =   2100
      TabIndex        =   0
      Top             =   840
      Width           =   1275
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FRMTRansferData.Show
End Sub

Private Sub Command2_Click()
FRMTRansferData2.Show
End Sub
