VERSION 5.00
Begin VB.Form FrmFarmer6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÌÏæá ãÚÇãá ÇáÊÝÑíÎ"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15135
   Icon            =   "FrmFarmer6.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   9630
      Left            =   -120
      Picture         =   "FrmFarmer6.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15285
   End
End
Attribute VB_Name = "FrmFarmer6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
