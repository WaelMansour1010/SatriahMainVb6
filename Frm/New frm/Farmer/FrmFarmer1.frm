VERSION 5.00
Begin VB.Form FrmFarmer1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  «‰Ê«⁄  «·”·«·« "
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   Icon            =   "FrmFarmer1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   6735
      Left            =   -120
      Picture         =   "FrmFarmer1.frx":000C
      Top             =   0
      Width           =   6690
   End
End
Attribute VB_Name = "FrmFarmer1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
