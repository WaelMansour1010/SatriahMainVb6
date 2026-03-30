VERSION 5.00
Begin VB.Form FrmFarmer5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  œ›⁄Â «·Ï „⁄„· «· ›—ÌŒ"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7995
   Icon            =   "FrmFarmer5.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   8445
      Left            =   -120
      Picture         =   "FrmFarmer5.frx":000C
      Top             =   0
      Width           =   8130
   End
End
Attribute VB_Name = "FrmFarmer5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
