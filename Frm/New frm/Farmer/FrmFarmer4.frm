VERSION 5.00
Begin VB.Form FrmFarmer4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÚÑÖ ÌÏæá  íæãíÉ ÏæÑÇÊ ÇáÏæÇÌä"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16740
   Icon            =   "FrmFarmer4.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmFarmer4.frx":000C
   RightToLeft     =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   16740
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   9690
      Left            =   -120
      Picture         =   "FrmFarmer4.frx":30177
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16875
   End
End
Attribute VB_Name = "FrmFarmer4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
