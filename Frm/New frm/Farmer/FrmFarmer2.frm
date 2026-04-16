VERSION 5.00
Begin VB.Form FrmFarmer2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "œÊ—…  —»Ì… œÊ«Ã‰"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7950
   Icon            =   "FrmFarmer2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Image Image1 
      Height          =   6900
      Left            =   -120
      Picture         =   "FrmFarmer2.frx":000C
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "FrmFarmer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
