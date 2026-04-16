VERSION 5.00
Begin VB.Form FrmFarmer9 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„Õ«÷— «·«⁄œ«„"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6555
   Icon            =   "FrmFarmer9.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   6555
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
      Height          =   6435
      Left            =   -120
      Picture         =   "FrmFarmer9.frx":000C
      Top             =   0
      Width           =   6675
   End
End
Attribute VB_Name = "FrmFarmer9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
