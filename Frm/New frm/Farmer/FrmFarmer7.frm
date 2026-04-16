VERSION 5.00
Begin VB.Form FrmFarmer7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÕÃÊ“«  «·⁄„·«¡"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   Icon            =   "FrmFarmer7.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   7350
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
      Height          =   7965
      Left            =   -120
      Picture         =   "FrmFarmer7.frx":000C
      Top             =   -360
      Width           =   7590
   End
End
Attribute VB_Name = "FrmFarmer7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
