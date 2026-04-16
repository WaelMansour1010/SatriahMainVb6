VERSION 5.00
Begin VB.Form FRMShowPic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "œÊ—… „—«ﬁ»Â «·ÃÊœ…"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8205
   Icon            =   "FRMShowPic.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   8205
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
      Height          =   8670
      Left            =   120
      Picture         =   "FRMShowPic.frx":000C
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "FRMShowPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
End Sub
