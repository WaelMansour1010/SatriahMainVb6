VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form FrmFarmer3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ”ÃÌ· ÌÊ„Ì… œÊ—… œÊ«Ã‰"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "FrmFarmer3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   7110
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
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   7095
      _cx             =   12515
      _cy             =   5318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   8454143
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "«·«œÊÌ…|«·⁄·›|«‰ «Ã «·»Ì÷|«·«Ê“«‰"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin VB.Image Image4 
         Height          =   2910
         Left            =   8040
         Picture         =   "FrmFarmer3.frx":000C
         Top             =   330
         Width           =   7260
      End
      Begin VB.Image Image3 
         Height          =   2925
         Left            =   7740
         Picture         =   "FrmFarmer3.frx":778C
         Top             =   330
         Width           =   7185
      End
      Begin VB.Image Image2 
         Height          =   2760
         Left            =   45
         Picture         =   "FrmFarmer3.frx":E702
         Top             =   330
         Width           =   7155
      End
   End
   Begin VB.Image Image1 
      Height          =   7965
      Left            =   0
      Picture         =   "FrmFarmer3.frx":16DDE
      Top             =   0
      Width           =   7320
   End
End
Attribute VB_Name = "FrmFarmer3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub

