VERSION 5.00
Begin VB.Form FrmViewPic 
   BackColor       =   &H00E0E0E0&
   Caption         =   "⁄—÷ «·’Ê—…"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15180
   Icon            =   "FrmViewPic.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   15180
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin Dynamic_Byte.NewViewBox MainView 
      Height          =   7965
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   19995
      _ExtentX        =   35269
      _ExtentY        =   14049
   End
   Begin VB.ComboBox CboViewStyle 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   11580
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Œ—ÊÃ"
      Height          =   495
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ—Ìﬁ… «·⁄—÷"
      Height          =   315
      Left            =   14130
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "FrmViewPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboViewStyle_Click()

    If CboViewStyle.ListIndex <> -1 Then
        Me.MainView.View = CboViewStyle.ItemData(CboViewStyle.ListIndex)
    End If

    Me.MainView.Refresh
End Sub

Private Sub CmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    With Me.CboViewStyle
        .AddItem "⁄«œÏ", 0
        .ItemData(0) = ViewConstants.Normal
        .AddItem "„„ œ…", 1
        .ItemData(1) = ViewConstants.Stretch
    
        .AddItem "ÕÃ„  ·ﬁ«∆Ï", 2
        .ItemData(2) = ViewConstants.AutoSize
    
        .AddItem "≈ŸÂ«— «⁄„œ…  „—Ì—", 3
        .ItemData(3) = ViewConstants.ScrollBars
    
        .AddItem "ÕÃ„ ÿ»Ì⁄Ï", 4
        .ItemData(4) = ViewConstants.RealStrecth
    
        .AddItem " „—Ì— »„ƒ‘— «·›√—…", 5
        .ItemData(5) = ViewConstants.HandScroll
    
    End With
 Me.MainView.View = 1
End Sub

Private Sub Form_Resize()
'    Me.MainView.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight - (CmdOk.Height + 200)
'    Me.CmdOk.Move Me.ScaleLeft, Me.ScaleHeight - CmdOk.Height
'
'    lbl.top = CmdOk.top
'    lbl.left = Me.ScaleWidth - Me.lbl.Width
'
''    Me.CboViewStyle.top = Me.CmdOk.top
 '   Me.CboViewStyle.left = Me.ScaleWidth - (lbl.Width + Me.CboViewStyle.Width)
End Sub

