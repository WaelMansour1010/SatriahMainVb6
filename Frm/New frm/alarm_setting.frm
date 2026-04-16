VERSION 5.00
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form alarm_setting 
   Caption         =   "   ≈⁄œ«œ  «· ‰»ÌÂ«   "
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Õ›Ÿ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "alarm_setting.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3840
      Width           =   4935
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "alarm_setting.frx":001C
         Left            =   120
         List            =   "alarm_setting.frx":0029
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ ﬁÌ·"
         Height          =   255
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3120
      Width           =   4935
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "alarm_setting.frx":003C
         Left            =   120
         List            =   "alarm_setting.frx":0049
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ ﬁÌ·"
         Height          =   255
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Width           =   4935
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "alarm_setting.frx":005C
         Left            =   120
         List            =   "alarm_setting.frx":0069
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ ﬁÌ·"
         Height          =   255
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   4935
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ ﬁÌ·"
         Height          =   255
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "alarm_setting.frx":007C
         Left            =   120
         List            =   "alarm_setting.frx":0089
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   4935
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "alarm_setting.frx":009C
         Left            =   120
         List            =   "alarm_setting.frx":00A9
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ ﬁÌ·"
         Height          =   255
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ «·«ﬁ«„«  «·„‰ ÂÌ…"
      Height          =   615
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "   ≈⁄œ«œ  «· ‰»ÌÂ«   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1095
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   8520
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ «·ÃÊ«“«  «·„‰ ÂÌ…"
      Height          =   615
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ —Œ’ «·ﬁÌ«œ… «·„‰ ÂÌ…"
      Height          =   615
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ Õ«›Ÿ… «·‰›Ê” «·„‰ÂÌ…"
      Height          =   615
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ  «·„ÊŸ›Ì‰ «·„‰ ÂÌ  √„Ì‰« Â„"
      Height          =   855
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4080
      Width           =   4455
   End
End
Attribute VB_Name = "alarm_setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub d3_Click()
End Sub

Private Sub d1_Click()
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
End Sub
