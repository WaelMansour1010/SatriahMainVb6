VERSION 5.00
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Car_alarms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "     ‰»ÌÂ«   ﬁÿ«⁄ «·‰ﬁ·Ì« "
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "Car_alarms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   7800
   Begin ALLButtonS.ALLButton ALLButton4 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "⁄—÷"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
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
      MICON           =   "Car_alarms.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton5 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "⁄—÷"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
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
      MICON           =   "Car_alarms.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "Car_alarms.frx":0044
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ  «·„ÊŸ›Ì‰  «· Ì ” ‰ ÂÌ  √„Ì‰« Â„"
      Height          =   855
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label d5 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   615
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ Õ«›Ÿ… «·‰›Ê”  «· Ì ” ‰ ÂÌ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   5295
   End
   Begin VB.Label d4 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "     ‰»ÌÂ«   ﬁÿ«⁄ «·‰ﬁ·Ì« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   2
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6600
   End
End
Attribute VB_Name = "Car_alarms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Askinterval As String
Dim Askcount As Integer

Private Sub ALLButton1t_Click()
    FrmCarExpireLicens.show
End Sub



Private Sub ChangeLang()

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    
End Sub
