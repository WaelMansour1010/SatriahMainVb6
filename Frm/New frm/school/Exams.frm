VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Exams 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘…  ⁄—Ì› «·«„ Õ«‰« "
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   5415
   Begin VB.CommandButton Command40 
      Height          =   492
      Index           =   1
      Left            =   600
      Picture         =   "Exams.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   492
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   " ⁄—Ì› «·«„ Õ«‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   " «—ÌŒ «·«„ Õ«‰"
      Height          =   375
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·«„ Õ«‰"
      Height          =   375
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·«„ Õ«‰"
      Height          =   375
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Exams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
