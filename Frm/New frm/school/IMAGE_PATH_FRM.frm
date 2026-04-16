VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form IMAGE_PATH_FRM 
   Caption         =   " ÕœÌœ „”«— «·»—‰«„Ã"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7785
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
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Õ›Ÿ"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5880
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox IMAGE_PATH 
      DataField       =   "IMAGE_PATH"
      DataSource      =   "Adodc3"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   7335
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   1440
      Top             =   2520
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "IMAGE_PATH"
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
   Begin VB.Label Label1 
      Caption         =   "„”«— «·’Ê—…"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   3960
      Width           =   2175
   End
End
Attribute VB_Name = "IMAGE_PATH_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc3.Recordset.Fields!IMAGE_PATH = IMAGE_PATH
Adodc3.Recordset.Update
End Sub

Private Sub Dir1_Change()
IMAGE_PATH.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

