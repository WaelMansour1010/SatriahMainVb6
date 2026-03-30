VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form MEMBER_CHILD 
   Caption         =   "«·«⁄÷«¡ «· «»⁄Ì‰"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "«· «»⁄Ì‰"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "IMAGE_PATH"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÃœÌœ"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "MEMBER_CHILD.frx":0000
      DataField       =   "MEMBER_TYPE"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "MEMBER_NAME"
      BoundColumn     =   "MEMBER_ID"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   0
      Top             =   4320
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1720
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=MAY.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=MAY.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "MEMBERS"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   5400
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=MAY.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=MAY.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "MEMBER_TYPES"
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
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·⁄÷Ê"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   11
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·⁄÷Ê"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·⁄÷ÊÌ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   9
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "«·’Ê—…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "„·› «·⁄÷Ê"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "MEMBER_CHILD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveLast
End Sub
