VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form SERVICE_TYPE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«‰Ê«⁄ «· Õ’Ì·« "
   ClientHeight    =   4620
   ClientLeft      =   7170
   ClientTop       =   2685
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "name"
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
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "value"
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
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2280
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÃœÌœ"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "no"
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
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   4815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   0
      Top             =   3960
      Width           =   6975
      _ExtentX        =   12303
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "services"
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
      Caption         =   "«”„ «· Õ’Ì·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5160
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "«·ﬁÌ„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "«‰Ê«⁄ «· Õ’Ì·« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1320
      TabIndex        =   6
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «· Õ’Ì·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "SERVICE_TYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Update
End Sub

