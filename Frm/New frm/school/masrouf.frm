VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form masrouf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‘«‘… «·„’—Ê›« "
   ClientHeight    =   4665
   ClientLeft      =   5415
   ClientTop       =   3495
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "«÷«›… ‰Ê⁄ ÃœÌœ"
      Height          =   255
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "masrouf.frx":0000
      Top             =   1920
      Width           =   5175
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3000
      TabIndex        =   8
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "ÃœÌœ"
      Height          =   612
      Left            =   -120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   2172
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "masrouf.frx":0004
      DataField       =   "OPERATION_TYPE"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "name"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "Fines_VALUE"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3000
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "”œ«œ"
      Enabled         =   0   'False
      Height          =   615
      Left            =   -120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   840
      Top             =   5760
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "SELECT * FROM operations WHERE MEMBER_ID=0"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   7200
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
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
      RecordSource    =   "masrouf"
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
   Begin MSAdodcLib.Adodc Adodc20 
      Height          =   375
      Left            =   960
      Top             =   600
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      RecordSource    =   "OPERATIONS"
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
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "„·«ÕŸ« "
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
      Left            =   5400
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·«Ì’«·"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      Left            =   5400
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·„’—Ê›"
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
      Left            =   5280
      TabIndex        =   3
      Top             =   240
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
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "masrouf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Command1_Click()
If Text5.Text = "" Then
MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·«Ì’«·", vbCritical

Exit Sub
End If


If Not IsNumeric(Text5.Text) Then
MsgBox "    —ﬁ„ «·«Ì’«· ÌÃ» «‰ ÌﬂÊ‰ «—ﬁ«„", vbCritical
Text5.Text = ""
Exit Sub
End If

Adodc20.CommandType = adCmdText
Adodc20.RecordSource = "select * from operations where bill_no=" & Text5.Text
Adodc20.Refresh

If Adodc20.Recordset.RecordCount > 0 Then
MsgBox "—ﬁ„ «·«Ì’«·  „” Œœ„ „‰ ﬁ»·", vbCritical
Exit Sub
End If






Adodc1.Recordset.Fields!operation_type = DataCombo1.Text
Adodc1.Recordset.Fields!TOTAL_VALUE = -(Text3.Text)

Adodc1.Recordset.Fields!MEMBER_NAME = "„’—Ê›« "
Adodc1.Recordset.Fields!ACTUAL_VALUE = -(Text3.Text)
Adodc1.Recordset.Fields!OPERATION_DATE = DateValue(Now)
Adodc1.Recordset.Fields!notes = Text8.Text
Adodc1.Recordset.Fields!payed = 1
Adodc1.Recordset.Fields!user_name = main.txtusername
Adodc1.Recordset.Fields!box_man = main.txtusername
Adodc1.Recordset.Fields!bill_no = Text5.Text
Adodc1.Recordset.Update
 

 
MsgBox " „ «· ‰›ÌÌ–", vbInformation
 
Text3.Text = 0
 
Text5.Text = ""
Text8.Text = ""

Command1.Enabled = False
Command2.Enabled = True
 
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
Command1.Enabled = True
   
  
End Sub

Private Sub Command3_Click()
MASROUF_TYPE.Show
End Sub
 


Private Sub Text3_Change()
If Not IsNumeric(Text3.Text) Then
 
MsgBox "ÌÊÃœ Œÿ√ ›Ì «·ﬁÌ„ «·„œŒ·… ÌÃ» «‰  ﬂÊ‰ «—ﬁ«„", vbCritical
Text3.Text = ""
End If
End Sub
 
