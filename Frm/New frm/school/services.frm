VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form services 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‘«‘… «· Õ’Ì·« "
   ClientHeight    =   4620
   ClientLeft      =   7110
   ClientTop       =   3495
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   7275
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
   Begin VB.CommandButton Command4 
      Caption         =   "«÷«›… ‰Ê⁄ ÃœÌœ"
      Height          =   255
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   3120
      TabIndex        =   17
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "ÃœÌœ"
      Height          =   612
      Left            =   360
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   2172
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "FINES_TOTAL"
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3480
      Width           =   2295
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "services.frx":0000
      DataField       =   "OPERATION_TYPE"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3000
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "name"
      Text            =   ""
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "Fines_PERCENT"
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
      Left            =   3120
      TabIndex        =   9
      Text            =   "1"
      Top             =   2880
      Width           =   2295
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "»ÕÀ"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1800
      Picture         =   "services.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "”œ«œ"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2400
      Top             =   4680
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
      Left            =   7320
      Top             =   1920
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   7320
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
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
   Begin MSAdodcLib.Adodc Adodc20 
      Height          =   375
      Left            =   0
      Top             =   0
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
      Left            =   5400
      TabIndex        =   18
      Top             =   3960
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
      Left            =   5520
      TabIndex        =   15
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "«·«Ã„«·Ì"
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
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«·⁄œœ"
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
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «· Õ’Ì·"
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
      TabIndex        =   8
      Top             =   1800
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
      Left            =   5400
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·ÿ«·»"
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
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»"
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
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHECKFRM As Integer
Dim VALUE_OF_MEMBER As Single
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








Adodc1.Recordset.Fields!member_id = Text1.Text
Adodc1.Recordset.Fields!MEMBER_NAME = Text6.Text
Adodc1.Recordset.Fields!operation_type = DataCombo1.Text
Adodc1.Recordset.Fields!OTHERS_FEES = Text3.Text
Adodc1.Recordset.Fields!TOTAL_VALUE = Text2.Text
Adodc1.Recordset.Fields!ACTUAL_VALUE = Text2.Text
Adodc1.Recordset.Fields!OPERATION_DATE = DateValue(Now)
Adodc1.Recordset.Fields!payed = 1
Adodc1.Recordset.Fields!user_name = main.txtusername
Adodc1.Recordset.Fields!box_man = main.txtusername
Adodc1.Recordset.Fields!bill_no = Text5.Text

Adodc1.Recordset.Update
 

 
MsgBox " „ «· ‰›ÌÌ–", vbInformation
Text1.Text = ""
Text2.Text = 0
Text3.Text = 0
Text4.Text = 0
Text5.Text = ""
Text6.Text = ""

Command1.Enabled = False
Command2.Enabled = True
 
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
Command1.Enabled = True
  Command3.Enabled = True
  Text4.Text = 1
End Sub

Private Sub Command3_Click()
member_search.Show
member_search.from = 9
End Sub
 


Private Sub Command4_Click()
SERVICE_TYPE.Show
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "SELECT  * FROM   services WHERE name LIKE'%" & DataCombo1.Text & "%'"
Adodc2.Refresh

Text3.Text = Adodc2.Recordset.Fields!Value
 
End Sub

Private Sub Text3_Change()
If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) Then
Text2.Text = Text3.Text * Text4.Text
Else
MsgBox "ÌÊÃœ Œÿ√ ›Ì «·ﬁÌ„ «·„œŒ·… ÌÃ» «‰  ﬂÊ‰ «—ﬁ«„", vbCritical
Text3.Text = ""
End If
End Sub

Private Sub Text4_Change()
If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) Then
Text2.Text = Text3.Text * Text4.Text
Else
MsgBox "ÌÊÃœ Œÿ√ ›Ì «·ﬁÌ„ «·„œŒ·… ÌÃ» «‰  ﬂÊ‰ «—ﬁ«„", vbCritical
End If
End Sub
