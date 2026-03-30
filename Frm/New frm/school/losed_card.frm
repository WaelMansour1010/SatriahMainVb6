VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form losed_card 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… »œ· ›«∆œ"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "year"
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "OPERATION_NO"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   4320
      Width           =   150
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   3960
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÃœÌœ"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "losed_card_value"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   372
      Left            =   4320
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "»ÕÀ «·ÿ«·»   "
      Enabled         =   0   'False
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   732
   End
   Begin VB.CommandButton Command4 
      Caption         =   "»ÕÀ ÿ«·»  «»⁄"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3840
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   0
      Top             =   4320
      Width           =   6972
      _ExtentX        =   12303
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   0
      Top             =   5040
      Width           =   6975
      _ExtentX        =   12303
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
      RecordSource    =   "losed_card"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   -360
      Top             =   2160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "this_year"
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "”‰… «· ÃœÌœ"
      Height          =   375
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·ÿ«·»  "
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
      Left            =   6600
      TabIndex        =   12
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»  "
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
      Left            =   6600
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "‘«‘… »œ· ›«∆œ"
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
      Height          =   612
      Left            =   1800
      TabIndex        =   10
      Top             =   0
      Width           =   3972
   End
   Begin VB.Label Label4 
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
      Height          =   615
      Left            =   6600
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label7 
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
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "losed_card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHECKN As Integer
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
CHECKN = 1
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command2_Click()


If Text2.Text <> "" And CHECKN = 1 Then
CHECKN = 0
Adodc1.Recordset.Fields!member_id = Text2.Text
'Adodc1.Recordset.Fields!CHILD_ID = Text5.Text
Adodc1.Recordset.Fields!MEMBER_NAME = Text3.Text

Adodc1.Recordset.Fields!operation_type = "»œ· ›«∆œ"
Adodc1.Recordset.Fields!OPERATION_DATE = DateValue(Now)
Adodc1.Recordset.Fields!user_name = main.txtusername
Adodc1.Recordset.Fields!ID_VALUE = Text1.Text
Adodc1.Recordset.Fields!TOTAL_VALUE = Text1.Text
Adodc1.Recordset.Fields!ACTUAL_VALUE = Text1.Text

Adodc1.Recordset.Fields!update_year = Text7.Text
Adodc1.Recordset.Fields!member_type = Text4.Text
Adodc1.Recordset.Update
MsgBox " „ «·Õ›Ÿ", vbInformation
Command3.Enabled = False
Command4.Enabled = False
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

Else
MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
End If


End Sub

Private Sub Command3_Click()
member_search.Show
member_search.from = 8

End Sub

Private Sub Command4_Click()
MEMBER_CHILD_SEARCH.Show
MEMBER_CHILD_SEARCH.from = 4
End Sub



Private Sub Form_Load()
 

CHECKN = 0
End Sub


