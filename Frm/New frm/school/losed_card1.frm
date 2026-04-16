VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form reprint_card 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÔÇÔÉ ÇÚÇÏÉ ÇáØÈÇÚå"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   8085
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "year"
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   " "
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
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
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "ÇÚÇÏÉ ØÈÇÚÉ ÇáßÇÑäíÉ"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÈÍË   "
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2400
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      RecordSource    =   "select * from ready_to_print where MEMBER_ID='0'"
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
      Left            =   2880
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
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
      Alignment       =   2  'Center
      Caption         =   "ÇáÓäÉ ÇáÍÇáíÉ"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ÑÞã ÇáØÇáÈ"
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
      Left            =   6600
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ÇÓã ÇáØÇáÈ"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "ÔÇÔÉ ÇÚÇÏÉ ÇáØÈÇÚÉ"
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
      TabIndex        =   3
      Top             =   0
      Width           =   3972
   End
End
Attribute VB_Name = "reprint_card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 

Private Sub Command2_Click()
If Adodc1.Recordset.RecordCount > 0 Then
member_id = Adodc1.Recordset.Fields!member_id
MEMBER_NAME = Adodc1.Recordset.Fields!MEMBER_NAME
member_type = Adodc1.Recordset.Fields!member_type
update_year = Adodc1.Recordset.Fields!update_year
sex = Adodc1.Recordset.Fields!sex
OPR_TYPE = "ÇÚÇÏÉ ØÈÇÚÉ "
IMAGE_PATH = Adodc1.Recordset.Fields!IMAGE_PATH
CENTER_MANAGER = Adodc1.Recordset.Fields!CENTER_MANAGER
user_name = main.txtusername.Caption
OPERATION_DATE = Adodc1.Recordset.Fields!OPERATION_DATE
type1 = Adodc1.Recordset.Fields!Type
 
 
 Adodc1.Recordset.AddNew
 Adodc1.Recordset.Fields!member_id = member_id
Adodc1.Recordset.Fields!MEMBER_NAME = MEMBER_NAME
Adodc1.Recordset.Fields!member_type = member_type
 Adodc1.Recordset.Fields!update_year = update_year
Adodc1.Recordset.Fields!sex = sex
Adodc1.Recordset.Fields!OPR_TYPE = OPR_TYPE
 Adodc1.Recordset.Fields!IMAGE_PATH = IMAGE_PATH
Adodc1.Recordset.Fields!CENTER_MANAGER = CENTER_MANAGER
 Adodc1.Recordset.Fields!user_name = user_name
 Adodc1.Recordset.Fields!OPERATION_DATE = OPERATION_DATE
Adodc1.Recordset.Fields!Type = type1
 
 
 Adodc1.Recordset.Update
MsgBox "Êã"

End If
End Sub

Private Sub Command3_Click()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from ready_to_print where MEMBER_ID='" & Text2.Text & "' and update_year='" & Text1.Text & "'"
Adodc1.Refresh





End Sub

Private Sub Command4_Click()

End Sub



Private Sub Text2_Change()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from ready_to_print where MEMBER_ID='" & Text2.Text & "' and update_year='" & Text1.Text & "'"
Adodc1.Refresh
End Sub
