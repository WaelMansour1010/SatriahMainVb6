VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form pay_losed_ard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "œ›⁄ —”Ê„ »œ· ›«∆œ"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10830
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      DataField       =   "USER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc4"
      Height          =   285
      Left            =   9720
      TabIndex        =   21
      Text            =   "Text11"
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "update_year"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      DataField       =   "CENTER_MANAGER"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      DataField       =   "ID_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_DATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_NO"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_TYPE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "member_type"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "”œ«œ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "bill_no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   2880
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   480
      Top             =   2880
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   767
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
      RecordSource    =   "select  *  from operations where PAYED=0  and operation_type='»œ· ›«∆œ'"
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   435
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   767
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
      RecordSource    =   "ready_to_print"
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   435
      Left            =   600
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   767
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
      RecordSource    =   "ready_to_print"
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   435
      Left            =   5880
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
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
      RecordSource    =   "CENTER_MANAGER"
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   435
      Left            =   3960
      Top             =   3480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
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
      RecordSource    =   "CENTER_MANAGER"
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "»Ê«”ÿ…"
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
      Left            =   8520
      TabIndex        =   23
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "«·⁄«„ «·œ—«”Ì"
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
      Left            =   2880
      TabIndex        =   20
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·ﬂ«—‰Ì… »œ·  ›«∆œ"
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
      Left            =   2880
      TabIndex        =   16
      Top             =   1200
      Width           =   2535
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
      Left            =   8160
      TabIndex        =   15
      Top             =   1680
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
      Left            =   7800
      TabIndex        =   14
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " «—ÌŒ «·⁄„·Ì…"
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
      Left            =   3120
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·⁄·„Ì…"
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
      Left            =   8040
      TabIndex        =   12
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
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
      Left            =   8880
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«·”‰… «·œ—«”Ì…"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Left            =   8040
      TabIndex        =   9
      Top             =   2760
      Width           =   3255
   End
End
Attribute VB_Name = "pay_losed_ard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text9.Text = "" Then
MsgBox "  ·«»œ „‰ ﬂ «»… —ﬁ„ «·«Ì’«· «Ê·«  ", vbCritical
Text9.BackColor = &HFF&
Text9.SetFocus
Exit Sub
End If

Adodc20.CommandType = adCmdText
Adodc20.RecordSource = "select * from operations where bill_no=" & Text9.Text
Adodc20.Refresh

If Adodc20.Recordset.RecordCount > 0 Then
MsgBox "—ﬁ„ «·«Ì’«·  „” Œœ„ „‰ ﬁ»·", vbCritical
Exit Sub
End If




  

If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "·« ÌÊÃœ «Ì ⁄„·Ì«  ·”œ«œÂ«", vbCritical
Exit Sub
End If


    Adodc4.Recordset.AddNew
    
        
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "SELECT * FROM OPERATIONS  WHERE OPERATION_TYPE= ' ÃœÌœ ⁄÷ÊÌ…' and  MEMBER_ID=" & Text1.Text
    Adodc3.Refresh
     
    
   If Text7.Text = "0" Or Text7.Text = "" Then
     Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "SELECT * FROM MEMBERS  WHERE MEMBER_ID=" & Text1.Text
    Adodc5.Refresh
 
   If Adodc3.Recordset.RecordCount = 0 Then
     Adodc4.Recordset.Fields!Type = "⁄÷ÊÌ… ÃœÌœ…"
   Else
   Adodc4.Recordset.Fields!Type = " ÃœÌœ ⁄÷ÊÌ…"
   End If
    
    Adodc4.Recordset.Fields!sex = Adodc5.Recordset.Fields!sex
   Adodc4.Recordset.Fields!IMAGE_PATH = Adodc5.Recordset.Fields!IMAGE_PATH
   Adodc4.Recordset.Fields!user_name = main.txtusername
    Adodc4.Recordset.Fields!bill_no = Text9.Text
    Adodc4.Recordset.Fields!CENTER_MANAGER = Text8.Text
    Adodc4.Recordset.Fields!member_id = Text1.Text
    Adodc4.Recordset.Fields!MEMBER_NAME = Text6.Text
    Adodc4.Recordset.Fields!member_type = Text5.Text
    Adodc4.Recordset.Fields!update_year = Text10.Text
    Adodc4.Recordset.Fields!OPERATION_DATE = DateValue(Now)
     Adodc4.Recordset.Fields!OPR_TYPE = "»œ· ›«∆œ"
       
    
    Adodc4.Recordset.Update
    Else
     
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "SELECT * FROM MEMBER_CHILD  WHERE MEMBER_ID=" & Text1.Text & " and MEMBER_CHILD_ID=" & Text7.Text
    Adodc5.Refresh
    
    Adodc4.Recordset.Fields!sex = Adodc5.Recordset.Fields!sex
   
      If Adodc3.Recordset.RecordCount = 0 Then
     Adodc4.Recordset.Fields!Type = "⁄÷ÊÌ… ÃœÌœ…"
     Else
     Adodc4.Recordset.Fields!Type = " ÃœÌœ ⁄÷ÊÌ…"
     End If
   
    
   Adodc4.Recordset.Fields!user_name = main.txtusername
    Adodc4.Recordset.Fields!IMAGE_PATH = Adodc5.Recordset.Fields!MEMBER_CHILD_iMAGE_PATH
   Adodc4.Recordset.Fields!bill_no = Text9.Text
   Adodc4.Recordset.Fields!CENTER_MANAGER = Text8.Text
    Adodc4.Recordset.Fields!member_id = Text1.Text & "-" & Text7.Text
    Adodc4.Recordset.Fields!MEMBER_NAME = Adodc5.Recordset.Fields!MEMBER_CHILD_NAME
    If Text7.Text = "1" Then
    Adodc4.Recordset.Fields!member_type = "⁄÷Ê ⁄«„·  «»⁄"
    Else
     Adodc4.Recordset.Fields!member_type = "⁄÷Ê  «»⁄"
     End If
    Adodc4.Recordset.Fields!update_year = Text10.Text
     Adodc4.Recordset.Fields!OPERATION_DATE = DateValue(Now)
    Adodc4.Recordset.Fields!OPR_TYPE = "»œ· ›«∆œ"
    Adodc4.Recordset.Update
    
    
    End If
Adodc1.Recordset.Fields!payed = 1
Adodc1.Recordset.Fields!ACTUAL_VALUE = Text18.Text
Adodc1.Recordset.Update
Adodc1.Refresh


    MsgBox " „ «·Õ›Ÿ", vbInformation
    Text9.Text = ""

End Sub

Private Sub Text7_Change()
If Text7.Text = "0" Then
Text7.Visible = False
Else
Text7.Visible = True
End If
End Sub

Private Sub Text9_Change()
If Text9.Text = "" Then Exit Sub
If Not IsNumeric(Text9.Text) Then
MsgBox "  —ﬁ„ «·«Ì’«· „ﬂÊ‰ „‰ «—ﬁ«„ ›ﬁÿ    ", vbCritical
Text9.Text = ""
Exit Sub
End If

End Sub

