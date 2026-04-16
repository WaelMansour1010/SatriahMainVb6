VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form renew_member_activity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÃœÌœ «·«‘ —«ﬂ ›Ì ‰‘«ÿ"
   ClientHeight    =   4365
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   7305
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "member_type"
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
      Height          =   468
      Left            =   720
      TabIndex        =   15
      Top             =   2160
      Width           =   4692
   End
   Begin VB.TextBox Text6 
      DataField       =   "INSTALLMENTS_TOTAL"
      DataSource      =   "Adodc5"
      Height          =   288
      Left            =   4440
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   5400
      Width           =   492
   End
   Begin VB.CommandButton Command4 
      Caption         =   "»ÕÀ ÿ«·»  «»⁄"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1560
      TabIndex        =   12
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton Command3 
      Caption         =   "»ÕÀ  ÿ«·»"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   840
      Width           =   732
   End
   Begin VB.TextBox Text1 
      DataField       =   "VALUE"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   372
      Left            =   600
      TabIndex        =   10
      Top             =   3240
      Width           =   4812
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "renew_member_activity.frx":0000
      DataField       =   "Activities_NAME"
      DataSource      =   "Adodc1"
      Height          =   288
      Left            =   600
      TabIndex        =   8
      Top             =   2760
      Width           =   4812
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Activities_NAME"
      Text            =   ""
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÃœÌœ"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "member_name"
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
      Height          =   492
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   4692
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      DataField       =   "member_id_FULL"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Width           =   1932
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   0
      Top             =   4440
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   492
      Left            =   240
      Top             =   5160
      Width           =   6972
      _ExtentX        =   12303
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   492
      Left            =   240
      Top             =   4920
      Width           =   6972
      _ExtentX        =   12303
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   372
      Left            =   0
      Top             =   5760
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
   Begin VB.Label Label70 
      Caption         =   "0"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label60 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   375
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
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   5640
      TabIndex        =   14
      Top             =   2160
      Width           =   1812
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
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   5640
      TabIndex        =   9
      Top             =   3240
      Width           =   1812
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " ÃœÌœ «·«‘ —«ﬂ ›Ì ‰‘«ÿ  "
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
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«·‰‘«ÿ"
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
      Height          =   612
      Left            =   5640
      TabIndex        =   4
      Top             =   2760
      Width           =   1812
   End
   Begin VB.Label Label2 
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
      Height          =   612
      Left            =   5640
      TabIndex        =   3
      Top             =   1680
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·ÿ«·»"
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
      Height          =   612
      Left            =   5640
      TabIndex        =   2
      Top             =   960
      Width           =   1812
   End
End
Attribute VB_Name = "renew_member_activity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHECKN As Integer

Private Sub Command1_Click()
    Adodc1.Recordset.AddNew
    Label60.Caption = 0
    Label70.Caption = 0
    CHECKN = 1
    Command3.Enabled = True
    Command4.Enabled = True
End Sub

Private Sub Command2_Click()

    If DataCombo1.text = "" Then
        MsgBox "·«»œ „‰ «Œ Ì«— «·‰‘«ÿ ﬁ»· «·Õ›Ÿ", vbCritical
        DataCombo1.BackColor = &HFF&
        Exit Sub

    End If

    Adodc1.Recordset.Fields!member_id = Label60.Caption
    Adodc1.Recordset.Fields!MEMBER_CHILD_ID = Label70.Caption
    Adodc1.Recordset.Fields!Activities_NAME = DataCombo1.text
    Adodc1.Recordset.Fields![value] = Text1.text

    Adodc1.Recordset.update

    If CHECKN = 1 Then
        Adodc5.Recordset.AddNew
        Adodc5.Recordset.Fields!member_id = Label60.Caption

        'Adodc5.Recordset.Fields!member_id = Text2.Text
        Adodc5.Recordset.Fields!CHILD_ID = Label70.Caption
        Adodc5.Recordset.Fields!member_name = Text3.text
        Adodc5.Recordset.Fields!activity_value = Text1.text
        Adodc5.Recordset.Fields!MEMBER_TYPE = Text4.text
        Adodc5.Recordset.Fields!total_value = Text1.text
        Adodc5.Recordset.Fields!activity_name = DataCombo1.text
        Adodc5.Recordset.Fields!OPERATION_DATE = DateValue(Now)
        'Adodc5.Recordset.Fields!User_Name = Main.TxtUserName.Caption
        Adodc5.Recordset.Fields!operation_type = " ÃœÌœ ‰‘«ÿ"
        Adodc5.Recordset.update
    End If

    MsgBox " „ «·Õ›Ÿ", vbInformation
    Command3.Enabled = False
    Command4.Enabled = False

End Sub

Private Sub Command3_Click()
    member_search.Show
    member_search.from = 10

End Sub

Private Sub Command4_Click()
    MEMBER_CHILD_SEARCH.Show
    MEMBER_CHILD_SEARCH.from = 5
End Sub

Private Sub DataCombo1_Click(Area As Integer)

    If DataCombo1.text = "" Then Exit Sub
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from ACTIVITIES_type where Activities_NAME='" & DataCombo1.text & "'"
    Adodc3.Refresh

    Text1.text = Adodc3.Recordset.Fields![value]
End Sub

Private Sub Form_Load()
    CHECKN = 0
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from member_activity where member_id=0 "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from  ACTIVITIES_type "
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  ACTIVITIES_type "
    Adodc3.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from  OPERATIONS "
    Adodc5.Refresh

End Sub

