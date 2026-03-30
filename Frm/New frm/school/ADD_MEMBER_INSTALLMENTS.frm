VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form ADD_MEMBER_INSTALLMENTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "     «÷«›… ﬁ”ÿ ⁄·Ï  «·ÿ«·»  "
   ClientHeight    =   5535
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   7185
   Begin VB.CommandButton Command1 
      Caption         =   "  ÿ»Ìﬁ ﬁ”ÿ ⁄·Ï «·ÿ«·» "
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   855
      Left            =   2040
      Picture         =   "ADD_MEMBER_INSTALLMENTS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      Height          =   492
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "Installments_VALUE"
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
      TabIndex        =   6
      Text            =   "0"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "INSTALLMENT_COUNT"
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
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "INSATLLMENT_DATE"
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
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "ÃœÌœ"
      Height          =   672
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   2292
   End
   Begin VB.TextBox Text9 
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc5"
      Height          =   615
      Left            =   8280
      TabIndex        =   1
      Text            =   "Text9"
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "  ÿ»Ìﬁ ﬁ”ÿ ⁄·Ï  «·“ÊÃ…"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "ADD_MEMBER_INSTALLMENTS.frx":1992
      DataField       =   "Installments_TYPE"
      DataSource      =   "Adodc3"
      Height          =   312
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Width           =   2292
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Installments_NAME"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   7440
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
      Height          =   336
      Left            =   7200
      Top             =   1680
      Visible         =   0   'False
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   582
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
      Left            =   1080
      Top             =   6360
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   336
      Left            =   480
      Top             =   2160
      Visible         =   0   'False
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   582
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
      Height          =   732
      Left            =   1440
      Top             =   5520
      Width           =   6972
      _ExtentX        =   12303
      _ExtentY        =   1296
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   336
      Left            =   720
      Top             =   2640
      Visible         =   0   'False
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   582
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
   Begin VB.Label « 
      Caption         =   "«”„ «·ÿ«·»  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   5640
      TabIndex        =   16
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·ÿ«·»  "
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
      TabIndex        =   15
      Top             =   360
      Width           =   3015
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
      Height          =   612
      Left            =   5400
      TabIndex        =   14
      Top             =   2160
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·ﬁ”ÿ"
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
      Left            =   5400
      TabIndex        =   13
      Top             =   1680
      Width           =   1812
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "⁄œœ «·œ›⁄« "
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
      Left            =   5400
      TabIndex        =   12
      Top             =   2760
      Width           =   1812
   End
   Begin VB.Label Label8 
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
      Height          =   612
      Left            =   5520
      TabIndex        =   11
      Top             =   3360
      Width           =   1812
   End
End
Attribute VB_Name = "ADD_MEMBER_INSTALLMENTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHECKFRM As Integer
Dim VALUE_OF_MEMBER As Single

Private Sub Command1_Click()

    Dim i As Integer
    Adodc3.Recordset.update

    If Text5.text <> "" Then
 
        For i = 1 To Text5.text
            Adodc5.Recordset.AddNew
            Adodc5.Recordset.Fields!INSTALLMENT_NO = i
            Adodc5.Recordset.Fields!member_id = Text1.text

            Adodc5.Recordset.Fields!installment_value = val(Text3.text) / val(Text5.text)
            '   If i = 1 Then
            '   Adodc5.Recordset.Fields!ACTIVATED = 1
            '   End If
         
            Adodc5.Recordset.update
        Next i

    Else
        MsgBox "·«»œ „‰ ﬂ «»… ⁄œœ «·œ›⁄« ", vbCritical
        Exit Sub
    End If

    MsgBox " „ «· ‰›ÌÌ–", vbInformation
    Command1.Enabled = False
    Command3.Enabled = False
 
    Command5.Enabled = False
End Sub

Private Sub Command2_Click()
    Adodc3.Recordset.AddNew
    Text7.text = DateValue(Now)
    Command1.Enabled = True
    Command3.Enabled = True
 
    Command5.Enabled = True
End Sub

Private Sub Command3_Click()
    member_search.Show
    member_search.from = 2
End Sub

Private Sub Command5_Click()
    Dim i As Integer
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "select * from member_child where member_id= " & Text1.text & " and MEMBER_CHILD_ID=1"
    Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then

        Adodc3.Recordset.Fields!CHILD_ID = 1
        Adodc3.Recordset.update

        If Text5.text <> "" Then
 
            For i = 1 To Text5.text
                Adodc5.Recordset.AddNew
                Adodc5.Recordset.Fields!INSTALLMENT_NO = i
                Adodc5.Recordset.Fields!member_id = Text1.text

                Adodc5.Recordset.Fields!installment_value = val(Text3.text) / val(Text5.text)
                Adodc5.Recordset.Fields!CHILD_ID = 1
                '   If i = 1 Then
                '   Adodc5.Recordset.Fields!ACTIVATED = 1
                '   End If
         
                Adodc5.Recordset.update
            Next i

        Else
            MsgBox "·«»œ „‰ ﬂ «»… ⁄œœ «·œ›⁄« ", vbCritical
            Exit Sub
        End If

        MsgBox " „ «· ‰›ÌÌ–", vbInformation
        Command1.Enabled = False
        Command3.Enabled = False
 
        Command5.Enabled = False
    Else
        MsgBox "·« ÌÊÃœ “ÊÃ… ·Â–« «·⁄÷Ê", vbCritical

    End If

End Sub

Private Sub DataCombo1_Click(Area As Integer)
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT  * FROM Installments_TYPES WHERE Installments_NAME LIKE'%" & DataCombo1.text & "%'"
    Adodc2.Refresh

    Text3.text = Adodc2.Recordset.Fields!Installments_VALUE
    Text5.text = Adodc2.Recordset.Fields!Installments_COUNT
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM MEMBERS where MEMBER_ID =0"
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT  * FROM Fine_TYPES WHERE Fines_ID=0  "
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "SELECT * FROM  INSTALLMENTS WHERE MEMBER_ID=0 "
    Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select * from  Installments_TYPES "
    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from  INSTALLMENT_DETAILS "
    Adodc5.Refresh

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    Adodc6.RecordSource = "select * from Fine_TYPES "
    Adodc6.Refresh

    'Adodc7.ConnectionString = connection_string
    'Adodc7.CommandType = adCmdText
    'Adodc7.RecordSource = "select * from  "
    'Adodc7.Refresh
    VALUE_OF_MEMBER = 0

End Sub

