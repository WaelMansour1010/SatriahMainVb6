VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form MEMBER_CHILD_SEARCH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "            ÔÇÔÉ ÈÍË ÇáØáÇÈ ÇáÊÇÈÚíä"
   ClientHeight    =   6495
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   6960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãæÇÝÞ"
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MEMBER_CHILD_SEARCH.frx":0000
      Height          =   4452
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   6612
      _ExtentX        =   11668
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   612
      Left            =   480
      Top             =   7080
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   1085
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
   Begin VB.Label Label6 
      Caption         =   "ÇÓã ÇáØÇáÈ ÇáÊÇÈÚ"
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   4680
      TabIndex        =   9
      Top             =   720
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "ÇáÕÝÉ"
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   5520
      TabIndex        =   7
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "ÑÞã ÇáØÇáÈ"
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label from 
      Caption         =   "Label3"
      Height          =   372
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Label Label3 
      Caption         =   "ÇÓã ÇáØÇáÈ"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "ÑÞã ÇáØÇáÈ"
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   1692
   End
End
Attribute VB_Name = "MEMBER_CHILD_SEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If from = 1 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        member_activity.Text2.text = Adodc1.Recordset.Fields!member_id & "-" & Adodc1.Recordset.Fields!MEMBER_CHILD_ID
        member_activity.Label60 = Adodc1.Recordset.Fields!member_id
        'member_activity.Label70 = Adodc1.Recordset.Fields!MEMBER_CHILD_ID
        member_activity.Text3.text = Adodc1.Recordset.Fields!MEMBER_CHILD_NAME
        member_activity.Text4.text = Adodc1.Recordset.Fields!MEMBER_TYPE

    End If

    If from.Caption = 3 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        delete_member_activity.Text1.text = Adodc1.Recordset.Fields!member_id & "-" & Adodc1.Recordset.Fields!MEMBER_CHILD_ID
        delete_member_activity.Text2.text = Adodc1.Recordset.Fields!MEMBER_CHILD_NAME
        'delete_member_activity.Text3.text = Adodc1.Recordset.Fields!member_type_name
    End If

    If from.Caption = 4 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        losed_card.Text2.text = Adodc1.Recordset.Fields!member_id
        'losed_card.Text5.Text = Adodc1.Recordset.Fields!MEMBER_CHILD_ID
        losed_card.Text5.Visible = True
        losed_card.Text3.text = Adodc1.Recordset.Fields!MEMBER_CHILD_NAME
        losed_card.Text4.text = Adodc1.Recordset.Fields!MEMBER_TYPE
    End If

    If from = 5 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        renew_member_activity.Text2.text = Adodc1.Recordset.Fields!member_id & "-" & Adodc1.Recordset.Fields!MEMBER_CHILD_ID
        renew_member_activity.Label60 = Adodc1.Recordset.Fields!member_id
        'renew_member_activity.Label70 = Adodc1.Recordset.Fields!MEMBER_CHILD_ID
        renew_member_activity.Text3.text = Adodc1.Recordset.Fields!MEMBER_CHILD_NAME
        renew_member_activity.Text4.text = Adodc1.Recordset.Fields!MEMBER_TYPE
    End If

    Unload Me
End Sub

Private Sub DataGrid1_Click()
    Command1_Click
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from  MEMBER_CHILD"
    Adodc1.Refresh
End Sub

Private Sub Text1_Change()

    If Text1.text <> "" Or IsNumeric(Text1.text) Then
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * FROM member_CHILD where MEMBER_ID=" & Text1.text
        Adodc1.Refresh
    End If

End Sub

Private Sub Text3_Change()
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * FROM member_CHILD where MEMBER_CHILD_NAME LIKE'%" & Text3.text & "%'"
    Adodc1.Refresh

End Sub
