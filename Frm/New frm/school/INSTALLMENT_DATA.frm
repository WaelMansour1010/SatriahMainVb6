VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form INSTALLMENT_DATA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„⁄·Ê„«  «·«ﬁ”«ÿ"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12195
   Icon            =   "INSTALLMENT_DATA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   12195
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   5640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ›·Ÿ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "INSTALLMENT_DATA.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check1 
      Caption         =   " ›€Ì· «·ﬁ”ÿ"
      DataField       =   "ACTIVATED"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   10680
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtname 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox id 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1200
      Top             =   6600
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
      Height          =   495
      Left            =   7200
      Top             =   7080
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "INSTALLMENT_DATA.frx":0028
      Height          =   3375
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "op_no"
         Caption         =   "op_no"
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
         DataField       =   "cust_id"
         Caption         =   "cust_id"
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
      BeginProperty Column02 
         DataField       =   "MEMBER_ID"
         Caption         =   "MEMBER_ID"
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
      BeginProperty Column03 
         DataField       =   "member_name"
         Caption         =   "member_name"
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
      BeginProperty Column04 
         DataField       =   "CHILD_ID"
         Caption         =   "ﬂÊœ «·ÿ«·»"
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
      BeginProperty Column05 
         DataField       =   "member_child_name"
         Caption         =   "«”„ «·ÿ«·»"
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
      BeginProperty Column06 
         DataField       =   "INSTALLMENT_NO"
         Caption         =   "—ﬁ„ «·ﬁ”ÿ"
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
      BeginProperty Column07 
         DataField       =   "INSTALLMENT_VALUE"
         Caption         =   "ﬁÌ„… «·ﬁ”ÿ"
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
      BeginProperty Column08 
         DataField       =   "ACTIVATED"
         Caption         =   "ACTIVATED"
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
      BeginProperty Column09 
         DataField       =   "PAYED"
         Caption         =   "„”œœ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "‰⁄„"
            FalseValue      =   "·«"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "DATE_OF_PAYED"
         Caption         =   " «—ÌŒ «·”œ«œ"
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
      BeginProperty Column11 
         DataField       =   "returned"
         Caption         =   " „ «” ⁄«œ …"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "‰⁄„"
            FalseValue      =   "·«"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "date_of_returned"
         Caption         =   " «—ÌŒ «·«” ⁄«œ…"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Label lblcustid 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label d3 
      Alignment       =   2  'Center
      Caption         =   "00"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   5640
      Width           =   3975
   End
   Begin VB.Label d2 
      Alignment       =   2  'Center
      Caption         =   "00"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label d1 
      Alignment       =   2  'Center
      Caption         =   "00"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì «·„ »ﬁÌ"
      Height          =   375
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì «·„”œœ"
      Height          =   375
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì «·«ﬁ”«ÿ"
      Height          =   375
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "«”„  Ê·Ì «·«„—"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "ﬂÊœ Ê·Ì «·«„—"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "INSTALLMENT_DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHECKFRM As Integer

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub ALLButton1_Click()

    Adodc1.Recordset.update
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select INSTALLMENT_NO,INSTALLMENT_VALUE,DATE_OF_PAYED,ACTIVATED  FROM INSTALLMENT_DETAILS where MEMBER_ID=0 "
    Adodc1.Refresh

End Sub

Private Sub Text1_Change()

End Sub

Function calc()
    On Error Resume Next
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select sum(INSTALLMENT_VALUE) as total FROM INSTALLMENT_DETAILS where cust_id=" & Me.lblcustid.Caption
    Adodc2.Refresh
    d1.Caption = Adodc2.Recordset.Fields!total
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select sum(INSTALLMENT_VALUE) as total_payed FROM INSTALLMENT_DETAILS where cust_id=" & Me.lblcustid.Caption & " and payed=1"
    Adodc2.Refresh
    d2.Caption = Adodc2.Recordset.Fields!total_payed
    d3.Caption = d1.Caption - d2.Caption
End Function

Private Sub lblcustid_Change()
    calc
End Sub

