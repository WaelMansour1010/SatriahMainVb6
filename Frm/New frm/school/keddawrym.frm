VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form keddawrym 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÇÓÊÏÚÇÁ  ÞíÏ ÏæÑí íÏæí"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
   Icon            =   "keddawrym.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   10335
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "des"
      DataSource      =   "Adodc1"
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3120
      Width           =   10335
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "ÇäÔÇÁ"
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
      MICON           =   "keddawrym.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "keddawrym.frx":0028
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ã"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ked_serial"
         Caption         =   "  ÑÞã ÇáÞíÏ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "des"
         Caption         =   "ÇáÔÑÍ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ked_date"
         Caption         =   "ÇáÊÇÑíÎ ÇáãÞÊÑÍ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "SallingPrice"
         Caption         =   "ÚÏÏ ÇáÞíæÏ ÇáãÊÈÞÉ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "akher_s3r_shera"
         Caption         =   "akher_s3r_shera"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "motwaset_taklefa"
         Caption         =   "motwaset_taklefa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "blocked"
         Caption         =   "blocked"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "akher_shera_date"
         Caption         =   "akher_shera_date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "akher_be3_date"
         Caption         =   "akher_be3_date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "akher_sarf_date"
         Caption         =   "akher_sarf_date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "hesab_taklefa_method"
         Caption         =   "hesab_taklefa_method"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   5999.812
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÇÎÊíÇÑ Çáßá"
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
      MICON           =   "keddawrym.frx":003D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÍÐÝ Çáßá"
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
      MICON           =   "keddawrym.frx":0059
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   330
      Left            =   6240
      TabIndex        =   7
      Top             =   840
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   582
      _Version        =   393216
      Format          =   94240769
      CurrentDate     =   38784
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1032
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
      Caption         =   "ÊÍÑíß"
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
   Begin VB.Label LBLTYPE 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label LBLCOUNT 
      Alignment       =   1  'Right Justify
      Caption         =   " "
      Height          =   375
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇÌãÇáí ÇáÞíæÏ ÇáãÊÈÞíÉ"
      Height          =   375
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ÊÇÑíÎ ÇáÚãáíÉ"
      Height          =   375
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "    ÇÓÊÏÚÇÁ ÞíÏ ÏæÑí íÏæí"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   9960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "    ÇäÔÇÁ ÞíÏ ÏæÑí íÏæí"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   10320
   End
End
Attribute VB_Name = "keddawrym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ALLButton1_Click()
    On Error Resume Next
If LBLTYPE = 0 Then
    If Adodc1.Recordset.RecordCount > 0 And Not IsNull(Adodc1.Recordset.Fields!ked_no) Then
        FrmAccEditJournal.retrive1 (Adodc1.Recordset.Fields!ked_no)
        FrmAccEditJournal.Txt.Text = Adodc1.Recordset.Fields!des
        FrmAccEditJournal.DTP_Date.value = Me.XPDtbBill.value
        Adodc1.Recordset.Fields!OK = 1
        Adodc1.Recordset.update
    End If
ElseIf LBLTYPE = 1 Then

   FrmAccEditJournal4.retrive1 (Adodc1.Recordset.Fields!ked_no)
        FrmAccEditJournal4.Txt.Text = Adodc1.Recordset.Fields!des
        FrmAccEditJournal4.DTP_Date.value = Me.XPDtbBill.value
        Adodc1.Recordset.Fields!OK = 1
        Adodc1.Recordset.update
End If

    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    XPDtbBill.value = Now

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    XPDtbBill.value = Date
    Dim sql As String
    connection_string = Cn.ConnectionString
    'Adodc1.ConnectionString = connection_string
    'Adodc1.CommandType = adCmdText

    'sql = "SELECT     ked_no, ked_serial, des, ked_date, [index],OK"
    'sql = sql + " from dbo.ked_dawry WHERE     ( OK=0 AND ked_date = CONVERT(DATETIME, '" & Format$(XPDtbBill.value, "dd-mm-yyyy") & " 00:00:00', 102))"
'
'    Adodc1.RecordSource = sql ' "   SELECT  * FROM KED_DAWRY WHERE KALEB=1"
'    Adodc1.Refresh
'    LBLCOUNT = Adodc1.Recordset.RecordCount
XPDtbBill_Change
End Sub

Private Sub ChangeLang()
    On Error Resume Next
    Me.Caption = "Create repeated GL Manual "
    Label1(0).Caption = Me.Caption
    Label2.Caption = " Date"
    Label3.Caption = "Total"

    DataGrid2.Columns(1).Caption = "NO#"
    DataGrid2.Columns(2).Caption = "description"
    DataGrid2.Columns(3).Caption = "Proposed Date"

    DataGrid2.RightToLeft = False
    ALLButton2.Caption = "Select All"
    ALLButton3.Caption = "Delete All"
    ALLButton1.Caption = "Execute"
End Sub

Private Sub XPDtbBill_Change()
    On Error Resume Next
    Dim sql As String
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText

    sql = "SELECT     ked_no, ked_serial, des, ked_date, [index],OK"
    sql = sql + " from dbo.ked_dawry WHERE     ( OK=0 AND  ked_date ='" & SQLDate(XPDtbBill.value) & "')"

    Adodc1.RecordSource = sql ' "   SELECT  * FROM KED_DAWRY WHERE KALEB=1"
    Adodc1.Refresh
    LBLCOUNT = Adodc1.Recordset.RecordCount

End Sub

