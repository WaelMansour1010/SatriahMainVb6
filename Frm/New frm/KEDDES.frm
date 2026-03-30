VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form KEDDES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬁÊ«·» «·‘—Õ"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   12435
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "«œ—«Ã"
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
      MICON           =   "KEDDES.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "ked"
      DataSource      =   "Adodc1"
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3360
      Width           =   11655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   6240
      Top             =   6000
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
      Caption         =   " Õ—Ìﬂ"
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
      Bindings        =   "KEDDES.frx":001C
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "code"
         Caption         =   "«·ﬂÊœ"
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
         DataField       =   "ked"
         Caption         =   "«·‘—Õ"
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
            Object.Visible         =   -1  'True
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   10005.17
         EndProperty
      EndProperty
   End
   Begin VB.Label rowno 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      Height          =   135
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—Õ"
      Height          =   375
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·ﬂÊœ"
      Height          =   375
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label case_id 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   615
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "KEDDES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ALLButton1_Click()

    If Me.case_id = 0 Then
        FrmAccEditJournal.Txt.Text = Me.Text1.Text
    ElseIf Me.case_id = 1 Then
        'FrmAccEditJournal.TxtDes.text = Me.Text1.text

        FrmAccEditJournal.Fg_Journal.Cell(flexcpData, Me.rowno, FrmAccEditJournal.Fg_Journal.ColIndex("Des")) = Me.Text1.Text
        FrmAccEditJournal.Fg_Journal.TextMatrix(Me.rowno, FrmAccEditJournal.Fg_Journal.ColIndex("des")) = Me.Text1.Text
   ElseIf Me.case_id = 2 Then
        FrmAccEditJournal3.Txt.Text = Me.Text1.Text
    ElseIf Me.case_id = 3 Then
        'FrmAccEditJournal.TxtDes.text = Me.Text1.text

        FrmAccEditJournal3.Fg_Journal.Cell(flexcpData, Me.rowno, FrmAccEditJournal3.Fg_Journal.ColIndex("Des")) = Me.Text1.Text
        FrmAccEditJournal3.Fg_Journal.TextMatrix(Me.rowno, FrmAccEditJournal3.Fg_Journal.ColIndex("des")) = Me.Text1.Text
   
    ElseIf Me.case_id = 4 Then
        FrmAccEditJournal4.Txt.Text = Me.Text1.Text
        
       ElseIf Me.case_id = 5 Then
        'FrmAccEditJournal.TxtDes.text = Me.Text1.text

        FrmAccEditJournal4.Fg_Journal.Cell(flexcpData, Me.rowno, FrmAccEditJournal3.Fg_Journal.ColIndex("Des")) = Me.Text1.Text
        FrmAccEditJournal4.Fg_Journal.TextMatrix(Me.rowno, FrmAccEditJournal3.Fg_Journal.ColIndex("des")) = Me.Text1.Text
   
   
        'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)

    End If

    Unload Me
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "  select * FROM ked_desc"
    Adodc1.Refresh

End Sub

Private Sub ChangeLang()
    Me.Caption = "Des Templates"
    DataGrid1.Columns(0).Caption = "Code"
    DataGrid1.Columns(1).Caption = "DES"
    ALLButton1.Caption = "Insert"
    Label1.Caption = "Code"
    Label2.Caption = "Des"
    DataGrid1.RightToLeft = False
End Sub

Private Sub Text2_Change()

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "  select * FROM ked_desc where code like '%" & Text2.Text & "%'"
    Adodc1.Refresh
End Sub

Private Sub Text3_Change()

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "  select * FROM ked_desc where ked like '%" & Text3.Text & "%'"
    Adodc1.Refresh
End Sub
