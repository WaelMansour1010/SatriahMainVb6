VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form KALEB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇÎÊíÇÑ ÞÇáÈ ÞíÏ"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10935
   Icon            =   "KALEB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   10935
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton Command1 
      Caption         =   "ÇÏÑÇÌ ÞÇáÈ ÇáÞíÏ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "Remark"
      DataSource      =   "Adodc1"
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2880
      Width           =   10575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   5280
      Top             =   5400
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "KALEB.frx":000C
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   23
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "NOTEDATE"
         Caption         =   "ÊÇÑíÎ ÇáÞíÏ"
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
         DataField       =   "REMARK"
         Caption         =   "ÇáÔÑÍ"
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
         DataField       =   "KALEB"
         Caption         =   "KALEB"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   8505.071
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
      EndProperty
   End
   Begin VB.Label LBLTYPE 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "KALEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Me.lbltype.Caption = 0 Then
    If Adodc1.Recordset.RecordCount > 0 And Not IsNull(Adodc1.Recordset.Fields!NoteID) Then
        FrmAccEditJournal.retrive1 (Adodc1.Recordset.Fields!NoteID)
        FrmAccEditJournal.Txt.Text = Text1.Text
    End If
ElseIf Me.lbltype.Caption = 1 Then
    If Adodc1.Recordset.RecordCount > 0 And Not IsNull(Adodc1.Recordset.Fields!NoteID) Then
        FrmAccEditJournal4.retrive1 (Adodc1.Recordset.Fields!NoteID)
        FrmAccEditJournal4.Txt.Text = Text1.Text
    End If


ElseIf Me.lbltype.Caption = 2 Then
    If Adodc1.Recordset.RecordCount > 0 And Not IsNull(Adodc1.Recordset.Fields!NoteID) Then
        FrmAccEditJournal3.retrive1 (Adodc1.Recordset.Fields!NoteID)
        FrmAccEditJournal3.Txt.Text = Text1.Text
    End If
    
End If

    Unload Me
End Sub

Private Sub DataGrid1_Click()
    Exit Sub
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "   SELECT  * FROM Notes WHERE KALEB=1"
    Adodc1.Refresh

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
End Sub

Private Sub ChangeLang()
    Me.Caption = "Select Template"
    Command1.Caption = "Insert Template"
    DataGrid1.RightToLeft = False

    DataGrid1.Columns(0).Caption = "Voucher Date"
    DataGrid1.Columns(1).Caption = "Description"

End Sub
