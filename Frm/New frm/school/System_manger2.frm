VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form System_manger2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16605
   Icon            =   "System_manger2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   16605
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "System_manger2.frx":000C
      Left            =   16800
      List            =   "System_manger2.frx":00A6
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   1200
      Width           =   3825
   End
   Begin VB.ComboBox CBOYearDigit 
      Height          =   315
      ItemData        =   "System_manger2.frx":0467
      Left            =   3600
      List            =   "System_manger2.frx":0471
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CheckBox ChkStore 
      Alignment       =   1  'Right Justify
      Caption         =   " ﬂÊÌœ ÿ»ﬁ« ··„Œ“‰"
      Height          =   195
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox xxx 
      Height          =   315
      ItemData        =   "System_manger2.frx":0486
      Left            =   2400
      List            =   "System_manger2.frx":0490
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Text            =   "4"
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtPrefix 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ì"
      Height          =   195
      Left            =   19560
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ì„·¡ «’›«—"
      Height          =   195
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "System_manger2.frx":049C
      Height          =   5295
      Left            =   120
      TabIndex        =   31
      Top             =   3120
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Caption         =   " "
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "BranchName"
         Caption         =   "«·›—⁄"
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
         DataField       =   "StoreCoding"
         Caption         =   " ﬂÊÌœ ÿ»ﬁ« ··„Œ“‰"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "1"
            FalseValue      =   "0"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "sanad_type"
         Caption         =   "‰Ê⁄ «·”‰œ"
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
         DataField       =   "YearDigit"
         Caption         =   "Œ«‰«  «·”‰…"
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
         DataField       =   "numbering_type"
         Caption         =   "«· —ﬁÌ„"
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
         DataField       =   "branch_no"
         Caption         =   "branch_no"
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
         DataField       =   "no_of_digit"
         Caption         =   "⁄œœ «·Œ«‰« "
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
         DataField       =   "zeros"
         Caption         =   "zeros"
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
         DataField       =   "start_at"
         Caption         =   "Ì»œ√ „‰"
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
         DataField       =   "departement"
         Caption         =   "departement"
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
      BeginProperty Column10 
         DataField       =   "end_at"
         Caption         =   "Ì‰ ÂÌ ›Ì"
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
            Alignment       =   3
            Object.Visible         =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3899.906
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3704.882
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   240
      TabIndex        =   25
      Top             =   9960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   480
      TabIndex        =   20
      Top             =   10080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Text            =   "1"
      Top             =   1680
      Width           =   3855
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Language  «··€…"
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "EN"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "System_manger2.frx":04B1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   20040
      TabIndex        =   17
      Top             =   2040
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   18720
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8880
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   9720
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   3360
      Top             =   10440
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Adodc1"
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
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "System_manger2.frx":04CD
      Left            =   5760
      List            =   "System_manger2.frx":04DD
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "System_manger2.frx":0502
      Left            =   9120
      List            =   "System_manger2.frx":05C6
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin ALLButtonS.ALLButton Command1 
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«÷«›…"
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
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "System_manger2.frx":0A66
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton Command2 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   8520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Õ–› ”ÿ—"
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
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "System_manger2.frx":0A82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   -240
      Top             =   10440
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Adodc1"
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
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   9120
      TabIndex        =   0
      Top             =   720
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ALLButtonS.ALLButton ApplyToall 
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«÷«›… ·ﬂ· «·›—Ê⁄"
      ENAB            =   0   'False
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
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "System_manger2.frx":0A9E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "System_manger2.frx":0ABA
      Height          =   5295
      Left            =   120
      TabIndex        =   40
      Top             =   3120
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Caption         =   " "
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "BranchName"
         Caption         =   "Branch Name"
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
         DataField       =   "StoreCoding"
         Caption         =   "Store Coding"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "1"
            FalseValue      =   "0"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "sanad_type"
         Caption         =   "Voucher "
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
         DataField       =   "YearDigit"
         Caption         =   "Year Digit"
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
         DataField       =   "numbering_type"
         Caption         =   "Numbering"
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
         DataField       =   "branch_no"
         Caption         =   "branch_no"
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
         DataField       =   "no_of_digit"
         Caption         =   "No Of Digit"
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
         DataField       =   "zeros"
         Caption         =   "zeros"
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
         DataField       =   "start_at"
         Caption         =   "Start From"
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
         DataField       =   "departement"
         Caption         =   "departement"
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
      BeginProperty Column10 
         DataField       =   "end_at"
         Caption         =   "End At"
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
            Alignment       =   3
            Object.Visible         =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3899.906
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3704.882
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ‰”Ìﬁ «·”‰…"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4800
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ã“¡ «·À«» "
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7080
      TabIndex        =   36
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13800
      TabIndex        =   35
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·Œ«‰« "
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13560
      TabIndex        =   33
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "   ﬂÊÌœ «·„” ‰œ« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   705
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   0
      Width           =   16455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ì‰ ÂÌ ›Ì"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6960
      TabIndex        =   30
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ì»œ√ „‰"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13560
      TabIndex        =   19
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «· —ﬁÌ„"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·”‰œ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13200
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "System_manger2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String
Dim branch_no As Integer
Dim departement_name As Integer

Private Sub ApplyToall_Click()
    On Error Resume Next

    If Not IsNumeric(Text3.text) Then
 
        If my_language = "E" Then
            MsgBox "Start from must be digit", vbCritical
        Else
            MsgBox "»œ«Ì… «· —ﬁÌ„ ÌÃ» «‰  ﬂÊ‰ «—ﬁ«„", vbCritical
        End If
    End If

    If Text3.text = "" Then
        Text3.text = 1
    End If

    If Combo1.text = "" Or Combo2.text = "" Then
        Exit Sub
    End If

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim My_SQL As String
    My_SQL = " select branch_id,branch_name,branch_namee  from TblBranchesData "
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly

    If rs.RecordCount = 0 Then
        Exit Sub
    End If

    For i = 1 To rs.RecordCount

        Adodc2.RecordSource = "select * from  sanad_numbering where branch_no=" & val(rs("branch_id").value) & "   and  sanad_no= " & Combo1.ListIndex
        Adodc2.Refresh

        If Adodc2.Recordset.RecordCount > 0 Then
            If my_language = "E" Then
                MsgBox "this voucher type alraedy defined with numbering method to change delete it and then try to add it again   " & rs("branch_namee").value, vbCritical
            Else
                MsgBox "Â–« «·‰Ê⁄ „‰ «·”‰œ«  „Õœœ  —ﬁÌ„… „‰ ﬁ»· ·«⁄«œÂ «· —ﬁÌ„ ﬁ„ »Õ–›… À„ ﬁ„ »«÷«› … „⁄ ‰Ê⁄ «· —ﬁÌ„ «·ÃœÌœ „—… «Œ—Ï ›Ì «·›—⁄   " & rs("branch_name").value, vbCritical
            End If

            Exit Sub
 
        End If

        rs.MoveNext
    Next i

    rs.MoveFirst

    For i = 1 To rs.RecordCount

        Adodc2.RecordSource = "select * from  sanad_numbering where branch_no=" & val(rs("branch_id").value) & "   and  sanad_no= " & Combo1.ListIndex
        Adodc2.Refresh

        If Adodc2.Recordset.RecordCount > 0 Then
            If my_language = "E" Then
                MsgBox "this voucher type alraedy defined with numbering method to change delete it and then try to add it again" & rs("branch_namee").value, vbCritical
            Else
                MsgBox "Â–« «·‰Ê⁄ „‰ «·”‰œ«  „Õœœ  —ﬁÌ„… „‰ ﬁ»· ·«⁄«œÂ «· —ﬁÌ„ ﬁ„ »Õ–›… À„ ﬁ„ »«÷«› … „⁄ ‰Ê⁄ «· —ﬁÌ„ «·ÃœÌœ „—… «Œ—Ï ›Ì «·›—⁄" & rs("branch_name").value, vbCritical
            End If
 
        End If

        branch_no = val(rs("branch_id").value)
        departement_name = 1
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields!Sanad_No = Combo1.ListIndex
        Adodc1.Recordset.Fields!SANAD_TYPE = Combo1.text
        Adodc1.Recordset.Fields!numbering_id = Combo2.ListIndex
        Adodc1.Recordset.Fields!numbering_type = Combo2.text
        Adodc1.Recordset.Fields!branch_no = val(rs("branch_id").value)
        Adodc1.Recordset.Fields!branchname = rs("branch_name").value
        Adodc1.Recordset.Fields!start_at = val(Text3.text)
        Adodc1.Recordset.Fields!end_at = val(Text2.text)
        Adodc1.Recordset.Fields!Departement = departement_name
        Adodc1.Recordset.Fields!no_of_digit = val(Text4.text)

        If Check1 = vbChecked Then
            Adodc1.Recordset.Fields!Zeros = 1
        Else
            Adodc1.Recordset.Fields!Zeros = 0
        End If



    Adodc1.Recordset.Fields!Prefix = (TxtPrefix.text)

    Adodc1.Recordset.Fields!YearDigit = Abs(val(CBOYearDigit.text)) ' IIf(val(CBOYearDigit.text) = 0, 4, 2)
 
    If ChkStore.value = vbChecked Then
        Adodc1.Recordset.Fields!StoreCoding = 1
    Else
        Adodc1.Recordset.Fields!StoreCoding = 0
    End If
    


        Adodc1.Recordset.update
        rs.MoveNext
    Next i
dcBranch.SetFocus
End Sub

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        ''Call Reload(Me)
 
    Else
        my_language = "A"
 
        ''Call Reload(Me)
    End If

End Sub

Function updateview()
   Dim str As String
    connection_string = Cn.ConnectionString

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText

  str = "select * from  sanad_numbering where 1=1 "
  If val(dcBranch.BoundText) <> 0 Then
 str = str & " and  branch_no=" & val(dcBranch.BoundText)
   End If
   
  If val(Combo1.ListIndex) <> -1 Then
 str = str & " and  sanad_no=" & val(Combo1.ListIndex)
   End If
   
str = str & "  order by sanad_type "
   
        Adodc1.RecordSource = str
    
 

    Adodc1.Refresh
    DataGrid2.Refresh



End Function

Private Sub Combo1_Click()
updateview
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, _
                         Shift As Integer)
 '   On Error Resume Next
 '   Combo1.Text = ""
If KeyCode = vbKeyReturn Then
updateview
End If

  AutoSel Combo1, KeyCode
  
End Sub

Private Sub Combo2_Click()

    If Combo2.ListIndex = 0 Then
        Text3.Enabled = False
    Else
        Text3.Enabled = True
    End If

    If Combo2.ListIndex = 2 Or Combo2.ListIndex = 3 Then
       CBOYearDigit.Visible = True
       Label5.Visible = True
    Else
            CBOYearDigit.Visible = False
       Label5.Visible = False
    End If
    


End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, _
                         Shift As Integer)
'    On Error Resume Next
'    Combo2.Text = ""
  AutoSel Combo2, KeyCode
  
End Sub

Private Sub Command1_Click()
    On Error Resume Next

    If Combo1.text = "" Or Combo2.text = "" Then
        Exit Sub
    End If

    Adodc2.RecordSource = "select * from  sanad_numbering where branch_no=" & val(dcBranch.BoundText) & "   and  sanad_no= " & Combo1.ListIndex & "and  Prefix='" & TxtPrefix & "'"

    If TxtPrefix.text = "" Then
        Adodc2.RecordSource = "select * from  sanad_numbering where branch_no=" & val(dcBranch.BoundText) & "   and  sanad_no= " & Combo1.ListIndex
    Else
        Adodc2.RecordSource = "select * from  sanad_numbering where branch_no=" & val(dcBranch.BoundText) & "   and  sanad_no= " & Combo1.ListIndex & " and  Prefix='" & TxtPrefix & "'"

    End If

    Adodc2.Refresh

    If Adodc2.Recordset.RecordCount > 0 Then
        If my_language = "E" Then
            MsgBox "this voucher type alraedy defined with numbering method to change delete it and then try to add it again", vbCritical
        Else
            MsgBox "Â–« «·‰Ê⁄ „‰ «·”‰œ«  „Õœœ  —ﬁÌ„… „‰ ﬁ»· ·«⁄«œÂ «· —ﬁÌ„ ﬁ„ »Õ–›… À„ ﬁ„ »«÷«› … „⁄ ‰Ê⁄ «· —ﬁÌ„ «·ÃœÌœ „—… «Œ—Ï", vbCritical
        End If

        Exit Sub
    End If

    If Not IsNumeric(Text3.text) Then
        If my_language = "E" Then
            MsgBox "Start from must be digit", vbCritical
        Else
            MsgBox "»œ«Ì… «· —ﬁÌ„ ÌÃ» «‰  ﬂÊ‰ «—ﬁ«„", vbCritical
        End If
    End If

    If Text3.text = "" Then
        Text3.text = 1
    End If

    branch_no = val(Me.dcBranch.BoundText)
    departement_name = 1
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!Sanad_No = Combo1.ListIndex
    Adodc1.Recordset.Fields!SANAD_TYPE = Combo1.text
    Adodc1.Recordset.Fields!numbering_id = Combo2.ListIndex
    Adodc1.Recordset.Fields!numbering_type = Combo2.text
    Adodc1.Recordset.Fields!branch_no = branch_no
    Adodc1.Recordset.Fields!branchname = Me.dcBranch.text
    Adodc1.Recordset.Fields!start_at = val(Text3.text)
    Adodc1.Recordset.Fields!end_at = val(Text2.text)
    Adodc1.Recordset.Fields!Departement = departement_name
    Adodc1.Recordset.Fields!no_of_digit = val(Text4.text)
    Adodc1.Recordset.Fields!Prefix = (TxtPrefix.text)

    Adodc1.Recordset.Fields!YearDigit = Abs(val(CBOYearDigit.text)) ' IIf(val(CBOYearDigit.text) = 0, 4, 2)


    If Check1.value = vbChecked Then
        Adodc1.Recordset.Fields!Zeros = 1
    Else
        Adodc1.Recordset.Fields!Zeros = 0
    End If

    If ChkStore.value = vbChecked Then
        Adodc1.Recordset.Fields!StoreCoding = 1
    Else
        Adodc1.Recordset.Fields!StoreCoding = 0
    End If
    
    Adodc1.Recordset.update
 
dcBranch.SetFocus
End Sub

Private Sub Command13_Click()
    On Error Resume Next
    DataCombo6.text = ""
    items_search.show
    items_search.case_id.Caption = 3

End Sub

Private Sub Command2_Click()
    On Error Resume Next
    X = MsgBox("   confirm delete  Â· «‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbCritical + vbYesNo)

    If X = vbNo Then
        Exit Sub
    End If

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.delete
        Adodc1.Refresh
        DataGrid1.Refresh
    End If

End Sub

Private Sub dcBranch_Change()

    'connection_string = Cn.ConnectionString

    'Adodc1.ConnectionString = connection_string
    'Adodc1.CommandType = adCmdText
'
  
'        Adodc1.RecordSource = "select * from  sanad_numbering where branch_no=" & val(dcBranch.BoundText) & " order by sanad_type "   '& " and departement='" & departement_name & "'"
'
'
'
'    Adodc1.Refresh
'    DataGrid2.Refresh
End Sub

Private Sub DcBranch_Click(Area As Integer)
updateview

End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
updateview
End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim My_SQL As String

If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select branch_id,branch_name from TblBranchesData    "
 Else
     My_SQL = "  select branch_id,branch_namee from TblBranchesData    "

 End If
 
    
    fill_combo dcBranch, My_SQL
  
    '   branch_no = 1
    ' departement_name = 1
    dcBranch.BoundText = branch_id
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    Me.dept_lbl = departement_name
    Me.emp_name_lbl = current_user_name
    emp_a.Caption = current_user_name
    dep_a.Caption = departement_name
    infoA.Visible = True
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChkStore.Caption = "Coding With Store"
        
        Label3.Caption = "Fixed Segment"
        Label4.Caption = "Voucher Type"
        Label6.Caption = "Coding Type"
        Label7.Caption = "Start From"
        Label9.Caption = "End At"
        DataGrid2.Columns(2).Caption = "Voucher Type"
        DataGrid2.Columns(4).Caption = "Coding Type"
        DataGrid2.Columns(8).Caption = "Start From"
        DataGrid2.Columns(10).Caption = "End At"
        Label2.Caption = "Branch"
             
        ApplyToall.Caption = "Add To All"
        Check1.Caption = "Zeros"
        check2.Caption = "Auto"
         
        '  InfoE.Visible = True
        ' infoA.Visible = False
 
        Label1.Caption = "No Of Digit"
        LblHeader = "Voucher Coding"
        Combo1.Clear
        Combo1.AddItem "Journal Entry Voucher"
        Combo1.AddItem "Expenses voucher "
        Combo1.AddItem "Cashing  voucher"
        Combo1.AddItem "Opening Balance"
        Combo1.AddItem "Payment Expenses"
        Combo1.AddItem "Bety cash Expenses"
        Combo1.AddItem "Purchase Invoice"
        Combo1.AddItem "Sales Invoice"
        Combo1.AddItem "Financial  Invoice"
        Combo1.AddItem "Recieve Voucher"
        Combo1.AddItem "Issue  Voucher"
        Combo1.AddItem "Product Order"
        Combo1.AddItem "íMoving Items Bt Inv"
        Combo1.AddItem "Stock Adjustement"
        Combo1.AddItem " Sales Return"
        Combo1.AddItem "íPurchase Return"
        Combo1.AddItem "íFinancial Transfer"
        Combo1.AddItem "Banks Deposite"
        Combo1.AddItem "Production Issue Vouchers"
        Combo1.AddItem "Production Recieve Vouchers"
        Combo1.AddItem "Collection and payment of checks"
        Combo1.AddItem "Typical Production Voucher Coding"
        Combo1.AddItem "Fixed Assets Purchase Invoices"
        Combo1.AddItem " Employee Allocations"
        Combo1.AddItem "Indirect Cost Vouchers"
        Combo1.AddItem "collection of premiums Vchr"
        Combo1.AddItem "Payment of premiums Vchr"
        Combo1.AddItem "Advanced Payment -Components"
        Combo1.AddItem "Disposal Of Fixed Assets " '28
        Combo1.AddItem "  Purchase order " '29
        Combo1.AddItem "Sales Order " '30
        Combo1.AddItem "Performa Invoices" '31
        Combo1.AddItem "Emp Advanced Payment Vchr  " '32
        Combo1.AddItem "Semi-Production Issue Vouchers" '33
        Combo1.AddItem "Semi-Production Recieve Vouchers" '34
        Combo1.AddItem "Era Vouchers" '35
        Combo1.AddItem "Issue Vocher For Danage Or Sample  " '36
        Combo1.AddItem "Trips Vocher   " '37
        Combo1.AddItem "Internal Order  " '38
        Combo1.AddItem "Reserve Items   " '39
        Combo1.AddItem "Bill compound   " '40
        Combo1.AddItem "Sales Quotations Request " '41
       Combo1.AddItem "Quotations " '42
       Combo1.AddItem "Sales Order Request " '43
        Combo1.AddItem "Sales Order  " '44
        
        Combo1.AddItem "Purchase Quotations Request  " '45
     Combo1.AddItem "Purchase Quotations  " '46
                
        Combo1.AddItem "Purchase  Request " '47
                        
        Combo1.AddItem "Purchase Order  " '48
                                
        Combo1.AddItem "Production Order  " '49
        Combo1.AddItem "Mintenance Bill  " '50
        Combo1.AddItem "Mintenance Entitlement to commissions  " '51
        Combo1.AddItem "ComputerCheck Bill  " '52
        Combo1.AddItem " Expenses Voucher-Multi" '53
        Combo1.AddItem "Shipment Order  " '54
        Combo1.AddItem "Shipment Voucher " '55
       Combo1.AddItem " Shipment Recieve Voucheri" '56
       Combo1.AddItem "Je Manual Entry" '57
       Combo1.AddItem "Expenses Request" '58
        Combo1.AddItem "General Cashing" '59
       Combo1.AddItem "Rent Contract" '60
              Combo1.AddItem "Reserve Production '8050 61"
              
              Combo1.AddItem "Allowance Discount" '8033  62
              
              Combo1.AddItem "Service Invoices" '8063  63
              
        Combo2.Clear
        Combo2.AddItem "Manual"
        Combo2.AddItem "Automatic"
        Combo2.AddItem "Monthly "
        Combo2.AddItem "Yearly"

        CMD_language.Caption = "⁄—»Ì"

        'Frame2.Visible = False
        'Frame1.Visible = True
        DataGrid2.Visible = False
        'Combo1.RightToLeft = False
        'Combo2.RightToLeft = False
        'Temp = Command1.left
        'Command1.left = Command2.left
        'Command2.left = Temp
        Command1.Caption = "Add"
        Command2.Caption = "Delete"
    End If

    'LoadSettings
    connection_string = Cn.ConnectionString

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText

    If SystemOptions.usertype = UserAdminAll Then
        Adodc1.RecordSource = "select * from  sanad_numbering  order by sanad_type  "
        ApplyToall.Enabled = True
    Else
        Adodc1.RecordSource = "select * from  sanad_numbering where branch_no=" & branch_id & " order by sanad_type "  '& " and departement='" & departement_name & "'"
        Me.dcBranch.Enabled = False
    End If

    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText

End Sub

