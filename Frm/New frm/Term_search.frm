VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Term_search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ «·»‰Êœ"
   ClientHeight    =   4425
   ClientLeft      =   3825
   ClientTop       =   2430
   ClientWidth     =   8430
   Icon            =   "Term_search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   8430
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "Term_search.frx":000C
      Height          =   2535
      Left            =   1560
      TabIndex        =   16
      Top             =   5520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "account_no"
         Caption         =   "—ﬁ„ «·Õ”«»"
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
         DataField       =   "account_name"
         Caption         =   "«”„ «·Õ”«»"
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
         DataField       =   "account_type"
         Caption         =   "account_type"
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
         DataField       =   "parent_no"
         Caption         =   "parent_no"
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
         DataField       =   "child_no"
         Caption         =   "child_no"
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
         DataField       =   "mezania_or_kayma"
         Caption         =   "mezania_or_kayma"
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
         DataField       =   "account_natural"
         Caption         =   "account_natural"
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
         DataField       =   "markas_taklefa"
         Caption         =   "markas_taklefa"
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
         DataField       =   "markas_taklefa_type"
         Caption         =   "markas_taklefa_type"
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
         DataField       =   "zmam"
         Caption         =   "zmam"
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
         DataField       =   "moazna"
         Caption         =   "moazna"
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
      BeginProperty Column12 
         DataField       =   "black_list"
         Caption         =   "black_list"
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
      BeginProperty Column13 
         DataField       =   "markas_taklefa_value"
         Caption         =   "markas_taklefa_value"
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
      BeginProperty Column14 
         DataField       =   "opening_balance"
         Caption         =   "opening_balance"
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
      BeginProperty Column15 
         DataField       =   "level"
         Caption         =   "level"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label2 
         Caption         =   "Term Name"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Term #"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   1455
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·»‰œ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "‘—Õ «·»‰œ"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "«Œ Ì«— "
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
      MICON           =   "Term_search.frx":0021
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
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   7200
      Top             =   7560
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Term_search.frx":003D
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "fullcode"
         Caption         =   "—ﬁ„ «·»‰œ"
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
         DataField       =   "des"
         Caption         =   "‘—Õ «·»‰œ"
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
         DataField       =   "account_type"
         Caption         =   "‰Ê⁄ «·Õ”«»"
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
         DataField       =   "parent_no"
         Caption         =   "parent_no"
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
         DataField       =   "child_no"
         Caption         =   "child_no"
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
         DataField       =   "mezania_or_kayma"
         Caption         =   "mezania_or_kayma"
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
         DataField       =   "account_natural"
         Caption         =   "account_natural"
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
         DataField       =   "markas_taklefa"
         Caption         =   "markas_taklefa"
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
         DataField       =   "markas_taklefa_type"
         Caption         =   "markas_taklefa_type"
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
         DataField       =   "zmam"
         Caption         =   "zmam"
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
         DataField       =   "moazna"
         Caption         =   "moazna"
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
      BeginProperty Column12 
         DataField       =   "black_list"
         Caption         =   "black_list"
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
      BeginProperty Column13 
         DataField       =   "markas_taklefa_value"
         Caption         =   "markas_taklefa_value"
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
      BeginProperty Column14 
         DataField       =   "opening_balance"
         Caption         =   "opening_balance"
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
      BeginProperty Column15 
         DataField       =   "level"
         Caption         =   "«·„” ÊÏ"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   5595.024
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Term_search.frx":0052
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "account_no"
         Caption         =   "account_no"
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
         DataField       =   "account_name"
         Caption         =   "account_name"
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
         DataField       =   "account_type"
         Caption         =   "account_type"
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
         DataField       =   "parent_no"
         Caption         =   "parent_no"
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
         DataField       =   "child_no"
         Caption         =   "child_no"
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
         DataField       =   "mezania_or_kayma"
         Caption         =   "mezania_or_kayma"
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
         DataField       =   "account_natural"
         Caption         =   "account_natural"
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
         DataField       =   "markas_taklefa"
         Caption         =   "markas_taklefa"
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
         DataField       =   "markas_taklefa_type"
         Caption         =   "markas_taklefa_type"
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
         DataField       =   "zmam"
         Caption         =   "zmam"
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
         DataField       =   "moazna"
         Caption         =   "moazna"
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
      BeginProperty Column12 
         DataField       =   "black_list"
         Caption         =   "black_list"
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
      BeginProperty Column13 
         DataField       =   "markas_taklefa_value"
         Caption         =   "markas_taklefa_value"
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
      BeginProperty Column14 
         DataField       =   "opening_balance"
         Caption         =   "opening_balance"
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
      BeginProperty Column15 
         DataField       =   "level"
         Caption         =   "level"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "Term_search.frx":0067
      Height          =   2535
      Left            =   1560
      TabIndex        =   12
      Top             =   5520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "account_no"
         Caption         =   "account_no"
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
         DataField       =   "account_name"
         Caption         =   "account_name"
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
         DataField       =   "account_type"
         Caption         =   "account_type"
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
         DataField       =   "parent_no"
         Caption         =   "parent_no"
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
         DataField       =   "child_no"
         Caption         =   "child_no"
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
         DataField       =   "mezania_or_kayma"
         Caption         =   "mezania_or_kayma"
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
         DataField       =   "account_natural"
         Caption         =   "account_natural"
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
         DataField       =   "markas_taklefa"
         Caption         =   "markas_taklefa"
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
         DataField       =   "markas_taklefa_type"
         Caption         =   "markas_taklefa_type"
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
         DataField       =   "zmam"
         Caption         =   "zmam"
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
         DataField       =   "moazna"
         Caption         =   "moazna"
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
      BeginProperty Column12 
         DataField       =   "black_list"
         Caption         =   "black_list"
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
      BeginProperty Column13 
         DataField       =   "markas_taklefa_value"
         Caption         =   "markas_taklefa_value"
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
      BeginProperty Column14 
         DataField       =   "opening_balance"
         Caption         =   "opening_balance"
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
      BeginProperty Column15 
         DataField       =   "level"
         Caption         =   "level"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   585
      Left            =   6600
      Top             =   6840
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
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   5640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·ﬂ·"
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
      MICON           =   "Term_search.frx":007C
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
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   7680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Õ–› «·ﬂ·"
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
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Term_search.frx":0098
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton4 
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   7680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
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
      MICON           =   "Term_search.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "Term_search.frx":00D0
      Height          =   2535
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483648
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "Fullcode"
         Caption         =   "project Code."
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
         DataField       =   "Project_name"
         Caption         =   "project Name"
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
         DataField       =   "account_type"
         Caption         =   "‰Ê⁄ «·Õ”«»"
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
         DataField       =   "parent_no"
         Caption         =   "parent_no"
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
         DataField       =   "child_no"
         Caption         =   "child_no"
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
         DataField       =   "mezania_or_kayma"
         Caption         =   "mezania_or_kayma"
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
         DataField       =   "account_natural"
         Caption         =   "account_natural"
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
         DataField       =   "markas_taklefa"
         Caption         =   "markas_taklefa"
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
         DataField       =   "markas_taklefa_type"
         Caption         =   "markas_taklefa_type"
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
         DataField       =   "zmam"
         Caption         =   "zmam"
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
         DataField       =   "moazna"
         Caption         =   "moazna"
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
      BeginProperty Column12 
         DataField       =   "black_list"
         Caption         =   "black_list"
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
      BeginProperty Column13 
         DataField       =   "markas_taklefa_value"
         Caption         =   "markas_taklefa_value"
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
      BeginProperty Column14 
         DataField       =   "opening_balance"
         Caption         =   "opening_balance"
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
      BeginProperty Column15 
         DataField       =   "level"
         Caption         =   "«·„” ÊÏ"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   5595.024
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Label case_index 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label SANAD_TYPE 
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label sandat_pc_no 
      Caption         =   "0"
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label case_id 
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Term_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim first_run As Boolean

Private Sub ALLButton1_Click()

    On Error Resume Next

    If case_id.Caption = 1 And Adodc1.Recordset.RecordCount >= 0 Then

        With FrmEmpSalary4.Grid
            ' .TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("to_project")) = Adodc1.Recordset.Fields!Project_name
            .TextMatrix(.Row, .ColIndex("to_projectid")) = Adodc1.Recordset.Fields!id
  
            '  FrmEmpSalary4.Grid_AfterEdit .Row, 21
            ' FrmAccEditJournal.Fg_Journal_AfterEdit .Row, 2
  
        End With
 
    End If

    If case_id.Caption = 55 And Adodc1.Recordset.RecordCount >= 0 Then
 
        FrmExpensesType.DboParentAccount.text = Adodc1.Recordset.Fields!account_name
 
    End If

    If case_id.Caption = 66 And Adodc1.Recordset.RecordCount >= 0 Then
 
        FrmRevenuesTypes.DboParentAccount.text = Adodc1.Recordset.Fields!account_name
 
    End If

    If case_id.Caption = 77 And Adodc1.Recordset.RecordCount >= 0 Then
 
        FrmCustemers.DboParentAccount.text = Adodc1.Recordset.Fields!account_name
 
    End If

    If case_id.Caption = 88 And Adodc1.Recordset.RecordCount >= 0 Then
 
        FrmCompany.DboParentAccount.text = Adodc1.Recordset.Fields!account_name
 
    End If

    If case_id.Caption = 1000 And Adodc1.Recordset.RecordCount >= 0 Then

        With FrmAccEditJournal1.Fg_Journal
            ' .TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmAccEditJournal1.Fg_Journal_AfterEdit .Row, 2
        End With
 
    End If

    If case_id.Caption = 0 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmAccountCharts.Retrive (Adodc1.Recordset.Fields!Account_Code)
    End If

    If case_id.Caption = 1 And Adodc1.Recordset.RecordCount >= 0 Then
        If SystemOptions.UserInterface = EnglishInterface Then
            FrmAccountingReport.Set_account_code Adodc1.Recordset.Fields!Account_Code, Adodc1.Recordset.Fields!Account_NameEng
        Else
            FrmAccountingReport.Set_account_code Adodc1.Recordset.Fields!Account_Code, Adodc1.Recordset.Fields!account_name
        End If

    End If

    If case_id.Caption = 2 And Adodc1.Recordset.RecordCount >= 0 Then
        frmsandat_kabd.DataCombo1.text = Adodc1.Recordset.Fields!account_no
        frmsandat_kabd.Text2.text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 3 And Adodc1.Recordset.RecordCount >= 0 Then
        Voucher_search.Text1(2).text = Adodc1.Recordset.Fields!account_serial
    End If

    If case_id.Caption = 40 And Adodc1.Recordset.RecordCount >= 0 Then
        mowazna.DataCombo1.text = Adodc1.Recordset.Fields!account_serial
        mowazna.Text2.text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 3 And Adodc1.Recordset.RecordCount >= 0 Then
        frmcustomer.DataCombo1.text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 4 And Adodc1.Recordset.RecordCount >= 0 Then
        frmVendors.DataCombo1.text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 5 And Adodc1.Recordset.RecordCount >= 0 Then
        rased_eftetahy_account.DataCombo1.text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 6 And Adodc1.Recordset.RecordCount >= 0 Then
        frmmasrouf.DataCombo1.text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 7 And Adodc1.Recordset.RecordCount >= 0 Then
        frmboxes.DataCombo1.text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 8 And Adodc1.Recordset.RecordCount >= 0 Then
        frmrdrod.DataCombo1.text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 90 And Adodc1.Recordset.RecordCount >= 0 Then
        REPORTSFRM.DataCombo8.text = Adodc1.Recordset.Fields!account_no
    End If

    If case_id.Caption = 700 And Adodc1.Recordset.RecordCount >= 0 Then
        baranches.DataCombo1(case_index).text = Adodc1.Recordset.Fields!account_no
    End If

    '
    Unload Me
End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next

    If Adodc1.Recordset.RecordCount >= 0 Then
        Adodc1.Recordset.MoveFirst

        For i = 1 To Adodc1.Recordset.RecordCount
            Adodc2.Recordset.AddNew
            Adodc2.Recordset.Fields!account_no = Adodc1.Recordset.Fields!account_no
            Adodc2.Recordset.Fields!account_name = Adodc1.Recordset.Fields!account_name
  
            Adodc2.Recordset.update
            Adodc1.Recordset.MoveNext

        Next i

        Adodc2.Refresh
        DataGrid3.Refresh
        DataGrid4.Refresh
    End If

End Sub

Private Sub ALLButton3_Click()
    On Error Resume Next

    x = MsgBox("Confirm Cancell All selection  √ﬂÌœ «·€«¡ ﬂ· «· ÕœÌœ", vbCritical + vbYesNo)

    If x = vbNo Then Exit Sub

    For i = 1 To Adodc2.Recordset.RecordCount
        Adodc2.Recordset.delete
        Adodc2.Recordset.MoveNext
    
    Next i

    DataGrid3.Refresh
    DataGrid4.Refresh
        
End Sub

Private Sub ALLButton4_Click()
    'On Error Resume Next
    On Error Resume Next

    If Adodc2.Recordset.RecordCount = 0 Then Exit Sub
    x = MsgBox("Confirm PROCESS  √ﬂÌœ «·⁄„·Ì…", vbInformation + vbYesNo)

    If x = vbNo Then Exit Sub

    If Adodc2.Recordset.RecordCount = 0 Then Exit Sub
    Adodc2.Recordset.MoveFirst

    If case_id.Caption = 30 Then 'ked

        For i = 0 To Adodc2.Recordset.RecordCount - 1
            'If frmtransactions.Text1.Text = "" Then MsgBox "you must select TRip first ·«»œ „‰ «Œ Ì«— —Õ·… ", vbCritical: Exit Sub
 
            frmsandat_ked.Adodc2.Recordset.AddNew
            frmsandat_ked.Adodc2.Recordset.Fields!account_no = Adodc2.Recordset.Fields!account_no
            frmsandat_ked.Adodc2.Recordset.Fields!account_name = Adodc2.Recordset.Fields!account_name
            frmsandat_ked.Adodc2.Recordset.Fields!sandat_pc_no = sandat_pc_no.Caption
            frmsandat_ked.Adodc2.Recordset.Fields!SANAD_TYPE = Me.SANAD_TYPE
            frmsandat_ked.Text6.text = frmsandat_ked.Text11.text
            frmsandat_ked.Adodc2.Recordset.update
 
            frmsandat_ked.Adodc2.Recordset.Fields!Sanad_No = frmsandat_ked.Text1.text
            frmsandat_ked.Adodc2.Recordset.Fields!sanad_source = "ÌœÊÌ"
            frmsandat_ked.Adodc2.Recordset.Fields!Date = DateValue(Now)

            Adodc2.Recordset.MoveNext

        Next i

        frmsandat_ked.Adodc2.RecordSource = "SELECT * FROM sandat_ked_details WHERE sandat_pc_no=" & Me.sandat_pc_no.Caption
        frmsandat_ked.Adodc2.Refresh
        frmsandat_ked.DataGrid1.Refresh
    End If

    If case_id.Caption = 2 Then 'kabd

        For i = 0 To Adodc2.Recordset.RecordCount - 1
            'If frmtransactions.Text1.Text = "" Then MsgBox "you must select TRip first ·«»œ „‰ «Œ Ì«— —Õ·… ", vbCritical: Exit Sub
 
            frmsandat_kabd.Adodc2.Recordset.AddNew
            frmsandat_kabd.Adodc2.Recordset.Fields!account_no = Adodc2.Recordset.Fields!account_no
            frmsandat_kabd.Adodc2.Recordset.Fields!account_name = Adodc2.Recordset.Fields!account_name
            frmsandat_kabd.Adodc2.Recordset.Fields!sandat_pc_no = sandat_pc_no.Caption
            frmsandat_kabd.Adodc2.Recordset.Fields!SANAD_TYPE = Me.SANAD_TYPE
            frmsandat_kabd.Text6.text = frmsandat_kabd.Text11.text

            frmsandat_kabd.Adodc2.Recordset.update
 
            frmsandat_kabd.Adodc2.Recordset.Fields!Sanad_No = frmsandat_kabd.Text1.text
            frmsandat_kabd.Adodc2.Recordset.Fields!sanad_source = "ÌœÊÌ"
            frmsandat_kabd.Adodc2.Recordset.Fields!Date = DateValue(Now)
            frmsandat_kabd.Adodc2.Recordset.Fields!BOX_name = frmsandat_kabd.DataCombo3.text
            frmsandat_kabd.Adodc2.Recordset.Fields!bona_3la = frmsandat_kabd.Text11.text
            Adodc2.Recordset.MoveNext

        Next i

        frmsandat_kabd.Adodc2.RecordSource = "SELECT * FROM sandat_ked_details WHERE sandat_pc_no=" & Me.sandat_pc_no.Caption
        frmsandat_kabd.Adodc2.Refresh
        frmsandat_kabd.DataGrid1.Refresh
    End If

    If case_id.Caption = 40 Then 'sarf

        For i = 0 To Adodc2.Recordset.RecordCount - 1
            'If frmtransactions.Text1.Text = "" Then MsgBox "you must select TRip first ·«»œ „‰ «Œ Ì«— —Õ·… ", vbCritical: Exit Sub
 
            frmsandat_sarf.Adodc2.Recordset.AddNew
            frmsandat_sarf.Adodc2.Recordset.Fields!account_no = Adodc2.Recordset.Fields!account_no
            frmsandat_sarf.Adodc2.Recordset.Fields!account_name = Adodc2.Recordset.Fields!account_name
            frmsandat_sarf.Adodc2.Recordset.Fields!sandat_pc_no = sandat_pc_no.Caption
            frmsandat_sarf.Adodc2.Recordset.Fields!SANAD_TYPE = Me.SANAD_TYPE
            frmsandat_sarf.Text6.text = frmsandat_sarf.Text11.text

            frmsandat_sarf.Adodc2.Recordset.update
 
            frmsandat_sarf.Adodc2.Recordset.Fields!Sanad_No = frmsandat_sarf.Text1.text
            frmsandat_sarf.Adodc2.Recordset.Fields!sanad_source = "ÌœÊÌ"
            frmsandat_sarf.Adodc2.Recordset.Fields!Date = DateValue(Now)
            frmsandat_sarf.Adodc2.Recordset.Fields!BOX_name = frmsandat_sarf.DataCombo3.text
            frmsandat_sarf.Adodc2.Recordset.Fields!bona_3la = frmsandat_sarf.Text11.text
            Adodc2.Recordset.MoveNext

        Next i

        frmsandat_kabd.Adodc2.RecordSource = "SELECT * FROM sandat_ked_details WHERE sandat_pc_no=" & Me.sandat_pc_no.Caption
        frmsandat_kabd.Adodc2.Refresh
        frmsandat_kabd.DataGrid1.Refresh
    End If

    If case_id.Caption = 50 Then 'rased eftetahy

        For i = 0 To Adodc2.Recordset.RecordCount - 1
            'If frmtransactions.Text1.Text = "" Then MsgBox "you must select TRip first ·«»œ „‰ «Œ Ì«— —Õ·… ", vbCritical: Exit Sub
 
            rased_eftetahy_account.Adodc2.Recordset.AddNew
            rased_eftetahy_account.Adodc2.Recordset.Fields!account_no = Adodc2.Recordset.Fields!account_no
            rased_eftetahy_account.Adodc2.Recordset.Fields!account_name = Adodc2.Recordset.Fields!account_name
            rased_eftetahy_account.Adodc2.Recordset.Fields!sandat_pc_no = sandat_pc_no.Caption
            rased_eftetahy_account.Adodc2.Recordset.Fields!SANAD_TYPE = Me.SANAD_TYPE
            rased_eftetahy_account.Text6.text = rased_eftetahy_account.Text11.text

            rased_eftetahy_account.Adodc2.Recordset.update
 
            rased_eftetahy_account.Adodc2.Recordset.Fields!Sanad_No = rased_eftetahy_account.Text1.text
            rased_eftetahy_account.Adodc2.Recordset.Fields!sanad_source = "ÌœÊÌ"
            rased_eftetahy_account.Adodc2.Recordset.Fields!Date = DateValue(Now)

            Adodc2.Recordset.MoveNext

        Next i

        rased_eftetahy_account.Adodc2.RecordSource = "SELECT * FROM sandat_ked_details WHERE sandat_pc_no=" & Me.sandat_pc_no.Caption
        rased_eftetahy_account.Adodc2.Refresh
        rased_eftetahy_account.DataGrid1.Refresh
    End If

    Unload Me
End Sub

Private Sub DataGrid1_Click()
    On Error Resume Next

    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    Adodc2.Recordset.AddNew
    Adodc2.Recordset.Fields!account_no = Adodc1.Recordset.Fields!account_no
    Adodc2.Recordset.Fields!account_name = Adodc1.Recordset.Fields!account_name
  
    Adodc2.Recordset.update
    DataGrid3.Refresh
    DataGrid4.Refresh

End Sub

Private Sub DataGrid2_Click()
    ALLButton1_Click

End Sub

Private Sub DataGrid3_KeyUp(KeyCode As Integer, _
                            Shift As Integer)
    On Error Resume Next

    If KeyCode = 46 Then
        If Adodc2.Recordset.RecordCount > 0 Then
            Adodc2.Recordset.delete
            Adodc2.Refresh
            DataGrid3.Refresh
            DataGrid4.Refresh
        End If

    End If

End Sub

Private Sub DataGrid4_KeyUp(KeyCode As Integer, _
                            Shift As Integer)
    On Error Resume Next

    If KeyCode = 46 Then
        If Adodc2.Recordset.RecordCount > 0 Then
            Adodc2.Recordset.delete
            Adodc2.Refresh
            DataGrid3.Refresh
            DataGrid4.Refresh
        End If

    End If

End Sub

Private Sub DataGrid5_Click()
    ALLButton1_Click
End Sub

Private Sub Form_Activate()
    On Error Resume Next

    'If first_run = False Then
    'first_run = True
    'If case_id.Caption = 1 Then

    'Adodc1.RecordSource = "select * from   ACCOUNTS where last_account=1"
    'Adodc1.Refresh
    '

    '
    'Else
    'If case_id.Caption = 90 Or case_id.Caption = 3 Or case_id.Caption = 4 Or case_id.Caption = 7 Or case_id.Caption = 2 Or case_id.Caption = 30 Or case_id.Caption = 40 Or case_id.Caption = 6 Or case_id.Caption = 50 Then
    '        Sql = "select * from accounts where last_account=1 and (account_type='›—⁄Ì' or  account_type='sub') and NOT (account_no IS NULL)   order by account_no"
    'Adodc1.RecordSource = Sql
    'Adodc1.Refresh

    'Else
 
    'Adodc1.RecordSource = "select *  from account_index where black_list=0 and  NOT (account_no IS NULL)   order by account_no"
    'Adodc1.Refresh
    'If case_id.Caption = 0 Then
    'Adodc1.RecordSource = "select *  from account_index where  NOT (account_no IS NULL)   order by account_no"
    'Adodc1.Refresh
    '
    '
    '
    'End If
    'End If
    'End If
    'End If

    'Sql = Adodc1.RecordSource
    'DataGrid1.Refresh
    'DataGrid2.Refresh
 
End Sub

Private Sub Form_Load()
    On Error Resume Next
    connection_string = Cn.ConnectionString

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    sql = "select * from projects_des "
    Adodc1.RecordSource = sql
    Adodc1.Refresh

    '

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
 
        Label6.Caption = "ACC.Code"
        Label1.Caption = "ACC. Name"
        Me.Caption = "Accounts Search"

        DataGrid5.Visible = True
        DataGrid2.Visible = False

    Else
 
        DataGrid2.Visible = True
        DataGrid5.Visible = False

    End If

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
   
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from accounts_temp"
    Adodc2.Refresh
      
    For i = 0 To Adodc2.Recordset.RecordCount
        Adodc2.Recordset.delete
        Adodc2.Recordset.MoveNext
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    first_run = False
End Sub

Private Sub Text1_KeyUp(Index As Integer, _
                        KeyCode As Integer, _
                        Shift As Integer)
    On Error Resume Next

    If Index = 0 Then
                     
        sql = "select * from projects where fullcode like '%" & Text1(Index).text & "%'"
    End If
                      
    If Index = 1 Then
    
        sql = "select * from projects where   Project_name like'%" & Text1(Index).text & "%'"
    End If
                       
    Adodc1.RecordSource = sql
    Adodc1.Refresh
                      
    If Adodc1.Recordset.RecordCount = 0 Then
          
        '  MsgBox "not fount  ·«ÌÊÃœ ‰ «∆Ã ··»ÕÀ", vbInformation
    End If
 
    ' End If
End Sub
