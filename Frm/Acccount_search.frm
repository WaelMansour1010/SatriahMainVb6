VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Account_search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ «·Õ”«»«  "
   ClientHeight    =   5685
   ClientLeft      =   3825
   ClientTop       =   2430
   ClientWidth     =   14070
   Icon            =   "Acccount_search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   14070
   Begin VB.CheckBox chkenter 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»ÕÀ »œÊ‰ «‰ —"
      Height          =   375
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox CboNameSearch 
      Height          =   315
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   840
      Width           =   1515
   End
   Begin VB.ComboBox CboAccountCodeSearch 
      Height          =   315
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   360
      Width           =   1515
   End
   Begin VB.Frame Frame4 
      Caption         =   "‰Ê⁄ «·ﬂÊœ"
      Height          =   495
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   240
      Width           =   5055
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»"
         Height          =   195
         Index           =   0
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄„Ì·"
         Height          =   195
         Index           =   1
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ê—œ"
         Height          =   195
         Index           =   2
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊŸ›"
         Height          =   195
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "Acccount_search.frx":000C
      Height          =   2535
      Left            =   2280
      TabIndex        =   16
      Top             =   8160
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
         Caption         =   "Account Name"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Account#"
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
      Left            =   12600
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂÊœ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·Õ”«»"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   285
         TabIndex        =   7
         Top             =   720
         Width           =   810
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   8400
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
      MICON           =   "Acccount_search.frx":0021
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
      Left            =   4800
      TabIndex        =   0
      Top             =   840
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   7920
      Top             =   10680
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
      Bindings        =   "Acccount_search.frx":003D
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16776960
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   19
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
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
         DataField       =   "Account_Serial"
         Caption         =   "ﬂÊœ «·Õ”«»"
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
         DataField       =   "FirstName"
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
         DataField       =   "ParentName"
         Caption         =   "«·Õ”«» «·⁄«„"
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
         DataField       =   "RootName"
         Caption         =   "«·Õ”«» «·—∆Ì”Ì"
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
         MarqueeStyle    =   1
         BeginProperty Column00 
            Alignment       =   2
            Object.Visible         =   0   'False
            WrapText        =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            WrapText        =   -1  'True
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            WrapText        =   -1  'True
            ColumnWidth     =   5355.213
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Object.Visible         =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   2894.74
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Object.Visible         =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   2894.74
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
      Bindings        =   "Acccount_search.frx":0052
      Height          =   2535
      Left            =   720
      TabIndex        =   2
      Top             =   8280
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
      Bindings        =   "Acccount_search.frx":0067
      Height          =   2535
      Left            =   2280
      TabIndex        =   12
      Top             =   8160
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
      Left            =   7320
      Top             =   9960
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
      Left            =   5040
      TabIndex        =   13
      Top             =   8280
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
      MICON           =   "Acccount_search.frx":007C
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
      Left            =   5160
      TabIndex        =   14
      Top             =   10800
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
      MICON           =   "Acccount_search.frx":0098
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
      Left            =   2520
      TabIndex        =   15
      Top             =   10800
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
      MICON           =   "Acccount_search.frx":00B4
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
      Bindings        =   "Acccount_search.frx":00D0
      Height          =   3855
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777088
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
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
         DataField       =   "Account_Serial"
         Caption         =   "Account Code"
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
         DataField       =   "FirstName"
         Caption         =   "Account Name"
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
         DataField       =   "ParentName"
         Caption         =   "Parent Name"
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
         DataField       =   "RootName"
         Caption         =   "RootName"
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
            ColumnWidth     =   5355.213
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2894.74
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2894.74
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   8
      Left            =   3630
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   870
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   11
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "›Ì Õ«·… ÊÃÊœ «ﬂÀ— „‰ Õ”«» ›Ì «·»ÕÀ «÷€ÿ «‰ — ·«Œ Ì«— «·Õ”«» «·„ÿ·Ê»"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1320
      Width           =   5895
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
      Left            =   6720
      TabIndex        =   18
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label sandat_pc_no 
      Caption         =   "0"
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   7800
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
Attribute VB_Name = "Account_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim first_run As Boolean
Dim mClick As Boolean
Public mIndex As Integer
Private Sub ALLButton1_Click()

    On Error Resume Next
    
    'If sandat_pc_no <> 0 Then
    'If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    ' Adodc2.Recordset.AddNew
    '  Adodc2.Recordset.Fields!account_no = Adodc1.Recordset.Fields!account_no
    '   Adodc2.Recordset.Fields!account_name = Adodc1.Recordset.Fields!account_name
  
    '  Adodc2.Recordset.Update
    '  DataGrid3.Refresh
    'DataGrid4.Refresh
    'End If
    
 If case_id.Caption = 78912 And Adodc1.Recordset.RecordCount >= 0 Then
    
     'FrmEditUsers.ListStoreSelected.AddItem ListStoreall.List(i)
        'FrmEditUsers.ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(i)
        
        Dim i As Long

        For i = 0 To FrmEditUsers.ListAllAccount.ListCount - 1
            If FrmEditUsers.ListAllAccount.ItemData(i) = (Adodc1.Recordset.Fields!Account_ID) Then
                FrmEditUsers.ListAllAccount.Selected(i) = True
                
                
            End If
        Next

        'FrmEditUsers.ListStoreall.Selected(val(Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("id")))) = True

End If

    
    If case_id.Caption = 20200201 And Adodc1.Recordset.RecordCount >= 0 Then
       
        Ageng_all.DboParentAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
        
    End If
 
 
    If case_id.Caption = 50 And Adodc1.Recordset.RecordCount >= 0 Then
        tahlil_maly.DataCombo1.Text = Adodc1.Recordset.Fields!account_serial
        tahlil_maly.Text4.Text = Adodc1.Recordset.Fields!FirstName
    End If

    If case_id.Caption = 60 And Adodc1.Recordset.RecordCount >= 0 Then
        tahlil_maly.DataCombo2.Text = Adodc1.Recordset.Fields!account_serial
        tahlil_maly.Text6.Text = Adodc1.Recordset.Fields!FirstName
    End If


    If case_id.Caption = 260219 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmBoxDrawing.DCAccounts2.BoundText = (Adodc1.Recordset.Fields!Account_code)
  '      Exit Sub
 End If
 
    If case_id.Caption = 126575 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpenses30.DcbAccount.BoundText = Adodc1.Recordset.Fields(0)
        
  '      Exit Sub
 End If
  
 
    If case_id.Caption = 120519 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpenses301.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
  '      Exit Sub
 End If
  
  
 If case_id.Caption = 789725 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmAccountCharts.DboParentAccount2.BoundText = (Adodc1.Recordset.Fields!Account_code)
  '      Exit Sub
 End If
  
 
 If case_id.Caption = 90519 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmCars.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
  '      Exit Sub
 End If
  
  
 If case_id.Caption = 10519 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpenses30.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
  '      Exit Sub
 End If
 

    If case_id.Caption = 2602191 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmBoxDrawing.DCAccounts1.BoundText = (Adodc1.Recordset.Fields!Account_code)
  '      Exit Sub
 End If
 



    If (case_id.Caption = 700 Or case_id.Caption = 1700) And Adodc1.Recordset.RecordCount >= 0 Then
        'tahlil_maly.DataCombo2.text = Adodc1.Recordset.Fields!account_serial
        If SystemOptions.UserInterface = EnglishInterface Then
            baranchesE.DataCombo1(Me.case_index).Text = Adodc1.Recordset.Fields!FirstName 'Account_NameEng
        Else
            baranches.DataCombo1(Me.case_index).Text = Adodc1.Recordset.Fields!FirstName
        End If
    End If

    If case_id.Caption = 200 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmAccEditJournal.Fg_Journal
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmAccEditJournal.Fg_Journal_AfterEdit .Row, 4
            FrmAccEditJournal.Fg_Journal.ShowCell FrmAccEditJournal.Fg_Journal.Row, 10 '.Col
            FrmAccEditJournal.Fg_Journal.SetFocus
        End With
    End If
        If case_id.Caption = 220011 And Adodc1.Recordset.RecordCount >= 0 Then
            With FrmExpensesInvestment.GridInstallments
                .TextMatrix(.Row, .ColIndex("AccontName")) = Adodc1.Recordset.Fields!FirstName
                .TextMatrix(.Row, .ColIndex("AccontCode")) = (Adodc1.Recordset.Fields!Account_code)
            End With
        End If
        If case_id.Caption = 26112014 And Adodc1.Recordset.RecordCount >= 0 Then
            With FrmAccEditJournal4.Fg_Journal
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmAccEditJournal4.Fg_Journal_AfterEdit .Row, 8
            FrmAccEditJournal4.Fg_Journal.ShowCell FrmAccEditJournal4.Fg_Journal.Row, 10 '.Col
            FrmAccEditJournal4.Fg_Journal.SetFocus
        End With
       End If
        If case_id.Caption = 26112015 And Adodc1.Recordset.RecordCount >= 0 Then
           With FrmAccEditJournal4.Fg_Journal
            .TextMatrix(.Row, .ColIndex("Account_Serial2")) = Adodc1.Recordset.Fields!account_serial
            FrmAccEditJournal4.Fg_Journal_AfterEdit .Row, val(FrmAccEditJournal4.Fg_Journal.ColIndex("Account_Serial2"))
            FrmAccEditJournal4.Fg_Journal.ShowCell FrmAccEditJournal4.Fg_Journal.Row, FrmAccEditJournal4.Fg_Journal.ColIndex("AccountName2") '.Col
            FrmAccEditJournal4.Fg_Journal.SetFocus
        End With
    End If
    
    If case_id.Caption = 2220011 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmReturnExpensInves.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 177 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmAccountDestribution.DCAccountMaster.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 25102017 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmBoxesData.DboParentAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
        If case_id.Caption = 31219 And Adodc1.Recordset.RecordCount >= 0 Then
        Nationality.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    
    
    
        If case_id.Caption = 201 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmCitiesDistance.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
        If case_id.Caption = 203 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmDiscounts.DcboDebitSide.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
            If case_id.Caption = 204 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmDiscounts.DcboCreditSide.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    If case_id.Caption = 202 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmCitiesDistance.DcbAccount2.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 86 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmDocType.DcAccount1.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 87 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmDocType.DcAccount2.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
 
    

    
    If case_id.Caption = 89 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmDocType.DcAccount3.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 92 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmDocType.DcAccount4.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 91 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmDocType.DCAccount5.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 2116 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmShareholders.DboParentAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
        
    If case_id.Caption = 11116 And Adodc1.Recordset.RecordCount >= 0 Then
        Frminvestment.DboParentAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 20150204 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmIqarCompnent.DcAccountsus.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 778899 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmDestriEpensItemSearch.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
'***********************************
    If case_id.Caption = 251161 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmSocialInsurance.DcbAccount1.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    
   If case_id.Caption = 240219 And Adodc1.Recordset.RecordCount >= 0 Then
   
       
    End If
     
     
    
    If case_id.Caption = 251162 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmSocialInsurance.DcbAccount2.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
          
    If case_id.Caption = 251163 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmSocialInsurance.DcbAccount3.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 251164 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmSocialInsurance.DcbAccount4.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 654879 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmEmpDepartments.Account_code(mIndex).BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
      
   If case_id.Caption = 7897278 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmStoreData.Account_code1(mIndex).BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
   If case_id.Caption = 7897279 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmStoreData.Account_code2(mIndex).BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    


If case_id.Caption = 7897286 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmPermission.cmbAccounts.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
   If case_id.Caption = 7897280 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmStoreData.Account_code3(mIndex).BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
   If case_id.Caption = 7897281 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmStoreData.Account_code4(mIndex).BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
         
    
    '***********************************
    If case_id.Caption = 667788 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmDistriItemAccount.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 29121 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmPripaidExpenses.DBCboClientName.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 29123 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmContStudent.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 29122 And Adodc1.Recordset.RecordCount >= 0 Then
       FrmPripaidExpenses.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 50150221 And Adodc1.Recordset.RecordCount >= 0 Then
        MOFRAD.DCAccounts.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
   
    If case_id.Caption = 22915 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmAccountingReport.DCAccounts.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
   
    If case_id.Caption = 17815 And Adodc1.Recordset.RecordCount >= 0 Then
        MOFRAD.DCAccounts1.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 178 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmAccountDestribution.DCAccountDist.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 188 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmNewGard1.DcAccount1.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 189 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmNewGard1.DcAccount2.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 190 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmCashing.DCAccounts.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 260815 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmCashing.CommdiscountAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 260816 And Adodc1.Recordset.RecordCount >= 0 Then
    
        FrmCashing.TxtAccount = getAccountSerial_Code("Account_Serial", "Account_Code", (Adodc1.Recordset.Fields!Account_code))
        FrmCashing.DcbAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
 
    If case_id.Caption = 260817 And Adodc1.Recordset.RecordCount >= 0 Then
    
        FrmCashing.TxtCustCode = getAccountSerial_Code("Account_Serial", "Account_Code", (Adodc1.Recordset.Fields!Account_code))
        FrmCashing.DCAccounts.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    If case_id.Caption = 1200 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmCashing1.DCAccounts.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 191 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmOut.DCExtraAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 192 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmEmpSalary5.DCAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 193 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmPayments.DBCboClientName.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 1300 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmPayments2.DBCboClientName.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 194 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpenses40.DCAccounts1.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 195 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmProductionAllocation.DcAccount1.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 196 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmItemsClass.DboAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 2001 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmAccEditJournal1.Fg_Journal
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmAccEditJournal1.Fg_Journal_AfterEdit .Row, 4
            FrmAccEditJournal1.Fg_Journal.ShowCell FrmAccEditJournal1.Fg_Journal.Row, 10 '.Col
            FrmAccEditJournal1.Fg_Journal.SetFocus
        End With
    End If
    If case_id.Caption = 110815 Then
        Voucher_search.DCExtraAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
        Voucher_search.Text1(2).Text = Adodc1.Recordset.Fields!account_serial
    End If

    If case_id.Caption = 270815 Then
        Voucher_search1.DCExtraAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
        Voucher_search1.Text1(2).Text = Adodc1.Recordset.Fields!account_serial
    End If

    If case_id.Caption = 5300 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmAccEditJournal3.Fg_Journal
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmAccEditJournal3.Fg_Journal.ShowCell FrmAccEditJournal3.Fg_Journal.Row, 10 '.Col
            FrmAccEditJournal3.Fg_Journal.SetFocus
            FrmAccEditJournal3.Fg_Journal_AfterEdit .Row, 8
        End With
    End If
    
    If case_id.Caption = 80 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmExpenses3.VSFlexGrid1
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmExpenses3.VSFlexGrid1_AfterEdit .Row, 5
        End With
    End If
 
    If case_id.Caption = 50115 And Adodc1.Recordset.RecordCount >= 0 Then
        frmsalebill.DCExtraAccount.BoundText = (Adodc1.Recordset.Fields!Account_code)
    End If
    
    If case_id.Caption = 350350 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmExpenses301.VSFlexGrid1
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmExpenses301.VSFlexGrid1_AfterEdit .Row, .ColIndex("Account_Serial")
        End With
    End If
    
    If case_id.Caption = 191120141 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmCompositeAccounts.VSFlexGrid1
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            .TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!FirstName
            FrmCompositeAccounts.VSFlexGrid1_AfterEdit .Row, 8
        End With
    End If
    
    If case_id.Caption = 29112014 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmBalanceSheet.VSFlexGrid1
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            .TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!FirstName
            .TextMatrix(.Row, .ColIndex("Add")) = Adodc1.Recordset.Fields!FirstName
            .TextMatrix(.Row, .ColIndex("AccountCode")) = Adodc1.Recordset.Fields!Account_code
            'FrmBalanceSheet.VSFlexGrid1_AfterEdit .Row, 6
        End With
    End If
    
    If case_id.Caption = 350 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmExpenses30.VSFlexGrid1
            .TextMatrix(.Row, .ColIndex("AccountCode")) = Adodc1.Recordset.Fields!Account_code
            .TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!FirstName
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmExpenses30.VSFlexGrid1_AfterEdit .Row, 6
        End With
    End If
       
    If case_id.Caption = 350053 And Adodc1.Recordset.RecordCount >= 0 Then
        With RsExpenses.Fg_Journal
            .TextMatrix(.Row, .ColIndex("AccountCode")) = Adodc1.Recordset.Fields!Account_code
            .TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!FirstName
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            RsExpenses.Fg_Journal_AfterEdit .Row, 6
        End With
    End If
        If case_id.Caption = 350055 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmBillBuy.Fg_Journal
            .TextMatrix(.Row, .ColIndex("Accountcode2")) = Adodc1.Recordset.Fields!Account_code
            .TextMatrix(.Row, .ColIndex("Account_Name2")) = Adodc1.Recordset.Fields!FirstName
        End With
    End If
   If case_id.Caption = 350054 And Adodc1.Recordset.RecordCount >= 0 Then
        RsExpenses.DcbAccount.BoundText = Adodc1.Recordset.Fields!Account_code
    End If
    If case_id.Caption = 55 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpensesType.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Account_code
    End If
   
    If case_id.Caption = 161115 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpensesType.DboAcc.BoundText = Adodc1.Recordset.Fields!Account_code
       'FrmExpensesType.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Parent_Account_Code
    End If
    
    If case_id.Caption = 291115 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmViolationTypes.dcDiscAccount.BoundText = Adodc1.Recordset.Fields!Account_code
       'FrmExpensesType.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Parent_Account_Code
    End If
    
    If case_id.Caption = 16112 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmBanksData.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Account_code
       'FrmExpensesType.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Parent_Account_Code
    End If
    
    If case_id.Caption = 16915 And Adodc1.Recordset.RecordCount >= 0 Then
        FixedAssetsGroup.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Account_code
    End If
    
    If case_id.Caption = 251115 And Adodc1.Recordset.RecordCount >= 0 Then
        FixedAssetsGroup.DboParentAccount1.BoundText = Adodc1.Recordset.Fields!Account_code
    End If

    If case_id.Caption = 66 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmRevenuesTypes.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Account_code
    End If

    If case_id.Caption = 77 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmCustemers.DboParentAccount.Text = Adodc1.Recordset.Fields!FirstName
    End If

    If case_id.Caption = 88 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmCompany.DboParentAccount.Text = Adodc1.Recordset.Fields!FirstName
    End If
    
    If case_id.Caption = 880 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmOtherCustomers.DboParentAccount.Text = Adodc1.Recordset.Fields!FirstName
    End If
    
    If case_id.Caption = 666 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmBeforeInventoryK.AccDibDC.Text = Adodc1.Recordset.Fields!FirstName
    End If
    
        If case_id.Caption = 6660 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmVocationEntitlements.ADDACC.Text = Adodc1.Recordset.Fields!FirstName
    End If
    
       If case_id.Caption = 6661 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmVocationEntitlements.DISACC.Text = Adodc1.Recordset.Fields!FirstName
    End If
    
    
    
    If case_id.Caption = 888 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmBeforeInventoryK.AccCirDC.Text = Adodc1.Recordset.Fields!FirstName
    End If
    
    If case_id.Caption = 999 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmBeforeInventoryK.AccountsDC.Text = Adodc1.Recordset.Fields!FirstName
    End If
    If case_id.Caption = 889 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmIncomAndExpenReports.AccountsDC.BoundText = Adodc1.Recordset.Fields!Account_code
        FrmIncomAndExpenReports.Text1.Text = Adodc1.Recordset.Fields!account_serial
    End If
    If case_id.Caption = 1000 And Adodc1.Recordset.RecordCount >= 0 Then
        With FrmAccEditJournal1.Fg_Journal
            '.TextMatrix(.Row, .ColIndex("AccountName")) = Adodc1.Recordset.Fields!account_name
            .TextMatrix(.Row, .ColIndex("Account_Serial")) = Adodc1.Recordset.Fields!account_serial
            FrmAccEditJournal1.Fg_Journal_AfterEdit .Row, 2
        End With
    End If

    If case_id.Caption = 0 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmAccountCharts.Retrive (Adodc1.Recordset.Fields!Account_code)
    End If

    If case_id.Caption = 1 And Adodc1.Recordset.RecordCount >= 0 Then
        If SystemOptions.UserInterface = EnglishInterface Then
            FrmAccountingReport.Set_account_code Adodc1.Recordset.Fields!Account_code, Adodc1.Recordset.Fields!FirstName, Adodc1.Recordset.Fields!account_serial
        Else
            FrmAccountingReport.Set_account_code Adodc1.Recordset.Fields!Account_code, Adodc1.Recordset.Fields!FirstName, Adodc1.Recordset.Fields!account_serial
        End If
    End If

    If case_id.Caption = 2 And Adodc1.Recordset.RecordCount >= 0 Then
        frmsandat_kabd.DataCombo1.Text = Adodc1.Recordset.Fields!account_no
        frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!FirstName
    End If

    If case_id.Caption = 3 And Adodc1.Recordset.RecordCount >= 0 Then
        Voucher_search.Text1(2).Text = Adodc1.Recordset.Fields!account_serial
    End If

    If case_id.Caption = 40 And Adodc1.Recordset.RecordCount >= 0 Then
        mowazna.DataCombo1.Text = Adodc1.Recordset.Fields!account_serial
        mowazna.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 3 And Adodc1.Recordset.RecordCount >= 0 Then
        frmcustomer.DataCombo1.Text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 2122014 And Adodc1.Recordset.RecordCount >= 0 Then
        RSOwner.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Account_code
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If
    
       If case_id.Caption = 21220142 And Adodc1.Recordset.RecordCount >= 0 Then
        RsCustomers.DboParentAccount.BoundText = Adodc1.Recordset.Fields!Account_code
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If
        
    If case_id.Caption = 4 And Adodc1.Recordset.RecordCount >= 0 Then
        frmVendors.DataCombo1.Text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 5 And Adodc1.Recordset.RecordCount >= 0 Then
        rased_eftetahy_account.DataCombo1.Text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 6 And Adodc1.Recordset.RecordCount >= 0 Then
        frmmasrouf.DataCombo1.Text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 7 And Adodc1.Recordset.RecordCount >= 0 Then
        frmboxes.DataCombo1.Text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 8 And Adodc1.Recordset.RecordCount >= 0 Then
        frmrdrod.DataCombo1.Text = Adodc1.Recordset.Fields!account_no
        'frmsandat_kabd.Text2.Text = Adodc1.Recordset.Fields!account_name
    End If

    If case_id.Caption = 90 And Adodc1.Recordset.RecordCount >= 0 Then
        REPORTSFRM.DataCombo8.Text = Adodc1.Recordset.Fields!account_no
    End If

    If case_id.Caption = 700 And Adodc1.Recordset.RecordCount >= 0 Then
        baranches.DataCombo1(case_index).Text = Adodc1.Recordset.Fields!account_no
    End If

    If case_id.Caption = 201301 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpenses4.DCAccounts.BoundText = Adodc1.Recordset.Fields!Account_code
    End If

    If case_id.Caption = 201302 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpenses5.DCAccounts.BoundText = Adodc1.Recordset.Fields!Account_code
    End If

    If case_id.Caption = 20190719 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmExpenses301.DcbAccount.BoundText = Adodc1.Recordset.Fields!Account_code
    End If
    

    If case_id.Caption = 23072014 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmReceiptPart.DCAccounts.BoundText = Adodc1.Recordset.Fields!Account_code
    End If
    
    If case_id.Caption = 555 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmLC.DboParentAccount.Text = Adodc1.Recordset.Fields!FirstName
    End If
    
    If case_id.Caption = 2014110501 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmPaymentType.DcAccountsus.BoundText = Adodc1.Recordset.Fields!Account_code
    End If
    
    If case_id.Caption = 2014110502 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmPaymentType.DcAccountcom.BoundText = Adodc1.Recordset.Fields!Account_code
    End If
    
    If case_id.Caption = 9915 And Adodc1.Recordset.RecordCount >= 0 Then
        FrmTypeExchange.DCAccounts.BoundText = Adodc1.Recordset.Fields!Account_code
    End If
    
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
If SystemOptions.UserInterface = ArabicInterface Then
    X = MsgBox("  √ﬂÌœ «·€«¡ ﬂ· «· ÕœÌœ", vbCritical + vbYesNo)
 Else
 X = MsgBox("Confirm Cancell All Selection", vbCritical + vbYesNo)
End If

    If X = vbNo Then Exit Sub
        For i = 1 To Adodc2.Recordset.RecordCount
            Adodc2.Recordset.delete
            Adodc2.Recordset.MoveNext
        Next i
        DataGrid3.Refresh
        DataGrid4.Refresh
End Sub
Private Sub ALLButton4_Click()
    
    On Error Resume Next

    If Adodc2.Recordset.RecordCount = 0 Then Exit Sub
        X = MsgBox("Confirm PROCESS  √ﬂÌœ «·⁄„·Ì…", vbInformation + vbYesNo)
            If X = vbNo Then Exit Sub
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
                    frmsandat_ked.Text6.Text = frmsandat_ked.Text11.Text
                    frmsandat_ked.Adodc2.Recordset.update
                    frmsandat_ked.Adodc2.Recordset.Fields!Sanad_No = frmsandat_ked.Text1.Text
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
                    frmsandat_kabd.Text6.Text = frmsandat_kabd.Text11.Text
                    frmsandat_kabd.Adodc2.Recordset.update
                    frmsandat_kabd.Adodc2.Recordset.Fields!Sanad_No = frmsandat_kabd.Text1.Text
                    frmsandat_kabd.Adodc2.Recordset.Fields!sanad_source = "ÌœÊÌ"
                    frmsandat_kabd.Adodc2.Recordset.Fields!Date = DateValue(Now)
                    frmsandat_kabd.Adodc2.Recordset.Fields!BOX_name = frmsandat_kabd.DataCombo3.Text
                    frmsandat_kabd.Adodc2.Recordset.Fields!bona_3la = frmsandat_kabd.Text11.Text
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
                    frmsandat_sarf.Text6.Text = frmsandat_sarf.Text11.Text
                    frmsandat_sarf.Adodc2.Recordset.update
                    frmsandat_sarf.Adodc2.Recordset.Fields!Sanad_No = frmsandat_sarf.Text1.Text
                    frmsandat_sarf.Adodc2.Recordset.Fields!sanad_source = "ÌœÊÌ"
                    frmsandat_sarf.Adodc2.Recordset.Fields!Date = DateValue(Now)
                    frmsandat_sarf.Adodc2.Recordset.Fields!BOX_name = frmsandat_sarf.DataCombo3.Text
                    frmsandat_sarf.Adodc2.Recordset.Fields!bona_3la = frmsandat_sarf.Text11.Text
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
                    rased_eftetahy_account.Text6.Text = rased_eftetahy_account.Text11.Text
                    rased_eftetahy_account.Adodc2.Recordset.update
                    rased_eftetahy_account.Adodc2.Recordset.Fields!Sanad_No = rased_eftetahy_account.Text1.Text
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

Private Sub CboAccountCodeSearch_Change()
    Text1_KeyUp 0, 0, 0
End Sub
Private Sub CboAccountCodeSearch_Click()
    CboAccountCodeSearch_Change
End Sub
Private Sub CboNameSearch_Change()
    Text1_KeyUp 1, 1, 1
End Sub
Private Sub CboNameSearch_Click()
    CboNameSearch_Change
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
  

         ' Prompt user for desired author's last name
'         FindLastName = InputBox("Please enter the author's last name you
'                                  want to search for", "Find")

        ' Adodc1.Find "au_lname = '" & FindLastName & "'", , , 1

         ' Append your bookmark to the collection of selected rows
       '  DataGrid1.SelBookmarks.Add Adodc1.Recordset.Bookmark

  '  mClick = True
End Sub
Function hilightGrid()
    
    Dim title As String
    
    If Adodc1.Recordset.RecordCount < 1 Then Exit Function
    'Remove previous bookmarks.
    Do While DataGrid2.SelBookmarks.count > 0
        DataGrid2.SelBookmarks.Remove 0
    Loop
 
    DataGrid2.SelBookmarks.Add Adodc1.Recordset.Bookmark
End Function

Private Sub DataGrid2_GotFocus()
    hilightGrid
End Sub
Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ALLButton1_Click
    Else
        hilightGrid
    End If
End Sub

Private Sub DataGrid3_KeyUp(KeyCode As Integer, Shift As Integer)
    
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
Private Sub DataGrid4_KeyUp(KeyCode As Integer, Shift As Integer)
    
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
Exit Sub
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
 
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
 
    If SystemOptions.UserInterface = ArabicInterface Then
        sql = "   SELECT  ACCOUNTS.Account_Code ,ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName,Account_ID  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
    Else
        sql = "   SELECT  ACCOUNTS.Account_Code ,ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_NameEng As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName,Account_ID  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
    End If

    If Me.case_id = 1700 Then
        sql = sql & " and ACCOUNTS.last_account=1"
    End If
                          
    If Me.case_id = 16112 Or Me.case_id = 700 Or Me.case_id = 55 Or Me.case_id = 16915 Or Me.case_id = 555 Or Me.case_id = 66 Or Me.case_id = 77 Or Me.case_id = 88 Or Me.case_id = 2116 Or Me.case_id = 11116 Then
        sql = sql & " and ACCOUNTS.last_account=0"
    End If
       sql = sql & GetAccountByBarnchUser
       sql = sql & GetAccountCodeHiding
    If SystemOptions.UserInterface = EnglishInterface Then
        sql = sql & "Order By ACCOUNTS.Account_NameENG"
    Else
        sql = sql & "Order By ACCOUNTS.Account_Name"
    End If
                             
    Adodc1.RecordSource = sql
    Adodc1.Refresh
 
End Sub
Private Sub Form_Load()
    
    On Error Resume Next
    
    connection_string = Cn.ConnectionString

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText

    If SystemOptions.UserInterface = ArabicInterface Then
        sql = "   SELECT  ACCOUNTS.Account_Code ,ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName ,ACCOUNTS.Account_ID  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code ='x'  "
        sql = sql & GetAccountByBarnchUser
        sql = sql & GetAccountCodeHiding
        sql = sql & " Order By ACCOUNTS.Account_Name"
    Else
        sql = "   SELECT  ACCOUNTS.Account_Code ,ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_NameEng As FirstName,ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName,ACCOUNTS.Account_ID   FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code = 'X'    "
        sql = sql & GetAccountByBarnchUser
        sql = sql & GetAccountCodeHiding
        sql = sql & "  Order By ACCOUNTS.Account_NameEng"
    End If

    'Sql = "select * from accounts where last_account=1"
    Adodc1.RecordSource = sql
    Adodc1.Refresh

    '
    If SystemOptions.UserInterface = EnglishInterface Then
        With Me.CboAccountCodeSearch
            .Clear
            .AddItem "Typical Search"
            .AddItem "From The Start"
            .AddItem "From The End"
            .AddItem "Any Where"
        End With
        
        With Me.CboNameSearch
            .Clear
            .AddItem "Typical Search"
            .AddItem "From The Start"
            .AddItem "From The End"
            .AddItem "Any Where"
        End With
    Else
        With Me.CboAccountCodeSearch
            .Clear
            .AddItem "»ÕÀ „ÿ«»ﬁ"
            .AddItem "»ÕÀ „‰ «·»œ«Ì…"
            .AddItem "»ÕÀ „‰ «·‰Â«Ì…"
            .AddItem "»ÕÀ ›Ï «Ï „ﬂ«‰"
        End With
        
        With Me.CboNameSearch
            .Clear
            .AddItem "»ÕÀ „ÿ«»ﬁ"
            .AddItem "»ÕÀ „‰ «·»œ«Ì…"
            .AddItem "»ÕÀ „‰ «·‰Â«Ì…"
            .AddItem "»ÕÀ ›Ï «Ï „ﬂ«‰"
        End With
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        Label6.Caption = "ACC.Code"
        Label1.Caption = "ACC. Name"
        Me.Caption = "Accounts Search"

        DataGrid5.Visible = True
        DataGrid2.Visible = False
        Frame4.Caption = "Search By  "
        OpTcode(0).value = "Acc"
        OpTcode(1).value = "Cus."
        OpTcode(2).value = "Supp"
        OpTcode(3).value = "Emp."
        Label4.Caption = "Press Enter to Select Account"
        OpTcode(0).Caption = "Account"
        OpTcode(1).Caption = "Customer"
        OpTcode(2).Caption = "Supplier"
        OpTcode(3).Caption = "Employee"
        lbl(11).Caption = "Match Type"
        lbl(8).Caption = "Match Type"
    Else
        DataGrid2.Visible = True
        DataGrid5.Visible = False
    End If
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
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
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    Dim autosearch As Boolean
    autosearch = False
    
    If KeyCode = vbKeyF6 Then
        'account_index.Show
    End If
    mClick = False
     
     If KeyCode = vbKeyReturn Then
                                        If Adodc1.Recordset.RecordCount = 1 Then
            DataGrid2.SetFocus
            ALLButton1_Click
        Else
          DataGrid2.SetFocus
        End If
                   End If
                   
                   
    
    If True = True Then

    
    'If KeyCode = 13 Then
    ' last_account=1 and

    Dim CusName  As String
    Dim emp_name  As String
    
    
    If Index = 0 Then
        ' sql = "   SELECT  ACCOUNTS.Account_Code ,ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'    AND ACCOUNTS.Account_Serial like '%" & Text1(Index).text & "%' Order By ACCOUNTS.Account_Name"
        If OpTcode(0).value = True Then
            sql = "   SELECT  ACCOUNTS.Account_Code ,ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName  ,ACCOUNTS.Account_ID FROM (ACCOUNTS  LEFT OUTER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)   LEFT OUTER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'   "
            If Me.CboAccountCodeSearch.ListIndex = 0 Then
                sql = sql & " AND ACCOUNTS.Account_Serial = '" & Text1(Index).Text & "' "
            ElseIf Me.CboAccountCodeSearch.ListIndex = 1 Then
                sql = sql & " AND ACCOUNTS.Account_Serial like '" & Text1(Index).Text & "%' "
            ElseIf Me.CboAccountCodeSearch.ListIndex = 2 Then
                sql = sql & " AND ACCOUNTS.Account_Serial like '%" & Text1(Index).Text & "' "
            ElseIf Me.CboAccountCodeSearch.ListIndex = 3 Then
                sql = sql & " AND ACCOUNTS.Account_Serial like '%" & Text1(Index).Text & "%' "
            ElseIf Me.CboAccountCodeSearch.ListIndex = -1 Then
                sql = sql & " AND ACCOUNTS.Account_Serial like '%" & Text1(Index).Text & "%' "
            End If
        ElseIf OpTcode(1).value = True Then
            GetCustomersDetail , , Text1(0).Text, 1, , , CusName
            If CusName = "" Then CusName = "?xc????"
                If SystemOptions.UserInterface = EnglishInterface Then
                    sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_NameENG As FirstName,ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName,ACCOUNTS.Account_ID   FROM (ACCOUNTS  LEFT OUTER JOIN  ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)  LEFT OUTER JOIN   ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'  "
                    'AND ACCOUNTS.Account_NameENG like'%" & CusName & "%'"
                    If Me.CboAccountCodeSearch.ListIndex = 0 Then
                        sql = sql & " AND ACCOUNTS.Account_NameENG = '" & CusName & "' "
                    ElseIf Me.CboAccountCodeSearch.ListIndex = 1 Then
                        sql = sql & " AND ACCOUNTS.Account_NameENG like '" & CusName & "%' "
                    ElseIf Me.CboAccountCodeSearch.ListIndex = 2 Then
                        sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & CusName & "' "
                    ElseIf Me.CboAccountCodeSearch.ListIndex = 3 Then
                        sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & CusName & "%' "
                    ElseIf Me.CboAccountCodeSearch.ListIndex = -1 Then
                        sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & CusName & "%' "
                    End If
                Else
                    sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName,ACCOUNTS.Account_ID   FROM (ACCOUNTS  LEFT OUTER JOIN  ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)  LEFT OUTER JOIN   ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                    'AND ACCOUNTS.Account_Name like'%" & CusName & "%' "
                    If Me.CboAccountCodeSearch.ListIndex = 0 Then
                        sql = sql & " AND ACCOUNTS.Account_Name = '" & CusName & "' "
                    ElseIf Me.CboAccountCodeSearch.ListIndex = 1 Then
                        sql = sql & " AND ACCOUNTS.Account_Name like '" & CusName & "%' "
                    ElseIf Me.CboAccountCodeSearch.ListIndex = 2 Then
                        sql = sql & " AND ACCOUNTS.Account_Name like '%" & CusName & "' "
                    ElseIf Me.CboAccountCodeSearch.ListIndex = 3 Then
                        sql = sql & " AND ACCOUNTS.Account_Name like '%" & CusName & "%' "
                    ElseIf Me.CboAccountCodeSearch.ListIndex = -1 Then
                        sql = sql & " AND ACCOUNTS.Account_Name like '%" & CusName & "%' "
                    End If
                End If
            ElseIf OpTcode(2).value = True Then
                GetCustomersDetail , , Text1(0).Text, 2, , , CusName
                If CusName = "" Then CusName = "?xc????"
                    If SystemOptions.UserInterface = EnglishInterface Then
                        sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_NameENG As FirstName,ACCOUNTS_1.Account_NameENG As ParentName, ACCOUNTS_2.Account_NameENG As RootName,ACCOUNTS.Account_ID   FROM (ACCOUNTS  LEFT OUTER JOIN  ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)  LEFT OUTER JOIN   ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'  "
                        'AND ACCOUNTS.Account_NameENG like'%" & CusName & "%' "
                        If Me.CboAccountCodeSearch.ListIndex = 0 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG = '" & CusName & "' "
                        ElseIf Me.CboAccountCodeSearch.ListIndex = 1 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG like '" & CusName & "%' "
                        ElseIf Me.CboAccountCodeSearch.ListIndex = 2 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & CusName & "' "
                        ElseIf Me.CboAccountCodeSearch.ListIndex = 3 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & CusName & "%' "
                        ElseIf Me.CboAccountCodeSearch.ListIndex = -1 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & CusName & "%' "
                        End If
                    Else
                        sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName ,ACCOUNTS.Account_ID  FROM (ACCOUNTS  LEFT OUTER JOIN  ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)  LEFT OUTER JOIN   ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                        'AND ACCOUNTS.Account_Name like'%" & CusName & "%' "
                        If Me.CboAccountCodeSearch.ListIndex = 0 Then
                            sql = sql & " AND ACCOUNTS.Account_Name = '" & CusName & "' "
                        ElseIf Me.CboAccountCodeSearch.ListIndex = 1 Then
                            sql = sql & " AND ACCOUNTS.Account_Name like '" & CusName & "%' "
                        ElseIf Me.CboAccountCodeSearch.ListIndex = 2 Then
                            sql = sql & " AND ACCOUNTS.Account_Name like '%" & CusName & "' "
                        ElseIf Me.CboAccountCodeSearch.ListIndex = 3 Then
                            sql = sql & " AND ACCOUNTS.Account_Name like '%" & CusName & "%' "
                        ElseIf Me.CboAccountCodeSearch.ListIndex = -1 Then
                            sql = sql & " AND ACCOUNTS.Account_Name like '%" & CusName & "%' "
                        End If
                    End If
                ElseIf OpTcode(3).value = True Then
                    GetEmployeeIDFromCode Text1(0).Text, , , , emp_name
                    emp_name = Trim(emp_name)
                    If emp_name = "" Then emp_name = "?xc????"
                        If SystemOptions.UserInterface = EnglishInterface Then
                            sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_NameENG As FirstName,ACCOUNTS_1.Account_NameENG As ParentName, ACCOUNTS_2.Account_NameENG As RootName ,ACCOUNTS.Account_ID  FROM (ACCOUNTS  LEFT OUTER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)  LEFT OUTER JOIN   ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'  "
                            'AND ACCOUNTS.Account_NameENG like'%" & emp_name & "%' "
                            If Me.CboAccountCodeSearch.ListIndex = 0 Then
                                sql = sql & " AND ACCOUNTS.Account_NameENG = '" & emp_name & "' "
                            ElseIf Me.CboAccountCodeSearch.ListIndex = 1 Then
                                sql = sql & " AND ACCOUNTS.Account_NameENG like '" & emp_name & "%' "
                            ElseIf Me.CboAccountCodeSearch.ListIndex = 2 Then
                                sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & emp_name & "' "
                            ElseIf Me.CboAccountCodeSearch.ListIndex = 3 Then
                                sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & emp_name & "%' "
                            ElseIf Me.CboAccountCodeSearch.ListIndex = -1 Then
                                sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & emp_name & "%' "
                            End If
                        Else
                            sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName ,ACCOUNTS.Account_ID  FROM (ACCOUNTS  LEFT OUTER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)  LEFT OUTER JOIN   ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                            'AND ACCOUNTS.Account_Name like'%" & emp_name & "%' "
                            If Me.CboAccountCodeSearch.ListIndex = 0 Then
                                sql = sql & " AND ACCOUNTS.Account_Name = '" & emp_name & "' "
                            ElseIf Me.CboAccountCodeSearch.ListIndex = 1 Then
                                sql = sql & " AND ACCOUNTS.Account_Name like '" & emp_name & "%' "
                            ElseIf Me.CboAccountCodeSearch.ListIndex = 2 Then
                                sql = sql & " AND ACCOUNTS.Account_Name like '%" & emp_name & "' "
                            ElseIf Me.CboAccountCodeSearch.ListIndex = 3 Then
                                sql = sql & " AND ACCOUNTS.Account_Name like '%" & emp_name & "%' "
                            ElseIf Me.CboAccountCodeSearch.ListIndex = -1 Then
                                sql = sql & " AND ACCOUNTS.Account_Name like '%" & emp_name & "%' "
                            End If
                        End If
                    End If
                    'Sql = "select * from accounts where Account_Serial like '%" & Text1(Index).text & "%'"
                    If Text1(0).Text = "" Then
                        sql = "   SELECT  ACCOUNTS.Account_Code ,ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName ,ACCOUNTS.Account_ID  FROM (ACCOUNTS  LEFT OUTER JOIN  ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)   LEFT OUTER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'   "
                        sql = sql & "AND ACCOUNTS.Account_Serial like '%X" & Text1(Index).Text & "%' " 'Order By ACCOUNTS.Account_Name
                    End If
                End If
                If Index = 1 Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        'sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_NameENG As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'  " & "   AND ACCOUNTS.Account_NameENG like'%" & Text1(Index).text & "%' "
                        sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_NameENG As FirstName,ACCOUNTS_1.Account_NameENG As ParentName, ACCOUNTS_2.Account_NameENG As RootName,ACCOUNTS.Account_ID   FROM (ACCOUNTS  LEFT OUTER JOIN  ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)  LEFT OUTER JOIN   ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'  "
                        'AND ACCOUNTS.Account_NameENG like'%" & Text1(Index).Text & "%' "
                        If Me.CboNameSearch.ListIndex = 0 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG = '" & Text1(Index).Text & "' "
                        ElseIf Me.CboNameSearch.ListIndex = 1 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG like '" & Text1(Index).Text & "%' "
                        ElseIf Me.CboNameSearch.ListIndex = 2 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & Text1(Index).Text & "' "
                        ElseIf Me.CboNameSearch.ListIndex = 3 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & Text1(Index).Text & "%' "
                        ElseIf Me.CboNameSearch.ListIndex = -1 Then
                            sql = sql & " AND ACCOUNTS.Account_NameENG like '%" & Text1(Index).Text & "%' "
                        End If
                        'Sql = "select * from accounts where   Account_NameENG like'%" & Text1(Index).text & "%'"
                    Else
                        'sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' " & "   AND ACCOUNTS.Account_Name like'%" & Text1(Index).text & "%' "
                        sql = "   SELECT ACCOUNTS.Account_Code , ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName ,ACCOUNTS.Account_ID  FROM (ACCOUNTS  LEFT OUTER JOIN  ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)  LEFT OUTER JOIN   ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                        'AND ACCOUNTS.Account_Name like'%" & Text1(Index).Text & "%' "
                        If Me.CboNameSearch.ListIndex = 0 Then
                            sql = sql & " AND ACCOUNTS.Account_Name = '" & Text1(Index).Text & "' "
                        ElseIf Me.CboNameSearch.ListIndex = 1 Then
                            sql = sql & " AND ACCOUNTS.Account_Name like '" & Text1(Index).Text & "%' "
                        ElseIf Me.CboNameSearch.ListIndex = 2 Then
                            sql = sql & " AND ACCOUNTS.Account_Name like '%" & Text1(Index).Text & "' "
                        ElseIf Me.CboNameSearch.ListIndex = 3 Then
                            sql = sql & " AND ACCOUNTS.Account_Name like '%" & Text1(Index).Text & "%' "
                        ElseIf Me.CboNameSearch.ListIndex = -1 Then
                            sql = sql & " AND ACCOUNTS.Account_Name like '%" & Text1(Index).Text & "%' "
                        End If
                        'Sql = "select * from accounts where   Account_Name like'%" & Text1(Index).text & "%'"
                    End If

                    If Text1(1).Text = "" Then
                        sql = "SELECT  ACCOUNTS.Account_Code ,ACCOUNTS.Account_SERIAL, ACCOUNTS.Account_Name As FirstName,ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName ,ACCOUNTS.Account_ID  FROM (ACCOUNTS  LEFT OUTER JOIN  ACCOUNTS AS ACCOUNTS_1 ON ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code)   LEFT OUTER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r'    AND ACCOUNTS.Account_Serial like '%X" & Text1(Index).Text & "%' " 'Order By ACCOUNTS.Account_Name
                    End If
                End If
  If Me.case_id = 16112 Or Me.case_id = 291115 Or Me.case_id = 22915 Or Me.case_id = 110815 Or Me.case_id = 778899 Or Me.case_id = 667788 Or Me.case_id = 20150204 Or Me.case_id = 26112014 Or Me.case_id = 26112015 Or Me.case_id = 2001 Or Me.case_id = 350 Or Me.case_id = 200 Or Me.case_id = 201301 Or Me.case_id = 201302 Or Me.case_id = 1700 Or Me.case_id = 177 Or Me.case_id = 178 Or Me.case_id = 188 Or Me.case_id = 189 Or Me.case_id = 190 Or Me.case_id = 1200 Or Me.case_id = 191 Or Me.case_id = 192 Or Me.case_id = 23072014 Or Me.case_id = 2014110501 Or Me.case_id = 2014110502 Or Me.case_id = 50150221 Or Me.case_id = 17815 Or Me.case_id = 80 Or Me.case_id = 350053 Or Me.case_id = 350054 Or Me.case_id = 6660 Or Me.case_id = 6661 Or Me.case_id = 203 Or Me.case_id = 204 Or Me.case_id = 190 Or Me.case_id = 260816 Or Me.case_id = 260817 Or Me.case_id = 654879 Or Me.case_id = 240219 Or Me.case_id = 260219 Or Me.case_id = 2602191 Or Me.case_id = 90519 Or Me.case_id = 10519 Or Me.case_id = 20190719 Or _
  Me.case_id = 31219 Then

                    sql = sql & " and ACCOUNTS.last_account=1"
                End If
                  
                If Me.case_id = 25102017 Or Me.case_id = 789725 Or Me.case_id = 16112 Or Me.case_id = 700 Or Me.case_id = 55 Or Me.case_id = 66 Or Me.case_id = 77 Or Me.case_id = 88 Then
                    sql = sql & " and ACCOUNTS.last_account=0"
                End If
                
                sql = sql & GetAccountByBarnchUser
                sql = sql & GetAccountCodeHiding
                
                If SystemOptions.UserInterface = EnglishInterface Then
                    sql = sql & "Order By ACCOUNTS.Account_Serial , ACCOUNTS.Account_NameENG"
                Else
                    sql = sql & "Order By ACCOUNTS.Account_Serial,ACCOUNTS.Account_Name"
                End If
                
                Adodc1.RecordSource = sql
                
                
                
                Adodc1.Refresh
                      
                      

    End If

                      
            '    If Adodc1.Recordset.RecordCount = 0 Then
                    'MsgBox "not fount  ·«ÌÊÃœ ‰ «∆Ã ··»ÕÀ", vbInformation
            '    End If
            'End If
End Sub
