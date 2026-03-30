VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form items_search2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… »ÕÀ «’‰«›"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   12150
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   840
      TabIndex        =   19
      Top             =   0
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   10680
      TabIndex        =   18
      Top             =   0
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "items_search2.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   10215
      _ExtentX        =   18018
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "item_code"
         Caption         =   " ﬂÊœ «·’‰›"
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
         DataField       =   "items_name"
         Caption         =   "«”„ «·’‰›"
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
         DataField       =   "part_no"
         Caption         =   "—ﬁ„ «·ﬁÿ⁄…"
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
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4995.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "items_search2.frx":0015
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   2160
      Width           =   12015
      _ExtentX        =   21193
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "ItemCode"
         Caption         =   "ﬂÊœ «·’‰›"
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
         DataField       =   "ItemName"
         Caption         =   "«”„ «·’‰›"
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
         DataField       =   "ItemID"
         Caption         =   "—ﬁ„ «·ﬁÿ⁄…"
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
         Caption         =   "”⁄— «·»Ì⁄"
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
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4995.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1484.787
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
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
      MICON           =   "items_search2.frx":002A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   3720
      Top             =   4440
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
      Bindings        =   "items_search2.frx":0046
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "fullcode"
         Caption         =   "code"
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
         DataField       =   "items_name"
         Caption         =   "items_name"
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
         DataField       =   "part_no"
         Caption         =   "part_no"
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
         DataField       =   "had_eltalab"
         Caption         =   "had_eltalab"
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
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4995.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1484.787
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
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "«œ—«Ã «·ﬂ·"
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
      MICON           =   "items_search2.frx":005B
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
      Height          =   585
      Left            =   5040
      Top             =   7920
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
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
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
      MICON           =   "items_search2.frx":0077
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
      Left            =   2040
      TabIndex        =   7
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "«·€«¡  «·ﬂ·"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "items_search2.frx":0093
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "items_search2.frx":00AF
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "item_code"
         Caption         =   "code"
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
         DataField       =   "items_name"
         Caption         =   "items_name"
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
         DataField       =   "part_no"
         Caption         =   "part_no"
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
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4995.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc6 
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
   Begin MSAdodcLib.Adodc Adodc4 
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
   Begin MSAdodcLib.Adodc Adodc12 
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   585
      Left            =   0
      Top             =   720
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
   Begin MSAdodcLib.Adodc Adodc5 
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
   Begin ALLButtonS.ALLButton ALLButton5 
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
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
      MICON           =   "items_search2.frx":00C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc numbering 
      Height          =   585
      Left            =   10320
      Top             =   1800
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin MSAdodcLib.Adodc detect_no 
      Height          =   585
      Left            =   10200
      Top             =   1920
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin MSAdodcLib.Adodc Adodc14 
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
   Begin MSAdodcLib.Adodc Adodc8 
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
   Begin MSAdodcLib.Adodc Adodc7 
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
   Begin MSAdodcLib.Adodc Adodc9 
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
   Begin MSAdodcLib.Adodc Adodc10 
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
   Begin MSAdodcLib.Adodc Adodc11 
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
   Begin VB.Label Label5 
      Caption         =   "—ﬁ„ «·ﬁÿ⁄…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   8520
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Part no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1080
      TabIndex        =   16
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label case_id 
      Caption         =   "0"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "«”„ «·’‰›"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   8520
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "ﬂÊœ «·’‰›"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   8400
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Item no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1080
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "items_search2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim auto_sanad_no As String
Dim numbering_type As Integer

Private Sub ALLButton1_Click()

    On Error Resume Next

    If case_id.Caption = 0 And Adodc1.Recordset.RecordCount >= 0 Then 'ok items
        '        Sql = "select * from items  where  branch_no=" & branch_no & " and      departement='" & departement_name & "' and   blocked=0 and  fullcode='" & Adodc1.Recordset.Fields!fullcode & "'"
        'FrmSaleBill.DCboItemsCode.text = Adodc1.Recordset.Fields!ItemCode
        'FrmItems.Adodc1.RecordSource = Sql
        'FrmItems.Adodc1.Refresh
        'InvoiceScreen.DCboItemCode.text = Adodc1.Recordset.Fields!ItemCode
        Unload items_search2

    End If

    If (case_id.Caption = 1 Or case_id.Caption = 5 Or case_id.Caption = 50 Or case_id.Caption = 60) And Adodc1.Recordset.RecordCount >= 0 Then

        Adodc2.Recordset.AddNew
        Adodc2.Recordset.Fields!item_code = Adodc1.Recordset.Fields!fullcode
        Adodc2.Recordset.Fields!items_name = Adodc1.Recordset.Fields!items_name
        Adodc2.Recordset.Fields!part_no = Adodc1.Recordset.Fields!part_no
  
        Adodc2.Recordset.update
  
        Adodc2.Refresh
        DataGrid3.Refresh
        DataGrid4.Refresh

    End If

    If case_id.Caption = 4 And Adodc1.Recordset.RecordCount >= 0 Then 'ok items units
        'Sql = "select * from items_units where fullcode='" & Adodc1.Recordset.Fields!fullcode & "'"
        'items_units.Adodc1.RecordSource = Sql
        items_units.Text2.text = Adodc1.Recordset.Fields!fullcode
        items_units.Text3.text = Adodc1.Recordset.Fields!items_name

        items_units.Adodc3.CommandType = adCmdText
        items_units.Adodc3.RecordSource = "select * from  items_units where item_code='" & Adodc1.Recordset.Fields!fullcode & "'"
        items_units.Adodc3.Refresh
        items_units.DataGrid1.Refresh

        Unload Me
    End If
  
    'If case_id.Caption = 5 And Adodc1.Recordset.RecordCount >= 0 Then
 
    'RASED_EFTETAHY.DataCombo1.Text = Adodc1.Recordset.Fields!items_name
    'RASED_EFTETAHY.Adodc3.Refresh
    'Unload Me
    'End If
    
End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next

    x = MsgBox("Confirm PROCESS  √ﬂÌœ «·⁄„·Ì…", vbInformation + vbYesNo)

    If x = vbNo Then Exit Sub

    If case_id.Caption = 60 Then ' SAND estlam

        Adodc5.CommandType = adCmdText
        Adodc5.RecordSource = "select * from  inventories where branch_no=" & Branch_NO & " and departement='" & departement_name & "'  and  master = 1 and   not (inventory_name='')"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount = 0 Then
            If my_language = "E" Then
                MsgBox "no primary stores defined operation will be cancelled", vbCritical
            Else
                MsgBox "·« ÌÊÃœ „Œ«“‰ «”«”Ì… ·œÌﬂ ”Ì „ «·€«¡ «·⁄„·Ì…", vbCritical
            End If

            Exit Sub

        End If

        Adodc10.CommandType = adCmdText
        Adodc10.RecordSource = "select * from  vendors where branch_no=" & Branch_NO & " and departement='" & departement_name & "'  and  master = 1 and   not (vendor_name='')"
        Adodc10.Refresh

        If Adodc10.Recordset.RecordCount = 0 Then
            If my_language = "E" Then
                MsgBox "no primary Vendors defined operation will be cancelled", vbCritical
            Else
                MsgBox "·« ÌÊÃœ „Ê—œ «”«”Ì ·œÌﬂ ”Ì „ «·€«¡ «·⁄„·Ì…", vbCritical
            End If

            Exit Sub

        End If

        If sand_ESTLAM_inventory.Text3.text = "" Then
            auto_sanad_no = sand_ESTLAM_inventory.sand_numbering()
        Else
            auto_sanad_no = sand_ESTLAM_inventory.Text3.text
            'Adodc1.Recordset.Fields!sanad_no = Adodc1.Recordset.Fields!sandat_pc_no
        End If

        If auto_sanad_no <> "" Then
         
        Else
          
            If my_language = "E" Then
          
                MsgBox "can't save define numbering method first in system manger", vbCritical: Exit Sub
            Else
                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ·«»œ „‰ «œŒ«· —ﬁ„ ··”‰œ «Ê·« ·«‰ﬂ «Œ —   —ﬁÌ„ ”‰œ«  ÌœÊÌ", vbCritical: Exit Sub
            End If
          
        End If
        
        If sand_ESTLAM_inventory.DataCombo4.text = "" Then
            If my_language = "E" Then
          
                MsgBox "can't save define Voucher Type first", vbCritical: Exit Sub
            Else
                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ·«»œ „‰  ÕœÌœ ‰Ê⁄  ··”‰œ  ", vbCritical: Exit Sub
            End If

            Exit Sub
        End If
        
        If sand_ESTLAM_inventory.Text2.text = "" Then
            If my_language = "E" Then
          
                MsgBox "can't save define Based ON    first", vbCritical: Exit Sub
            Else
                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ·«»œ „‰  ÕœÌœ »‰«¡ ⁄·Ïø  ", vbCritical: Exit Sub
            End If

            Exit Sub
        End If
        
        numbering.ConnectionString = connection_string
        numbering.CommandType = adCmdText
        numbering.RecordSource = "select * from sanad_numbering where branch_no=" & Branch_NO & " and departement='" & departement_name & "' and sanad_no=5"
        numbering.Refresh

        If numbering.Recordset.RecordCount = 0 Then
            numbering_type = 0
        Else
            numbering_type = numbering.Recordset.Fields!numbering_id
        End If
        
        For i = 1 To Adodc2.Recordset.RecordCount

            Adodc3.CommandType = adCmdText
            Adodc3.RecordSource = "select * from  items_units where master = 1 and  item_code='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc3.Refresh

            If Adodc3.Recordset.RecordCount > 0 Then
                Adodc3.Recordset.MoveFirst
            Else

                If my_language = "E" Then
                    MsgBox "ther is no primary unit to item code : " & Adodc2.Recordset.Fields!item_code & " and name is  " & Adodc2.Recordset.Fields!items_name, vbCritical: GoTo HX
                Else
                    MsgBox "·« ÌÊÃœ ÊÕœÂ «”«”Ì… ···’‰› —ﬁ„ " & Adodc2.Recordset.Fields!item_code & " Ê«”„Â  " & Adodc2.Recordset.Fields!items_name, vbCritical: GoTo HX
                End If

            End If

            'Command1(1).Enabled = False
            'Command1(0).Enabled = True
 
            'txt_item_total.Text = Val(txt_rased.Text) * Val(txt_item_by_unit)

            sand_ESTLAM_inventory.Adodc1.Recordset.AddNew
            
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!vendor_id = Adodc10.Recordset.Fields!fullcode
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!vendor = Adodc10.Recordset.Fields!vendor_name

            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!Sanad_No = auto_sanad_no
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!Branch_NO = Branch_NO
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!Departement = departement_name
              
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!user_name = current_user_name

            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!numbering_type = numbering_type 'TRIC
                  
            If numbering_type = 2 Then
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!sanad_year = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!sanad_month = Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            End If

            If numbering_type = 3 Then
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!sanad_year = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
            
            End If

            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!Transaction_Type = "”‰œ «” ·«„"
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!SANAD_TYPE = sand_ESTLAM_inventory.DataCombo4.text '"”‰œ ’—› «·Ì €Ì— „Õœœ"
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!bona_3la = sand_ESTLAM_inventory.Text2.text ' "”‰œ ’—› «·Ì" 'Text2.Text
         
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!Transaction_Date = DateValue(Now)
         
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!item_code = Adodc2.Recordset.Fields!item_code
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!items_name = Adodc2.Recordset.Fields!items_name
         
            Adodc8.RecordSource = "select * from  items where branch_no=" & Branch_NO & " and departement='" & departement_name & "' and   blocked=0 AND   fullcode='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc8.Refresh
                  
            ' Adodc10.RecordSource = "select * from  items where branch_no=" & branch_no & " and departement='" & departement_name & "' and   blocked=0 AND  item_code='" & Adodc2.Recordset.Fields!item_code & "'"
            ' Adodc10.Refresh

            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!part_no = Adodc8.Recordset.Fields!part_no
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!item_unit = Adodc3.Recordset.Fields!Unit_name
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!items_no_by_one = Adodc3.Recordset.Fields!unit_value 'Val(frm_Amr_shogl.Text6.Text) * Val(frm_Amr_shogl.DataCombo10.BoundText)
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!ta2ther_makhzan = Adodc3.Recordset.Fields!unit_value   'Val(frm_Amr_shogl.Text6.Text) * Val(frm_Amr_shogl.DataCombo10.BoundText) ' Val(frm_Amr_shogl.Text6.Text) * -1 'Val(frm_Amr_shogl.Adodc8.Recordset.Fields!items_no_by_one) * -1
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!Qty = 1 'Val(frm_Amr_shogl.Text6.Text)
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!inventory_name = Adodc5.Recordset.Fields!inventory_name
            sand_ESTLAM_inventory.Adodc1.Recordset.Fields!inventory_id = Adodc5.Recordset.Fields!fullcode

            '«Œ— ”⁄— ‘—«¡ Ê «· ﬂ·›…
            Adodc12.CommandType = adCmdText
            Adodc12.RecordSource = "select * from  items where fullcode='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc12.Refresh

            Adodc11.CommandType = adCmdText
            Adodc11.RecordSource = "select sum(ta2ther_makhzan) as total_qty from  inventory where item_code='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc11.Refresh
         
            If Adodc12.Recordset.RecordCount > 0 And Not IsNull(Adodc12.Recordset.Fields!motwaset_taklefa) And Not IsNull(Adodc11.Recordset.Fields!total_qty) Then
 
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!average_cost = (Adodc11.Recordset.Fields!total_qty * Adodc12.Recordset.Fields!motwaset_taklefa + Adodc3.Recordset.Fields!unit_value * Adodc12.Recordset.Fields!akher_s3r_shera) / (Adodc11.Recordset.Fields!total_qty + Adodc3.Recordset.Fields!unit_value)
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!total_price = sand_ESTLAM_inventory.Adodc1.Recordset.Fields!average_cost * Adodc3.Recordset.Fields!unit_value
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!last_price = Adodc12.Recordset.Fields!akher_s3r_shera
 
            Else
 
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!average_cost = 0
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!total_price = 0
                sand_ESTLAM_inventory.Adodc1.Recordset.Fields!last_price = 0

            End If

            Adodc12.Recordset.Fields!motwaset_taklefa = sand_ESTLAM_inventory.Adodc1.Recordset.Fields!average_cost
            Adodc12.Recordset.update

            sand_ESTLAM_inventory.Adodc1.Recordset.update
            ' Adodc1.Recordset.MoveLast
HX:
            Adodc2.Recordset.MoveNext

        Next i

        sqlx = "select * from inventory where branch_no=" & Branch_NO & " and transaction_type='" & "”‰œ «” ·«„" & "' and  sanad_no= '" & auto_sanad_no & "'"
        sand_ESTLAM_inventory.Adodc1.RecordSource = sqlx

        sand_ESTLAM_inventory.Adodc1.Refresh
        sand_ESTLAM_inventory.DataGrid1.Refresh
        sand_ESTLAM_inventory.Text3.text = auto_sanad_no

    End If

    '##################################################

    '########################################################### SAND SARF INVENTORY
    If case_id.Caption = 50 Then ' SAND SARF

        Adodc5.CommandType = adCmdText
        Adodc5.RecordSource = "select * from  inventories where branch_no=" & Branch_NO & " and departement='" & departement_name & "'  and  master = 1 and   not (inventory_name='')"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount = 0 Then
            If my_language = "E" Then
                MsgBox "no primary stores defined operation will be cancelled", vbCritical
            Else
                MsgBox "·« ÌÊÃœ „Œ«“‰ «”«”Ì… ·œÌﬂ ”Ì „ «·€«¡ «·⁄„·Ì…", vbCritical
            End If

            Exit Sub

        End If

        If sand_sarf_inventory.Text3.text = "" Then
            auto_sanad_no = sand_sarf_inventory.sand_numbering()
        Else
            auto_sanad_no = sand_sarf_inventory.Text3.text
            'Adodc1.Recordset.Fields!sanad_no = Adodc1.Recordset.Fields!sandat_pc_no
        End If

        If auto_sanad_no <> "" Then
         
        Else
          
            If my_language = "E" Then
          
                MsgBox "can't save define numbering method first in system manger", vbCritical: Exit Sub
            Else
                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ·«»œ „‰ «œŒ«· —ﬁ„ ··”‰œ «Ê·« ·«‰ﬂ «Œ —   —ﬁÌ„ ”‰œ«  ÌœÊÌ", vbCritical: Exit Sub
            End If
          
        End If
        
        If sand_sarf_inventory.DataCombo4.text = "" Then
            If my_language = "E" Then
          
                MsgBox "can't save define Voucher Type first", vbCritical: Exit Sub
            Else
                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ·«»œ „‰  ÕœÌœ ‰Ê⁄  ··”‰œ  ", vbCritical: Exit Sub
            End If

            Exit Sub
        End If
        
        If sand_sarf_inventory.Text2.text = "" Then
            If my_language = "E" Then
          
                MsgBox "can't save define Based ON    first", vbCritical: Exit Sub
            Else
                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ·«»œ „‰  ÕœÌœ »‰«¡ ⁄·Ïø  ", vbCritical: Exit Sub
            End If

            Exit Sub
        End If
        
        numbering.ConnectionString = connection_string
        numbering.CommandType = adCmdText
        numbering.RecordSource = "select * from sanad_numbering where branch_no=" & Branch_NO & " and departement='" & departement_name & "' and sanad_no=4"
        numbering.Refresh

        If numbering.Recordset.RecordCount = 0 Then
            numbering_type = 0
        Else
            numbering_type = numbering.Recordset.Fields!numbering_id
        End If
        
        For i = 1 To Adodc2.Recordset.RecordCount

            Adodc3.CommandType = adCmdText
            Adodc3.RecordSource = "select * from  items_units where master = 1 and  item_code='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc3.Refresh

            If Adodc3.Recordset.RecordCount > 0 Then
                Adodc3.Recordset.MoveFirst
            Else

                If my_language = "E" Then
                    MsgBox "ther is no primary unit to item code : " & Adodc2.Recordset.Fields!item_code & " and name is  " & Adodc2.Recordset.Fields!items_name, vbCritical: GoTo HXm
                Else
                    MsgBox "·« ÌÊÃœ ÊÕœÂ «”«”Ì… ···’‰› —ﬁ„ " & Adodc2.Recordset.Fields!item_code & " Ê«”„Â  " & Adodc2.Recordset.Fields!items_name, vbCritical: GoTo HXm
                End If

            End If
 
            Adodc7.RecordSource = "select SUM(ta2ther_makhzan) AS AVILABLE_ITEMS from  inventory WHERE branch_no=" & Branch_NO & " and item_code='" & Adodc2.Recordset.Fields!item_code & "' AND inventory_ID='" & Adodc5.Recordset.Fields!fullcode & "'"
            Adodc7.Refresh

            If Adodc7.Recordset.Fields!AVILABLE_ITEMS < Adodc3.Recordset.Fields!unit_value Then
                If my_language = "E" Then
                    x = MsgBox("items in BASIC inventory less than < item in this instrument - do you want continue any case ", vbCritical + vbYesNo)
                Else
                    x = MsgBox("ﬂ„Ì… «·«’‰«› «·„ÊÃÊœ… «ﬁ· „‰ «·ﬂ„Ì… «·„ÿ·Ê» ’—›Â« „‰ «·„Œ“‰ «·—∆Ì”Ì Â·  —Ìœ  ﬂ„·… «·⁄„·Ì… ⁄·Ï «Ì Õ«·", vbYesNo + vbCritical)
                
                End If

                If x = vbNo Then
                    Exit Sub
                End If

            End If

            'Command1(1).Enabled = False
            'Command1(0).Enabled = True
 
            'txt_item_total.Text = Val(txt_rased.Text) * Val(txt_item_by_unit)

            sand_sarf_inventory.Adodc1.Recordset.AddNew
            sand_sarf_inventory.Adodc1.Recordset.Fields!Sanad_No = auto_sanad_no
            sand_sarf_inventory.Adodc1.Recordset.Fields!Branch_NO = Branch_NO
            sand_sarf_inventory.Adodc1.Recordset.Fields!Departement = departement_name
              
            sand_sarf_inventory.Adodc1.Recordset.Fields!user_name = current_user_name

            sand_sarf_inventory.Adodc1.Recordset.Fields!numbering_type = numbering_type 'TRIC
                  
            If numbering_type = 2 Then
                sand_sarf_inventory.Adodc1.Recordset.Fields!sanad_year = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                sand_sarf_inventory.Adodc1.Recordset.Fields!sanad_month = Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            End If

            If numbering_type = 3 Then
                sand_sarf_inventory.Adodc1.Recordset.Fields!sanad_year = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
            
            End If

            sand_sarf_inventory.Adodc1.Recordset.Fields!Transaction_Type = "”‰œ ’—› „Œ“‰Ì"
            sand_sarf_inventory.Adodc1.Recordset.Fields!SANAD_TYPE = sand_sarf_inventory.DataCombo4.text '"”‰œ ’—› «·Ì €Ì— „Õœœ"
            sand_sarf_inventory.Adodc1.Recordset.Fields!bona_3la = sand_sarf_inventory.Text2.text ' "”‰œ ’—› «·Ì" 'Text2.Text
         
            sand_sarf_inventory.Adodc1.Recordset.Fields!Transaction_Date = DateValue(Now)
         
            sand_sarf_inventory.Adodc1.Recordset.Fields!item_code = Adodc2.Recordset.Fields!item_code
            sand_sarf_inventory.Adodc1.Recordset.Fields!items_name = Adodc2.Recordset.Fields!items_name
         
            Adodc8.RecordSource = "select * from  items where branch_no=" & Branch_NO & " and departement='" & departement_name & "' and   blocked=0 AND   fullcode='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc8.Refresh
                  
            ' Adodc10.RecordSource = "select * from  items where branch_no=" & branch_no & " and departement='" & departement_name & "' and   blocked=0 AND  item_code='" & Adodc2.Recordset.Fields!item_code & "'"
            ' Adodc10.Refresh

            sand_sarf_inventory.Adodc1.Recordset.Fields!part_no = Adodc8.Recordset.Fields!part_no
            sand_sarf_inventory.Adodc1.Recordset.Fields!item_unit = Adodc3.Recordset.Fields!Unit_name
            sand_sarf_inventory.Adodc1.Recordset.Fields!items_no_by_one = Adodc3.Recordset.Fields!unit_value 'Val(frm_Amr_shogl.Text6.Text) * Val(frm_Amr_shogl.DataCombo10.BoundText)
            sand_sarf_inventory.Adodc1.Recordset.Fields!ta2ther_makhzan = -1 * Adodc3.Recordset.Fields!unit_value 'Val(frm_Amr_shogl.Text6.Text) * Val(frm_Amr_shogl.DataCombo10.BoundText) ' Val(frm_Amr_shogl.Text6.Text) * -1 'Val(frm_Amr_shogl.Adodc8.Recordset.Fields!items_no_by_one) * -1
            sand_sarf_inventory.Adodc1.Recordset.Fields!Qty = 1 'Val(frm_Amr_shogl.Text6.Text)
            sand_sarf_inventory.Adodc1.Recordset.Fields!inventory_name = Adodc5.Recordset.Fields!inventory_name
            sand_sarf_inventory.Adodc1.Recordset.Fields!inventory_id = Adodc5.Recordset.Fields!fullcode

            '«Œ— ”⁄— ‘—«¡ Ê «· ﬂ·›…
            Adodc12.CommandType = adCmdText
            Adodc12.RecordSource = "select * from  items where fullcode='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc12.Refresh
         
            If Adodc12.Recordset.RecordCount > 0 And Not IsNull(Adodc12.Recordset.Fields!motwaset_taklefa) Then
 
                sand_sarf_inventory.Adodc1.Recordset.Fields!average_cost = Adodc12.Recordset.Fields!motwaset_taklefa 'Val(frm_Amr_shogl.Text9.Text) / Val(frm_Amr_shogl.DataCombo10.BoundText)
                sand_sarf_inventory.Adodc1.Recordset.Fields!total_price = Adodc12.Recordset.Fields!motwaset_taklefa * Adodc3.Recordset.Fields!unit_value ' frm_Amr_shogl.Adodc8.Recordset.Fields!items_no_by_one
                sand_sarf_inventory.Adodc1.Recordset.Fields!last_price = Adodc12.Recordset.Fields!akher_s3r_shera
    
            Else
 
                sand_sarf_inventory.Adodc1.Recordset.Fields!average_cost = 0
                sand_sarf_inventory.Adodc1.Recordset.Fields!total_price = 0
                sand_sarf_inventory.Adodc1.Recordset.Fields!last_price = 0

            End If

            sand_sarf_inventory.Adodc1.Recordset.update
            ' Adodc1.Recordset.MoveLast
HXm:
            Adodc2.Recordset.MoveNext

        Next i

        sqlx = "select * from inventory where branch_no=" & Branch_NO & " and transaction_type='" & "”‰œ ’—› „Œ“‰Ì" & "' and  sanad_no= '" & auto_sanad_no & "'"
        sand_sarf_inventory.Adodc1.RecordSource = sqlx

        sand_sarf_inventory.Adodc1.Refresh
        sand_sarf_inventory.DataGrid1.Refresh
        sand_sarf_inventory.Text3.text = auto_sanad_no

    End If

    '##################################################
    If case_id.Caption = 5 Then ' rased eftetahy
        Adodc5.CommandType = adCmdText
        Adodc5.RecordSource = "select * from  inventories where master=1 and  branch_no=" & Branch_NO & " and departement='" & departement_name & "' and not (inventory_name='')"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount = 0 Then
            If my_language = "E" Then
                MsgBox "you have no  primary inventory operation will be cancelled", vbCritical
            Else
              
                MsgBox "·« ÌÊÃœ „Œ“‰ —∆Ì”Ì  ·œÌﬂ ”Ì „ «·€«¡ «·⁄„·Ì…", vbCritical
            End If

            Exit Sub

        End If

        For i = 1 To Adodc2.Recordset.RecordCount
            Adodc3.CommandType = adCmdText
            Adodc3.RecordSource = "select * from  items_units where  master =1 and item_code='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc3.Refresh

            If Adodc3.Recordset.RecordCount = 0 Then
 
                If my_language = "E" Then
                    MsgBox "ther is no primary unit to item code : " & Adodc2.Recordset.Fields!fullcode & " and name is  " & Adodc2.Recordset.Fields!items_name, vbCritical: GoTo mm
                Else
                    MsgBox "·« ÌÊÃœ ÊÕœÂ «”«”Ì… ···’‰› —ﬁ„ " & Adodc2.Recordset.Fields!item_code & " Ê«”„Â  " & Adodc2.Recordset.Fields!items_name, vbCritical: GoTo mm
                End If

                GoTo mm
            End If

            RASED_EFTETAHY.Adodc1.Recordset.AddNew
            RASED_EFTETAHY.Adodc1.Recordset.Fields!inventory_name = Adodc5.Recordset.Fields!inventory_name
            RASED_EFTETAHY.Adodc1.Recordset.Fields!inventory_id = Adodc5.Recordset.Fields!fullcode
            RASED_EFTETAHY.Adodc1.Recordset.Fields!item_unit = Adodc3.Recordset.Fields!Unit_name
            RASED_EFTETAHY.Adodc1.Recordset.Fields!items_no_by_one = Adodc3.Recordset.Fields!unit_value 'Val(RASED_EFTETAHY.Text6.Text) * Val(RASED_EFTETAHY.DataCombo10.BoundText)
            RASED_EFTETAHY.Adodc1.Recordset.Fields!ta2ther_makhzan = Adodc3.Recordset.Fields!unit_value  'Val(RASED_EFTETAHY.Text6.Text) * Val(RASED_EFTETAHY.DataCombo10.BoundText) ' Val(RASED_EFTETAHY.Text6.Text) * -1 'Val(RASED_EFTETAHY.Adodc1.Recordset.Fields!items_no_by_one) * -1
            RASED_EFTETAHY.Adodc1.Recordset.Fields!Branch_NO = Branch_NO
            RASED_EFTETAHY.Adodc1.Recordset.Fields!Departement = departement_name
            RASED_EFTETAHY.Adodc1.Recordset.Fields!user_name = current_user_name
            RASED_EFTETAHY.Adodc1.Recordset.Fields!item_code = Adodc2.Recordset.Fields!item_code
            RASED_EFTETAHY.Adodc1.Recordset.Fields!items_name = Adodc2.Recordset.Fields!items_name
            RASED_EFTETAHY.Adodc1.Recordset.Fields!part_no = Adodc2.Recordset.Fields!part_no
            RASED_EFTETAHY.Adodc1.Recordset.Fields!Qty = 1 'Val(RASED_EFTETAHY.Text6.Text)
            RASED_EFTETAHY.Adodc1.Recordset.Fields!Transaction_Type = "—’Ìœ «›  «ÕÌ"
            RASED_EFTETAHY.Adodc1.Recordset.Fields!process_type = "«·Ï"
            RASED_EFTETAHY.Adodc1.Recordset.Fields!Transaction_Date = DateValue(Now)
            RASED_EFTETAHY.Adodc1.Recordset.update
mm:
            Adodc2.Recordset.MoveNext
            RASED_EFTETAHY.Adodc1.Refresh
            RASED_EFTETAHY.DataGrid1.Refresh

        Next i

    End If

    If case_id.Caption = 1 Then ' work order

        Adodc5.CommandType = adCmdText
        Adodc5.RecordSource = "select * from  inventories where branch_no=" & Branch_NO & " and departement='" & departement_name & "'  and  master = 1 and   not (inventory_name='')"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount = 0 Then
            If my_language = "E" Then
                MsgBox "no primary stores defined operation will be cancelled", vbCritical
            Else
                MsgBox "·« ÌÊÃœ „Œ«“‰ «”«”Ì… ·œÌﬂ ”Ì „ «·€«¡ «·⁄„·Ì…", vbCritical
            End If

            Exit Sub

        End If

        For i = 1 To Adodc2.Recordset.RecordCount

            Adodc3.CommandType = adCmdText
            Adodc3.RecordSource = "select * from  items_units where master = 1 and  item_code='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc3.Refresh

            If Adodc3.Recordset.RecordCount > 0 Then
                Adodc3.Recordset.MoveFirst
            Else

                If my_language = "E" Then
                    MsgBox "ther is no primary unit to item code : " & Adodc2.Recordset.Fields!item_code & " and name is  " & Adodc2.Recordset.Fields!items_name, vbCritical: GoTo ll
                Else
                    MsgBox "·« ÌÊÃœ ÊÕœÂ «”«”Ì… ···’‰› —ﬁ„ " & Adodc2.Recordset.Fields!item_code & " Ê«”„Â  " & Adodc2.Recordset.Fields!items_name, vbCritical: GoTo ll
                End If

            End If

            Adodc14.RecordSource = "select SUM(ta2ther_makhzan) AS AVILABLE_ITEMS from  inventory WHERE branch_no=" & Branch_NO & " and  item_code='" & Adodc2.Recordset.Fields!item_code & "' AND inventory_id='" & Adodc5.Recordset.Fields!fullcode & "'"
            Adodc14.Refresh

            If IsNull(Adodc14.Recordset.Fields!AVILABLE_ITEMS) Then
                If my_language = "E" Then
                    x = MsgBox("this item :" & Adodc2.Recordset.Fields!item_code & " does not found in this store continue any case", vbYesNo + vbCritical)

                Else
                    x = MsgBox("Â–« «·’‰› " & Adodc2.Recordset.Fields!item_code & "€Ì— „ÊÃÊœ ›Ì «·„Œ“‰ «·„Õœœ Â·  —Ìœ  ﬂ„·… ⁄„·Ì… «·’—› ⁄·Ï «Ì… Õ«·", vbYesNo + vbCritical)
                End If

                If x = vbNo Then
                    GoTo ll
                End If

            End If

            If Adodc14.Recordset.Fields!AVILABLE_ITEMS < Adodc3.Recordset.Fields!unit_value Then

                If my_language = "E" Then
                    x = MsgBox("ITEM" & Adodc2.Recordset.Fields!item_code & " QTY IN THIS ORDER GREATER TAHN > item qty in Basic store continue in case ", vbYesNo + vbCritical)
                Else
                    x = MsgBox("⁄œœ «·ﬁÿ⁄ «·„ÊÃÊœ… ›Ì «·„Œ“‰ «·—∆Ì”Ì „‰ Â–« «·’‰›" & Adodc2.Recordset.Fields!item_code & " «ﬁ· „‰ «·„ÿ·Ê» ’—›…  —Ìœ  ﬂ„·… ⁄„·Ì… «·’—› ⁄·Ï «Ì… Õ«·", vbYesNo + vbCritical)
    
                End If

                If x = vbNo Then
                    GoTo ll
                End If

            End If

            '«Œ— ”⁄— ‘—«¡ Ê «· ﬂ·›…
            Adodc12.CommandType = adCmdText
            Adodc12.RecordSource = "select * from  items where fullcode='" & Adodc2.Recordset.Fields!item_code & "'"
            Adodc12.Refresh
         
            If Adodc12.Recordset.RecordCount > 0 Then

                If IsNull(Adodc12.Recordset.Fields!motwaset_taklefa) Then
     
                    If my_language = "E" Then
                        x = MsgBox("ITEM" & Adodc2.Recordset.Fields!item_code & " cost not defined continue any case", vbYesNo + vbCritical)
                    Else
                        x = MsgBox("«· ﬂ·›… €Ì— „Õœœ… ›Ì „·› «·«’‰«›" & Adodc2.Recordset.Fields!item_code & "    —Ìœ  ﬂ„·… ⁄„·Ì… «·’—› ⁄·Ï «Ì… Õ«·", vbYesNo + vbCritical)
             
                    End If
             
                    If x = vbNo Then
                        GoTo ll
                    End If
     
                End If

            End If

            frm_Amr_shogl.Adodc8.Recordset.AddNew

            If IsNull(Adodc12.Recordset.Fields!motwaset_taklefa) Then

                frm_Amr_shogl.Adodc8.Recordset.Fields!average_cost = 0 'Val(frm_Amr_shogl.Text9.Text) / Val(frm_Amr_shogl.DataCombo10.BoundText)
                frm_Amr_shogl.Adodc8.Recordset.Fields!total_price = 0
                frm_Amr_shogl.Adodc8.Recordset.Fields!last_price = 0
            Else
  
                frm_Amr_shogl.Adodc8.Recordset.Fields!average_cost = Adodc12.Recordset.Fields!motwaset_taklefa 'Val(frm_Amr_shogl.Text9.Text) / Val(frm_Amr_shogl.DataCombo10.BoundText)
                frm_Amr_shogl.Adodc8.Recordset.Fields!total_price = Adodc12.Recordset.Fields!motwaset_taklefa * Adodc3.Recordset.Fields!unit_value ' frm_Amr_shogl.Adodc8.Recordset.Fields!items_no_by_one
                frm_Amr_shogl.Adodc8.Recordset.Fields!last_price = Adodc12.Recordset.Fields!akher_s3r_shera
            End If
    
            frm_Amr_shogl.Adodc8.Recordset.Fields!Branch_NO = Branch_NO
            frm_Amr_shogl.Adodc8.Recordset.Fields!Departement = departement_name

            frm_Amr_shogl.Adodc8.Recordset.Fields!user_name = current_user_name
            frm_Amr_shogl.Adodc8.Recordset.Fields!amr_shogl_fk = frm_Amr_shogl.Text1.text
            frm_Amr_shogl.Adodc8.Recordset.Fields!item_code = Adodc2.Recordset.Fields!item_code
            frm_Amr_shogl.Adodc8.Recordset.Fields!items_name = Adodc2.Recordset.Fields!items_name
            frm_Amr_shogl.Adodc8.Recordset.Fields!part_no = Adodc2.Recordset.Fields!part_no
            frm_Amr_shogl.Adodc8.Recordset.Fields!item_unit = Adodc3.Recordset.Fields!Unit_name
            frm_Amr_shogl.Adodc8.Recordset.Fields!items_no_by_one = Adodc3.Recordset.Fields!unit_value 'Val(frm_Amr_shogl.Text6.Text) * Val(frm_Amr_shogl.DataCombo10.BoundText)
            frm_Amr_shogl.Adodc8.Recordset.Fields!ta2ther_makhzan = -1 * Adodc3.Recordset.Fields!unit_value 'Val(frm_Amr_shogl.Text6.Text) * Val(frm_Amr_shogl.DataCombo10.BoundText) ' Val(frm_Amr_shogl.Text6.Text) * -1 'Val(frm_Amr_shogl.Adodc8.Recordset.Fields!items_no_by_one) * -1
            frm_Amr_shogl.Adodc8.Recordset.Fields!Qty = 1 'Val(frm_Amr_shogl.Text6.Text)
            frm_Amr_shogl.Adodc8.Recordset.Fields!Transaction_Type = "”‰œ ’—› „Œ“‰Ì"
            frm_Amr_shogl.Adodc8.Recordset.Fields!process_type = "«·Ï"
            frm_Amr_shogl.Adodc8.Recordset.Fields!process_text = "«„— ‘€·"
            frm_Amr_shogl.Adodc8.Recordset.Fields!Sanad_No = "W" & frm_Amr_shogl.Text14.text
            frm_Amr_shogl.Adodc8.Recordset.Fields!bona_3la = "«„— ‘€· «·Ï Ê—ﬁ„…" & frm_Amr_shogl.Text14.text
            frm_Amr_shogl.Adodc8.Recordset.Fields!Transaction_Date = DateValue(Now)
            frm_Amr_shogl.Adodc8.Recordset.Fields!item_departement = frm_Amr_shogl.txtdepartement.text
            frm_Amr_shogl.Adodc8.Recordset.Fields!technical = frm_Amr_shogl.DataCombo8.text
            frm_Amr_shogl.Adodc8.Recordset.Fields!technical_notes = frm_Amr_shogl.Text11.text
            frm_Amr_shogl.Adodc8.Recordset.Fields!inventory_name = Adodc5.Recordset.Fields!inventory_name
            frm_Amr_shogl.Adodc8.Recordset.Fields!inventory_id = Adodc5.Recordset.Fields!fullcode
            frm_Amr_shogl.Adodc8.Recordset.update
 
ll:
            Adodc2.Recordset.MoveNext

        Next i

        frm_Amr_shogl.Adodc8.Refresh
        frm_Amr_shogl.DataGrid1.Refresh

    End If

    Unload Me
End Sub

Private Sub ALLButton3_Click()
    On Error Resume Next

    If (case_id.Caption = 1 Or case_id.Caption = 5 Or case_id.Caption = 50 Or case_id.Caption = 60) And Adodc1.Recordset.RecordCount >= 0 Then
        Adodc1.Recordset.MoveFirst

        For i = 1 To Adodc1.Recordset.RecordCount
            Adodc2.Recordset.AddNew
            Adodc2.Recordset.Fields!item_code = Adodc1.Recordset.Fields!fullcode
            Adodc2.Recordset.Fields!items_name = Adodc1.Recordset.Fields!items_name
            Adodc2.Recordset.Fields!part_no = Adodc1.Recordset.Fields!part_no
  
            Adodc2.Recordset.update
            Adodc1.Recordset.MoveNext

        Next i

        Adodc2.Refresh
        DataGrid3.Refresh
        DataGrid4.Refresh
    End If

End Sub

Private Sub ALLButton4_Click()
    On Error Resume Next

    x = MsgBox("Confirm Cancell All selection  √ﬂÌœ «·€«¡ ﬂ· «· ÕœÌœ", vbCritical + vbYesNo)

    If x = vbNo Then Exit Sub

    For i = 1 To Adodc2.Recordset.RecordCount
        Adodc2.Recordset.delete
        Adodc2.Recordset.MoveNext
        DataGrid3.Refresh
        DataGrid4.Refresh
    Next i

End Sub

Private Sub ALLButton5_Click()
    On Error Resume Next

    If case_id.Caption = 1 And Adodc1.Recordset.RecordCount >= 0 Then

        frm_Amr_shogl.DataCombo6.text = Adodc1.Recordset.Fields!fullcode
        Unload Me
    End If

    If case_id.Caption = 5 And Adodc1.Recordset.RecordCount >= 0 Then

        sql = "select * from items  where  branch_no=" & Branch_NO & " and      departement='" & departement_name & "' and   blocked=0 and  fullcode = '" & Adodc1.Recordset.Fields!fullcode & "' "
        RASED_EFTETAHY.Adodc2.RecordSource = sql
        RASED_EFTETAHY.Adodc2.Refresh
        RASED_EFTETAHY.DataCombo1.ReFill

        RASED_EFTETAHY.DataCombo1.text = Adodc1.Recordset.Fields!items_name
 
        RASED_EFTETAHY.Adodc3.RecordSource = "select * from  items_units where item_code='" & Adodc1.Recordset.Fields!fullcode & "'"
        RASED_EFTETAHY.Adodc3.Refresh
        'Adodc1.RecordSource = "select * from  inventory where transaction_type='—’Ìœ «›  «ÕÌ' and items_name='" & DataCombo1.Text & "'"
        'Adodc1.Refresh

        RASED_EFTETAHY.DataCombo2.text = ""
        RASED_EFTETAHY.DataCombo2.ReFill

        Unload Me
    End If

    If case_id.Caption = 60 And Adodc1.Recordset.RecordCount >= 0 Then
 
        sql = "select * from items  where  branch_no=" & Branch_NO & " and      departement='" & departement_name & "' and   blocked=0 and  fullcode = '" & Adodc1.Recordset.Fields!fullcode & "' "
        sand_ESTLAM_inventory.Adodc2.RecordSource = sql
        sand_ESTLAM_inventory.Adodc2.Refresh
        sand_ESTLAM_inventory.DataCombo1.ReFill
  
        sand_ESTLAM_inventory.Adodc3.RecordSource = "select * from  items_units where item_code='" & Adodc1.Recordset.Fields!fullcode & "'"
        sand_ESTLAM_inventory.Adodc3.Refresh
  
        sand_ESTLAM_inventory.DataCombo1.text = Adodc1.Recordset.Fields!items_name
        'RASED_EFTETAHY.Adodc3.Refresh
        Unload Me
    End If

    If case_id.Caption = 50 And Adodc1.Recordset.RecordCount >= 0 Then
 
        sql = "select * from items  where  branch_no=" & Branch_NO & " and      departement='" & departement_name & "' and   blocked=0 and  fullcode = '" & Adodc1.Recordset.Fields!fullcode & "' "
        sand_sarf_inventory.Adodc2.RecordSource = sql
        sand_sarf_inventory.Adodc2.Refresh
        sand_sarf_inventory.DataCombo1.ReFill
  
        sand_sarf_inventory.Adodc3.RecordSource = "select * from  items_units where item_code='" & Adodc1.Recordset.Fields!fullcode & "'"
        sand_sarf_inventory.Adodc3.Refresh
  
        sand_sarf_inventory.DataCombo1.text = Adodc1.Recordset.Fields!items_name
        'RASED_EFTETAHY.Adodc3.Refresh
        Unload Me
    End If

    If case_id.Caption = 8 And Adodc1.Recordset.RecordCount >= 0 Then
 
        REPORTSFRM.DataCombo2.text = Adodc1.Recordset.Fields!fullcode
        'RASED_EFTETAHY.Adodc3.Refresh
        Unload Me
    End If

End Sub

Private Sub DataGrid1_Click()
    On Error Resume Next

    ALLButton1_Click

End Sub

Private Sub DataGrid2_Click()
    On Error Resume Next

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

Public Function sand_numbering1()
    On Error Resume Next
    Dim start_at As Integer

    auto_sanad_no = ""
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & Branch_NO & " and sanad_no=5"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at

    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  inventory where branch_no=" & Branch_NO & " and  transaction_type='" & "”‰œ «” ·«„" & "' and numbering_type=" & numbering_type
        detect_no.Refresh
    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  inventory where branch_no=" & Branch_NO & " and  transaction_type='" & "”‰œ «” ·«„" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  inventory where branch_no=" & Branch_NO & " and  transaction_type='" & "”‰œ «” ·«„" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh
 
            End If
 
        End If
    End If

    If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

        If numbering_type = 0 Then
            ' auto_sanad_no = 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = start_at
            Else
                
                If numbering_type = 2 Then
                    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

                    End If
                End If
            End If
        End If

    Else

        If numbering_type = 0 Then
            'auto_sanad_no = x + 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
            Else
                
                If numbering_type = 2 Then
                    '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                    ' no = 1
                    '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                    '  Else
                    no = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (no + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        no = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (no + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    'MsgBox auto_sanad_no

End Function

Private Sub Form_Activate()
    On Error Resume Next

    If case_id.Caption = 0 Or case_id.Caption = 4 Or case_id.Caption = 8 Then
        Me.Height = 5670
        ALLButton3.Visible = False
    End If

End Sub

Private Sub Form_Load()
    On Error Resume Next

    On Error Resume Next

    '
    If my_language = "E" Then
        ALLButton3.Caption = "Select All"
        ALLButton5.Caption = "Attach"
        ALLButton4.Caption = "Cancel All"
        ALLButton2.Caption = "Attach All"
        Frame2.Visible = True
        Check1.Caption = "Matched"
        Check2.Caption = "Have keywords"
        Me.Caption = "Items search"

        For i = 0 To Text1.count - 1
            Text1(i).Alignment = 0
        Next i

        Frame2.Visible = False
    Else

        Frame1.Visible = False

    End If

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    connection_string = Cn.ConnectionString
    
    Adodc11.ConnectionString = connection_string
    Adodc11.CommandType = adCmdText
 
    Adodc12.ConnectionString = connection_string
    Adodc12.CommandType = adCmdText
 
    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
 
    Adodc8.ConnectionString = connection_string
    Adodc8.CommandType = adCmdText
 
    Adodc10.ConnectionString = connection_string
    Adodc10.CommandType = adCmdText
 
    Adodc14.ConnectionString = connection_string
    Adodc14.CommandType = adCmdText
 
    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
 
    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
 
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = " select * from TblItems  where ItemID=0" '   branch_no=" & branch_no & " and   blocked=0 AND   departement='" & departement_name & "'"
    Adodc1.Refresh

    sql = Adodc1.RecordSource

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    'Adodc2.RecordSource = " select * from items_temp "
    'Adodc2.Refresh
      
    For i = 0 To Adodc2.Recordset.RecordCount
        Adodc2.Recordset.delete
        Adodc2.Recordset.MoveNext
    Next i

    If my_language = "E" Then
        DataGrid1.Visible = True
        DataGrid2.Visible = False
    End If

    If case_id.Caption = 1 Then
   
        If my_language = "E" Then
            '   ALLButton1.Caption = "Add"
            ALLButton5.Caption = "Attach "
            ALLButton3.Caption = "Select ALL "
                  
            ALLButton2.Caption = " Attach  ALL"
            ALLButton4.Caption = "Cancel All"
            
            DataGrid3.Visible = True
            DataGrid1.Visible = True
            DataGrid2.Visible = False
            DataGrid4.Visible = False
       
        End If
   
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    case_id.Caption = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, _
                        KeyCode As Integer, _
                        Shift As Integer)
    On Error Resume Next

    If KeyCode = 13 Then

        On Error Resume Next
             
        If Index = 0 Then
            ' If Not IsNumeric(Text1(Index).Text) Then Exit Sub
            sql = "select * from TblItems  where ItemCode like '%" & Text1(Index).text & "%' "
    
        End If
                   
        If Index = 1 Then
            sql = "select * from  TblItems  where      ItemName like '%" & Text1(Index).text & "%' "
        End If
                      
        If Index = 2 Then
            sql = "select * from  TblItems  where    ItemID like '%" & Text1(Index).text & "%' "
        End If
             
        Adodc1.RecordSource = sql
        Adodc1.Refresh

        If my_language = "E" Then
            DataGrid1.Refresh
        Else
            DataGrid2.Refresh
        End If

        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "not fount  ·«ÌÊÃœ ‰ «∆Ã ··»ÕÀ", vbInformation
        End If
 
    End If

End Sub

