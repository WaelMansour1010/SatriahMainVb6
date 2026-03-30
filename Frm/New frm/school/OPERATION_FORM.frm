VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form OPERATION_FORM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… «·Œ“Ì‰… ··⁄÷ÊÌ«  "
   ClientHeight    =   8910
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13485
   Begin VB.TextBox Text15 
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc4"
      Height          =   615
      Left            =   1920
      TabIndex        =   24
      Text            =   "Text15"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      DataField       =   "update_year"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   23
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "«·€«¡ «·€—«„…"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      DataField       =   "discounts"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      TabIndex        =   21
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "«⁄«œ… «·Õ”«Ì« "
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      DataField       =   "CHILD_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      DataField       =   "TOTAL_VALUE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "OTHERS_FEES"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "ID_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "FINES_TOTAL"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "INSTALLMENTS_TOTAL"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_TYPE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_NO"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_DATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "”œ«œ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "»ÕÀ"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "»ÕÀ"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "«Œ Ì«—"
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "«Œ Ì«—"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   " ÿ»Ìﬁ «·€—«„…"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      DataField       =   "CHILD_NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   4920
      Top             =   8400
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
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Left            =   4320
      Top             =   9120
      Visible         =   0   'False
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "OPERATION_FORM.frx":0000
      Height          =   2655
      Left            =   1680
      TabIndex        =   25
      Top             =   5760
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   0   'False
      HeadLines       =   1
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "op_id"
         Caption         =   "op_id"
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
         DataField       =   "CHILD_ID"
         Caption         =   "CHILD_ID"
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
         DataField       =   "CHILD_NAME"
         Caption         =   "CHILD_NAME"
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
         DataField       =   "MEMBER_VALUE"
         Caption         =   "MEMBER_VALUE"
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
         DataField       =   "member_card_value"
         Caption         =   "member_card_value"
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
         DataField       =   "FINES_name"
         Caption         =   "FINES_name"
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
         DataField       =   "FINES_value"
         Caption         =   "FINES_value"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   120
      Top             =   9360
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   495
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   495
      Left            =   360
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   495
      Left            =   1200
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "KEST"
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
      Height          =   495
      Left            =   1320
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "GHRAMA"
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
      Height          =   495
      Left            =   1080
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "GHRAMA"
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
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «· «»⁄"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   49
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "”‰… «· ÃœÌœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   48
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "›—Êﬁ«   Œ’„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   47
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11400
      TabIndex        =   46
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·€—«„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   45
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·€—«„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·€—«„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   43
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·ﬂ«—‰Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   42
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "«”„ «· «»⁄"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   41
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "«Ã„«·Ì «‘ —«ﬂ«  «· «»⁄Ì‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   40
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "«Ã„«·Ì «·„ÿ·Ê»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   39
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "„’«—Ì› «Œ—Ï"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   38
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… ÿ»«⁄Â «·ﬂ«—‰Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   37
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·«‘ —«ﬂ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   36
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·«ﬁ”«ÿ  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   35
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   34
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·⁄·„Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   33
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
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
      Height          =   615
      Left            =   5400
      TabIndex        =   32
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·«‘ —«ﬂ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   31
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·⁄÷Ê"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   30
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·⁄÷Ê"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   29
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·€—«„…  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   28
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·“ÊÃ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·“ÊÃ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   26
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "OPERATION_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim check As Integer

Private Sub Command1_Click()
    Dim MEMBER_ID_V As String
    MEMBER_ID_V = Text1.text
    MEMBER_NAME_V = Text6.text

    If Text1.text = "" Then

        MsgBox "·«ÌÊÃœ «Ì ⁄„·Ì« ", vbCritical
        Exit Sub
    End If

    'Text1_Change
    Adodc1.Recordset.Fields!payed = 1
    Adodc1.Recordset.Fields!ACTUAL_VALUE = Text11.text
    Adodc1.Recordset.update

    If Label14.Caption <> "" Then
        Adodc6.CommandType = adCmdText
        Adodc6.RecordSource = "SELECT * FROM INSTALLMENT_DETAILS  WHERE MEMBER_ID=" & MEMBER_ID_V & "AND INSTALLMENT_NO=" & Label14.Caption
        Adodc6.Refresh
    
        If Adodc6.Recordset.RecordCount > 0 Then
            Adodc6.Recordset.Fields!payed = 1
            Adodc6.Recordset.Fields!DATE_OF_PAYED = Date
            Adodc6.Recordset.Fields!ACTIVATED = 0
            Adodc6.Recordset.update

            DoEvents
        
        End If
    
        Adodc6.CommandType = adCmdText
        Adodc6.RecordSource = "SELECT * FROM INSTALLMENT_DETAILS  WHERE MEMBER_ID=" & MEMBER_ID_V & "AND INSTALLMENT_NO=" & Label14.Caption + 1
        Adodc6.Refresh

        If Adodc6.Recordset.RecordCount > 0 Then
            Adodc6.Recordset.Fields!ACTIVATED = 1
            Adodc6.Recordset.update
        End If
    
    End If

    If Label15.Caption <> "" Then
        Adodc7.CommandType = adCmdText
        Adodc7.RecordSource = "SELECT * FROM fines_DETAILS  WHERE MEMBER_ID=" & MEMBER_ID_V & " AND fines_no=" & Label15.Caption
        Adodc7.Refresh
    
        If Adodc7.Recordset.RecordCount > 0 Then
            Adodc7.Recordset.Fields!payed = 1
            Adodc7.Recordset.Fields!ACTIVATED = 0
            Adodc7.Recordset.Fields!PAYED_DATE = Date
            Adodc7.Recordset.update
        End If
    
        Adodc7.CommandType = adCmdText
        Adodc7.RecordSource = "SELECT * FROM fines_DETAILS  WHERE MEMBER_ID=" & MEMBER_ID_V & "AND fines_no=" & Label15.Caption + 1
        Adodc7.Refresh
    
        If Adodc7.Recordset.RecordCount > 0 Then
            Adodc7.Recordset.Fields!ACTIVATED = 1
            Adodc7.Recordset.update
        End If
    
    End If

    If Adodc2.Recordset.RecordCount > O Then
        Adodc2.Recordset.MoveFirst

        For i = 1 To Adodc2.Recordset.RecordCount

            Adodc4.Recordset.AddNew
            Adodc5.CommandType = adCmdText
            Adodc5.RecordSource = "SELECT * FROM MEMBER_CHILD  WHERE MEMBER_ID=" & MEMBER_ID_V & " AND MEMBER_CHILD_ID=" & Adodc2.Recordset.Fields!CHILD_ID
            Adodc5.Refresh

            Adodc4.Recordset.Fields!image_path = Adodc5.Recordset.Fields!MEMBER_CHILD_iMAGE_PATH
            Adodc4.Recordset.Fields!member_id = Adodc5.Recordset.Fields!member_id & "-" & Adodc2.Recordset.Fields!CHILD_ID
            Adodc4.Recordset.Fields!member_name = Adodc2.Recordset.Fields!CHILD_NAME
            Adodc4.Recordset.Fields!MEMBER_TYPE = "⁄÷Ê  «»⁄"
    
            Adodc4.Recordset.Fields!update_year = Text14.text
    
            Adodc2.Recordset.Fields!payed = 1
            Adodc2.Recordset.update
            Adodc4.Recordset.update
            Adodc2.Recordset.MoveNext
        Next i

    End If

    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT * FROM members  WHERE MEMBER_ID=" & MEMBER_ID_V
    Adodc2.Refresh

    Adodc4.Recordset.AddNew
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "SELECT * FROM MEMBERS  WHERE MEMBER_ID=" & MEMBER_ID_V
    Adodc5.Refresh
    Adodc4.Recordset.Fields!image_path = Adodc5.Recordset.Fields!image_path
    Adodc4.Recordset.Fields!member_id = MEMBER_ID_V
    Adodc4.Recordset.Fields!member_name = MEMBER_NAME_V
    Adodc4.Recordset.Fields!MEMBER_TYPE = "⁄÷Ê ⁄«„·"
    Adodc4.Recordset.Fields!update_year = Text14.text
    Adodc4.Recordset.update

    Adodc1.Refresh

End Sub

Private Sub Command2_Click()
    Text3_Change
End Sub

Private Sub Command3_Click()
    x = InputBox("«œŒ· «·—ﬁ„ «Ê Ã“¡ „‰ «·—ﬁ„", "‘«‘… «·»ÕÀ »«·—ﬁ„")

    'select * from operations where PAYED=0 and
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM OPERATIONS where  PAYED=0 and MEMBER_ID LIKE'%" & x & "%'"
    Adodc1.Refresh

    If Text1.text = "" Then Text1.text = 0
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,Fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND MEMBER_ID=" & Text1.text
    Adodc2.Refresh
End Sub

Private Sub Command4_Click()
    x = InputBox("«œŒ· «·«”„ «Ê Ã“¡ „‰ «·«”„", "‘«‘… «·»ÕÀ »«·«”„")

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM OPERATIONS where PAYED=0 and MEMBER_NAME LIKE'%" & x & "%'"
    Adodc1.Refresh

    If Text1.text = "" Then Text1.text = 0
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,Fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND MEMBER_ID=" & Text1.text
    Adodc2.Refresh
End Sub

Private Sub Command5_Click()
    Adodc2.Recordset.Fields!Fines_NAME = ""
    Adodc2.Recordset.Fields!fines_value = 0
    Adodc2.Recordset.update
    Text3_Change
    '    Adodc2.Refresh
    DataGrid2.Refresh

End Sub

Private Sub Command6_Click()
    installments_update.Show
    installments_update.Adodc1.CommandType = adCmdText
    installments_update.Adodc1.RecordSource = "select *  FROM Installments where MEMBER_ID =" & Text1.text & " and child_id=" & Text16.text & " ORDER BY CHILD_ID DESC"
    installments_update.Adodc1.Refresh

    installments_update.Adodc2.CommandType = adCmdText
    installments_update.Adodc2.RecordSource = "select INSTALLMENT_NO,INSTALLMENT_VALUE,DATE_OF_PAYED,ACTIVATED,payed  FROM INSTALLMENT_DETAILS where payed=0 and MEMBER_ID =" & Text1.text & "and child_id=" & Text16.text
    installments_update.Adodc2.Refresh

End Sub

Private Sub Command7_Click()
    fines_update.Show
    'fines_update..Adodc1.CommandType = adCmdText
    'fines_update.Adodc1.RecordSource = "select *  FROM FINES where MEMBER_ID =" & Text1.text & "and child_id=" & Text16.text
    'fines_update.Adodc1.Refresh

    fines_update.Adodc2.CommandType = adCmdText
    fines_update.Adodc2.RecordSource = "select FINES_NO,FINES_VALUE,PAYED_DATE,ACTIVATED,payed  FROM FINES_DETAILS where payed =0 and MEMBER_ID =" & Text1.text & "and child_id=" & Text16.text
    fines_update.Adodc2.Refresh

End Sub

Private Sub DataCombo1_Click(Area As Integer)

    If DataCombo1.text <> "" And Adodc2.Recordset.EOF = False And Adodc2.Recordset.BOF = False Then
        Adodc2.Recordset.Fields!Fines_NAME = DataCombo1.text
        Adodc2.Recordset.update
        Adodc2.Recordset.MoveFirst
        '    Adodc2.Refresh
        DataGrid2.Refresh
        Text3_Change
    End If
    
End Sub

Private Sub Command8_Click()

    If Not IsNull(Adodc2.Recordset.Fields!op_id) Then
        Fine_TYPES.Show
        Fine_TYPES.Text1.Enabled = False
        Fine_TYPES.Text2.Enabled = False
        Fine_TYPES.Text3.Enabled = False
        Fine_TYPES.Text4.Enabled = False
        Fine_TYPES.Command1.Visible = False
        Fine_TYPES.Command2.Visible = False
        'Fine_TYPES.Command3.Visible = True
        'Fine_TYPES.Text5.text = Adodc2.Recordset.Fields!MEMBER_VALUE
        Fine_TYPES.Label5.Caption = "«Œ «— ‰Ê⁄  «·€—«„… „‰ ›÷·ﬂ"
    End If

End Sub

Public Function updatedata()
    Text3_Change
End Function

Private Sub Form_Activate()
    check = 1
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select  *  from operations where PAYED=0 "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value ,FINES_name,FINES_value ,PAYED FROM OPERATION_DETAILS  WHERE MEMBER_ID=0  AND PAYED =0 "
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  Fine_TYPES"
    Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select * from  ready_to_print "
    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from  MEMBER_CHILD "
    Adodc5.Refresh

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    Adodc6.RecordSource = "select * from   INSTALLMENT_DETAILS"
    Adodc6.Refresh

    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "select * from  FINES_DETAILS"
    Adodc7.Refresh

    Adodc8.ConnectionString = connection_string
    Adodc8.CommandType = adCmdText
    Adodc8.RecordSource = "select * from FINES_DETAILS  "
    Adodc8.Refresh

    check = 0

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
    End If

End Sub

Private Sub Text10_Change()

    If IsNumeric(Text10.text) Then
        Text3_Change
 
    End If

End Sub

Private Sub Text13_Change()

    If IsNumeric(Text13.text) Then
        Text3_Change
 
    End If

End Sub

Private Sub Text3_Change()

    If Text16.text = "0" Then
        Adodc2.CommandType = adCmdText
        Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Text1.text
        Adodc2.Refresh
    Else
        Adodc2.CommandType = adCmdText
        Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=0"
        Adodc2.Refresh

    End If

    On Error GoTo ll

    If Text3.text <> "" And check = 1 Then
        Dim SUM As Single
        SUM = 0
    
        For i = 1 To Adodc2.Recordset.RecordCount
            SUM = SUM + Adodc2.Recordset.Fields!MEMBER_VALUE + Adodc2.Recordset.Fields!member_card_value + Adodc2.Recordset.Fields!fines_value
    
            Adodc2.Recordset.MoveNext
        Next i
 
        If Adodc2.Recordset.RecordCount > 0 Then
            Adodc2.Recordset.MoveFirst
        End If
    
        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.Fields!CHILD_VALUE = SUM
        End If
      
        'If Text5.Text = "" Then Text5.Text = 0
        'If Text7.Text = "" Then Text7.Text = 0
        'If Text8.Text = "" Then Text8.Text = 0
        'If Text9.Text = "" Then Text9.Text = 0
        'If Text10.Text = "" Then Text10.Text = 0
        'If Text13.Text = "" Then Text13.Text = 0
        'If Text13.Text = "" Then Text13.Text = 0
        'MsgBox Adodc1.Recordset.Fields!MEMBER_VALUE
      
        If Text3.text <> "" Then
            Adodc1.Recordset.Fields!total_value = SUM + Text7.text + Text5.text + Text8.text + Text9.text + Text10.text - Text13.text
            Adodc1.Recordset.update
            ' Adodc1.Recordset.Update
    
        End If

    End If

ll:
End Sub

Private Sub Text5_Change()

    If IsNumeric(Text5.text) Then
        Adodc1.Recordset.Fields!INSTALLMENTS_TOTAL = Text5.text
        Adodc1.Recordset.update
        Text3_Change
    End If

End Sub

Private Sub Text8_Change()

    If IsNumeric(Text8.text) Then
        Adodc1.Recordset.Fields!FINES_TOTAL = Text8.text
        Adodc1.Recordset.update
        Text3_Change

    End If

End Sub

Private Sub Text9_Change()

    If IsNumeric(Text9.text) Then
        Text3_Change
    End If

End Sub

