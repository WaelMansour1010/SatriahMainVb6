VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form SECURITY_FORM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "              „Ê«›ﬁ… «·Ê“«—…"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox Combo1 
      DataField       =   "year"
      DataSource      =   "Adodc13"
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Text            =   " "
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "CHILD_NAME"
      DataSource      =   "Adodc9"
      Height          =   375
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "Text7"
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "OPERATION_NO"
      DataSource      =   "Adodc8"
      Height          =   285
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox CARD_VALUE 
      Alignment       =   1  'Right Justify
      DataField       =   "CARD_VALUE"
      DataSource      =   "Adodc10"
      Height          =   375
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "INSTALLMENT_NO"
      DataSource      =   "Adodc6"
      Height          =   495
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_CHILD_NAME"
      DataSource      =   "Adodc4"
      Height          =   285
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   7560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "index"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "„Ê«›ﬁ… «·Ê“«—…"
      Height          =   855
      Left            =   1920
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   4575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SECURITY_FORM.frx":0000
      Height          =   4215
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7435
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
      Height          =   330
      Left            =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   582
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
      Height          =   330
      Left            =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   582
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
      RecordSource    =   " "
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Left            =   120
      Top             =   8280
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   495
      Left            =   240
      Top             =   7800
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   495
      Left            =   -1560
      Top             =   5640
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   495
      Left            =   240
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   495
      Left            =   -120
      Top             =   9240
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   495
      Left            =   7920
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   495
      Left            =   8160
      Top             =   5640
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   495
      Left            =   120
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   495
      Left            =   240
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   495
      Left            =   240
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "”‰… «·„Ê«›ﬁ…"
      Height          =   255
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line7 
      BorderWidth     =   5
      X1              =   2040
      X2              =   8400
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      BorderWidth     =   5
      X1              =   8400
      X2              =   8400
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·ÿ«·»"
      Height          =   255
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   3480
      X2              =   3480
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   " «—ÌŒ «·ÿ·»"
      Height          =   255
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   255
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "„”·”·"
      Height          =   255
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·’Ê—…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   2040
      X2              =   8400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   2040
      X2              =   2040
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   5520
      X2              =   5520
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line5 
      BorderWidth     =   5
      X1              =   7200
      X2              =   7200
      Y1              =   0
      Y2              =   360
   End
End
Attribute VB_Name = "SECURITY_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim dtmTest As Date
    Dim last_id As Integer
 
    Dim member_id  As String
    Dim MEMBER_CHILD_ID  As Integer

    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    Dim check As String
    Dim installment_value As Single
    check = "singel"
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select *  FROM new_members WHERE [INDEX]=" & Text1.text
    Adodc2.Refresh

    Adodc12.CommandType = adCmdText
    Adodc12.RecordSource = "select *  FROM MEMBER_TYPES Where [member_id] = " & Adodc1.Recordset.Fields!MEMBER_TYPE
    Adodc12.Refresh

    Adodc14.CommandType = adCmdText
    Adodc14.RecordSource = "select max(member_id) as last_id FROM MEMBERs"
    Adodc14.Refresh

    If Not IsNull(Adodc14.Recordset.Fields!last_id) Then
        last_id = Adodc14.Recordset.Fields!last_id + 1
    Else
        last_id = 1
    End If

    Adodc3.Recordset.AddNew
    Adodc3.Recordset.Fields!member_id = last_id

    Adodc3.Recordset.Fields!VALUE_OF_MEMBER = Adodc12.Recordset.Fields!value
    Adodc3.Recordset.Fields!member_name = Adodc2.Recordset.Fields!member_name
    Adodc3.Recordset.Fields!MEMBER_TYPE = Adodc12.Recordset.Fields!member_name
    Adodc3.Recordset.Fields!MEMBER_DOB = Adodc2.Recordset.Fields!MEMBER_DOB
    Adodc3.Recordset.Fields!MEMBER_born_place = Adodc2.Recordset.Fields!MEMBER_born_place
    Adodc3.Recordset.Fields!MEMBER_address = Adodc2.Recordset.Fields!MEMBER_address
    Adodc3.Recordset.Fields!MEMBER_certificate = Adodc2.Recordset.Fields!MEMBER_certificate
    Adodc3.Recordset.Fields!MEMBER_telephone = Adodc2.Recordset.Fields!MEMBER_telephone
    Adodc3.Recordset.Fields!MEMBER_job = Adodc2.Recordset.Fields!MEMBER_job
    Adodc3.Recordset.Fields!MEMBER_job_address = Adodc2.Recordset.Fields!MEMBER_job_address
    Adodc3.Recordset.Fields!MEMBER_NATIONAL_id = Adodc2.Recordset.Fields!original_NATIONAL_id
    Adodc3.Recordset.Fields!MEMBER_date_of_issue = DateValue(Now)

    Adodc3.Recordset.Fields!Sex = "–ﬂ—"
    Adodc3.Recordset.update

    member_id = Adodc3.Recordset.Fields!member_id
    MEMBER_CHILD_ID = 0

    If Not IsNull(Adodc2.Recordset.Fields!wife_order_number) Then

        'Adodc12.CommandType = adCmdText
        'Adodc12.RecordSource = "select *  FROM MEMBER_TYPES Where [MEMBER_ID] = 2"
        'Adodc12.Refresh

        check = "family"
        Adodc4.Recordset.AddNew
        Adodc4.Recordset.Fields!member_id = member_id
        MEMBER_CHILD_ID = MEMBER_CHILD_ID + 1
        Adodc4.Recordset.Fields!VALUE_OF_MEMBER = Adodc12.Recordset.Fields!value
        Adodc4.Recordset.Fields!MEMBER_CHILD_ID = MEMBER_CHILD_ID
        Adodc4.Recordset.Fields!MEMBER_CHILD_NAME = Adodc2.Recordset.Fields!wife_NAME
        Adodc4.Recordset.Fields!MEMBER_DOB = Adodc2.Recordset.Fields!wife_DOB
        Adodc4.Recordset.Fields!MEMBER_born_place = Adodc2.Recordset.Fields!wife_born_place
        Adodc4.Recordset.Fields!MEMBER_address = Adodc2.Recordset.Fields!wife_address
        Adodc4.Recordset.Fields!MEMBER_certificate = Adodc2.Recordset.Fields!wife_certificate
        Adodc4.Recordset.Fields!MEMBER_telephone = Adodc2.Recordset.Fields!wife_telephone
        Adodc4.Recordset.Fields!MEMBER_job = Adodc2.Recordset.Fields!wife_job
        Adodc4.Recordset.Fields!MEMBER_job_address = Adodc2.Recordset.Fields!wife_job_address
        Adodc4.Recordset.Fields!MEMBER_NATIONAL_id = Adodc2.Recordset.Fields!original_NATIONAL_id
        Adodc4.Recordset.Fields!MEMBER_TITLE = " «»⁄ 1"
        'Adodc4.Recordset.Fields!member_type = "⁄«∆·Ï   «»⁄"
        Adodc4.Recordset.Fields!member_type_name = "⁄÷Ê ⁄«„·  «»⁄"
        Adodc4.Recordset.Fields!MEMBER_date_of_issue = DateValue(Now)
        'Adodc4.Recordset.Fields!sex = "«‰ÀÏ"
        Adodc4.Recordset.update

        'Adodc7.CommandType = adCmdText
        'Adodc7.RecordSource = "select * from Installments_TYPES where Installments_ID=2"
        'Adodc7.Refresh

        'Adodc5.Recordset.AddNew
        'Adodc5.Recordset.Fields!member_id = member_id
        'Adodc5.Recordset.Fields!MEMBER_NAME = Adodc2.Recordset.Fields!wife_NAME
        'Adodc5.Recordset.Fields!CHILD_ID = MEMBER_CHILD_ID
        'Adodc5.Recordset.Fields!Installments_VALUE = Adodc7.Recordset.Fields!Installments_VALUE
        'Adodc5.Recordset.Fields!Installments_TYPE = "ﬁ”ÿ  √”Ì” ›—œÌ"

        'Adodc5.Recordset.Update
        'Dim i As Integer
        'For i = 1 To 10
        'Adodc6.Recordset.AddNew
        'Adodc6.Recordset.Fields!member_id = member_id
        'Adodc6.Recordset.Fields!CHILD_ID = MEMBER_CHILD_ID
        'Adodc6.Recordset.Fields!INSTALLMENT_NO = i
        'Adodc6.Recordset.Fields!installment_value = Round(Adodc7.Recordset.Fields!Installments_VALUE / 10, 2)
        'installment_value = Round(Adodc7.Recordset.Fields!Installments_VALUE / 10, 2)
        'If i = 1 Then
        'Adodc6.Recordset.Fields!ACTIVATED = 1
        'End If
        'Adodc6.Recordset.Update
        'Next i

    End If

    'Adodc12.CommandType = adCmdText
    'Adodc12.RecordSource = "select *  FROM MEMBER_TYPES Where [MEMBER_NAME] = '⁄«∆·Ï   «»⁄'"
    'Adodc12.Refresh

    If Not IsNull(Adodc2.Recordset.Fields!SISTER_order_number) Then
        check = "family"
        Adodc4.Recordset.AddNew
        Adodc4.Recordset.Fields!member_id = member_id
        MEMBER_CHILD_ID = MEMBER_CHILD_ID + 1
        Adodc4.Recordset.Fields!VALUE_OF_MEMBER = Adodc12.Recordset.Fields!value
        Adodc4.Recordset.Fields!MEMBER_CHILD_ID = MEMBER_CHILD_ID
        Adodc4.Recordset.Fields!MEMBER_CHILD_NAME = Adodc2.Recordset.Fields!SISTER_NAME
        Adodc4.Recordset.Fields!MEMBER_DOB = Adodc2.Recordset.Fields!SISTER_DOB
        Adodc4.Recordset.Fields!MEMBER_born_place = Adodc2.Recordset.Fields!SISTER_born_place
        Adodc4.Recordset.Fields!MEMBER_address = Adodc2.Recordset.Fields!SISTER_address
        Adodc4.Recordset.Fields!MEMBER_certificate = Adodc2.Recordset.Fields!SISTER_certificate
        Adodc4.Recordset.Fields!MEMBER_telephone = Adodc2.Recordset.Fields!SISTER_telephone
        Adodc4.Recordset.Fields!MEMBER_job = Adodc2.Recordset.Fields!SISTER_job
        Adodc4.Recordset.Fields!MEMBER_job_address = Adodc2.Recordset.Fields!SISTER_job_address
        Adodc4.Recordset.Fields!MEMBER_NATIONAL_id = Adodc2.Recordset.Fields!original_NATIONAL_id
        Adodc4.Recordset.Fields!MEMBER_TITLE = "«·«Œ "
        Adodc4.Recordset.Fields!MEMBER_TYPE = " ⁄«∆·Ï   «»⁄"
        Adodc4.Recordset.Fields!member_type_name = "⁄÷Ê  «»⁄"
        Adodc4.Recordset.Fields!MEMBER_date_of_issue = DateValue(Now)

        Adodc4.Recordset.Fields!Sex = "«‰ÀÏ"
        Adodc4.Recordset.update
    End If

    If Not IsNull(Adodc2.Recordset.Fields!GRANDFATHER_order_number) Then
        check = "family"
        Adodc4.Recordset.AddNew
        Adodc4.Recordset.Fields!member_id = member_id
        MEMBER_CHILD_ID = MEMBER_CHILD_ID + 1
        Adodc4.Recordset.Fields!MEMBER_CHILD_ID = MEMBER_CHILD_ID
        Adodc4.Recordset.Fields!VALUE_OF_MEMBER = Adodc12.Recordset.Fields!value
        Adodc4.Recordset.Fields!MEMBER_CHILD_NAME = Adodc2.Recordset.Fields!GRANDFATHER_NAME
        Adodc4.Recordset.Fields!MEMBER_DOB = Adodc2.Recordset.Fields!GRANDFATHER_DOB
        Adodc4.Recordset.Fields!MEMBER_born_place = Adodc2.Recordset.Fields!GRANDFATHER_born_place
        Adodc4.Recordset.Fields!MEMBER_address = Adodc2.Recordset.Fields!GRANDFATHER_address
        Adodc4.Recordset.Fields!MEMBER_certificate = Adodc2.Recordset.Fields!GRANDFATHER_certificate
        Adodc4.Recordset.Fields!MEMBER_telephone = Adodc2.Recordset.Fields!GRANDFATHER_telephone
        Adodc4.Recordset.Fields!MEMBER_job = Adodc2.Recordset.Fields!GRANDFATHER_job
        Adodc4.Recordset.Fields!MEMBER_job_address = Adodc2.Recordset.Fields!GRANDFATHER_job_address
        Adodc4.Recordset.Fields!MEMBER_NATIONAL_id = Adodc2.Recordset.Fields!original_NATIONAL_id
        Adodc4.Recordset.Fields!MEMBER_TITLE = "«·Ãœ"
        Adodc4.Recordset.Fields!MEMBER_TYPE = "⁄«∆·Ï   «»⁄"
        Adodc4.Recordset.Fields!member_type_name = "⁄÷Ê  «»⁄"
        Adodc4.Recordset.Fields!MEMBER_date_of_issue = DateValue(Now)
        Adodc4.Recordset.Fields!Sex = "–ﬂ—"
        Adodc4.Recordset.update
    End If

    If Not IsNull(Adodc2.Recordset.Fields!GRANDMOTHER_order_number) Then
        check = "family"
        Adodc4.Recordset.AddNew
        Adodc4.Recordset.Fields!member_id = member_id
        Adodc4.Recordset.Fields!VALUE_OF_MEMBER = Adodc12.Recordset.Fields!value
        MEMBER_CHILD_ID = MEMBER_CHILD_ID + 1
        Adodc4.Recordset.Fields!MEMBER_CHILD_ID = MEMBER_CHILD_ID
        Adodc4.Recordset.Fields!MEMBER_CHILD_NAME = Adodc2.Recordset.Fields!GRANDMOTHER_NAME
        Adodc4.Recordset.Fields!MEMBER_DOB = Adodc2.Recordset.Fields!GRANDMOTHER_DOB
        Adodc4.Recordset.Fields!MEMBER_born_place = Adodc2.Recordset.Fields!GRANDMOTHER_born_place
        Adodc4.Recordset.Fields!MEMBER_address = Adodc2.Recordset.Fields!GRANDMOTHER_address
        Adodc4.Recordset.Fields!MEMBER_certificate = Adodc2.Recordset.Fields!GRANDMOTHER_certificate
        Adodc4.Recordset.Fields!MEMBER_telephone = Adodc2.Recordset.Fields!GRANDMOTHER_telephone
        Adodc4.Recordset.Fields!MEMBER_job = Adodc2.Recordset.Fields!GRANDMOTHER_job
        Adodc4.Recordset.Fields!MEMBER_job_address = Adodc2.Recordset.Fields!GRANDMOTHER_job_address
        Adodc4.Recordset.Fields!MEMBER_NATIONAL_id = Adodc2.Recordset.Fields!original_NATIONAL_id
        Adodc4.Recordset.Fields!MEMBER_TITLE = "«·Ãœ…"
        Adodc4.Recordset.Fields!MEMBER_TYPE = "⁄«∆·Ï   «»⁄"
        Adodc4.Recordset.Fields!member_type_name = "⁄÷Ê  «»⁄"
        Adodc4.Recordset.Fields!MEMBER_date_of_issue = DateValue(Now)
        Adodc4.Recordset.Fields!Sex = "«‰ÀÏ"
        Adodc4.Recordset.update
    End If

    Adodc5.Recordset.AddNew
    Adodc5.Recordset.Fields!member_id = member_id
    Adodc5.Recordset.Fields!member_name = Adodc2.Recordset.Fields!member_name

    If check = "family" Then
        Adodc7.CommandType = adCmdText
        Adodc7.RecordSource = "select * from Installments_TYPES where Installments_ID=1"
        Adodc7.Refresh

        Adodc5.Recordset.Fields!Installments_TYPE = "ﬁ”ÿ  √”Ì” ⁄«∆·Ì"
        Adodc5.Recordset.Fields!Installments_VALUE = Adodc7.Recordset.Fields!Installments_VALUE
        Adodc5.Recordset.update

        For i = 1 To 4
            Adodc6.Recordset.AddNew
            Adodc6.Recordset.Fields!member_id = member_id
            Adodc6.Recordset.Fields!INSTALLMENT_NO = i
            Adodc6.Recordset.Fields!installment_value = Round(Adodc7.Recordset.Fields!Installments_VALUE / 4, 2)
            installment_value = Round(Adodc7.Recordset.Fields!Installments_VALUE / 4, 2)

            If i = 1 Then
                Adodc6.Recordset.Fields!ACTIVATED = 1
            End If

            Adodc6.Recordset.update
        Next i

    Else
        Adodc7.CommandType = adCmdText
        Adodc7.RecordSource = "select * from Installments_TYPES where Installments_ID=2"
        Adodc7.Refresh
        Adodc5.Recordset.Fields!Installments_TYPE = "ﬁ”ÿ  √”Ì” ›—œÌ"
        Adodc5.Recordset.Fields!Installments_VALUE = Adodc7.Recordset.Fields!Installments_VALUE
        Adodc5.Recordset.update

        For i = 1 To 10
            Adodc6.Recordset.AddNew
            Adodc6.Recordset.Fields!member_id = member_id
            Adodc6.Recordset.Fields!INSTALLMENT_NO = i
            Adodc6.Recordset.Fields!installment_value = Round(Adodc7.Recordset.Fields!Installments_VALUE / 10, 2)
            installment_value = Round(Adodc7.Recordset.Fields!Installments_VALUE / 10, 2)

            If i = 1 Then
                Adodc6.Recordset.Fields!ACTIVATED = 1
            End If

            Adodc6.Recordset.update
        Next i

    End If

    ' Adodc11.CommandType = adCmdText
    ' Adodc11.RecordSource = "select*  FROM member_TYPES WHERE MEMBER_ID=1"
    ' Adodc11.Refresh

    Adodc8.Recordset.AddNew '··⁄÷Ê «·⁄«„·
    Adodc8.Recordset.Fields!member_id = member_id
    Adodc8.Recordset.Fields!member_name = Adodc2.Recordset.Fields!member_name
    Adodc8.Recordset.Fields!MEMBER_TYPE = Adodc12.Recordset.Fields!member_name
    Adodc8.Recordset.Fields!operation_type = "⁄÷ÊÌ… ÃœÌœ…"
    Adodc8.Recordset.Fields!OPERATION_DATE = DateValue(Now)
    'Adodc8.Recordset.Fields!User_Name = Main.TxtUserName
    Adodc8.Recordset.Fields!MEMBER_VALUE = Adodc12.Recordset.Fields![value] '' «·«‘ —«ﬂ
    Adodc8.Recordset.Fields!INSTALLMENTS_TOTAL = installment_value '«·ﬁ”ÿ
    Adodc8.Recordset.Fields!INSTALLMENTS_NO = 1
    Adodc8.Recordset.Fields!ID_VALUE = CARD_VALUE.text 'À„‰ «·ﬂ«—‰Ì…
    Adodc8.Recordset.Fields!update_year = Combo1.text

    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select*  FROM member_CHILD WHERE MEMBER_ID=" & member_id
    Adodc4.Refresh

    Adodc8.Recordset.update

    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select*  FROM member_CHILD WHERE MEMBER_ID=" & member_id
    Adodc4.Refresh

    Adodc11.CommandType = adCmdText
    Adodc11.RecordSource = "select *  FROM MEMBER_TYPES Where [MEMBER_NAME] = '⁄«∆·Ï   «»⁄'"
    Adodc11.Refresh
 
    ' Adodc12.CommandType = adCmdText
    'Adodc12.RecordSource = "select *  FROM MEMBER_TYPES Where [MEMBER_ID] = 1"
    'Adodc12.Refresh
 
    For i = 1 To Adodc4.Recordset.RecordCount

        Adodc9.Recordset.AddNew '··⁄÷Ê «· «»⁄
        Adodc9.Recordset.Fields!member_id = member_id
        Adodc9.Recordset.Fields!CHILD_ID = Adodc4.Recordset.Fields!MEMBER_CHILD_ID
        Adodc9.Recordset.Fields!CHILD_NAME = Adodc4.Recordset.Fields!MEMBER_CHILD_NAME
        
        If Adodc4.Recordset.Fields!MEMBER_TITLE = "«·“ÊÃ…" Then
            Adodc9.Recordset.Fields!MEMBER_VALUE = Adodc12.Recordset.Fields![value]
            Adodc9.Recordset.Fields!wife = 1
        Else
            '         Adodc9.Recordset.Fields!MEMBER_VALUE = Adodc11.Recordset.Fields![Value]
        End If
       
        Adodc9.Recordset.Fields!member_card_value = CARD_VALUE.text 'À„‰ «·ﬂ«—‰Ì…
        Adodc9.Recordset.update
       
        Adodc4.Recordset.MoveNext
    Next i

    Adodc2.Recordset.Fields!SECURITY = 1
    Adodc2.Recordset.update
    Adodc2.Refresh

    DoEvents
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select  [INDEX],original_order_number,original_order_date,MEMBER_NAME FROM new_members WHERE security=0"
    Adodc1.Refresh
 
    DataGrid1.Refresh
    MsgBox " „  ”ÃÌ· «·ÿ«·» «·ÃœÌœ Ê«·ÿ·«» «· «»⁄Ì‰ ·… »‰Ã«Õ —ﬁ„ «·ÿ«·» ÂÊ :" & last_id, vbInformation
 
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select  [index],original_order_number,original_order_date,MEMBER_NAME,member_type FROM new_members WHERE security=0 "
    Adodc1.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  MEMBERS "
    Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select * from MEMBER_CHILD "
    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from  Installments"
    Adodc5.Refresh

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    Adodc6.RecordSource = "select * from  INSTALLMENT_DETAILS"
    Adodc6.Refresh

    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "select * from  Installments_TYPES "
    Adodc7.Refresh

    Adodc8.ConnectionString = connection_string
    Adodc8.CommandType = adCmdText
    Adodc8.RecordSource = "select * from OPERATIONS "
    Adodc8.Refresh

    Adodc9.ConnectionString = connection_string
    Adodc9.CommandType = adCmdText
    Adodc9.RecordSource = "select * from OPERATION_DETAILS "
    Adodc9.Refresh

    Adodc10.ConnectionString = connection_string
    Adodc10.CommandType = adCmdText
    Adodc10.RecordSource = "select * from  CARD "
    Adodc10.Refresh

    Adodc11.ConnectionString = connection_string
    Adodc11.CommandType = adCmdText
    Adodc11.RecordSource = "select * from  MEMBER_TYPES"
    Adodc11.Refresh

    Adodc12.ConnectionString = connection_string
    Adodc12.CommandType = adCmdText
    Adodc12.RecordSource = "select * from MEMBER_TYPES "
    Adodc12.Refresh

    Adodc13.ConnectionString = connection_string
    Adodc13.CommandType = adCmdText
    Adodc13.RecordSource = "select * from  this_year"
    Adodc13.Refresh

    Adodc14.ConnectionString = connection_string
    Adodc14.CommandType = adCmdText
    Adodc14.RecordSource = "select * from MEMBERS "
    Adodc14.Refresh

    For i = 2010 To 2040
        Combo1.AddItem i & "-" & i + 1
    Next i

    Combo1.text = Adodc13.Recordset.Fields!year
End Sub
