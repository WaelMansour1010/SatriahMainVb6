VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form update_member 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÃœÌœ «· Õ«ﬁ «·ÿ·«»"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   2730
   ScaleWidth      =   8670
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox Text27 
      Alignment       =   2  'Center
      DataField       =   "image_location"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2160
      Width           =   1575
   End
   Begin DBPIXLib.DBPix20 DBPix1 
      Height          =   855
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   1508
      _StockProps     =   1
      _Image          =   "update_member.frx":0000
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_TYPE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc10"
      Height          =   285
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox fines_value 
      Alignment       =   1  'Right Justify
      DataField       =   "Fines_VALUE"
      DataSource      =   "Adodc6"
      Height          =   375
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "Text7"
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc5"
      Height          =   375
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "year"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "⁄—÷ ⁄„·Ì«  «· ÃœÌœ «·„ÊÃÊœ… Õ«·Ì« ·œÏ «·Œ“Ì‰…"
      Height          =   735
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc9"
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox Text2 
      DataField       =   "OPERATION_TYPE"
      DataSource      =   "Adodc8"
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox CARD_VALUE 
      DataField       =   "CARD_VALUE"
      DataSource      =   "Adodc4"
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   3240
      Picture         =   "update_member.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   " ÃœÌœ"
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1800
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   2640
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Left            =   7080
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Left            =   1680
      Top             =   6600
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
      CommandType     =   2
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
      Height          =   495
      Left            =   2040
      Top             =   5400
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
      CommandType     =   2
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
      Height          =   975
      Left            =   3960
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1720
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Height          =   615
      Left            =   6960
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Height          =   615
      Left            =   3720
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Left            =   2040
      Top             =   6960
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
      CommandType     =   2
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
   Begin MSAdodcLib.Adodc Adodc20 
      Height          =   615
      Left            =   1320
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   6840
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Left            =   1320
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Left            =   1320
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Left            =   1320
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "„”·”· «·’Ê—…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1320
      TabIndex        =   20
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "«·’› «·œ—«”Ì"
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
      Left            =   2520
      TabIndex        =   17
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   6960
      TabIndex        =   12
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " ÃœÌœ «· Õ«ﬁ «·ÿ·«»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·ÿ«·»"
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
      Left            =   6960
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "update_member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    operatiomn_update_frm.Show
    operatiomn_update_frm.Caption = "‘«‘…  ÃœÌœ «·⁄÷ÊÌ…"

End Sub

Private Sub Command2_Click()
   
    Dim dtmTest As Date

    If Text1.text = "" Or Not IsNumeric(Text1.text) Then Exit Sub
    Adodc11.CommandType = adCmdText
    Adodc11.RecordSource = "select*  FROM operations WHERE payed=0 and MEMBER_ID=" & Text1.text
    Adodc11.Refresh
 
    If Adodc11.Recordset.RecordCount > 0 Then
        MsgBox "ÌÊÃœ ·Â–« «·⁄÷Ê ⁄„·Ì… ›Ì «·Œ“Ì‰… ·«»œ „‰ ”œ«œÂ« √Ê·«", vbCritical
        Exit Sub
    End If

    If Text1.text = "" Or Text4.text = "" Then
        Exit Sub
    End If

    Adodc11.CommandType = adCmdText
    Adodc11.RecordSource = "select*  FROM member_TYPES WHERE MEMBER_NAME='" & Text8.text & "'"
    Adodc11.Refresh

    Dim update_year As String
    'Adodc2.CommandType = adCmdText
    'Adodc2.RecordSource = "SELECT SUM([VALUE]) AS TOTAL_ACTIVITY FROM member_activity WHERE MEMBER_CHILD_ID =0 AND MEMBER_ID=" & Adodc1.Recordset.Fields!member_id
    'Adodc2.Refresh

    'Adodc20.CommandType = adCmdText
    'Adodc20.RecordSource = "select * from update_year_qry where member_id=" & Text1.Text
    'Adodc20.Refresh

    Adodc20.CommandType = adCmdText
    Adodc20.RecordSource = "select * from members where member_id=" & Text1.text
    Adodc20.Refresh

    Dim no_of_year_with_fines As Integer
    no_of_year_with_fines = 1

    If Not IsNull(Adodc20.Recordset.Fields!last_update_year) Then
        no_of_year_with_fines = val(Mid(Text5.text, 1, 4)) - val(Mid(Adodc20.Recordset.Fields!last_update_year, 1, 4))
    End If

    If no_of_year_with_fines = 0 Then
        MsgBox "·« Ì„ﬂ‰  ‰›Ì– €„·Ì… «· ÃœÌœ ÕÌÀ «‰ Â–« «·⁄÷Ê „Ãœœ »«·›⁄· ··” … «·„«·Ì… " & Text5.text, vbCritical
        Exit Sub

        'MsgBox no_of_year_with_fines
        'For i = 1 To no_of_year_with_fines

        'ADD_MEMBER_FINES
        'Next i

    End If

    If no_of_year_with_fines = 0 Then
        MsgBox "·« Ì„ﬂ‰  ‰›Ì– €„·Ì… «· ÃœÌœ ÕÌÀ «‰ «·”‰… «·„«·Ì… «·Õ«·Ì… «’€— „‰ «Œ— ” …  ÃœÌœ ·Â–« «·⁄÷Ê " & Text5.text, vbCritical
        Exit Sub

    End If

    'update_year = Val(Mid(Adodc20.Recordset.Fields!update_year, 1, 4)) + 1 & "-" & Val(Mid(Adodc20.Recordset.Fields!update_year, 6, 4)) + 1

    Adodc8.Recordset.AddNew '··⁄÷Ê «·⁄«„·
    Adodc8.Recordset.Fields!member_id = Adodc1.Recordset.Fields!member_id
    Adodc8.Recordset.Fields!member_name = Adodc1.Recordset.Fields!member_name
    Adodc8.Recordset.Fields!operation_type = " ÃœÌœ ⁄÷ÊÌ…"
    Adodc8.Recordset.Fields!MEMBER_TYPE = Adodc1.Recordset.Fields!MEMBER_TYPE  ' cc "⁄÷Ê⁄«„·"
    Adodc8.Recordset.Fields!OPERATION_DATE = DateValue(Now)
    'Adodc8.Recordset.Fields!User_Name = Main.TxtUserName
    Adodc8.Recordset.Fields!MEMBER_VALUE = Adodc11.Recordset.Fields![value] '' «·«‘ —«ﬂ
    Adodc8.Recordset.Fields!update_year = Text5.text

    Adodc8.Recordset.Fields!last_update_year = Adodc20.Recordset.Fields!last_update_year
    Adodc8.Recordset.Fields!ID_VALUE = CARD_VALUE.text 'À„‰ «·ﬂ«—‰Ì…

    'If Adodc2.Recordset.RecordCount <> 0 Then
    'Adodc8.Recordset.Fields!activity_value = Adodc2.Recordset.Fields!TOTAL_ACTIVITY
    'End If
    Adodc8.Recordset.update

    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select*  FROM member_CHILD WHERE MEMBER_ID=" & Adodc1.Recordset.Fields!member_id
    Adodc4.Refresh
 
    Adodc11.CommandType = adCmdText
    Adodc11.RecordSource = "select*  FROM member_TYPES WHERE MEMBER_NAME='⁄÷Ê  «»⁄'"
    Adodc11.Refresh
    Adodc12.CommandType = adCmdText
    Adodc12.RecordSource = "select*  FROM member_TYPES WHERE MEMBER_NAME='" & Text8.text & "'"
    Adodc12.Refresh
  
    For i = 1 To Adodc4.Recordset.RecordCount

        If Adodc4.Recordset.Fields!MEMBER_TYPE <> 1 Then
            'Adodc2.CommandType = adCmdText
            'Adodc2.RecordSource = "SELECT SUM([VALUE]) AS TOTAL_ACTIVITY FROM member_activity WHERE MEMBER_ID=" & Adodc1.Recordset.Fields!member_id & "AND MEMBER_CHILD_ID=" & Adodc4.Recordset.Fields!MEMBER_CHILD_ID
            'Adodc2.Refresh
  
            Adodc11.CommandType = adCmdText
            'Adodc11.RecordSource = "select*  FROM member_TYPES WHERE MEMBER_NAME='n" & Adodc4.Recordset.Fields!MEMBER_TYPE & "'"
            Adodc11.RecordSource = "SELECT     * from MEMBER_TYPES WHERE     (MEMBER_NAME = '⁄«∆·Ï   «»⁄')"
            Adodc11.Refresh
        
            Adodc9.Recordset.AddNew '··⁄÷Ê «· «»⁄
            Adodc9.Recordset.Fields!member_id = Adodc1.Recordset.Fields!member_id
            Adodc9.Recordset.Fields!CHILD_ID = Adodc4.Recordset.Fields!MEMBER_CHILD_ID
            Adodc9.Recordset.Fields!CHILD_NAME = Adodc4.Recordset.Fields!MEMBER_CHILD_NAME
            Adodc9.Recordset.Fields!MEMBER_TITLE = Adodc4.Recordset.Fields!MEMBER_TITLE  ' cc "⁄÷Ê⁄«„·"

            If Adodc4.Recordset.Fields!MEMBER_TITLE = "«·“ÊÃ…" Then
                Adodc9.Recordset.Fields!MEMBER_VALUE = Adodc12.Recordset.Fields![value]
                Adodc9.Recordset.Fields!wife = 1
            Else
                ' Adodc9.Recordset.Fields!MEMBER_VALUE = Adodc11.Recordset.Fields![Value]
                Adodc9.Recordset.Fields!MEMBER_VALUE = 2400 'Adodc11.Recordset.Fields![Value]
            End If
      
            Adodc9.Recordset.Fields!member_card_value = CARD_VALUE.text 'À„‰ «·ﬂ«—‰Ì…
            '      If Not IsNull(Adodc2.Recordset.Fields!TOTAL_ACTIVITY) <> 0 Then
            '       Adodc9.Recordset.Fields!activity_value = Adodc2.Recordset.Fields!TOTAL_ACTIVITY
            '       Else
            '       Adodc9.Recordset.Fields!activity_value = 0
            '      End If
        
            Adodc9.Recordset.update
        End If

        Adodc4.Recordset.MoveNext
    Next i

    MsgBox " „  ÃœÌœ «·€÷ÊÌ…", vbInformation
    operatiomn_update_frm.Show
    Unload Me

End Sub

Function ADD_MEMBER_FINES()
    Adodc5.Recordset.AddNew

    Adodc5.Recordset.Fields!member_id = Text1.text
    Adodc5.Recordset.Fields!member_name = Text4.text

    Adodc5.Recordset.Fields!FINES_TYPE = "€—«„…  √ŒÌ—"
    Adodc5.Recordset.Fields!FINE_DATE = DateValue(Now)
    Adodc5.Recordset.Fields!Fine_nO_OF_YEAR = 5
    Adodc5.Recordset.Fields!FINES_TOTAL = fines_value.text
 
    Adodc5.Recordset.Fields!fines_value = fines_value.text
    
    Adodc5.Recordset.update
 
    Dim i As Integer

    For i = 1 To 5
        Adodc10.Recordset.AddNew
        Adodc10.Recordset.Fields!fines_no = i
        Adodc10.Recordset.Fields!member_id = Text1.text
        Adodc10.Recordset.Fields!member_name = Text4.text
        Adodc10.Recordset.Fields!fines_value = val(fines_value) / 5

        Adodc10.Recordset.Fields!FINES_TYPE = "€—«„…  √ŒÌ—"
        Adodc10.Recordset.Fields!FINES_TOTAL = fines_value.text
        Adodc10.Recordset.Fields!FINES_DATE = DateValue(Now)
        Adodc10.Recordset.update
     
    Next i
 
End Function

Private Sub DataCombo1_Click(Area As Integer)
 
End Sub

Private Sub Command3_Click()
    member_search.Show
    member_search.from = 7
 
End Sub

Private Sub Form_Load()
    system_path = App.path ' "D:\my works\accountant\28  01 2011\SourceCode\SourceCode"
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT * FROM MEMBERS WHERE MEMBER_ID=0 "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from  member_activity "
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  this_year "
    Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select * from  CARD "
    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from FineS "
    Adodc5.Refresh

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    Adodc6.RecordSource = "select * from Fine_TYPES where Fines_ID=2 "
    Adodc6.Refresh

    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "select * from MEMBERS  "
    Adodc7.Refresh

    Adodc8.ConnectionString = connection_string
    Adodc8.CommandType = adCmdText
    Adodc8.RecordSource = "select * from   OPERATIONS"
    Adodc8.Refresh

    Adodc9.ConnectionString = connection_string
    Adodc9.CommandType = adCmdText
    Adodc9.RecordSource = "select * from  OPERATION_DETAILS "
    Adodc9.Refresh

    Adodc10.ConnectionString = connection_string
    Adodc10.CommandType = adCmdText
    Adodc10.RecordSource = "select * from FINES_DETAILS "
    Adodc10.Refresh

    Adodc11.ConnectionString = connection_string
    Adodc11.CommandType = adCmdText
    Adodc11.RecordSource = "select * from  OPERATIONS"
    Adodc11.Refresh

    Adodc12.ConnectionString = connection_string
    Adodc12.CommandType = adCmdText
    Adodc12.RecordSource = "select * from  OPERATIONS "
    Adodc12.Refresh

    Adodc20.ConnectionString = connection_string
    Adodc20.CommandType = adCmdText
    Adodc20.RecordSource = "select * from   member_activity"
    Adodc20.Refresh

End Sub

Private Sub Text1_Change()

    If Text1.text = "" Or Not IsNumeric(Text1.text) Then Exit Sub
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM members WHERE MEMBER_ID=" & Text1.text
    Adodc1.Refresh

    DBPix1.ImageViewFile (system_path & "\IMAGES\" & Text27.text & ".JPG")

End Sub

