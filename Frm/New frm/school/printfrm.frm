VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form printfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4428
   ClientLeft      =   108
   ClientTop       =   408
   ClientWidth     =   6204
   LinkTopic       =   "Form1"
   ScaleHeight     =   4428
   ScaleWidth      =   6204
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ÿ»«⁄…  «·ﬂ·"
      Height          =   432
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   2652
   End
   Begin VB.CheckBox Check1 
      Caption         =   " ÕœÌœ ··ÿ»«⁄…"
      DataField       =   "SELECTED"
      DataSource      =   "Adodc1"
      Height          =   192
      Left            =   4320
      TabIndex        =   13
      Top             =   3000
      Width           =   1812
   End
   Begin DBPIXLib.DBPix20 DBPIX1 
      DataField       =   "IMAGE_PATH"
      DataSource      =   "Adodc1"
      Height          =   1452
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   2561
      _StockProps     =   1
      BackColor       =   -2147483633
      _Image          =   "printfrm.frx":0000
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   492
      Left            =   480
      Top             =   3360
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   868
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ready_to_print"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "ÿ»«⁄… «·„Õœœ ›ﬁÿ"
      Height          =   432
      Left            =   3480
      TabIndex        =   0
      Top             =   3960
      Width           =   2652
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      DataField       =   "CENTER_MANAGER"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1200
      TabIndex        =   12
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "„œÌ— ⁄«„ «·„—ﬂ“"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1080
      TabIndex        =   11
      Top             =   2280
      Width           =   1212
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "update_year"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   2280
      TabIndex        =   10
      Top             =   2280
      Width           =   2532
   End
   Begin VB.Shape Shape1 
      Height          =   2772
      Left            =   360
      Top             =   120
      Width           =   5772
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·«Ì’«·"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   4800
      TabIndex        =   9
      Top             =   2040
      Width           =   1212
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "BILL_NO"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2520
      TabIndex        =   8
      Top             =   2040
      Width           =   2052
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„”·”·"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   4800
      TabIndex        =   7
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "inedx"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2520
      TabIndex        =   6
      Top             =   1680
      Width           =   2052
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·⁄÷“Ì…"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   4800
      TabIndex        =   5
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   2052
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_TYPE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   2052
   End
End
Attribute VB_Name = "printfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
