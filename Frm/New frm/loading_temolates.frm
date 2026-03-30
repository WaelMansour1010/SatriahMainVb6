VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form loading_temolates 
   BorderStyle     =   0  'None
   Caption         =   " ⁄œÌ· «·’Ê—"
   ClientHeight    =   10980
   ClientLeft      =   105
   ClientTop       =   75
   ClientWidth     =   20010
   DrawStyle       =   3  'Dash-Dot
   DrawWidth       =   2
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   20010
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin DBPIXLib.DBPix20 DBPix201 
      Height          =   1215
      Left            =   2760
      TabIndex        =   53
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   2143
      _StockProps     =   1
      _Image          =   "loading_temolates.frx":0000
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   49
      Top             =   3840
      Width           =   1455
      Begin VB.CommandButton Command19 
         Caption         =   "L"
         Height          =   495
         Left            =   840
         TabIndex        =   52
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command18 
         Caption         =   "P"
         Height          =   495
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "« Ã«… «·’›Õ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   0
      TabIndex        =   47
      Top             =   8520
      Width           =   1455
      Begin VB.CommandButton Command12 
         Caption         =   "ÿ»«⁄…"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1455
      Left            =   0
      TabIndex        =   40
      Top             =   4800
      Width           =   1455
      Begin VB.CommandButton Command17 
         Caption         =   "ﬁ·„"
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command33 
         Caption         =   "‰’"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "‰Ê⁄ «·⁄„·Ì…"
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
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   0
      TabIndex        =   37
      Top             =   3000
      Width           =   1455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Text            =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "«·’›Õ… «·Õ«·Ì…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   2415
      Left            =   0
      TabIndex        =   31
      Top             =   600
      Width           =   1455
      Begin VB.CommandButton Command13 
         Caption         =   "«··Ê‰"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1800
         Width           =   915
      End
      Begin VB.CommandButton Command10 
         Caption         =   "ÕÃ„ «·ŒŸ"
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "«⁄œ«œ«  «·Õﬁ·"
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
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ÕÃ„ «·Õﬁ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "·Ê‰ «·Œ·›Ì…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   915
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   25
      Top             =   8760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text2 
      DataField       =   "templates_id"
      DataSource      =   "templates_details"
      Height          =   285
      Left            =   1920
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      DataField       =   "templates_id"
      DataSource      =   "templates"
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Left            =   0
      TabIndex        =   16
      Top             =   6240
      Width           =   1455
      Begin VB.CommandButton Command21 
         BackColor       =   &H000000FF&
         Caption         =   "Õ›Ÿ2"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Õ›Ÿ"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         Caption         =   " Õ„Ì· ‰„Ê–Ã"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Õ›Ÿ ﬂ‰„Ê–Ã"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_save 
         Caption         =   "Õ›Ÿ"
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "‰„Ê–Ã ÃœÌœ"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÿ»«⁄…  «·Œ·›Ì…"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   4080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   -2160
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         Picture         =   "loading_temolates.frx":0018
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Picture         =   "loading_temolates.frx":04BE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         Picture         =   "loading_temolates.frx":0954
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         Picture         =   "loading_temolates.frx":0DE9
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   5
         Height          =   615
         Left            =   1320
         Top             =   840
         Width           =   735
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   5
         Height          =   615
         Left            =   0
         Top             =   840
         Width           =   735
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   5
         Height          =   735
         Left            =   720
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   5
         Height          =   735
         Left            =   720
         Top             =   120
         Width           =   735
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   720
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   3720
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Õ–› «·‰’"
      Height          =   495
      Left            =   -120
      TabIndex        =   9
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   16440
      ScaleHeight     =   675
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   17280
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "-"
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
      Left            =   360
      TabIndex        =   5
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "+"
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
      Left            =   720
      TabIndex        =   4
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "-"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   10320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+"
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
      Left            =   840
      TabIndex        =   2
      Top             =   10560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "«Œ›«¡ «·‰’"
      Height          =   375
      Left            =   -120
      TabIndex        =   1
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000007&
      Height          =   3375
      Left            =   2040
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   0
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc templates 
      Height          =   375
      Left            =   3240
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc templates_details 
      Height          =   375
      Left            =   2640
      Top             =   3480
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3000
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      Height          =   375
      Left            =   3120
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
   Begin VB.Label SUBJECT_NO 
      BackColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   45
      Top             =   10440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   30
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   5400
      TabIndex        =   29
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbl_templates_id 
      Caption         =   "Label9"
      DataField       =   "templates_id"
      DataSource      =   "templates"
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      Height          =   735
      Left            =   -120
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   3420
      Left            =   16680
      Picture         =   "loading_temolates.frx":1285
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "„Õ«–«… «·‰’"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   -1920
      TabIndex        =   10
      Top             =   9960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   495
      Left            =   120
      Top             =   11280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label IMAGE_NAME 
      Alignment       =   2  'Center
      Caption         =   "Label5"
      Height          =   495
      Left            =   16800
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "«”„ «·’Ê—…"
      Height          =   375
      Left            =   17280
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   9825
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13515
   End
End
Attribute VB_Name = "loading_temolates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Dim edit_index As Integer
Dim image_direction As String
Dim MAIN_INDEX As Integer
Dim image_path(1000) As String
Dim TEXT_CONTENT As String
Dim CURRENT_CHAR As String
Dim pen As Boolean

Private TXTsend(1000) As TextBox
Dim TEXTCOUNT As Integer
Dim TxtName As String
Dim check As Boolean
Dim txtindex As Integer
Dim X1, Y1 As Integer

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38

Private Const RC_PALETTE As Long = &H100

Private Const SIZEPALETTE As Long = 104

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Declare Function CreateCompatibleDC _
                Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleBitmap _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long) As Long

Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal iCapabilitiy As Long) As Long

Private Declare Function GetSystemPaletteEntries _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal wStartIndex As Long, _
                             ByVal wNumEntries As Long, _
                             lpPaletteEntries As PALETTEENTRY) As Long

Private Declare Function CreatePalette _
                Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long

Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long

Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDCDest As Long, _
                             ByVal XDest As Long, _
                             ByVal YDest As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hDCSrc As Long, _
                             ByVal XSrc As Long, _
                             ByVal YSrc As Long, _
                             ByVal dwRop As Long) As Long

Private Declare Function DeleteDC _
                Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function GetForegroundWindow _
                Lib "user32" () As Long

Private Declare Function SelectPalette _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hPalette As Long, _
                             ByVal bForceBackground As Long) As Long

Private Declare Function RealizePalette _
                Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function GetWindowDC _
                Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDC _
                Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetWindowRect _
                Lib "user32" (ByVal hWnd As Long, _
                              lpRect As RECT) As Long

Private Declare Function ReleaseDC _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hDC As Long) As Long

Private Declare Function GetDesktopWindow _
                Lib "user32" () As Long

Private Type PicBmp
    Size As Long
    type As Long
        hBmp As Long
        hPal As Long
        reserved As Long
    End Type

    Private Declare Function OleCreatePictureIndirect _
                    Lib "olepro32.dll" (PicDesc As PicBmp, _
                                        RefIID As GUID, _
                                        ByVal fPictureOwnsHandle As Long, _
                                        Ipic As IPicture) As Long

Private Sub Combo1_Click()
    On Error Resume Next
    Image1.Picture = LoadPicture(system_path & "\images\" & image_path(Combo1.ListIndex + 1))

    IMAGE_NAME.Caption = image_path(Combo1.ListIndex + 1)

    Dim txtindex, i As Integer

    Dim astrSplitItems() As String

    'On Error Resume Next
    If List1.ListCount > 0 Then

        For i = 0 To List1.ListCount - 1
            
            astrSplitItems = Split(List1.List(i), "-")
            txtindex = val(astrSplitItems(0))
            'txtindex = Val(Mid(List1.List(i), 1, 1))
            ' Unload Text3(txtindex)
            Text3(txtindex).Visible = False

            DoEvents
            Shape1.Visible = False
            'List1.List(List1.ListIndex) = "##Deleted##"
            
        Next i

    End If

    List1.Clear

    'Dim i As Integer
    templates_details.CommandType = adCmdText
    templates_details.RecordSource = "select * from templates_detailsnew where templates_id=" & Label6.Caption & "and image_id=" & Combo1.text
    templates_details.Refresh

    If templates_details.Recordset.RecordCount > 0 Then
        templates_details.Recordset.MoveFirst

        If templates_details.Recordset.Fields!image_direction = "p" Then
            Command18_Click
        Else

            If templates_details.Recordset.Fields!image_direction = "l" Then
                Command19_Click
            End If
        End If

    End If

    Static Index As Integer
    'MAIN_INDEX = 0
        
    'Dim i As Integer
    Dim X1, X2, Y1, Y2 As Integer
    Dim text, color, BackColor, FontName, FontSize, FontBold, FontItalic, FontUnderline, Strikethrough As String
    List1.Clear

    For i = 1 To templates_details.Recordset.RecordCount
        X1 = templates_details.Recordset.Fields!X1
        Y1 = templates_details.Recordset.Fields!Y1
        X2 = templates_details.Recordset.Fields!X2
        Y2 = templates_details.Recordset.Fields!Y2
        text = templates_details.Recordset.Fields!text
        color = templates_details.Recordset.Fields!color
        BackColor = templates_details.Recordset.Fields!BackColor

        FontName = templates_details.Recordset.Fields!FontName
        FontSize = templates_details.Recordset.Fields!FontSize
        FontBold = templates_details.Recordset.Fields!FontBold
        FontItalic = templates_details.Recordset.Fields!FontItalic
        FontUnderline = templates_details.Recordset.Fields!FontUnderline
        Strikethrough = templates_details.Recordset.Fields!Strikethrough

        'TXTNAME = "AA" & Trim(Str(TEXTCOUNT))
        'Set TXTsend(TEXTCOUNT) = loading_temolates.Controls.Add("VB.TextBox", TXTNAME)
        '  Index = i ' Index + 1
        MAIN_INDEX = MAIN_INDEX + 1
        Load Text3(MAIN_INDEX)
 
        With Text3(MAIN_INDEX)
            .top = X1
            .left = Y1
            .Width = X2
            .Height = Y2
            .text = text
            .Visible = True
            .BorderStyle = 0
            .Alignment = 1
            .RightToLeft = True
            .ForeColor = color
            .BackColor = BackColor ' Label1.backcolor
            .Font.name = FontName
            .Font.Size = FontSize
            .Font.Bold = FontBold
            .Font.Italic = FontItalic
            .Font.Underline = FontUnderline
            .Font.Strikethrough = Strikethrough
            List1.AddItem MAIN_INDEX & "-" & .top & "-" & .left & "-" & .Width & "-" & .Height
            List1.ListIndex = List1.ListCount - 1

        End With

        'TEXTCOUNT = TEXTCOUNT + 1

        templates_details.Recordset.MoveNext
    Next i

    'image_path(i) = CommonDialog1.FileTitle
End Sub

Private Sub Command1_Click()

    Dim astrSplitItems() As String

    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))

    Text3(MAIN_INDEX).left = Text3(MAIN_INDEX).left + 30
    change_object_data
End Sub

Private Sub Command10_Click()
    'Dim txtindex As Integer
    'txtindex = edit_index ' Val(Mid(List1.List(List1.ListIndex), 1, 1))
 
    dlgFont.Flags = cdlCFScreenFonts
    dlgFont.CancelError = True
    On Error Resume Next
    dlgFont.ShowFont

    If Err = cdlCancel Then Exit Sub

    ' Use the dialog's properties.
    With dlgFont
        Text3(edit_index).Font.name = .FontName
        Text3(edit_index).Font.Size = .FontSize
        Text3(edit_index).Font.Bold = .FontBold
        Text3(edit_index).Font.Italic = .FontItalic
        Text3(edit_index).Font.Underline = .FontUnderline
        Text3(edit_index).Font.Strikethrough = .FontStrikethru
    End With

End Sub

Private Sub Command11_Click()
    Dim txtindex, i As Integer

    'MAIN_INDEX = 0
    Dim x As String
    Dim Y As String
    x = InputBox("„‰ ›÷·ﬂ Õœœ ⁄œœ «Ê—«ﬁ «·„” ‰œ", vbExclamation)

    If Not IsNumeric(x) Then
        MsgBox "·«»œ „‰ ﬂ «»… «—ﬁ«„ ›ﬁÿ ", vbCritical
        Exit Sub
    End If

    Combo1.Clear

    For i = 1 To x
        Combo1.AddItem i
    Next i

    Y = InputBox("Õœœ «”„ «·‰„Ê–Ã", vbExclamation)
    templates.Recordset.AddNew
    templates.Recordset.Fields!templates_name = Y
    'templates.Recordset.Fields!IMAGE_NAME = "T#" & template_name & "-" & Mid(Date, 1, 2) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 7, 4) & ".JPG"
    'templates.Recordset.Fields!departement_name = dep_name
    templates.Recordset.Fields!no_of_images = x

    templates.Recordset.update
    templates.Recordset.MoveLast
    Label6.Caption = lbl_templates_id.Caption
    'Dim txtindex, i As Integer

    Dim astrSplitItems() As String

    If List1.ListCount > 0 Then

        For i = 0 To List1.ListCount - 1

            astrSplitItems = Split(List1.List(i), "-")
            txtindex = val(astrSplitItems(0))
            'txtindex = Val(Mid(List1.List(i), 1, 1))
            ' Text3(txtindex).Visible = False
            Text3(txtindex).Visible = False
             
            Shape1.Visible = False
            'List1.List(List1.ListIndex) = "##Deleted##"
            
        Next i

    End If

    List1.Clear

    Image1.Picture = Nothing

    For i = 1 To x
        MsgBox "«œŒ· «·’Ê—… —ﬁ„" & i
        CommonDialog1.ShowOpen
        Image1.Picture = LoadPicture(CommonDialog1.FileName)
        IMAGE_NAME.Caption = CommonDialog1.FileTitle

        image_path(i) = CommonDialog1.FileTitle

        DBPix201.ImageLoadFile (CommonDialog1.FileName)
        DBPix201.ImageSaveFile (system_path & "\images\" & CommonDialog1.FileTitle)

        DoEvents
        DoEvents
        DoEvents
    Next i

    new_templates.Show
    new_templates.Text1 = Y
    new_templates.id = templates.Recordset.Fields!templates_id
End Sub

Private Sub Command12_Click()

    If Check1.value = 0 Then

        txtlostfocusfun
    Else
        Shape1.Visible = False
    End If

    DoEvents

    DoEvents
    Set Picture1.Picture = CaptureScreen()

    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
End Sub

Private Sub Command13_Click()

    ' Dim txtindex As Integer
 
    'Dim astrSplitItems() As String
    'astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    'txtindex = Val(astrSplitItems(0))

    dlgFont.ShowColor
    Text3(edit_index).ForeColor = dlgFont.color
    Command13.BackColor = dlgFont.color
End Sub

Private Sub Command14_Click()
    Dim txtindex As Integer
    Dim astrSplitItems() As String

    If List1.ListIndex >= 0 Then

        astrSplitItems = Split(List1.List(List1.ListIndex), "-")

        txtindex = val(astrSplitItems(0))

        Text3(MAIN_INDEX).Visible = False
        Shape1.Visible = False
        List1.List(List1.ListIndex) = "##Deleted##"
    End If

End Sub

Public Function save_templates(template_name As String, _
                               dep_name As String)

    Dim i As Integer

    txtlostfocusfun
    Shape1.Visible = False

    DoEvents

    'Dim template_name As String
    'template_name = InputBox("«ﬂ » «”„ «·‰„Ê–Ã")

    templates.Recordset.AddNew
    templates.Recordset.Fields!templates_name = template_name
    'templates.Recordset.Fields!IMAGE_NAME = "T#" & template_name & "-" & Mid(Date, 1, 2) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 7, 4) & ".JPG"
    templates.Recordset.Fields!departement_name = dep_name

    templates.Recordset.update

    For i = 0 To List1.ListCount - 1

        If List1.List(i) = "##Deleted##" Then GoTo ll
        Dim astrSplitItems() As String

        astrSplitItems = Split(List1.List(i), "-")
        'astrSplitItems (0)
        templates_details.Recordset.AddNew
        templates_details.Recordset.Fields!templates_id = templates.Recordset.Fields!templates_id
        templates_details.Recordset.Fields!X1 = astrSplitItems(1)
        templates_details.Recordset.Fields!Y1 = astrSplitItems(2)
        templates_details.Recordset.Fields!X2 = astrSplitItems(3)
        templates_details.Recordset.Fields!Y2 = astrSplitItems(4)
        templates_details.Recordset.Fields!Index = astrSplitItems(0)
        templates_details.Recordset.update
ll:
    Next i

    'Set Picture1.Picture = CaptureScreen()
   
    Picture1.Picture = loading_temolates.Image1.Picture
   
    SavePicture Picture1.Picture, system_path & "\templates\" & "T#" & template_name & "-" & Mid(Date, 1, 2) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 7, 4) & ".JPG"

End Function

Private Sub Command15_Click()
    txtlostfocusfun
    new_templates.Show
    'save_templates
End Sub

Private Sub Command16_Click()
    Image1.Picture = Nothing

    Dim txtindex, i As Integer

    Dim astrSplitItems() As String

    If List1.ListCount > 0 Then

        For i = 0 To List1.ListCount - 1

            astrSplitItems = Split(List1.List(i), "-")
            txtindex = val(astrSplitItems(0))
            'txtindex = Val(Mid(List1.List(i), 1, 1))
            Text3(txtindex).Visible = False
            Shape1.Visible = False
            'List1.List(List1.ListIndex) = "##Deleted##"

        Next i

    End If

    List1.Clear
    'Image1.Picture = Null
    frm_templates.Show

End Sub

Private Sub Command17_Click()
    pen = True
End Sub

Private Sub Command18_Click()
    Image1.Width = 9945
    Image1.Height = 11025
    image_direction = "p"
End Sub

Private Sub Command19_Click()
    Image1.Width = 13515
    Image1.Height = 9825
    image_direction = "l"
End Sub

Private Sub Command2_Click()
    Dim txtindex As Integer
    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))

    Text3(MAIN_INDEX).Visible = False
    Shape1.Visible = False
End Sub

Private Sub Command20_Click()

    'On Error Resume Next
    If Combo1.text = "" Then Exit Sub

    If List1.ListCount = 0 Then MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ·⁄œ„ ÊÃÊœ ÕﬁÊ· ›Ì «·‰„Ê–Ã", vbCritical: Exit Sub
    Dim i As Integer
    templates_details.CommandType = adCmdText
    templates_details.RecordSource = "select * from templates_detailsnew where templates_id=" & Label6.Caption & "and image_id=" & Combo1.text
    templates_details.Refresh

    For i = 1 To templates_details.Recordset.RecordCount
        templates_details.Recordset.delete
        templates_details.Recordset.MoveNext

    Next i

    For i = 0 To List1.ListCount - 1

        If List1.List(i) = "##Deleted##" Then GoTo ll
        Dim astrSplitItems() As String

        astrSplitItems = Split(List1.List(i), "-")
        'astrSplitItems (0)
        templates_details.Recordset.AddNew
        templates_details.Recordset.Fields!templates_id = Label6.Caption
        templates_details.Recordset.Fields!IMAGE_NAME = image_path(Combo1.ListIndex + 1)
        templates_details.Recordset.Fields!image_id = Combo1.text
        templates_details.Recordset.Fields!image_direction = image_direction

        templates_details.Recordset.Fields!X1 = astrSplitItems(1)
        templates_details.Recordset.Fields!Y1 = astrSplitItems(2)
        templates_details.Recordset.Fields!X2 = astrSplitItems(3)
        templates_details.Recordset.Fields!Y2 = astrSplitItems(4)
        templates_details.Recordset.Fields!Index = astrSplitItems(0)
        templates_details.Recordset.Fields!text = Text3(astrSplitItems(0)).text
        templates_details.Recordset.Fields!color = Text3(astrSplitItems(0)).ForeColor
        templates_details.Recordset.Fields!BackColor = Text3(astrSplitItems(0)).BackColor

        templates_details.Recordset.Fields!FontName = Text3(astrSplitItems(0)).Font.name
        templates_details.Recordset.Fields!FontSize = Text3(astrSplitItems(0)).Font.Size
        templates_details.Recordset.Fields!FontBold = Text3(astrSplitItems(0)).Font.Bold
        templates_details.Recordset.Fields!FontItalic = Text3(astrSplitItems(0)).Font.Italic
        templates_details.Recordset.Fields!FontUnderline = Text3(astrSplitItems(0)).Font.Underline
        templates_details.Recordset.Fields!Strikethrough = Text3(astrSplitItems(0)).Font.Strikethrough

        templates_details.Recordset.update
ll:
    Next i

End Sub

Private Sub Command21_Click()
    Dim i As Integer
    templates_details.CommandType = adCmdText
    templates_details.RecordSource = "select * from templates_detailsnew where templates_id=" & Label6.Caption & "and image_id=" & Combo1.text
    templates_details.Refresh

    For i = 1 To templates_details.Recordset.RecordCount
        templates_details.Recordset.delete
        templates_details.Recordset.MoveNext

    Next i

    For i = 0 To List1.ListCount - 1

        If List1.List(i) = "##Deleted##" Then GoTo ll
        Dim astrSplitItems() As String

        astrSplitItems = Split(List1.List(i), "-")
        'astrSplitItems (0)
        templates_details.Recordset.AddNew
        templates_details.Recordset.Fields!templates_id = Label6.Caption
        templates_details.Recordset.Fields!IMAGE_NAME = image_path(Combo1.ListIndex + 1)
        templates_details.Recordset.Fields!image_id = Combo1.text
        templates_details.Recordset.Fields!image_direction = image_direction

        templates_details.Recordset.Fields!X1 = astrSplitItems(1)
        templates_details.Recordset.Fields!Y1 = astrSplitItems(2)
        templates_details.Recordset.Fields!X2 = astrSplitItems(3)
        templates_details.Recordset.Fields!Y2 = astrSplitItems(4)
        templates_details.Recordset.Fields!Index = astrSplitItems(0)
        templates_details.Recordset.Fields!text = Text3(astrSplitItems(0)).text
        templates_details.Recordset.Fields!color = Text3(astrSplitItems(0)).ForeColor
        templates_details.Recordset.Fields!BackColor = Text3(astrSplitItems(0)).BackColor

        templates_details.Recordset.Fields!FontName = Text3(astrSplitItems(0)).Font.name
        templates_details.Recordset.Fields!FontSize = Text3(astrSplitItems(0)).Font.Size
        templates_details.Recordset.Fields!FontBold = Text3(astrSplitItems(0)).Font.Bold
        templates_details.Recordset.Fields!FontItalic = Text3(astrSplitItems(0)).Font.Italic
        templates_details.Recordset.Fields!FontUnderline = Text3(astrSplitItems(0)).Font.Underline
        templates_details.Recordset.Fields!Strikethrough = Text3(astrSplitItems(0)).Font.Strikethrough

        templates_details.Recordset.update
ll:
    Next i

End Sub

Private Sub Command3_Click()
    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))

    Text3(MAIN_INDEX).left = Text3(MAIN_INDEX).left - 30
    change_object_data
End Sub

Private Sub Command33_Click()
    pen = False
    CURRENT_CHAR = "T"
End Sub

Private Sub Command4_Click()
    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))

    Text3(MAIN_INDEX).top = Text3(MAIN_INDEX).top - 30
    change_object_data
End Sub

Private Sub Command5_Click()
    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))

    Text3(MAIN_INDEX).top = Text3(MAIN_INDEX).top + 30
    change_object_data
End Sub

Private Sub Command6_Click()
    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))
    Text3(MAIN_INDEX).Width = Text3(MAIN_INDEX).Width + 30
    change_object_data
End Sub

Private Sub Command7_Click()
    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))

    Text3(MAIN_INDEX).Width = Text3(MAIN_INDEX).Width - 30
    change_object_data
End Sub

Private Sub Command8_Click()

    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))
    Text3(MAIN_INDEX).Height = Text3(MAIN_INDEX).Height + 30
    change_object_data
End Sub

Private Sub Command9_Click()
    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))
    Text3(MAIN_INDEX).Height = Text3(MAIN_INDEX).Height - 30
    change_object_data
End Sub

Private Sub Form_Load()
    system_path = App.path
    On Error Resume Next
    'LoadSettings
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from  templates_detailsnew"
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from  templates_detailsnew"
    Adodc2.Refresh
 
    templates.ConnectionString = connection_string
    templates.CommandType = adCmdText
    templates.RecordSource = "select * from  templatesnew"
    templates.Refresh

    templates_details.ConnectionString = connection_string
    templates_details.CommandType = adCmdText
    templates_details.RecordSource = "select * from  templates_detailsnew"
    templates_details.Refresh

    pen = False
    '    log_files_form.Adodc1.Recordset.AddNew: log_files_form.Adodc1.Recordset.Fields!log_date = Date
    '    log_files_form.Adodc1.Recordset.Fields!log_time = Time
    '    log_files_form.Adodc1.Recordset.Fields!User_Name = login.Adodc1.Recordset.Fields!name
    '
    '      log_files_form.Adodc1.Recordset.Fields!process_name = "œŒÊ· «·Ï      ‘«‘…  " & Me.Caption
    '       log_files_form.Adodc1.Recordset.Fields!process_text = ""
    '
    '        log_files_form.Adodc1.Recordset.update: DoEvents
    TEXTCOUNT = 0
    check = False
End Sub

Function change_object_data2(Index As Integer)
    Dim i As Integer
    Dim astrSplitItems() As String

    For i = 0 To List1.ListCount - 1

        astrSplitItems = Split(List1.List(i), "-")

        txtindex = val(astrSplitItems(0))

        If txtindex = Index Then

            List1.List(i) = txtindex & "-" & Text3(Index).top & "-" & Text3(Index).left & "-" & Text3(Index).Width & "-" & Text3(Index).Height
            Shape1.top = Text3(Index).top
            Shape1.left = Text3(Index).left
            Shape1.Width = Text3(Index).Width
            Shape1.Height = Text3(Index).Height
 
        End If

    Next i

    'Shape1.Visible = True
 
    'TEXT3(MAIN_index).Height = TEXT3(MAIN_index).Height - 30
End Function

Function change_object_data()
    Dim i As Integer

    For i = 1 To Len(List1.List(List1.ListIndex))

        If Mid(List1.List(List1.ListIndex), i, 1) = "-" Then
            GoTo ll
        End If

    Next i

ll:

    Dim astrSplitItems() As String
    astrSplitItems = Split(List1.List(List1.ListIndex), "-")

    txtindex = val(astrSplitItems(0))

    If Text3(MAIN_INDEX).Width = 40 And Text3(MAIN_INDEX).Height = 40 Then Exit Function
    List1.List(List1.ListIndex) = txtindex & "-" & Text3(MAIN_INDEX).top & "-" & Text3(MAIN_INDEX).left & "-" & Text3(MAIN_INDEX).Width & "-" & Text3(MAIN_INDEX).Height
    Shape1.top = Text3(MAIN_INDEX).top
    Shape1.left = Text3(MAIN_INDEX).left
    Shape1.Width = Text3(MAIN_INDEX).Width
    Shape1.Height = Text3(MAIN_INDEX).Height
 
    Shape1.Visible = True
 
    'TEXT3(MAIN_index).Height = TEXT3(MAIN_index).Height - 30
End Function

Private Sub TXTsend_DragDrop(Index As Integer, _
                             Source As Control, _
                             x As Single, _
                             Y As Single)
    TXTsend(Index).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' log_files_form.Adodc1.Recordset.AddNew: log_files_form.Adodc1.Recordset.Fields!log_date = Date
    ' log_files_form.Adodc1.Recordset.Fields!log_time = Time
    ' log_files_form.Adodc1.Recordset.Fields!User_Name = login.Adodc1.Recordset.Fields!name

    'log_files_form.Adodc1.Recordset.Fields!process_name = "  Œ—ÊÃ „‰  ‘«‘…" & Me.Caption
    ' log_files_form.Adodc1.Recordset.Fields!process_text = ""
    '
    '  log_files_form.Adodc1.Recordset.update: DoEvents
End Sub

Function loading_templates_data()
    'Dim oCtl As Control
    Static Index As Integer
        
    Dim i As Integer
    Dim X1, X2, Y1, Y2 As Integer
    List1.Clear

    For i = 1 To templates_details.Recordset.RecordCount
        X1 = templates_details.Recordset.Fields!X1
        Y1 = templates_details.Recordset.Fields!Y1
        X2 = templates_details.Recordset.Fields!X2
        Y2 = templates_details.Recordset.Fields!Y2
        'TXTNAME = "AA" & Trim(Str(TEXTCOUNT))
        'Set TXTsend(TEXTCOUNT) = loading_temolates.Controls.Add("VB.TextBox", TXTNAME)
        Index = Index + 1
        MAIN_INDEX = Index
        Load Text3(MAIN_INDEX)
 
        With Text3(MAIN_INDEX)
            .top = X1
            .left = Y1
            .Width = X2
            .Height = Y2

            .Visible = True
            .BorderStyle = 0
            .Alignment = 1
            .RightToLeft = True

            .BackColor = Label1.BackColor
  
            List1.AddItem MAIN_INDEX & "-" & .top & "-" & .left & "-" & .Width & "-" & .Height
            List1.ListIndex = List1.ListCount - 1

        End With

        'TEXTCOUNT = TEXTCOUNT + 1

        templates_details.Recordset.MoveNext
    Next i

    'If Button = 1 Then
    'If CHECK = False Then
    'X1 = x: Y1 = y: CHECK = True

    'Else

    'Print "X:" & X & "   Y:" & Y

    'TEXTCOUNT = TEXTCOUNT + 1
    '
    'CHECK = False
    'End If
    'End If

End Function

Private Sub Image1_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             Y As Single)
    Static Index As Integer
 
    'Index = Index + 1
    'MAIN_INDEX = Index
    'Load Text3(MAIN_INDEX)
    'TEXT3(MAIN_index).Visible = True
    'Text3Index).Move Text1(Index - 1).Left, Text1(Index - 1).Top + Text1(Index - 1).Height

    'Dim oCtl As Control

    If pen = True Then
        Shape1.Visible = False
        ' Index = Index + 1
        MAIN_INDEX = MAIN_INDEX + 1
        Load Text3(MAIN_INDEX)
         
        With Text3(MAIN_INDEX)
            .left = x
            .top = Y
            .top = .top - Image1.top
            .left = .left + Image1.left
            .Font.Size = 8
            .Width = 150
            .Height = 285
            .Visible = True
            
            .BackColor = Label8.BackColor
        
            List1.AddItem MAIN_INDEX & "-" & .top & "-" & .left & "-" & .Width & "-" & .Height
            List1.ListIndex = List1.ListCount - 1
        End With
         
        Exit Sub
    End If

    If Button = 1 Then
        If check = False Then
            X1 = x: Y1 = Y: check = True

        Else
            Index = Index + 1
            MAIN_INDEX = MAIN_INDEX + 1
            Load Text3(MAIN_INDEX)
 
            With Text3(MAIN_INDEX)

                If X1 < x Then
                    .left = X1

                    If Y1 < Y Then
                        .top = Y1
                    Else
                        .top = Y
                    End If
    
                Else
   
                    .left = x
    
                    If Y1 < Y Then
                        .top = Y1
                    Else
                        .top = Y
  
                    End If
                End If

                .top = .top - Image1.top
                .left = .left + Image1.left

                .Width = Abs(X1 - x)
                .Height = Abs(Y1 - Y)
                .Visible = True
                .BorderStyle = 0
                .Alignment = 1
                .RightToLeft = True
                .SetFocus
                .Font.Size = 12
  
                '   .Text = "........................................"
                .BackColor = Label1.BackColor
                ' .BorderStyle = 0
                List1.AddItem MAIN_INDEX & "-" & .top & "-" & .left & "-" & .Width & "-" & .Height
                List1.ListIndex = List1.ListCount - 1
            End With

            'Print "Abs(X1 - X):" & Abs(X1 - X) & "   Abs(Y1 - Y):" & Abs(Y1 - Y)
            'TEXTCOUNT = TEXTCOUNT + 1

            check = False
        End If
    End If

End Sub

Function txtlostfocusfun()
    On Error Resume Next
    Dim astrSplitItems() As String

    If List1.ListCount = 0 Then Exit Function
    Dim i As Integer

    For i = 0 To List1.ListCount - 1

        astrSplitItems = Split(List1.List(i), "-")

        txtindex = val(astrSplitItems(0))

        If Text3(txtindex).Width = 150 And Text3(txtindex).Height = 285 Then
            GoTo ll
        Else
            Text3(txtindex).BackColor = vbWhite
        End If

ll:
    Next i

    Shape1.Visible = False
End Function

Private Sub Label1_Click()
    On Error Resume Next

    If List1.ListCount = 0 Then Exit Sub
    Dim txtindex As Integer
    Dim i As Integer
    Dim x As Integer
    dlgFont.ShowColor
    Label1.BackColor = dlgFont.color
    Dim astrSplitItems() As String
    x = MsgBox("Â·  —Ìœ  €ÌÌ— ·Ê‰ «·Œ·›Ì… ··‰’Ê’ «·ﬁœÌ„… «Ì÷«", vbExclamation + vbYesNo)

    If x = vbYes Then

        For i = 0 To List1.ListCount - 1

            astrSplitItems = Split(List1.List(i), "-")

            txtindex = val(astrSplitItems(0))
 
            Text3(txtindex).BackColor = Label1.BackColor

        Next i

    End If

    'Shape2.FillColor = Label1.backcolor
    Text3(edit_index).BackColor = dlgFont.color

End Sub

Private Sub Label10_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    CURRENT_CHAR = "W"
End Sub

Private Sub Label3_Click()
    'CURRENT_CHAR = "W"
End Sub

Private Sub Label6_Change()
    'templates_details.CommandType = adCmdText
    'templates_details.RecordSource = "select * from templates_details where templates_id= " & Val(Label6.Caption)
    'templates_details.Refresh
    'If templates_details.Recordset.RecordCount <> 0 Then

    'Call loading_templates_data
    'End If

End Sub

Private Sub Label7_Change()
    On Error Resume Next
    Image1.Picture = LoadPicture(system_path & "\templates\" & Label7.Caption)

    IMAGE_NAME.Caption = CommonDialog1.FileTitle
End Sub

Private Sub Label8_Click()
    dlgFont.ShowColor
    Label8.BackColor = dlgFont.color
    Dim astrSplitItems() As String

End Sub

Private Sub Label9_Change()
    On Error Resume Next
    'CURRENT_CHAR = "T"

    Combo1.Clear
    Dim i As Integer

    For i = 1 To Label9.Caption

        Combo1.AddItem i
    Next i

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = " select distinct  image_id,image_name from templates_detailsnew where templates_id=" & Label6.Caption & " order by image_id"
    Adodc1.Refresh

    For i = 1 To Adodc1.Recordset.RecordCount
        image_path(i) = Adodc1.Recordset.Fields!IMAGE_NAME
        Adodc1.Recordset.MoveNext
    Next i

End Sub

Private Sub List1_Click()
    Dim txtindex As Integer
    Dim i As Integer

    If List1.List(List1.ListIndex) = "##Deleted##" Or List1.ListIndex < 0 Then Exit Sub

    For i = 1 To Len(List1.List(List1.ListIndex))

        If Mid(List1.List(List1.ListIndex), i, 1) = "-" Then
            GoTo ll
        End If

    Next i

ll:

    txtindex = val(Mid(List1.List(List1.ListIndex), 1, i - 1))
    'txtlostfocusfun
    'TEXT3(MAIN_index).BackColor = Label1.BackColor
    change_object_data

End Sub

Private Sub List1_DblClick()
    Dim txtindex As Integer
    txtindex = val(Mid(List1.List(List1.ListIndex), 1, 1))
    'Print txtindex
    Text3(MAIN_INDEX).BackColor = vbWhite
    Text3(MAIN_INDEX).Visible = True
End Sub

' Capture the entire screen.
'Private Sub Command1_Click()
'   Set Picture1.Picture = CaptureScreen()
'End Sub

' Capture the entire form including title and border.
'Private Sub Command2_Click()
'    Set Picture1.Picture = CaptureForm(Me)
'End Sub

' Capture the client area of the form.
'Private Sub Command3_Click()
'    Set Picture1.Picture = CaptureClient(Me)
'End Sub

' Capture the active window after two seconds.
'Private Sub Command4_Click()
'    MsgBox "Two seconds after you close this dialog the active window will be captured."
' Wait for two seconds.
'    Dim EndTime As Date
'    EndTime = DateAdd("s", 2, Now)
'    Do Until Now > EndTime
'       DoEvents
'       Loop
'    Set Picture1.Picture = CaptureActiveWindow()
'    ' Set focus back to form.
'    Me.SetFocus
'End Sub

' Print the current contents of the picture box.
'   Private Sub Command5_Click()
'    PrintPictureToFitPage Printer, Picture1.Picture
'    Printer.EndDoc
'End Sub

' Clear out the picture box.
'Private Sub Command6_Click()
'    Set Picture1.Picture = Nothing
'End Sub

Public Function CreateBitmapPicture(ByVal hBmp As Long, _
                                    ByVal hPal As Long) As Picture
    Dim r As Long

    Dim Pic As PicBmp
    ' IPicture requires a reference to "Standard OLE Types."
    Dim Ipic As IPicture
    Dim IID_IDispatch As GUID

    ' Fill in with IDispatch Interface ID.
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    ' Fill Pic with necessary parts.
    With Pic
        .Size = Len(Pic)          ' Length of structure.
        .type = vbPicTypeBitmap   ' Type of Picture (bitmap).
        .hBmp = hBmp              ' Handle to bitmap.
        .hPal = hPal              ' Handle to palette (may be null).
    End With

    ' Create Picture object.
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, Ipic)

    ' Return the new Picture object.
    Set CreateBitmapPicture = Ipic
End Function

Public Function CaptureWindow(ByVal hWndSrc As Long, _
                              ByVal Client As Boolean, _
                              ByVal LeftSrc As Long, _
                              ByVal TopSrc As Long, _
                              ByVal WidthSrc As Long, _
                              ByVal HeightSrc As Long) As Picture

    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE

    ' Depending on the value of Client get the proper device context.
    If Client Then
        hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
    Else
        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
        ' window.
    End If

    ' Create a memory device context for the copy process.
    hDCMemory = CreateCompatibleDC(hDCSrc)
    ' Create a bitmap and place it in the memory DC.
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    ' Get screen properties.
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    ' capabilities.
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
    ' support.
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
    ' palette.

    ' If the screen has a palette make a copy and realize it.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        ' Create a copy of the system palette.
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        ' Select the new palette into the memory DC and realize it.
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If

    ' Copy the on-screen image into the memory DC.
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    ' Remove the new copy of the  on-screen image.
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    ' If the screen has a palette get back the palette that was
    ' selected in previously.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    ' Release the device context resources back to the system.
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)

    ' Call CreateBitmapPicture to create a picture object from the
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureScreen
'    - Captures the entire screen.
'
' Returns
'    - Returns a Picture object containing a bitmap of the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureScreen() As Picture
    Dim hWndScreen As Long

    ' Get a handle to the desktop window.
    hWndScreen = GetDesktopWindow()

    ' Call CaptureWindow to capture the entire desktop give the handle
    ' and return the resulting Picture object.

    Set CaptureScreen = CaptureWindow(hWndScreen, False, Image1.left / Screen.TwipsPerPixelX, 0, Image1.Width \ Screen.TwipsPerPixelX, Image1.Height \ Screen.TwipsPerPixelY)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureForm
'    - Captures an entire form including title bar and border.
'
' frmSrc
'    - The Form object to capture.
'
' Returns
'    - Returns a Picture object containing a bitmap of the entire
'      form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureForm(frmSrc As Form) As Picture
    ' Call CaptureWindow to capture the entire form given its window
    ' handle and then return the resulting Picture object.
    Set CaptureForm = CaptureWindow(frmSrc.hWnd, False, 200, 300, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureClient
'    - Captures the client area of a form.
'
' frmSrc
'    - The Form object to capture.
'
' Returns
'    - Returns a Picture object containing a bitmap of the form's
'      client area.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureClient(frmSrc As Form) As Picture
    ' Call CaptureWindow to capture the client area of the form given
    ' its window handle and return the resulting Picture object.
    Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth - 1000, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
    'Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 220, 0, 12000, 12000)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureActiveWindow
'    - Captures the currently active window on the screen.
'
' Returns
'    - Returns a Picture object containing a bitmap of the active
'      window.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureActiveWindow() As Picture

    Dim hWndActive As Long
    Dim r As Long
    
    Dim RectActive As RECT
    
    ' Get a handle to the active/foreground window.
    hWndActive = GetForegroundWindow()
    
    ' Get the dimensions of the window.
    r = GetWindowRect(hWndActive, RectActive)
    
    ' Call CaptureWindow to capture the active window given its
    ' handle and return the Resulting Picture object.
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 2880, 0, RectActive.right - RectActive.left, RectActive.bottom - RectActive.top)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintPictureToFitPage
'    - Prints a Picture object as big as possible.
'
' Prn
'    - Destination Printer object.
'
' Pic
'    - Source Picture object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Sub PrintPictureToFitPage(Prn As Printer, _
                                 Pic As Picture)
    On Error Resume Next
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double
    
    ' Determine if picture should be printed in landscape or portrait
    ' and set the orientation.
    If Pic.Height >= Pic.Width Then
        Prn.Orientation = vbPRORPortrait   ' Taller than wide.
    Else
        Prn.Orientation = vbPRORLandscape  ' Wider than tall.
    End If
    
    ' Calculate device independent Width-to-Height ratio for picture.
    PicRatio = Pic.Width / Pic.Height
    
    ' Calculate the dimentions of the printable area in HiMetric.
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    ' Calculate device independent Width to Height ratio for printer.
    PrnRatio = PrnWidth / PrnHeight
    
    ' Scale the output to the printable area.
    If PicRatio >= PrnRatio Then
        ' Scale picture to fit full width of printable area.
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
        ' Scale picture to fit full height of printable area.
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
    
    ' Print the picture using the PaintPicture method.
    If Prn.Orientation = vbPRORLandscape Then
        Prn.PaintPicture Pic, 0, 50, PrnPicWidth, PrnPicHeight + 740
    
    Else
        Prn.PaintPicture Pic, 0, 50, PrnPicWidth, PrnPicHeight + 2150
    End If

End Sub
'-------------------------------------------------------------------

Private Sub Cmd_save_Click()

    If Check1.value = 0 Then

        txtlostfocusfun
    Else
        Shape1.Visible = False
    End If

    DoEvents
    Dim LASTIMAGENO As Integer
    imaged.Adodc2.CommandType = adCmdText
    imaged.Adodc2.RecordSource = "SELECT MAX(image_no)  AS LASTIMAGENO FROM subjects_images WHERE subject_no= " & imaged.SUBJECT_NO.Caption
    imaged.Adodc2.Refresh

    If imaged.Adodc2.Recordset.RecordCount = 0 Or IsNull(imaged.Adodc2.Recordset.Fields!LASTIMAGENO) Then
        LASTIMAGENO = 1
    Else
        LASTIMAGENO = (imaged.Adodc2.Recordset.Fields!LASTIMAGENO) + 1
    End If

    Set Picture1.Picture = CaptureScreen()
   
    SavePicture Picture1.Picture, system_path & "\images\" & imaged.SUBJECT_NO & "-" & LASTIMAGENO & "#" & Mid(Date, 1, 2) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 7, 4) & "-M.JPG"
    
    DoEvents
    DoEvents

    imaged.Adodc1.Recordset.AddNew
    imaged.Text1.text = imaged.SUBJECT_NO.Caption
    imaged.Text2.text = imaged.SUBJECT_NO & "-" & LASTIMAGENO & "#" & Mid(Date, 1, 2) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 7, 4) & "-M"
    imaged.Text3.text = Now
    imaged.Text4.text = LASTIMAGENO
    imaged.Text5.text = imaged.Departement.Caption
    imaged.Adodc1.Recordset.update

    '       log_files_form.Adodc1.Recordset.AddNew: log_files_form.Adodc1.Recordset.Fields!log_date = Date
    '       log_files_form.Adodc1.Recordset.Fields!log_time = Time
    '       log_files_form.Adodc1.Recordset.Fields!User_Name = login.Adodc1.Recordset.Fields!name

    '      log_files_form.Adodc1.Recordset.Fields!process_name = "    ‘«‘…   " & Me.Caption
    '       log_files_form.Adodc1.Recordset.Fields!process_text = "  „   ⁄œÌ· ’Ê—… —ﬁ„ " & LASTIMAGENO & "  ··„” ‰œ —ﬁ„  " & imaged.SUBJECT_NO
    '
    '        log_files_form.Adodc1.Recordset.update: DoEvents

    'CommonDialog1.DefaultExt = ".jpg"
    'CommonDialog1.Filter = "Bitmap Image (*.bmp)|*.bmp"
    'CommonDialog1.ShowSave
    'If CommonDialog1.FileName <> "" Then
    '        Set Picture1.Picture = CaptureClient(Me)

    'End If

End Sub

Private Sub Txtsend_KeyPress(Index As Integer, _
                             KeyAscii As Integer)
    MsgBox KeyAscii
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    edit_index = Index
    Shape1.top = Text3(Index).top
    Shape1.left = Text3(Index).left
    Shape1.Width = Text3(Index).Width
    Shape1.Height = Text3(Index).Height
End Sub

Private Sub Text3_KeyDown(Index As Integer, _
                          KeyCode As Integer, _
                          Shift As Integer)

    'MsgBox Index
    If Shift = 1 And KeyCode = 37 Then ' LEFT
        Text3(Index).left = Text3(Index).left - 30
        GoTo ll
    End If

    If Shift = 1 And KeyCode = 38 Then 'UP
        Text3(Index).top = Text3(Index).top - 30
        GoTo ll
    End If

    If Shift = 1 And KeyCode = 39 Then ' RIGHT
        Text3(Index).left = Text3(Index).left + 30
        GoTo ll
    End If

    If Shift = 1 And KeyCode = 40 Then 'DOWN
        Text3(Index).top = Text3(Index).top + 30
        GoTo ll
    End If

    If CURRENT_CHAR = "W" Then
        If KeyCode = 39 Then 'ADD
            Text3(Index).Width = Text3(Index).Width + 30
            ' Text3(Index).text = ""
            GoTo ll
        Else

            If KeyCode = 37 Then '109
                Text3(Index).Width = Text3(Index).Width - 30
                ' Text3(Index).text = ""
                GoTo ll
            End If
        End If
 
        'Text3(Index).text = ""
    End If

    If CURRENT_CHAR = "W" Then
        If KeyCode = 40 Then 'ADD
            Text3(Index).Height = Text3(Index).Height + 30
            GoTo ll
        Else

            If KeyCode = 38 Then '109
                Text3(Index).Height = Text3(Index).Height - 30
                GoTo ll
            End If
        End If
 
        'Text3(Index).text = ""
    End If

    If KeyCode = 46 Then
        Dim txtindex As Integer
        Dim astrSplitItems() As String
        Dim i As Integer

        If List1.ListIndex >= 0 Then
        
            For i = 0 To List1.ListCount - 1
        
                astrSplitItems = Split(List1.List(i), "-")
        
                txtindex = val(astrSplitItems(0))

                If txtindex = Index Then
                    List1.List(i) = "##Deleted##"
                End If
        
            Next i

            Unload Text3(Index)
            GoTo ll
            'Text3(Index).Visible = False
            Shape1.Visible = False
         
        End If

    End If

ll:
    change_object_data2 (Index)
End Sub

Private Sub Timer2_Timer()

    If Shape2.BorderColor = &HFFFF& Then
        Shape2.BorderColor = vbBlack
    Else
        Shape2.BorderColor = &HFFFF&
    End If

End Sub
