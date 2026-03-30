VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form CONTRACT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "`"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   14535
   Begin VB.TextBox Text25 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "inventory_name"
      DataSource      =   "Adodc1"
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
      Left            =   2280
      TabIndex        =   109
      Top             =   5400
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "inventory_name"
      DataSource      =   "Adodc1"
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
      Left            =   2280
      TabIndex        =   107
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "inventory_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   105
      Top             =   4080
      Width           =   3135
   End
   Begin VB.ComboBox Combo5 
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
      Height          =   480
      ItemData        =   "CONTRACT.frx":0000
      Left            =   2280
      List            =   "CONTRACT.frx":000A
      RightToLeft     =   -1  'True
      TabIndex        =   104
      Top             =   480
      Width           =   3135
   End
   Begin VB.ComboBox Combo4 
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
      Height          =   480
      ItemData        =   "CONTRACT.frx":001A
      Left            =   2280
      List            =   "CONTRACT.frx":0027
      TabIndex        =   103
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "inventory_name"
      DataSource      =   "Adodc1"
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
      Left            =   2280
      TabIndex        =   101
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "inventory_name"
      DataSource      =   "Adodc1"
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
      Left            =   2280
      TabIndex        =   98
      Top             =   1080
      Width           =   3135
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   9480
      TabIndex        =   97
      Top             =   6120
      Visible         =   0   'False
      Width           =   2775
      _Version        =   524288
      _ExtentX        =   4895
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2010
      Month           =   11
      Day             =   23
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   4200
      TabIndex        =   5
      Top             =   8880
      Width           =   5535
      Begin ALLButtonS.ALLButton Command1 
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "ÃœÌœ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "CONTRACT.frx":003F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "CONTRACT.frx":005B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Õ–›"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "CONTRACT.frx":0077
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
         Height          =   330
         Left            =   1560
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
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
      Begin ALLButtonS.ALLButton Command1 
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "ÿ»«⁄…"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "CONTRACT.frx":0093
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "„—›ﬁ« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "CONTRACT.frx":00AF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "»«·‰”»… ··»Ì⁄"
      Height          =   3135
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   79
      Top             =   5760
      Width           =   6975
      Begin VB.TextBox Text24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   88
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
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
         Height          =   480
         ItemData        =   "CONTRACT.frx":00CB
         Left            =   0
         List            =   "CONTRACT.frx":00D5
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   86
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   85
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   600
         TabIndex        =   84
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   600
         TabIndex        =   83
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   5
         Left            =   120
         Picture         =   "CONTRACT.frx":00ED
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1680
         Width           =   492
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   81
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   4
         Left            =   3000
         Picture         =   "CONTRACT.frx":094F
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   120
         Width           =   492
      End
      Begin ALLButtonS.ALLButton ALLButton1 
         Height          =   255
         Left            =   600
         TabIndex        =   89
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Õ”«» «·«ﬁ”«ÿ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "CONTRACT.frx":11B1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label35 
         Caption         =   "ﬁÌ„… «·ÊÕœ…"
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
         Height          =   495
         Left            =   5280
         TabIndex        =   96
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label34 
         Caption         =   "‰Ê⁄ «·”œ«œ"
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
         Height          =   375
         Left            =   2160
         TabIndex        =   95
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label33 
         Caption         =   "⁄œœ «·«ﬁ”«ÿ"
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
         Height          =   375
         Left            =   5280
         TabIndex        =   94
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "«·„œ›Ê⁄"
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
         Height          =   375
         Left            =   5520
         TabIndex        =   93
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "«·„ »ﬁÌ"
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
         Height          =   375
         Left            =   2520
         TabIndex        =   92
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label30 
         Caption         =   " «—ÌŒ «Ê· ﬁ”ÿ"
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
         Height          =   375
         Left            =   2160
         TabIndex        =   91
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label29 
         Caption         =   " «—ÌŒ  Õ—Ì— «·⁄ﬁœ"
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
         Height          =   375
         Left            =   5160
         TabIndex        =   90
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "»«·‰”»… ··«ÌÃ«—"
      Height          =   3135
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   5760
      Width           =   6975
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   3
         Left            =   3000
         Picture         =   "CONTRACT.frx":11CD
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   120
         Width           =   492
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   77
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   2
         Left            =   120
         Picture         =   "CONTRACT.frx":1A2F
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1680
         Width           =   492
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   600
         TabIndex        =   74
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   600
         TabIndex        =   72
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   71
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   70
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
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
         Height          =   480
         ItemData        =   "CONTRACT.frx":2291
         Left            =   0
         List            =   "CONTRACT.frx":229B
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   68
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   0
         Left            =   0
         Picture         =   "CONTRACT.frx":22B3
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   720
         Width           =   492
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   600
         TabIndex        =   66
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   1
         Left            =   3120
         Picture         =   "CONTRACT.frx":2B15
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   720
         Width           =   492
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   3600
         TabIndex        =   64
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   480
         ItemData        =   "CONTRACT.frx":3377
         Left            =   0
         List            =   "CONTRACT.frx":3384
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "inventory_name"
         DataSource      =   "Adodc1"
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
         Left            =   1080
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
      Begin ALLButtonS.ALLButton Command100 
         Height          =   255
         Left            =   600
         TabIndex        =   61
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Õ”«» «·«ﬁ”«ÿ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "CONTRACT.frx":3397
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label28 
         Caption         =   " «—ÌŒ  Õ—Ì— «·⁄ﬁœ"
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
         Height          =   375
         Left            =   5160
         TabIndex        =   76
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   " «—ÌŒ «Ê· ﬁ”ÿ"
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
         Height          =   375
         Left            =   2160
         TabIndex        =   73
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "«·„ »ﬁÌ"
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
         Height          =   375
         Left            =   2520
         TabIndex        =   60
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "«·„œ›Ê⁄"
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
         Height          =   375
         Left            =   5520
         TabIndex        =   59
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "⁄œœ «·«ﬁ”«ÿ"
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
         Height          =   375
         Left            =   5280
         TabIndex        =   58
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "‰Ê⁄ «·”œ«œ"
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
         Height          =   375
         Left            =   2160
         TabIndex        =   57
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "«·ﬁÌ„… «·«ÌÃ«—Ì…"
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
         Height          =   735
         Left            =   5160
         TabIndex        =   56
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Ì‰ ÂÌ ›Ì"
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
         Height          =   495
         Left            =   2160
         TabIndex        =   55
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Ì»œ√ ›Ì"
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
         Height          =   735
         Left            =   5280
         TabIndex        =   54
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "„œ… «·⁄ﬁœ"
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
         Height          =   735
         Left            =   2160
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text9 
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
      Height          =   480
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "1"
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   8640
      Picture         =   "CONTRACT.frx":33B3
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "»ÕÀ ⁄‰ ”‰œF3"
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   3495
      Left            =   13920
      TabIndex        =   28
      Top             =   12240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   12840
      Picture         =   "CONTRACT.frx":4D45
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "»ÕÀ ⁄‰ ”‰œF3"
      Top             =   -120
      Width           =   975
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "code"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   10920
      TabIndex        =   21
      Top             =   9600
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   5160
      TabIndex        =   16
      Top             =   10800
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   8280
      TabIndex        =   15
      Top             =   11160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   240
      TabIndex        =   12
      Top             =   10920
      Visible         =   0   'False
      Width           =   7095
      Begin MSAdodcLib.Adodc user_priviliges_adodc 
         Height          =   495
         Left            =   120
         Top             =   240
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
      Begin VB.Label adodc4error 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         DataField       =   "employee_id"
         DataSource      =   "user_priviliges_adodc"
         Height          =   495
         Left            =   2160
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M35"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtname 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "inventory_name"
      DataSource      =   "Adodc1"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "code"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "code"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "code"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "inventory_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2640
      Width           =   3135
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   360
      TabIndex        =   29
      ToolTipText     =   "Language  «··€…"
      Top             =   0
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
      MICON           =   "CONTRACT.frx":66D7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodccode1 
      Height          =   465
      Left            =   2760
      Top             =   6840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
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
   Begin MSAdodcLib.Adodc Adodccode2 
      Height          =   345
      Left            =   3360
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
   Begin MSAdodcLib.Adodc Adodc_prifix 
      Height          =   465
      Left            =   4080
      Top             =   6840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "CONTRACT.frx":66F3
      DataField       =   "Manger"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   9360
      TabIndex        =   30
      Top             =   2640
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "Name"
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "CONTRACT.frx":6708
      DataField       =   "Manger"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   9360
      TabIndex        =   31
      Top             =   3120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "Name"
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "CONTRACT.frx":671D
      DataField       =   "Manger"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   9360
      TabIndex        =   32
      Top             =   3600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "Name"
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Bindings        =   "CONTRACT.frx":6732
      DataField       =   "Manger"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   9360
      TabIndex        =   33
      Top             =   2160
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "Name"
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Bindings        =   "CONTRACT.frx":6747
      DataField       =   "Manger"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   9360
      TabIndex        =   100
      Top             =   1560
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "Name"
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label38 
      Caption         =   "’Ì«‰… „ﬁœ„…"
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
      Left            =   5640
      TabIndex        =   110
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label37 
      Caption         =   "„Ì«… „ﬁœ„…"
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
      Left            =   5640
      TabIndex        =   108
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "«·„ÊÃÊœ« "
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
      Left            =   5640
      TabIndex        =   106
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "«·„”«Õ…"
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
      Left            =   5640
      TabIndex        =   102
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label20 
      Caption         =   "ﬂÊœ «·ÊÕœ…"
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
      Height          =   495
      Left            =   5640
      TabIndex        =   99
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "‰Ê⁄ «·⁄ﬁœ"
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
      Height          =   375
      Left            =   5760
      TabIndex        =   50
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "—ﬁ„ «·⁄ﬁœ"
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
      Height          =   375
      Left            =   12840
      TabIndex        =   49
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "⁄ﬁœ ≈ÌÃ«—/ »Ì⁄ ÊÕœ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5280
      TabIndex        =   47
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label10 
      Caption         =   "ÿ—› «Ê· «·„«·ﬂ"
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
      Height          =   375
      Left            =   12840
      TabIndex        =   46
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label v 
      Caption         =   "—ﬁ„ «·⁄ﬁ«—"
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
      Height          =   375
      Left            =   12840
      TabIndex        =   45
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "⁄œœ «·€—› "
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
      Left            =   5640
      TabIndex        =   44
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "«·„œÌ‰…"
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
      Left            =   12840
      TabIndex        =   43
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "«·ÕÌ"
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
      Left            =   12840
      TabIndex        =   42
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "«·‘«—⁄"
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
      Left            =   12840
      TabIndex        =   41
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "«·⁄‰Ê«‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   12360
      TabIndex        =   40
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label16 
      Caption         =   "—ﬁ„ «·»‰«Ì…"
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
      Left            =   12840
      TabIndex        =   39
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "„œŒ·"
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
      Left            =   12840
      TabIndex        =   38
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "«·œÊ—"
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
      Left            =   12840
      TabIndex        =   37
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label19 
      Caption         =   "‰‘«ÿ «·ÊÕœ…"
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
      Height          =   375
      Left            =   5640
      TabIndex        =   36
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "«·„‰«›⁄"
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
      Left            =   5640
      TabIndex        =   35
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "ÿ—› À«‰Ì"
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
      Height          =   375
      Left            =   12840
      TabIndex        =   34
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "CONTRACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bindex As Integer
Private Sub Label39_Click()

End Sub

Private Sub Calendar1_Click()
On Error Resume Next
If bindex = 3 Then
Text18.Text = Calendar1.Value
End If

If bindex = 1 Then
Text11.Text = Calendar1.Value
 
Dim flag As String

If Combo1.ListIndex = 0 Then
flag = "d"
Else
If Combo1.ListIndex = 1 Then
flag = "m"
Else
If Combo1.ListIndex = 2 Then
flag = "yyyy"
End If
End If
End If


Text12.Text = DateAdd(flag, Text10.Text, Text11.Text)
Text12.Text = DateAdd("d", -1, Text12.Text)

End If

If bindex = 0 Then
Text12.Text = Calendar1.Value
End If

If bindex = 2 Then
Text17.Text = Calendar1.Value
End If

If bindex = 4 Then
Text19.Text = Calendar1.Value
End If

If bindex = 5 Then
Text20.Text = Calendar1.Value
End If


 
Calendar1.Visible = False


End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
If Index = 4 Then
        
        
        If my_language = "E" Then
        If Text9.Text = "" Then MsgBox "Select Voucher First": Exit Sub
        
        Else
        If Text9.Text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— ”‰œ  «Ê·«": Exit Sub
        End If
        
        imaged.Show
        imaged.txtopeation_type = "⁄ﬁœ"
        imaged.SUBJECT_NO = Text9.Text
         If my_language = "E" Then
         imaged.Label6.Caption = "contract #"
         imaged.Caption = "contract Attachments"
         Else
        imaged.Label6.Caption = "—ﬁ„ «·⁄ﬁœ"
         imaged.Caption = "„—›ﬁ«  «·⁄ﬁœ"
        End If
        imaged.Adodc1.CommandType = adCmdText
        imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '⁄ﬁœ' and subject_no='" & Text9.Text & "'"
        imaged.Adodc1.Refresh
        If imaged.Adodc1.Recordset.RecordCount > 0 Then
        
        imaged.DBPix201.Visible = True
        Else
        imaged.DBPix201.Visible = False
        End If
End If
End Sub

Private Sub Command100_Click()
AKSAT.Show
AKSAT.Text14.Text = Text14.Text
AKSAT.Text16.Text = Text16.Text
AKSAT.Text17.Text = Text17.Text

End Sub

Private Sub Command40_Click(Index As Integer)
bindex = Index
Calendar1.Visible = True
Calendar1.Value = Date
End Sub

Private Sub Form_Load()
      On Error Resume Next
          login.SkinFramework.ApplyWindow Me.hWnd
Me.Left = (MDIForm1.Width - Me.Width) / 2
   Me.Top = (MDIForm1.Height - Me.Height) / 2 - 500
End Sub

Private Sub Frame7_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub Text15_Change()
If IsNumeric(Text13.Text) And IsNumeric(Text15.Text) Then
Text16.Text = Text13.Text - Text15.Text
Else
MsgBox "Â‰«ﬂ Œÿ√ ›Ì «·ﬁÌ„… «·«ÌÃ«—Ì… «Ê «·„œ›Ê⁄"
End If
End Sub
