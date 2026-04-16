VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmsandat_ked 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”‰œ ﬁÌœ"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   1080
   ClientWidth     =   13605
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   13605
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtDEVID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "Double_Entry_Vouchers_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   0
      TabIndex        =   130
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox NoteID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "NoteID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3600
      TabIndex        =   129
      Top             =   0
      Width           =   2055
   End
   Begin DBPIXLib.DBPix20 DBPix202 
      Height          =   615
      Left            =   840
      TabIndex        =   123
      Top             =   7440
      Width           =   3015
      _Version        =   131072
      _ExtentX        =   5318
      _ExtentY        =   1085
      _StockProps     =   1
      BackColor       =   12632256
      _Image          =   "frmsandat_ked.frx":0000
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
   Begin VB.Frame Frame17 
      Height          =   855
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   111
      Top             =   600
      Width           =   3735
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "„·€Ì"
         Height          =   195
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÌœ œÊ—Ì"
         Height          =   195
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁ«·»"
         Height          =   195
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   114
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   " „ «⁄ „«œÂ"
         Height          =   195
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   113
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œÌ„ «· √ÀÌ—"
         Height          =   195
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   112
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSACAL.Calendar Calendar1 
      DataSource      =   "Adodc2"
      Height          =   2295
      Left            =   5400
      TabIndex        =   53
      Top             =   1080
      Visible         =   0   'False
      Width           =   4095
      _Version        =   524288
      _ExtentX        =   7223
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2010
      Month           =   11
      Day             =   25
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
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      TabIndex        =   29
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label46 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Move"
         Height          =   255
         Left            =   600
         TabIndex        =   90
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         TabIndex        =   32
         Top             =   -120
         Width           =   300
      End
      Begin VB.Label Label18 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   31
         Top             =   0
         Width           =   310
      End
      Begin VB.Label Label17 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   -120
         Picture         =   "frmsandat_ked.frx":0018
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2385
      End
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "Remark"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1560
      Width           =   11415
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   4560
      TabIndex        =   80
      Top             =   9360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3360
         TabIndex        =   84
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   83
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   82
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   4560
      TabIndex        =   75
      Top             =   9360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4800
         TabIndex        =   78
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   77
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   76
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame12 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   -600
      TabIndex        =   73
      Top             =   2040
      Width           =   3495
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   720
      TabIndex        =   68
      Top             =   8040
      Width           =   12735
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   1
         Left            =   10680
         TabIndex        =   69
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
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
         MICON           =   "frmsandat_ked.frx":0F7E
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
         Height          =   495
         Index           =   2
         Left            =   9720
         TabIndex        =   70
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Õ–›"
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
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":0F9A
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
         Height          =   495
         Index           =   0
         Left            =   11640
         TabIndex        =   71
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÃœÌœ"
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
         MICON           =   "frmsandat_ked.frx":0FB6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton1 
         Height          =   495
         Left            =   1560
         TabIndex        =   99
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "«·„—›ﬁ« "
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
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":0FD2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Print_cmd 
         Height          =   255
         Left            =   2160
         TabIndex        =   100
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "ÿ»«⁄…"
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
         BCOL            =   65535
         BCOLO           =   65535
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":0FEE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   255
         Left            =   7800
         TabIndex        =   101
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "»ÕÀ ﬁÌœ"
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
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":100A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton20 
         Height          =   255
         Left            =   8880
         TabIndex        =   117
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "«⁄ „«œ"
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
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":1026
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton6 
         Height          =   255
         Left            =   6840
         TabIndex        =   118
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "ﬁÌœ œÊ—Ì"
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
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":1042
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton7 
         Height          =   255
         Left            =   5640
         TabIndex        =   119
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   " ÕÊÌ· «·Ï ﬁ«·»"
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
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":105E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton8 
         Height          =   255
         Left            =   4680
         TabIndex        =   120
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "«·€«¡ «· √ÀÌ—"
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
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":107A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton9 
         Height          =   255
         Left            =   3120
         TabIndex        =   121
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "ÿ»«⁄Â ⁄·Ï «·‘«‘…"
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
         BCOL            =   65535
         BCOLO           =   65535
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":1096
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton10 
         Height          =   255
         Left            =   1080
         TabIndex        =   122
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "«” œ⁄«¡ ﬁ«·»"
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
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsandat_ked.frx":10B2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   72
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   13560
      Picture         =   "frmsandat_ked.frx":10CE
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "»ÕÀ ⁄‰ ”‰œF3"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   840
      TabIndex        =   55
      Top             =   9600
      Visible         =   0   'False
      Width           =   12495
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "„"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   11760
         TabIndex        =   98
         Top             =   120
         Width           =   135
      End
      Begin VB.Line Line1 
         X1              =   12190
         X2              =   12190
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line9 
         X1              =   11685
         X2              =   11685
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line10 
         X1              =   6390
         X2              =   6390
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line11 
         X1              =   5090
         X2              =   5090
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "œ«∆‰"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5520
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line12 
         X1              =   7700
         X2              =   7700
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line13 
         X1              =   9690
         X2              =   9690
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·Õ”«»"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10080
         TabIndex        =   59
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·Õ”«»"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   8160
         TabIndex        =   58
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "„œÌ‰"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   6840
         TabIndex        =   57
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘—Õ"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2160
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   720
      TabIndex        =   61
      Top             =   9480
      Visible         =   0   'False
      Width           =   12735
      Begin VB.Line Line5 
         X1              =   810
         X2              =   810
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "I"
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
         Height          =   375
         Left            =   600
         TabIndex        =   96
         Top             =   120
         Width           =   255
      End
      Begin VB.Line Line14 
         X1              =   7410
         X2              =   7410
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
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
         Left            =   6360
         TabIndex        =   66
         Top             =   120
         Width           =   975
      End
      Begin VB.Line Line8 
         X1              =   300
         X2              =   300
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line6 
         X1              =   2800
         X2              =   2800
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line4 
         X1              =   6110
         X2              =   6110
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C#"
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
         Height          =   375
         Left            =   1440
         TabIndex        =   65
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Name"
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
         Left            =   3120
         TabIndex        =   64
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Depit"
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
         Left            =   5040
         TabIndex        =   63
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   495
         Left            =   9480
         TabIndex        =   62
         Top             =   120
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   4800
         X2              =   4800
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   8880
      TabIndex        =   48
      Top             =   480
      Width           =   3255
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataField       =   "NoteSerial"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   975
         Begin VB.Label Label43 
            Caption         =   "Entry#"
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
            Height          =   495
            Left            =   0
            TabIndex        =   92
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label40 
            Caption         =   "Index"
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
            Height          =   375
            Left            =   0
            TabIndex        =   52
            Top             =   -960
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label41 
            Caption         =   "Type"
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
            Left            =   0
            TabIndex        =   51
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "sandat_pc_no"
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
         Height          =   600
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   -480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmsandat_ked.frx":2A60
         DataField       =   "sanad_type"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "ked_name"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame10 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   4080
      TabIndex        =   44
      Top             =   480
      Width           =   3735
      Begin VB.CommandButton Command40 
         Height          =   375
         Index           =   0
         Left            =   1080
         Picture         =   "frmsandat_ked.frx":2A75
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   600
         Width           =   492
      End
      Begin VB.TextBox txtdate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataField       =   "NoteDate"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   600
         Width           =   2055
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         Begin VB.Label Label1 
            Caption         =   "Date"
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
            Index           =   1
            Left            =   0
            TabIndex        =   95
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label44 
            Caption         =   "Source"
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
            Height          =   495
            Left            =   0
            TabIndex        =   47
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataField       =   "sanad_source"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   45
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame14 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   12120
      TabIndex        =   41
      Top             =   480
      Width           =   1455
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "„”·”· «·”‰œ"
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
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   -600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "‰Ê⁄ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame15 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   7560
      TabIndex        =   39
      Top             =   480
      Width           =   1935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   " «—ÌŒ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   93
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "„’œ— «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   -120
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame16 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2880
      TabIndex        =   38
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   2040
      TabIndex        =   35
      Top             =   9840
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
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M19"
         Height          =   255
         Left            =   3360
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label adodc4error 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         DataField       =   "employee_id"
         DataSource      =   "user_priviliges_adodc"
         Height          =   495
         Left            =   2160
         TabIndex        =   36
         Top             =   120
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   1680
      Top             =   6720
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   585
      Left            =   -480
      Top             =   6960
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   585
      Left            =   2880
      Top             =   6120
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   585
      Left            =   2880
      Top             =   6720
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   585
      Left            =   1800
      Top             =   5880
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   585
      Left            =   2880
      Top             =   8040
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
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF0000&
      Caption         =   "Õ›Ÿ ‰Â«∆Ì"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Õ–›"
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
      Left            =   -600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -5280
      Top             =   6840
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "description"
      DataSource      =   "Adodc2"
      Height          =   840
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3360
      Width           =   11415
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "credit_value"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "depet_value"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "account_name"
      Height          =   285
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF0000&
      Caption         =   " ÕœÌÀ"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   -480
      TabIndex        =   54
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
      MICON           =   "frmsandat_ked.frx":32D7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   16
      Top             =   7080
      Width           =   12735
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8400
         TabIndex        =   19
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   18
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DcboUsers 
         Height          =   315
         Left            =   10440
         TabIndex        =   126
         Top             =   480
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "Õ—— »Ê«”ÿ…"
         Height          =   495
         Left            =   10680
         TabIndex        =   127
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "«Ã„«·Ì «·„œÌ‰"
         Height          =   495
         Left            =   8520
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "«Ã„«·Ì «·œ«∆‰"
         Height          =   495
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "«·›—ﬁ"
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   6120
      TabIndex        =   87
      Top             =   -240
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "”ÿ— ÃœÌœ"
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
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   3
         Height          =   855
         Left            =   -120
         Shape           =   5  'Rounded Square
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc numbering 
      Height          =   585
      Left            =   9120
      Top             =   0
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
      Left            =   7920
      Top             =   0
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
   Begin ALLButtonS.ALLButton Command3 
      Height          =   375
      Left            =   2280
      TabIndex        =   103
      Top             =   2880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      BCOL            =   65280
      BCOLO           =   65280
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmsandat_ked.frx":32F3
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
      Height          =   375
      Left            =   1080
      TabIndex        =   104
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«· Ê“Ì⁄"
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
      BCOL            =   65280
      BCOLO           =   65280
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmsandat_ked.frx":330F
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
      Left            =   0
      TabIndex        =   105
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
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
      BCOL            =   65280
      BCOLO           =   65280
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmsandat_ked.frx":332B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton5 
      Height          =   495
      Left            =   0
      TabIndex        =   106
      Top             =   5280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   " ÕœÌÀ ”ÿ—"
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
      BCOL            =   65280
      BCOLO           =   65280
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmsandat_ked.frx":3347
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   0
      Left            =   2115
      TabIndex        =   107
      Top             =   120
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmsandat_ked.frx":3363
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   3
      Left            =   1425
      TabIndex        =   108
      Top             =   120
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmsandat_ked.frx":36FD
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   1
      Left            =   2745
      TabIndex        =   109
      Top             =   120
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmsandat_ked.frx":3A97
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   2
      Left            =   720
      TabIndex        =   110
      Top             =   120
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmsandat_ked.frx":3E31
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin MSDataListLib.DataCombo DcAccount1 
      Height          =   315
      Left            =   10920
      TabIndex        =   124
      Top             =   2880
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcAccount2 
      Height          =   315
      Left            =   6960
      TabIndex        =   125
      Top             =   2880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmsandat_ked.frx":41CB
      Height          =   2535
      Left            =   840
      TabIndex        =   131
      Top             =   4200
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   4471
      _Version        =   393216
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "DEV_ID_Line_No"
         Caption         =   "—ﬁ„ «·”ÿ—"
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
         DataField       =   "Double_Entry_Vouchers_ID"
         Caption         =   "Double_Entry_Vouchers_ID"
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
         DataField       =   "Value"
         Caption         =   "Value"
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
         DataField       =   "Credit_Or_Debit"
         Caption         =   "Credit_Or_Debit"
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
         DataField       =   "Double_Entry_Vouchers_Description"
         Caption         =   "Double_Entry_Vouchers_Description"
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
         DataField       =   "Account_Serial"
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
      BeginProperty Column06 
         DataField       =   "Account_Code"
         Caption         =   "Account_Code"
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
         DataField       =   "Account_Name"
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
      BeginProperty Column08 
         DataField       =   "RecordDate"
         Caption         =   "RecordDate"
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
         DataField       =   "Notes_ID"
         Caption         =   "Notes_ID"
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
         DataField       =   "UserID"
         Caption         =   "UserID"
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
         DataField       =   "depet_value"
         Caption         =   "„œÌ‰"
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
         DataField       =   "credit_value"
         Caption         =   "œ«∆‰"
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
         DataField       =   "Double_Entry_Vouchers_Description"
         Caption         =   "«·‘—Õ"
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
         DataField       =   "Account_Interval_ID"
         Caption         =   "Account_Interval_ID"
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
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   2624.882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   3300.095
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1484.787
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   585
      Left            =   -120
      Top             =   6120
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
   Begin VB.Label LineNo 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   128
      Top             =   5880
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4560
      Picture         =   "frmsandat_ked.frx":41E0
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6480
      Picture         =   "frmsandat_ked.frx":468F
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "     ”‰œ ﬁÌœ         "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   -600
      TabIndex        =   102
      Top             =   0
      Width           =   14295
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      Caption         =   "Desc"
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
      Left            =   -600
      TabIndex        =   89
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Caption         =   "General Desc"
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
      Height          =   735
      Left            =   -600
      TabIndex        =   88
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Caption         =   "«·‘—Õ «·⁄«„"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11760
      TabIndex        =   86
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "«·‘—Õ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11760
      TabIndex        =   85
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Õ–› ”ÿ—"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   -480
      TabIndex        =   34
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   -600
      X2              =   -600
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "«· Ê“Ì⁄"
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
      Height          =   495
      Left            =   -1800
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Õ›Ÿ ‰Â«∆Ì"
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
      Height          =   495
      Left            =   -480
      TabIndex        =   28
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   " ÕœÌÀ ”ÿ—"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   -480
      TabIndex        =   27
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   " ÕœÌÀ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   -1320
      TabIndex        =   26
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   11880
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "„—«ﬂ“ «· ﬂ·›…"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   720
      TabIndex        =   14
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "œ«∆‰"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3480
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "„œÌ‰"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5280
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   8160
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11400
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   5040
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
End
Attribute VB_Name = "frmsandat_ked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Dim new_record As Boolean
Dim end_save_ok As Boolean
Dim auto_sanad_no As String
Dim numbering_type As Integer
Dim Branch_NO As Integer
Dim departement_name As Integer

Private Sub ALLButton1_Click()
    On Error Resume Next

    If my_language = "E" Then
        If Text1.text = "" Then MsgBox "Select Voucher First": Exit Sub

    Else

        If Text1.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— ”‰œ  «Ê·«": Exit Sub
    End If

    imaged.Show
    imaged.txtopeation_type = "”‰œ ﬁÌœ"
    imaged.SUBJECT_NO = Text1.text

    If my_language = "E" Then
        imaged.Label6.Caption = "Voucher #"
        imaged.Caption = "Voucher Attachments"
    Else
        imaged.Label6.Caption = "—ﬁ„ «·”‰œ"
        imaged.Caption = "„—›ﬁ«  «·”‰œ« "
    End If

    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '”‰œ ﬁÌœ' and subject_no='" & Text1.text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next

    Voucher_search.Show
    Voucher_search.title_lbl = "”‰œ ﬁÌœ"

End Sub

Private Sub ALLButton20_Click()

    If Dir(system_path & "\images\sign" & user_id & ".JPG") <> "" Then
        DBPix202.ImageLoadFile (system_path & "\images\sign" & user_id & ".JPG")
    End If

    Check2.value = 1
End Sub

Private Sub ALLButton3_Click()
    Label26_Click
End Sub

Private Sub ALLButton4_Click()
    On Error Resume Next

    If my_language = "E" Then
        If Text1.text = "" Then MsgBox "enter voucher no first", vbCritical: Exit Sub
    Else

        If Text1.text = "" Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·”‰œ", vbCritical: Exit Sub

    End If

    On Error Resume Next

    If my_language = "E" Then
        x = MsgBox("Confirm delete", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ Õ–› ”ÿ—", vbCritical + vbYesNo)
              
    End If
            
    If x = vbNo Then
        Exit Sub
    End If

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.delete
        Adodc2.Refresh
        detect_error
    End If

    'Label32_Click
End Sub

Private Sub ALLButton5_Click()
    Label21_Click
End Sub

Private Sub ALLButton6_Click()

    If Text1.text = "" Then MsgBox "«Œ — ﬁÌœ «Ê·«", vbCritical: Exit Sub
    ked_dawry.Show
    ked_dawry.id = Text1.text
    ked_dawry.desc = Text11.text

End Sub

Private Sub ALLButton7_Click()
    x = MsgBox(" √ﬂÌœ «· ÕÊÌ· «·Ï ﬁ«·»", vbInformation + vbYesNo)

    If x = vbYes Then
        Check2.value = 1
    End If

End Sub

Private Sub ALLButton8_Click()
    Check1.value = 1
End Sub

Private Sub ALLButton9_Click()
    On Error Resume Next

    Form3.Show
 
    Form3.case_id = 16
End Sub

Private Sub Calendar1_Click()
    On Error Resume Next
    txtdate.text = Calendar1.value
    Calendar1.Visible = False
End Sub

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        '''Call Reload(Me)
 
    Else
        my_language = "A"
 
        '''Call Reload(Me)
    End If

End Sub

Private Sub Command1_Click(Index As Integer)

    'On Error Resume Next
    Select Case Index

        Case 0
            'Text1.Text = ""
            Adodc1.Recordset.AddNew
            'Text1.text = new_id("sandat_ked", "sanad_no", "")
            sand_numbering

            'Adodc1.Recordset.Fields!sanad_no = Adodc1.Recordset.Fields!sandat_pc_no
            If auto_sanad_no <> "" Then
                Dim NoteSerial As Integer
                Text1.text = auto_sanad_no
                Text3.text = "ÌœÊÌ"
                NoteSerial = auto_sanad_no
                Adodc1.Recordset.Fields!NoteSerial = NoteSerial
                Adodc1.Recordset.Fields!NoteID = CStr(new_id("notes", "NoteID", "", True))

                If TxtDEVID.text = "" Then
                    Me.TxtDEVID.text = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
                End If
  
                ' Me.TxtDEV_NO.text = Me.TxtDEVID.text
    
                Adodc1.Recordset.Fields!NoteType = 200
                Adodc1.Recordset.Fields!NoteDate = DateValue(Now)
                Adodc1.Recordset.Fields!numbering_type = numbering_type
                Adodc1.Recordset.Fields!UserID = val(Me.DcboUsers.BoundText)
             
                If numbering_type = 2 Then
                    Adodc1.Recordset.Fields!sanad_year = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                    Adodc1.Recordset.Fields!sanad_month = Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
                End If
            
                If numbering_type = 3 Then
                    Adodc1.Recordset.Fields!sanad_year = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
            
                End If
             
                Adodc1.Recordset.Fields!type = "”‰œ ﬁÌœ"
                Adodc1.Recordset.Fields!Branch_NO = Branch_NO
                Adodc1.Recordset.Fields!user_name = current_user_name
                Adodc1.Recordset.Fields!Departement = departement_name
            
                Adodc1.Recordset.update
                Adodc1.Recordset.MoveLast
 
            Else
        
                If my_language = "E" Then
          
                    MsgBox "can't save define numbering method first in system manger", vbCritical: Exit Sub
                Else
                    MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ·«»œ „‰ «œŒ«· —ﬁ„ ··”‰œ «Ê·« ·«‰ﬂ «Œ —   —ﬁÌ„ ”‰œ«  ÌœÊÌ", vbCritical: Exit Sub
                End If
        
            End If
   
            ' DataGrid1.Visible = False
            '  Command3.Visible = False
    
            '  Frame1.Visible = False
    
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=0"
            Adodc2.Refresh

        Case 1

            If my_language = "E" Then

                'If Text1.Text = "" Then MsgBox "Specify voucher no", vbCritical: Exit Sub
                If DataCombo2.text = "" Then MsgBox " specify type of constraint", vbCritical: Exit Sub
            Else

                'If Text1.Text = "" Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·”‰œ", vbCritical: Exit Sub
                If DataCombo2.text = "" Then MsgBox " Õœœ ‰Ê⁄ «·ﬁÌœ", vbCritical: Exit Sub

            End If

            Adodc1.Recordset.update

            end_save

            ' If new_record = True Then
            '  Adodc1.Recordset.MoveLast
            ' new_record = False
            ' End If
    
            '  Adodc1.Refresh
            If Text7.text = "" Or Not IsNumeric(Text7) Then
                'Command3.Visible = False
            Else
                'Command3.Visible = True
            End If
 
        Case 2

            If my_language = "E" Then
                If Text1.text = "" Then MsgBox "enter voucher first", vbCritical: Exit Sub
            
            Else
             
                If Text1.text = "" Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·”‰œ", vbCritical: Exit Sub
            End If

            If my_language = "E" Then
                x = MsgBox("Confirm delete", vbCritical + vbYesNo)
            Else
                x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
              
            End If
            
            If x = vbNo Then
                Exit Sub
            End If

            Check5.value = 1
            'If Adodc2.Recordset.RecordCount > 0 Then
            'For I = 1 To Adodc2.Recordset.RecordCount
            'Adodc2.Recordset.Delete
            'Adodc2.Recordset.MoveNext

            'Next I
            'Adodc2.Refresh
            'End If

            'If Adodc1.Recordset.RecordCount > 0 Then
            'Adodc1.Recordset.Delete
            'Adodc1.Refresh
            'DataGrid1.Refresh
            'Label21_Click
            'End If

            '           If my_language = "E" Then
            '              x = MsgBox("Confirm delete", vbCritical + vbYesNo)
            '            Else
            '            x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
            '
            '            End If
            
            'If x = vbNo Then
            'Exit Sub
            'End If
            'If my_language = "E" Then
            'If Text1.Text = "" Then MsgBox "Select Voucher First", vbCritical: Exit Sub
            '
            'Else
            'If Text1.Text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— ”‰œ  «Ê·«": Exit Sub
            'End If
            '
            '

            '    If Adodc1.Recordset.RecordCount > 0 Then
            '        Adodc2.CommandType = adCmdText
            'Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            'Adodc2.Refresh
    
            '    Adodc1.Recordset.Delete
            '    Adodc1.Refresh
            '    DataGrid1.Refresh
            '    End If

        Case 4
            On Error Resume Next

            If my_language = "E" Then
                x = InputBox("«œŒ· —ﬁ„ «·”‰œ")
            Else
                x = InputBox("Enter Voucher no")

            End If

            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            Adodc2.Refresh

            If IsNumeric(x) Then
                Adodc1.CommandType = adCmdText
                Adodc1.RecordSource = "select * from  sandat_ked where type='”‰œ ﬁÌœ' and sanad_no=" & x
                Adodc1.Refresh
                Text7_Change
            Else
       
                If my_language = "E" Then
                    MsgBox "Enter Digit only", vbCritical
                Else
                    MsgBox "«œŒ· «—ﬁ«„ ›ﬁÿ", vbCritical
                End If
              
            End If

        Case 5
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            Adodc2.Refresh

            If my_language = "E" Then
                x = InputBox("«œŒ· ‰Ê⁄ «·”‰œ")
            Else
                x = InputBox("Enter Voucher Type")

            End If
    
            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from sandat_ked where type='”‰œ ﬁÌœ' and sanad_type like '%" & x & "%'"
            Adodc1.Refresh
            Text7_Change

        Case 6
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            Adodc2.Refresh

            frm_search_date.Show
            frm_search_date.case_id = 1
            Text7_Change
  
        Case 7
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            Adodc2.Refresh

            If my_language = "E" Then
                x = InputBox("«œŒ· „’œ— «·”‰œ")
            Else
                x = InputBox("Enter Voucher Source")

            End If

            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from sandat_ked where type='”‰œ ﬁÌœ' and sanad_source like '%" & x & "%'"
            Adodc1.Refresh
            Text7_Change

        Case 8
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            Adodc2.Refresh
   
            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from sandat_ked type='”‰œ ﬁÌœ' and where attachment=1"
            Adodc1.Refresh
        
            Text7_Change

        Case 9
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            Adodc2.Refresh

            If my_language = "E" Then
                x = InputBox("«œŒ· Ê’› «·”‰œ")
            Else
                x = InputBox("Enter Voucher description")

            End If

            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from sand_all_details_qry where description like '%" & x & "%'"
            Adodc1.Refresh
            Text7_Change

        Case 10
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            Adodc2.Refresh

            If my_language = "E" Then
                x = InputBox("«œŒ· ﬁÌ„… «·”‰œ")
            Else
                x = InputBox("Enter Voucher Value")

            End If
    
            If IsNumeric(x) Then
                Adodc1.CommandType = adCmdText
                Adodc1.RecordSource = "select * from sand_all_details_qry where depet_value = " & x & " or credit_value=" & x
                Adodc1.Refresh
            Else
       
                If my_language = "E" Then
                    MsgBox "Enter Digit only", vbCritical
                Else
                    MsgBox "«œŒ· «—ﬁ«„ ›ﬁÿ", vbCritical
                End If
              
            End If

            Text7_Change
        
        Case 12
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=0"
            Adodc2.Refresh
   
            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from sandat_ked  where type='”‰œ ﬁÌœ'  "
            Adodc1.Refresh
            Text7_Change
        
    End Select

End Sub
 
Private Sub Command13_Click()
    On Error Resume Next

    Voucher_search.Show
    Voucher_search.title_lbl = "”‰œ ﬁÌœ"

End Sub

Private Sub Command2_Click()
    On Error Resume Next

    Voucher_search.Show
End Sub

Private Sub Command3_Click()
    On Error Resume Next

    '  If my_language = "E" Then
    '    If Text1.text = "" Then MsgBox "enter Voucher first you select manual numbering", vbCritical: Exit Sub
    '    If Text7.text = "" Then Exit Sub
    '    If DataCombo1.text = "" Or DataCombo3.text = "" Or (Text4.text = "" And Text5.text = "") Then MsgBox "must fill all fields", vbCritical: Exit Sub

    '  Else
    '    If Text1.text = "" Then MsgBox "·«»œ „‰ Õ›Ÿ «·”‰œ «Ê·«", vbCritical: Exit Sub
    '    If Text7.text = "" Then Exit Sub
    '    If DataCombo1.text = "" Or DataCombo3.text = "" Or (Text4.text = "" And Text5.text = "") Then MsgBox "·«»œ „‰  ”ÃÌ· »Ì«‰«  ﬁÌœ ’ÕÌÕ…", vbCritical: Exit Sub
    '
    '    End If

    Adodc7.RecordSource = "select max(DEV_ID_Line_No) as last_line_no from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
    Adodc7.Refresh

    If Not IsNull(Adodc7.Recordset.Fields!last_line_no) Then
        LineNo.Caption = Adodc7.Recordset.Fields!last_line_no + 1
    Else
        LineNo.Caption = 1
    End If

    If (Text4.text <> "" And Text5.text <> "") Then
        If my_language = "E" Then
            MsgBox "íMust enter depit/credit value", vbCritical
        Else
            MsgBox "ÌÃ»  ”ÃÌ· ﬁÌ„… Ê«Õœ… œ«∆‰… «Ê „œÌ‰…", vbCritical
        End If

        Exit Sub
    End If
            
    Dim description As String
    description = Text6.text
    Dim IntDEV_Type As Integer
    Dim SngDEV_Value As Single

    If DCAccount1.text <> "" Then
        If val(Text4.text) > 0 Then
            IntDEV_Type = 0
            SngDEV_Value = val(Text4.text)
        Else
            IntDEV_Type = 1
            SngDEV_Value = val(Text5.text)
        End If

        '   Me.TxtDEVID.text = x
        If ModAccounts.AddNewDev(val(Me.TxtDEVID.text), val(LineNo.Caption), DCAccount1.BoundText, SngDEV_Value, IntDEV_Type, CStr(description), NoteID.text, , , SystemOptions.SysCurrentAccountIntervalID, Me.txtdate.text, Me.DcboUsers.BoundText, , Me.Text1.text) = False Then

            DoEvents
                
            '         If ModAccounts.AddNewDev(Val(Me.TxtDEVID.text), .TextMatrix(I, .ColIndex("LineNo")), _
            '       .TextMatrix(I, .ColIndex("AccountCode")), SngDEV_Value, IntDEV_Type, _
            '       CStr(.Cell(flexcpData, I, .ColIndex("Des"))), Val(Me.TxtNoteID.text), , , _
            '       SystemOptions.SysCurrentAccountIntervalID, Me.DTP_Date.value, Me.DcboUsers.BoundText, , Me.TxtSerial.text) = False Then
                
            GoTo ErrTrap
            'LineNo.Caption = LineNo.Caption + 1
        End If
    End If

    'Adodc2.Recordset.AddNew

    'Text6.text = Text11.text
    'Adodc2.Recordset.Fields!sandat_pc_no = Text7.text
    'Adodc2.Recordset.Fields!sanad_no = Text1.text
    'Adodc2.Recordset.Fields!account_no = DataCombo1.text
    'Adodc2.Recordset.Fields!Account_Name = DataCombo3.text
    'Adodc2.Recordset.Fields!depet_value = Val(Text4.text)
    'Adodc2.Recordset.Fields!credit_value = Val(Text5.text)
    'Adodc2.Recordset.Fields!description = Text6.text 'description

    'Adodc2.Recordset.Fields!SANAD_TYPE = "”‰œ ﬁÌœ"
    'Adodc2.Recordset.Fields!sanad_source = "ÌœÊÌ"
    'Adodc2.Recordset.Fields!box_name = DataCombo3.Text
    'Adodc2.Recordset.Fields!bona_3la = Text11.Text
    'Adodc2.Recordset.Fields!Date = DateValue(Now)

    'Adodc2.Recordset.update
    'Adodc2.Refresh
    'Adodc2.Recordset.MoveLast

    If NoteID.text = "" Then Exit Sub
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
    Adodc2.Refresh
    DataGrid1.Refresh
  
    'Adodc2.Recordset.update
    'Adodc2.Refresh
    'Adodc2.Recordset.MoveLast

    detect_error
    DCAccount1.text = ""
    DCAccount2.text = ""
    Text4.text = ""
    Text5.text = ""
    Exit Sub
ErrTrap:
    MsgBox ""

End Sub

Function detect_error()

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Refresh
    End If

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.MoveFirst
    End If

    For i = 1 To Adodc2.Recordset.RecordCount

        If Adodc2.Recordset.Fields!Credit_Or_Debit = 0 Then
            Adodc2.Recordset.Fields!depet_value = Adodc2.Recordset.Fields![value]
        Else
            Adodc2.Recordset.Fields!credit_value = Adodc2.Recordset.Fields![value]
        End If

        Adodc2.Recordset.Fields!DEV_ID_Line_No = i
        Adodc2.Recordset.update
        Adodc2.Recordset.MoveNext

    Next i

    On Error Resume Next
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
 
    Frame1.Visible = True

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.MoveFirst
        Adodc6.RecordSource = "select sum(depet_value) as depet_sum from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
        Adodc6.Refresh

        If Adodc6.Recordset.RecordCount > 0 And Not IsNull(Adodc6.Recordset.Fields!depet_sum) Then
            Text8.text = Adodc6.Recordset.Fields!depet_sum
        Else
            Text8.text = 0
            ' Exit Function
            
        End If
    
        Adodc6.RecordSource = "select sum(credit_value) as credit_sum from double_entry_voucher_with_name where  Notes_ID=" & NoteID.text
        Adodc6.Refresh

        If Adodc6.Recordset.RecordCount > 0 And Not IsNull(Adodc6.Recordset.Fields!credit_sum) Then
            Text9.text = Adodc6.Recordset.Fields!credit_sum
        Else
            Text9.text = 0
            ' Exit Function
            
        End If

    Else
        Text8.text = 0
        Text9.text = 0
    End If

    If Text8.text = "" Or Text9.text = "" Then

    Else
        Text10.text = Abs(val(Text8.text) - val(Text9.text))
    End If

    If Text8.text = 0 And Text9.text = 0 Then
        Text10.text = 0
    End If

    If Adodc2.Recordset.RecordCount > 0 Then Adodc2.Recordset.MoveFirst

    For i = 1 To Adodc2.Recordset.RecordCount
        Adodc2.Recordset.Fields![Index] = i
        Adodc2.Recordset.update
        Adodc2.Recordset.MoveNext
  
    Next i

    If Adodc2.Recordset.RecordCount > 0 Then Adodc2.Recordset.MoveFirst
  
End Function

Private Sub Command4_Click()
    On Error Resume Next

    If my_language = "E" Then
        If Text1.text = "" Then MsgBox "enter voucher no first", vbCritical: Exit Sub
    Else

        If Text1.text = "" Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·”‰œ", vbCritical: Exit Sub

    End If

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.delete
        Adodc2.Refresh
        detect_error
    End If

End Sub

Private Sub Command40_Click(Index As Integer)
    On Error Resume Next
    Calendar1.Visible = True
    Calendar1.value = DateValue(Now)
    Calendar1.left = Frame10.left + Command40(Index).left
    Calendar1.top = Frame10.top + Command40(Index).top
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    Dim x As Integer

    'x = Adodc2.Recordset.Index
    If Adodc2.Recordset.RecordCount > 0 Then

        For i = 1 To Adodc2.Recordset.RecordCount
            Adodc2.Recordset.Fields![Index] = i

            Adodc2.Recordset.update
        Next i

        'For i = 0 To x
        'Adodc2.Recordset.MoveNext

        'Adodc2.Recordset.Update
        'Next i

        'Adodc2.Refresh

        DataGrid1.Refresh
        'Adodc2.Recordset.MoveLast

        detect_error
    End If

End Sub

Private Sub Command6_Click()
    On Error Resume Next
    end_save

End Sub

Function end_save()
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
    Adodc2.Refresh

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.MoveFirst
    End If

    detect_error
    On Error Resume Next
    Dim x As Integer

    If val(Text10.text) > 0 Then
        If my_language = "E" Then
            x = MsgBox("can not save because different between depit and credit value" & Text10.text, vbCritical)
        Else
            x = MsgBox("·« Ì„ﬂ‰ Õ›Ÿ Â–« «·”‰œ ·ÊÃÊœ ›—ﬁ »Ì‰ «·œ«∆‰ Ê «·„œÌ‰ ÊﬁÌ„ … Â·  —Ìœ «·Œ—ÊÃ »œÊ‰ Õ›Ÿ" & Text10.text, vbCritical + vbYesNo)
                    
        End If

        If x = vbYes Then
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
            Adodc2.Refresh

            If Adodc2.Recordset.RecordCount > 0 Then
                Adodc2.Recordset.MoveFirst
            End If

            If Adodc2.Recordset.RecordCount > 0 Then
                Adodc2.Recordset.MoveFirst

                For i = 1 To Adodc2.Recordset.RecordCount
                    Adodc2.Recordset.delete
                    Adodc2.Recordset.MoveNext
                Next i
                     
                Text10.text = 0
            End If

            end_save_ok = True
        End If

        end_save_ok = False
    Else
        end_save_ok = True
 
    End If

End Function

Private Sub DataCombo1_Click(Area As Integer)
    On Error Resume Next

    If DataCombo1.text <> "" Then
        Text4.text = ""
        Text5.text = ""
        Adodc5.RecordSource = "select * from accounts where  last_account=1 and  Account_Serial='" & DataCombo1.text & "'"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount > 0 And Not IsNull(Adodc5.Recordset.Fields!account_serial) Then

            DataCombo3.text = Adodc5.Recordset.Fields!account_name
        Else

            If my_language = "E" Then
                MsgBox "error in this account name to fix error goto account index screen", vbCritical
            Else
                MsgBox "Â‰«ﬂ Œÿ√ ›Ì «”„ «·Õ”«» —«Ã⁄ «·œ·Ì· «·„Õ«”»Ì", vbCritical
                DataCombo1.text = ""
            End If
        End If

    End If

End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
    On Error Resume Next
    account_info_bar.Show
    account_info_bar.item_code = DataCombo1.text
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Acccount_search.Show
        Acccount_search.case_id = 30
        Acccount_search.sandat_pc_no = Adodc1.Recordset.Fields!sandat_pc_no
        Acccount_search.SANAD_TYPE = "”‰œ ﬁÌœ"
    End If

    If KeyCode = vbKeyF6 Then
        account_index.Show
    End If

    If KeyCode = vbKeyF5 Then
        Adodc4.Refresh
        DataCombo1.ReFill
    End If

End Sub

Private Sub DataCombo2_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF6 Then
        ked_types.Show
    End If

    If KeyCode = vbKeyF5 Then
        Adodc3.Refresh
        DataCombo2.ReFill
    End If

End Sub

Private Sub DataCombo3_Change()
    On Error Resume Next

    If DataCombo23.text <> "" Then
        Text4.text = ""
        Text5.text = ""
        Adodc5.RecordSource = "select * from accounts where last_account=1 and  Account_Name like '%" & DataCombo3.text & "%'"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount > 0 And Not IsNull(Adodc5.Recordset.Fields!account_serial) Then

            DataCombo1.text = Adodc5.Recordset.Fields!account_serial
        Else
            DataCombo1.text = ""
        End If

    End If

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, _
                            Shift As Integer)
    On Error Resume Next

    If KeyCode = 13 Then
        Label21_Click
    End If

    If KeyCode = 46 Then
        ALLButton4_Click
    End If

End Sub

Function sand_numbering()
    Branch_NO = 1
    Departement = 1
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    auto_sanad_no = ""
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & Branch_NO & " and departement='" & departement_name & "' and  sanad_no=0"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at
        end_at = numbering.Recordset.Fields!end_at
    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where  branch_no=" & Branch_NO & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
        detect_no.Refresh
    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where  branch_no=" & Branch_NO & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where  branch_no=" & Branch_NO & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
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

Private Sub DcAccount1_Click(Area As Integer)
    On Error Resume Next

    If DCAccount1.text = "" Then Exit Sub
    Dim My_SQL As String

    My_SQL = "select Account_Name from ACCOUNTS where Account_Serial='" & DCAccount1.text & "'"
 
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    DCAccount2.text = rec.Fields("Account_Name").value
End Sub

Private Sub DcAccount2_Change()
    On Error Resume Next

    If DCAccount2.text = "" Then Exit Sub
    Dim My_SQL As String

    My_SQL = "select Account_Serial from ACCOUNTS where Account_Name ='" & DCAccount2.text & "'"
 
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    DCAccount1.text = rec.Fields("Account_Serial").value
    'DcAccount2.SetFocus
    'SendKeys "{f4}"
End Sub

Private Sub DcAccount2_Click(Area As Integer)
    On Error Resume Next

    If DCAccount2.text = "" Then Exit Sub
    Dim My_SQL As String

    My_SQL = "select Account_Serial from ACCOUNTS where Account_Name ='" & DCAccount2.text & "'"
 
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    DCAccount1.text = rec.Fields("Account_Serial").value
End Sub

Private Sub Form_Activate()
    On Error Resume Next

    If first_run = True Then
        'Command1_Click (0)
        first_run = False
    End If
 
    'user_priviliges_adodc.ConnectionString = connection_string: user_priviliges_adodc.CommandType = adCmdText
    '    If my_language = "E" Then
    '    user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Val(current_user) & "and [no]='" & screen_name.Caption & "'"
    '    Else
    '    user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Val(current_user) & "and [no]='" & screen_name.Caption & "'"
    '
    '    End If
    'user_priviliges_adodc.Refresh

    '    If user_priviliges_adodc.Recordset.RecordCount = 0 Then
    '            If my_language = "E" Then
    '        MsgBox "NOT allowed ", vbCritical
    '
    '        Else
    '        MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
    '        End If
    '   Unload Me
    '    End If
    '
    'If user_priviliges_adodc.Recordset.Fields![View] = False Then
    '        If my_language = "E" Then
    '        MsgBox "NOT allowed ", vbCritical
    '
    '        Else
    '        MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
    '        End If
    '
    'Unload Me
    'End If

    'Command1(0).Enabled = user_priviliges_adodc.Recordset.Fields![add_new]
    'Command1(1).Enabled = user_priviliges_adodc.Recordset.Fields![Save]
    'Command1(2).Enabled = user_priviliges_adodc.Recordset.Fields![Delete]

End Sub

Private Sub Form_Load()

    Dim i As Integer
    Dim My_SQL As String

    'My_SQL = "select * from Notes where NoteType=200"
    'Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient

    'RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    'Me.TxtModFlg.text = "R"
    'Resize_Form Me
    'load ACCOUNTS -----------------------------------------------
    My_SQL = "  select Account_code,Account_Serial from ACCOUNTS  where last_account=1"

    fill_combo DCAccount1, My_SQL

    My_SQL = "  select Account_Serial,Account_Name from ACCOUNTS  where last_account=1"

    fill_combo DCAccount2, My_SQL

    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DcboUsers
 
    Me.DcboUsers.BoundText = user_id

    On Error Resume Next
    Branch_NO = 1
    departement_name = 1

    '
    
    If my_language = "E" Then
        Print_cmd.Caption = "Print"

        CMD_language.ToolTipText = "Change Language"
        Command13.ToolTipText = "F3 Voucher Search "

        Label46.Visible = True
        Label42.Visible = True
        Label45.Visible = True
        Frame3.left = 12840

        Label31.Visible = False
        Label39.Visible = False

        Label26.left = 9960

        Me.dept_lbl = departement_name
        Me.emp_name_lbl = current_user_name
        InfoE.Visible = True
        infoA.Visible = False
    Else

        emp_a.Caption = current_user_name
        dep_a.Caption = departement_name
   
        infoA.Visible = True
        InfoE.Visible = False
    End If

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If my_language = "E" Then
        Label9.left = 1380
        DataCombo1.left = 1380
        DataCombo1.RightToLeft = False
        DataCombo2.RightToLeft = False
        DataCombo3.left = 3480
        'DataCombo3.Alignment = 0
        Label3.left = 3480
     
        Text5.left = 7600
        Text5.Alignment = 0
        Label4.left = 6000
     
        ' Text6.Left = 7680 + 1800
        Text6.Alignment = 0
     
        Label11.left = Label26.left
     
        ' Label26.Left = 1380
     
        Text4.left = 6000
        Text4.Alignment = 0
     
        Label10.left = Label4.left + 1500
     
        DataGrid1.RightToLeft = False
        Frame5.Visible = True
        Frame7.Visible = False
 
        txtdate.Alignment = 0
        Text7.Alignment = 0
        Text1.Alignment = 0
        Text3.Alignment = 0
        Text12.Alignment = 0
        Text11.Alignment = 0
   
        DataCombo2.RightToLeft = False
        DataCombo3.RightToLeft = False
     
        CMD_language.Caption = "⁄—»Ì"
        'Frame7.Visible = True
        Frame9.Visible = True
        Frame11.Visible = True
        Frame13.Visible = True
     
        Frame16.Visible = False
        Frame15.Visible = False
        Frame14.Visible = False
        Frame12.left = 9500
        Frame8.left = 0
        ALLButton1.Caption = "Attachments"
        Label16.Caption = "Exit"
        Label33.Caption = " Journal entry"
 
        Me.Caption = Label33.Caption
        Label26.Caption = "Distribution"

        Label9.Caption = "Customer ID"
        Label3.Caption = "Customer Name"
        Label4.Caption = "depit"
   
        Label10.Caption = "Credit"
        Label11.Caption = "Cost Center"
        Label25.Caption = "New Row"
        Label32.Caption = "Delete Row"
        Label21.Caption = "Update Row"
        Label7.Caption = "Search"
        Command1(0).Caption = "new"
        Command1(1).Caption = "save"
        Command1(2).Caption = "delete"
  
        Command1(4).Caption = "ID"
        Command1(5).Caption = "Type"
        Command1(6).Caption = "Date"
        Command1(7).Caption = "Source"
        Command1(8).Caption = "Attachments"
        Command1(9).Caption = "Description"
        Command1(10).Caption = "Value"
        Command1(11).Caption = "Copy From"
        Command1(12).Caption = "end"
   
        Label22.Caption = "End save"
        Adodc1.Caption = "move"
        Label13.Caption = "Total depit"
        Label14.Caption = "Total credit"
        Label15.Caption = "Total"
  
    End If
 
    'LoadSettings
    connection_string = Cn.ConnectionString
    Branch_NO = 1
    departement_name = 1

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from notes  where branch_no=" & Branch_NO & " and departement='" & departement_name & "' and  NoteType=200 "
    Adodc1.Refresh

    ' NOT (sanad_no IS NULL) and
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveLast

        'If IsNull(Adodc1.Recordset.Fields!sanad_no) Then
        'GoTo ll
        'End If

    End If

    '  Adodc1.Recordset.AddNew
    '  Adodc1.Recordset.Fields!branch_no = branch_no
    'Adodc1.Recordset.Fields!user_name = current_user_name
    '   Adodc1.Recordset.Fields!DEPARTEMENT = departement_name
    '
    '
    '    Adodc1.Recordset.Fields!type = "”‰œ ﬁÌœ"
    '    Text3.text = "ÌœÊÌ"
    '       txtdate.text = DateValue(Now)
    '    Text3.text = "ÌœÊÌ"
    '
    '     Adodc1.Recordset.Fields!attachment = vbFalse
    '    Adodc1.Recordset.update
    '    Adodc1.Recordset.MoveLast
    '

    'll:

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select *  from double_entry_voucher_with_name where Notes_ID=0"
    Adodc2.Refresh

    'Adodc2.Recordset.AddNew
    'Adodc2.Recordset.Fields!sandat_pc_no = Text7.Text
    ''Adodc2.Recordset.Fields!sanad_no = Text1.Text

    'Adodc2.Recordset.Fields!SANAD_TYPE = "”‰œ ﬁÌœ"
    'Adodc2.Recordset.Fields!sanad_source = "ÌœÊÌ"
    'Adodc2.Recordset.Fields!box_name = DataCombo3.Text
    'Adodc2.Recordset.Fields!bona_3la = Text11.Text
    'Adodc2.Recordset.Fields!Date = DateValue(Now)

    'Adodc2.Recordset.Update
    'Adodc2.Refresh
    'Adodc2.Recordset.MoveLast
    'DataGrid1.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select *  from ked_types  where type_id=0"
    Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    'Adodc4.RecordSource = "select * from account_index where black_list=0 and ( account_type='›—⁄Ì' or account_type='sub' )"
    Adodc4.RecordSource = "select * from ACCOUNTS where last_account=1"
    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from ACCOUNTS  where last_account=1"
    Adodc5.Refresh

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    Adodc6.RecordSource = "select * from sandat_ked_details "
    Adodc6.Refresh

    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText

    'new_record = False
    If NoteID.text = "" Then Exit Sub
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
    Adodc2.Refresh
    end_save_ok = False
    first_run = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    end_save

    If end_save_ok = True Then
        Unload Me
    Else
        Exit Sub
    End If

End Sub

Private Sub Image2_Click()
    frmCalculator.Show
    frmCalculator.case_id = 0

End Sub

Private Sub Image3_Click()
    frmCalculator.Show
    frmCalculator.case_id = 1
End Sub

Private Sub Label16_Click()
    On Error Resume Next
    end_save

    If end_save_ok = True Then
        Unload Me
    End If

End Sub

Private Sub Label17_Click()
    On Error Resume Next

    'detect_error
    'end_save
    If NoteID.text = "" Then Exit Sub
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
    Adodc2.Refresh
    DataGrid1.Refresh

    If end_save_ok = True Then
        If Adodc1.Recordset.EOF = False Then
            Adodc1.Recordset.MoveNext

        End If
    End If

    detect_error

End Sub

Private Sub Label18_Click()
    On Error Resume Next

    'detect_error
    'end_save
    If NoteID.text = "" Then Exit Sub
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
    Adodc2.Refresh
    DataGrid1.Refresh

    If end_save_ok = True Then

        If Adodc1.Recordset.BOF = False Then
            Adodc1.Recordset.MovePrevious
 
        End If
    End If

    detect_error

End Sub

Private Sub Label21_Click()
    On Error Resume Next

    If my_language = "E" Then
        If Text1.text = "" Then MsgBox "enter voucher no first", vbCritical: Exit Sub

    Else

        If Text1.text = "" Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·”‰œ", vbCritical: Exit Sub
    End If

    Command5_Click
End Sub

Private Sub Label22_Click()
    On Error Resume Next
    Label21_Click
    Command6_Click
End Sub

Private Sub Label23_Click()
    On Error Resume Next

    'detect_error
    'end_save
    If NoteID.text = "" Then Exit Sub
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
    Adodc2.Refresh
    DataGrid1.Refresh

    If end_save_ok = True Then
        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.MoveLast
 
        End If
    End If

    detect_error
End Sub

Private Sub Label24_Click()
    On Error Resume Next

    'detect_error
    'end_save
    If NoteID.text = "" Then Exit Sub
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from double_entry_voucher_with_name where Notes_ID=" & NoteID.text
    Adodc2.Refresh
    DataGrid1.Refresh

    If end_save_ok = True Then
        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.MoveFirst

        End If
    End If

    detect_error
End Sub

Private Sub Label25_Click()
    On Error Resume Next
 
    Command3_Click

End Sub

Private Sub Label26_Click()
    On Error Resume Next

    If Adodc2.Recordset.RecordCount = 0 Then Exit Sub

    'If Adodc2.Recordset.Fields!dist = vbTrue Then
    'marakes_taklefa_tawze3.Show
    'marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
    'marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
    'marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where opr_id =" & Adodc2.Recordset.Fields!opr_id
    'marakes_taklefa_tawze3.Adodc3.Refresh
    'marakes_taklefa_tawze3.DataGrid1.Refresh
    '
    'Exit Sub
    'End If

    If Adodc2.Recordset.Fields!depet_value <> 0 Then
        marakes_taklefa_tawze3.Show

        marakes_taklefa_tawze3.value.Caption = Adodc2.Recordset.Fields!depet_value ' Text4.Text
        marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"

    End If

    If Adodc2.Recordset.Fields!credit_value <> 0 Then
        marakes_taklefa_tawze3.Show

        marakes_taklefa_tawze3.value.Caption = Adodc2.Recordset.Fields!credit_value 'Text5.Text
        marakes_taklefa_tawze3.type.Caption = "œ«∆‰"
 
    End If

    marakes_taklefa_tawze3.opr_type = "”‰œ ﬁÌœ"
    marakes_taklefa_tawze3.opr_id = Adodc2.Recordset.Fields!opr_id
    marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
    marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
    marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where opr_id =" & Adodc2.Recordset.Fields!opr_id
    marakes_taklefa_tawze3.Adodc3.Refresh

    If Adodc2.Recordset.Fields!depet_value = 0 And Adodc2.Recordset.Fields!credit_value = 0 Then
        If my_language = "E" Then
            MsgBox "enter depit/credit value first to complete this process", vbCritical

        Else
            MsgBox "·« Ì„ﬂ‰ « „«„ «·⁄„·Ì… «œŒ· «·ﬁÌ„… œ«∆‰ «Ê „œÌ‰ «Ê·«", vbCritical
        End If
    End If

    'On Error Resume Next

    'If Text4.Text <> "" Then
    'marakes_taklefa_tawze3.value.Caption = Text4.Text
    'marakes_taklefa_tawze3.type.Caption = "„œÌ‰"
    'marakes_taklefa_tawze3.Show
    'End If

    'If Text5.Text <> "" Then
    'marakes_taklefa_tawze3.value.Caption = Text5.Text
    'marakes_taklefa_tawze3.type.Caption = "œ«∆‰"
    'marakes_taklefa_tawze3.Show
    'End If
    '

    'If Text4.Text = "" And Text5.Text = "" Then
    ' If my_language = "E" Then
    ' MsgBox "enter depit/creit value to complete this process", vbCritical
    '
    ' Else
    'MsgBox "·« Ì„ﬂ‰ « „«„ «·⁄„·Ì… «œŒ· «·ﬁÌ„… œ«∆‰ «Ê „œÌ‰ «Ê·«", vbCritical
    'End If
    '
    'End If

End Sub

Private Sub Label32_Click()
    On Error Resume Next

    If my_language = "E" Then
        x = MsgBox("Confirm delete", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ Õ–› ”ÿ—", vbCritical + vbYesNo)
              
    End If
            
    If x = vbNo Then
        Exit Sub
    End If

    Command4_Click
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, _
                        Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Voucher_search.Show
        Voucher_search.title_lbl = "”‰œ ﬁÌœ"
    End If

End Sub

Private Sub Text10_Change()
    On Error Resume Next
    Timer1.Enabled = True
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, _
                         Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Voucher_search.Show
        Voucher_search.title_lbl = "”‰œ ﬁÌœ"
    End If

End Sub

Private Sub DataCombo3_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF6 Then
        account_index.Show
    End If

    If KeyCode = vbKeyF3 Then
        Acccount_search.Show
        Acccount_search.case_id = 30
        Acccount_search.sandat_pc_no = Adodc1.Recordset.Fields!sandat_pc_no
        Acccount_search.SANAD_TYPE = "”‰œ ﬁÌœ"
    End If

End Sub

Private Sub Text3_Click()
    On Error Resume Next
    Calendar1.Visible = False

End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, _
                        Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Voucher_search.Show
        Voucher_search.title_lbl = "”‰œ ﬁÌœ"
    End If

    If KeyCode = 13 Then
        On Error Resume Next
        Command3_Click
 
        Exit Sub
    End If

    For i = 1 To Len(Text4.text)

        If Asc(Mid$(Text4.text, i, 1)) < 48 Or Asc(Mid$(Text4.text, i, 1)) > 57 Then
       
            If my_language = "E" Then
                MsgBox "Enter Digit only", vbCritical
            Else
                MsgBox "«œŒ· «—ﬁ«„ ›ﬁÿ", vbCritical
            End If
              
            Text4.text = 0
            Text4.BackColor = vbRed
            Exit Sub
        End If

    Next i

End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, _
                        Shift As Integer)
    On Error Resume Next
 
    If KeyCode = vbKeyF3 Then
        Voucher_search.Show
        Voucher_search.title_lbl = "”‰œ ﬁÌœ"
    End If

    If KeyCode = 13 Then
        On Error Resume Next
        Command3_Click
 
        Exit Sub
    End If

    For i = 1 To Len(Text5.text)

        If Asc(Mid$(Text5.text, i, 1)) < 48 Or Asc(Mid$(Text5.text, i, 1)) > 57 Then
        
            If my_language = "E" Then
                MsgBox "Enter Digit only", vbCritical
            Else
                MsgBox "«œŒ· «—ﬁ«„ ›ﬁÿ", vbCritical
            End If
              
            Text5.text = 0
            Text5.BackColor = vbRed
            Exit Sub
        End If

    Next i

End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, _
                        Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Voucher_search.Show
        Voucher_search.title_lbl = "”‰œ ﬁÌœ"
    End If

End Sub

Private Sub Text7_Change()
    'On Error Resume Next
    'Adodc2.ConnectionString = connection_string
    'Adodc2.CommandType = adCmdText
    '
    '
    'If Text7.text = "" Or Not IsNumeric(Text7) Then Exit Sub
    'Adodc2.CommandType = adCmdText
    'Adodc2.RecordSource = "select * from sandat_ked_details where sandat_pc_no=" & Text7.text
    'Adodc2.Refresh
    detect_error
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next

    If val(Text10.text) = 0 Then
        Text10.BackColor = vbWhite
        Timer1.Enabled = False
 
    End If

    If Text10.BackColor = vbWhite Then
        Text10.BackColor = vbRed
    Else
        Text10.BackColor = vbWhite
    End If

End Sub

Private Sub txtdate_KeyUp(KeyCode As Integer, _
                          Shift As Integer)
16777215    On Error Resume Next

            If KeyCode = vbKeyF3 Then
                Voucher_search.Show
                Voucher_search.title_lbl = "”‰œ ﬁÌœ"
            End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    Select Case Index

        Case 0
            Label17_Click

        Case 1
            Label23_Click

        Case 2
            Label24_Click

        Case 3
            Label18_Click
    End Select

End Sub
