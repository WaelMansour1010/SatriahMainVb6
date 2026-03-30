VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form frm_maintenace 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7770
   ClientLeft      =   1065
   ClientTop       =   1845
   ClientWidth     =   14535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   14535
   Begin VB.Frame Frame9 
      Height          =   1215
      Left            =   5040
      TabIndex        =   75
      Top             =   5880
      Width           =   3375
      Begin ALLButtonS.ALLButton Command1 
         Height          =   350
         Index           =   3
         Left            =   1200
         TabIndex        =   76
         Top             =   240
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   609
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
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frm_maintenace.frx":0000
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
         Height          =   350
         Index           =   6
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   609
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
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frm_maintenace.frx":001C
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
         Height          =   350
         Index           =   7
         Left            =   2280
         TabIndex        =   78
         Top             =   240
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   609
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
         MICON           =   "frm_maintenace.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   465
         Left            =   840
         Top             =   720
         Width           =   1800
         _ExtentX        =   3175
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
      Begin VB.Label Label13 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   79
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   120
      TabIndex        =   61
      Top             =   720
      Width           =   5775
      Begin MSACAL.Calendar Calendar1 
         Height          =   2295
         Left            =   2280
         TabIndex        =   62
         Top             =   360
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
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   0
         Left            =   5040
         Picture         =   "frm_maintenace.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   360
         Width           =   492
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "motahamel_taklefa"
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
         ItemData        =   "frm_maintenace.frx":08B6
         Left            =   2280
         List            =   "frm_maintenace.frx":08C0
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   3360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "error_description"
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
         Height          =   1560
         Left            =   2280
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   72
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "date_in_warsha"
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   360
         Width           =   2655
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
         Begin VB.Label Label43 
            Caption         =   "Date IN"
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
            Left            =   0
            TabIndex        =   70
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label52 
            Caption         =   "Error Description "
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
            Left            =   0
            TabIndex        =   69
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label53 
            Caption         =   "Error By"
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
            Left            =   120
            TabIndex        =   68
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label54 
            Caption         =   "Cost paid by"
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
            Left            =   0
            TabIndex        =   67
            Top             =   4080
            Width           =   2415
         End
         Begin VB.Label Label9 
            Caption         =   "TIME IN"
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
            Left            =   0
            TabIndex        =   66
            Top             =   960
            Width           =   1815
         End
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "TIME_IN"
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "...."
         Height          =   255
         Left            =   5160
         TabIndex        =   63
         Top             =   1080
         Width           =   375
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "’Ì«‰… Œ«—ÃÌ…"
      DataField       =   "OUT_WARSHA"
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
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5880
      TabIndex        =   59
      Top             =   480
      Width           =   2775
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   3600
      TabIndex        =   54
      Top             =   7080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4800
         TabIndex        =   57
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   55
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   0
      TabIndex        =   49
      Top             =   6120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   52
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   51
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   12480
      Picture         =   "frm_maintenace.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "»ÕÀ ⁄‰ ’Ì«‰… ”Ì«—…F3"
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   6240
      TabIndex        =   25
      Top             =   840
      Width           =   2295
      Begin VB.Label Label4 
         Caption         =   "Êﬁ  »œ«Ì… «·’Ì«‰…"
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
         Left            =   480
         TabIndex        =   60
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "„ Õ„· ‰ﬂ·›… «·«’·«Õ"
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
         Left            =   240
         TabIndex        =   28
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Ê’› «·’Ì«‰…"
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
         Left            =   720
         TabIndex        =   27
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " «—ÌŒ  »œ«Ì… «·’Ì«‰…"
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
         Left            =   480
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2415
      Left            =   120
      TabIndex        =   39
      Top             =   840
      Width           =   1575
      Begin VB.Frame Frame7 
         Height          =   2175
         Left            =   6960
         TabIndex        =   44
         Top             =   6120
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "frm_maintenace.frx":2268
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
            Left            =   120
            TabIndex        =   45
            Top             =   1560
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "frm_maintenace.frx":2284
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
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "frm_maintenace.frx":22A0
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
            TabIndex        =   47
            Top             =   1440
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1935
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "»—ﬁ„ «·„⁄œÂ/«·”Ì«—…"
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
            MICON           =   "frm_maintenace.frx":22BC
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
            Index           =   5
            Left            =   120
            TabIndex        =   42
            Top             =   1320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "»√”„ «·”«∆ﬁ"
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
            MICON           =   "frm_maintenace.frx":22D8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SuperLablel.SuperLabel SuperLabel2 
            Height          =   615
            Left            =   240
            TabIndex        =   43
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1085
            Text            =   "»ÕÀ"
            ColorGeneral    =   16711680
            ColorGeneral    =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   4815
      Left            =   6000
      TabIndex        =   30
      Top             =   720
      Width           =   5655
      Begin VB.CheckBox Check1 
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   5280
         TabIndex        =   6
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "opr_date"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
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
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "opr_id"
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   -480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   120
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label Label59 
            Alignment       =   2  'Center
            Caption         =   "Maintenance Type"
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
            Left            =   120
            TabIndex        =   38
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label Label62 
            Caption         =   "insurance"
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
            Left            =   120
            TabIndex        =   37
            Top             =   3000
            Width           =   2415
         End
         Begin VB.Label Label60 
            Caption         =   "work shop"
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
            Left            =   120
            TabIndex        =   36
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label58 
            Caption         =   "Opr#"
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
            Left            =   120
            TabIndex        =   35
            Top             =   -240
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label57 
            Caption         =   "Operator"
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
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label56 
            Caption         =   "equipment Code"
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
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label55 
            Caption         =   "Opr Date"
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
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_maintenace.frx":22F4
         DataField       =   "car_no"
         DataSource      =   "Adodc1"
         Height          =   480
         Left            =   2760
         TabIndex        =   2
         Top             =   840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         BackColor       =   12632256
         ListField       =   "fullcode"
         BoundColumn     =   "fullcode"
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
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frm_maintenace.frx":2309
         DataField       =   "driver_name"
         DataSource      =   "Adodc1"
         Height          =   480
         Left            =   2760
         TabIndex        =   3
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         BackColor       =   12632256
         ListField       =   "driver_name"
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
         Bindings        =   "frm_maintenace.frx":231E
         DataField       =   "maintenance_type"
         DataSource      =   "Adodc1"
         Height          =   480
         Left            =   2760
         TabIndex        =   4
         Top             =   2040
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         BackColor       =   12632256
         ListField       =   "maintenance_type"
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
         Bindings        =   "frm_maintenace.frx":2333
         DataField       =   "warsha_name"
         DataSource      =   "Adodc1"
         Height          =   480
         Left            =   2760
         TabIndex        =   5
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         BackColor       =   12632256
         ListField       =   "wersha_name"
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
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Language  «··€…"
      Top             =   120
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
      MICON           =   "frm_maintenace.frx":2348
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
      Height          =   690
      Left            =   2160
      Top             =   7200
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   1217
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
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   11760
      TabIndex        =   17
      Top             =   840
      Width           =   1695
      Begin VB.Label Label23 
         Caption         =   "—ﬁ„ «·Õ—ﬂ…"
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
         Left            =   240
         TabIndex        =   24
         Top             =   -480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "„€ÿÏ »«· √„Ì‰"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label21 
         Caption         =   "«”„ «·„” √Ã—"
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
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "«”„ «·„«·ﬂ"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "ﬂÊœ «·ÊÕœ… /«·⁄ﬁ«—"
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
         Left            =   0
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   " «—ÌŒ «·ÌÊ„"
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
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   1560
      TabIndex        =   14
      Top             =   7560
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
         Caption         =   "M24"
         Height          =   255
         Left            =   3360
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.CommandButton Command40 
      Height          =   492
      Index           =   1
      Left            =   -720
      Picture         =   "frm_maintenace.frx":2364
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   -240
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "tklef"
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
      Left            =   10920
      TabIndex        =   12
      Top             =   7440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "date_out_warsha"
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2040
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2040
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   2040
      Top             =   6480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   2040
      Top             =   6720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc numbering 
      Height          =   585
      Left            =   -720
      Top             =   5640
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
      Left            =   0
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   690
      Left            =   -600
      Top             =   5400
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   1217
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
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   " ÕÊÌ· «·Ï «„— ‘€·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label11 
      Caption         =   " «—ÌŒ «·Œ—ÊÃ „‰ «·Ê—‘…"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "’Ì«‰… «·ÊÕœ«  Ê «·⁄ﬁ«—« "
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
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frm_maintenace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Dim first_run As Boolean
Dim bindex As Integer
Dim auto_sanad_no As String
Dim numbering_type As Integer


Private Sub Calendar1_Click()
On Error Resume Next
If bindex = 0 Then
Text3.Text = Calendar1.value
End If

If bindex = 1 Then
    Text5.Text = Calendar1.value
End If
Calendar1.Visible = False
End Sub

 
 

Private Sub Check2_Click()
On Error Resume Next
If Check2.value = 1 Then
Label14.Visible = False
Else
Label14.Visible = True

End If
End Sub

Private Sub CMD_language_Click()
On Error Resume Next

If CMD_language.Caption = "EN" Then
my_language = "E"
 
Call Reload(Me)

 
Else
my_language = "A"
 
Call Reload(Me)
End If
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Select Case Index

Case 7
'Adodc1.ConnectionString = connection_string
'Adodc1.CommandType = adCmdText
'Adodc1.RecordSource = "select * from maintenance  where   Amr_shogl=0"
'Adodc1.Refresh


    Adodc1.Recordset.AddNew
    
 Adodc1.Recordset.Fields!branch_no = branch_no
 
   Adodc1.Recordset.Fields!DEPARTEMENT = departement_name


  Adodc1.Recordset.Fields!user_name = current_user_name
    Adodc1.Recordset.Fields!opr_date = DateValue(Now)

     Adodc1.Recordset.Fields!out_maintenance = vbFalse
Adodc1.Recordset.Fields!repaired = vbFalse
Adodc1.Recordset.Fields!moghat_belt2men = vbFalse
Adodc1.Recordset.Fields!Amr_shogl = vbFalse
Adodc1.Recordset.Fields!converted = vbFalse
Adodc1.Recordset.Fields!OUT_WARSHA = vbFalse

'     Adodc1.Recordset.Fields!moghat_belt2men = vbTrue
'        Adodc1.Recordset.Fields!OUT_WARSHA = vbFalse
'           Adodc1.Recordset.Fields!converted = 0
    Adodc1.Recordset.Update
Adodc1.Recordset.MoveLast

'Adodc1.RecordSource = "select * from maintenance  where converted=0"
'Adodc1.Refresh

'If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
'    Adodc1.Recordset.MoveLast
    
Case 3
    If my_language = "E" Then
    If DataCombo1.Text = "" Then MsgBox "select Equipments  ", vbCritical: Exit Sub
    If DataCombo2.Text = "" Then MsgBox " select driver« ", vbCritical: Exit Sub
    If DataCombo3.Text = "" Then MsgBox " select maintenance type ", vbCritical: Exit Sub
    If DataCombo4.Text = "" Then MsgBox " select work shop", vbCritical: Exit Sub
   
    If Text3.Text = "" Then MsgBox "specify date in ", vbCritical: Exit Sub
    If Text2.Text = "" Then MsgBox "write error description", vbCritical: Exit Sub
    Else
    If DataCombo1.Text = "" Then MsgBox "«Œ — —ﬁ„ «·„⁄œÂ «Ê·«  ", vbCritical: Exit Sub
    If DataCombo2.Text = "" Then MsgBox " «Œ — «”„ «·”«∆ﬁ «Ê·« ", vbCritical: Exit Sub
    If DataCombo3.Text = "" Then MsgBox " «Œ — ‰Ê⁄ «·’Ì«‰… «Ê·« ", vbCritical: Exit Sub
    If DataCombo4.Text = "" Then MsgBox " «Œ — «·Ê—‘… «Ê·« ", vbCritical: Exit Sub
   
    If Text3.Text = "" Then MsgBox "Õœœ  «—ÌŒ œŒÊ· «·Ê—‘… ", vbCritical: Exit Sub
    If Text2.Text = "" Then MsgBox "—Ã«¡ Ê’› «·⁄ÿ· ", vbCritical: Exit Sub
    End If
'        Adodc1.Recordset.Fields!opr_date = DateValue(Now)
Adodc1.Recordset.Fields!cars_NO = DataCombo1.Text

 Adodc1.Recordset.Fields!branch_no = branch_no
 Adodc1.Recordset.Fields!DEPARTEMENT = departement_name


  Adodc1.Recordset.Fields!user_name = current_user_name
  
Adodc1.Recordset.Fields!driver_name = DataCombo2.Text
Adodc1.Recordset.Fields!maintenance_type = DataCombo3.Text
Adodc1.Recordset.Fields!warsha_name = DataCombo4.Text
Adodc1.Recordset.Fields!error_person = DataCombo5.Text
   Adodc1.Recordset.Update
   
' Adodc1.Recordset.MoveLast
Case 6
 If my_language = "E" Then
x = MsgBox("confirm delete", vbCritical + vbYesNo)

Else
x = MsgBox("Â· «‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbCritical + vbYesNo)
End If

If x = vbNo Then
Exit Sub
End If

    If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    DataGrid1.Refresh
    End If

 





End Select

End Sub

Private Sub Command13_Click()
On Error Resume Next

workorder_search.Show
workorder_search.case_id = 1

End Sub

Private Sub Command2_Click()
 On Error Resume Next
 Unload frmspecify_time
frmspecify_time.Show
'frmspecify_time.Left = Command2(Index).Left
'frmspecify_time.Top = Command2(Index).Top
frmspecify_time.case_id = 500
End Sub

Private Sub Command40_Click(Index As Integer)
On Error Resume Next
bindex = Index
Calendar1.Visible = True
Calendar1.value = DateValue(Now)
'Calendar1.Top = Command40(Index).Top
'Calendar1.Left = Command40(Index).Left
End Sub

 

Private Sub DataCombo1_Click(Area As Integer)
'Adodc2.Refresh
'DataCombo1.ReFill
End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
On Error Resume Next

'If DataCombo1.Text = "" Then Exit Sub
Car_info_bar.Show
Car_info_bar.Car_no = DataCombo1.Text
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF3 Then
Cars_search.Show
Cars_search.case_id = 1
End If

If KeyCode = vbKeyF6 Then
frmcars.Show
End If


If KeyCode = vbKeyF5 Then
Adodc2.Refresh
DataCombo1.ReFill
End If


End Sub

Private Sub DataCombo2_Click(Area As Integer)
'Adodc3.Refresh
'DataCombo2.ReFill
End Sub

Private Sub DataCombo2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF3 Then
emp_search.Show
emp_search.case_id = 2
End If

If KeyCode = vbKeyF6 Then
EMPLOYEES.Show
End If


If KeyCode = vbKeyF5 Then
Adodc3.Refresh
DataCombo2.ReFill
End If


End Sub

Private Sub DataCombo3_Click(Area As Integer)
Adodc4.Refresh
DataCombo3.ReFill
End Sub

Private Sub DataCombo3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF6 Then
frmmaintenace_type.Show
End If


If KeyCode = vbKeyF5 Then
Adodc4.Refresh
DataCombo3.ReFill
End If


End Sub

Private Sub DataCombo4_Click(Area As Integer)
'Adodc5.Refresh
'DataCombo4.ReFill
End Sub

Private Sub DataCombo4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF3 Then
warsha_search.Show
warsha_search.case_id = 1
End If

If KeyCode = vbKeyF6 Then
frmwarsha.Show
End If

If KeyCode = vbKeyF5 Then
Adodc5.Refresh
DataCombo4.ReFill
End If


End Sub

Private Sub DataCombo5_Click(Area As Integer)
Adodc3.Refresh
DataCombo5.ReFill
End Sub

Private Sub DataCombo5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF3 Then
emp_search.Show
emp_search.case_id = 3
End If

If KeyCode = vbKeyF6 Then
EMPLOYEES.Show
End If



If KeyCode = vbKeyF5 Then
Adodc3.Refresh
DataCombo5.ReFill
End If

End Sub

Function sand_numbering()
On Error Resume Next

auto_sanad_no = ""
numbering.ConnectionString = connection_string
numbering.CommandType = adCmdText
numbering.RecordSource = "select * from sanad_numbering where branch_no=" & branch_no & " and departement='" & departement_name & "' and  sanad_no=7"
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
detect_no.RecordSource = "select max(Amr_shogl_no) as last_sand_no from  maintenance where  branch_no=" & branch_no & " and departement='" & departement_name & "' and numbering_type=" & numbering_type
detect_no.Refresh
Else
If numbering_type = 2 Then
 
detect_no.ConnectionString = connection_string
detect_no.CommandType = adCmdText
detect_no.RecordSource = "select max(Amr_shogl_no) as last_sand_no from  maintenance where  branch_no=" & branch_no & " and departement='" & departement_name & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
detect_no.Refresh

Else
If numbering_type = 3 Then
 
detect_no.ConnectionString = connection_string
detect_no.CommandType = adCmdText
detect_no.RecordSource = "select max(Amr_shogl_no) as last_sand_no from  maintenance where  branch_no=" & branch_no & " and departement='" & departement_name & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
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
End Function

Private Sub Form_Activate()
On Error Resume Next
If first_run = False Then
'    Adodc1.Recordset.AddNew
'    Adodc1.Recordset.Fields!opr_date = DateValue(Now)
'     Adodc1.Recordset.Fields!moghat_belt2men = vbTrue
'        Adodc1.Recordset.Fields!OUT_WARSHA = vbFalse
'           Adodc1.Recordset.Fields!converted = 0
'    Adodc1.Recordset.Update

'Adodc1.RecordSource = "select * from maintenance  where NOT (car_no IS NULL) and converted=0"
'Adodc1.Refresh

'If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
'    Adodc1.Recordset.MoveLast

'first_run = True
End If

 
user_priviliges_adodc.ConnectionString = connection_string: user_priviliges_adodc.CommandType = adCmdText
    If my_language = "E" Then
    user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Val(current_user) & "and [no]='" & screen_name.Caption & "'"
    Else
    user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Val(current_user) & "and [no]='" & screen_name.Caption & "'"
    
    End If
user_priviliges_adodc.Refresh

    If user_priviliges_adodc.Recordset.RecordCount = 0 Then
            If my_language = "E" Then
        MsgBox "NOT allowed ", vbCritical
        
        Else
        MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
        End If
 Unload Me
    End If

If user_priviliges_adodc.Recordset.Fields![View] = False Then
        If my_language = "E" Then
        MsgBox "NOT allowed ", vbCritical
        
        Else
        MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
        End If

Unload Me
End If

Command1(0).Enabled = user_priviliges_adodc.Recordset.Fields![add_new]
Command1(1).Enabled = user_priviliges_adodc.Recordset.Fields![Save]
Command1(2).Enabled = user_priviliges_adodc.Recordset.Fields![Delete]

End Sub

Private Sub Form_Load()
On Error Resume Next


    login.SkinFramework.ApplyWindow Me.hWnd


If my_language = "E" Then
CMD_language.ToolTipText = "Change Language"
Command13.ToolTipText = "F3 Maintenance process Search "

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

Dim tleft As Integer
 Me.Left = (MDIForm1.Width - Me.Width) / 2
    Me.Top = (MDIForm1.Height - Me.Height) / 2 - 500

On Error Resume Next
 If my_language = "E" Then
 Check2.Caption = "Out maintenance"
 Text1.Alignment = 0
  Text2.Alignment = 0
   Text3.Alignment = 0
    txtid.Alignment = 0
      DataCombo1.RightToLeft = False
    DataCombo2.RightToLeft = False
      DataCombo3.RightToLeft = False
        DataCombo4.RightToLeft = False
          DataCombo5.RightToLeft = False
             Combo1.RightToLeft = False
        Combo1.Clear
        Combo1.AddItem "Operator"
        Combo1.AddItem "Company"
        
 
  tleft = Frame3.Left
Frame3.Left = Frame4.Left
Frame4.Left = tleft
Frame8.Left = 11750
Check1.Left = 2760

 CMD_language.Caption = "⁄—»Ì"
 
Frame10.Visible = True
Frame11.Visible = True
Adodc1.Caption = "move"

Frame5.Visible = False
Frame2.Visible = False
 Command1(0).Caption = "New"
 Command1(1).Caption = "save"
 Command1(2).Caption = "Delete"
SuperLabel2.Text = "search"

Label3.Caption = "Equipments maintenance "
Me.Caption = Label3.Caption
Label14.Caption = "Convert to work order"
 

Command1(4).Caption = "by Equipments"
Command1(5).Caption = "by driver"
Me.Width = 13470
End If

LoadSettings


   
'Adodc1.RecordSource = "select * from maintenance  where NOT (car_no IS NULL) and  converted=0"
'Adodc1.Refresh
'Adodc1.RecordSource = "select * from maintenance  where NOT (car_no IS NULL) and  converted=0"
'Adodc1.Refresh

'If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
' Adodc1.Recordset.MoveLast
 
Adodc8.ConnectionString = connection_string
Adodc8.CommandType = adCmdText
Adodc8.RecordSource = "select * from inventory"
Adodc8.Refresh


 
Adodc2.ConnectionString = connection_string
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select * from CARS  where not(Car_no is null) "
Adodc2.Refresh

Adodc3.ConnectionString = connection_string
Adodc3.CommandType = adCmdText
Adodc3.RecordSource = "select * from drivers  where  NOT (driver_name ='')  "
Adodc3.Refresh


Adodc4.ConnectionString = connection_string
Adodc4.CommandType = adCmdText
Adodc4.RecordSource = "select * from maintenance_type where not(maintenance_type='')  "
Adodc4.Refresh


Adodc5.ConnectionString = connection_string
Adodc5.CommandType = adCmdText
Adodc5.RecordSource = "select * from wersha  where not (wersha_name='') "
Adodc5.Refresh


Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from maintenance  where converted=0 and branch_no=" & branch_no & " and departement='" & departement_name & "' and     Amr_shogl=0"
Adodc1.Refresh

'NOT (car_no IS NULL) and
If Adodc1.Recordset.RecordCount > 0 Then
Adodc1.Recordset.MoveLast

If IsNull(Adodc1.Recordset.Fields!DEPARTEMENT) Then
GoTo LL
End If

End If


Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields!branch_no = branch_no
  Adodc1.Recordset.Fields!user_name = current_user_name
  
     Adodc1.Recordset.Fields!opr_date = DateValue(Now)
 
     Adodc1.Recordset.Fields!out_maintenance = vbFalse
Adodc1.Recordset.Fields!repaired = vbFalse
Adodc1.Recordset.Fields!moghat_belt2men = vbFalse
Adodc1.Recordset.Fields!Amr_shogl = vbFalse
Adodc1.Recordset.Fields!converted = vbFalse
Adodc1.Recordset.Fields!OUT_WARSHA = vbFalse

       Adodc1.Recordset.Fields!cars_NO = ""
     
     
     'Adodc1.Recordset.Fields!moghat_belt2men = vbTrue
     ' Adodc1.Recordset.Fields!OUT_WARSHA = vbFalse
     ' Adodc1.Recordset.Fields!converted = 0
    Adodc1.Recordset.Update
    'Adodc1.Refresh
    Adodc1.Recordset.MoveLast
LL:
 


End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
first_run = False
'Adodc1.ConnectionString = connection_string
'Adodc1.CommandType = adCmdText
'Adodc1.RecordSource = "select * from maintenance  where  NOT (car_no IS NULL) and converted=0"
'Adodc1.Refresh
End Sub

Private Sub Label14_Click()
On Error Resume Next

' Command1_Click (1)

If Text1.Text = "" Then
Exit Sub
End If
    If my_language = "E" Then
    If DataCombo1.Text = "" Then MsgBox "select Equipments  ", vbCritical: Exit Sub
    If DataCombo2.Text = "" Then MsgBox " select driver« ", vbCritical: Exit Sub
    If DataCombo3.Text = "" Then MsgBox " select maintenance type ", vbCritical: Exit Sub
    If DataCombo4.Text = "" Then MsgBox " select work shop", vbCritical: Exit Sub
   
    If Text3.Text = "" Then MsgBox "specify date in ", vbCritical: Exit Sub
    If Text2.Text = "" Then MsgBox "write error description", vbCritical: Exit Sub
    Else
    If DataCombo1.Text = "" Then MsgBox "«Œ — —ﬁ„ «·„⁄œÂ «Ê·«  ", vbCritical: Exit Sub
    If DataCombo2.Text = "" Then MsgBox " «Œ — «”„ «·”«∆ﬁ «Ê·« ", vbCritical: Exit Sub
    If DataCombo3.Text = "" Then MsgBox " «Œ — ‰Ê⁄ «·’Ì«‰… «Ê·« ", vbCritical: Exit Sub
    If DataCombo4.Text = "" Then MsgBox " «Œ — «·Ê—‘… «Ê·« ", vbCritical: Exit Sub
    
    If Text3.Text = "" Then MsgBox "Õœœ  «—ÌŒ œŒÊ· «·Ê—‘… ", vbCritical: Exit Sub
    If Text2.Text = "" Then MsgBox "—Ã«¡ Ê’› «·⁄ÿ· ", vbCritical: Exit Sub
    End If
'    Adodc1.Recordset.Fields!opr_date = DateValue(Now)
   
Adodc1.Recordset.Update


Dim x As Integer
 If my_language = "E" Then
x = MsgBox("confirm convert to work order", vbCritical + vbYesNo)

Else
x = MsgBox(" √ﬂÌœ «· ÕÊÌ· «·Ï «„— ‘€·", vbCritical + vbYesNo)
End If

If x = vbNo Then
Exit Sub
End If


 
sand_numbering
'Adodc1.Recordset.Fields!sanad_no = Adodc1.Recordset.Fields!sandat_pc_no
        If auto_sanad_no <> "" Then
           Adodc1.Recordset.Fields!Amr_shogl_no = auto_sanad_no
          Else
                         If my_language = "E" Then
                                  MsgBox "can't save because numbering type not defined in system manger", vbCritical: Exit Sub
                         Else
                               If my_language = "E" Then
                               MsgBox "work order can't save please define numbering method in system manger screen ", vbCritical: Exit Sub
                                Else
                                MsgBox "·„ Ì „ Õ›Ÿ «„— «·‘€· ·«‰ﬂ ·„ ‰Õœœ ‰Ê⁄ «· —ﬁÌ„ ›Ì „œÌ— «·‰Ÿ«„", vbCritical: Exit Sub
                                End If
                        End If
        End If
        
        
  Adodc1.Recordset.Fields!numbering_type = numbering_type
  If numbering_type = 2 Then
Adodc1.Recordset.Fields!sanad_year = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
Adodc1.Recordset.Fields!sanad_month = Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
End If

If numbering_type = 3 Then
Adodc1.Recordset.Fields!sanad_year = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)

End If


If Adodc1.Recordset.RecordCount > 0 Then
    Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields!amr_shogl_fk = Text1.Text
Adodc8.Recordset.Update


Adodc1.Recordset.Fields!cars_NO = DataCombo1.Text
Adodc1.Recordset.Fields!driver_name = DataCombo2.Text
Adodc1.Recordset.Fields!maintenance_type = DataCombo3.Text
Adodc1.Recordset.Fields!warsha_name = DataCombo4.Text
Adodc1.Recordset.Fields!error_person = DataCombo5.Text


Adodc1.Recordset.Fields!Amr_shogl = vbTrue
Adodc1.Recordset.Fields!converted = vbTrue
Adodc1.Recordset.Fields!branch_no = branch_no
 'Adodc1.Recordset.Fields!branch_no = branch_no
 
   Adodc1.Recordset.Fields!DEPARTEMENT = departement_name

  Adodc1.Recordset.Fields!user_name = current_user_name
Adodc1.Recordset.Update

'Adodc1.ConnectionString = connection_string
'Adodc1.CommandType = adCmdText
'Adodc1.RecordSource = "select * from maintenance  where   Amr_shogl=0"
Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from maintenance  where converted=0 and  branch_no=" & branch_no & " and departement='" & departement_name & "' and     Amr_shogl=0"
Adodc1.Refresh
 
'Adodc1.Recordset.Update
'DoEvents
'Adodc1.ConnectionString = connection_string
'Adodc1.CommandType = adCmdText
'Adodc1.RecordSource = "select * from maintenance  where  NOT (car_no IS NULL) and converted=0"
''Adodc1.Refresh

'Adodc1.Refresh
End If
     
     If my_language = "E" Then
         MsgBox "Convert to work order done", vbInformation

    Else
    MsgBox " „ «· ÕÊÌ·", vbInformation
    End If

End Sub

Private Sub Text1_change()
On Error Resume Next

'On Error Resume Next
'If Text1.Text = "" Or Text1.Text = "text1" Then Exit Sub
'If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
'If Not IsNull(Adodc1.Recordset.Fields!OUT_WARSHA) Then
'If (Adodc1.Recordset.Fields!OUT_WARSHA) = vbTrue Then
'Label14.Visible = False
'Else
'Label14.Visible = True
'End If
'
'End If
End Sub
