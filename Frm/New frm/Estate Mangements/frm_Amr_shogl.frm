VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form frm_Amr_shogl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«Ê«„— «·‘€·"
   ClientHeight    =   5295
   ClientLeft      =   2055
   ClientTop       =   -15
   ClientWidth     =   12360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   12360
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame16 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5400
      TabIndex        =   133
      Top             =   0
      Width           =   2415
      Begin VB.Label Label80 
         Caption         =   "«Ê«„— «·‘€·"
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
         Height          =   495
         Left            =   360
         TabIndex        =   138
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   735
      Left            =   7320
      Picture         =   "frm_Amr_shogl.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   105
      ToolTipText     =   "»ÕÀ ⁄‰ «„— ‘€·F3"
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   2880
      TabIndex        =   128
      Top             =   4560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   132
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   131
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2520
         TabIndex        =   130
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   129
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   1920
      TabIndex        =   123
      Top             =   8880
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   127
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   126
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   125
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   124
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   " "
      Height          =   1335
      Left            =   0
      TabIndex        =   119
      Top             =   7080
      Width           =   855
      Begin VB.Label Label32 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   120
         TabIndex        =   121
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "«÷«›…"
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
         Height          =   480
         Left            =   120
         TabIndex        =   120
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   960
      TabIndex        =   96
      Top             =   6960
      Width           =   8055
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÊÕœ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3480
         TabIndex        =   104
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·’‰›"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5520
         TabIndex        =   103
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ﬁÿ⁄…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   4320
         TabIndex        =   102
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·’‰›"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6840
         TabIndex        =   101
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·›‰Ì"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -120
         TabIndex        =   100
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Select Store"
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
         Left            =   10080
         TabIndex        =   99
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Œ“‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1080
         TabIndex        =   98
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂ„Ì…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2520
         TabIndex        =   97
         Top             =   120
         Width           =   615
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   2230
         X2              =   2230
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   1230
         X2              =   1230
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   6750
         X2              =   6750
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   4230
         X2              =   4230
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   7760
         X2              =   7760
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   3240
         X2              =   3240
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5250
         X2              =   5250
         Y1              =   720
         Y2              =   120
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   960
      TabIndex        =   110
      Top             =   6960
      Width           =   8055
      Begin VB.Line Line1 
         Index           =   20
         X1              =   4840
         X2              =   4840
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   19
         X1              =   2820
         X2              =   2820
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   18
         X1              =   6830
         X2              =   6830
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   3850
         X2              =   3850
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   5830
         X2              =   5830
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   320
         X2              =   320
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   1325
         X2              =   1325
         Y1              =   720
         Y2              =   120
      End
      Begin VB.Label Label71 
         BackStyle       =   0  'Transparent
         Caption         =   "Part No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2880
         TabIndex        =   118
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label70 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   117
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label69 
         Caption         =   "Select Store"
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
         Left            =   10080
         TabIndex        =   116
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   115
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label67 
         BackStyle       =   0  'Transparent
         Caption         =   "Technical"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6960
         TabIndex        =   114
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5040
         TabIndex        =   113
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventoty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5880
         TabIndex        =   112
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4200
         TabIndex        =   111
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   -1440
      TabIndex        =   106
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   107
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "»—ﬁ„ «·„⁄œÂ/«·”Ì«—…"
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
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_Amr_shogl.frx":1992
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
         TabIndex        =   108
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "»√”„ «·”«∆ﬁ"
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
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_Amr_shogl.frx":19AE
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
         Alignment       =   2  'Center
         Caption         =   "«·»ÕÀ"
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
         Left            =   240
         TabIndex        =   109
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   480
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Frame Frame13 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   9120
      TabIndex        =   91
      Top             =   6960
      Width           =   3135
      Begin VB.Frame Frame1 
         Height          =   1935
         Left            =   720
         TabIndex        =   92
         Top             =   -120
         Width           =   1935
         Begin ALLButtonS.ALLButton Command2 
            Height          =   375
            Left            =   120
            TabIndex        =   93
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "«·«‰ Â«¡ „‰ «·«’·«Õ"
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
            BCOL            =   65535
            BCOLO           =   65535
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_Amr_shogl.frx":19CA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton command3 
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   1440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Œ—ÊÃ „‰ «·Ê—‘…"
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
            BCOL            =   49152
            BCOLO           =   49152
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_Amr_shogl.frx":19E6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command100 
            Height          =   450
            Left            =   120
            TabIndex        =   136
            Top             =   650
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   794
            BTYPE           =   3
            TX              =   " «‰ Â«¡ «·«’·«Õ ›Ì «·ﬁ”„"
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
            MICON           =   "frm_Amr_shogl.frx":1A02
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
            TabIndex        =   95
            Top             =   1440
            Width           =   855
         End
      End
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   55
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
      MICON           =   "frm_Amr_shogl.frx":1A1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   1
      Left            =   5520
      TabIndex        =   86
      Top             =   240
      Width           =   1455
      Begin VB.Label Label42 
         Caption         =   " ﬂ·›… «·’Ì«‰…"
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
         TabIndex        =   140
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label800 
         Caption         =   "Êﬁ  «·œŒÊ·"
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
         TabIndex        =   134
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "«Ê«„— «·‘€·"
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
         Left            =   7080
         TabIndex        =   122
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label Label8 
         Caption         =   " «—ÌŒ  «·œŒÊ·"
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
         Index           =   1
         Left            =   120
         TabIndex        =   90
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "„ Õ„· «· ﬂ·›…"
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
         Index           =   1
         Left            =   120
         TabIndex        =   89
         Top             =   3120
         Width           =   1695
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
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   88
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " «—ÌŒ «·Œ—ÊÃ"
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
         Index           =   1
         Left            =   120
         TabIndex        =   87
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSDataListLib.DataCombo DataCombo11 
      Bindings        =   "frm_Amr_shogl.frx":1A3A
      Height          =   480
      Left            =   9480
      TabIndex        =   26
      Top             =   6240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   12632256
      ListField       =   "inventory_name"
      BoundColumn     =   "fullcode"
      Text            =   ""
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
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   0
      Left            =   10920
      TabIndex        =   30
      Top             =   120
      Width           =   1455
      Begin VB.Label Label1 
         Caption         =   " «—ÌŒ «·Õ—ﬂ…"
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
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
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
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "«·„«·ﬂ"
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
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label17 
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
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "—ﬁ„ «·«„—"
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
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   -600
      TabIndex        =   27
      Top             =   10080
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
         Caption         =   "M26"
         Height          =   255
         Left            =   3360
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox txt_last_price 
      Height          =   405
      Left            =   -480
      TabIndex        =   25
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      Height          =   405
      Left            =   6720
      TabIndex        =   24
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   11040
      Picture         =   "frm_Amr_shogl.frx":1A50
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "»ÕÀ ’‰›  F3"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtdepartement 
      Height          =   615
      Left            =   3000
      TabIndex        =   22
      Text            =   "Text12"
      Top             =   -360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Height          =   840
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2415
      Left            =   -4560
      TabIndex        =   11
      Top             =   -3000
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox Text7 
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
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text8 
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
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo DataCombo7 
         Bindings        =   "frm_Amr_shogl.frx":33E2
         Height          =   480
         Left            =   720
         TabIndex        =   14
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         BackColor       =   12632256
         ListField       =   "departement_name"
         Text            =   ""
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
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "«—”«·"
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
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "«—”«· «·Ï ﬁ”„"
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
         Left            =   3600
         TabIndex        =   17
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "«·ﬁ”„ «·Õ«·Ì"
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
         Left            =   3720
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "«·„” Œœ„ «·Õ«·Ì"
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
         Left            =   3480
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   9
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
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
      Left            =   6120
      TabIndex        =   2
      Top             =   5400
      Width           =   855
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2775
      Left            =   1560
      TabIndex        =   7
      Top             =   -3360
      Visible         =   0   'False
      Width           =   3615
      _Version        =   524288
      _ExtentX        =   6376
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   16711680
      Year            =   2009
      Month           =   6
      Day             =   4
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   16777215
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   16777215
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   16777215
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
         Charset         =   0
         Weight          =   700
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
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Bindings        =   "frm_Amr_shogl.frx":33F7
      Height          =   480
      Left            =   9480
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   12632256
      ListField       =   "items_name"
      BoundColumn     =   "fullcode"
      Text            =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_Amr_shogl.frx":340C
      Height          =   1695
      Left            =   960
      TabIndex        =   8
      Top             =   7245
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   23
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
      ColumnCount     =   22
      BeginProperty Column00 
         DataField       =   "transaction_id"
         Caption         =   "transaction_id"
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
         Caption         =   "—ﬁ„ «·’‰›"
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
         DataField       =   "item_unit"
         Caption         =   "«·ÊÕœ…"
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
         DataField       =   "qty"
         Caption         =   "«·ﬂ„Ì…"
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
         DataField       =   "items_no_by_one"
         Caption         =   "items_no_by_one"
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
         DataField       =   "ta2ther_makhzan"
         Caption         =   "ta2ther_makhzan"
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
         DataField       =   "average_cost"
         Caption         =   "average_cost"
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
         DataField       =   "last_price"
         Caption         =   "last_price"
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
         DataField       =   "total_price"
         Caption         =   "total_price"
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
         DataField       =   "transaction_type"
         Caption         =   "transaction_type"
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
      BeginProperty Column12 
         DataField       =   "transaction_date"
         Caption         =   "transaction_date"
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
      BeginProperty Column13 
         DataField       =   "sanad_no"
         Caption         =   "sanad_no"
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
      BeginProperty Column14 
         DataField       =   "sanad_type"
         Caption         =   "sanad_type"
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
      BeginProperty Column15 
         DataField       =   "bona_3la"
         Caption         =   "bona_3la"
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
      BeginProperty Column16 
         DataField       =   "inventory_id"
         Caption         =   "inventory_id"
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
      BeginProperty Column17 
         DataField       =   "inventory_name"
         Caption         =   "«·„Œ“‰"
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
      BeginProperty Column18 
         DataField       =   "amr_shogl_fk"
         Caption         =   "amr_shogl_fk"
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
      BeginProperty Column19 
         DataField       =   "technical"
         Caption         =   "«·›‰Ì"
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
      BeginProperty Column20 
         DataField       =   "technical_notes"
         Caption         =   "technical_notes"
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
      BeginProperty Column21 
         DataField       =   "item_departement"
         Caption         =   "item_departement"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo8 
      Bindings        =   "frm_Amr_shogl.frx":3421
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   3120
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   12632256
      ListField       =   "name"
      Text            =   "€Ì— „Õœœ"
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
   Begin MSDataListLib.DataCombo DataCombo9 
      Bindings        =   "frm_Amr_shogl.frx":3436
      Height          =   480
      Left            =   9120
      TabIndex        =   20
      Top             =   10560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   12632256
      ListField       =   "departement_name"
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DataCombo10 
      Bindings        =   "frm_Amr_shogl.frx":344B
      Height          =   480
      Left            =   8160
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   12632256
      ListField       =   "unit_name"
      BoundColumn     =   "unit_value"
      Text            =   ""
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   -240
      Top             =   9120
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
      Left            =   -120
      Top             =   9360
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
      Left            =   -120
      Top             =   9600
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
      Left            =   -120
      Top             =   9720
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   -120
      Top             =   10080
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   -120
      Top             =   8880
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   -120
      Top             =   9240
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   -120
      Top             =   9600
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   -120
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   -120
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   -120
      Top             =   10680
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
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   330
      Left            =   -240
      Top             =   4680
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
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   330
      Left            =   -480
      Top             =   8880
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
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   330
      Left            =   -840
      Top             =   8400
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
   Begin VB.Frame Frame12 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   5280
      TabIndex        =   57
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "Amr_shogl_no"
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   137
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "opr_date"
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   960
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   0
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         DataField       =   "moghat_belt2men"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   5280
         TabIndex        =   66
         Top             =   4080
         Width           =   255
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   360
         TabIndex        =   58
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
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
            TabIndex        =   65
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label56 
            Caption         =   "Equipmet#"
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
            TabIndex        =   64
            Top             =   1440
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
            TabIndex        =   63
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label58 
            Caption         =   "Work order#"
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
            TabIndex        =   62
            Top             =   360
            Width           =   1575
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
            Height          =   615
            Left            =   120
            TabIndex        =   61
            Top             =   3240
            Width           =   1335
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
            Height          =   495
            Left            =   120
            TabIndex        =   60
            Top             =   3960
            Width           =   2415
         End
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
            TabIndex        =   59
            Top             =   2640
            Width           =   2415
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_Amr_shogl.frx":3461
         DataField       =   "car_no"
         DataSource      =   "Adodc1"
         Height          =   480
         Left            =   2880
         TabIndex        =   67
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   12632256
         ListField       =   "Car_no"
         Text            =   ""
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
         Bindings        =   "frm_Amr_shogl.frx":3476
         DataField       =   "driver_name"
         DataSource      =   "Adodc1"
         Height          =   480
         Left            =   2880
         TabIndex        =   70
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   12632256
         ListField       =   "driver_name"
         Text            =   ""
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
         Bindings        =   "frm_Amr_shogl.frx":348B
         DataField       =   "maintenance_type"
         DataSource      =   "Adodc1"
         Height          =   480
         Left            =   2880
         TabIndex        =   71
         Top             =   2760
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   12632256
         ListField       =   "maintenance_type"
         Text            =   ""
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
         Bindings        =   "frm_Amr_shogl.frx":34A0
         DataField       =   "warsha_name"
         DataSource      =   "Adodc1"
         Height          =   480
         Left            =   2880
         TabIndex        =   72
         Top             =   3360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   12632256
         ListField       =   "wersha_name"
         Text            =   ""
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
   Begin VB.Frame Frame8 
      Height          =   2295
      Left            =   0
      TabIndex        =   46
      Top             =   6120
      Visible         =   0   'False
      Width           =   12255
      Begin VB.Label Label41 
         Caption         =   "Unit"
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
         Left            =   3840
         TabIndex        =   54
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label40 
         Caption         =   "Technical Notes"
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
         Left            =   -240
         TabIndex        =   53
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label39 
         Caption         =   "Total"
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
         Left            =   6960
         TabIndex        =   52
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label38 
         Caption         =   "Cost"
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
         Left            =   5160
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label37 
         Caption         =   "Qty"
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
         Left            =   6120
         TabIndex        =   50
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label36 
         Caption         =   "Technical name"
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
         Left            =   8520
         TabIndex        =   49
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "Select Item"
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
         Left            =   1680
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label34 
         Caption         =   "Select Store"
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
         Left            =   1320
         TabIndex        =   47
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame Frame7 
      Height          =   2295
      Left            =   0
      TabIndex        =   38
      Top             =   6360
      Width           =   11415
      Begin VB.Label Label14 
         Caption         =   "«Œ «— «·’‰›"
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
         Left            =   9600
         TabIndex        =   56
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label31 
         Caption         =   "«Œ «— «·„Œ“‰"
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
         Left            =   9720
         TabIndex        =   45
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label26 
         Caption         =   "«·›‰Ì «·ﬁ«∆„ »«· —ﬂÌ»"
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
         Left            =   3120
         TabIndex        =   44
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "«·ﬂ„Ì…"
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
         Left            =   6240
         TabIndex        =   43
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "«· ﬂ·›…"
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
         Left            =   7200
         TabIndex        =   42
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   " ⁄·Ìﬁ ›‰Ì «· —ﬂÌ»"
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
         Left            =   600
         TabIndex        =   40
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label30 
         Caption         =   "«·ÊÕœ…"
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
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "«·«Ã„«·Ì"
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
         Left            =   5040
         TabIndex        =   41
         Top             =   120
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc16 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc17 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   -360
      TabIndex        =   73
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "Time_in"
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
         TabIndex        =   141
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   360
         TabIndex        =   74
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
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
            Height          =   375
            Left            =   0
            TabIndex        =   79
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label450 
            Caption         =   "Time In"
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
            TabIndex        =   78
            Top             =   1080
            Width           =   2415
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
            TabIndex        =   77
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
            Left            =   0
            TabIndex        =   76
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
            TabIndex        =   75
            Top             =   3960
            Width           =   2415
         End
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "Time_in"
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
         TabIndex        =   135
         Top             =   1080
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
         Left            =   -240
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   2280
         Visible         =   0   'False
         Width           =   2775
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   600
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
         Height          =   1080
         Left            =   2760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   83
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0C0C0&
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
         ItemData        =   "frm_Amr_shogl.frx":34B5
         Left            =   2760
         List            =   "frm_Amr_shogl.frx":34BF
         TabIndex        =   82
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   1
         Left            =   2280
         Picture         =   "frm_Amr_shogl.frx":34D5
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1680
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.CommandButton Command40 
         Height          =   492
         Index           =   0
         Left            =   2280
         Picture         =   "frm_Amr_shogl.frx":3D37
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   720
         Visible         =   0   'False
         Width           =   492
      End
   End
   Begin ALLButtonS.ALLButton Command1 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   139
      Top             =   4680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÿ»«⁄Â «„— «·‘€· "
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Amr_shogl.frx":4599
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "«—”«· «·«„— «·Ï ﬁ”„"
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
      Left            =   11280
      TabIndex        =   21
      Top             =   9480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "«—”«·"
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
      Height          =   840
      Left            =   7440
      TabIndex        =   19
      Top             =   10560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "«Ê«„— «·‘€·"
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
      Left            =   6600
      TabIndex        =   5
      Top             =   -600
      Width           =   2055
   End
End
Attribute VB_Name = "frm_Amr_shogl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Dim bindex As Integer
Dim start_load As Boolean

 

Private Sub Calendar1_Click()
On Error Resume Next
If bindex = 0 Then
Text3.Text = Calendar1.Value
End If

If bindex = 1 Then
    Text5.Text = Calendar1.Value
End If
Calendar1.Visible = False
End Sub

Private Sub Check1_Click()
On Error Resume Next
'If Check1.value = 0 Then
'Combo1.Visible = True
'Text4.Visible = True
 'Adodc1.Recordset.Fields!moghat_belt2men = vbFalse
'Else
'Combo1.Visible = False
' Text4.Visible = False
 ' Adodc1.Recordset.Fields!moghat_belt2men = vbTrue
' Text4.Text = 0
'End If
'Adodc1.Recordset.Update
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

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Combo1.Text = ""
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Select Case Index

Case 0
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!opr_date = DateValue(Now)
     Adodc1.Recordset.Fields!moghat_belt2men = vbTrue
     
Case 1
    Adodc1.Recordset.Update

Case 2
    If my_language = "E" Then
    x = MsgBox("CONFIRM DELETE", vbCritical + vbYesNo)
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

Case 3
    If Adodc1.Recordset.RecordCount > 0 Then
    
    Adodc17.RecordSource = "select * from maintenance_all_details_qry where opr_id=" & Text1.Text
Adodc17.Refresh

If Adodc17.Recordset.RecordCount > 1 Then
    Form3.sqllbl = "select * from maintenance_all_details_qry where not (item_code is null) and opr_id=" & Text1.Text

Else
    Form3.sqllbl = "select * from maintenance_all_details_qry where opr_id=" & Text1.Text

End If

    Form3.case_id = 4
   
    Form3.Show
    Else
            If my_language = "E" Then
            MsgBox "NO Job order", vbCritical
            Else
            MsgBox "·« ÌÊÃœ «„— ‘€· ·ÿ»«⁄ …", vbCritical
            End If
           
           
    End If

Case 4
On Error Resume Next
        If my_language = "E" Then
        x = InputBox("Enter Car#")
        Else
        x = InputBox("«œŒ· «·—ﬁ„ «·„ÿ·Ê» «·»ÕÀ ⁄‰…")
        End If
        If IsNumeric(x) Then
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from  maintenance where car_no='" & x & "'"
        Adodc1.Refresh
        Else
                If my_language = "E" Then
                MsgBox "Must enter Digit only", vbCritical
                Else
                MsgBox "·«»œ „‰ ﬂ «»… «—ﬁ«„", vbCritical
                End If
       
        End If

Case 5
If my_language = "E" Then
 x = InputBox("enter Driver Name")
Else
    x = InputBox("«œŒ· ﬂ·„… «·»ÕÀ")
End If
            Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from maintenance where driver_name like '%" & x & "%'"
        Adodc1.Refresh


End Select

End Sub

Private Sub Command100_Click()
On Error Resume Next

If Adodc8.Recordset.RecordCount = 0 Then
If my_language = "E" Then
x = MsgBox("this departement does not attach any parts to this car", vbCritical)

Else

x = MsgBox("Â–« «·ﬁ”„ ·„ Ì÷› «Ì ﬁÿ⁄ €Ì«— ·Â–… «·„⁄œÂ/«·”Ì«—…", vbCritical)
End If

Exit Sub
End If
Adodc8.Recordset.MoveFirst
For i = 1 To Adodc8.Recordset.RecordCount

Adodc8.Recordset.Fields!finish_date = Now
Adodc8.Recordset.Update
Adodc8.Recordset.MoveNext
Next i


End Sub

Private Sub Command13_Click()
On Error Resume Next
DataCombo6.Text = ""
items_search2.Show
items_search2.case_id = 1
End Sub

Private Sub Command2_Click()
On Error Resume Next
If my_language = "E" Then
x = MsgBox("confirm maintenance finish", vbCritical + vbYesNo)

Else

x = MsgBox(" √ﬂÌœ ⁄„·Ì… «·«‰Â«¡", vbCritical + vbYesNo)
End If

If x = vbNo Then
Exit Sub
End If

    If my_language = "E" Then
        If Adodc1.Recordset.RecordCount = 0 Then MsgBox "No Job order found", vbCritical: Exit Sub
        Else
        If Adodc1.Recordset.RecordCount = 0 Then MsgBox "·« ÌÊÃœ «„— ‘€· ·· ⁄«„· „⁄…", vbCritical: Exit Sub
    End If
Adodc1.Recordset.Fields!repaired = 1
Adodc1.Recordset.Fields!end_maintenance_date = DateValue(Now)
Adodc1.Recordset.Fields!end_maintenance_TIME = time
Adodc1.Recordset.Fields!Amr_shogl = vbFalse
Adodc1.Recordset.Fields!no_of_maintenance_days = DateDiff("D", Text3.Text, Now)
Adodc1.Recordset.Update
'Adodc1.RecordSource = "select * from maintenance where Amr_shogl=1  "
'Adodc1.Refresh
'Adodc1.Refresh
If my_language = "E" Then
MsgBox "maintenance finish done    ", vbInformation

Else
MsgBox "   „ «‰ Â«¡ «·«’·«Õ", vbInformation
End If
UPDATE_RECORDS
End Sub

Private Sub Command3_Click()
On Error Resume Next
If my_language = "E" Then
x = MsgBox("confirm car out from workshop", vbCritical + vbYesNo)

Else
x = MsgBox(" √ﬂÌœ ⁄„·Ì… «·Œ—ÊÃ ··”Ì«—… „‰ «·Ê—‘…", vbCritical + vbYesNo)
End If

If x = vbNo Then
Exit Sub
End If

    If my_language = "E" Then
    If Adodc1.Recordset.RecordCount = o Then MsgBox "No Job Order to Clost it", vbCritical: Exit Sub
    Else
    If Adodc1.Recordset.RecordCount = o Then MsgBox "·« ÌÊÃœ «„— ‘€· ·«€·«ﬁ…", vbCritical: Exit Sub
    End If
Adodc1.Recordset.Fields!repaired = 1
Adodc1.Recordset.Fields!date_out_warsha = DateValue(Now)
 Adodc1.Recordset.Fields!TIME_OUT = time

Adodc1.Recordset.Fields!Amr_shogl = vbFalse

Adodc1.Recordset.Update
Adodc1.RecordSource = "select * from maintenance where Amr_shogl=1  "
Adodc1.Refresh
Adodc1.Refresh
If my_language = "E" Then
MsgBox " car out done", vbInformation
Else
MsgBox "   „ «·Œ—ÊÃ", vbInformation

End If
UPDATE_RECORDS

End Sub

Private Sub Command4_Click()
 On Error Resume Next

workorder_search.Show
 workorder_search.case_id = 0
 
End Sub

Private Sub Command40_Click(Index As Integer)
On Error Resume Next
bindex = Index
Calendar1.Visible = True
Calendar1.Value = DateValue(Now)
Calendar1.Top = Command40(Index).Top
Calendar1.Left = Command40(Index).Left
End Sub

Function UPDATE_RECORDS()
 On Error Resume Next
frm_Amr_shogl.Adodc1.CommandType = adCmdText
frm_Amr_shogl.Adodc1.RecordSource = "select * from maintenance where Amr_shogl=1" ' and  departement_name_now = '" & Me.txtdepartement.Text & "'"
frm_Amr_shogl.Adodc1.Refresh
 
frm_Amr_shogl.Adodc8.CommandType = adCmdText
frm_Amr_shogl.Adodc8.RecordSource = "select * from inventory where  branch_no=" & branch_no & " and amr_shogl_fk=" & Val(Me.Text1.Text) & " and item_departement='" & Me.txtdepartement.Text & "'"
frm_Amr_shogl.Adodc8.Refresh
frm_Amr_shogl.DataGrid1.Refresh
End Function

  

Private Sub DataCombo1_DblClick(Area As Integer)
On Error Resume Next

Car_info_bar.Show
Car_info_bar.Car_no = DataCombo1.Text
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

 

If KeyCode = vbKeyF6 Then
frmcars.Show
End If


End Sub

Private Sub DataCombo10_Click(Area As Integer)


On Error Resume Next
If DataCombo6.Text = "" Or DataCombo10.Text = "" Then Exit Sub
Adodc12.CommandType = adCmdText
Adodc12.RecordSource = "select * from  items where  branch_no=" & branch_no & " and departement='" & departement_name & "' and   blocked=0 AND  fullcode='" & DataCombo6.BoundText & "'"
Adodc12.Refresh
         
         
If Adodc12.Recordset.RecordCount = 0 Then Exit Sub
Adodc12.Recordset.MoveLast
Text9.Text = Val(Adodc12.Recordset.Fields!motwaset_taklefa) * Val(DataCombo10.BoundText)
 txt_last_price.Text = Adodc12.Recordset.Fields!akher_s3r_shera





End Sub

Private Sub DataCombo10_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then
Label19_Click
Else
DataCombo10.Text = ""
End If

If KeyCode = vbKeyF6 Then
items_units.Show
End If


End Sub

Private Sub DataCombo11_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then
Label19_Click
Else
DataCombo11.Text = ""
End If









On Error Resume Next
 
If KeyCode = vbKeyF3 Then
inventory_search.Show
inventory_search.case_id = 4
End If

 


If KeyCode = vbKeyF6 Then
inventory.Show
End If

If KeyCode = vbKeyF5 Then
Adodc15.Refresh
DataCombo11.ReFill
End If

End Sub

 

Private Sub DataCombo2_Click(Area As Integer)
On Error Resume Next

If KeyCode = vbKeyF3 Then
Driver_search.Show
Driver_search.case_id = 3
End If

If KeyCode = vbKeyF6 Then
drivers.Show
End If

End Sub

Private Sub DataCombo3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF6 Then
frmmaintenace_type.Show
End If
End Sub

Private Sub DataCombo4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF6 Then
frmwarsha.Show
End If
End Sub

Private Sub DataCombo5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF6 Then
drivers.Show
End If
End Sub

Private Sub DataCombo6_Change()
On Error Resume Next
 
DataCombo10.Text = ""
Text9.Text = ""
Text6.Text = ""
Text10.Text = ""
Text11.Text = ""
DataCombo8.Text = ""

If DataCombo6.Text = "" Then Exit Sub
Adodc11.CommandType = adCmdText
Adodc11.RecordSource = "select * from  items_units where item_code='" & DataCombo6.BoundText & "'"
Adodc11.Refresh

DataCombo10.ReFill

 
Text6.Text = 1
End Sub

Private Sub DataCombo6_Click(Area As Integer)
On Error Resume Next
If DataCombo6.Text = "" Then Exit Sub
Adodc11.CommandType = adCmdText
Adodc11.RecordSource = "select * from  items_units where  item_code='" & DataCombo6.BoundText & "'"
Adodc11.Refresh

DataCombo10.ReFill

 
Text6.Text = 1


End Sub

Private Sub DataCombo6_DblClick(Area As Integer)
On Error Resume Next

items_info_bar.Show
items_info_bar.item_code = DataCombo6.BoundText
items_info_bar.inventory_id = DataCombo11.Text
 
 items_info_bar.item_name = DataCombo6.Text
items_info_bar.inventory_name = DataCombo11.Text

End Sub

Private Sub DataCombo6_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF3 Then
items_search2.Show
items_search2.case_id = 1
DataCombo6.Text = ""
End If

If KeyCode = 13 Then
Label19_Click
End If


If KeyCode = vbKeyF6 Then
frmitems.Show
End If


If KeyCode = vbKeyF5 Then
Adodc6.Refresh
DataCombo6.ReFill
End If



End Sub

 

Private Sub DataCombo6_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

If DataCombo6.Text = "" Then Exit Sub
'If items_info_bar.Visible = True Then
'Unload items_info_bar
'End If

End Sub

Private Sub DataCombo8_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then
Label19_Click
End If
If KeyCode = vbKeyF6 Then
EMPLOYEES.Show
End If

If KeyCode = vbKeyF5 Then
Adodc9.Refresh
DataCombo8.ReFill
End If

If KeyCode = vbKeyF3 Then
emp_search.Show
emp_search.case_id = 4
End If

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 46 Then
Label32_Click
End If

If KeyCode = 13 Then
If Adodc8.Recordset.RecordCount > 0 Then
Adodc8.Recordset.Fields!total_price = Adodc8.Recordset.Fields!qty * Adodc8.Recordset.Fields!average_cost
Adodc8.Recordset.Update
End If

End If

End Sub

Private Sub Form_Activate()
On Error Resume Next



If start_load = False Then

Adodc16.ConnectionString = connection_string
 Adodc16.CommandType = adCmdText
Adodc16.RecordSource = "select * from departement where departement_no=" & current_user_departement
Adodc16.Refresh
 
Me.txtdepartement.Text = Adodc16.Recordset.Fields!departement_name



frm_Amr_shogl.Adodc8.CommandType = adCmdText
frm_Amr_shogl.Adodc8.RecordSource = "select * from inventory where   amr_shogl_fk=" & Val(Me.Text1.Text) & " and item_departement='" & Me.txtdepartement.Text & "'"
frm_Amr_shogl.Adodc8.Refresh
frm_Amr_shogl.DataGrid1.Refresh
start_load = True

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
Command100.Caption = "Maintenence Finish in Dept"
CMD_language.ToolTipText = "Change Language"
Command13.ToolTipText = "F3 JobOrders Search "
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

 Me.Left = (MDIForm1.Width - Me.Width) / 2
    Me.Top = (MDIForm1.Height - Me.Height) / 2 - 500
    
 Dim tleft As Integer
 
 If my_language = "E" Then
 Frame15.Left = 9120
 Frame14.Visible = False
Command13.Left = 1200 - 1000
 
Label35.Left = 1000
Label41.Left = 3840 + 220 - 1000
Label38.Left = 5160 + 150 - 1000
Label37.Left = 6120 + 150 - 1000
Label39.Left = 6960 + 150 - 1000
Label36.Left = 7800 + 250 - 1000
Label40.Left = 10440 - 1000
Label34.Left = 1000
DataCombo8.Left = 9000 - 1700
DataCombo6.Left = 2040 - 1000
DataCombo11.Left = 2040 - 1000

DataCombo10.Left = 4680 - 1700
Text9.Left = 6000 - 1700
Text6.Left = 6960 - 1700
Text10.Left = 7920 - 1700
Text11.Left = 11040 - 1700


  tleft = Frame12.Left
Frame12.Left = Frame6.Left
Frame6.Left = tleft
' Label19.Left = 13800
 ' Label32.Left = 13800

 'Frame13.Left = 12480
 
  Text11.Alignment = 0
   Text10.Alignment = 0
   Text14.Alignment = 0
    
     
 DataCombo6.RightToLeft = False
 DataCombo11.RightToLeft = False
 DataCombo10.RightToLeft = False
 DataCombo8.RightToLeft = False
 
  Text1.Alignment = 0
  txtid.Alignment = 0
   Text2.Alignment = 0
    Text3.Alignment = 0
     Text4.Alignment = 0
      Text5.Alignment = 0
      Text13.Alignment = 0
      
     Check1.Left = 2760
  
  
  DataCombo1.RightToLeft = False
    DataCombo2.RightToLeft = False
      DataCombo3.RightToLeft = False
        DataCombo4.RightToLeft = False
          DataCombo5.RightToLeft = False
             Combo1.RightToLeft = False
        Combo1.Clear
        Combo1.AddItem "Driver"
        Combo1.AddItem "Company"
 CMD_language.Caption = "⁄—»Ì"
 DataGrid1.RightToLeft = False
Frame8.Visible = True
Frame9.Visible = True
Frame10.Visible = True
Frame11.Visible = True
Adodc1.Caption = "move"

Frame5(0).Visible = False
Frame5(1).Visible = False

'Frame6.Visible = False
'Frame7.Visible = False

Label80.Caption = "Job Orders"
Me.Caption = Label80.Caption

Label19.Caption = "Add"
Label32.Caption = "Del"
Label7.Caption = "search"
Command1(4).Caption = "by car"
Command1(5).Caption = "by driver"

Command1(3).Caption = "view final Job order"
Command2.Caption = "Finish maintenance"
Command3.Caption = "out of workshop"

Label29.Caption = "Send to Dep"
Label28.Caption = "send"
End If

'On Error Resume Next
LoadSettings
Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from maintenance where  branch_no=" & branch_no & " and departement='" & departement_name & "' and Amr_shogl=1  "
Adodc1.Refresh
DoEvents
If Adodc1.Recordset.RecordCount > 0 Then
Adodc1.Recordset.MoveLast
If my_language = "E" Then

MsgBox "You Have:  " & Adodc1.Recordset.RecordCount & "    Job Ordesr"
Else
MsgBox "  ·œÌﬂ   " & Adodc1.Recordset.RecordCount & "    «„— ‘€·"
End If

 Else
 
 If my_language = "E" Then

MsgBox " NO Job Orders"
Else
MsgBox "·« ÌÊÃœ «Ê«„— ‘€·"
End If


End If

'Adodc1.RecordSource = "select * from maintenance where opr_id=0 "
'Adodc1.Refresh

'Adodc16.ConnectionString = connection_string
' Adodc16.CommandType = adCmdText
'Adodc16.RecordSource = "select * from departement where departement_no=" & current_user_departement
'Adodc16.Refresh
 
Me.txtdepartement.Text = departement_name ' Adodc16.Recordset.Fields!departement_name


'Adodc2.ConnectionString = connection_string
'Adodc2.CommandType = adCmdText
'Adodc2.RecordSource = "select  * from CARS  "
'Adodc2.Refresh

'Adodc3.ConnectionString = connection_string
'Adodc3.CommandType = adCmdText
'Adodc3.RecordSource = "select  * from drivers  "
'Adodc3.Refresh

'Adodc4.ConnectionString = connection_string
'Adodc4.CommandType = adCmdText
'Adodc4.RecordSource = "select  * from maintenance_type  "
'Adodc4.Refresh

'Adodc5.ConnectionString = connection_string
'Adodc5.CommandType = adCmdText
'Adodc5.RecordSource = "select  * from wersha  "
'Adodc5.Refresh

'Adodc6.ConnectionString = connection_string
'Adodc6.CommandType = adCmdText
'Adodc6.RecordSource = "select distinct items_name,item_code from  inventory  where not (item_code is null) "
'Adodc6.Refresh


'Adodc7.ConnectionString = connection_string
'Adodc7.CommandType = adCmdText
'Adodc7.RecordSource = "select * from departement where warsha=1  "
'Adodc7.Refresh

Adodc8.ConnectionString = connection_string
Adodc8.CommandType = adCmdText
Adodc8.RecordSource = "SELECT * FROM inventory WHERE branch_no=" & branch_no & " and  amr_shogl_fk=0"
Adodc8.Refresh
DataGrid1.Refresh

Adodc9.ConnectionString = connection_string
Adodc9.CommandType = adCmdText
Adodc9.RecordSource = "select * from employees where  branch_no=" & branch_no & "and departement='" & departement_name & "'"
Adodc9.Refresh


'Adodc10.ConnectionString = connection_string
'Adodc10.CommandType = adCmdText
'Adodc10.RecordSource = "select  * from technicals  "
'Adodc10.Refresh


Adodc11.ConnectionString = connection_string
Adodc11.CommandType = adCmdText
Adodc11.RecordSource = "select * from  items_units where item_id=0"
Adodc11.Refresh

Adodc12.ConnectionString = connection_string
Adodc12.CommandType = adCmdText
Adodc12.RecordSource = "select * from  items_units where item_id=0"
Adodc12.Refresh

Adodc13.ConnectionString = connection_string
Adodc13.CommandType = adCmdText
Adodc13.RecordSource = "select  * from items WHERE branch_no=" & branch_no & " and   blocked=0   "
Adodc13.Refresh

Adodc14.ConnectionString = connection_string
Adodc14.CommandType = adCmdText
Adodc14.RecordSource = "  select * from inventory  where  branch_no=" & branch_no
Adodc14.Refresh

Adodc15.ConnectionString = connection_string
Adodc15.CommandType = adCmdText
Adodc15.RecordSource = "select  * from inventories where  branch_no=" & branch_no & " and departement='" & departement_name & "' and not(inventory_name='')  "
Adodc15.Refresh


Adodc17.ConnectionString = connection_string
Adodc17.CommandType = adCmdText







'frm_Amr_shogl.Adodc1.CommandType = adCmdText
'frm_Amr_shogl.Adodc1.RecordSource = "select * from maintenance where Amr_shogl=1 and  departement_name_now = '" & Me.txtdepartement.Text & "'"
'frm_Amr_shogl.Adodc1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
start_load = False
End Sub

Private Sub Label19_Click()
On Error Resume Next
'If Not IsNumeric(DataCombo10.BoundText) Then MsgBox "SELECT UNIT «Œ «— ÊÕœ…", vbCritical: Exit Sub
'If Not IsNumeric(DataCombo11.BoundText) Then MsgBox "SELECT INVENTORY «Œ «— „Œ“‰", vbCritical: Exit Sub

 If my_language = "E" Then
     If Text1.Text = "" Then MsgBox "opr# not found error", vbCritical: Exit Sub
    
    
    If DataCombo6.Text = "" Then MsgBox "must select item name", vbCritical: Exit Sub
    
    If DataCombo10.Text = "" Then MsgBox "select item unit", vbCritical: Exit Sub
    If Text6.Text = "" Then MsgBox "specify Qty", vbCritical: Exit Sub
    If Text9.Text = "" Then MsgBox "error in cost item have error", vbCritical: Exit Sub
    
    
    If Text10.Text = "" Then MsgBox "error in total", vbCritical: Exit Sub
    If DataCombo6.Text = "" Then MsgBox "", vbCritical: Exit Sub
    
    
    If DataCombo8.Text = "" Then MsgBox "specify technical name", vbCritical: Exit Sub
    If Text11.Text = "" Then MsgBox "specify technical notes", vbCritical: Exit Sub
    If DataCombo11.Text = "" Then
    MsgBox "select store«", vbCritical
    DataCombo11.SetFocus
    SendKeys "{f4}"
    Exit Sub
    End If

 Else
    If Text1.Text = "" Then MsgBox "·« ÌÊÃœ —ﬁ„ Õ—ﬂ… ·«„— «·‘€·", vbCritical: Exit Sub
    
    
    If DataCombo6.Text = "" Then MsgBox "·«»œ „‰  ”ÃÌ· «”„ «·’‰›", vbCritical: Exit Sub
    
    If DataCombo10.Text = "" Then MsgBox "«Œ — ÊÕœ… «·’‰› «Ê·«", vbCritical: Exit Sub
    If Text6.Text = "" Then MsgBox "Õœœ «·ﬂ„Ì… «·„—«œ ’—›Â«", vbCritical: Exit Sub
    If Text9.Text = "" Then MsgBox "«· ﬂ·›… €Ì— „Õœœ… ·„   „ «Ì ⁄„·Ì«  ⁄·Ï Â–« «·’‰›", vbCritical: Exit Sub
    
    
    If Text10.Text = "" Then MsgBox "Œÿ√ ›Ì «·«Ã„«·Ì", vbCritical: Exit Sub
    If DataCombo6.Text = "" Then MsgBox "", vbCritical: Exit Sub
    
    
    If DataCombo8.Text = "" Then MsgBox "·«»œ „‰  ”ÃÌ· «”„ «·›‰Ì", vbCritical: Exit Sub
    If Text11.Text = "" Then MsgBox "·«»œ „‰  ”ÃÌ·  ⁄·Ìﬁ  «·›‰Ì", vbCritical: Exit Sub
    
        If DataCombo11.Text = "" Then
   MsgBox "«Œ — «·„Œ“‰ «Ê·«", vbCritical
    DataCombo11.SetFocus
    SendKeys "{f4}"
    Exit Sub
    End If
    
    
 

End If

Adodc14.RecordSource = "select SUM(ta2ther_makhzan) AS AVILABLE_ITEMS from  inventory WHERE branch_no=" & branch_no & " and  item_code='" & DataCombo6.BoundText & "' AND inventory_id='" & DataCombo11.BoundText & "'"
Adodc14.Refresh

If IsNull(Adodc14.Recordset.Fields!AVILABLE_ITEMS) Then
    If my_language = "E" Then
    x = MsgBox("this item does not found in this store continue any case", vbYesNo + vbCritical)

    Else
    x = MsgBox("Â–« «·’‰› €Ì— „ÊÃÊœ ›Ì «·„Œ“‰ «·„Õœœ Â·  —Ìœ  ﬂ„·… ⁄„·Ì… «·’—› ⁄·Ï «Ì… Õ«·", vbYesNo + vbCritical)
    End If

If x = vbNo Then
Exit Sub
End If

End If

If Adodc14.Recordset.Fields!AVILABLE_ITEMS < Val(Val(Text6.Text) * Val(DataCombo10.BoundText)) Then

    If my_language = "E" Then
    x = MsgBox("ITEM QTY IN THIS ORDER GREATER TAHN > item qty in this store continue in case", vbYesNo + vbCritical)
    Else
    x = MsgBox("⁄œœ «·«’‰«› «·„ÊÃÊœ… «ﬁ· „‰ «·ﬂ„Ì… «·„ÿ·Ê» ’—›Â« „‰ «·„Œ“‰ Â·  —Ìœ  ﬂ„·… «·⁄„·Ì… ⁄·Ï «Ì Õ«·", vbYesNo + vbCritical)
    
    End If

If x = vbNo Then
Exit Sub
End If

End If





Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields!branch_no = branch_no
 Adodc8.Recordset.Fields!user_name = current_user_name
  
Adodc8.Recordset.Fields!amr_shogl_fk = Text1.Text

Adodc8.Recordset.Fields!item_code = DataCombo6.BoundText

Adodc13.CommandType = adCmdText
Adodc13.RecordSource = "select * from  items where branch_no=" & branch_no & " and   blocked=0 AND   fullcode='" & DataCombo6.BoundText & "'"
Adodc13.Refresh
Adodc8.Recordset.Fields!part_no = Adodc13.Recordset.Fields!part_no
Adodc8.Recordset.Fields!items_name = Adodc13.Recordset.Fields!items_name

Adodc8.Recordset.Fields!item_unit = DataCombo10.Text
Adodc8.Recordset.Fields!qty = Val(Text6.Text)
Adodc8.Recordset.Fields!items_no_by_one = Val(Text6.Text) * Val(DataCombo10.BoundText)
Adodc8.Recordset.Fields!ta2ther_makhzan = Val(Adodc8.Recordset.Fields!items_no_by_one) * -1


Adodc8.Recordset.Fields!average_cost = Val(Text9.Text) / Val(DataCombo10.BoundText)
Adodc8.Recordset.Fields!last_price = txt_last_price.Text
Adodc8.Recordset.Fields!transaction_type = "”‰œ ’—› „Œ“‰Ì"
 
Adodc8.Recordset.Fields!process_type = "«·Ï"
Adodc8.Recordset.Fields!process_text = "«„— ‘€·"
'tric
Adodc8.Recordset.Fields!sanad_no = "W" & Text14.Text
Adodc8.Recordset.Fields!bona_3la = "«„— ‘€· «·Ï Ê—ﬁ„…" & Text14.Text

Adodc8.Recordset.Fields!branch_no = branch_no
Adodc8.Recordset.Fields!departement = departement_name

Adodc8.Recordset.Fields!transaction_date = DateValue(Now)

Adodc8.Recordset.Fields!item_departement = Adodc16.Recordset.Fields!departement_name ' & "'" 'txtdepartement.Text


Adodc8.Recordset.Fields!total_price = Text10.Text

Adodc8.Recordset.Fields!technical = DataCombo8.Text
Adodc8.Recordset.Fields!technical_notes = Text11.Text

Adodc8.Recordset.Fields!inventory_name = DataCombo11.Text
Adodc8.Recordset.Fields!inventory_id = DataCombo11.BoundText


Adodc8.Recordset.Update
Adodc8.CommandType = adCmdText
Adodc8.RecordSource = "select * from inventory where   amr_shogl_fk=" & Val(Me.Text1.Text) & " and item_departement='" & Me.txtdepartement.Text & "'"
Adodc8.Refresh
'Adodc8.RecordSource = "select * from amr_soghl_details where opr_id=" & Text1.Text & " and item_departement='" & txtdepartement.Text & "'"
Adodc8.Refresh
DataGrid1.Refresh

DataCombo6.Text = ""
DataCombo10.Text = ""
Text9.Text = ""
Text6.Text = ""
Text10.Text = ""
Text11.Text = ""
DataCombo8.Text = ""

End Sub

Private Sub Label28_Click()
On Error Resume Next
    If my_language = "E" Then
        x = MsgBox("confirm send", vbCritical + vbYesNo)

    Else
    x = MsgBox(" √ﬂÌœ ⁄„·Ì… «·«—”«·", vbCritical + vbYesNo)
    End If
If x = vbNo Then
Exit Sub
End If

    If my_language = "E" Then
        If Adodc1.Recordset.RecordCount = o Then MsgBox "NO Job order To Close it", vbCritical: Exit Sub
    Else
        If Adodc1.Recordset.RecordCount = o Then MsgBox "·« ÌÊÃœ «„— ‘€· ·«—”«·…", vbCritical: Exit Sub
    End If
Adodc1.Recordset.Fields!departement_name_now = DataCombo9.Text
Adodc1.Recordset.Update
UPDATE_RECORDS
  If my_language = "E" Then
   MsgBox "send done to departement   " & DataCombo9.Text, vbInformation
  Else
  MsgBox " „ «·«—”«· «·Ï ﬁ”„   " & DataCombo9.Text, vbInformation
  End If
End Sub

Private Sub Label32_Click()
On Error Resume Next
If my_language = "E" Then

        If Text1.Text = "" Then MsgBox "no operation to delete", vbCritical: Exit Sub
        
        x = MsgBox("confirm delete", vbCritical + vbYesNo)
        If x = vbNo Then
        Exit Sub
        End If

Else

        If Text1.Text = "" Then MsgBox "·« ÌÊÃœ Õ—ﬂ… ·Õ›–›Â«", vbCritical: Exit Sub
        
        x = MsgBox("Â· «‰  „ √ﬂœ „‰ «·Õ–›", vbCritical + vbYesNo)
        If x = vbNo Then
        Exit Sub
        End If

End If


If Adodc8.Recordset.RecordCount > 0 Then
Adodc8.Recordset.Delete
Adodc8.Refresh

End If
End Sub

Private Sub Text1_change()
On Error Resume Next
Adodc8.ConnectionString = connection_string
Adodc8.CommandType = adCmdText
Adodc16.ConnectionString = connection_string
 Adodc16.CommandType = adCmdText
Adodc16.RecordSource = "select * from departement where departement_no=" & current_user_departement
Adodc16.Refresh
 
Me.txtdepartement.Text = Adodc16.Recordset.Fields!departement_name


If Text1.Text = "" Then Exit Sub
Adodc8.CommandType = adCmdText
Adodc8.RecordSource = "select * from inventory where   amr_shogl_fk=" & Val(Me.Text1.Text) & " and item_departement='" & Me.txtdepartement.Text & "'"
Adodc8.Refresh
DataGrid1.Refresh
'UPDATE_RECORDS
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then
Label19_Click
End If

End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then
Label19_Click
End If

End Sub

Private Sub Text6_Change()
On Error Resume Next
 For i = 1 To Len(Text6.Text)
    If Asc(Mid$(Text6.Text, i, 1)) < 48 Or Asc(Mid$(Text6.Text, i, 1)) > 57 Then
                If my_language = "E" Then
                MsgBox "Must enter Digit only", vbCritical
                Else
                MsgBox "·«»œ „‰ ﬂ «»… «—ﬁ«„", vbCritical
                End If
        Text6.Text = 0
        Text6.BackColor = vbRed
        Exit Sub
    End If
Next i
Text10.Text = Val(Text6.Text) * Val(Text9.Text)
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then
Label19_Click
End If

End Sub

Private Sub Text9_Change()
On Error Resume Next
'For i = 1 To Len(Text9.Text)
'    If Asc(Mid$(Text9.Text, i, 1)) < 48 Or Asc(Mid$(Text9.Text, i, 1)) > 57 Then
'                If my_language = "E" Then
'                MsgBox "Must enter Digit only", vbCritical
'                Else
'                MsgBox "·«»œ „‰ ﬂ «»… «—ﬁ«„", vbCritical
'                End If
'        Text9.Text = 0
'        Text9.BackColor = vbRed
'        Exit Sub
'    End If
'Next i
Text10.Text = Val(Text6.Text) * Val(Text9.Text)
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then
Label19_Click
End If

End Sub
