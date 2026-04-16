VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSubcontractorContract 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9150
   ClientLeft      =   -1485
   ClientTop       =   -1905
   ClientWidth     =   18225
   Icon            =   "frmSubcontractorContractl.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9152.811
   ScaleMode       =   0  'User
   ScaleWidth      =   18225
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   6480
      TabIndex        =   148
      Top             =   24960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   152
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4440
         TabIndex        =   151
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   150
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   372
         Left            =   2760
         TabIndex        =   149
         Top             =   120
         Width           =   1452
      End
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc1"
      Height          =   288
      ItemData        =   "frmSubcontractorContractl.frx":000C
      Left            =   28440
      List            =   "frmSubcontractorContractl.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   147
      Top             =   1440
      Visible         =   0   'False
      Width           =   4932
   End
   Begin VB.TextBox Text5 
      DataField       =   "last_root"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   2280
      TabIndex        =   138
      Text            =   "Text5"
      Top             =   16440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2280
      TabIndex        =   137
      Top             =   13920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   2280
      TabIndex        =   134
      Top             =   13920
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
         TabIndex        =   136
         Top             =   120
         Width           =   495
      End
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M15"
         Height          =   255
         Left            =   3360
         TabIndex        =   135
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   10920
      TabIndex        =   133
      Top             =   25920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   1080
      TabIndex        =   128
      Top             =   13920
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label3 
         Caption         =   "Center#"
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
         TabIndex        =   132
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Center Name"
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
         Height          =   372
         Left            =   360
         TabIndex        =   131
         Top             =   1320
         Width           =   1812
      End
      Begin VB.Label Label13 
         Caption         =   "Center Type"
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
         TabIndex        =   130
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Major Center"
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
         TabIndex        =   129
         Top             =   1920
         Width           =   2415
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   19800
      TabIndex        =   123
      Top             =   12480
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Line Line1 
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Center Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   127
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5640
         TabIndex        =   126
         Top             =   240
         Width           =   735
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3780
         X2              =   3780
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Center Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   125
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Center #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   124
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   5520
         X2              =   5520
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   720
      TabIndex        =   118
      Top             =   12720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Line Line1 
         Index           =   4
         X1              =   2700
         X2              =   2700
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6180
         X2              =   6180
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·„” ÊÏ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   122
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·„—ﬂ“"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   1440
         TabIndex        =   121
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4460
         X2              =   4460
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„—ﬂ“"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5040
         TabIndex        =   120
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„—ﬂ“"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   119
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   960
         X2              =   960
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   4095
      Left            =   12600
      TabIndex        =   114
      Top             =   13680
      Width           =   1455
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   0
         TabIndex        =   115
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   116
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "»«·—ﬁ„"
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
            MICON           =   "frmSubcontractorContractl.frx":003A
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
            TabIndex        =   117
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
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
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   2760
      TabIndex        =   109
      Top             =   13920
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   113
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   112
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   111
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   110
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox total1x 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "total"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   1560
      TabIndex        =   108
      Top             =   14400
      Width           =   1095
   End
   Begin VB.TextBox note_id 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2520
      TabIndex        =   106
      Top             =   12480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "⁄„·Ì«  ﬂ· »‰œ"
      Height          =   3615
      Left            =   13080
      RightToLeft     =   -1  'True
      TabIndex        =   98
      Top             =   26160
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox txt_opr_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   2760
         Width           =   3015
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   2340
         Left            =   120
         TabIndex        =   100
         Top             =   360
         Width           =   15000
         _cx             =   26458
         _cy             =   4128
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSubcontractorContractl.frx":0056
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin ALLButtonS.ALLButton opr_items 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   101
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··»‰Êœ"
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmSubcontractorContractl.frx":03D4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Show_items 
         Height          =   375
         Index           =   0
         Left            =   11280
         TabIndex        =   102
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "„Ê«œ "
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmSubcontractorContractl.frx":03F0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton employee_details 
         Height          =   375
         Index           =   0
         Left            =   9000
         TabIndex        =   103
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»Ì«‰«  «·⁄„«·…"
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmSubcontractorContractl.frx":040C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton opr_expenses 
         Height          =   375
         Index           =   0
         Left            =   6720
         TabIndex        =   104
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "„’«—Ì›"
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmSubcontractorContractl.frx":0428
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
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã„«·Ì"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5640
         TabIndex        =   105
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "„Ê«œ «·⁄„·Ì… —ﬁ„"
      Height          =   3615
      Left            =   14880
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   13920
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox XPTxtBillID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   360
         Left            =   840
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   1080
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox TxtFillData 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   0
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox XPTxtSum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1530
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   690
         Index           =   2
         Left            =   360
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   240
         Width           =   14715
         _cx             =   25956
         _cy             =   1217
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   900
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   300
            Width           =   1755
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   10500
            MaxLength       =   20
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   1140
            Width           =   2310
         End
         Begin VB.TextBox TxtQuantity 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2730
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   300
            Width           =   1770
         End
         Begin VB.ComboBox CboItemCase 
            Height          =   288
            Left            =   6870
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   300
            Width           =   1920
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   8805
            TabIndex        =   85
            Top             =   300
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo DCboItemsCode 
            Height          =   315
            Left            =   11790
            TabIndex        =   86
            Top             =   300
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
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
         Begin ImpulseButton.ISButton CmdAdd 
            Height          =   375
            Left            =   120
            TabIndex        =   87
            Top             =   270
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ButtonImage     =   "frmSubcontractorContractl.frx":0444
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—"
            Height          =   255
            Index           =   26
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   0
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì…"
            Height          =   255
            Index           =   27
            Left            =   2925
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   0
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ì—Ì«·"
            Height          =   255
            Index           =   28
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   0
            Width           =   2205
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·’‰›"
            Height          =   255
            Index           =   29
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   0
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈”„ «·’‰›"
            Height          =   255
            Index           =   30
            Left            =   9150
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   0
            Width           =   2640
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   31
            Left            =   11985
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   0
            Width           =   2700
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   1905
         Left            =   240
         TabIndex        =   94
         Top             =   960
         Width           =   14835
         _cx             =   26167
         _cy             =   3360
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSubcontractorContractl.frx":07DE
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   0
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin ALLButtonS.ALLButton Show_items 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   95
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«Œ›«¡"
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmSubcontractorContractl.frx":09A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblItemsCount 
         Caption         =   "Label27"
         Height          =   135
         Left            =   240
         TabIndex        =   97
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·«’‰«›"
         Height          =   255
         Index           =   2
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   3000
         Width           =   1575
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "«·„’—Ê›« "
      Height          =   3615
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   12600
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox txt_expenses_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   2340
         Left            =   240
         TabIndex        =   73
         Top             =   360
         Width           =   14760
         _cx             =   26035
         _cy             =   4128
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSubcontractorContractl.frx":09C2
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin ALLButtonS.ALLButton opr_expenses 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   74
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«Œ›«¡ «·„’—Ê›« "
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmSubcontractorContractl.frx":0AD0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·„’—Ê›« "
         Height          =   255
         Index           =   6
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "«”„«¡ «·⁄«„·Ì‰ ›Ì «·„‘—Ê⁄"
      Height          =   3615
      Left            =   29400
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   6480
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox txt_emp_salary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   6360
         TabIndex        =   66
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_employee_count 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   9960
         TabIndex        =   65
         Top             =   3000
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   1860
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Width           =   15120
         _cx             =   26670
         _cy             =   3281
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSubcontractorContractl.frx":0AEC
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin ALLButtonS.ALLButton employee_details 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   68
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmSubcontractorContractl.frx":0CC5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label24 
         Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·"
         Height          =   255
         Left            =   8040
         TabIndex        =   70
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "«Ã„«·Ì ⁄œœ «·⁄„·"
         Height          =   255
         Left            =   11640
         TabIndex        =   69
         Top             =   3120
         Width           =   1815
      End
   End
   Begin VB.TextBox TXTsub_contractor_id 
      Height          =   375
      Left            =   29280
      TabIndex        =   63
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TXTEnd_user_id 
      Height          =   285
      Left            =   29160
      TabIndex        =   62
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtModFlg 
      Height          =   285
      Left            =   29400
      TabIndex        =   61
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtsubaccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   29160
      TabIndex        =   60
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtendaccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   29400
      TabIndex        =   59
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtdate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   29280
      TabIndex        =   58
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9150
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18225
      _cx             =   32147
      _cy             =   16140
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   3540
         Left            =   -15960
         TabIndex        =   245
         TabStop         =   0   'False
         Top             =   6930
         Visible         =   0   'False
         Width           =   16095
         _cx             =   28390
         _cy             =   6244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   3150
            Left            =   -4470
            TabIndex        =   246
            Tag             =   "1"
            Top             =   240
            Width           =   15870
            _cx             =   27993
            _cy             =   5556
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   14871017
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSubcontractorContractl.frx":0CE1
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   -1  'True
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   17400
            RightToLeft     =   -1  'True
            TabIndex        =   250
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Index           =   6
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   247
            Top             =   6120
            Width           =   7095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   615
         Left            =   0
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   8535
         Width           =   18225
         _cx             =   32147
         _cy             =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   2
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   0
            Left            =   16815
            TabIndex        =   159
            Top             =   0
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   1270
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
            MICON           =   "frmSubcontractorContractl.frx":0E24
            PICN            =   "frmSubcontractorContractl.frx":0E40
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   1
            Left            =   14265
            TabIndex        =   160
            Top             =   0
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   1270
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
            MICON           =   "frmSubcontractorContractl.frx":76A2
            PICN            =   "frmSubcontractorContractl.frx":76BE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   3
            Left            =   15645
            TabIndex        =   161
            Top             =   0
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   1270
            BTYPE           =   3
            TX              =   " ⁄œÌ·"
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
            MICON           =   "frmSubcontractorContractl.frx":DF20
            PICN            =   "frmSubcontractorContractl.frx":DF3C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   6
            Left            =   12885
            TabIndex        =   162
            Top             =   0
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   1270
            BTYPE           =   3
            TX              =   " —«Ã⁄"
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
            MICON           =   "frmSubcontractorContractl.frx":1479E
            PICN            =   "frmSubcontractorContractl.frx":147BA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   7
            Left            =   4380
            TabIndex        =   163
            Top             =   0
            Visible         =   0   'False
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   1270
            BTYPE           =   3
            TX              =   "ÿ»«⁄… «·„” Œ·’"
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
            MICON           =   "frmSubcontractorContractl.frx":1B01C
            PICN            =   "frmSubcontractorContractl.frx":1B038
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   8
            Left            =   0
            TabIndex        =   164
            Top             =   0
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1270
            BTYPE           =   3
            TX              =   "ÿ»«⁄Â «·ﬁÌœ"
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
            MICON           =   "frmSubcontractorContractl.frx":2189A
            PICN            =   "frmSubcontractorContractl.frx":218B6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   9
            Left            =   11775
            TabIndex        =   165
            Top             =   0
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   1270
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
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "frmSubcontractorContractl.frx":28118
            PICN            =   "frmSubcontractorContractl.frx":28134
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   10
            Left            =   10440
            TabIndex        =   166
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   1270
            BTYPE           =   3
            TX              =   "»ÕÀ"
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
            MICON           =   "frmSubcontractorContractl.frx":2E996
            PICN            =   "frmSubcontractorContractl.frx":2E9B2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   11
            Left            =   5970
            TabIndex        =   167
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   1270
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
            BCOL            =   255
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "frmSubcontractorContractl.frx":35214
            PICN            =   "frmSubcontractorContractl.frx":35230
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   12
            Left            =   9090
            TabIndex        =   168
            Top             =   0
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   1270
            BTYPE           =   3
            TX              =   "‰”Œ… „„«À·…"
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
            MICON           =   "frmSubcontractorContractl.frx":3BA92
            PICN            =   "frmSubcontractorContractl.frx":3BAAE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   15
            Left            =   7230
            TabIndex        =   175
            Top             =   0
            Visible         =   0   'False
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   1270
            BTYPE           =   3
            TX              =   "≈‰‘«¡ «·›Ê« Ì— «·‘Â—Ì…"
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
            MICON           =   "frmSubcontractorContractl.frx":42310
            PICN            =   "frmSubcontractorContractl.frx":4232C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   720
            Index           =   16
            Left            =   2850
            TabIndex        =   205
            Top             =   0
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1270
            BTYPE           =   3
            TX              =   "ÿ»«⁄… „Êﬁ› «·„” Œ·’« "
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
            MICON           =   "frmSubcontractorContractl.frx":48B8E
            PICN            =   "frmSubcontractorContractl.frx":48BAA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Accredit 
            Height          =   345
            Left            =   1320
            TabIndex        =   248
            Top             =   0
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··«⁄ „«œ"
            BackColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   -2147483635
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   345
            Left            =   1320
            TabIndex        =   249
            Top             =   360
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
            ButtonPositionImage=   1
            Caption         =   "Õ«·Â «·«⁄ „«œ"
            BackColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   -2147483635
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   135
         Left            =   21735
         TabIndex        =   143
         Top             =   1560
         Width           =   1830
      End
      Begin C1SizerLibCtl.C1Elastic Frame2 
         Height          =   4770
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3180
         Width           =   18120
         _cx             =   31962
         _cy             =   8414
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "»‰Êœ «·„‘—Ê⁄"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.TextBox txtManulaVat 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   435
            Left            =   5400
            TabIndex        =   243
            Top             =   3960
            Width           =   795
         End
         Begin XtremeSuiteControls.CheckBox ChQty 
            Height          =   255
            Left            =   11700
            TabIndex        =   242
            Top             =   -1740
            Width           =   2775
            _Version        =   786432
            _ExtentX        =   4895
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·ﬂ„Ì… «·„‰›–…  ”«ÊÌ «·ﬂ„Ì… «·›⁄·Ì…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmSubcontractorContractl.frx":4F40C
            Left            =   5160
            List            =   "frmSubcontractorContractl.frx":4F42E
            RightToLeft     =   -1  'True
            TabIndex        =   240
            Top             =   4440
            Width           =   5235
         End
         Begin VB.TextBox TxtNetValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   11040
            Locked          =   -1  'True
            TabIndex        =   238
            Top             =   3960
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.TextBox TxtPerforValue 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   435
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   236
            Top             =   3960
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.TextBox TxtTotalValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   3960
            Width           =   3180
         End
         Begin VB.TextBox TxtFATValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   3960
            Width           =   1875
         End
         Begin VB.TextBox TxtFATYou 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   194
            Top             =   3960
            Width           =   795
         End
         Begin VB.Frame Frame7 
            Height          =   3015
            Left            =   -3540
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   2250
            Width           =   4215
            Begin VB.TextBox TxtValueTemp 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   0
               TabIndex        =   216
               Top             =   0
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox TxtRemarks2 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   795
               Left            =   360
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   185
               Top             =   1920
               Width           =   1815
            End
            Begin VB.TextBox TxtBillNo 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   360
               TabIndex        =   181
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txtPeriod 
               Height          =   315
               Index           =   0
               Left            =   1440
               TabIndex        =   179
               Top             =   1440
               Width           =   735
            End
            Begin VB.ComboBox DcbPeriodType 
               Height          =   315
               ItemData        =   "frmSubcontractorContractl.frx":4F483
               Left            =   360
               List            =   "frmSubcontractorContractl.frx":4F490
               TabIndex        =   178
               Top             =   1440
               Width           =   975
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   10
               Left            =   3360
               TabIndex        =   177
               Top             =   120
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   476
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«€·«ﬁ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "frmSubcontractorContractl.frx":4F4A3
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker StartDate 
               Height          =   315
               Left            =   360
               TabIndex        =   183
               Top             =   1080
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   147587073
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DateTemp 
               Height          =   315
               Left            =   2040
               TabIndex        =   188
               Top             =   2400
               Visible         =   0   'False
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   147587073
               CurrentDate     =   41640
            End
            Begin VB.Label Label42 
               Alignment       =   2  'Center
               Caption         =   "≈‰‘«¡ «·›Ê« Ì— «·‘Â—Ì…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   480
               TabIndex        =   187
               Top             =   240
               Width           =   1890
            End
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               Caption         =   "„·«ÕŸ«  "
               Height          =   240
               Left            =   2160
               TabIndex        =   186
               Top             =   2280
               Width           =   1890
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "Ì»œ« „‰  «—ÌŒ"
               Height          =   240
               Index           =   1
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   1125
               Width           =   1890
            End
            Begin VB.Label Label40 
               Alignment       =   2  'Center
               Caption         =   "⁄œœ «·›Ê« Ì—"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1740
               TabIndex        =   182
               Top             =   2070
               Width           =   1890
            End
            Begin VB.Label Label39 
               Alignment       =   2  'Center
               Caption         =   "«·„œ… »Ì‰ «·›Ê« Ì—"
               Height          =   240
               Left            =   2280
               TabIndex        =   180
               Top             =   1440
               Width           =   1890
            End
         End
         Begin VB.TextBox total 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   3960
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.TextBox Results 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   14880
            TabIndex        =   50
            Top             =   3960
            Width           =   1875
         End
         Begin VB.TextBox txtDiscountG 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   435
            Left            =   12975
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   3960
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
            Height          =   3135
            Left            =   120
            TabIndex        =   52
            Top             =   540
            Width           =   17880
            _cx             =   31538
            _cy             =   5530
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16777152
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   53
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSubcontractorContractl.frx":4FA3D
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   -1  'True
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   435
            Left            =   16965
            TabIndex        =   53
            Tag             =   "Delete Row"
            Top             =   3705
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   767
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16776960
            BCOLO           =   16776960
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSubcontractorContractl.frx":50440
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   11640
            TabIndex        =   173
            Top             =   4440
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AccountVat 
            Bindings        =   "frmSubcontractorContractl.frx":5045C
            Height          =   315
            Left            =   0
            TabIndex        =   199
            Top             =   4440
            Visible         =   0   'False
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«œŒ«· «·‰”»… «·ÌœÊÌ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Index           =   148
            Left            =   4560
            TabIndex        =   244
            Top             =   3720
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ ‰„Ê–Ã"
            Height          =   345
            Index           =   0
            Left            =   10425
            RightToLeft     =   -1  'True
            TabIndex        =   241
            Top             =   4455
            Width           =   1065
         End
         Begin VB.Label Label60 
            Alignment       =   2  'Center
            Caption         =   "’«›Ì «·›« Ê—…"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   11160
            TabIndex        =   239
            Top             =   3720
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label Label59 
            Alignment       =   2  'Center
            Caption         =   "Œ’„ Õ”‰ «·«œ«¡"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   9240
            TabIndex        =   237
            Top             =   3720
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·›« "
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   67
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   3720
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   68
            Left            =   1245
            RightToLeft     =   -1  'True
            TabIndex        =   198
            Top             =   3720
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»…«·›« "
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   66
            Left            =   6285
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   3720
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   22
            Left            =   17160
            TabIndex        =   174
            Top             =   4440
            Width           =   900
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   120
            TabIndex        =   172
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   2400
            TabIndex        =   171
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   20
            Left            =   3120
            TabIndex        =   170
            Top             =   4440
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   21
            Left            =   1260
            TabIndex        =   169
            Top             =   4440
            Width           =   1065
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            Caption         =   "«Ã„«·Ì «·›« Ê—…"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   14985
            TabIndex        =   56
            Top             =   3720
            Width           =   1545
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "Œ’„ «·›« Ê—…"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   13005
            TabIndex        =   55
            Top             =   3720
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "’«›Ì «·›« Ê—…"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   7380
            TabIndex        =   54
            Top             =   3720
            Visible         =   0   'False
            Width           =   1530
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frame13 
         Height          =   2205
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   990
         Width           =   18135
         _cx             =   31988
         _cy             =   3889
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.TextBox txtqtySubContractor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   5550
            TabIndex        =   266
            Top             =   1440
            Width           =   1050
         End
         Begin VB.TextBox txtcostSubContractor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   4140
            TabIndex        =   264
            Top             =   1440
            Width           =   1380
         End
         Begin VB.TextBox txtPeriod 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Index           =   1
            Left            =   3060
            TabIndex        =   261
            Top             =   1440
            Width           =   1050
         End
         Begin VB.TextBox Txtqty 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   345
            Left            =   8070
            TabIndex        =   259
            Top             =   1470
            Width           =   1050
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H0080FFFF&
            Caption         =   "»Ì«‰«  «·œ›⁄«  «·„ﬁœ„…"
            Height          =   2190
            Left            =   -16860
            RightToLeft     =   -1  'True
            TabIndex        =   209
            Top             =   2745
            Visible         =   0   'False
            Width           =   14895
            Begin VB.TextBox TxtPreBalaVATYu 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   8280
               TabIndex        =   231
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaNet 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   229
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaTransPyed 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1560
               TabIndex        =   224
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaRemain 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   3000
               TabIndex        =   223
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaPayed 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   4440
               TabIndex        =   222
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaTotal 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   5880
               TabIndex        =   221
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaVAT 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   7320
               TabIndex        =   219
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   8760
               TabIndex        =   217
               Top             =   480
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   195
               Left            =   13080
               RightToLeft     =   -1  'True
               TabIndex        =   211
               Top             =   540
               Width           =   1200
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H8000000B&
               Caption         =   "«·€«¡ «·”œ«œ"
               Height          =   315
               Left            =   11400
               RightToLeft     =   -1  'True
               TabIndex        =   210
               Top             =   480
               Width           =   1695
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
               CausesValidation=   0   'False
               Height          =   1740
               Left            =   120
               TabIndex        =   212
               Top             =   840
               Width           =   14640
               _cx             =   25823
               _cy             =   3069
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   12648447
               ForeColorSel    =   16711680
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   5
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   18
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmSubcontractorContractl.frx":50471
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   1
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   -1  'True
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   255
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   233
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2640
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   255
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   232
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2640
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "«·’«›Ì «·„” Õﬁ"
               Height          =   255
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   230
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "„”œœ «·Õ—ﬂ…"
               Height          =   255
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   228
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "„ »ﬁÌ"
               Height          =   255
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   227
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "„”œœ „”»ﬁ«"
               Height          =   255
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   226
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "«·ﬁÌ„… «·‘«„·…"
               Height          =   255
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   225
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·›« "
               Height          =   255
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   220
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "«·ﬁÌ„…"
               Height          =   255
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   218
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì"
               Height          =   255
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   215
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2640
               Width           =   1575
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   214
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2640
               Width           =   3015
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   14520
               RightToLeft     =   -1  'True
               TabIndex        =   213
               Top             =   240
               Width           =   135
            End
         End
         Begin VB.TextBox TxtPreVAT 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   225
            Left            =   -1350
            TabIndex        =   206
            Top             =   3000
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtcost 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   345
            Left            =   6630
            TabIndex        =   203
            Top             =   1470
            Width           =   1380
         End
         Begin VB.ComboBox DcbExPercen 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmSubcontractorContractl.frx":50730
            Left            =   2130
            List            =   "frmSubcontractorContractl.frx":5073D
            TabIndex        =   201
            Top             =   2745
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   14640
            TabIndex        =   193
            Top             =   75
            Width           =   1740
         End
         Begin VB.TextBox TxtAccountUnderImp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   1140
            TabIndex        =   192
            Top             =   2430
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "›⁄·Ì"
            Height          =   165
            Left            =   -3960
            TabIndex        =   191
            Top             =   2475
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0FFFF&
            Caption         =   " ﬁœÌ—Ì"
            Height          =   165
            Left            =   -2775
            TabIndex        =   190
            Top             =   2475
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0FFFF&
            Caption         =   " Õ  «· ‰›Ì–"
            Height          =   165
            Left            =   -5550
            TabIndex        =   189
            Top             =   2475
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.TextBox DcAccount1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   360
            Left            =   4335
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   690
            Visible         =   0   'False
            Width           =   5235
         End
         Begin VB.TextBox DcAccount2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   -630
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   3000
            Visible         =   0   'False
            Width           =   4710
         End
         Begin VB.ComboBox billto 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmSubcontractorContractl.frx":50752
            Left            =   -1095
            List            =   "frmSubcontractorContractl.frx":5075C
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   2385
            Visible         =   0   'False
            Width           =   4710
         End
         Begin VB.ComboBox bill_Type 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmSubcontractorContractl.frx":50778
            Left            =   3765
            List            =   "frmSubcontractorContractl.frx":5077A
            TabIndex        =   19
            Top             =   2160
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox txtid 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   14640
            TabIndex        =   18
            Top             =   120
            Width           =   1740
         End
         Begin VB.TextBox txtprojectname 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            DataField       =   "project_name"
            DataSource      =   "Adodc1"
            Height          =   345
            Left            =   13185
            TabIndex        =   17
            Top             =   1500
            Width           =   3300
         End
         Begin VB.TextBox TxtRemarks 
            Height          =   510
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   3030
            Width           =   3810
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   15135
            TabIndex        =   15
            Top             =   765
            Width           =   1275
         End
         Begin VB.ComboBox cboDiscount1 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmSubcontractorContractl.frx":5077C
            Left            =   4590
            List            =   "frmSubcontractorContractl.frx":50789
            TabIndex        =   14
            Top             =   2145
            Visible         =   0   'False
            Width           =   2580
         End
         Begin VB.ComboBox cboDiscount2 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmSubcontractorContractl.frx":5079E
            Left            =   4590
            List            =   "frmSubcontractorContractl.frx":507AB
            TabIndex        =   13
            Top             =   2385
            Visible         =   0   'False
            Width           =   2580
         End
         Begin VB.TextBox txtDiscount1 
            BackColor       =   &H00C0FFFF&
            Height          =   195
            Left            =   2475
            TabIndex        =   12
            Top             =   2145
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox txtDiscount2 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2475
            TabIndex        =   11
            Top             =   2400
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox txtManualNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   11640
            TabIndex        =   10
            Top             =   120
            Width           =   1515
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   225
            Left            =   2130
            TabIndex        =   9
            Top             =   2475
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.TextBox advancedPayment 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   225
            Left            =   -1350
            TabIndex        =   8
            Top             =   2745
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   16485
            TabIndex        =   23
            Top             =   1500
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
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
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   300
            Left            =   14610
            TabIndex        =   24
            Top             =   435
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Format          =   147259393
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker dueDate1 
            Height          =   210
            Left            =   -1335
            TabIndex        =   25
            Top             =   2175
            Visible         =   0   'False
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   370
            _Version        =   393216
            Format          =   147193857
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Height          =   315
            Left            =   4335
            TabIndex        =   26
            Top             =   75
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSComCtl2.DTPicker dueDate 
            Height          =   210
            Left            =   2175
            TabIndex        =   27
            Top             =   2175
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   370
            _Version        =   393216
            Format          =   118816769
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbosubContractor 
            Height          =   315
            Left            =   11700
            TabIndex        =   28
            Top             =   765
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            _Version        =   393216
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
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   330
            Left            =   -1140
            TabIndex        =   208
            Tag             =   "Delete Row"
            Top             =   2430
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   582
            BTYPE           =   3
            TX              =   "⁄—÷ «·œ›⁄«  «·„ﬁœ„…"
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
            BCOL            =   16776960
            BCOLO           =   16776960
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSubcontractorContractl.frx":507C0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker StartDateProje 
            Height          =   210
            Left            =   -4230
            TabIndex        =   234
            Top             =   2175
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   370
            _Version        =   393216
            Format          =   118816769
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo cmbBand 
            Height          =   315
            Left            =   9180
            TabIndex        =   257
            Top             =   1500
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
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
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   315
            Left            =   60
            TabIndex        =   263
            Tag             =   "Delete Row"
            Top             =   1440
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "≈÷«›… ﬂ· «·»‰Êœ"
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
            BCOL            =   16776960
            BCOLO           =   16776960
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSubcontractorContractl.frx":507DC
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
            Height          =   315
            Left            =   1590
            TabIndex        =   268
            Tag             =   "Delete Row"
            Top             =   1440
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "≈÷«›… «·»‰œ «·„Õœœ"
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
            BCOL            =   16776960
            BCOLO           =   16776960
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSubcontractorContractl.frx":507F8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label63 
            Alignment       =   2  'Center
            Caption         =   "ﬂ„Ì… «· ⁄«ﬁœ"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   5580
            TabIndex        =   267
            Top             =   1110
            Width           =   1065
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            Caption         =   "”⁄— «· ⁄«ﬁœ"
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   2
            Left            =   4110
            TabIndex        =   265
            Top             =   1110
            Width           =   1605
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            Caption         =   "„œ… «· ‰›Ì–"
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   1
            Left            =   2850
            TabIndex        =   262
            Top             =   1080
            Width           =   1605
         End
         Begin VB.Label Label62 
            Alignment       =   2  'Center
            Caption         =   "ﬂ„Ì… «·»‰œ"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   8100
            TabIndex        =   260
            Top             =   1140
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "«·»‰œ"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   10500
            TabIndex        =   258
            Top             =   1170
            Width           =   1065
         End
         Begin VB.Label Label58 
            Alignment       =   2  'Center
            Caption         =   " «—ÌŒ »œ«Ì… «·„‘—Ê⁄"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   -2790
            TabIndex        =   235
            Top             =   2175
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            Caption         =   "VAT «·œ›⁄… «·„ﬁœ„…"
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   450
            TabIndex        =   207
            Top             =   3000
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            Caption         =   "«·”⁄—"
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   0
            Left            =   6630
            TabIndex        =   204
            Top             =   1140
            Width           =   1905
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "‰Ê⁄ «·„” Œ·’"
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   1170
            TabIndex        =   202
            Top             =   2955
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   $"frmSubcontractorContractl.frx":50814
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   735
            Index           =   5
            Left            =   -1050
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   2685
            Width           =   4200
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Õ Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   975
            TabIndex        =   46
            Top             =   2175
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„·«ÕŸ… Â«„…:-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   135
            Index           =   4
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2835
            Width           =   1200
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   810
            Left            =   -1530
            Top             =   2160
            Visible         =   0   'False
            Width           =   4125
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "«”„ „ﬁ«Ê· «·»«ÿ‰"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10035
            TabIndex        =   44
            Top             =   690
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            Caption         =   "‰Ê⁄ «·„” Œ·’"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5160
            TabIndex        =   43
            Top             =   2160
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„” Œ·’ «·Ï"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   3120
            TabIndex        =   42
            Top             =   2355
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "«· «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   16755
            TabIndex        =   41
            Top             =   435
            Width           =   1230
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   15900
            TabIndex        =   40
            Top             =   120
            Width           =   2085
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ «·⁄„Ì· «·‰Â«∆Ì"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2670
            TabIndex        =   39
            Top             =   3135
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "«·„‘—Ê⁄"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   15330
            TabIndex        =   38
            Top             =   1170
            Width           =   1065
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "„·«ÕŸ« "
            ForeColor       =   &H00000000&
            Height          =   150
            Left            =   1680
            TabIndex        =   37
            Top             =   2730
            Width           =   1905
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   10500
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4245
            TabIndex        =   35
            Top             =   2175
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„ﬁ«Ê· «·»«ÿ‰"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   15930
            TabIndex        =   34
            Top             =   765
            Width           =   2085
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ’„ ÷„«‰ «·«⁄„«·"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   6690
            TabIndex        =   33
            Top             =   2145
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ’„  œ›⁄Â „ﬁœ„…"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6690
            TabIndex        =   32
            Top             =   2385
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   13305
            TabIndex        =   31
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "Õ”„Ì« "
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   4245
            TabIndex        =   30
            Top             =   2475
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            Caption         =   "«·œ›⁄Â «·„ﬁœ„…"
            ForeColor       =   &H00000000&
            Height          =   150
            Left            =   465
            TabIndex        =   29
            Top             =   2745
            Visible         =   0   'False
            Width           =   1890
         End
      End
      Begin ALLButtonS.ALLButton CMD_language 
         Height          =   690
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Language  «··€…"
         Top             =   150
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   1217
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
         MICON           =   "frmSubcontractorContractl.frx":508CF
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
         Height          =   510
         Index           =   0
         Left            =   2970
         TabIndex        =   2
         Top             =   240
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   900
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmSubcontractorContractl.frx":508EB
         ColorButton     =   16777215
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
         Height          =   510
         Index           =   2
         Left            =   1755
         TabIndex        =   3
         Top             =   240
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   900
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmSubcontractorContractl.frx":50C85
         ColorButton     =   16777215
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
         Height          =   510
         Index           =   1
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
         ButtonStyle     =   1
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmSubcontractorContractl.frx":5101F
         ColorButton     =   16777215
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         Alignment       =   0
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         RightToLeft     =   -1  'True
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   510
         Index           =   3
         Left            =   2370
         TabIndex        =   5
         Top             =   240
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   900
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmSubcontractorContractl.frx":513B9
         ColorButton     =   16777215
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   405
         Index           =   13
         Left            =   23940
         TabIndex        =   156
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "»ÕÀ"
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
         MICON           =   "frmSubcontractorContractl.frx":51753
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
         Height          =   405
         Index           =   14
         Left            =   19800
         TabIndex        =   157
         Top             =   0
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "»ÕÀ"
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
         MICON           =   "frmSubcontractorContractl.frx":5176F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   256
         Top             =   8160
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   315
         Index           =   10
         Left            =   4080
         TabIndex        =   255
         Top             =   8160
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   315
         Index           =   9
         Left            =   7080
         TabIndex        =   254
         Top             =   8160
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»«·€ «·„Õ Ã“Â  «Ã„«·Ì "
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   8
         Left            =   1560
         TabIndex        =   253
         Top             =   8160
         Width           =   2025
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»«·€ «·„Õ Ã“Â «·Õ«·ÌÂ"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   7
         Left            =   5160
         TabIndex        =   252
         Top             =   8160
         Width           =   2025
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»«·€ «·„Õ Ã“Â «·”«»ﬁ…"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   3
         Left            =   8400
         TabIndex        =   251
         Top             =   8160
         Width           =   2025
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   9360
         Picture         =   "frmSubcontractorContractl.frx":5178B
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   9765
         Picture         =   "frmSubcontractorContractl.frx":553F3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "—ﬁ„ «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   21270
         TabIndex        =   145
         Top             =   3000
         Width           =   2025
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   19995
         TabIndex        =   144
         Top             =   2235
         Width           =   1515
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄ﬁœ „ﬁ«Ê· »«ÿ‰"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   840
         Left            =   -1050
         TabIndex        =   6
         Top             =   0
         Width           =   19305
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   7935
         Picture         =   "frmSubcontractorContractl.frx":56960
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSubcontractorContractl.frx":57ECD
      Height          =   2892
      Left            =   9240
      TabIndex        =   139
      Top             =   12720
      Visible         =   0   'False
      Width           =   7572
      _ExtentX        =   13361
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
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
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "account_no"
         Caption         =   "—ﬁ„ «·„‘—Ê⁄"
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
         DataField       =   "account_name"
         Caption         =   "«”„ «·„‘—Ê⁄"
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
         DataField       =   "account_type"
         Caption         =   "‰Ê⁄ «·„‘—Ê⁄"
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
         DataField       =   "parent_no"
         Caption         =   "parent_no"
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
         DataField       =   "child_no"
         Caption         =   "child_no"
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
         DataField       =   "level"
         Caption         =   "«·„” ÊÏ"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4500.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   576
      Left            =   960
      Top             =   13200
      Visible         =   0   'False
      Width           =   1788
      _ExtentX        =   3149
      _ExtentY        =   1005
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
      Height          =   336
      Left            =   10560
      Top             =   13800
      Visible         =   0   'False
      Width           =   1788
      _ExtentX        =   3149
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
      Height          =   336
      Left            =   10560
      Top             =   12720
      Visible         =   0   'False
      Width           =   1788
      _ExtentX        =   3149
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
      Height          =   336
      Left            =   10560
      Top             =   14160
      Visible         =   0   'False
      Width           =   1788
      _ExtentX        =   3149
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
      Height          =   336
      Left            =   10560
      Top             =   12840
      Visible         =   0   'False
      Width           =   1788
      _ExtentX        =   3149
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
      Height          =   336
      Left            =   10560
      Top             =   15840
      Visible         =   0   'False
      Width           =   1788
      _ExtentX        =   3149
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
      Height          =   336
      Left            =   4320
      Top             =   12840
      Width           =   7692
      _ExtentX        =   13573
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
      Caption         =   "Adodc2"
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
   Begin ALLButtonS.ALLButton opr_items 
      Height          =   372
      Index           =   0
      Left            =   28560
      TabIndex        =   153
      Top             =   1440
      Visible         =   0   'False
      Width           =   4344
      _ExtentX        =   7673
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "⁄„·Ì«  «·»‰œ"
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
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "frmSubcontractorContractl.frx":57EE2
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
      Height          =   372
      Index           =   2
      Left            =   28800
      TabIndex        =   154
      Top             =   480
      Visible         =   0   'False
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   661
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
      BCOL            =   255
      BCOLO           =   192
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "frmSubcontractorContractl.frx":57EFE
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
      Height          =   372
      Index           =   5
      Left            =   28080
      TabIndex        =   155
      Top             =   120
      Visible         =   0   'False
      Width           =   1092
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "»ÕÀ"
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
      MICON           =   "frmSubcontractorContractl.frx":57F1A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtrevenue_account 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   240
      TabIndex        =   107
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      Caption         =   "·œ›⁄Â „ÕœœÂ"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   30360
      TabIndex        =   146
      Top             =   2640
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label5 
      Caption         =   "Label2"
      Height          =   15
      Left            =   12120
      TabIndex        =   142
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   12720
      TabIndex        =   141
      Top             =   12840
      Width           =   2172
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "·œ›⁄Â „ÕœœÂ"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   28560
      TabIndex        =   140
      Top             =   5280
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   12120
      TabIndex        =   57
      Top             =   9720
      Width           =   855
   End
End
Attribute VB_Name = "frmSubcontractorContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim x As Long
Dim last_root As Integer
Dim last_geeral As Integer
Dim last_branch As Integer
Dim mod_flad As String
Dim first_run  As Boolean
Dim rs As ADODB.Recordset
Dim RsDev As ADODB.Recordset
Dim current_terms As String
Dim current_opr As String
Dim NewGrid As New ClsGrid
Dim expanses_account As String
Dim AcountGood As String
Dim maa_rs As ADODB.Recordset
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
Dim cCompanyInfo As New ClsCompanyInfo
Dim FlgBillBuy As Boolean

Sub ReloadContrac(Optional project_no As Double)
Dim Dcombos As ClsDataCombos
Dim StrSQL As String
Set Dcombos = New ClsDataCombos
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select CusID,CusName From TblCustemers"
Else
    StrSQL = "Select CusID,CusNamee From TblCustemers"
End If
StrSQL = StrSQL & " Where Type = 3"
'StrSQL = StrSQL & " where CusID in(SELECT     sub_contractor_id"
'StrSQL = StrSQL & " From dbo.projects_des)"
'StrSQL = StrSQL & " WHERE     (project_id = " & project_no & "))"
'Dcombos.ClearMyDataCombo DcbosubContractor

'fill_combo Me.DcbosubContractor, StrSQL

StrSQL = "  SELECT oprid,des FROM projects_des "
StrSQL = StrSQL & " WHERE     (project_id = " & project_no & ")"
fill_combo Me.cmbBand, StrSQL

End Sub
'ma
Public Sub Search(ID As Integer)


   Set maa_rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.SubcontractorContract  where 1 =1"
 
 StrSQL = " SELECT *  From dbo.SubcontractorContract  where  ID=  " & ID & " Order by ID "
   
    
   maa_rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If maa_rs.RecordCount < 1 Then
        '    XPTxtCurrent.Caption = 0
        '    XPTxtCount.Caption = 0
        Exit Sub
    End If
    If maa_rs.EOF Or maa_rs.BOF Then
        Exit Sub
    Else
    
    
    maaRetrive
End If
End Sub


Private Sub maaRetrive()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If maa_rs.RecordCount < 1 Then
        '    XPTxtCurrent.Caption = 0
        '    XPTxtCount.Caption = 0
        Exit Sub
    End If

    If maa_rs.EOF Or maa_rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            maa_rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If maa_rs.EOF Or maa_rs.BOF Then
                Exit Sub
            End If
        End If
    End If
Me.dcBranch.BoundText = IIf(IsNull(maa_rs("branch_no").value), "", maa_rs("branch_no").value)
    txtid.Text = IIf(IsNull(maa_rs("id").value), 0, val(maa_rs("id").value))

    XPDtbTrans.value = IIf(IsNull(maa_rs("bill_date").value), Date, maa_rs("bill_date").value)
dueDate.value = IIf(IsNull(maa_rs("dueDate").value), Date, maa_rs("dueDate").value)
dueDate1.value = IIf(IsNull(maa_rs("dueDate1").value), Date, maa_rs("dueDate1").value)

    DataCombo2.BoundText = IIf(IsNull(rs("project_no").value), "", rs("project_no").value)
'*************************************************
DcbosubContractor.BoundText = IIf(IsNull(maa_rs("subContractorId").value), "", maa_rs("subContractorId").value)
txtDiscount1.Text = IIf(IsNull(maa_rs("discount1value").value), 0, (maa_rs("discount1value").value))
txtDiscount2.Text = IIf(IsNull(maa_rs("discount2value").value), 0, (maa_rs("discount2value").value))

cboDiscount1.ListIndex = IIf(IsNull(maa_rs("discount1ID").value), 0, (maa_rs("discount1ID").value))
cboDiscount2.ListIndex = IIf(IsNull(maa_rs("discount2ID").value), 0, (maa_rs("discount2ID").value))

'*************************************************


    txtprojectname.Text = IIf(IsNull(maa_rs("project_name").value), "", maa_rs("project_name").value)
'    DcAccount1.text = IIf(IsNull(rs("Sub_user_name").value), "", rs("Sub_user_name").value)
'    DcAccount2.text = IIf(IsNull(rs("End_user_name").value), "", rs("End_user_name").value)

    txtendaccount.Text = IIf(IsNull(maa_rs("End_user_account").value), "", maa_rs("End_user_account").value)
    txtsubaccount.Text = IIf(IsNull(maa_rs("Sub_user_account").value), "", maa_rs("Sub_user_account").value)
    txtrevenue_account.Text = IIf(IsNull(maa_rs("revenue_account").value), "", maa_rs("revenue_account").value)

    'DcAccount4.text = IIf(IsNull(Rs("sub_contractor_name").value), "", Rs("sub_contractor_name").value)

    'DcAccount2.text = IIf(IsNull(Rs("End_user_name").value), "", Rs("End_user_name").value)

    billto.ListIndex = IIf(IsNull(maa_rs("bill_to").value), -1, maa_rs("bill_to").value)
    bill_Type.Text = IIf(IsNull(maa_rs("bill_type").value), 0, val(maa_rs("bill_type").value))
    Me.note_id.Text = IIf(IsNull(maa_rs("note_id").value), "", maa_rs("note_id").value)
    TxtNoteSerial.Text = IIf(IsNull(maa_rs("NoteSerial").value), "", maa_rs("NoteSerial").value)
    TxtRemarks.Text = IIf(IsNull(maa_rs("Remarks").value), "", maa_rs("Remarks").value)
        TxtManualNO.Text = IIf(IsNull(maa_rs("ManualNo").value), "", maa_rs("ManualNo").value)
        

'rs("Remarks").value = Trim(TxtRemarks.text)
'rs("ManualNo").value = Trim(txtManualNo.text)

    total.Text = Round(IIf(IsNull(maa_rs("total").value), 0, maa_rs("total").value), Decimal_Places)

    'Exit Sub

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "SELECT     item_id,id, project_no, item, cost, exe, percentage, exedate, bill_id,item_unit ,Unit_id,Quantity,Price,Pre_Quantity,Pre_Value,Pre_Percent,Curr_Quantity,Curr_value,curr_Percent,tot_quantity,tot_value,tot_percent "
        StrSQL = StrSQL + " from dbo.SubcontractorContract2 "
        StrSQL = StrSQL + " Where bill_id =" & Me.txtid.Text
    
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or maa_rs.EOF) Then
            'Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            'Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
    
            RsDev.MoveFirst
    
            With Me.Fg_Journal
                .Rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("item")) = IIf(IsNull(RsDev("item").value), "", RsDev("item").value)
    
                    .TextMatrix(i, .ColIndex("item_id")) = IIf(IsNull(RsDev("item_id").value), "", RsDev("item_id").value)
            
                    .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("cost").value), "", RsDev("cost").value)
            
                    .TextMatrix(i, .ColIndex("exe")) = IIf(IsNull(RsDev("exe").value), "", RsDev("exe").value)
           
                    .TextMatrix(i, .ColIndex("percentage")) = IIf(IsNull(RsDev("percentage").value), "", RsDev("percentage").value)
        
                    .TextMatrix(i, .ColIndex("exedate")) = IIf(IsNull(RsDev("exedate").value), "", RsDev("exedate").value)
                    
                    
                          .TextMatrix(i, .ColIndex("Unit")) = IIf(IsNull(RsDev("item_unit").value), "", RsDev("item_unit").value)
                           .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(RsDev("Quantity").value), "", RsDev("Quantity").value)
                            .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), "", RsDev("Price").value)
                             .TextMatrix(i, .ColIndex("Pre_Quantity")) = IIf(IsNull(RsDev("Pre_Quantity").value), "", RsDev("Pre_Quantity").value)
                              .TextMatrix(i, .ColIndex("Pre_Value")) = IIf(IsNull(RsDev("Pre_Value").value), "", RsDev("Pre_Value").value)
                              .TextMatrix(i, .ColIndex("Pre_Percent")) = IIf(IsNull(RsDev("Pre_Percent").value), "", RsDev("Pre_Percent").value)
                              
                          
                            .TextMatrix(i, .ColIndex("Curr_Quantity")) = IIf(IsNull(RsDev("Curr_Quantity").value), "", RsDev("Curr_Quantity").value)
                            .TextMatrix(i, .ColIndex("Curr_value")) = IIf(IsNull(RsDev("Curr_value").value), "", RsDev("Curr_value").value)
                            .TextMatrix(i, .ColIndex("curr_Percent")) = IIf(IsNull(RsDev("curr_Percent").value), "", RsDev("curr_Percent").value)
                 .TextMatrix(i, .ColIndex("tot_quantity")) = IIf(IsNull(RsDev("tot_quantity").value), "", RsDev("tot_quantity").value)
          
                 .TextMatrix(i, .ColIndex("tot_value")) = IIf(IsNull(RsDev("tot_value").value), "", RsDev("tot_value").value)
                 .TextMatrix(i, .ColIndex("tot_percent")) = IIf(IsNull(RsDev("tot_percent").value), "", RsDev("tot_percent").value)
                
                    
        
                    RsDev.MoveNext
                Next i

                'Me.txt_total_sum.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
                '  Me.txt_sub_discount.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("discount"), .Rows - 1, .ColIndex("discount"))
                '    Me.txt_sub_net.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("net"), .Rows - 1, .ColIndex("net"))
           
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
                '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
                '  .Rows - 1, .ColIndex("CreditValue"))
                '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
                '  .Rows - 1, .ColIndex("DebitValue"))
            End With

        End If

    End If

    '-----------------------------------------------------------------------------
    'XPTxtCurrent.Caption = Rs.AbsolutePosition
    'XPTxtCount.Caption = Rs.RecordCount
GET_PROJECT_DATA
    ReLineGrid
    Exit Sub
ErrTrap:

End Sub

Private Sub Accredit_Click()
 Dim BeginTrans As Boolean
If val(txtid.Text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
 
    SendTopost Me.Name, "SubcontractorContract", "id", 0, val(dcBranch.BoundText), val(txtid.Text), TxtNoteSerial1.Text, val(Me.note_id)
    
  rs.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
    Retrive (val(Me.txtid.Text))


C1Elastic3.Visible = True
End Sub

Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.txtid.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
Accredit.Enabled = False
Else
Accredit.Enabled = True
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
End If
 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.Rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label24.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label24.Caption = "Approved"
                                 End If
                            Label24.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label24.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label24.Caption = "Currently required Approve"
                            End If
                 Label24.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.Rows = 1
    End If
RsDetails.Close

End Function




Private Sub advancedPayment_Change()
calcnet
End Sub

Private Sub ALLButton1_Click()
Frame15.Visible = True
If Me.TxtModFlg = "R" Then

Frame15.Enabled = False
Exit Sub
Else
 
Frame15.Enabled = True
End If
 VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
 VSFlexGrid4.Rows = 2
TxtPreBalaNet.Text = 0
TxtPreBalaValue.Text = 0
TxtPreBalaVAT.Text = 0
TxtPreBalaTotal.Text = 0
TxtPreBalaPayed.Text = 0
TxtPreBalaRemain.Text = 0
TxtPreBalaTransPyed.Text = 0
advancedPayment.Text = 0
BillCustomer
ClcalteOpiningBalance
End Sub
Sub ClcalteOpiningBalance()
Dim Valu As Double
Dim Percetage As Double
Dim Percetage2 As Double
Dim RecDate As Date
 GetBalanceProject Valu, RecDate
 TxtPreBalaTotal.Text = Round(Valu, Decimal_Places)
 If val(TxtPreBalaTotal.Text) > 0 Then
 PercentgValueAddedAccount_Transec RecDate, 6, 1, , Percetage
 TxtPreBalaVATYu.Text = Percetage
     If Percetage <> 0 Then
   Percetage2 = Percetage / 100 + 1
   TxtPreBalaValue.Text = Round(val(TxtPreBalaTotal.Text) / Percetage2, Decimal_Places)
   TxtPreBalaVAT.Text = Round(val(TxtPreBalaValue.Text) * Percetage / 100, Decimal_Places)
   Else
     TxtPreBalaValue.Text = Round(val(TxtPreBalaTotal.Text), Decimal_Places)
   TxtPreBalaVAT.Text = 0
   End If
 End If
TxtPreBalaPayed.Text = Round(GetValue(), Decimal_Places)
TxtPreBalaRemain.Text = Round(val(TxtPreBalaTotal.Text) - val(TxtPreBalaPayed.Text), Decimal_Places)
End Sub
Sub BillCustomer(Optional Ind As Integer = 0)

Exit Sub
Dim Msg As String
Dim mIsFoundRow As Boolean
If VSFlexGrid4.Rows > 1 Then
    If val(VSFlexGrid4.TextMatrix(1, VSFlexGrid4.ColIndex("NoteSerial1"))) <> 0 Then
        mIsFoundRow = True
    Else
        mIsFoundRow = False
        
    End If
End If

    If val(TXTEnd_user_id.Text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ ≈Œ Ì«— «·⁄„Ì· «Ê·«"
        Else
            MsgBox "Please Select Customer"
        End If
        Exit Sub
    Else
        If Ind = 0 Then
            Frame15.Visible = True
        End If
        If Me.TxtModFlg.Text <> "R" Then
           VSFlexGrid4.Enabled = True
        Else
         '   VSFlexGrid4.Enabled = False
        End If
            If Ind = 0 Then
            'XPTxtVal.Text = 0
            End If
            If Me.TxtModFlg.Text = "N" Then
                If val(billto.ListIndex) = 0 Then
                    RetriveBillBuy val(TXTEnd_user_id.Text)
                ElseIf val(billto.ListIndex) = 1 Then
                    RetriveBillBuy val(DcbosubContractor.BoundText)
                    
                End If
            End If
            If Me.TxtModFlg.Text = "E" And (FlgBillBuy = True Or Not mIsFoundRow) Then
     '           VSFlexGrid4.Editable = True
            Else
             '   VSFlexGrid4.Editable = False
            End If
                If val(billto.ListIndex) = 0 Then
                    RetriveBillBuy val(TXTEnd_user_id.Text)
                ElseIf val(billto.ListIndex) = 1 Then
                    RetriveBillBuy val(DcbosubContractor.BoundText)
                    
                End If

            
        End If
    
End Sub
Sub RetriveBillBuy(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset
With VSFlexGrid4
.Clear flexClearScrollable, flexClearEverything
.Rows = 1
End With
sql = " SELECT     dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.Notes.ManulaNO, dbo.Notes.NoteID, dbo.Notes.branch_no, "
sql = sql & "                       dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
sql = sql & "                       dbo.Notes.NCashingType , dbo.TblCustemers.fullcode, dbo.Notes.Note_Value, dbo.Notes.vat, dbo.Notes.totalPayed"
sql = sql & "  FROM         dbo.Notes LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
sql = sql & "  Where (dbo.Notes.NCashingType = 3) and dbo.Notes.CusID=" & CuID & " and (dbo.Notes.totalPayed=0 or  dbo.Notes.totalPayed is null)"
sql = sql & "  ORDER BY dbo.Notes.NoteSerial1"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
'VSFlexGrid4.Enabled = True
With VSFlexGrid4
.Clear flexClearScrollable, flexClearEverything
.Rows = 1
.Rows = .Rows + Rs8.RecordCount
.Rows = .FixedRows + Rs8.RecordCount
 Rs8.MoveFirst
For i = .FixedRows To Rs8.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("branch_no").value), 0, Rs8("branch_no").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If

.TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("NoteID").value), 0, Rs8("NoteID").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(Rs8("Note_Value").value), 0, Rs8("Note_Value").value)
.TextMatrix(i, .ColIndex("VAT")) = IIf(IsNull(Rs8("VAT").value), 0, Rs8("VAT").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("NoteDate").value), "", Rs8("NoteDate").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("ManulaNO").value), "", Rs8("ManulaNO").value)
.TextMatrix(i, .ColIndex("Note_Value")) = val(IIf(IsNull(Rs8("Note_Value").value), 0, Rs8("Note_Value").value)) + val(IIf(IsNull(Rs8("VAT").value), 0, Rs8("VAT").value))
If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteBillBuy(val(.TextMatrix(i, .ColIndex("NoteID"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If
End Sub
Function GeteBillBuy(Optional Transaction_ID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblProjePayPrePayed"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillBuy = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillBuy = 0
End If
End Function
Public Sub RetriveBillBuyData(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Double
    Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
  StrSQL = "   SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblPayPrePayed.*"
  StrSQL = StrSQL & "  FROM         dbo.TblPayPrePayed LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblPayPrePayed.branch_no = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & "  Where (dbo.TblPayPrePayed.NoteID1 = " & val(txtid.Text) & ")"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid4
    .Clear flexClearScrollable, flexClearEverything
    .Rows = .FixedRows
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To RsDetails.RecordCount
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("VATLine")) = IIf(IsNull(RsDetails("VATLine").value), 0, RsDetails("VATLine").value)
            .TextMatrix(i, .ColIndex("ValueLine")) = IIf(IsNull(RsDetails("ValueLine").value), 0, RsDetails("ValueLine").value)
            .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
            End If
            .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsDetails("NoteID").value), 0, RsDetails("NoteID").value)
            .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDetails("NoteSerial1").value), 0, RsDetails("NoteSerial1").value)
            .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsDetails("Note_Value").value), 0, RsDetails("Note_Value").value)
            .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
            .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
            .TextMatrix(i, .ColIndex("too")) = IIf(IsNull(RsDetails("too").value), "", RsDetails("too").value)
            .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
            .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(((RsDetails("NoteDate").value))), "", ((RsDetails("NoteDate").value)))
            .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
            RsDetails.MoveNext
        Next i
    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub
Sub DeleteBillBuy()
'Dim i As Integer
'Dim StrSQL As String
'With VSFlexGrid4
' For i = .FixedRows To .Rows - 1
' If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
'      StrSQL = "Update Notes Set  TotalPayed=0 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
'                Cn.Execute StrSQL, , adExecuteNoRecords
'     End If
'     Next i
' End With
End Sub
 Function saveBillBuy()
'    Dim StrSQL As String
'    Dim i As Double
'    Dim Diff As Double
'    Dim Note_Value1 As Double
'    Diff = 0
'Dim RsDetails As ADODB.Recordset
'      If Me.TxtModFlg.Text = "E" Then
'    StrSQL = "Delete From TblPayPrePayed Where NoteID1=" & val(txtid.Text)
'    Cn.Execute StrSQL, , adExecuteNoRecords
'        StrSQL = "Delete From TblProjePayPrePayed Where  NoteID=" & val(txtid.Text)
'    Cn.Execute StrSQL, , adExecuteNoRecords
'    End If
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     * from dbo.TblPayPrePayed Where (1 = -1)"
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    With VSFlexGrid4
'    TxtValueTemp.Text = val(Label47.Caption)
'
'
'
'    For i = .FixedRows To .Rows - 1
'        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > 0 Then
''Dim LineDiscountPercent As Double
''        LineDiscountPercent = val(.TextMatrix(i, .ColIndex("NetExe"))) / Results.Text
'            RsDetails.AddNew
'
'            RsDetails("NoteID1").value = val(txtid.Text)
'            RsDetails("VATLine").value = val(.TextMatrix(i, .ColIndex("VATLine")))
'            RsDetails("ValueLine").value = val(.TextMatrix(i, .ColIndex("ValueLine")))
'            RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
'            RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
'            RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
'            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
'            Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
'            Diff = 0
'            If val(TxtValueTemp.Text) > 0 Then
'          If val(TxtValueTemp.Text) <= Note_Value1 Then
'          Diff = val(TxtValueTemp.Text)
'          TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
'          Else
'          Diff = Note_Value1
'          TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
'          End If
'          End If
'            .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
'            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
'            RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
'            RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
'            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
'           .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
'            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
'            RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
'            RsDetails.update
'            If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
'            StrSQL = "Update Notes Set  TotalPayed=1 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
'                Cn.Execute StrSQL, , adExecuteNoRecords
'             Else
'                 StrSQL = "Update Notes Set  TotalPayed=Null Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
'                Cn.Execute StrSQL, , adExecuteNoRecords
'            End If
'      End If
'    Next i
'End With
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     * from dbo.TblProjePayPrePayed Where (1 = -1)"
'   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    With VSFlexGrid4
'    For i = .FixedRows To .Rows - 1
'        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) Then
'            RsDetails.AddNew
'            RsDetails("NoteID").value = val(txtid.Text)
'            RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
'            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
'            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
'            RsDetails.update
'        End If
'    Next i
'End With

End Function

Private Sub ALLButton2_Click()
FillAllBandsToGrid
End Sub

Private Sub ALLButton3_Click()
Dim LngNewRow  As Long
Dim Row As Long
Dim Rs1 As New ADODB.Recordset

       Dim netexe As Double
       Dim QtyExe As Double
    Dim VATPer  As Double
    Dim oldPerforValue  As Double
    Dim discountHasmyat As Double
    
'LngNewRow = val(ModFgLib.SetFgForNewRow(Fg_Journal, Fg_Journal.ColIndex("project_id")))
If Fg_Journal.Rows > 1 Then
    If val(Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("project_id"))) = 0 Then
        LngNewRow = 1
    Else
        LngNewRow = val(SetFgForNewRow(Fg_Journal, Fg_Journal.ColIndex("project_id")))
    End If
Else
LngNewRow = 1
End If

With Fg_Journal

.Rows = .Rows + 1

.TextMatrix(LngNewRow, .ColIndex("PrMainDesID")) = cmbBand.BoundText
.TextMatrix(LngNewRow, .ColIndex("MainDes")) = cmbBand.Text

.TextMatrix(LngNewRow, .ColIndex("Period")) = TxtPeriod(1).Text


'Fg_Journal_AfterEdit LngNewRow, .ColIndex("MainDes")
.TextMatrix(LngNewRow, .ColIndex("item")) = cmbBand.Text
.TextMatrix(LngNewRow, .ColIndex("oprid")) = cmbBand.BoundText
'Fg_Journal_AfterEdit LngNewRow, .ColIndex("item")
.TextMatrix(LngNewRow, .ColIndex("PrMainDesID")) = cmbBand.BoundText
.TextMatrix(LngNewRow, .ColIndex("MainDes")) = cmbBand.Text


.TextMatrix(LngNewRow, .ColIndex("project_id")) = DataCombo2.BoundText
.TextMatrix(LngNewRow, .ColIndex("FullCode")) = DataCombo2.Text
.TextMatrix(LngNewRow, .ColIndex("Projectname")) = txtprojectname


Row = LngNewRow

sql = sql & " FROM         dbo.ProjectMainDes LEFT OUTER JOIN"
sql = sql & "                      dbo.projects_des ON dbo.ProjectMainDes.ID = dbo.projects_des.PrMainDesID"

 StrSQL = "SELECT  projects_des.qtySubContractor,projects_des.costSubContractor,projects_des.PrMainDesID, ProjectMainDes.Name, dbo.projects_des.TotalExe, dbo.projects_des.QtyExe, dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, dbo.projects_des.cost, "
              StrSQL = StrSQL & "        dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id,"
              StrSQL = StrSQL & "        dbo.projects_des.esQty, dbo.projects_des.Remark, dbo.projects_des.fullcode, dbo.projects_des.PandUnitID, dbo.TblProcessUnites.UnitName,"
              StrSQL = StrSQL & "        dbo.TblProcessUnites.UnitNamee"
              StrSQL = StrSQL & " FROM         dbo.projects_des LEFT OUTER JOIN"
              StrSQL = StrSQL & "               dbo.TblProcessUnites ON dbo.projects_des.PandUnitID = dbo.TblProcessUnites.UnitID"
              
              StrSQL = StrSQL & "     LEFT OUTER JOIN      ProjectMainDes On dbo.ProjectMainDes.ID = dbo.projects_des.PrMainDesID"
              
              StrSQL = StrSQL & "  WHERE dbo.projects_des.oprid='" & .TextMatrix(LngNewRow, .ColIndex("oprid")) & "'"
                    StrSQL = StrSQL & " and dbo.projects_des.project_id =" & val(DataCombo2.BoundText)
                    Set Rs1 = New ADODB.Recordset
                  '  StrSQL = "SELECT   line_no, oprid,des, net, project_id  ,[unit] ,[Quantity],[Price] ,[Pre_Quantity] ,[Pre_Value],[Pre_Percent] ,[Curr_Quantity]  ,[Curr_value] ,[curr_Percent] ,[tot_quantity] ,[tot_value] ,[tot_percent]   from dbo.projects_des  WHERE fullcode='" & .ComboData & "'"  ' project_id =" & Val(DataCombo2.BoundText) & "and line_no=" & Val(.ComboItem)
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    If Not Rs1.EOF Then
                      .TextMatrix(Row, .ColIndex("qty")) = IIf(IsNull(Rs1("qty").value), 0, Rs1("qty").value)
                      If Me.ChQty.value = vbChecked Then
                      .TextMatrix(Row, .ColIndex("quntExc")) = val(.TextMatrix(Row, .ColIndex("qty")))
                      End If
                      
                  .TextMatrix(Row, .ColIndex("PrMainDesID")) = Rs1!PrMainDesID & ""
                    .TextMatrix(Row, .ColIndex("MainDes")) = Rs1!Name & ""
                    
                      .TextMatrix(Row, .ColIndex("cost")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
        
                      
                         .TextMatrix(Row, .ColIndex("qtySubContractor")) = IIf(IsNull(Rs1("qtySubContractor").value), 0, Rs1("qtySubContractor").value)
                      .TextMatrix(Row, .ColIndex("costSubContractor")) = IIf(IsNull(Rs1("costSubContractor").value), 0, Rs1("costSubContractor").value)
                      
                      If val(.TextMatrix(Row, .ColIndex("qtySubContractor"))) = 0 Then
                      .TextMatrix(Row, .ColIndex("qtySubContractor")) = .TextMatrix(Row, .ColIndex("qty"))
                      End If
                      
                      
                     If val(.TextMatrix(Row, .ColIndex("costSubContractor"))) = 0 Then
                      .TextMatrix(Row, .ColIndex("costSubContractor")) = .TextMatrix(Row, .ColIndex("exe"))
                      End If
                                  
                                  
                                  
                      .TextMatrix(Row, .ColIndex("total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("cost")))
                      .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
                      
               If billto.ListIndex = 0 Then
                      .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
                   Else
                   .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("costSubContractor").value), 0, Rs1("costSubContractor").value)
                   End If
                   
                      If SystemOptions.UserInterface = ArabicInterface Then
                      .TextMatrix(Row, .ColIndex("Unit")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                      Else
                      .TextMatrix(Row, .ColIndex("Unit")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
                      End If
                      .TextMatrix(Row, .ColIndex("unit_id")) = IIf(IsNull(Rs1("PandUnitID").value), 0, Rs1("PandUnitID").value)
                      .TextMatrix(Row, .ColIndex("item_id")) = IIf(IsNull(Rs1("fullcode").value), "", Rs1("fullcode").value)
                       .TextMatrix(Row, .ColIndex("discount")) = IIf(IsNull(Rs1("discount").value), 0, Rs1("discount").value)
                       .TextMatrix(Row, .ColIndex("net")) = IIf(IsNull(Rs1("net").value), 0, Rs1("net").value)
                      .TextMatrix(Row, .ColIndex("percentage")) = (val(.TextMatrix(Row, .ColIndex("qty"))) - val(.TextMatrix(Row, .ColIndex("quntExc")))) / 100
                     '  .TextMatrix(Row, .ColIndex("unit")) = IIf(IsNull(Rs1("unit").value), 0, Rs1("unit").value)
                     '   .TextMatrix(Row, .ColIndex("Quantity")) = IIf(IsNull(Rs1("Quantity").value), 0, Rs1("Quantity").value)
                        .TextMatrix(Row, .ColIndex("Pre_Value")) = IIf(IsNull(Rs1("TotalExe").value), 0, Rs1("TotalExe").value)
                        .TextMatrix(Row, .ColIndex("Pre_Quantity")) = IIf(IsNull(Rs1("QtyExe").value), 0, Rs1("QtyExe").value)
                          If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                          GetTermsTotals val(.TextMatrix(Row, .ColIndex("oprid"))), val(txtid.Text), XPDtbTrans.value, netexe, QtyExe, TxtModFlg.Text
                         End If
                         If SystemOptions.AllowNoRoudProjectInvoices = True Then
                         .TextMatrix(Row, .ColIndex("Pre_Quantity")) = val(.TextMatrix(Row, .ColIndex("Pre_Quantity"))) + Round(QtyExe, val(cCompanyInfo.NoRoudProjectInvoices))
                         .TextMatrix(Row, .ColIndex("Pre_Value")) = .TextMatrix(Row, .ColIndex("Pre_Value")) + Round(netexe, val(cCompanyInfo.NoRoudProjectInvoices))
                         Else
                         .TextMatrix(Row, .ColIndex("Pre_Quantity")) = val(.TextMatrix(Row, .ColIndex("Pre_Quantity"))) + Round(QtyExe, 2)
                         .TextMatrix(Row, .ColIndex("Pre_Value")) = .TextMatrix(Row, .ColIndex("Pre_Value")) + Round(netexe, 2)
                         End If
                     '      .TextMatrix(Row, .ColIndex("Pre_Value")) = IIf(IsNull(Rs1("Pre_Value").value), 0, Rs1("Pre_Value").value)
                     '       .TextMatrix(Row, .ColIndex("Pre_Percent")) = IIf(IsNull(Rs1("Pre_Percent").value), 0, Rs1("Pre_Percent").value)
                     '        .TextMatrix(Row, .ColIndex("Curr_Quantity")) = IIf(IsNull(Rs1("Curr_Quantity").value), 0, Rs1("Curr_Quantity").value)
                     ' .TextMatrix(Row, .ColIndex("Curr_value")) = IIf(IsNull(Rs1("Curr_value").value), 0, Rs1("Curr_value").value)
                     ' .TextMatrix(Row, .ColIndex("curr_Percent")) = IIf(IsNull(Rs1("curr_Percent").value), 0, Rs1("curr_Percent").value)
                     ' .TextMatrix(Row, .ColIndex("tot_quantity")) = IIf(IsNull(Rs1("tot_quantity").value), 0, Rs1("tot_quantity").value)
                     ' .TextMatrix(Row, .ColIndex("tot_value")) = IIf(IsNull(Rs1("tot_value").value), 0, Rs1("tot_value").value)
                     ' .TextMatrix(Row, .ColIndex("tot_percent")) = IIf(IsNull(Rs1("tot_percent").value), 0, Rs1("tot_percent").value)
            End If
            
            If val(txtQty.Text) <> 0 Then
                .TextMatrix(LngNewRow, .ColIndex("qty")) = txtQty.Text
            End If
            If val(TxtCost.Text) <> 0 Then
                .TextMatrix(LngNewRow, .ColIndex("cost")) = TxtCost.Text
            End If
            If val(txtqtySubContractor.Text) <> 0 Then
                .TextMatrix(LngNewRow, .ColIndex("qtySubContractor")) = txtqtySubContractor.Text
            End If
            If val(txtcostSubContractor.Text) <> 0 Then
                .TextMatrix(LngNewRow, .ColIndex("costSubContractor")) = txtcostSubContractor.Text
            End If
                  '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    CalCultePers Row
        
ReLineGrid






CalCultePers LngNewRow
calcnet
ReLineGrid
End With


End Sub

Private Sub billto_Change()

ClculteVAT
ReLineGrid
ReLineGrid
End Sub

Private Sub billto_Click()
billto_Change
End Sub

Private Sub cboDiscount1_Change()
calcnet
End Sub

Private Sub cboDiscount1_Click()
cboDiscount1_Change
End Sub

Private Sub Check1_Click()
    Dim i As Integer

    If Check1.value = vbChecked Then

        With Me.VSFlexGrid4
 
            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid4

            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    RelineBuy
End Sub

Private Sub cmbBand_Change()
Dim StrSQL  As String
Dim Rs1 As New ADODB.Recordset
        StrSQL = "SELECT  projects_des.qtySubContractor,projects_des.costSubContractor,  dbo.projects_des.TotalExe, dbo.projects_des.QtyExe, dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, dbo.projects_des.cost, "
              StrSQL = StrSQL & "        dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id,"
              StrSQL = StrSQL & "        dbo.projects_des.esQty, dbo.projects_des.Remark, dbo.projects_des.fullcode, dbo.projects_des.PandUnitID, dbo.TblProcessUnites.UnitName,"
              StrSQL = StrSQL & "        dbo.TblProcessUnites.UnitNamee"
              StrSQL = StrSQL & " FROM         dbo.projects_des LEFT OUTER JOIN"
              StrSQL = StrSQL & "               dbo.TblProcessUnites ON dbo.projects_des.PandUnitID = dbo.TblProcessUnites.UnitID"
              StrSQL = StrSQL & "  WHERE dbo.projects_des.oprid='" & cmbBand.BoundText & "'"
                    StrSQL = StrSQL & " and dbo.projects_des.project_id =" & val(DataCombo2.BoundText)
                    Set Rs1 = New ADODB.Recordset
                  '  StrSQL = "SELECT   line_no, oprid,des, net, project_id  ,[unit] ,[Quantity],[Price] ,[Pre_Quantity] ,[Pre_Value],[Pre_Percent] ,[Curr_Quantity]  ,[Curr_value] ,[curr_Percent] ,[tot_quantity] ,[tot_value] ,[tot_percent]   from dbo.projects_des  WHERE fullcode='" & .ComboData & "'"  ' project_id =" & Val(DataCombo2.BoundText) & "and line_no=" & Val(.ComboItem)
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    If Not Rs1.EOF Then
                      txtQty.Text = IIf(IsNull(Rs1("qty").value), 0, Rs1("qty").value)
                     
                      
                        
                      TxtCost.Text = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
        
                      
                         txtqtySubContractor.Text = IIf(IsNull(Rs1("qtySubContractor").value), 0, Rs1("qtySubContractor").value)
                      txtcostSubContractor.Text = IIf(IsNull(Rs1("costSubContractor").value), 0, Rs1("costSubContractor").value)
                 End If
End Sub

Private Sub Cmd_Click(Index As Integer)
Frame7.Visible = False
End Sub

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        ''Call Reload(Me)
 
    Else
        my_language = "A"
 
        ''Call Reload(Me)
    End If

End Sub

Function SaveData()
    calcnet
    Dim accountdep As String
Dim PerforVLineDiscount  As Double
    If billto.ListIndex = -1 Then MsgBox "Õœœ «·„” Œ·’ „ﬁœ„  «·Ï „‰ ", vbCritical: Exit Function
        
   If billto.ListIndex = 1 And DcbosubContractor.BoundText = "" Then MsgBox "Õœœ „ﬁ«Ê· «·»«ÿ‰  ", vbCritical: Exit Function
   
    Dim StrSQL As String
     Dim j As Integer
     Dim found As Boolean
     
   j = Fg_Journal.FixedRows
     found = False
     
  For j = Fg_Journal.FixedRows To Fg_Journal.Rows - 1
  If Fg_Journal.TextMatrix(j, Fg_Journal.ColIndex("item")) <> "" Then
        found = True
  End If
Next

If found = False Then

MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„ ›Ï «·›« Ê—… ", vbCritical: Exit Function

End If
Dim ii As Long
  For j = Fg_Journal.FixedRows To Fg_Journal.Rows - 1
  
        For ii = Fg_Journal.FixedRows To Fg_Journal.Rows - 1
            If val(Fg_Journal.TextMatrix(ii, Fg_Journal.ColIndex("oprid"))) = val(Fg_Journal.TextMatrix(j, Fg_Journal.ColIndex("oprid"))) _
            And val(Fg_Journal.TextMatrix(ii, Fg_Journal.ColIndex("project_id"))) = val(Fg_Journal.TextMatrix(ii, Fg_Journal.ColIndex("project_id"))) And ii <> j Then
                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ · ﬂ—«— «·»‰œ ”ÿ— —ﬁ„ " & ii & " „⁄ «·”ÿ— —ﬁ„ " & j, vbCritical: Exit Function
                
            End If
        Next
  If Fg_Journal.TextMatrix(j, Fg_Journal.ColIndex("item")) <> "" Then
        found = True
  End If
Next


        
        
        
        
    If billto.ListIndex = 0 Then
   x = val(TXTEnd_user_id.Text)
        'accountdep = txtendaccount.text
    Else

        If billto.ListIndex = 1 Then
        x = val(TXTsub_contractor_id.Text)
        '    accountdep = txtsubaccount.text
        End If
    End If
x = val(TXTEnd_user_id.Text)
  '  Dim x As Double
  '  x = get_Customer_id(accountdep)
        
    '  total.text = gettotal(txtid.text)
    Dim Rs1 As New ADODB.Recordset
'    StrSQL = "select * From Notes where NoteType=65000 and NoteSerial='" & Me.TxtNoteSerial.Text & "' order by NoteID"
'    Rs1.Open StrSQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
'
    If TxtModFlg.Text = "N" Then
   
        If x = 0 Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "An error in customer Number", vbCritical: Exit Function
            Else
                MsgBox "ÌÊÃœ Œÿ√ ›Ì —ﬁ„ «·⁄„Ì·", vbCritical: Exit Function
            End If
        End If
          note_id.Text = CStr(new_id("Notes", "NoteID", "", True))
            txtid.Text = CStr(new_id("SubcontractorContract", "id", "", True))
            
        rs.AddNew

    Else
'        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(note_id.Text)
'        Cn.Execute StrSQL, , adExecuteNoRecords
'        StrSQL = "Delete From notes  Where NoteID=" & val(note_id.Text)
'        Cn.Execute StrSQL, , adExecuteNoRecords
'
'        StrSQL = "Delete From SubcontractorContract2 Where bill_id=" & val(Me.txtid.Text)
'        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
'
'    Rs1.AddNew
' 'branch_id
''     If TxtNoteSerial1.Text = "" Then
''     TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbTrans.value, 65, 65, , , , , val(billto.ListIndex))
''     End If
'
'    Rs1("branch_no").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'    Rs1("NoteID").value = val(note_id.Text)
'    Rs1("Note_Value").value = IIf(Total.Text = "", Null, val(Total.Text))
'    Rs1("CusID").value = X
'    Rs1("NoteType").value = 500
'    Rs1("NoteType").value = 65000
'    Rs1("NoteDate").value = XPDtbTrans.value
'    Rs1("UserID").value = user_id
'   Rs1("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", Trim(TxtNoteSerial1.Text), Null)
'
'   Rs1("RemarkE").value = IIf(Me.TxtRemarks <> "", Trim(TxtRemarks.Text), Null)
'   Rs1("Remark").value = IIf(Me.TxtRemarks <> "", Trim(TxtRemarks.Text), Null)
'
  '  rs("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", Trim(TxtNoteSerial1.Text), Null)
    rs("ExPercen").value = val(TxtCost.Text)
    rs("ExPercenID").value = val(DcbExPercen.ListIndex)
    rs("PreVAT").value = val(TxtPreVAT.Text)
    rs("FATYou").value = val(TxtFATYou.Text)
    rs("FATValue").value = val(TxtFATValue.Text)
    rs("TotalValue").value = val(TxtTotalValue.Text)
    rs("AccountCodeVat").value = Me.AccountVat.BoundText
    rs("NetValue").value = val(TxtNetValue.Text)
    rs("PerforValue").value = val(TxtPerforValue.Text)
    ''/////////
    rs("StartDateProje").value = StartDateProje.value
    rs("PreBalaValue").value = val(TxtPreBalaValue.Text)
    rs("PreBalaVAT").value = val(TxtPreBalaVAT.Text)
    rs("PreBalaTotal").value = val(TxtPreBalaTotal.Text)
    rs("PreBalaPayed").value = val(TxtPreBalaPayed.Text)
    rs("PreBalaRemain").value = val(TxtPreBalaRemain.Text)
    rs("PreBalaTransPyed").value = val(TxtPreBalaTransPyed.Text)
    rs("PreBalaNet").value = val(TxtPreBalaNet.Text)
    rs("PreBalaVATYu").value = val(TxtPreBalaVATYu.Text)
    rs("SumVATLine").value = val(Label57.Caption)
    rs("SumValueLine").value = val(Label56.Caption)
    ''/////
' If Option7.value = True Then
'  rs("UnderImp").value = 0
' ElseIf Option6.value = True Then
'  rs("UnderImp").value = 1
' ElseIf Option8.value = True Then
'  rs("UnderImp").value = 2
'End If
If TxtManualNO.Text = "" Then
TxtManualNO.Text = TxtNoteSerial1.Text
End If

'    If SystemOptions.UserInterface = ArabicInterface Then
'        Rs1("remark").value = "„” Œ·’ —ﬁ„  :  " & txtManualNo & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text
'    Else
'        Rs1("remark").value = "  Project Invoice No  :  " & txtManualNo & CHR(13) & "  To Project " & TXTprojectname.Text
'    End If
'
'    '   Rs1("remark").value = "„” Œ·’ —ﬁ„ :     " & txtid & "    " & Chr(13) & "  ··„‘—Ê⁄  " & txtprojectname.text
''
''    If TxtNoteSerial = "" Then
''        TxtNoteSerial.Text = Notes_coding(val(my_branch), XPDtbTrans.value)
''    End If
''
'
'
'    Rs1("NoteSerial").value = txtid
'
'  '  Rs1("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.Text) '„”·”·
'    Rs1("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
'    '  rs("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
'
'    Rs1("sanad_year").value = year(XPDtbTrans.value)
'    Rs1("sanad_month").value = Month(XPDtbTrans.value)
'    Rs1("note_value_by_characters").value = WriteNo(Format(Me.Results.Text, "0.00"), 0, True, ".")
'
'    Rs1.update
    
   ' rs("id").value = val(Me.txtid.Text)
    
    rs("bill_date").value = XPDtbTrans.value
  'branch_id
    rs("branch_no").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
    rs("project_no").value = IIf(Not IsNumeric(DataCombo2.BoundText), "", DataCombo2.BoundText)
    rs("project_name").value = txtprojectname.Text
'    rs("Sub_user_name").value = IIf(IsNull(DcAccount1.Text), "", DcAccount1.Text)
    rs("End_user_name").value = IIf(IsNull(DcAccount2.Text), "", DcAccount2.Text)
    rs("End_user_account").value = IIf(IsNull(txtendaccount.Text), "", txtendaccount.Text)
    'rs("Sub_user_account").value = IIf(IsNull(txtsubaccount.Text), "", txtsubaccount.Text)
    rs("revenue_account").value = IIf(IsNull(txtrevenue_account.Text), "", txtrevenue_account.Text)
    rs("UserID").value = IIf(DCboUserName.BoundText <> "", val((DCboUserName.BoundText)), Null)
    rs("bill_to").value = billto.ListIndex
    rs("bill_type").value = bill_Type.ListIndex '  IIf(IsNull(bill_Type.text), "", bill_Type.text)
    rs("note_id").value = IIf(IsNull(note_id.Text), "", note_id.Text)
    rs("NoteSerial").value = IIf(IsNull(TxtNoteSerial.Text), "", TxtNoteSerial.Text)
    rs("total").value = IIf(Not IsNumeric(total.Text), 0, total.Text)
'    rs("AccountUnderImp").value = TxtAccountUnderImp.Text
    
    
    '26082015
    'rs("Discount").value = IIf(Not IsNumeric(TxtDiscount.Text), 0, TxtDiscount.Text)
    'rs("AdvancedPayment").value = IIf(Not IsNumeric(advancedPayment.Text), 0, advancedPayment.Text)
    
    rs("Results").value = IIf(Not IsNumeric(Results.Text), 0, Results.Text)
   ''///////23 05 2016
  rs("BillNo").value = val(TxtBillNo.Text)
  'rs("StartDate").value = startDate.value
  
  rs("PeriodType").value = val(DcbPeriodType.ListIndex)
 ' rs("Remarks2").value = TxtRemarks2.Text
  
 
'26082015


         rs("dueDate").value = dueDate.value
'rs("dueDate1").value = dueDate1.value


'*************************************************
rs("subContractorId").value = IIf(Not IsNumeric(DcbosubContractor.BoundText), Null, DcbosubContractor.BoundText)
'rs("discount1ID").value = val(cboDiscount1.ListIndex)
'rs("discount2ID").value = val(cboDiscount2.ListIndex)
'rs("discount1value").value = val(txtDiscount1.Text)
'rs("discount2value").value = val(txtDiscount2.Text)
'rs("Remarks").value = Trim(TxtRemarks.Text)
'rs("ManualNo").value = Trim(txtManualNo.Text)

 
'*************************************************
If val(Me.TxtBillNo.Text) > 0 Then
'SaveBillMonthly
End If

    rs.update

    
'    Set RsDev = New ADODB.Recordset
' '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'               StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
'   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'   Dim LngDevID As Long
'  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'  accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", X, "Account_code")
'
' Dim Posted As Integer
'            If CheckAprroveScreen(Me.Name) = True Then
'            Posted = 1
'            Else
'            Posted = 0
'            End If
'
'
'If billto.ListIndex = 0 Then
'  Dim lineno As Integer
'  lineno = 1
''    If accountdep = "" Then GoTo ll
'    '«·ÿ—› «·„œÌ‰
'    RsDev.AddNew
'
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'If Option8.value = True Then
'    RsDev("Account_Code").value = Me.TxtAccountUnderImp.Text
' Else
'    RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
'End If
'        RsDev("Value").value = val(Me.Total.Text) + val(TxtFATValue.Text)
'    RsDev("Credit_Or_Debit").value = 0
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & " Manual #   " & txtManualNo & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'     RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'
'    RsDev.update
''ll:
'lineno = lineno + 1
'
''«·Õ”„Ì« 
''Account_Code_dynamic1
'If val(Me.TxtDiscount.Text) > 0 Then
'    RsDev.AddNew
'
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Account_Code").value = Account_Code_dynamic1 '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
'    RsDev("Value").value = val(Me.TxtDiscount.Text)
'    RsDev("Credit_Or_Debit").value = 0
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & "  Manual# " & txtManualNo & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'
'   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'
'
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'   RsDev("Posted").value = Posted
'    RsDev.update
''ll:
'lineno = lineno + 1
'
'End If
'If val(Me.TxtPerforValue.Text) > 0 Then
'    RsDev.AddNew
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Account_Code").value = AcountGood ' get_account_code_branch(152, my_branch) 'Õ”«» Õ”‰ «·«œ«¡
'    RsDev("Value").value = val(Me.TxtPerforValue.Text)
'    RsDev("Credit_Or_Debit").value = 0
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & " Manual#  " & txtManualNo & CHR(13) & TxtRemarks.Text
'    End If
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'    RsDev.update
''ll:
''lineno = lineno + 1
''
''    RsDev.AddNew
''    RsDev("branch_id").value = IIf(Trim$(Me.DcBranch.BoundText) = "", Null, Trim$(Me.DcBranch.BoundText))
''    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
''    RsDev("DEV_ID_Line_No").value = lineno
''    RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
''    RsDev("Value").value = val(Me.TxtPerforValue.Text)
''    RsDev("Credit_Or_Debit").value = 1
''    If SystemOptions.UserInterface = ArabicInterface Then
''        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & txtid & Chr(13) & "  ··„‘—Ê⁄ " & txtprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNo
''    Else
''        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & txtid & Chr(13) & "  To Project " & txtprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNo
''    End If
''    RsDev("Notes_ID").value = val(note_id.Text)
''    RsDev("project_bill_no").value = val(txtid.Text)
''   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
''    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
''    RsDev("UserID").value = user_id
''    RsDev("branch_id").value = my_branch
''    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
''    RsDev.update
''ll:
'lineno = lineno + 1
'
'End If
'
'
''«·œ›⁄«  «·„ﬁœ„…
''Account_Code_dynamic2
'If val(Me.advancedPayment.Text) > 0 Then
'    RsDev.AddNew
'
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Account_Code").value = Account_Code_dynamic2 '    Õ”«» œ›⁄«  „ﬁœ„…
'    RsDev("Value").value = val(Me.advancedPayment.Text)
'    RsDev("Credit_Or_Debit").value = 0
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & "œ›⁄«  „ﬁœ„…" & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & "  Manuall#  " & txtManualNo & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'    RsDev.update
''ll:
'lineno = lineno + 1
'
'  RsDev.AddNew
'
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Account_Code").value = accountdep '    Õ”«» œ›⁄«  „ﬁœ„…
'    RsDev("Value").value = val(Me.advancedPayment.Text)
'    RsDev("Credit_Or_Debit").value = 1
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & "œ›⁄«  „ﬁœ„…" & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & "   Manual   " & txtManualNo & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'     RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'    RsDev.update
''ll:
'lineno = lineno + 1
'
'End If
'
'
''«·«Ì—œ« 
'
'    '«·ÿ—› «·œ«∆‰
'   If Option8.value = False Then
'    If Me.txtrevenue_account.Text = "" Then Exit Function
'
'   Else
'   If (accountdep = "") Then Exit Function
'  End If
'    RsDev.AddNew
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'    RsDev("branch_id").value = my_branch
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
' 'If SystemOptions.Revenueowed = True Then
' If Option8.value = True Then
' RsDev("Account_Code").value = accountdep
' Else
' ''
' RsDev("Account_Code").value = Me.txtrevenue_account.Text ' Account_Code_dynamic1
'
' End If
' '   Else
'    'RsDev("Account_Code").value = Me.txtrevenue_account .text
' '   End If
'
'    RsDev("Value").value = val(Me.Results.Text)  '«·«Ì—«œ« 
'    RsDev("Credit_Or_Debit").value = 1
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
' '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
' Else
'   RsDev("project_id").value = val(Me.DataCombo2.BoundText)
' End If
'
'    RsDev("RecordDate").value = XPDtbTrans.value
'    RsDev("UserID").value = user_id
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'    RsDev.update
'lineno = lineno + 1
''///////////////////
'If Me.AccountVat.BoundText <> "" And val(Me.TxtFATValue.Text) > 0 Then
'    RsDev.AddNew
'    RsDev("Account_Code").value = Me.AccountVat.BoundText
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'    RsDev("branch_id").value = my_branch
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Value").value = val(Me.TxtFATValue.Text)  '«·«Ì—«œ« 
'    RsDev("Credit_Or_Debit").value = 1
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & txtid & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & " Õ”«» «·ﬁÌ„… «·„÷«›…" & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & "Account VAT" & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'    If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
'     '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'     Else
'     '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'     End If
'
'    RsDev("RecordDate").value = XPDtbTrans.value
'    RsDev("UserID").value = user_id
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'    RsDev("Posted").value = Posted
'    RsDev.update
'    lineno = lineno + 1
'''///////////////////////›«  «·œ›⁄«  «·„ﬁœ„…
'    If val(TxtPreVAT.Text) > 0 Then
'        RsDev.AddNew
'        RsDev("Account_Code").value = Me.AccountVat.BoundText
'        RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'        RsDev("branch_id").value = my_branch
'        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'        RsDev("DEV_ID_Line_No").value = lineno
'        RsDev("Value").value = val(Me.TxtPreVAT.Text)  '«·«Ì—«œ« 
'        RsDev("Credit_Or_Debit").value = 0
'        If SystemOptions.UserInterface = ArabicInterface Then
'            RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & " Õ”«» «·ﬁÌ„… «·„÷«›… œ›⁄«  „ﬁœ„…" & CHR(13) & TxtRemarks.Text
'        Else
'            RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & "Account VAT" & CHR(13) & TxtRemarks.Text
'        End If
'
'        RsDev("Notes_ID").value = val(note_id.Text)
'        RsDev("project_bill_no").value = val(txtid.Text)
'        If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
'         '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'         Else
'        '   RsDev("project_id").value = val(Me.DataCombo2.BoundText) 'xxxxxxxxxxxxxx
'         End If
'
'        RsDev("RecordDate").value = XPDtbTrans.value
'        RsDev("UserID").value = user_id
'        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'        RsDev("Posted").value = Posted
'        RsDev.update
'        lineno = lineno + 1
'''//////////////
'        RsDev.AddNew
'        If Option8.value = True Then
'            RsDev("Account_Code").value = Me.TxtAccountUnderImp.Text
'        Else
'            RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
'        End If
'        RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'        RsDev("branch_id").value = my_branch
'        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'        RsDev("DEV_ID_Line_No").value = lineno
'        RsDev("Value").value = val(Me.TxtPreVAT.Text)  '«·«Ì—«œ« 
'        RsDev("Credit_Or_Debit").value = 1
'        If SystemOptions.UserInterface = ArabicInterface Then
'            RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & " Õ”«» «·ﬁÌ„… «·„÷«›… œ›⁄«  „ﬁœ„…" & CHR(13) & TxtRemarks.Text
'        Else
'            RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & "Account VAT" & CHR(13) & TxtRemarks.Text
'        End If
'
'        RsDev("Notes_ID").value = val(note_id.Text)
'        RsDev("project_bill_no").value = val(txtid.Text)
'        If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
'         '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'         Else
'        '   RsDev("project_id").value = val(Me.DataCombo2.BoundText) 'xxxxxxxxxxxxxx
'         End If
' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'        RsDev("RecordDate").value = XPDtbTrans.value
'        RsDev("UserID").value = user_id
'        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'        RsDev("Posted").value = Posted
'        RsDev.update
'        lineno = lineno + 1
'
'    End If
'End If
'''/////////
'Else
''
'        'If SystemOptions.SubContactorHave3Account = True Then
'                Dim Discount1 As Double
'                Dim Discount2 As Double
'                Dim netvalue As Double
'                Dim TotalValue As Double
'                Dim AdvancedAccount As String
'                Dim GuranteeAccount As String
'                Dim line_no As Integer
'                Dim des As String
'                            If cboDiscount1.ListIndex = 0 Then
'                                Discount1 = 0
'                            ElseIf cboDiscount1.ListIndex = 1 Then
'                                Discount1 = val(txtDiscount1) * val(Me.TxtNetValue.Text) / 100
'                            ElseIf cboDiscount1.ListIndex = 2 Then
'                                Discount1 = val(txtDiscount1)
'                            End If
'
'                            If cboDiscount2.ListIndex = 0 Then
'                                Discount2 = 0
'                            ElseIf cboDiscount2.ListIndex = 1 Then
'                                Discount2 = val(txtDiscount2) * val(TxtNetValue.Text) / 100
'                            ElseIf cboDiscount2.ListIndex = 2 Then
'                                Discount2 = val(txtDiscount2)
'                            End If
'               AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code2")
'               GuranteeAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code1")
'               accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code")
'
'               line_no = 1
'               If SystemOptions.AllowNoRoudProjectInvoices = True Then
'                Discount1 = Round(Discount1, val(cCompanyInfo.NoRoudProjectInvoices))
'                Discount2 = Round(Discount2, val(cCompanyInfo.NoRoudProjectInvoices))
'               netvalue = Round(val(TxtNetValue.Text) - Discount1 - Discount2, val(cCompanyInfo.NoRoudProjectInvoices))
'               TotalValue = Round(val(TxtNetValue), val(cCompanyInfo.NoRoudProjectInvoices))
'               Else
'               Discount1 = Round(Discount1, 2)
'                Discount2 = Round(Discount2, 2)
'               netvalue = Round(val(TxtNetValue.Text) - Discount1 - Discount2, 2)
'               TotalValue = Round(val(TxtNetValue), Decimal_Places)
'              End If
'               If Option8.value = True Then
'                              des = "„’—Ê›«  «·„‘«—Ì⁄ " & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & CHR(13) & TxtRemarks.Text
'                         Else
'                         des = "„” Œ·’ «·„‘«—Ì⁄ " & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & "Manual#" & txtManualNo & CHR(13) & TxtRemarks.Text
'                   End If
'           If TotalValue > 0 Then '
'
'
'
'            If Option8.value = True Then
'               If ModAccounts.AddNewDev(LngDevID, line_no, TxtAccountUnderImp.Text, TotalValue, 0, Msg & des & "  " & "    " & TXTprojectname.Text, val(note_id.Text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , DataCombo2.BoundText, , , , , , , val(Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
'                                        GoTo ErrTrap
'                                    End If
'
'                                    line_no = line_no + 1
'                   Else
'                    If ModAccounts.AddNewDev(LngDevID, line_no, expanses_account, TotalValue, 0, Msg & des & "  " & "    " & TXTprojectname.Text, val(note_id.Text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , DataCombo2.BoundText, , , , , , , val(Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
'                                        GoTo ErrTrap
'                                    End If
'
'                                    line_no = line_no + 1
'                   End If
'                 If val(TxtFATValue.Text) > 0 Then
'                  If ModAccounts.AddNewDev(LngDevID, line_no, Me.AccountVat.BoundText, val(TxtFATValue.Text), 0, Msg & "  " & "    " & TXTprojectname.Text & "    VAT  " & "   INV# " & TxtNoteSerial1.Text, val(note_id.Text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
'                                        GoTo ErrTrap
'                                    End If
'
'                                    line_no = line_no + 1
'                    End If
'              '  End If
'  '////////////////////////////////////////////////////
'        ' End If
'
'      If SystemOptions.UserInterface = ArabicInterface Then
'               des = "Œ’„ ÷„«‰ «⁄„«· " & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & " ··„” Œ·’ —ﬁ„ " & TxtNoteSerial1.Text & CHR(13) & TxtRemarks.Text
'       Else
'                      des = " Discount " & "   " & TxtRemarks & "Inv# " & TxtNoteSerial1 & " manual#" & txtManualNo & CHR(13) & TxtRemarks.Text
'
'       End If
'           If Discount1 > 0 Then '÷„«‰ «·«⁄„«·
'
'                If GuranteeAccount = "" Then
'                GuranteeAccount = accountdep
'                End If
'
'
'               If ModAccounts.AddNewDev(LngDevID, line_no, GuranteeAccount, Discount1, 1, Msg & des & "  " & "··„‘—Ê⁄   " & TXTprojectname.Text, val(note_id.Text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
'                                        GoTo ErrTrap
'                                    End If
'
'                                    line_no = line_no + 1
'
'
'         End If
'       If SystemOptions.UserInterface = ArabicInterface Then
'         des = "Œ’„ œ›⁄«  „ﬁœ„…   " & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & CHR(13) & TxtRemarks.Text
'        Else
'        des = "Advance Discount " & "   " & TxtRemarks & "InvE " & TxtNoteSerial1 & "Manuall#   " & txtManualNo & CHR(13) & TxtRemarks.Text
'        End If
'           If Discount2 > 0 Then '
'
'               If AdvancedAccount = "" Then
'                AdvancedAccount = accountdep
'                End If
'
'               If ModAccounts.AddNewDev(LngDevID, line_no, AdvancedAccount, Discount2, 1, Msg & des & "  " & "··„‘—Ê⁄   " & TXTprojectname.Text, val(note_id.Text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , DataCombo2.BoundText, , , , , , , val(Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
'                                        GoTo ErrTrap
'                                    End If
'
'                                    line_no = line_no + 1
'
'
'         End If
'
'        '«·œ›⁄«  «·„ﬁœ„…
''Account_Code_dynamic2
'LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'If val(Me.advancedPayment.Text) > 0 Then
'    RsDev.AddNew
'
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Account_Code").value = accountdep '    Õ”«» œ›⁄«  „ﬁœ„…
'    RsDev("Value").value = val(Me.advancedPayment.Text) + val(Me.TxtPreVAT.Text)
'    RsDev("Credit_Or_Debit").value = 0
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & "œ›⁄«  „ﬁœ„…" & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & "Manual" & txtManualNo & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'    RsDev.update
''ll:
'lineno = lineno + 1
'
'  RsDev.AddNew
'
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Account_Code").value = AdvancedAccount   '    Õ”«» œ›⁄«  „ﬁœ„…
'    RsDev("Value").value = val(Me.advancedPayment.Text)
'    RsDev("Credit_Or_Debit").value = 1
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & "œ›⁄«  „ﬁœ„…" & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & "  manual  " & txtManualNo & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
' '   RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'    RsDev.update
''ll:
'lineno = lineno + 1
'
'
'
'
'  RsDev.AddNew
'
'    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Account_Code").value = Me.AccountVat.BoundText   '    Õ”«» œ›⁄«  „ﬁœ„…
'    RsDev("Value").value = val(Me.TxtPreVAT.Text)
'    RsDev("Credit_Or_Debit").value = 1
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & TXTprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & "œ›⁄«  „ﬁœ„…" & CHR(13) & TxtRemarks.Text
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & TXTprojectname.Text & " Manual" & txtManualNo & CHR(13) & TxtRemarks.Text
'    End If
'
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
' '   RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev("Posted").value = Posted
'    RsDev.update
'
'End If
'
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'         des = " «⁄„«·" & " ··„” Œ·’ —ﬁ„ " & TxtNoteSerial1.Text & CHR(13) & TxtRemarks.Text
'         Else
'         des = " Works" & "  Inv#" & TxtNoteSerial1.Text & CHR(13) & TxtRemarks.Text
'         End If
'
'           If netvalue > 0 Then '
'
'                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'                 If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, netvalue + val(TxtFATValue.Text), 1, Msg & des & "  " & "    " & TXTprojectname.Text, val(note_id.Text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
'                                        GoTo ErrTrap
'                                    End If
'
'                                    line_no = line_no + 1
'
'
'         End If
'
'
'
'End If
'End If
'
    Dim Rs3 As New ADODB.Recordset
 '   Rs3.Open "SubcontractorContract2", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
 
 Cn.Execute "Delete SubcontractorContract2  where bill_id =  " & val(Me.txtid.Text)
               StrSQL = "SELECT     * from dbo.SubcontractorContract2 Where (1 = -1)"
   Rs3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
          
    Dim i As Integer
Dim LineDiscountPercent As Double
Dim LineDiscount As Double
Dim linenetaftermainDiscount As Double
Dim linenetaftermainDiscountBeforevat As Double
Dim LineVat As Double
Dim linenetaftermainDiscountWithvat As Double
Dim OLDTotalwithVat  As Double
Dim CurrenttotalWithvat  As Double
Dim Totalwitvat  As Double
Dim oldPerforValue  As Double
Dim totalPerforValue  As Double





    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            '        Dim IntDEV_Type As Integer
            '        Dim SngDEV_Value As Single
            If .TextMatrix(i, .ColIndex("item")) <> "" Then

                Rs3.AddNew
                Rs3("bill_id").value = Me.txtid.Text
                Rs3("ExPercen").value = IIf(.TextMatrix(i, .ColIndex("ExPercen")) = "", Null, val(.TextMatrix(i, .ColIndex("ExPercen"))))
                
                Rs3("FullCode").value = Trim(.TextMatrix(i, .ColIndex("FullCode")))
                Rs3("project_id").value = IIf(.TextMatrix(i, .ColIndex("project_id")) = "", Null, val(.TextMatrix(i, .ColIndex("project_id"))))
                Rs3("projectName").value = .TextMatrix(i, .ColIndex("projectName"))
                
                
                Rs3("item").value = IIf(.TextMatrix(i, .ColIndex("item")) = "", Null, .TextMatrix(i, .ColIndex("item")))
                Rs3("item_id").value = IIf(.TextMatrix(i, .ColIndex("item_id")) = "", Null, .TextMatrix(i, .ColIndex("item_id")))
                Rs3("cost").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("cost"))), 0, .TextMatrix(i, .ColIndex("cost")))
                Rs3("exe").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("exe"))), 0, .TextMatrix(i, .ColIndex("exe")))
                
                Rs3("qtySubContractor").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("qtySubContractor"))), 0, .TextMatrix(i, .ColIndex("qtySubContractor")))
                Rs3("costSubContractor").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("costSubContractor"))), 0, .TextMatrix(i, .ColIndex("costSubContractor")))
                
                Rs3("OLDTotalwithVat").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("OLDTotalwithVat"))), 0, .TextMatrix(i, .ColIndex("OLDTotalwithVat")))
                Rs3("CurrenttotalWithvat").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("CurrenttotalWithvat"))), 0, .TextMatrix(i, .ColIndex("CurrenttotalWithvat")))
                Rs3("Totalwitvat").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Totalwitvat"))), 0, .TextMatrix(i, .ColIndex("Totalwitvat")))
                    
                                              
                Rs3("oldPerforValue").value = val(lbl(9).Caption)
                Rs3("totalPerforValue").value = val(lbl(11).Caption)
                 
                
                Rs3("percentage").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("percentage"))), 0, .TextMatrix(i, .ColIndex("percentage")))
                Rs3("exedate").value = IIf(Not IsDate(.TextMatrix(i, .ColIndex("exedate"))), Date, .TextMatrix(i, .ColIndex("exedate")))
                ''////29 07 2015
                Rs3("PrMainDesID").value = IIf(.TextMatrix(i, .ColIndex("PrMainDesID")) = "", Null, val(.TextMatrix(i, .ColIndex("PrMainDesID"))))
                Rs3("qty").value = IIf(.TextMatrix(i, .ColIndex("qty")) = "", Null, val(.TextMatrix(i, .ColIndex("qty"))))
                Rs3("total").value = IIf(.TextMatrix(i, .ColIndex("total")) = "", Null, val(.TextMatrix(i, .ColIndex("total"))))
                Rs3("discount").value = IIf(.TextMatrix(i, .ColIndex("discount")) = "", Null, val(.TextMatrix(i, .ColIndex("discount"))))
                Rs3("net").value = IIf(.TextMatrix(i, .ColIndex("net")) = "", Null, val(.TextMatrix(i, .ColIndex("net"))))
                Rs3("quntExc").value = IIf(.TextMatrix(i, .ColIndex("quntExc")) = "", Null, val(.TextMatrix(i, .ColIndex("quntExc"))))
                Rs3("totEx").value = IIf(.TextMatrix(i, .ColIndex("totEx")) = "", Null, val(.TextMatrix(i, .ColIndex("totEx"))))
                Rs3("Period").value = IIf(.TextMatrix(i, .ColIndex("Period")) = "", Null, val(.TextMatrix(i, .ColIndex("Period"))))
                
                
                LineDiscountPercent = val(.TextMatrix(i, .ColIndex("NetExe"))) / Results.Text
                
                 LineDiscount = (val(txtDiscountG.Text)) * LineDiscountPercent
         
                 PerforVLineDiscount = val(TxtPerforValue.Text) * LineDiscountPercent
                 
               
                 linenetaftermainDiscount = val(.TextMatrix(i, .ColIndex("NetExe"))) - LineDiscount
                 
                
                
                
                
                
                Rs3("LineDiscountPercent").value = LineDiscountPercent
                  Rs3("LineDiscount").value = LineDiscount
                   Rs3("PerforVLineDiscount").value = PerforVLineDiscount
                  linenetaftermainDiscountBeforevat = val(.TextMatrix(i, .ColIndex("NetExe"))) - LineDiscount
                  
                 Rs3("linenetaftermainDiscountBeforevat").value = linenetaftermainDiscountBeforevat
                 
                LineVat = Rs3("linenetaftermainDiscountBeforevat").value * val(TxtFATYou.Text) / 100
                
        '         Rs3("LineVat").value = LineVat
                 linenetaftermainDiscountWithvat = linenetaftermainDiscount + LineVat
                 
                 
                 
                  Rs3("linenetaftermainDiscountWithvat").value = linenetaftermainDiscountWithvat
             '     Rs3("linenetaftermainDiscountWithvat").value = linenetaftermainDiscountWithvat - PerforVLineDiscount
                  Rs3("LineFinal").value = linenetaftermainDiscountWithvat - PerforVLineDiscount
                  
              'newwwwwwwwwwwwwwwwwwwwwwww
              Rs3("QtyApprov").value = IIf(.TextMatrix(i, .ColIndex("QtyApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("QtyApprov"))))
              Rs3("QtyApprov").value = IIf(.TextMatrix(i, .ColIndex("QtyApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("QtyApprov"))))
              
              Rs3("QtyApprov").value = IIf(.TextMatrix(i, .ColIndex("QtyApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("QtyApprov"))))
              Rs3("TotalApprov").value = IIf(.TextMatrix(i, .ColIndex("TotalApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("TotalApprov"))))
              Rs3("PriceApprov").value = IIf(.TextMatrix(i, .ColIndex("PriceApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("PriceApprov"))))
              Rs3("DiscApprov").value = IIf(.TextMatrix(i, .ColIndex("DiscApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("DiscApprov"))))
              Rs3("NetApprov").value = IIf(.TextMatrix(i, .ColIndex("NetApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("NetApprov"))))
              '''//////////////
                Rs3("discountEXE").value = IIf(.TextMatrix(i, .ColIndex("discountEXE")) = "", Null, val(.TextMatrix(i, .ColIndex("discountEXE"))))
                
                Rs3("NetExe").value = IIf(.TextMatrix(i, .ColIndex("NetExe")) = "", Null, val(.TextMatrix(i, .ColIndex("NetExe"))))
                
                Rs3("Percentage1").value = IIf(.TextMatrix(i, .ColIndex("Percentage1")) = "", Null, val(.TextMatrix(i, .ColIndex("Percentage1"))))

                Rs3("Percentage1").value = IIf(.TextMatrix(i, .ColIndex("Percentage1")) = "", Null, val(.TextMatrix(i, .ColIndex("Percentage1"))))
                Rs3("Pre_Percent1").value = IIf(.TextMatrix(i, .ColIndex("Pre_Percent1")) = "", Null, val(.TextMatrix(i, .ColIndex("Pre_Percent1"))))
                Rs3("tot_percent1").value = IIf(.TextMatrix(i, .ColIndex("tot_percent1")) = "", Null, val(.TextMatrix(i, .ColIndex("tot_percent1"))))
                
               'newwwwwwwwwwwwwwwwwwwwwwwwwwwwww
     
                
                Rs3("oprid").value = IIf(.TextMatrix(i, .ColIndex("oprid")) = "", Null, val(.TextMatrix(i, .ColIndex("oprid"))))
                
               'Rs3("unit").value = IIf(.TextMatrix(i, .ColIndex("unit")) = "", Null, .TextMatrix(i, .ColIndex("unit")))
               Rs3("item_unit").value = IIf(.TextMatrix(i, .ColIndex("unit")) = "", Null, .TextMatrix(i, .ColIndex("unit")))
               Rs3("Unit_id").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Unit_id"))), 0, val(.TextMatrix(i, .ColIndex("Unit_id"))))
               Rs3("Quantity").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Quantity"))), 0, .TextMatrix(i, .ColIndex("Quantity")))
               Rs3("Price").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Price"))), 0, .TextMatrix(i, .ColIndex("Price")))
               Rs3("Pre_Quantity").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Pre_Quantity"))), 0, .TextMatrix(i, .ColIndex("Pre_Quantity")))
               Rs3("Pre_Value").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Pre_Value"))), 0, .TextMatrix(i, .ColIndex("Pre_Value")))
               Rs3("Pre_Percent").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Pre_Percent"))), 0, .TextMatrix(i, .ColIndex("Pre_Percent")))
               Rs3("Curr_Quantity").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Curr_Quantity"))), 0, .TextMatrix(i, .ColIndex("Curr_Quantity")))
                 Rs3("Curr_value").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Curr_value"))), 0, .TextMatrix(i, .ColIndex("Curr_value")))
                 Rs3("curr_Percent").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("curr_Percent"))), 0, .TextMatrix(i, .ColIndex("curr_Percent")))
                 Rs3("tot_quantity").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("tot_quantity"))), 0, .TextMatrix(i, .ColIndex("tot_quantity")))
                Rs3("tot_value").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("tot_value"))), 0, .TextMatrix(i, .ColIndex("tot_value")))
                 Rs3("tot_percent").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("tot_percent"))), 0, .TextMatrix(i, .ColIndex("tot_percent")))
                 
                 
                 
                Rs3.update
            End If

        Next i

    End With
updateNotesValueAndNobytext val(note_id.Text)
saveBillBuy
    TxtModFlg.Text = "R"
fillapprovData
    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Saved", vbInformation
    Else
        MsgBox " „ Õ›Ÿ «·»Ì«‰« ", vbInformation
  
    End If
    Retrive
    Exit Function
ErrTrap:
    
    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "error During Saving", vbInformation
    Else
        MsgBox "ÕœÀ Œÿ√ „« «À‰«¡ Õ›Ÿ «·»Ì«‰«  ", vbInformation
  
    End If
End Function

Private Sub CmdRemove_Click()
    Dim x As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If x = vbNo Then Exit Sub
    Dim sql As String
    
    If Fg_Journal.Rows > 1 Then
        If Fg_Journal.Rows = 2 Then
            Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Fg_Journal.Rows > 1 Then
                If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                    Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                End If
            End If
        End If
    End If
            
    ReLineGrid

End Sub

Private Sub Command1_Click(Index As Integer)
    'On Error Resume Next
    Select Case Index
Case 12
txtid.Text = ""
    TxtModFlg.Text = "N"

            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            Command1(1).Enabled = True
            XPDtbTrans_Change
     ALLButton1_Click
     
            ClculteVAT
            
ReLineGrid
ReLineGrid

        Case 0
 
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            Accredit.Caption = ""


            Me.DCboUserName.BoundText = user_id
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 1
            Fg_Journal.Enabled = True
            Command1(1).Enabled = True
            XPDtbTrans.value = DateValue(Now)
            Results.Text = 0
            XPDtbTrans.value = Date
            TxtNoteSerial.Text = ""
            Me.dcBranch.BoundText = Current_branch
            cboDiscount1.ListIndex = 0
            cboDiscount1.ListIndex = 0
            cboDiscount2.ListIndex = 0
billto.ListIndex = 0

        Case 1
ClculteVAT
                         If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
  '«· «ﬂœ „‰ «·Õ„Ì« 
  If val(txtDiscount.Text) > 0 Then
          Account_Code_dynamic1 = get_account_code_branch(103, my_branch)
        
        If Account_Code_dynamic1 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        
            'create_accounts = False
            Exit Sub
        Else

            If Account_Code_dynamic1 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» „’—Ê›«   «·Õ”„Ì«  ·„” Œ·’« /›Ê« Ì— «·„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        
             '   create_accounts = False
                Exit Sub
            End If
        End If
        
End If
    
    If val(TxtPerforValue.Text) > 0 Then
          
        
        If AcountGood = "" Then
            MsgBox "·„ Ì „ «‰‘«¡ Õ”«» Õ”‰ «·«œ«¡ ·Â–« «·„‘—Ê⁄", vbCritical
        
            'create_accounts = False
            Exit Sub

        End If
        
End If
    
    
   '«· «ﬂœ „‰ «·œ›⁄«  «·„ﬁœ„…
  If val(advancedPayment.Text) > 0 Then
  
  
  
        If SystemOptions.CustomerhavethreeAccounts = True Then
             Account_Code_dynamic2 = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.TXTEnd_user_id.Text), "Account_code2")
         
                     If Account_Code_dynamic2 = "" Then
                     
                            Account_Code_dynamic2 = get_account_code_branch(104, my_branch)
                    
                            If Account_Code_dynamic2 = "NO branch" Then
                                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                            
                                'create_accounts = False
                                Exit Sub
                            Else
                    
                                        If Account_Code_dynamic2 = "NO account" Then
                                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»   „œ›Ê⁄«  „ﬁœ„… ··⁄„·«¡", vbCritical
                                    
                                         '   create_accounts = False
                                            Exit Sub
                                        End If
                            End If
                    
                
                
                     End If
         Else
         
         
         Account_Code_dynamic2 = get_account_code_branch(104, my_branch)
        
        If Account_Code_dynamic2 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        
            'create_accounts = False
            Exit Sub
        Else

            If Account_Code_dynamic2 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „œ›Ê⁄«  „ﬁœ„…  „œ›Ê⁄«  „ﬁœ„… ··⁄„·«¡ ", vbCritical
        
             '   create_accounts = False
                Exit Sub
            End If
        End If
        
        
         End If
  

        
        
        
End If



    
If val(total.Text) <= 0 Then MsgBox "Õœœ  ﬂ·›… «·„„‰›– «Ê·«", vbCritical: Exit Sub
If Not IsNumeric(dcBranch.BoundText) Then MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical: Exit Sub
my_branch = val(dcBranch.BoundText)
    
            If TxtNoteSerial.Text = "" Then
                If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                    If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        '       TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
                    End If
                End If
            End If
                
            If SystemOptions.UserInterface = EnglishInterface Then
                If billto.ListIndex = -1 Then MsgBox "Specify Bill TO", vbCritical: Exit Sub
                'If DcAccount1.text = "" And billto.ListIndex = 1 Then MsgBox "this project have no subcontractor", vbCritical: Exit Sub

            Else

                If billto.ListIndex = -1 Then MsgBox "Õœœ «·„” Œ·’  «·Ï «Ê·«", vbCritical: Exit Sub
               ' If DcAccount1.text = "" And billto.ListIndex = 1 Then MsgBox "·«Ì„ﬂ‰ Õ›Ÿ «·„” Œ·’ ·«‰ﬂ «Œ —  „ﬁ«Ê· »«ÿ‰ Ê«·„‘—Ê⁄ ·Ì” ·Â „ﬁ«Ê· »«ÿ‰", vbCritical: Exit Sub
            End If
       If val(TxtBillNo.Text) > 0 Then
       
             If val(DcbPeriodType.ListIndex) = -1 Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                      MsgBox "Ì—ÃÏ «œŒ«· ‰Ê⁄ «·› —… »Ì‰ «·›Ê« Ì—"
                  Else
                      MsgBox "Please Enter Type Period"
                  End If
                  DcbPeriodType.SetFocus
                  Exit Sub
            
            End If
       End If
       
'GET_PROJECT_DATA
         Dim TxtNoteSerial1str As String
my_branch = val(Me.dcBranch.BoundText)
    If TxtNoteSerial1.Text = "" Then
     TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 65, 65, , , , , val(billto.ListIndex))
    
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…  Õ—ﬂ…  ÃœÌœ…  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  «·Õ—ﬂ… ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                  ' TxtNoteSerial1.text = TxtNoteSerial1str
                        '             txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, DCPreFix.text)
                    End If
                End If
    End If
Dim AccountVATDept As String
Dim str As String
str = "30/05/2017"
If AccountVat.BoundText = "" And True = True And CheckAnyVAT(XPDtbTrans.value) = True And StartDateProje.value > CDate(str) Then
MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
Exit Sub
End If
            SaveData

            ''Adodc1.Recordset.Fields!  project_no = DataCombo2.text
        Case 11

            On Error Resume Next

ShowAttachments txtid.Text, "24122020001"


Exit Sub

            If SystemOptions.UserInterface = EnglishInterface Then
                If txtid.Text = "" Then MsgBox "Select Bill firstly": Exit Sub

            Else

                If txtid.Text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „” Œ·’ «Ê·«": Exit Sub

            End If

            imaged.show

            If SystemOptions.UserInterface = EnglishInterface Then

                imaged.Label9.Caption = "Attachment For Project Bill "
                imaged.Caption = "Project  Bill Attachment  "
                imaged.Label6.Caption = "   Bill NO"
                Label5.Caption = "Documents"
                Label8.Caption = "Forms"

            Else

                imaged.Label9.Caption = "„—›ﬁ«    „” Œ·’ „‘—Ê⁄  —ﬁ„"
                imaged.Caption = "„—›ﬁ«  «·„” Œ·’     "
                imaged.Label6.Caption = "—ﬁ„ «·„” Œ·’   "

            End If

            imaged.SUBJECT_NO = txtid.Text
            imaged.txtopeation_type = "„—›ﬁ«  „” Œ·’"

            imaged.Adodc1.CommandType = adCmdText
            imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  „” Œ·’' and subject_no='" & txtid.Text & "'"
            imaged.Adodc1.Refresh

            If imaged.Adodc1.Recordset.RecordCount > 0 Then

                imaged.DBPix201.Visible = True
            Else
                imaged.DBPix201.Visible = False
            End If

        Case 3
        Frame15.Enabled = True
        Frame15.Visible = False
                     If ScreenAproved(val(txtid.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
       
       
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            Dim Msg As String
            Dim StrSQL As String
 
            Dim RsTemp As New ADODB.Recordset
'            StrSQL = "select * From ProjectBillBuy where Bill_id=" & val(txtid.Text)
'            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'            If Not (RsTemp.EOF Or RsTemp.BOF) Then
'                Msg = "·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  Â–« «·›« Ê—… " & CHR(13)
'                Msg = Msg + "·«‰Â«  „ ⁄·ÌÂ« ⁄„·Ì«  ”œ«œ"
'                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                Exit Sub
'            End If
          
            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id
            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            Command1(1).Enabled = True

        Case 4

        Case 5

        Case 6
            Undo

        Case 9
                     If ScreenAproved(val(txtid.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–› .Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
       
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 7 'ÿ»«⁄Â «·›« Ê—…

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            print_report val(DataCombo2.BoundText)

        Case 8

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.Text, , 200
            
    Case 10
    
          '  projectsbill_Search.show
  Case 15
  If Me.TxtModFlg.Text = "N" Then
  StartDate.value = XPDtbTrans.value
  End If
  Frame7.Visible = True
  Case 16
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            print_report2 val(DataCombo2.BoundText)
    End Select

End Sub
'
'Sub SaveBillMonthly()
'Dim RsDevsub As ADODB.Recordset
'Dim StrSQL As String
'Dim i As Integer
'StrSQL = "Delete from SubcontractorContract_Month where Bill_ID =" & val(Me.txtid.Text) & ""
'Cn.Execute StrSQL
'   Set RsDevsub = New ADODB.Recordset
'    StrSQL = "SELECT  *  from SubcontractorContract_Month Where (1 = -1)"
'    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'For i = 1 To val(Me.TxtBillNo.Text)
'If i = 1 Then
'DateTemp.value = StartDate.value
'Else
'If val(DcbPeriodType.ListIndex) = 0 Then
'DateTemp.value = DateAdd("D", val(Me.TxtPeriod.Text), DateTemp.value)
'ElseIf val(DcbPeriodType.ListIndex) = 1 Then
'DateTemp.value = DateAdd("M", val(Me.TxtPeriod.Text), DateTemp.value)
'ElseIf val(DcbPeriodType.ListIndex) = 2 Then
'DateTemp.value = DateAdd("YYYY", val(Me.TxtPeriod.Text), DateTemp.value)
'End If
'End If
'RsDevsub.AddNew
'RsDevsub("Bill_ID").value = val(txtid.Text)
'RsDevsub("RecordDate").value = DateTemp.value
'RsDevsub.update
'Next i
'End Sub
Function print_report(Optional NoteSerial As Integer)
    
     On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
   'new
    
    MySQL = "SELECT         SubcontractorContract.PerforValue,   dbo.SubcontractorContract.id, dbo.SubcontractorContract.bill_date, dbo.SubcontractorContract.ManualNO, dbo.SubcontractorContract.duedate1, dbo.SubcontractorContract.discount, dbo.SubcontractorContract.dueDate, dbo.SubcontractorContract.NoteSerial, dbo.SubcontractorContract.total, "
MySQL = MySQL & "                           dbo.SubcontractorContract.Remarks, dbo.SubcontractorContract.Results, dbo.SubcontractorContract.advancedPayment, dbo.SubcontractorContract.discount2value, dbo.SubcontractorContract.discount1value, dbo.SubcontractorContract.bill_type, dbo.SubcontractorContract.project_no,"
MySQL = MySQL & "                                 dbo.projects.Fullcode, dbo.SubcontractorContract.project_name, dbo.SubcontractorContract.End_user_name, dbo.SubcontractorContract.Sub_user_name, dbo.SubcontractorContract.End_user_account, dbo.SubcontractorContract.bill_to, dbo.SubcontractorContract.Sub_user_account,"
MySQL = MySQL & "                                 dbo.SubcontractorContract.revenue_account, dbo.SubcontractorContract.subContractorId, dbo.TblCustemers.Address, dbo.TblCustemers.VATNO, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.SubcontractorContract.Branch_NO,"
MySQL = MySQL & "                                 dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.SubcontractorContract.discount1ID, dbo.SubcontractorContract.discount2ID, dbo.SubcontractorContract.note_id, dbo.SubcontractorContract2.project_no AS project_noDet,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.item, dbo.SubcontractorContract2.cost, dbo.SubcontractorContract2.exe, dbo.SubcontractorContract2.percentage * 100 AS percentage, dbo.SubcontractorContract2.exedate, dbo.SubcontractorContract2.bill_id,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.line_no, dbo.SubcontractorContract2.item_id, dbo.SubcontractorContract2.Quantity, dbo.SubcontractorContract2.Price, dbo.SubcontractorContract2.Pre_Quantity, dbo.SubcontractorContract2.Pre_Value,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.Pre_Percent * 100 AS Pre_Percent, dbo.SubcontractorContract2.Curr_Quantity, dbo.SubcontractorContract2.Curr_value, dbo.SubcontractorContract2.curr_Percent * 100 AS curr_Percent,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.tot_quantity, dbo.SubcontractorContract2.tot_value, dbo.SubcontractorContract2.tot_percent * 100 AS tot_percent, dbo.SubcontractorContract2.Unit_id, dbo.TblProcessUnites.UnitName,"
MySQL = MySQL & "                                 dbo.TblProcessUnites.UnitNamee, dbo.SubcontractorContract2.oprid, dbo.SubcontractorContract2.totEx, dbo.SubcontractorContract2.quntExc, dbo.SubcontractorContract2.net, dbo.SubcontractorContract2.discount AS discountDet,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.total AS totalDet, dbo.SubcontractorContract2.qty, dbo.SubcontractorContract2.item_unit, dbo.SubcontractorContract2.discountEXE, dbo.SubcontractorContract2.NetExe,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.percentage1 * 100 AS percentage1, dbo.SubcontractorContract2.Pre_Percent1 * 100 AS Pre_Percent1, dbo.SubcontractorContract2.tot_percent1, dbo.SubcontractorContract2.percentage1 AS Expr1,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.Pre_Percent1 AS Expr2, dbo.SubcontractorContract2.QtyApprov, dbo.SubcontractorContract2.PriceApprov, dbo.SubcontractorContract2.TotalApprov, dbo.SubcontractorContract2.DiscApprov,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.NetApprov, dbo.SubcontractorContract.NoteSerial1, dbo.SubcontractorContract.FATYou, dbo.SubcontractorContract.FATValue, dbo.SubcontractorContract.TotalValue, dbo.SubcontractorContract.ExPercen, dbo.SubcontractorContract.ExPercenID,"
MySQL = MySQL & "                                 dbo.SubcontractorContract2.ExPercen AS ExPercenDet, dbo.SubcontractorContract.PreVAT, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.projects.REVENUE_account_balance, dbo.projects.Project_nameE"
MySQL = MySQL & "        FROM            dbo.TblCustemers INNER JOIN"
MySQL = MySQL & "                                 dbo.projects ON dbo.TblCustemers.CusID = dbo.projects.End_user_id RIGHT OUTER JOIN"
 MySQL = MySQL & "                                dbo.TblProcessUnites RIGHT OUTER JOIN"
MySQL = MySQL & "                                 dbo.SubcontractorContract2 ON dbo.TblProcessUnites.UnitID = dbo.SubcontractorContract2.Unit_id RIGHT OUTER JOIN"
MySQL = MySQL & "                                 dbo.SubcontractorContract ON dbo.SubcontractorContract2.bill_id = dbo.SubcontractorContract.id LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblBranchesData ON dbo.SubcontractorContract.Branch_NO = dbo.TblBranchesData.branch_id ON dbo.projects.id = dbo.SubcontractorContract.project_no"
MySQL = MySQL & "  Where dbo.SubcontractorContract.id  = " & val(txtid.Text)
    MySQL = MySQL + " order by SubcontractorContract2.id"

    If val(Combo2.ListIndex) = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects.rpt"
    Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects_E.rpt"
    End If
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects" & val(Combo2.ListIndex) & ".rpt"
    Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects_E" & val(Combo2.ListIndex) & ".rpt"
    End If
    End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng

    End If
     If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(3).AddCurrentValue GetRegVATNo(val(dcBranch.BoundText))
    End If

'xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.VATRegNo
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_report2(Optional NoteSerial As Integer)
    
     On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
Dim currentProjectid As Integer
Dim billlID As Integer
 billlID = val(txtid.Text)
currentProjectid = val(DataCombo2.BoundText)
subContractorId = val(DcbosubContractor.BoundText)

 'newwwwwwwwwwwwwwwww


MySQL = " SELECT OLDTotalwithVat, CurrenttotalWithvat, Totalwitvat, oldPerforValue, totalPerforValue ,        ROUND(dbo.SubcontractorContract2.totEx * (1 + dbo.SubcontractorContract.FATYou / 100), 3) AS CurrenttotalWithVatSalim, dbo.SubcontractorContract.project_no, dbo.projects.Fullcode, dbo.SubcontractorContract.project_name, dbo.TblProcessUnites.UnitName,"
MySQL = MySQL & "                           dbo.TblProcessUnites.UnitNamee, dbo.projects.id AS ProjectID, dbo.GetPaymentValue(dbo.projects.id, dbo.SubcontractorContract.bill_date) AS TotalPayment, dbo.SubcontractorContract2.item, dbo.SubcontractorContract2.exe,"
MySQL = MySQL & "                           dbo.SubcontractorContract.discount AS Sumdiscount, dbo.SubcontractorContract2.discountEXE AS SumdiscountEXE, dbo.SubcontractorContract.advancedPayment AS SumadvancedPayment, dbo.SubcontractorContract2.oprid, dbo.SubcontractorContract.FATYou,"
MySQL = MySQL & "                           dbo.SubcontractorContract.FATValue, dbo.SubcontractorContract.NoteSerial1, dbo.SubcontractorContract2.cost, dbo.SubcontractorContract.PreVAT, dbo.SubcontractorContract.PreBalaValue, dbo.SubcontractorContract.PreBalaVAT, dbo.SubcontractorContract.PreBalaTotal,"
MySQL = MySQL & "                           dbo.SubcontractorContract.PreBalaPayed, dbo.SubcontractorContract.PreBalaRemain, dbo.SubcontractorContract.PreBalaTransPyed, dbo.SubcontractorContract.PreBalaNet, dbo.SubcontractorContract.SumVATLine, dbo.SubcontractorContract.PreBalaVATYu, dbo.SubcontractorContract.SumValueLine,"
MySQL = MySQL & "                           dbo.SubcontractorContract.StartDateProje, dbo.SubcontractorContract.NetValue, dbo.SubcontractorContract.PerforValue, dbo.SubcontractorContract.PostedDate, dbo.SubcontractorContract.Posted, dbo.SubcontractorContract2.Curr_Quantity, dbo.SubcontractorContract2.Curr_value,"
MySQL = MySQL & "                           dbo.SubcontractorContract2.tot_quantity, dbo.projects.End_user_name, dbo.projects.sub_contractor_Account, dbo.SubcontractorContract.bill_date, dbo.SubcontractorContract.NoteSerial, dbo.SubcontractorContract.Branch_NO, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                           dbo.TblBranchesData.branch_namee, dbo.SubcontractorContract.Remarks, dbo.SubcontractorContract.ManualNO, dbo.SubcontractorContract.UserID, dbo.SubcontractorContract.BillNo, dbo.TblUsers.UserName, dbo.SubcontractorContract.StartDate, dbo.SubcontractorContract.Period,"
MySQL = MySQL & "                           dbo.SubcontractorContract.Remarks2, dbo.SubcontractorContract.subContractorId, dbo.TblCustemers.CusName AS Subcontractorname, dbo.TblCustemers.CusNamee AS Subcontractornamee, dbo.TblCustemers.CusID,"
MySQL = MySQL & "                           ROUND(dbo.GetTotEx(dbo.SubcontractorContract2.oprid, 9), 2) AS OldtotlEXEValueSalim, ROUND(dbo.GetQuntExc(dbo.SubcontractorContract2.oprid, 9), 2) AS oldtotQtysalim, ROUND(dbo.GetOLDPerforValue(18, 9), 2)"
MySQL = MySQL & "                           AS tOLDPerforValue, dbo.SubcontractorContract2.qty, ROUND(dbo.GetOLDPerforValuebysubContractorId(18, 9, 328), 2) AS OLDPerforValuebyCONTRACTORiD, dbo.SubcontractorContract2.quntExc, dbo.SubcontractorContract2.totEx,"
MySQL = MySQL & "                           dbo.SubcontractorContract2.LineDiscountPercent, dbo.SubcontractorContract2.LineDiscount, dbo.SubcontractorContract2.linenetaftermainDiscount, dbo.SubcontractorContract2.linenetaftermainDiscountBeforevat, dbo.SubcontractorContract2.LineVat,"
MySQL = MySQL & "                           dbo.SubcontractorContract2.linenetaftermainDiscountWithvat, dbo.SubcontractorContract2.PerforVLineDiscount, dbo.SubcontractorContract2.LineFinal, dbo.SubcontractorContract2.qtySubContractor, dbo.SubcontractorContract2.costSubContractor,"
MySQL = MySQL & "                           dbo.SubcontractorContract2.percentage1, dbo.SubcontractorContract2.Pre_Percent1, dbo.SubcontractorContract2.tot_percent1, dbo.SubcontractorContract2.Pre_Quantity, dbo.SubcontractorContract2.Pre_Value, dbo.SubcontractorContract2.Pre_Percent,"
MySQL = MySQL & "                           dbo.SubcontractorContract2.curr_Percent, dbo.SubcontractorContract2.percentage, dbo.SubcontractorContract2.exedate, dbo.SubcontractorContract2.Quantity, dbo.SubcontractorContract2.Price, dbo.SubcontractorContract2.tot_value,"
MySQL = MySQL & "                           dbo.SubcontractorContract2.tot_percent, dbo.SubcontractorContract2.total, dbo.SubcontractorContract2.discount, dbo.SubcontractorContract2.net, dbo.SubcontractorContract2.QtyApprov, dbo.SubcontractorContract2.TotalApprov,"
MySQL = MySQL & "                           dbo.SubcontractorContract2.PriceApprov , dbo.SubcontractorContract2.DiscApprov, dbo.SubcontractorContract2.NetApprov, dbo.SubcontractorContract2.PrMainDesID"
MySQL = MySQL & "   FROM            dbo.TblBranchesData RIGHT OUTER JOIN"
MySQL = MySQL & "                           dbo.SubcontractorContract ON dbo.TblBranchesData.branch_id = dbo.SubcontractorContract.Branch_NO LEFT OUTER JOIN"
MySQL = MySQL & "                           dbo.TblUsers ON dbo.SubcontractorContract.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
MySQL = MySQL & "                           dbo.TblCustemers ON dbo.SubcontractorContract.subContractorId = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                           dbo.TblProcessUnites RIGHT OUTER JOIN"
MySQL = MySQL & "                           dbo.SubcontractorContract2 ON dbo.TblProcessUnites.UnitID = dbo.SubcontractorContract2.Unit_id ON dbo.SubcontractorContract.id = dbo.SubcontractorContract2.bill_id LEFT OUTER JOIN"
MySQL = MySQL & "                           dbo.projects ON dbo.SubcontractorContract.project_no = dbo.projects.id"
 MySQL = MySQL & "  Where (dbo.SubcontractorContract.ID = " & billlID & ") And (dbo.Projects.ID = " & currentProjectid & ")     ORDER BY SubcontractorContract2.ID"



If billto.ListIndex = 0 Then
   ' MySQL = MySQL + " order by SubcontractorContract2.id"
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects2.rpt"
    Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects2.rpt"
    End If
    
 Else
 
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects3.rpt"
    Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects3.rpt"
    End If
 
 End If
 

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng

    End If
     If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(3).AddCurrentValue GetRegVATNo(val(dcBranch.BoundText))
    End If

'xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.VATRegNo
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function




Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    'On Error GoTo ErrTrap
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If Me.txtid.Text <> "" Then
'        StrSQL = "select * From SubcontractorContract where Bill_id=" & val(txtid.Text)
'        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'        If Not (RsTemp.EOF Or RsTemp.BOF) Then
'            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·›« Ê—… " & CHR(13)
'            Msg = Msg + "·«‰Â«  „ ⁄·ÌÂ« ⁄„·Ì«  ”œ«œ"
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            Exit Sub
'        End If
    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtid.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
'            StrSQL = "Delete  Notes  where NoteSerial ='" & TxtNoteSerial & "'"
'            Cn.Execute StrSQL, , adExecuteNoRecords
'            StrSQL = "Delete from SubcontractorContract_Month where Bill_ID =" & val(Me.txtid.Text) & ""
'            Cn.Execute StrSQL
'            StrSQL = "Delete From TblPayPrePayed Where NoteID1=" & val(Me.txtid.Text)
'           Cn.Execute StrSQL, , adExecuteNoRecords
'           StrSQL = "Delete From TblProjePayPrePayed Where   NoteID=" & val(Me.txtid.Text)
'          Cn.Execute StrSQL, , adExecuteNoRecords
'          DeleteBillBuy
            StrSQL = "Delete from SubcontractorContract2 where bill_id =" & val(Me.txtid.Text) & ""
            Cn.Execute StrSQL
          VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
          VSFlexGrid4.Rows = 1
            If Not rs.RecordCount < 1 Then
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.Rows = 3
                    Fg_Journal.Enabled = False
           
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     
        Exit Sub
    End If
 
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



Function GET_PROJECT_DATA(Optional IDx As Integer = 0)
    On Error Resume Next

    If DataCombo2.Text = "" Then Exit Function
    Dim My_SQL As String

    My_SQL = "select * from projects where id =" & DataCombo2.BoundText
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Rec.CursorLocation = adUseClient

    Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If SystemOptions.UserInterface = ArabicInterface Then
    txtprojectname.Text = Rec.Fields("Project_name").value
Else
txtprojectname.Text = Rec.Fields("Project_nameE").value
End If
    txtsubaccount.Text = IIf(IsNull(Rec.Fields("sub_contractor_Account").value), "", Rec.Fields("sub_contractor_Account").value)
    TxtAccountUnderImp.Text = IIf(IsNull(Rec.Fields("AccountUnderImp").value), "", Rec.Fields("AccountUnderImp").value)

    DcAccount1.Text = IIf(IsNull(Rec.Fields("sub_contractor_name").value), "", Rec.Fields("sub_contractor_name").value)
    txtendaccount.Text = IIf(IsNull(Rec.Fields("End_user_Account").value), "", Rec.Fields("End_user_Account").value)
    DcAccount2.Text = IIf(IsNull(Rec.Fields("End_user_name").value), "", Rec.Fields("End_user_name").value)
 Dim End_user_id As Double
 Dim sub_contractor_id As Double
 If Me.TxtModFlg = "N" Then
 StartDateProje.value = IIf(IsNull(Rec.Fields("StartDate").value), Date, Rec.Fields("StartDate").value)
 StartDateProje_Change
 End If
 End_user_id = IIf(IsNull(Rec.Fields("End_user_id").value), 0, Rec.Fields("End_user_id").value)
 sub_contractor_id = IIf(IsNull(Rec.Fields("sub_contractor_id").value), 0, Rec.Fields("sub_contractor_id").value)
 DcAccount2.Text = GET_ACCOUNT_name_by_Code(get_Customer_Account(End_user_id))
 DcAccount1.Text = GET_ACCOUNT_name_by_Code(get_Customer_Account(sub_contractor_id))
 If SystemOptions.Revenueowed = True Then
    txtrevenue_account.Text = IIf(IsNull(Rec.Fields("legal").value), "", Rec.Fields("legal").value) 'Õ”«» «·„” Œ·’« \
  Else
      txtrevenue_account.Text = IIf(IsNull(Rec.Fields("REVENUE_account").value), "", Rec.Fields("REVENUE_account").value) 'Õ”«» «·«Ì—«œ« \

  End If
  
TXTEnd_user_id.Text = IIf(IsNull(Rec.Fields("End_user_id").value), "", Rec.Fields("End_user_id").value) '—ﬁ„ «·⁄„Ì· «·‰Â«∆Ì
TXTsub_contractor_id.Text = IIf(IsNull(Rec.Fields("sub_contractor_id").value), "", Rec.Fields("sub_contractor_id").value) '—ﬁ„   „ﬁ«Ê· «·»«ÿ‰

 expanses_account = IIf(IsNull(Rec.Fields("expanses_account").value), "", Rec.Fields("expanses_account").value) 'Õ”«»  «·„’—Ê›« \
 AcountGood = IIf(IsNull(Rec.Fields("AcountGood").value), "", Rec.Fields("AcountGood").value)
    If Not IsNull(Rec("UnderImp").value) Then
   If Rec("UnderImp").value = 0 Then
   Option7.value = True
   ElseIf Rec("UnderImp").value = 1 Then
   Option6.value = True
   ElseIf Rec("UnderImp").value = 2 Then
   Option8.value = True
   End If
   Else
   Option7.value = True
   End If
   If IDx = 1 Then
 ReloadContrac (val(DataCombo2.BoundText))
 End If
    'My_SQL = "  select net,des from projects_des  where project_id='" & DataCombo2.BoundText & "'"
    'fill_combo DataCombo5, My_SQL

End Function

Private Sub Command10_Click()
Dim i As Integer
Dim StrSQL As String
Exit Sub
If Me.TxtModFlg.Text = "E" Then
DeleteBillBuy
VSFlexGrid4.Enabled = True
        Check1.Enabled = True
      StrSQL = "Delete From TblPayPrePayed Where NoteID1=" & val(Me.txtid.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblProjePayPrePayed Where   NoteID=" & val(Me.txtid.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
VSFlexGrid4.Rows = 1

FlgBillBuy = True
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If
ALLButton1_Click
    With Me.VSFlexGrid4

            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i


        End With
End If
End Sub

Private Sub DataCombo2_Change()
    GET_PROJECT_DATA 1
   
End Sub

Private Sub DataCombo2_Click(Area As Integer)
'    GET_PROJECT_DATA 1
   ' ReloadContrac (val(DataCombo2.BoundText))
End Sub

Private Sub DataCombo5_Click(Area As Integer)

    If DataCombo5.BoundText <> "" Then
        txtcostSubContractor.Text = DataCombo5.BoundText
        Text9.Text = ""
    Else
        DataCombo5 = ""
    End If

End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = 46 Then
        If Adodc7.Recordset.RecordCount > 0 Then
            Adodc7.Recordset.delete
            DataGrid2.Refresh
            Command1_Click (1)
            total.Text = Round(gettotal(txtid.Text), Decimal_Places)

        End If

    End If

End Sub

Function gettotal(x As String) As Double
    Dim My_SQL As String

    My_SQL = "  select Sum(exe) as total  from SubcontractorContract2 where bill_id=" & x

    Set Rec = New ADODB.Recordset
    Rec.CursorLocation = adUseClient

    Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    gettotal = IIf(IsNull(Rec.Fields("total").value), 0, Rec.Fields("total").value)

End Function

Private Sub DataCombo2_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        My_SQL = "  select id,Fullcode from Projects"
        fill_combo DataCombo2, My_SQL
    End If


        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 8
             FrmProjectSearch.show vbModal
           
        End If
        
        
End Sub

Private Sub DcbosubContractor_KeyUp(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 10
           FrmCompanySearch.show vbModal
           
        End If
        
End Sub

Private Sub Dcbranch_Change()
    If ChekSanNumber(Current_branch, 65) = True Then
          TxtNoteSerial1.Text = ""
      End If
      TxtNoteSerial.Text = ""
End Sub

Private Sub employee_details_Click(Index As Integer)

    Select Case Index

        Case 0

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame14.Visible = True

                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))
                Retrive4 current_opr
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame14.Caption = "⁄„«· «·⁄„·Ì… —ﬁ„ :   " & "  " & current_opr
                Else
                    Frame10.Caption = "Labors For Process No:   " & "  " & current_opr
                End If

                XPTxtSum.Text = 0
            End If

        Case 1
            Frame14.Visible = False
            VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("total_salary1")) = val(txt_emp_salary)
            ReLineGrid

    End Select

End Sub
Sub RelineBu22()
    Dim IntCounter As Integer
    Dim Percetage As Double
    Dim PercetageAdv As Double
    Dim SumVATLine As Double
    Dim SumValueLine As Double
    Dim Sm As Double
    Sm = 0
    SumVATLine = 0
    SumValueLine = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid4
        For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
              If val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > val(.TextMatrix(i, .ColIndex("RemainingValue"))) And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) <> 0 Then
              If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ﬁÌ„… «·œ›⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
              Else
              MsgBox "Can Not PaymentValue Larger Than Total Value "
              End If
              .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
              Exit Sub
              End If
    If val(TxtFATYou.Text) <> 0 And val(.TextMatrix(i, .ColIndex("VAT"))) <> 0 Then
    
    
    PercentgValueAddedAccount_Transec .TextMatrix(i, .ColIndex("NoteDate")), 6, 1, "", PercetageAdv
  PercetageAdv = PercetageAdv / 100 + 1
  
   .TextMatrix(i, .ColIndex("ValueLine")) = val(.TextMatrix(i, .ColIndex("TransPayedValue"))) / PercetageAdv
   .TextMatrix(i, .ColIndex("VATLine")) = val(.TextMatrix(i, .ColIndex("TransPayedValue"))) - val(.TextMatrix(i, .ColIndex("ValueLine")))  '"  .TextMatrix(i, .ColIndex("ValueLine")) * PercetageAdv / 100
   Else
     .TextMatrix(i, .ColIndex("ValueLine")) = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
   .TextMatrix(i, .ColIndex("VATLine")) = 0
   End If
           Sm = Sm + val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           SumVATLine = SumVATLine + val(.TextMatrix(i, .ColIndex("VATLine")))
           SumValueLine = SumValueLine + val(.TextMatrix(i, .ColIndex("ValueLine")))
   
           End If
           Next i
  
    End With
    Label56.Caption = Round(SumValueLine, 2)
    Label57.Caption = Round(SumVATLine, 2)
   Label47.Caption = Sm
SumVAT
End Sub
Sub SumVAT()
Dim Percetage2 As Double
If Me.TxtModFlg.Text <> "R" Then
Dim val1 As Double
TxtPreVAT.Text = 0
advancedPayment.Text = 0
     advancedPayment.Text = val(Label56.Caption)
   TxtPreVAT.Text = val(Label57.Caption)
 If val(TxtPreBalaTransPyed.Text) > 0 Then
     If val(TxtPreBalaVATYu.Text) <> 0 Then
   Percetage2 = val(TxtPreBalaVATYu.Text) / 100 + 1
   val1 = Round(val(TxtPreBalaTransPyed.Text) / Percetage2, 4)
   advancedPayment.Text = val(advancedPayment.Text) + val1
   TxtPreVAT.Text = val(TxtPreVAT.Text) + Round(val1 * val(TxtPreBalaVATYu.Text) / 100, 4)
   Else
    advancedPayment.Text = val(advancedPayment.Text) + val(TxtPreBalaTransPyed.Text)
   TxtPreVAT.Text = 0
   End If
 End If
 End If
End Sub

Private Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim netexe As Double
    Dim QtyExe As Double
    With Fg_Journal
    If val(.TextMatrix(Row, .ColIndex("ExPercen"))) = 0 Then
    If val(TxtCost.Text) <> 0 Then
     .TextMatrix(Row, .ColIndex("ExPercen")) = val(TxtCost.Text)
    Else
    .TextMatrix(Row, .ColIndex("ExPercen")) = 100
    End If
    End If
        Select Case .ColKey(Col)
            Case "MainDes"
           'If val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) = 0 Then
            If 0 = 0 Then
                   StrAccountCode = .ComboData
                   LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("PrMainDesID"), False, True)
                   .TextMatrix(Row, .ColIndex("PrMainDesID")) = StrAccountCode
                    .TextMatrix(Row, .ColIndex("oprid")) = ""
                    .TextMatrix(Row, .ColIndex("cost")) = 0
                    .TextMatrix(Row, .ColIndex("exe")) = 0
                    .TextMatrix(Row, .ColIndex("percentage")) = 0
                    .TextMatrix(Row, .ColIndex("item_id")) = ""
                    .TextMatrix(Row, .ColIndex("Item")) = ""
                       Else
                    
                    
                    
                   End If
                  
                Set Rs1 = New ADODB.Recordset
                StrSQL = "select FullCode from ProjectMainDes where ProjectID= " & val(DataCombo2.BoundText) & " and  ID =" & val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) & ""
                Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Rs1.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("CodeMain")) = IIf(IsNull(Rs1("FullCode").value), "", Rs1("FullCode").value)
                Else
                .TextMatrix(Row, .ColIndex("CodeMain")) = ""
                End If
           Case "CodeMain"
                   Set Rs1 = New ADODB.Recordset
                StrSQL = "select * from ProjectMainDes where ProjectID= " & val(DataCombo2.BoundText) & " and  FullCode ='" & (.TextMatrix(Row, .ColIndex("CodeMain"))) & "'"
                Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Rs1.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("PrMainDesID")) = IIf(IsNull(Rs1("ID").value), 0, Rs1("ID").value)
                .TextMatrix(Row, .ColIndex("MainDes")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                Else
                .TextMatrix(Row, .ColIndex("MainDes")) = ""
                .TextMatrix(Row, .ColIndex("PrMainDesID")) = 0
                End If
                
            Case "item"
'            If val(.TextMatrix(Row, .ColIndex("oprid"))) = 0 Then
            If 0 = 0 Then
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("oprid"), False, True)
                .TextMatrix(Row, .ColIndex("oprid")) = StrAccountCode
                Else
                StrAccountCode = val(.TextMatrix(Row, .ColIndex("oprid")))
            End If
                If StrAccountCode <> "" And val(StrAccountCode) <> 0 Then
              StrSQL = "SELECT  projects_des.qtySubContractor,projects_des.costSubContractor,  dbo.projects_des.TotalExe, dbo.projects_des.QtyExe, dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, dbo.projects_des.cost, "
              StrSQL = StrSQL & "        dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id,"
              StrSQL = StrSQL & "        dbo.projects_des.esQty, dbo.projects_des.Remark, dbo.projects_des.fullcode, dbo.projects_des.PandUnitID, dbo.TblProcessUnites.UnitName,"
              StrSQL = StrSQL & "        dbo.TblProcessUnites.UnitNamee"
              StrSQL = StrSQL & " FROM         dbo.projects_des LEFT OUTER JOIN"
              StrSQL = StrSQL & "               dbo.TblProcessUnites ON dbo.projects_des.PandUnitID = dbo.TblProcessUnites.UnitID"
              StrSQL = StrSQL & "  WHERE dbo.projects_des.oprid='" & StrAccountCode & "'"
                    StrSQL = StrSQL & " and dbo.projects_des.project_id =" & val(DataCombo2.BoundText)
              
                  '  StrSQL = "SELECT   line_no, oprid,des, net, project_id  ,[unit] ,[Quantity],[Price] ,[Pre_Quantity] ,[Pre_Value],[Pre_Percent] ,[Curr_Quantity]  ,[Curr_value] ,[curr_Percent] ,[tot_quantity] ,[tot_value] ,[tot_percent]   from dbo.projects_des  WHERE fullcode='" & .ComboData & "'"  ' project_id =" & Val(DataCombo2.BoundText) & "and line_no=" & Val(.ComboItem)
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    .TextMatrix(Row, .ColIndex("qty")) = IIf(IsNull(Rs1("qty").value), 0, Rs1("qty").value)
                    If Me.ChQty.value = vbChecked Then
                    .TextMatrix(Row, .ColIndex("quntExc")) = val(.TextMatrix(Row, .ColIndex("qty")))
                    End If
                    
                
                    .TextMatrix(Row, .ColIndex("cost")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
      
                    
                       .TextMatrix(Row, .ColIndex("qtySubContractor")) = IIf(IsNull(Rs1("qtySubContractor").value), 0, Rs1("qtySubContractor").value)
                    .TextMatrix(Row, .ColIndex("costSubContractor")) = IIf(IsNull(Rs1("costSubContractor").value), 0, Rs1("costSubContractor").value)
                    
                    If val(.TextMatrix(Row, .ColIndex("qtySubContractor"))) = 0 Then
                    .TextMatrix(Row, .ColIndex("qtySubContractor")) = .TextMatrix(Row, .ColIndex("qty"))
                    End If
                    
                    
                   If val(.TextMatrix(Row, .ColIndex("costSubContractor"))) = 0 Then
                    .TextMatrix(Row, .ColIndex("costSubContractor")) = .TextMatrix(Row, .ColIndex("exe"))
                    End If
                                
                                
                                
                    .TextMatrix(Row, .ColIndex("total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("cost")))
                    .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
                    
             If billto.ListIndex = 0 Then
                    .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
                 Else
                 .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("costSubContractor").value), 0, Rs1("costSubContractor").value)
                 End If
                 
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("Unit")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                    Else
                    .TextMatrix(Row, .ColIndex("Unit")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
                    End If
                    .TextMatrix(Row, .ColIndex("unit_id")) = IIf(IsNull(Rs1("PandUnitID").value), 0, Rs1("PandUnitID").value)
                    .TextMatrix(Row, .ColIndex("item_id")) = IIf(IsNull(Rs1("fullcode").value), "", Rs1("fullcode").value)
                     .TextMatrix(Row, .ColIndex("discount")) = IIf(IsNull(Rs1("discount").value), 0, Rs1("discount").value)
                     .TextMatrix(Row, .ColIndex("net")) = IIf(IsNull(Rs1("net").value), 0, Rs1("net").value)
                    .TextMatrix(Row, .ColIndex("percentage")) = (val(.TextMatrix(Row, .ColIndex("qty"))) - val(.TextMatrix(Row, .ColIndex("quntExc")))) / 100
                   '  .TextMatrix(Row, .ColIndex("unit")) = IIf(IsNull(Rs1("unit").value), 0, Rs1("unit").value)
                   '   .TextMatrix(Row, .ColIndex("Quantity")) = IIf(IsNull(Rs1("Quantity").value), 0, Rs1("Quantity").value)
                      .TextMatrix(Row, .ColIndex("Pre_Value")) = IIf(IsNull(Rs1("TotalExe").value), 0, Rs1("TotalExe").value)
                      .TextMatrix(Row, .ColIndex("Pre_Quantity")) = IIf(IsNull(Rs1("QtyExe").value), 0, Rs1("QtyExe").value)
                        If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                        GetTermsTotals val(.TextMatrix(Row, .ColIndex("oprid"))), val(txtid.Text), XPDtbTrans.value, netexe, QtyExe, TxtModFlg.Text
                       End If
                       If SystemOptions.AllowNoRoudProjectInvoices = True Then
                       .TextMatrix(Row, .ColIndex("Pre_Quantity")) = val(.TextMatrix(Row, .ColIndex("Pre_Quantity"))) + Round(QtyExe, val(cCompanyInfo.NoRoudProjectInvoices))
                       .TextMatrix(Row, .ColIndex("Pre_Value")) = .TextMatrix(Row, .ColIndex("Pre_Value")) + Round(netexe, val(cCompanyInfo.NoRoudProjectInvoices))
                       Else
                       .TextMatrix(Row, .ColIndex("Pre_Quantity")) = val(.TextMatrix(Row, .ColIndex("Pre_Quantity"))) + Round(QtyExe, 2)
                       .TextMatrix(Row, .ColIndex("Pre_Value")) = .TextMatrix(Row, .ColIndex("Pre_Value")) + Round(netexe, 2)
                       End If
                   '      .TextMatrix(Row, .ColIndex("Pre_Value")) = IIf(IsNull(Rs1("Pre_Value").value), 0, Rs1("Pre_Value").value)
                   '       .TextMatrix(Row, .ColIndex("Pre_Percent")) = IIf(IsNull(Rs1("Pre_Percent").value), 0, Rs1("Pre_Percent").value)
                   '        .TextMatrix(Row, .ColIndex("Curr_Quantity")) = IIf(IsNull(Rs1("Curr_Quantity").value), 0, Rs1("Curr_Quantity").value)
                   ' .TextMatrix(Row, .ColIndex("Curr_value")) = IIf(IsNull(Rs1("Curr_value").value), 0, Rs1("Curr_value").value)
                   ' .TextMatrix(Row, .ColIndex("curr_Percent")) = IIf(IsNull(Rs1("curr_Percent").value), 0, Rs1("curr_Percent").value)
                   ' .TextMatrix(Row, .ColIndex("tot_quantity")) = IIf(IsNull(Rs1("tot_quantity").value), 0, Rs1("tot_quantity").value)
                   ' .TextMatrix(Row, .ColIndex("tot_value")) = IIf(IsNull(Rs1("tot_value").value), 0, Rs1("tot_value").value)
                   ' .TextMatrix(Row, .ColIndex("tot_percent")) = IIf(IsNull(Rs1("tot_percent").value), 0, Rs1("tot_percent").value)
                Else
                    .TextMatrix(Row, .ColIndex("cost")) = 0
                    .TextMatrix(Row, .ColIndex("exe")) = 0
                    .TextMatrix(Row, .ColIndex("percentage")) = 0
                    .TextMatrix(Row, .ColIndex("item_id")) = ""
                    
                    
                   ' .TextMatrix(Row, .ColIndex("unit")) = 0
                   '   .TextMatrix(Row, .ColIndex("Quantity")) = 0
                   '    .TextMatrix(Row, .ColIndex("Price")) = 0
                   '     .TextMatrix(Row, .ColIndex("Pre_Quantity")) = 0
                   '      .TextMatrix(Row, .ColIndex("Pre_Value")) = 0
                   '       .TextMatrix(Row, .ColIndex("Pre_Percent")) = 0
                   '        .TextMatrix(Row, .ColIndex("Curr_Quantity")) = 0
                   '         .TextMatrix(Row, .ColIndex("Curr_value")) = 0
                   ' .TextMatrix(Row, .ColIndex("curr_Percent")) = 0
                   '' .TextMatrix(Row, .ColIndex("tot_quantity")) = 0
                   ' .TextMatrix(Row, .ColIndex("tot_value")) = 0
                   ' .TextMatrix(Row, .ColIndex("tot_percent")) = 0
             
                End If
            
                '     StrSQL = "SELECT   line_no, oprid,des, net, project_id from dbo.projects_des WHERE project_id =" & Val(DataCombo2.BoundText) & "and line_no"
                '    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                '
                '             .TextMatrix(Row, .ColIndex("cost")) = _
                '            IIf(IsNull(Rs("net").value), 0, Rs("net").value)
 
            Case "quntExc"
            .TextMatrix(Row, .ColIndex("percentage")) = (val(.TextMatrix(Row, .ColIndex("qty"))) - val(.TextMatrix(Row, .ColIndex("quntExc")))) / 100
            .TextMatrix(Row, .ColIndex("totEx")) = (val(.TextMatrix(Row, .ColIndex("exe"))) * val(.TextMatrix(Row, .ColIndex("quntExc"))))
           Case "exe"
            .TextMatrix(Row, .ColIndex("percentage")) = (val(.TextMatrix(Row, .ColIndex("qty"))) - val(.TextMatrix(Row, .ColIndex("quntExc")))) / 100
            .TextMatrix(Row, .ColIndex("totEx")) = (val(.TextMatrix(Row, .ColIndex("exe"))) * val(.TextMatrix(Row, .ColIndex("quntExc"))))
            Case "cost"
            .TextMatrix(Row, .ColIndex("total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("cost")))
            .TextMatrix(Row, .ColIndex("net")) = val(.TextMatrix(Row, .ColIndex("total"))) - val(.TextMatrix(Row, .ColIndex("discount")))
            Case "Unit"
                      StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("unit_id"), False, True)
                .TextMatrix(Row, .ColIndex("unit_id")) = StrAccountCode
    
        End Select

        '  Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    CalCultePers Row
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With
ReLineGrid
    ReLineGrid

End Sub
Sub FillAllBandsToGrid()
Dim sql As String
Dim i As Long
Dim rs2 As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim Row As Long

       Dim netexe As Double
       Dim QtyExe As Double
    Dim VATPer  As Double
    Dim oldPerforValue  As Double
    Dim discountHasmyat As Double
    
    linenetaftermainDiscountWithvat = 0
    
' Fg_Journal.Clear flexClearScrollable, flexClearEverything
       '     Fg_Journal.Rows = 1
sql = " SELECT     dbo.projects_des.PrMainDesID, dbo.ProjectMainDes.ProjectID, dbo.ProjectMainDes.Name, dbo.ProjectMainDes.FullCode, dbo.projects_des.des, "
sql = sql & "  dbo.projects_des.oprid"
sql = sql & " FROM         dbo.ProjectMainDes LEFT OUTER JOIN"
sql = sql & "                      dbo.projects_des ON dbo.ProjectMainDes.ID = dbo.projects_des.PrMainDesID"
sql = sql & " Where (dbo.ProjectMainDes.ProjectID = " & val(DataCombo2.BoundText) & ")"

sql = " SELECT     * "
sql = sql & " FROM         dbo.projects_des "
sql = sql & " LEFT OUTER JOIN"
sql = sql & "                      dbo.projects_des ON dbo.projects_des.ID = dbo.projects_des.PrMainDesID"
sql = sql & " Where (dbo.projects_des.Project_ID = " & val(DataCombo2.BoundText) & ")"

sql = " SELECT     dbo.projects_des.PrMainDesID, dbo.ProjectMainDes.ProjectID, dbo.ProjectMainDes.Name, dbo.ProjectMainDes.FullCode, dbo.projects_des.des, "
sql = sql & "  dbo.projects_des.oprid"
sql = sql & " FROM         dbo.ProjectMainDes LEFT OUTER JOIN"
sql = sql & "                      dbo.projects_des ON dbo.ProjectMainDes.ID = dbo.projects_des.PrMainDesID"
sql = sql & " Where (dbo.ProjectMainDes.ProjectID = " & val(DataCombo2.BoundText) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText


Dim LngNewRow  As Long
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then

LngNewRow = val(SetFgForNewRow(Fg_Journal, Fg_Journal.ColIndex("project_id")))
With Fg_Journal
rs2.MoveFirst


If LngNewRow = 0 Then LngNewRow = 1
.Rows = rs2.RecordCount + LngNewRow
For i = LngNewRow To .Rows - 1
.TextMatrix(i, .ColIndex("PrMainDesID")) = IIf(IsNull(rs2("PrMainDesID").value), "", rs2("PrMainDesID").value)
.TextMatrix(i, .ColIndex("MainDes")) = IIf(IsNull(rs2("Name").value), "", rs2("Name").value)
'.TextMatrix(i, .ColIndex("MainDes")) = IIf(IsNull(rs2("PrMainDesID").value), "", rs2("PrMainDesID").value)


.TextMatrix(i, .ColIndex("project_id")) = DataCombo2.BoundText
.TextMatrix(i, .ColIndex("FullCode")) = DataCombo2.Text
.TextMatrix(i, .ColIndex("Projectname")) = txtprojectname
Fg_Journal_StartEdit i, .ColIndex("MainDes"), True
'Fg_Journal_AfterEdit i, .ColIndex("MainDes")
.TextMatrix(i, .ColIndex("item")) = IIf(IsNull(rs2("des").value), "", rs2("des").value)
.TextMatrix(i, .ColIndex("oprid")) = IIf(IsNull(rs2("oprid").value), "", rs2("oprid").value)


'Fg_Journal_StartEdit i, .ColIndex("item"), True
'Fg_Journal_AfterEdit i, .ColIndex("item")

Row = i
  StrSQL = "SELECT  projects_des.qtySubContractor,projects_des.costSubContractor,  dbo.projects_des.TotalExe, dbo.projects_des.QtyExe, dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, dbo.projects_des.cost, "
              StrSQL = StrSQL & "        dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id,"
              StrSQL = StrSQL & "        dbo.projects_des.esQty, dbo.projects_des.Remark, dbo.projects_des.fullcode, dbo.projects_des.PandUnitID, dbo.TblProcessUnites.UnitName,"
              StrSQL = StrSQL & "        dbo.TblProcessUnites.UnitNamee"
              StrSQL = StrSQL & " FROM         dbo.projects_des LEFT OUTER JOIN"
              StrSQL = StrSQL & "               dbo.TblProcessUnites ON dbo.projects_des.PandUnitID = dbo.TblProcessUnites.UnitID"
              StrSQL = StrSQL & "  WHERE dbo.projects_des.oprid='" & .TextMatrix(i, .ColIndex("oprid")) & "'"
                    StrSQL = StrSQL & " and dbo.projects_des.project_id =" & val(DataCombo2.BoundText)
                    Set Rs1 = New ADODB.Recordset
                  '  StrSQL = "SELECT   line_no, oprid,des, net, project_id  ,[unit] ,[Quantity],[Price] ,[Pre_Quantity] ,[Pre_Value],[Pre_Percent] ,[Curr_Quantity]  ,[Curr_value] ,[curr_Percent] ,[tot_quantity] ,[tot_value] ,[tot_percent]   from dbo.projects_des  WHERE fullcode='" & .ComboData & "'"  ' project_id =" & Val(DataCombo2.BoundText) & "and line_no=" & Val(.ComboItem)
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    If Not Rs1.EOF Then
                      .TextMatrix(Row, .ColIndex("qty")) = IIf(IsNull(Rs1("qty").value), 0, Rs1("qty").value)
                      If Me.ChQty.value = vbChecked Then
                      .TextMatrix(Row, .ColIndex("quntExc")) = val(.TextMatrix(Row, .ColIndex("qty")))
                      End If
                      
                  
                      .TextMatrix(Row, .ColIndex("cost")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
        
                      
                         .TextMatrix(Row, .ColIndex("qtySubContractor")) = IIf(IsNull(Rs1("qtySubContractor").value), 0, Rs1("qtySubContractor").value)
                      .TextMatrix(Row, .ColIndex("costSubContractor")) = IIf(IsNull(Rs1("costSubContractor").value), 0, Rs1("costSubContractor").value)
                      
                      If val(.TextMatrix(Row, .ColIndex("qtySubContractor"))) = 0 Then
                      .TextMatrix(Row, .ColIndex("qtySubContractor")) = .TextMatrix(Row, .ColIndex("qty"))
                      End If
                      
                      
                     If val(.TextMatrix(Row, .ColIndex("costSubContractor"))) = 0 Then
                      .TextMatrix(Row, .ColIndex("costSubContractor")) = .TextMatrix(Row, .ColIndex("exe"))
                      End If
                                  
                                  
                                  
                      .TextMatrix(Row, .ColIndex("total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("cost")))
                      .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
                      
               If billto.ListIndex = 0 Then
                      .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
                   Else
                   .TextMatrix(Row, .ColIndex("exe")) = IIf(IsNull(Rs1("costSubContractor").value), 0, Rs1("costSubContractor").value)
                   End If
                   
                      If SystemOptions.UserInterface = ArabicInterface Then
                      .TextMatrix(Row, .ColIndex("Unit")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                      Else
                      .TextMatrix(Row, .ColIndex("Unit")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
                      End If
                      .TextMatrix(Row, .ColIndex("unit_id")) = IIf(IsNull(Rs1("PandUnitID").value), 0, Rs1("PandUnitID").value)
                      .TextMatrix(Row, .ColIndex("item_id")) = IIf(IsNull(Rs1("fullcode").value), "", Rs1("fullcode").value)
                       .TextMatrix(Row, .ColIndex("discount")) = IIf(IsNull(Rs1("discount").value), 0, Rs1("discount").value)
                       .TextMatrix(Row, .ColIndex("net")) = IIf(IsNull(Rs1("net").value), 0, Rs1("net").value)
                      .TextMatrix(Row, .ColIndex("percentage")) = (val(.TextMatrix(Row, .ColIndex("qty"))) - val(.TextMatrix(Row, .ColIndex("quntExc")))) / 100
                     '  .TextMatrix(Row, .ColIndex("unit")) = IIf(IsNull(Rs1("unit").value), 0, Rs1("unit").value)
                     '   .TextMatrix(Row, .ColIndex("Quantity")) = IIf(IsNull(Rs1("Quantity").value), 0, Rs1("Quantity").value)
                        .TextMatrix(Row, .ColIndex("Pre_Value")) = IIf(IsNull(Rs1("TotalExe").value), 0, Rs1("TotalExe").value)
                        .TextMatrix(Row, .ColIndex("Pre_Quantity")) = IIf(IsNull(Rs1("QtyExe").value), 0, Rs1("QtyExe").value)
                          If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                          GetTermsTotals val(.TextMatrix(Row, .ColIndex("oprid"))), val(txtid.Text), XPDtbTrans.value, netexe, QtyExe, TxtModFlg.Text
                         End If
                         If SystemOptions.AllowNoRoudProjectInvoices = True Then
                         .TextMatrix(Row, .ColIndex("Pre_Quantity")) = val(.TextMatrix(Row, .ColIndex("Pre_Quantity"))) + Round(QtyExe, val(cCompanyInfo.NoRoudProjectInvoices))
                         .TextMatrix(Row, .ColIndex("Pre_Value")) = .TextMatrix(Row, .ColIndex("Pre_Value")) + Round(netexe, val(cCompanyInfo.NoRoudProjectInvoices))
                         Else
                         .TextMatrix(Row, .ColIndex("Pre_Quantity")) = val(.TextMatrix(Row, .ColIndex("Pre_Quantity"))) + Round(QtyExe, 2)
                         .TextMatrix(Row, .ColIndex("Pre_Value")) = .TextMatrix(Row, .ColIndex("Pre_Value")) + Round(netexe, 2)
                         End If
                     '      .TextMatrix(Row, .ColIndex("Pre_Value")) = IIf(IsNull(Rs1("Pre_Value").value), 0, Rs1("Pre_Value").value)
                     '       .TextMatrix(Row, .ColIndex("Pre_Percent")) = IIf(IsNull(Rs1("Pre_Percent").value), 0, Rs1("Pre_Percent").value)
                     '        .TextMatrix(Row, .ColIndex("Curr_Quantity")) = IIf(IsNull(Rs1("Curr_Quantity").value), 0, Rs1("Curr_Quantity").value)
                     ' .TextMatrix(Row, .ColIndex("Curr_value")) = IIf(IsNull(Rs1("Curr_value").value), 0, Rs1("Curr_value").value)
                     ' .TextMatrix(Row, .ColIndex("curr_Percent")) = IIf(IsNull(Rs1("curr_Percent").value), 0, Rs1("curr_Percent").value)
                     ' .TextMatrix(Row, .ColIndex("tot_quantity")) = IIf(IsNull(Rs1("tot_quantity").value), 0, Rs1("tot_quantity").value)
                     ' .TextMatrix(Row, .ColIndex("tot_value")) = IIf(IsNull(Rs1("tot_value").value), 0, Rs1("tot_value").value)
                     ' .TextMatrix(Row, .ColIndex("tot_percent")) = IIf(IsNull(Rs1("tot_percent").value), 0, Rs1("tot_percent").value)
            End If
                  '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    CalCultePers Row
        
ReLineGrid
rs2.MoveNext
Next i
End With
End If

End Sub

Sub CalCultePers(Optional i As Long)
With Fg_Journal
If .TextMatrix(i, .ColIndex("item")) <> "" Then
If val(.TextMatrix(i, .ColIndex("ExPercen"))) = 0 Then
    If val(TxtCost.Text) <> 0 Then
     .TextMatrix(i, .ColIndex("ExPercen")) = val(TxtCost.Text)
    Else
    .TextMatrix(i, .ColIndex("ExPercen")) = 100
    End If
 End If

If val(DcbExPercen.ListIndex) = 0 Then
.TextMatrix(i, .ColIndex("totEx")) = .TextMatrix(i, .ColIndex("ExPercen"))
.TextMatrix(i, .ColIndex("TotalApprov")) = .TextMatrix(i, .ColIndex("ExPercen"))
Else
.TextMatrix(i, .ColIndex("totEx")) = (val(.TextMatrix(i, .ColIndex("ExPercen"))) * val(.TextMatrix(i, .ColIndex("quntExc"))) * val(.TextMatrix(i, .ColIndex("exe")))) / 100
.TextMatrix(i, .ColIndex("TotalApprov")) = (val(.TextMatrix(i, .ColIndex("ExPercen"))) * val(.TextMatrix(i, .ColIndex("QtyApprov"))) * val(.TextMatrix(i, .ColIndex("PriceApprov")))) / 100
End If
'If val(.TextMatrix(i, .ColIndex("QtyApprov"))) = 0 Or SystemOptions.AllowChangePriceApprove = False Then
'.TextMatrix(i, .ColIndex("QtyApprov")) = .TextMatrix(i, .ColIndex("quntExc"))
'End If
'If val(.TextMatrix(i, .ColIndex("PriceApprov"))) = 0 Or SystemOptions.AllowChangePriceApprove = False Then
'.TextMatrix(i, .ColIndex("PriceApprov")) = .TextMatrix(i, .ColIndex("exe"))
'End If
If val(.TextMatrix(i, .ColIndex("TotalApprov"))) = 0 Then
.TextMatrix(i, .ColIndex("TotalApprov")) = .TextMatrix(i, .ColIndex("totEx"))
End If
If val(.TextMatrix(i, .ColIndex("DiscApprov"))) = 0 Then
.TextMatrix(i, .ColIndex("DiscApprov")) = .TextMatrix(i, .ColIndex("discountEXE"))
End If
End If
.TextMatrix(i, .ColIndex("NetExe")) = val(.TextMatrix(i, .ColIndex("totEx"))) - val(.TextMatrix(i, .ColIndex("discountEXE")))
.TextMatrix(i, .ColIndex("NetApprov")) = val(.TextMatrix(i, .ColIndex("TotalApprov"))) - val(.TextMatrix(i, .ColIndex("DiscApprov")))
.TextMatrix(i, .ColIndex("totEx")) = val(.TextMatrix(i, .ColIndex("qtySubContractor"))) * val(.TextMatrix(i, .ColIndex("costSubContractor")))

.TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Qty"))) * val(.TextMatrix(i, .ColIndex("cost")))
.TextMatrix(i, .ColIndex("netexe")) = val(.TextMatrix(i, .ColIndex("qtySubContractor"))) * val(.TextMatrix(i, .ColIndex("costSubContractor")))

End With
   
End Sub
Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If
        Select Case .ColKey(Col)
               
                Case "PriceApprov", "QtyApprov"
            '    If SystemOptions.AllowChangePriceApprove = True Then
                .ComboList = ""
            '    Else
                Cancel = True
            '    End If
                
                    Case "qtySubContractor"
              '  Cancel = True
                    Case "costSubContractor"
              '  Cancel = True
                
          Case "percentage"
                Cancel = True
         Case "net"
                Cancel = True
                
         Case "discount"
                Cancel = True
        Case "LineNo"
                Cancel = True
          
                Case "cost"
'                If SystemOptions.AllowChanProjectBillPrice = True Then
'                .ComboList = ""
'                Else
'                Cancel = True
'                End If
                Case "total"
                Cancel = True
                
                Case "item_id"
                Cancel = True
                Case "qty"
              '  Cancel = True
              Case "CodeMain"
                .ComboList = ""
            Case "item"
                .ComboList = ""

            Case "cost"
                .ComboList = ""
        
            Case "exe"
                .ComboList = ""
        ''///////////////////
       '  Case "QtyApprov"
       '         .ComboList = ""
         Case "TotalApprov"
               Cancel = True
        ' Case "PriceApprov"
        '        .ComboList = ""
         Case "DiscApprov"
                .ComboList = ""
          Case "NetApprov"
                Cancel = True
                ''///////////
            Case "exedate"
                .ComboList = ""
                '  Cancel = True
                
             Case "exedate"
                .ComboList = ""
                
                 Case "unit"
                .ComboList = ""
                
                 Case "Quantity"
                .ComboList = ""
                
                 Case "Price"
                .ComboList = ""
                
                 Case "Pre_Quantity"
                .ComboList = ""
                
                 Case "Pre_Value"
                .ComboList = ""
                
                 Case "Pre_Percent"
                .ComboList = ""
                
                 Case "Curr_Quantity"
                .ComboList = ""
                
                 Case "Curr_value"
                .ComboList = ""
                
                 Case "curr_Percent"
                .ComboList = ""
                
                 Case "tot_quantity"
                .ComboList = ""
            
            Case "tot_value"
                .ComboList = ""
                
                Case "tot_percent"
                .ComboList = ""
            Case "quntExc"
                .ComboList = ""
                Case "exe"
                .ComboList = ""
                Case "totEx"
                .ComboList = ""
                
        End Select

    End With
 
End Sub

Private Sub Fg_Journal_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

If Col = 1 Then
  Frame16.Visible = True
  End If
End Sub

Private Sub Fg_Journal_Click()
    current_terms = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("item_id"))
    
    
  
End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)

    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    
  Dim Rs4 As New ADODB.Recordset
    Dim StrComboList_1 As String
     Dim StrSQL_2 As String
     
    Dim Msg As String

    With Fg_Journal

        Select Case .ColKey(Col)
            Case "MainDes"
                     StrSQL = " SELECT     Name, ID"
                     StrSQL = StrSQL & "       FROM         dbo.ProjectMainDes WHERE ProjectID =" & val(DataCombo2.BoundText)
        
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Name", "ID")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
                
            Case "item"
       
                'Full Path Display
                StrSQL = "SELECT  oprid, fullcode,line_no, oprid,des, net, project_id from dbo.projects_des WHERE project_id =" & val(DataCombo2.BoundText)
                If DcbosubContractor.Text <> "" And val(DcbosubContractor.BoundText) <> 0 Then
                StrSQL = StrSQL & " and  sub_contractor_id =" & val(DcbosubContractor.BoundText) & ""
                End If
                If val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) <> 0 Then
                StrSQL = StrSQL & " and  PrMainDesID =" & val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) & ""
                End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "des", "oprid")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
              Case "Unit"
              
                 StrSQL_2 = "SELECT    UnitID  ,UnitName      ,UnitNamee  FROM TblProcessUnites"
                Rs4.Open StrSQL_2, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList_1 = Fg_Journal.BuildComboList(Rs4, "UnitName", "UnitID")

                If StrComboList_1 <> "" Then
                    StrComboList_1 = "|" & StrComboList_1
                End If

                .ComboList = StrComboList_1
                    
        End Select

    End With

End Sub

Private Sub Form_Load()
    On Error Resume Next
    TxtModFlg.Text = "R"
    Dim StrSQL As String
    Dim my_language As String
    Set rs = New ADODB.Recordset
    
      If SystemOptions.AllowEditVaTManulay = True Then
txtManulaVat.Enabled = True
txtManulaVat.Visible = True
Else
txtManulaVat.Enabled = False
txtManulaVat.Text = 0
txtManulaVat.Visible = False
End If

  '  StrSQL = "SELECT  dueDate,  branch_no,NoteSerial, id, bill_date, project_no, project_name, Sub_user_name, End_user_name, End_user_account, bill_to, Sub_user_account, bill_type, revenue_account, note_id,"
  '  StrSQL = StrSQL + "total  From dbo.SubcontractorContract Order by ID"
    
  'StrSQL = "SELECT  dueDate,  branch_no,NoteSerial, id, bill_date, project_no, project_name, Sub_user_name, End_user_name, End_user_account, bill_to, Sub_user_account, bill_type, revenue_account, note_id,"
    
    
    
    StrSQL = StrSQL + "SELECT *  From dbo.SubcontractorContract  where 1=1"
    StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ") Order by ID"
    
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    '
   If SystemOptions.UserInterface = ArabicInterface Then
   With DcbExPercen
   .Clear
   .AddItem "ﬁÌ„…"
   .AddItem "‰”»…"
   End With
   With DcbPeriodType
   .Clear
   .AddItem "ÌÊ„"
   .AddItem "‘Â—"
   .AddItem "”‰…"
   End With
   Else
   With DcbExPercen
   .Clear
   .AddItem "Value"
   .AddItem "Percentage"
   End With
      With DcbPeriodType
   .Clear
   .AddItem "Day"
   .AddItem "Month"
   .AddItem "Year"
   End With
   End If
With bill_Type
.Clear
.AddItem "«Ì—«œ"
.AddItem "«Ì—«œ „” Õﬁ"
End With
    'first_run = True
    Dim My_SQL As String
 
    My_SQL = "  select id,Fullcode from Projects where not (Fullcode is null) and Fullcode <>N'""' "
    My_SQL = My_SQL & "  AND      branch_no in(" & Current_branchSql & ")"
    fill_combo DataCombo2, My_SQL

    Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
     Dcombos.GetBranches dcBranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetPersons Me.DcbosubContractor
    
    


Set Dcombos = New ClsDataCombos
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select CusID,CusName From TblCustemers"
Else
    StrSQL = "Select CusID,CusNamee From TblCustemers"
End If
StrSQL = StrSQL & " Where Type = 3"
'StrSQL = StrSQL & " where CusID in(SELECT     sub_contractor_id"
'StrSQL = StrSQL & " From dbo.projects_des)"
'StrSQL = StrSQL & " WHERE     (project_id = " & project_no & "))"
'Dcombos.ClearMyDataCombo DcbosubContractor

'fill_combo Me.DcbosubContractor, StrSQL

Dcombos.GetPersons Me.DcbosubContractor

    If my_language = "E" Then
        CMD_language.ToolTipText = "Change Language"

        'Me.dept_lbl = departement_nam
        'Me.emp_name_lbl = current_user_name
        InfoE.Visible = True
        infoA.Visible = False
    Else

        'emp_a.Caption = current_user_name
        'dep_a.Caption = departement_name
   
        infoA.Visible = True
        InfoE.Visible = False
    End If

    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    'LoadSettings
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    Set NewGrid.Grid = fg
    'NewGrid.GridTrans = Destruction
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    'Set NewGrid.DtpBillDate = Me.XPDtbBill
    'Set NewGrid.StoreName = Me.DCboStoreName
    'Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName

    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.txtPrice = txtPrice
    NewGrid.FillGrid
    
    
    ReloadContrac (val(DataCombo2.BoundText))
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    Frame7.Visible = False
If SystemOptions.UserInterface = EnglishInterface Then
    ChangeLang
 End If
    
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Command1_Click (0)
    End If

End Sub

Function ChangeLang()

    If SystemOptions.UserInterface = EnglishInterface Then
    ChQty.Caption = "UnEqual Qty"
    Label34.Caption = "Discounts"
    Label40.Caption = "No.Bills"
    Cmd(10).Caption = "Close"
    Command1(15).Caption = "Issuing Bills"
    Label42.Caption = "Issuing Bills"
    lbl(1).Caption = "Start"
    ALLButton2.Caption = "Add All Terms"
    Label39.Caption = "Period"
    Label41.Caption = "Remarks"
    Label43.Caption = "Type"
    Label1(66).Caption = "VAT %"
    Label1(67).Caption = "VAT"
    Label1(68).Caption = "Total"
    Check1.Caption = "Select All"
    Command1(16).Caption = "Print "
    Label58.Caption = "Beginning Project"
    Label60.Caption = "Net Value"
    Label59.Caption = "Perfor.Discount"
With VSFlexGrid4
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
'.TextMatrix(0, .ColIndex("InstalValue")) = "Installment Value"
.TextMatrix(0, .ColIndex("payed")) = "Select"
.TextMatrix(0, .ColIndex("NoteSerial1")) = "No"
.TextMatrix(0, .ColIndex("too")) = "No.Manual"
.TextMatrix(0, .ColIndex("NoteDate")) = "Date"
.TextMatrix(0, .ColIndex("branch_name")) = "Branch"
.TextMatrix(0, .ColIndex("Valu")) = "Original Value"
.TextMatrix(0, .ColIndex("Note_Value")) = "Total Value"
.TextMatrix(0, .ColIndex("PayedValue")) = "Payed Value"
.TextMatrix(0, .ColIndex("RemainingValue")) = "Remaining"
.TextMatrix(0, .ColIndex("TransPayedValue")) = "Payed Trans"
.TextMatrix(0, .ColIndex("NetValue")) = "Net Value"
.TextMatrix(0, .ColIndex("VAT")) = "VAT"
End With
Label49.Caption = "Value"
Label50.Caption = "VAT"
Label51.Caption = "Total"
Label52.Caption = "Payed Value"
Label53.Caption = "Remaining"
Label54.Caption = "Total Value"
Label55.Caption = "Net Value"

Label45.Caption = "Pre VAT"
ALLButton1.Caption = "Show Pre Payments"
Frame15.Caption = "Data of Pre Payments"
Command10.Caption = "Cancel"
Label46.Caption = "Total"
With cboDiscount1
.Clear
.AddItem "NA"
.AddItem "Perce%"
.AddItem "Value"
End With

With cboDiscount2
.Clear
.AddItem "NA"
.AddItem "Perce%"
.AddItem "Value"
End With

With billto
.Clear
.AddItem "End User"
.AddItem "Sub-Contractor"
 
End With

With Me.bill_Type

.Clear
.AddItem "Revenue"
.AddItem "Issue Revenue"
 
End With

  lbl(20).Caption = "Current Record"
    lbl(21).Caption = "NO. Recordes"
lbl(22).Caption = "By"
 

        temp = XPBtnMove(1).Left
        XPBtnMove(1).Left = XPBtnMove(2).Left
        XPBtnMove(2).Left = temp
Label26.Caption = "Branch"

        temp = XPBtnMove(0).Left
        XPBtnMove(0).Left = XPBtnMove(3).Left
        XPBtnMove(3).Left = temp
        SetInterface Me
        Me.Caption = "         Project Invoice"
        Label9.Caption = Me.Caption

        Label20.Caption = "Bill No."
        Label25.Caption = "Date"

        Label6.Caption = "Project Code"
        Label1(0).Caption = "Project Name"
         Label15.Caption = "End User"
        Label23.Caption = "Sub-Contractor"
        Label18.Caption = "Bill To"
        Label30.Caption = "Bill Type"
        Label8.Caption = "To Date"
        Label29.Caption = "Total"
        Label17.Caption = "Notes"

        Frame14.Caption = "Labors Data"
  
        DataGrid1.RightToLeft = False
        CMD_language.Caption = "⁄—»Ì"
        Frame4.Visible = True
        Frame3.Visible = True
        Frame8.Visible = True
  
   
        Adodc1.Caption = "move"
  
        With Fg_Journal
            .TextMatrix(0, .ColIndex("ExPercen")) = "Percentage"
            .TextMatrix(0, .ColIndex("LineNo")) = "Index"
            .TextMatrix(0, .ColIndex("Item_ID")) = "Term#"
            .TextMatrix(0, .ColIndex("QtyApprov")) = "Approved Qty"
            .TextMatrix(0, .ColIndex("TotalApprov")) = "Approved Total"
            .TextMatrix(0, .ColIndex("PriceApprov")) = "Approved Price"
            .TextMatrix(0, .ColIndex("DiscApprov")) = "Approved Discount"
            .TextMatrix(0, .ColIndex("NetApprov")) = "Approved Net"
            .TextMatrix(0, .ColIndex("MainDes")) = "Main Term Desc."
            .TextMatrix(0, .ColIndex("CodeMain")) = "Code"
            .TextMatrix(0, .ColIndex("item")) = "Term Desc."
            .TextMatrix(0, .ColIndex("cost")) = "Cost"
            .TextMatrix(0, .ColIndex("exe")) = "Price"
            .TextMatrix(0, .ColIndex("percentage")) = "Percentage"
            .TextMatrix(0, .ColIndex("exedate")) = "Exe Date"

  .TextMatrix(0, .ColIndex("Unit")) = "Unit"
  .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
.TextMatrix(0, .ColIndex("Price")) = "Price"
.TextMatrix(0, .ColIndex("Pre_Quantity")) = "Pre. Exe. Quantity"
.TextMatrix(0, .ColIndex("Pre_Value")) = "Pre. Exe Value "
.TextMatrix(0, .ColIndex("Pre_Percent")) = "Pre. Exe Percentage"
.TextMatrix(0, .ColIndex("Curr_Quantity")) = " Current Exe Quantity"
.TextMatrix(0, .ColIndex("Curr_value")) = " Current Exe Value"
.TextMatrix(0, .ColIndex("curr_Percent")) = "Current Exe Percentage"
.TextMatrix(0, .ColIndex("tot_quantity")) = "Total Quantity"
.TextMatrix(0, .ColIndex("tot_value")) = "Total Value"
.TextMatrix(0, .ColIndex("tot_percent")) = "Total Percent"

.TextMatrix(0, .ColIndex("net")) = "Net"
.TextMatrix(0, .ColIndex("qty")) = "Act. Qty"
.TextMatrix(0, .ColIndex("total")) = "Total"
.TextMatrix(0, .ColIndex("discount")) = "Discount"
.TextMatrix(0, .ColIndex("quntExc")) = "Exc Qty"
.TextMatrix(0, .ColIndex("totEx")) = "Exc Total"
.TextMatrix(0, .ColIndex("tot_percent")) = "Exc Qty"
.TextMatrix(0, .ColIndex("discountEXE")) = "EXc. Dis. "

.TextMatrix(0, .ColIndex("netexe")) = "Exc Net"
.TextMatrix(0, .ColIndex("percentage")) = "Exc Qty%"
.TextMatrix(0, .ColIndex("percentage1")) = "Exc Value%"

.TextMatrix(0, .ColIndex("Pre_Percent")) = "Pre Exc Qty%"
.TextMatrix(0, .ColIndex("Pre_Percent1")) = "Pre Exc Value%"

 .TextMatrix(0, .ColIndex("tot_percent")) = "Total Exc Qty%"
.TextMatrix(0, .ColIndex("tot_percent1")) = "Total Exc Value%"
 

        End With

        opr_items(0).Caption = "View Term Operations"
        Frame11.Caption = "Term Operaions"
 
        Label27.Caption = "Labors Count"
        Label24.Caption = "Total"

        With VSFlexGrid1
            .TextMatrix(0, .ColIndex("LineNo")) = "Index"
            .TextMatrix(0, .ColIndex("code")) = "Labor Code"
            .TextMatrix(0, .ColIndex("name")) = "name"

            .TextMatrix(0, .ColIndex("jobname")) = "Job"
            .TextMatrix(0, .ColIndex("daysalary")) = "Day Salary"
            .TextMatrix(0, .ColIndex("Start")) = "Start"
            .TextMatrix(0, .ColIndex("End")) = "End"
            .TextMatrix(0, .ColIndex("Count")) = "No Of Days"
            .TextMatrix(0, .ColIndex("total")) = "Total"

        End With

        With VSFlexGrid2
            .TextMatrix(0, .ColIndex("LineNo")) = "Index"
            .TextMatrix(0, .ColIndex("fullcode")) = "OPR Code"

            .TextMatrix(0, .ColIndex("name")) = "Operation Desc."
            .TextMatrix(0, .ColIndex("total_items")) = "Total Items Cost"
            .TextMatrix(0, .ColIndex("total_salary")) = "Total Salary"
            .TextMatrix(0, .ColIndex("total_expenses")) = "Total Expenses"
            .TextMatrix(0, .ColIndex("total")) = "Total"
            .TextMatrix(0, .ColIndex("total_items1")) = "Total Items Cost EXE"
            .TextMatrix(0, .ColIndex("total_salary1")) = "Total Salary EXE"
            .TextMatrix(0, .ColIndex("total_expenses1")) = "Total Expenses EXE"
            .TextMatrix(0, .ColIndex("total1")) = "Total EXE"

        End With

        CmdRemove.Caption = "Remove Line"
        Show_items(0).Caption = "Show Items"
        employee_details(0).Caption = "Show Labors"
        employee_details(1).Caption = "Return To OPR"
        opr_expenses(0).Caption = "Show Expenses"
        Label28.Caption = "Total"
        opr_items(1).Caption = "Retuen To Term"

        Frame12.Caption = "Expenses"
        opr_expenses(1).Caption = "Return To Opr"
        lbl(6).Caption = "Total Expenses"
Command1(12).Caption = "Same Copy"

        With Me.VSFlexGrid3
            .TextMatrix(0, .ColIndex("LineNo")) = "Index"
            .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Names"
            .TextMatrix(0, .ColIndex("value")) = "Value"

            .TextMatrix(0, .ColIndex("des")) = "Des"
 
        End With
Label21.Caption = "Due Date"
Label32.Caption = "Deduct Adv. Payment"
Label31.Caption = "Deduct Ensure Business "
Label22.Caption = "Sub-contractor"
Label33.Caption = "Manual No."

        Frame1.Caption = "OPR Items"
        lbl(31).Caption = "Item Code"
        lbl(30).Caption = "Item Name"
        lbl(29).Caption = "Status"
        lbl(28).Caption = "Serial"
        lbl(27).Caption = "QTY"
        lbl(26).Caption = "Price"
        'lbl(0).Caption = "Avilable"
        'lbl(1).Caption = "Reserved"
        'lbl(3).Caption = "ON order"
        lbl(2).Caption = "Total"
        Command1(3).Caption = "Edit"
        Command1(9).Caption = "Delete"
        Command1(6).Caption = "Undo"
        Command1(8).Caption = "Print Jl Entery"
        Command1(7).Caption = "Print Bill "
        Command1(10).Caption = " Search "
 
        opr_items(1).Caption = "Return To Opr"
        Show_items(1).Caption = "Return To Opr"
        Label5.Caption = "Entry No."
        Frame2.Caption = "Terms"
        Shape1.Visible = False
        lbl(4).Visible = False
        lbl(5).Visible = False
        ' Me.Width = 10000
    Else
        billto.Clear
        billto.AddItem "⁄„Ì· ‰Â«∆Ì"
        billto.AddItem "„ﬁ«Ê· »«ÿ‰"
        bill_Type.Clear
        bill_Type.AddItem "«Ì—«œ« "
        bill_Type.AddItem "«Ì—«œ«  „” Õﬁ…"
 
    End If

     Command1(0).Caption = "New"
        Command1(1).Caption = "Save"
        Command1(2).Caption = "Attachments"
        Command1(3).Caption = "Edit"
        Command1(9).Caption = "Delete"
  
        SuperLabel2.Text = "Search"
        Command1(4).Caption = "By ID"
        Command1(5).Caption = "Search"
  Command1(11).Caption = "Attachement"


Label29.Caption = "Total"
Label35.Caption = "Discount"
Label36.Caption = "Net"
Label37.Caption = "Advanced"

End Function

Private Sub retrive1(Item_ID As String)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
    VSFlexGrid2.Enabled = True
    txt_opr_total.Text = 0
          
    StrSQL = "select * from terms_operations_project_bill where term_fullcode='" & Item_ID & "' and bill_id=" & val(Me.txtid.Text)
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid2
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsDev("fullcode").value), "", RsDev("fullcode").value)
            
                .TextMatrix(i, .ColIndex("item_id")) = IIf(IsNull(RsDev("item_id").value), "", RsDev("item_id").value)
            
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
            
                .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("id").value), "", RsDev("id").value)
 
                .TextMatrix(i, .ColIndex("period")) = IIf(IsNull(RsDev("period").value), "", RsDev("period").value)
                .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDev("count").value), "", RsDev("count").value)
            
                .TextMatrix(i, .ColIndex("salary")) = IIf(IsNull(RsDev("salary").value), "", RsDev("salary").value)
 
                .TextMatrix(i, .ColIndex("total_items")) = IIf(IsNull(RsDev("total_items").value), "", RsDev("total_items").value)
                .TextMatrix(i, .ColIndex("total_salary")) = IIf(IsNull(RsDev("total_salary").value), "", RsDev("total_salary").value)
                .TextMatrix(i, .ColIndex("total_expenses")) = IIf(IsNull(RsDev("total_expenses").value), "", RsDev("total_expenses").value)
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), "", RsDev("total").value)
            
                .TextMatrix(i, .ColIndex("total_items1")) = IIf(IsNull(RsDev("total_items1").value), "", RsDev("total_items1").value)
                .TextMatrix(i, .ColIndex("total_salary1")) = IIf(IsNull(RsDev("total_salary1").value), "", RsDev("total_salary1").value)
                .TextMatrix(i, .ColIndex("total_expenses1")) = IIf(IsNull(RsDev("total_expenses1").value), "", RsDev("total_expenses1").value)
                .TextMatrix(i, .ColIndex("total1")) = IIf(IsNull(RsDev("total1").value), "", RsDev("total1").value)
            
                RsDev.MoveNext
            Next i

            Me.txt_opr_total.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        
        End With

    End If
          
    ReLineGrid

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub ImgFavorites_Click()
    AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
C1Elastic3.Visible = True

End Sub

Private Sub Label48_Click()
Frame15.Visible = False
End Sub

Private Sub Label61_Click()
C1Elastic3.Visible = False
End Sub

Private Sub opr_expenses_Click(Index As Integer)

    Select Case Index

        Case 0
  
            VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.Rows = 2
            VSFlexGrid3.Enabled = True

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame12.Visible = True

                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))
                Retrive3 current_opr

                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame1.Caption = "„’«—Ì› «·⁄„·Ì… —ﬁ„ :   " & "  " & current_opr
                Else
                    Frame1.Caption = "Expenses For Operation No: " & "  " & current_opr
                End If

                XPTxtSum.Text = 0
            End If

        Case 1
  
            Frame12.Visible = False
    End Select

End Sub

Private Sub Retrive4(current_opr As String)
    Dim RsDev As ADODB.Recordset
 
    StrSQL = "SELECT  * from opr_employee_details  where  (opr_type=0 or opr_type=3)  and opr_Fullcode='" & current_opr & "' and  (Start_date<='" & SQLDate(DTPicker1.value) & "')"
  
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsDev.RecordCount > 0 Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
   
            .Rows = .FixedRows + RsDev.RecordCount
   
            For i = .FixedRows To .Rows - 1
            
                .TextMatrix(i, .ColIndex("LineNo")) = i
    
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(RsDev("Emp_code").value), "", RsDev("Emp_code").value)
            
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("emp_name").value), "", RsDev("emp_name").value)
            
                .TextMatrix(i, .ColIndex("jobname")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                .TextMatrix(i, .ColIndex("daysalary")) = IIf(IsNull(RsDev("daysalary").value), "", RsDev("daysalary").value)
            
                .TextMatrix(i, .ColIndex("Start")) = IIf(IsNull(RsDev("Start_date").value), "", RsDev("Start_date").value)

                If DateDiff("d", IIf(IsNull(RsDev("end_date").value), Date, RsDev("end_date").value), DTPicker1.value) >= 0 Then
            
                    .TextMatrix(i, .ColIndex("End")) = IIf(IsNull(RsDev("end_date").value), Date, RsDev("end_date").value)
                Else
                    .TextMatrix(i, .ColIndex("End")) = DTPicker1.value
                End If
  
                .TextMatrix(i, .ColIndex("Count")) = DateDiff("d", .TextMatrix(i, .ColIndex("Start")), .TextMatrix(i, .ColIndex("End")))
                .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("daysalary")))
 
                RsDev.MoveNext
            Next i

            '  If RsDev.RecordCount > 0 Then
            Me.txt_emp_salary.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
            Me.txt_employee_count.Text = .Aggregate(flexSTCount, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
            '   End If
        End With

    End If

End Sub

Private Sub Retrive3(current_opr As String)
    Dim RsDev As ADODB.Recordset
 
    StrSQL = "SELECT  * from gl_cc  where  bill_id is null   and  recorddate<='" & SQLDate(DTPicker1.value) & "' and opr_fullcode='" & current_opr & "'"
  
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsDev.RecordCount > 0 Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid3
   
            .Rows = .FixedRows + RsDev.RecordCount
   
            For i = .FixedRows To .Rows - 1
            
                .TextMatrix(i, .ColIndex("LineNo")) = i
            
                '              .TextMatrix(I, .ColIndex("ExpensesID")) = IIf(IsNull(RsDev("ExpensesID").value), _
                '      "", RsDev("ExpensesID").value)
            
                '  .TextMatrix(I, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), _
                '      "", RsDev("AccountCode").value)
            
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
   
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
                Dim des As String

                If SystemOptions.UserInterface = ArabicInterface Then
                    des = "»‰«¡ ⁄·Ï "
                Else
                    des = "Based On"
                End If
          
                des = des & "  " & IIf(IsNull(RsDev("NotesTypeName").value), "", RsDev("NotesTypeName").value)
         
                If SystemOptions.UserInterface = ArabicInterface Then
                    des = des & "  »—ﬁ„  "
                Else
                    des = "  NO :"
                End If
          
                des = des & "  " & IIf(IsNull(RsDev("NoteSerial1").value), "", RsDev("NoteSerial1").value)
            
                .TextMatrix(i, .ColIndex("des")) = des
                RsDev.MoveNext
            Next i

            '  If RsDev.RecordCount > 0 Then
            Me.txt_expenses_total.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
            '   End If
        End With

    End If

End Sub

Private Sub opr_items_Click(Index As Integer)

    Select Case Index

        Case 0

            DTPicker1.value = Date

            If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("item_id")) = "" Then
                Frame11.Visible = True
        
                current_terms = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("item_id"))
                retrive1 current_terms

                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame11.Caption = "⁄„·Ì«  «·»‰œ —ﬁ„ : " & current_terms
                Else
                    Frame11.Caption = "Operations For Term NO:" & current_terms
                End If
            End If

        Case 1
            ReLineGrid current_terms
'            StrSQL = "Delete From terms_operations_project_bill Where term_fullcode ='" & current_terms & "' and bill_id=" & val(Me.txtid.Text) ' Val(Me.txt_project_id.text) & "AND item_id=" & current_terms
'            Cn.Execute StrSQL, , adExecuteNoRecords
            ' ⁄„·Ì«  «·»‰Êœ
            Set RsDev = New ADODB.Recordset
            RsDev.Open "terms_operations_project_bill", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            Dim i As Integer

            With Me.VSFlexGrid2

                For i = .FixedRows To .Rows - 1

                    '
                    If .TextMatrix(i, .ColIndex("fullcode")) <> "" Then

                        RsDev.AddNew
                        RsDev("bill_id").value = val(Me.txtid.Text)
                        RsDev("fullcode").value = .TextMatrix(i, .ColIndex("fullcode"))
                      '  RsDev("project_id").value = DataCombo2.BoundText
                        RsDev("term_fullcode").value = current_terms
                        RsDev("id").value = .TextMatrix(i, .ColIndex("LineNo"))
        
                        RsDev("name").value = .TextMatrix(i, .ColIndex("name"))
                        RsDev("period").value = IIf(.TextMatrix(i, .ColIndex("period")) = "", 0, .TextMatrix(i, .ColIndex("period")))
                        RsDev("count").value = IIf(.TextMatrix(i, .ColIndex("count")) = "", 0, .TextMatrix(i, .ColIndex("count")))
                        RsDev("salary").value = IIf(.TextMatrix(i, .ColIndex("salary")) = "", 0, .TextMatrix(i, .ColIndex("salary")))
                        RsDev("total_items").value = IIf(.TextMatrix(i, .ColIndex("total_items")) = "", 0, .TextMatrix(i, .ColIndex("total_items")))
                        RsDev("total_salary").value = IIf(.TextMatrix(i, .ColIndex("total_salary")) = "", 0, .TextMatrix(i, .ColIndex("total_salary")))
                        RsDev("total_expenses").value = IIf(.TextMatrix(i, .ColIndex("total_expenses")) = "", 0, .TextMatrix(i, .ColIndex("total_expenses")))
                        RsDev("total").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("total"))), 0, .TextMatrix(i, .ColIndex("total")))
        
                        RsDev("total_items1").value = IIf(.TextMatrix(i, .ColIndex("total_items1")) = "", 0, .TextMatrix(i, .ColIndex("total_items1")))
                        RsDev("total_salary1").value = IIf(.TextMatrix(i, .ColIndex("total_salary1")) = "", 0, .TextMatrix(i, .ColIndex("total_salary1")))
                        RsDev("total_expenses1").value = IIf(.TextMatrix(i, .ColIndex("total_expenses1")) = "", 0, .TextMatrix(i, .ColIndex("total_expenses1")))
                        RsDev("total1").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("total1"))), 0, .TextMatrix(i, .ColIndex("total1")))
        
                        RsDev.update
                    End If

                Next i
    
            End With

            Frame11.Visible = False

    End Select

End Sub



Private Sub Results_Change()
calcnet
End Sub

Private Sub Show_items_Click(Index As Integer)

    Select Case Index

        Case 0

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame10.Visible = True

                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))
                Retrive2 current_opr

                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame10.Caption = "„Ê«œ «·⁄„·Ì… —ﬁ„ :   " & "  " & current_opr
                Else
                    Frame10.Caption = "Items For Operation No:   " & "  " & current_opr
                End If

                XPTxtSum.Text = 0
            End If

        Case 1
            Frame10.Visible = False

    End Select

End Sub

Private Sub Retrive2(current_opr As String)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    fg.Clear flexClearScrollable, flexClearEverything
    fg.Rows = 2
    fg.Enabled = True
 
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where    bill_id is null and (payed =1 )  and opr_fullcode='" & current_opr & "' and Transaction_Date<='" & SQLDate(DTPicker1.value) & "'"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        fg.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            fg.TextMatrix(Num, fg.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            fg.TextMatrix(Num, fg.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            fg.TextMatrix(Num, fg.ColIndex("HaveSerial")) = True
            fg.TextMatrix(Num, fg.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
            fg.TextMatrix(Num, fg.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            fg.TextMatrix(Num, fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            fg.TextMatrix(Num, fg.ColIndex("Valu")) = IIf(IsNull(RsDetails("Quantity")), 0, (RsDetails("Quantity").value)) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value))
            fg.Cell(flexcpData, Num, fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            fg.TextMatrix(Num, fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Sub



Private Sub StartDateProje_Change()
Dim str As String
str = "11/05/2020"
If StartDateProje.value <= CDate(str) Then
TxtFATYou.Text = 5

If val(txtManulaVat) > 0 Then TxtFATYou.Text = txtManulaVat

Else
ClculteVAT
End If
End Sub

Private Sub total_Change()
Calculte
End Sub

Private Sub TxtBillNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtBillNo.Text, 0)
End Sub

Private Sub txtDiscount_Change()
txtDiscountG.Text = txtDiscount.Text
calcnet
ReLineGrid
ReLineGrid
End Sub

Function calcnet()
TxtNetValue.Text = Round(val(Results.Text) - val(txtDiscountG), Decimal_Places) ' - val(advancedPayment.text)
If val(cboDiscount1.ListIndex) = 1 Then
TxtPerforValue.Text = Round(val(txtDiscount1.Text) * val(TxtNetValue.Text) / 100, Decimal_Places)
ElseIf val(cboDiscount1.ListIndex) = 2 Then
TxtPerforValue.Text = Round(val(txtDiscount1.Text), Decimal_Places)
Else
TxtPerforValue.Text = 0
End If
total.Text = Round(val(TxtNetValue.Text) - val(TxtPerforValue.Text), Decimal_Places)
Calculte
End Function
Private Sub TxtDiscount_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, txtDiscount.Text, 0)
End Sub

Private Sub TxtDiscount1_Change()
calcnet
ReLineGrid
ReLineGrid

End Sub

Private Sub txtDiscount2_Change()
ReLineGrid
ReLineGrid

End Sub

Private Sub txtDiscountG_Change()
calcnet
End Sub

Private Sub TxtCost_Change()
ReLineGrid
End Sub

Private Sub TxtID_Change()
    ' "select * from SubcontractorContract2 where bill_id=" & Val(txtid.text)

End Sub
Sub ClculteVAT()
changegridFildssd

If Me.TxtModFlg.Text <> "R" Then
Dim Percetage As Double
Dim account As String
If val(billto.ListIndex) = 0 Then
PercentgValueAddedAccount_Transec XPDtbTrans.value, 6, 1, account, Percetage
ElseIf val(billto.ListIndex) = 1 Then
PercentgValueAddedAccount_Transec XPDtbTrans.value, 7, 0, account, Percetage
End If
TxtFATYou.Text = Percetage
If val(txtManulaVat) > 0 Then TxtFATYou.Text = txtManulaVat

'str = "11/05/2020"
'If StartDateProje.value <= CDate(str) Then
'TxtFATYou.Text = 5
'Else
'TxtFATYou.Text = Percetage
'End If

AccountVat.BoundText = account
Calculte
End If




End Sub
Sub Calculte()
If Me.TxtModFlg.Text <> "R" Then
If val(txtManulaVat) > 0 Then TxtFATYou.Text = txtManulaVat
If val(TxtFATYou.Text) > 0 Then
TxtFATValue.Text = Round((val(TxtNetValue.Text) * val(TxtFATYou.Text)) / 100, Decimal_Places)
Else
TxtFATValue.Text = 0
End If
TxtTotalValue.Text = Round(val(total.Text) + val(TxtFATValue.Text), Decimal_Places)
End If
End Sub

Public Sub Retrive(Optional Lngid As Long, Optional note_id As Double = 0)
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
            XPTxtCurrent.Caption = 0
            XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If


If note_id <> 0 Then

            rs.Find "note_id=" & note_id, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
  
        GoTo ll
End If

        If Lngid <> 0 Then
            rs.Find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
      
    End If
ll:
TxtNetValue.Text = Round(IIf(IsNull(rs("NetValue").value), IIf(IsNull(rs("total").value), 0, rs("total").value), rs("NetValue").value), Decimal_Places)
TxtPerforValue.Text = Round(IIf(IsNull(rs("PerforValue").value), 0, rs("PerforValue").value), Decimal_Places)

StartDateProje.value = IIf(IsNull(rs("StartDateProje").value), Date, rs("StartDateProje").value)
Me.dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    txtid.Text = IIf(IsNull(rs("id").value), 0, (rs("id").value))
    TxtPreVAT.Text = IIf(IsNull(rs("PreVAT").value), 0, rs("PreVAT").value)
'Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
Me.TxtNoteSerial1.Text = IIf(IsNull(rs("Id").value), "", rs("Id").value)
    XPDtbTrans.value = IIf(IsNull(rs("bill_date").value), Date, rs("bill_date").value)
dueDate.value = IIf(IsNull(rs("dueDate").value), Date, rs("dueDate").value)
'dueDate1.value = IIf(IsNull(rs("dueDate1").value), Date, rs("dueDate1").value)
'TxtAccountUnderImp.Text = IIf(IsNull(rs("AccountUnderImp").value), "", rs("AccountUnderImp").value)
TxtFATYou.Text = IIf(IsNull(rs("FATYou").value), 0, (rs("FATYou").value))
txtManulaVat.Text = IIf(IsNull(rs("FATYou").value), 0, (rs("FATYou").value))


TxtFATValue.Text = IIf(IsNull(rs("FATValue").value), 0, (rs("FATValue").value))
TxtTotalValue.Text = Round(IIf(IsNull(rs("TotalValue").value), 0, (rs("TotalValue").value)), Decimal_Places)
Me.AccountVat.BoundText = IIf(IsNull(rs("AccountCodeVat").value), "", (rs("AccountCodeVat").value))
DataCombo2.BoundText = IIf(IsNull(rs("project_no").value), "", rs("project_no").value)
'*************************************************
DcbosubContractor.BoundText = IIf(IsNull(rs("subContractorId").value), "", rs("subContractorId").value)
'txtDiscount1.Text = IIf(IsNull(rs("discount1value").value), 0, (rs("discount1value").value))
'txtDiscount2.Text = IIf(IsNull(rs("discount2value").value), 0, (rs("discount2value").value))

'cboDiscount1.ListIndex = IIf(IsNull(rs("discount1ID").value), 0, (rs("discount1ID").value))
'cboDiscount2.ListIndex = IIf(IsNull(rs("discount2ID").value), 0, (rs("discount2ID").value))
DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
TxtCost.Text = IIf(IsNull(rs("ExPercen").value), 0, (rs("ExPercen").value))
DcbExPercen.ListIndex = IIf(IsNull(rs("ExPercenID").value), -1, (rs("ExPercenID").value))
'*************************************************
'   If Not IsNull(rs("UnderImp").value) Then
'   If rs("UnderImp").value = 0 Then
'   Option7.value = True
'   ElseIf rs("UnderImp").value = 1 Then
'   Option6.value = True
'   ElseIf rs("UnderImp").value = 2 Then
'   Option8.value = True
'   End If
'   Else
'   Option7.value = True
'   End If
'26082015
'TxtDiscount.Text = Round(IIf(IsNull(rs("Discount").value), 0, (rs("Discount").value)), Decimal_Places)
Results.Text = Round(IIf(IsNull(rs("Results").value), 0, (rs("Results").value)), Decimal_Places)
'advancedPayment.Text = Round(IIf(IsNull(rs("advancedPayment").value), 0, (rs("advancedPayment").value)), Decimal_Places)
 ''//////////23 05 2016
  TxtBillNo.Text = IIf(IsNull(rs("BillNo").value), 0, rs("BillNo").value)
  
  Me.DcbPeriodType.ListIndex = IIf(IsNull(rs("PeriodType").value), -1, rs("PeriodType").value)
  'Me.TxtRemarks2.Text = IIf(IsNull(rs("Remarks2").value), "", rs("Remarks2").value)
  'startDate.value = IIf(IsNull(rs("StartDate").value), Date, rs("StartDate").value)
 ''//
 TxtPreBalaValue.Text = IIf(IsNull(rs("PreBalaValue").value), 0, rs("PreBalaValue").value)
 TxtPreBalaVAT.Text = IIf(IsNull(rs("PreBalaVAT").value), 0, rs("PreBalaVAT").value)
 TxtPreBalaTotal.Text = Round(IIf(IsNull(rs("PreBalaTotal").value), 0, rs("PreBalaTotal").value), Decimal_Places)
 TxtPreBalaPayed.Text = IIf(IsNull(rs("PreBalaPayed").value), 0, rs("PreBalaPayed").value)
 TxtPreBalaRemain.Text = IIf(IsNull(rs("PreBalaRemain").value), 0, rs("PreBalaRemain").value)
 TxtPreBalaTransPyed.Text = IIf(IsNull(rs("PreBalaTransPyed").value), 0, rs("PreBalaTransPyed").value)
 TxtPreBalaNet.Text = Round(IIf(IsNull(rs("PreBalaNet").value), 0, rs("PreBalaNet").value), Decimal_Places)
 TxtPreBalaVATYu.Text = IIf(IsNull(rs("PreBalaVATYu").value), 0, rs("PreBalaVATYu").value)
 Label57.Caption = IIf(IsNull(rs("SumVATLine").value), 0, rs("SumVATLine").value)
 Label56.Caption = IIf(IsNull(rs("SumValueLine").value), 0, rs("SumValueLine").value)
 
'26082015
    txtprojectname.Text = IIf(IsNull(rs("project_name").value), "", rs("project_name").value)
'    DcAccount1.text = IIf(IsNull(rs("Sub_user_name").value), "", rs("Sub_user_name").value)
'    DcAccount2.text = IIf(IsNull(rs("End_user_name").value), "", rs("End_user_name").value)

    txtendaccount.Text = IIf(IsNull(rs("End_user_account").value), "", rs("End_user_account").value)
    'txtsubaccount.Text = IIf(IsNull(rs("Sub_user_account").value), "", rs("Sub_user_account").value)
    txtrevenue_account.Text = IIf(IsNull(rs("revenue_account").value), "", rs("revenue_account").value)

    'DcAccount4.text = IIf(IsNull(Rs("sub_contractor_name").value), "", Rs("sub_contractor_name").value)

    'DcAccount2.text = IIf(IsNull(Rs("End_user_name").value), "", Rs("End_user_name").value)

    billto.ListIndex = IIf(IsNull(rs("bill_to").value), -1, rs("bill_to").value)
  If IsNull(rs("bill_type").value) Then
    bill_Type.ListIndex = 0
    Else
    bill_Type.ListIndex = IIf(IsNull(rs("bill_type").value), 0, val(rs("bill_type").value))
    End If
    Me.note_id.Text = IIf(IsNull(rs("note_id").value), "", rs("note_id").value)
    TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    'TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
     '   txtManualNo.Text = IIf(IsNull(rs("ManualNo").value), "", rs("ManualNo").value)
        

'rs("Remarks").value = Trim(TxtRemarks.text)
'rs("ManualNo").value = Trim(txtManualNo.text)

    total.Text = Round(IIf(IsNull(rs("total").value), 0, rs("total").value), Decimal_Places)

    'Exit Sub

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee, dbo.ProjectMainDes.FullCode, dbo.ProjectMainDes.Name,"
        StrSQL = StrSQL + "              dbo.SubcontractorContract2.*"
        StrSQL = StrSQL + "    FROM         dbo.SubcontractorContract2 LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.ProjectMainDes ON dbo.SubcontractorContract2.PrMainDesID = dbo.ProjectMainDes.ID LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.TblProcessUnites ON dbo.SubcontractorContract2.Unit_id = dbo.TblProcessUnites.UnitID"
        StrSQL = StrSQL + " Where dbo.SubcontractorContract2.bill_id =" & Me.txtid.Text
       StrSQL = StrSQL + " order by SubcontractorContract2.id"
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            'Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            'Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
    
            RsDev.MoveFirst
    
            With Me.Fg_Journal
                .Rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("ExPercen")) = IIf(IsNull(RsDev("ExPercen").value), 0, RsDev("ExPercen").value)
                .TextMatrix(i, .ColIndex("PrMainDesID")) = IIf(IsNull(RsDev("PrMainDesID").value), 0, RsDev("PrMainDesID").value)
                .TextMatrix(i, .ColIndex("CodeMain")) = IIf(IsNull(RsDev("FullCode").value), "", RsDev("FullCode").value)
                .TextMatrix(i, .ColIndex("MainDes")) = IIf(IsNull(RsDev("Name").value), "", RsDev("Name").value)
                .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(RsDev("qty").value), 0, RsDev("qty").value)
                .TextMatrix(i, .ColIndex("oprid")) = IIf(IsNull(RsDev("oprid").value), 0, RsDev("oprid").value)
                .TextMatrix(i, .ColIndex("totEx")) = IIf(IsNull(RsDev("totEx").value), 0, RsDev("totEx").value)
                .TextMatrix(i, .ColIndex("quntExc")) = IIf(IsNull(RsDev("quntExc").value), 0, RsDev("quntExc").value)
                .TextMatrix(i, .ColIndex("net")) = IIf(IsNull(RsDev("net").value), 0, RsDev("net").value)
                .TextMatrix(i, .ColIndex("discount")) = IIf(IsNull(RsDev("discount").value), 0, RsDev("discount").value)
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), 0, RsDev("total").value)
                .TextMatrix(i, .ColIndex("Period")) = IIf(IsNull(RsDev("Period").value), 0, RsDev("Period").value)
                                
                'newwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
                .TextMatrix(i, .ColIndex("QtyApprov")) = IIf(IsNull(RsDev("QtyApprov").value), 0, RsDev("QtyApprov").value)
                .TextMatrix(i, .ColIndex("TotalApprov")) = IIf(IsNull(RsDev("TotalApprov").value), 0, RsDev("TotalApprov").value)
                .TextMatrix(i, .ColIndex("PriceApprov")) = IIf(IsNull(RsDev("PriceApprov").value), 0, RsDev("PriceApprov").value)
                .TextMatrix(i, .ColIndex("DiscApprov")) = IIf(IsNull(RsDev("DiscApprov").value), 0, RsDev("DiscApprov").value)
                .TextMatrix(i, .ColIndex("NetApprov")) = IIf(IsNull(RsDev("NetApprov").value), 0, RsDev("NetApprov").value)
                '''////
                 .TextMatrix(i, .ColIndex("discountEXE")) = IIf(IsNull(RsDev("discountEXE").value), 0, RsDev("discountEXE").value)
                  .TextMatrix(i, .ColIndex("NetExe")) = IIf(IsNull(RsDev("NetExe").value), 0, RsDev("NetExe").value)
                  
                  .TextMatrix(i, .ColIndex("project_id")) = IIf(IsNull(RsDev("project_id").value), "", RsDev("project_id").value)
                  .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(RsDev("FullCode").value), "", RsDev("FullCode").value)
                  .TextMatrix(i, .ColIndex("projectName")) = IIf(IsNull(RsDev("projectName").value), "", RsDev("projectName").value)

                 
                
                  .TextMatrix(i, .ColIndex("Percentage1")) = IIf(IsNull(RsDev("Percentage1").value), 0, RsDev("Percentage1").value)
                  .TextMatrix(i, .ColIndex("Pre_Percent1")) = IIf(IsNull(RsDev("Pre_Percent1").value), 0, RsDev("Pre_Percent1").value)
                  .TextMatrix(i, .ColIndex("tot_percent1")) = IIf(IsNull(RsDev("tot_percent1").value), 0, RsDev("tot_percent1").value)
                  
               'newwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
               
                
                
                     .TextMatrix(i, .ColIndex("unit_id")) = IIf(IsNull(RsDev("unit_id").value), "", RsDev("unit_id").value)
                    .TextMatrix(i, .ColIndex("item")) = IIf(IsNull(RsDev("item").value), "", RsDev("item").value)
    
                    .TextMatrix(i, .ColIndex("item_id")) = IIf(IsNull(RsDev("item_id").value), "", RsDev("item_id").value)
            
                    .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("cost").value), "", RsDev("cost").value)
                    
                    .TextMatrix(i, .ColIndex("qtySubContractor")) = IIf(IsNull(RsDev("qtySubContractor").value), "", RsDev("qtySubContractor").value)
                    .TextMatrix(i, .ColIndex("costSubContractor")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    
                    
                                        If val(.TextMatrix(i, .ColIndex("qtySubContractor"))) = 0 Then
                    .TextMatrix(i, .ColIndex("qtySubContractor")) = .TextMatrix(i, .ColIndex("qty"))
                    End If
                    
                    
                   If val(.TextMatrix(i, .ColIndex("costSubContractor"))) = 0 Then
                    .TextMatrix(i, .ColIndex("costSubContractor")) = .TextMatrix(i, .ColIndex("exe"))
                    End If
                                
                                
                                
                 .TextMatrix(i, .ColIndex("OLDTotalwithVat")) = IIf(IsNull(RsDev("OLDTotalwithVat").value), 0, RsDev("OLDTotalwithVat").value)
                 .TextMatrix(i, .ColIndex("CurrenttotalWithvat")) = IIf(IsNull(RsDev("CurrenttotalWithvat").value), 0, RsDev("CurrenttotalWithvat").value)
                  .TextMatrix(i, .ColIndex("Totalwitvat")) = IIf(IsNull(RsDev("Totalwitvat").value), 0, RsDev("Totalwitvat").value)
                                  
                               lbl(9).Caption = IIf(IsNull(RsDev("oldPerforValue").value), 0, RsDev("oldPerforValue").value)
                              lbl(11).Caption = IIf(IsNull(RsDev("totalPerforValue").value), 0, RsDev("totalPerforValue").value)
                                  
            
            
               
                    
                    .TextMatrix(i, .ColIndex("exe")) = IIf(IsNull(RsDev("exe").value), "", RsDev("exe").value)
           
                    .TextMatrix(i, .ColIndex("percentage")) = IIf(IsNull(RsDev("percentage").value), "", RsDev("percentage").value)
        
                    .TextMatrix(i, .ColIndex("exedate")) = IIf(IsNull(RsDev("exedate").value), "", RsDev("exedate").value)
                    
                    If SystemOptions.UserInterface = ArabicInterface Then
                          .TextMatrix(i, .ColIndex("Unit")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
                     Else
                     .TextMatrix(i, .ColIndex("Unit")) = IIf(IsNull(RsDev("UnitNamee").value), "", RsDev("UnitNamee").value)
                     End If
                           .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(RsDev("Quantity").value), "", RsDev("Quantity").value)
                            .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), "", RsDev("Price").value)
                             .TextMatrix(i, .ColIndex("Pre_Quantity")) = IIf(IsNull(RsDev("Pre_Quantity").value), "", RsDev("Pre_Quantity").value)
                              .TextMatrix(i, .ColIndex("Pre_Value")) = IIf(IsNull(RsDev("Pre_Value").value), "", RsDev("Pre_Value").value)
                              .TextMatrix(i, .ColIndex("Pre_Percent")) = IIf(IsNull(RsDev("Pre_Percent").value), "", RsDev("Pre_Percent").value)
                              
                          
                            .TextMatrix(i, .ColIndex("Curr_Quantity")) = IIf(IsNull(RsDev("Curr_Quantity").value), "", RsDev("Curr_Quantity").value)
                            .TextMatrix(i, .ColIndex("Curr_value")) = IIf(IsNull(RsDev("Curr_value").value), "", RsDev("Curr_value").value)
                            .TextMatrix(i, .ColIndex("curr_Percent")) = IIf(IsNull(RsDev("curr_Percent").value), "", RsDev("curr_Percent").value)
                 .TextMatrix(i, .ColIndex("tot_quantity")) = IIf(IsNull(RsDev("tot_quantity").value), "", RsDev("tot_quantity").value)
          
                 .TextMatrix(i, .ColIndex("tot_value")) = IIf(IsNull(RsDev("tot_value").value), "", RsDev("tot_value").value)
                 .TextMatrix(i, .ColIndex("tot_percent")) = IIf(IsNull(RsDev("tot_percent").value), "", RsDev("tot_percent").value)
                
                    
        
                    RsDev.MoveNext
                Next i

                'Me.txt_total_sum.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
                '  Me.txt_sub_discount.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("discount"), .Rows - 1, .ColIndex("discount"))
                '    Me.txt_sub_net.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("net"), .Rows - 1, .ColIndex("net"))
           
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
                '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
                '  .Rows - 1, .ColIndex("CreditValue"))
                '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
                '  .Rows - 1, .ColIndex("DebitValue"))
            End With

        End If

    End If

    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
'GET_PROJECT_DATA
RetriveBillBuyData
    ReLineGrid
    ReLineGrid
    
    fillapprovData
    changegridFildssd
     Exit Sub
ErrTrap:

End Sub
Function changegridFildssd()
  With Fg_Journal
    If billto.ListIndex = 0 Then
    If Me.TxtModFlg <> "R" Then
   ' DcbosubContractor.BoundText = 0
    End If
'    DcbosubContractor.Visible = True
'    Text2.Visible = False
'    Label22.Visible = False
    
'     .ColHidden(.ColIndex("qty")) = False
'       .ColHidden(.ColIndex("cost")) = False
'       .ColHidden(.ColIndex("total")) = False
'       .ColHidden(.ColIndex("discount")) = False
'       .ColHidden(.ColIndex("net")) = False
''             .ColHidden(.ColIndex("qtySubContractor")) = True
''       .ColHidden(.ColIndex("costSubContractor")) = True
     Else '„ﬁ«Ê·
    
    DcbosubContractor.Visible = True
    Text2.Visible = True
    Label22.Visible = True
    
'      .ColHidden(.ColIndex("qty")) = True
'         .ColHidden(.ColIndex("cost")) = True
'       .ColHidden(.ColIndex("total")) = True
'       .ColHidden(.ColIndex("discount")) = True
'       .ColHidden(.ColIndex("net")) = True
'      .ColHidden(.ColIndex("qtySubContractor")) = False
'       .ColHidden(.ColIndex("costSubContractor")) = False
 
     
     End If
     
     
      
      End With
 
End Function
Sub RelineBuy()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid4
        For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("RemainingValue")))
           End If
           Next i
  
    End With
   Label47.Caption = Sm
End Sub
Private Sub ReLineGrid(Optional current_terms As String = "")
    
    On Error Resume Next
    Dim i As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim IntCounter As Integer
    Dim XRound As Integer
    Dim LineDiscountPercent As Double
    Dim LineDiscount As Double
    Dim linenetaftermainDiscount  As Double
changegridFildssd

        If SystemOptions.AllowNoRoudProjectInvoices = True Then
        XRound = val(cCompanyInfo.NoRoudProjectInvoices)
        Else
        XRound = 2
        End If
    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("item")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
              '  .TextMatrix(i, .ColIndex("cost")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("cost"))), 1, .TextMatrix(i, .ColIndex("cost")))
           
                ' sql = "  From terms_operations Where term_fullcode='" & .TextMatrix(I, .ColIndex("fullcode")) & "'"
              '  sql = "select sum(total1) as total  from terms_operations_project_bill where term_fullcode='" & .TextMatrix(i, .ColIndex("item_id")) & "' and bill_id=" & val(Me.txtid.text)
         
              '  Set rs = New ADODB.Recordset
              '  rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

              '  If rs.RecordCount > 0 And Not IsNull(rs("total").value) Then
                  .TextMatrix(i, .ColIndex("totEx")) = val(.TextMatrix(i, .ColIndex("exe"))) * val(.TextMatrix(i, .ColIndex("quntExc")))
         
              '  Else
                '    .TextMatrix(i, .ColIndex("exe")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("exe"))), 0, .TextMatrix(i, .ColIndex("exe")))
              '  End If
        
                '.TextMatrix(i, .ColIndex("percentage")) = Round(.TextMatrix(i, .ColIndex("exe")) / .TextMatrix(i, .ColIndex("cost")) * 100, 2)
             
                .TextMatrix(i, .ColIndex("exedate")) = IIf(.TextMatrix(i, .ColIndex("exedate")) = "", Date, .TextMatrix(i, .ColIndex("exedate")))
                
             ''///////////
                 .TextMatrix(i, .ColIndex("TotalApprov")) = (val(.TextMatrix(i, .ColIndex("PriceApprov"))) * val(.TextMatrix(i, .ColIndex("QtyApprov"))))
                .TextMatrix(i, .ColIndex("NetApprov")) = val(.TextMatrix(i, .ColIndex("TotalApprov"))) - val(.TextMatrix(i, .ColIndex("DiscApprov")))

  
       
            .TextMatrix(i, .ColIndex("totEx")) = (val(.TextMatrix(i, .ColIndex("exe"))) * val(.TextMatrix(i, .ColIndex("quntExc"))))

.TextMatrix(i, .ColIndex("NetExe")) = val(.TextMatrix(i, .ColIndex("totEx"))) - val(.TextMatrix(i, .ColIndex("discountEXE")))

 
       Dim netexe As Double
       Dim QtyExe As Double
    Dim VATPer  As Double
    Dim oldPerforValue  As Double
    Dim discountHasmyat As Double
    Dim linenetaftermainDiscountWithvat As Double
    linenetaftermainDiscountWithvat = 0
  '    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
       GetTermsTotals val(.TextMatrix(i, .ColIndex("oprid"))), val(txtid.Text), XPDtbTrans.value, netexe, QtyExe, TxtModFlg.Text, billto.ListIndex, val(DcbosubContractor.BoundText), VATPer, oldPerforValue, discountHasmyat, linenetaftermainDiscountWithvat
  '    End If
  
  
  '************************** cancelled***********************************************
  .TextMatrix(i, .ColIndex("Pre_Quantity")) = QtyExe
      .TextMatrix(i, .ColIndex("Pre_Value")) = netexe
      
      
      
     
       
   .TextMatrix(i, .ColIndex("OLDTotalwithVat")) = Round(linenetaftermainDiscountWithvat, 2)
   
   
     LineDiscountPercent = val(.TextMatrix(i, .ColIndex("NetExe"))) / Results.Text
              LineDiscount = (val(txtDiscountG.Text)) * LineDiscountPercent
         
              
               
                 linenetaftermainDiscount = val(.TextMatrix(i, .ColIndex("NetExe"))) - LineDiscount
                 
                
 
   .TextMatrix(i, .ColIndex("CurrenttotalWithvat")) = Round(((linenetaftermainDiscount - val(.TextMatrix(i, .ColIndex("discountEXE"))))) * (1 + TxtFATYou / 100), 2)
 .TextMatrix(i, .ColIndex("Totalwitvat")) = Round(val(.TextMatrix(i, .ColIndex("CurrenttotalWithvat"))) + .TextMatrix(i, .ColIndex("OLDTotalwithVat")), 2)
          
          lbl(9).Caption = Round(oldPerforValue, 2)
          lbl(10).Caption = Round(val(TxtPerforValue.Text), 2)
          lbl(11).Caption = val(lbl(9).Caption) + val(lbl(10).Caption)
          
      
       
      If billto.ListIndex = 0 Then '⁄„Ì·
      
          .TextMatrix(i, .ColIndex("percentage")) = ((val(.TextMatrix(i, .ColIndex("quntExc"))) / val(.TextMatrix(i, .ColIndex("qty")))))
    .TextMatrix(i, .ColIndex("percentage1")) = ((val(.TextMatrix(i, .ColIndex("NetExe"))) / val(.TextMatrix(i, .ColIndex("net")))))

                    If val(.TextMatrix(i, .ColIndex("Pre_Quantity"))) > 0 Then
                    .TextMatrix(i, .ColIndex("Pre_Percent")) = (val(.TextMatrix(i, .ColIndex("Pre_Quantity"))) / val(.TextMatrix(i, .ColIndex("Qty"))))
                    Else
                    .TextMatrix(i, .ColIndex("Pre_Percent")) = 0
                    End If
                    
                        If val(.TextMatrix(i, .ColIndex("Pre_Value"))) > 0 Then
                    .TextMatrix(i, .ColIndex("Pre_Percent1")) = (val(.TextMatrix(i, .ColIndex("Pre_Value"))) / val(.TextMatrix(i, .ColIndex("net"))))
                    Else
                    .TextMatrix(i, .ColIndex("Pre_Percent1")) = 0
                    End If
                    
                             .TextMatrix(i, .ColIndex("tot_quantity")) = (val(.TextMatrix(i, .ColIndex("Pre_Quantity"))) + val(.TextMatrix(i, .ColIndex("quntExc"))))
          .TextMatrix(i, .ColIndex("tot_percent")) = (val(.TextMatrix(i, .ColIndex("tot_quantity"))) / val(.TextMatrix(i, .ColIndex("Qty"))))
          
          .TextMatrix(i, .ColIndex("tot_value")) = (.TextMatrix(i, .ColIndex("Pre_Value")) + val(.TextMatrix(i, .ColIndex("netexe"))))
          
          .TextMatrix(i, .ColIndex("tot_percent1")) = (val(.TextMatrix(i, .ColIndex("tot_value"))) / val(.TextMatrix(i, .ColIndex("net"))))

     
     Else '„ﬁ«Ê·
     
                                        If val(.TextMatrix(i, .ColIndex("qtySubContractor"))) = 0 Then
                    .TextMatrix(i, .ColIndex("qtySubContractor")) = .TextMatrix(i, .ColIndex("qty"))
                    End If
                    
                    
                   If val(.TextMatrix(i, .ColIndex("costSubContractor"))) = 0 Then
                    .TextMatrix(i, .ColIndex("costSubContractor")) = .TextMatrix(i, .ColIndex("exe"))
                    End If
                                
                                
          .TextMatrix(i, .ColIndex("percentage")) = ((val(.TextMatrix(i, .ColIndex("quntExc"))) / val(.TextMatrix(i, .ColIndex("qtySubContractor")))))
    .TextMatrix(i, .ColIndex("percentage1")) = ((val(.TextMatrix(i, .ColIndex("NetExe"))) / val(.TextMatrix(i, .ColIndex("qtySubContractor")) * .TextMatrix(i, .ColIndex("costSubContractor")))))
   
   
         If val(.TextMatrix(i, .ColIndex("Pre_Value"))) > 0 Then
                    .TextMatrix(i, .ColIndex("Pre_Percent1")) = (val(.TextMatrix(i, .ColIndex("Pre_Value"))) / val(.TextMatrix(i, .ColIndex("qtySubContractor")) * .TextMatrix(i, .ColIndex("costSubContractor"))))
                    Else
                    .TextMatrix(i, .ColIndex("Pre_Percent1")) = 0
                    End If
                    
             If val(.TextMatrix(i, .ColIndex("Pre_Quantity"))) > 0 Then
                    .TextMatrix(i, .ColIndex("Pre_Percent")) = (val(.TextMatrix(i, .ColIndex("Pre_Quantity"))) / val(.TextMatrix(i, .ColIndex("qtySubContractor"))))
                    Else
                    .TextMatrix(i, .ColIndex("Pre_Percent")) = 0
                    End If
                    
                    
                    
            .TextMatrix(i, .ColIndex("tot_quantity")) = (val(.TextMatrix(i, .ColIndex("Pre_Quantity"))) + val(.TextMatrix(i, .ColIndex("quntExc"))))
          .TextMatrix(i, .ColIndex("tot_percent")) = (val(.TextMatrix(i, .ColIndex("tot_quantity"))) / val(.TextMatrix(i, .ColIndex("qtySubContractor"))))
          
          .TextMatrix(i, .ColIndex("tot_value")) = (.TextMatrix(i, .ColIndex("Pre_Value")) + val(.TextMatrix(i, .ColIndex("netexe"))))
          
          .TextMatrix(i, .ColIndex("tot_percent1")) = (val(.TextMatrix(i, .ColIndex("tot_value"))) / val(.TextMatrix(i, .ColIndex("qtySubContractor")) * .TextMatrix(i, .ColIndex("costSubContractor"))))
          
     
     End If
     
                    
                
                    
'************************** cancelled***********************************************
 
        
            End If
CalCultePers i
        Next i



Me.Results.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("totEx"), .Rows - 1, .ColIndex("totEx")), Decimal_Places)
      '  Me.total.Text = val(Me.Results.Text) - val(txtDiscountG.Text)  ' .Aggregate(flexSTSum, .FixedRows, .ColIndex("totEx"), .Rows - 1, .ColIndex("totEx"))
      calcnet
         
    End With

    IntCounter = 0

    With VSFlexGrid2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
          
                .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("total_expenses1"))) + val(.TextMatrix(i, .ColIndex("total_salary1"))) + val(.TextMatrix(i, .ColIndex("total_items1")))
           
            End If

        Next i

        Me.txt_opr_total.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total")), Decimal_Places)
    End With

    IntCounter = 0

    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
             
            End If

        Next i
 
    End With

End Sub

Private Sub txtManulaVat_Change()
Calculte
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„‘—Ê⁄«  "
            Else
                Me.Caption = "Projects"
            End If
        
            Me.Command1(0).Enabled = True 'ÃœÌœ
            Me.Command1(3).Enabled = True ' ⁄œÌ·
            Me.Command1(1).Enabled = False 'Õ›Ÿ
            Me.Command1(9).Enabled = True 'Õ–›
            Me.Command1(6).Enabled = False ' —«Ã⁄
            Me.Command1(10).Enabled = True '»ÕÀ
         
            Me.Command1(7).Enabled = True 'ÿ»«⁄Â ›« Ê—…
            Me.Command1(8).Enabled = True 'ÿ»«⁄Â  ﬁÌœ
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
 
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Command1(9).Enabled = False
                Me.Command1(3).Enabled = False
            
            End If
        
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„‘—Ê⁄«  (ÃœÌœ)"
            Else
                Me.Caption = " Projects(New Record)"
            End If
        
            Frame12.Enabled = True
            Frame1.Enabled = True
            Frame10.Enabled = True
            Frame11.Enabled = True
            Frame2.Enabled = True
            Frame13.Enabled = True
        
            Me.Command1(0).Enabled = False 'ÃœÌœ
            Me.Command1(3).Enabled = False ' ⁄œÌ·
            Me.Command1(1).Enabled = True 'Õ›Ÿ
            Me.Command1(9).Enabled = False 'Õ–›
            Me.Command1(6).Enabled = True ' —«Ã⁄
            Me.Command1(10).Enabled = False '»ÕÀ
         
            Me.Command1(7).Enabled = False 'ÿ»«⁄Â ›« Ê—…
            Me.Command1(8).Enabled = False 'ÿ»«⁄Â ﬁÌœ
         
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„‘—Ê⁄« (  ⁄œÌ· )"
            Else
                Me.Caption = "Projects (Edit Current Record)"
            End If
            FlgBillBuy = False
            Frame12.Enabled = True
            Frame1.Enabled = True
            Frame10.Enabled = True
            Frame11.Enabled = True
            Frame2.Enabled = True
            Frame13.Enabled = True
             
            Me.Command1(0).Enabled = False 'ÃœÌœ
            Me.Command1(3).Enabled = False ' ⁄œÌ·
            Me.Command1(1).Enabled = True 'Õ›Ÿ
            Me.Command1(9).Enabled = False 'Õ–›
            Me.Command1(6).Enabled = True ' —«Ã⁄
            Me.Command1(10).Enabled = False '»ÕÀ
         
            Me.Command1(7).Enabled = False 'ÿ»«⁄Â ›« Ê—…
            Me.Command1(8).Enabled = False 'ÿ»«⁄Â  ﬁÌœ
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
  
    End Select

    Exit Sub
ErrTrap:

End Sub



Private Sub TxtPreBalaTransPyed_Change()
ClculteVATBalan
SumVAT
End Sub

Private Sub TxtPreBalaTransPyed_LostFocus()
If val(TxtPreBalaTransPyed.Text) > val(TxtPreBalaRemain.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Paid value is greater than remaining"
End If
TxtPreBalaTransPyed.Text = 0
Exit Sub
End If
ClculteVATBalan
SumVAT
End Sub
Sub ClculteVATBalan()
Dim Percetage2 As Double
TxtPreBalaNet.Text = val(TxtPreBalaRemain.Text) - val(TxtPreBalaTransPyed.Text)

End Sub
Function GetBalanceProject(Optional ByRef Valu As Double, Optional ByRef RecDate As Date) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     OpenBalance8, OpenBalanceDate"
sql = sql & " From dbo.Projects"
sql = sql & " Where (ID = " & val(DataCombo2.BoundText) & ") "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
Valu = IIf(IsNull(rs2("OpenBalance8").value), 0, rs2("OpenBalance8").value)
RecDate = IIf(IsNull(rs2("OpenBalanceDate").value), getFirstPeriodDateInthisYear2, rs2("OpenBalanceDate").value)
GetBalanceProject = True
Else
GetBalanceProject = False
End If
End Function

Function GetValue() As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(PreBalaTransPyed) AS SumValue"
sql = sql & " From dbo.SubcontractorContract"
sql = sql & " WHERE     (id <> " & val(txtid.Text) & ") AND (project_no = N'" & DataCombo2.BoundText & "')"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetValue = IIf(IsNull(rs2("SumValue").value), 0, rs2("SumValue").value)
Else
GetValue = 0
End If
End Function

Private Sub txtprojectname_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
             FrmProjectSearch.lblSearchtype.Caption = 8
             FrmProjectSearch.show vbModal
        End If
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid2

        Select Case .ColKey(Col)
 
            Case "name"
                StrAccountCode = .ComboItem
       
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("name"), False, True)
                .TextMatrix(Row, .ColIndex("name")) = StrAccountCode
            
                If StrAccountCode <> "" Then
                    StrSQL = "SELECT   * from dbo.terms_operations WHERE  fullcode ='" & .ComboData & "'"
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             
                    .TextMatrix(Row, .ColIndex("fullcode")) = IIf(IsNull(Rs1("fullcode").value), 0, Rs1("fullcode").value)
                    .TextMatrix(Row, .ColIndex("total_items")) = IIf(IsNull(Rs1("total_items").value), 0, Rs1("total_items").value)
            
                    .TextMatrix(Row, .ColIndex("total_salary")) = IIf(IsNull(Rs1("total_salary").value), 0, Rs1("total_salary").value)
                    .TextMatrix(Row, .ColIndex("total_expenses")) = IIf(IsNull(Rs1("total_expenses").value), 0, Rs1("total_expenses").value)
                    .TextMatrix(Row, .ColIndex("total")) = IIf(IsNull(Rs1("total").value), 0, Rs1("total").value)
                    .TextMatrix(Row, .ColIndex("total_items1")) = get_opr_material_total(.ComboData, DTPicker1.value)
                    .TextMatrix(Row, .ColIndex("total_expenses1")) = get_opr_expenses_total(.ComboData, DTPicker1.value)
             
                Else
 
                    .TextMatrix(Row, .ColIndex("fullcode")) = ""
             
                End If

        End Select

        '  Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
    
    End With

    ReLineGrid
End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid2
        .ComboList = ""

        Select Case .ColKey(Col)
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid2_Click()
    current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))

    With VSFlexGrid2
   
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
 
    End With

    ReLineGrid

End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With VSFlexGrid2

        Select Case .ColKey(Col)

            Case "name"
       
                'Full Path Display
                StrSQL = "SELECT   fullcode,name from dbo.terms_operations WHERE term_fullcode ='" & current_terms & "'"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "name", "fullcode")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                    
        End Select

    End With

End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "id='" & val(txtid.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub VSFlexGrid4_AfterEdit(ByVal Row As Long, ByVal Col As Long)
RelineBuy
RelineBu22
End Sub

Private Sub VSFlexGrid4_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)




  
'ElseIf VSFlexGrid4.ColIndex("TransPayedValue") <> Col Then
'  VSFlexGrid4.ComboList = ""
'End If
End Sub

Private Sub VSFlexGrid4_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Discount1 As Double
If VSFlexGrid4.ColIndex("payed") = Col Then
 
                            If cboDiscount2.ListIndex = 0 Then
                                Discount2 = 0
                            ElseIf cboDiscount2.ListIndex = 1 Then
                                Discount2 = val(txtDiscount2) * val(VSFlexGrid4.TextMatrix(Row, VSFlexGrid4.ColIndex("Note_Value"))) / 100
                            ElseIf cboDiscount2.ListIndex = 2 Then
                                Discount2 = val(txtDiscount2)
                            End If
                 
        
   VSFlexGrid4.TextMatrix(Row, VSFlexGrid4.ColIndex("TransPayedValue")) = Discount2
Exit Sub
End If
If VSFlexGrid4.ColIndex("payed") = Col Or VSFlexGrid4.ColIndex("TransPayedValue") = Col Then

                     If VSFlexGrid4.ColIndex("TransPayedValue") = Col And VSFlexGrid4.Cell(flexcpChecked, Row, 4) <> 2 Then
                    ' MsgBox VSFlexGrid4.Cell(flexcpChecked, Row, 4)
                 Else
                  Cancel = True
                    Exit Sub
                End If
 Else
                   Cancel = True
                   
  End If

   
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub
Private Sub XPDtbTrans_Change()
    TxtNoteSerial.Text = ""
   If Me.TxtModFlg.Text <> "R" Then
      If ChekSanNumber(Current_branch, 65) = True Then
          TxtNoteSerial1.Text = ""
      End If
      TxtNoteSerial.Text = ""
   End If
End Sub
' aladein cod
Private Sub txtManualNo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    DataCombo2.SetFocus
End If
 Exit Sub
ErrTrap:
End Sub
Private Sub DataCombo2_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = vbKeyReturn Then
GET_PROJECT_DATA 1
End If
If KeyAscii = 13 Then
    DcAccount2.SetFocus
End If
 Exit Sub
ErrTrap:
End Sub
Private Sub DcAccount2_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    billto.SetFocus
End If
 Exit Sub
ErrTrap:
End Sub
Private Sub billto_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    DcbosubContractor.SetFocus
    End If
  Exit Sub
ErrTrap:
End Sub
Private Sub DcbosubContractor_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    cboDiscount1.SetFocus
End If
  Exit Sub
ErrTrap:
End Sub
Private Sub cboDiscount1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    txtDiscount1.SetFocus
End If

  Exit Sub
ErrTrap:
End Sub
Private Sub txtDiscount1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    cboDiscount2.SetFocus
End If
calcnet
  Exit Sub
ErrTrap:
End Sub
Private Sub cboDiscount2_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    txtDiscount2.SetFocus
End If
  Exit Sub
ErrTrap:
End Sub
Private Sub txtDiscount2_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    dcBranch.SetFocus
End If
  Exit Sub
ErrTrap:
End Sub
Private Sub Dcbranch_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    txtprojectname.SetFocus
End If
  Exit Sub
ErrTrap:
End Sub
Private Sub txtprojectname_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    DcAccount1.SetFocus
End If

  Exit Sub
ErrTrap:
End Sub
Private Sub DcAccount1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    bill_Type.SetFocus
End If
  Exit Sub
ErrTrap:
End Sub
Private Sub bill_Type_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
    TxtRemarks.SetFocus
End If
  Exit Sub
ErrTrap:
End Sub
Private Sub TxtRemarks_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
If KeyAscii = 13 Then
If Command1(1).Enabled = True Then
   Command1(1).SetFocus
   Else
   Command1(3).SetFocus
   End If
End If
  Exit Sub
ErrTrap:
End Sub


 

Public Function SetFgForNewRow(fg As Object, _
                               IntColIndex As Long) As Long

    Dim i As Long
    Dim LngTempRow As Long

    With fg

        For i = .FixedRows - 1 To .Rows - 1

            If Trim$(.TextMatrix(i, IntColIndex)) = "" Then
                Exit For
            End If

        Next i

        If .FixedRows = .Rows Then
            .Rows = .Rows + 1
            LngTempRow = .Rows - 1
        ElseIf i - 1 = .Rows - 1 Then
            .Rows = .Rows + 1
            LngTempRow = .Rows - 1
        ElseIf i < .Rows - 1 Then
            LngTempRow = i
        Else
            .Rows = .Rows + 1
            LngTempRow = .Rows - 1
        End If
    If LngTempRow > 1 Then
        If Trim$(.TextMatrix(LngTempRow - 1, IntColIndex)) = "" Then
            LngTempRow = LngTempRow - 1
        End If
    End If
        SetFgForNewRow = LngTempRow
    End With

End Function

