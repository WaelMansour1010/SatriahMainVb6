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
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form projectsbill 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   13800
   ClientLeft      =   -1485
   ClientTop       =   -1905
   ClientWidth     =   22635
   Icon            =   "projectsbill.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13804.24
   ScaleMode       =   0  'User
   ScaleWidth      =   22635
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
      TabIndex        =   138
      Top             =   24960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   142
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4440
         TabIndex        =   141
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   140
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   372
         Left            =   2760
         TabIndex        =   139
         Top             =   120
         Width           =   1452
      End
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc1"
      Height          =   288
      ItemData        =   "projectsbill.frx":000C
      Left            =   28440
      List            =   "projectsbill.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   137
      Top             =   1440
      Visible         =   0   'False
      Width           =   4932
   End
   Begin VB.TextBox Text5 
      DataField       =   "last_root"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   2280
      TabIndex        =   128
      Text            =   "Text5"
      Top             =   16440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2280
      TabIndex        =   127
      Top             =   13920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   2280
      TabIndex        =   124
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
         TabIndex        =   126
         Top             =   120
         Width           =   495
      End
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M15"
         Height          =   255
         Left            =   3360
         TabIndex        =   125
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   10920
      TabIndex        =   123
      Top             =   25920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   1080
      TabIndex        =   118
      Top             =   14220
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
         TabIndex        =   122
         Top             =   120
         Visible         =   0   'False
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
         TabIndex        =   121
         Top             =   1320
         Visible         =   0   'False
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
         TabIndex        =   120
         Top             =   1440
         Visible         =   0   'False
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
         TabIndex        =   119
         Top             =   1500
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   765
      Left            =   12660
      TabIndex        =   114
      Top             =   14550
      Visible         =   0   'False
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
            MICON           =   "projectsbill.frx":003A
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
         FormatString    =   $"projectsbill.frx":0056
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
         MICON           =   "projectsbill.frx":03D4
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
         MICON           =   "projectsbill.frx":03F0
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
         MICON           =   "projectsbill.frx":040C
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
         MICON           =   "projectsbill.frx":0428
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
      Top             =   12570
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
            ButtonImage     =   "projectsbill.frx":0444
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
         FormatString    =   $"projectsbill.frx":07DE
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
         MICON           =   "projectsbill.frx":09A6
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
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   6690
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
         Left            =   -6510
         TabIndex        =   73
         Top             =   1440
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
         FormatString    =   $"projectsbill.frx":09C2
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
         MICON           =   "projectsbill.frx":0AD0
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
         FormatString    =   $"projectsbill.frx":0AEC
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
         MICON           =   "projectsbill.frx":0CC5
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
      Height          =   13800
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   22635
      _cx             =   39926
      _cy             =   24342
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
      Begin VB.TextBox TXTIban 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2100
         TabIndex        =   290
         Top             =   30
         Width           =   1020
      End
      Begin VB.ComboBox DefaultInvoicetype 
         Height          =   315
         ItemData        =   "projectsbill.frx":0CE1
         Left            =   18570
         List            =   "projectsbill.frx":0CE3
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   283
         Top             =   720
         Width           =   1890
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6360
         Left            =   2280
         TabIndex        =   212
         TabStop         =   0   'False
         Top             =   4860
         Visible         =   0   'False
         Width           =   20010
         _cx             =   35295
         _cy             =   11218
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
         Begin VB.TextBox txtTotalBondHistory 
            Alignment       =   1  'Right Justify
            Height          =   525
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   280
            Top             =   3030
            Width           =   5085
         End
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   2760
            Left            =   -150
            TabIndex        =   213
            Tag             =   "1"
            Top             =   300
            Visible         =   0   'False
            Width           =   15870
            _cx             =   27993
            _cy             =   4868
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
            FormatString    =   $"projectsbill.frx":0CE5
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
         Begin VSFlex8UCtl.VSFlexGrid GrdBondHistory 
            Height          =   3060
            Left            =   90
            TabIndex        =   277
            ToolTipText     =   "«÷€ÿ „— Ì‰ ·› Õ «·›« Ê—…"
            Top             =   60
            Visible         =   0   'False
            Width           =   15675
            _cx             =   27649
            _cy             =   5397
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
            ForeColor       =   -2147483640
            BackColorFixed  =   14871017
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16777215
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
            Rows            =   50
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"projectsbill.frx":0E28
            ScrollTrack     =   -1  'True
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
            RightToLeft     =   0   'False
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
            TabIndex        =   217
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Index           =   6
            Left            =   3990
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   4200
            Width           =   7095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   900
         Left            =   0
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   12900
         Width           =   22635
         _cx             =   39926
         _cy             =   1588
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
            Height          =   1050
            Index           =   0
            Left            =   20910
            TabIndex        =   149
            Top             =   0
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":101B
            PICN            =   "projectsbill.frx":1037
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
            Height          =   1050
            Index           =   1
            Left            =   17730
            TabIndex        =   150
            Top             =   0
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":7899
            PICN            =   "projectsbill.frx":78B5
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
            Height          =   1050
            Index           =   3
            Left            =   19425
            TabIndex        =   151
            Top             =   0
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":E117
            PICN            =   "projectsbill.frx":E133
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
            Height          =   1050
            Index           =   6
            Left            =   15990
            TabIndex        =   152
            Top             =   0
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":14995
            PICN            =   "projectsbill.frx":149B1
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
            Height          =   1065
            Index           =   7
            Left            =   6525
            TabIndex        =   153
            Top             =   -15
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1879
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
            MICON           =   "projectsbill.frx":1B213
            PICN            =   "projectsbill.frx":1B22F
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
            Height          =   1050
            Index           =   8
            Left            =   0
            TabIndex        =   154
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":21A91
            PICN            =   "projectsbill.frx":21AAD
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
            Height          =   1050
            Index           =   9
            Left            =   14655
            TabIndex        =   155
            Top             =   0
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":2830F
            PICN            =   "projectsbill.frx":2832B
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
            Height          =   1050
            Index           =   10
            Left            =   12960
            TabIndex        =   156
            Top             =   0
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":2EB8D
            PICN            =   "projectsbill.frx":2EBA9
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
            Height          =   1050
            Index           =   11
            Left            =   7830
            TabIndex        =   157
            Top             =   0
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":3540B
            PICN            =   "projectsbill.frx":35427
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
            Height          =   1050
            Index           =   12
            Left            =   11280
            TabIndex        =   158
            Top             =   0
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":3BC89
            PICN            =   "projectsbill.frx":3BCA5
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
            Height          =   1050
            Index           =   15
            Left            =   8970
            TabIndex        =   165
            Top             =   0
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":42507
            PICN            =   "projectsbill.frx":42523
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
            Height          =   1050
            Index           =   16
            Left            =   3555
            TabIndex        =   195
            Top             =   0
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   1852
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
            MICON           =   "projectsbill.frx":48D85
            PICN            =   "projectsbill.frx":48DA1
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
            Height          =   495
            Left            =   1620
            TabIndex        =   215
            Top             =   0
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "Submit for Approval"
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
            Height          =   525
            Index           =   0
            Left            =   1620
            TabIndex        =   216
            Top             =   450
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   926
            ButtonPositionImage=   1
            Caption         =   "Approval Status"
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
         Begin ALLButtonS.ALLButton Command1 
            Height          =   1065
            Index           =   17
            Left            =   1710
            TabIndex        =   281
            Top             =   18240
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   1879
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
            MICON           =   "projectsbill.frx":4F603
            PICN            =   "projectsbill.frx":4F61F
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
            Height          =   1065
            Index           =   18
            Left            =   5370
            TabIndex        =   282
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1879
            BTYPE           =   3
            TX              =   "Print 2"
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
            MICON           =   "projectsbill.frx":55E81
            PICN            =   "projectsbill.frx":55E9D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   26955
         TabIndex        =   133
         Top             =   2355
         Width           =   2325
      End
      Begin C1SizerLibCtl.C1Elastic Frame2 
         Height          =   6510
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   5730
         Width           =   22500
         _cx             =   39688
         _cy             =   11483
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
         Begin VB.TextBox txtTotalBefore 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   690
            Left            =   13620
            Locked          =   -1  'True
            TabIndex        =   275
            Top             =   5280
            Width           =   1515
         End
         Begin VB.TextBox txtDiscount4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   690
            Left            =   12105
            TabIndex        =   273
            Top             =   5280
            Width           =   1455
         End
         Begin VB.TextBox txtDiscountG2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   690
            Left            =   18540
            Locked          =   -1  'True
            TabIndex        =   269
            Top             =   5280
            Width           =   2025
         End
         Begin VB.TextBox Results2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   705
            Left            =   20565
            TabIndex        =   268
            Top             =   5265
            Width           =   1830
         End
         Begin VB.TextBox total2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   690
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   267
            Top             =   5280
            Width           =   2325
         End
         Begin VB.TextBox TxtFATValue2 
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
            Height          =   690
            Left            =   6030
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   266
            Top             =   5280
            Width           =   1530
         End
         Begin VB.TextBox TxtTotalValue2 
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
            Height          =   690
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   265
            Top             =   5280
            Width           =   2130
         End
         Begin VB.TextBox TxtPerforValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   705
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   264
            Top             =   5265
            Width           =   1920
         End
         Begin VB.TextBox advancedPayment22 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   690
            Left            =   15225
            Locked          =   -1  'True
            TabIndex        =   263
            Top             =   5280
            Width           =   1530
         End
         Begin VB.TextBox TxtPreVAT22 
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
            Height          =   690
            Left            =   4260
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   262
            Top             =   5280
            Width           =   1590
         End
         Begin VB.TextBox txtDiscountGMater 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   690
            Left            =   16785
            TabIndex        =   257
            Top             =   5280
            Width           =   1710
         End
         Begin VB.TextBox TxtPreVAT2 
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
            Height          =   690
            Left            =   4260
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   255
            Top             =   5280
            Width           =   1590
         End
         Begin VB.TextBox advancedPayment2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   690
            Left            =   15225
            Locked          =   -1  'True
            TabIndex        =   253
            Top             =   5280
            Width           =   1530
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H0080FFFF&
            Caption         =   "»Ì«‰«  «·œ›⁄«  «·„ﬁœ„…"
            Height          =   4590
            Left            =   1935
            RightToLeft     =   -1  'True
            TabIndex        =   229
            Top             =   195
            Visible         =   0   'False
            Width           =   18405
            Begin VB.TextBox TxtPreBalaVATYu 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   8280
               TabIndex        =   239
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
               TabIndex        =   238
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaTransPyed 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1560
               TabIndex        =   237
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaRemain 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   3000
               TabIndex        =   236
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaPayed 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   4440
               TabIndex        =   235
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaTotal 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   5880
               TabIndex        =   234
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaVAT 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   7320
               TabIndex        =   233
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox TxtPreBalaValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   8760
               TabIndex        =   232
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
               TabIndex        =   231
               Top             =   540
               Width           =   1200
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H8000000B&
               Caption         =   "«·€«¡ «·”œ«œ"
               Height          =   315
               Left            =   11400
               RightToLeft     =   -1  'True
               TabIndex        =   230
               Top             =   480
               Width           =   1695
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
               CausesValidation=   0   'False
               Height          =   1740
               Left            =   120
               TabIndex        =   240
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
               FormatString    =   $"projectsbill.frx":5C6FF
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
               TabIndex        =   252
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
               TabIndex        =   251
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
               TabIndex        =   250
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
               TabIndex        =   249
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
               TabIndex        =   248
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
               TabIndex        =   247
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
               TabIndex        =   246
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
               TabIndex        =   245
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
               TabIndex        =   244
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
               TabIndex        =   243
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
               TabIndex        =   242
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
               TabIndex        =   241
               Top             =   240
               Width           =   135
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   4770
            Left            =   0
            ScaleHeight     =   4710
            ScaleWidth      =   4230
            TabIndex        =   228
            Top             =   0
            Visible         =   0   'False
            Width           =   4290
         End
         Begin VB.TextBox txtManulaVat 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   690
            Left            =   7695
            TabIndex        =   210
            Top             =   5280
            Width           =   1005
         End
         Begin XtremeSuiteControls.CheckBox ChQty 
            Height          =   405
            Left            =   15060
            TabIndex        =   209
            Top             =   0
            Width           =   3435
            _Version        =   786432
            _ExtentX        =   6059
            _ExtentY        =   714
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
            ItemData        =   "projectsbill.frx":5C9BE
            Left            =   6405
            List            =   "projectsbill.frx":5C9E0
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   6030
            Width           =   6495
         End
         Begin VB.TextBox TxtNetValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   270
            Left            =   17880
            Locked          =   -1  'True
            TabIndex        =   204
            Top             =   6090
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox TxtPerforValue 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   690
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   202
            Top             =   5280
            Width           =   1920
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
            Height          =   690
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   5280
            Width           =   2145
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
            Height          =   690
            Left            =   6000
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   5280
            Width           =   1590
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
            Height          =   690
            Left            =   8745
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   5280
            Width           =   1005
         End
         Begin VB.Frame Frame7 
            Height          =   4230
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   510
            Width           =   5205
            Begin VB.TextBox TxtValueTemp 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   0
               TabIndex        =   199
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
               TabIndex        =   175
               Top             =   1920
               Width           =   1815
            End
            Begin VB.TextBox TxtBillNo 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   360
               TabIndex        =   171
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox TxtPeriod 
               Height          =   315
               Left            =   1440
               TabIndex        =   169
               Top             =   1440
               Width           =   735
            End
            Begin VB.ComboBox DcbPeriodType 
               Height          =   315
               ItemData        =   "projectsbill.frx":5CA35
               Left            =   360
               List            =   "projectsbill.frx":5CA42
               TabIndex        =   168
               Top             =   1440
               Width           =   975
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   10
               Left            =   3360
               TabIndex        =   167
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
               ButtonImage     =   "projectsbill.frx":5CA55
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker StartDate 
               Height          =   315
               Left            =   360
               TabIndex        =   173
               Top             =   1080
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   214237185
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DateTemp 
               Height          =   315
               Left            =   2040
               TabIndex        =   178
               Top             =   2400
               Visible         =   0   'False
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   214237185
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
               TabIndex        =   177
               Top             =   240
               Width           =   1890
            End
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               Caption         =   "„·«ÕŸ«  "
               Height          =   240
               Left            =   2160
               TabIndex        =   176
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
               TabIndex        =   174
               Top             =   1125
               Width           =   1890
            End
            Begin VB.Label Label40 
               Alignment       =   2  'Center
               Caption         =   "⁄œœ «·›Ê« Ì—"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   2280
               TabIndex        =   172
               Top             =   720
               Width           =   1890
            End
            Begin VB.Label Label39 
               Alignment       =   2  'Center
               Caption         =   "«·„œ… »Ì‰ «·›Ê« Ì—"
               Height          =   240
               Left            =   2280
               TabIndex        =   170
               Top             =   1440
               Width           =   1890
            End
         End
         Begin VB.TextBox total 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   690
            Left            =   9810
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   5280
            Width           =   2325
         End
         Begin VB.TextBox Results 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   705
            Left            =   20625
            TabIndex        =   50
            Top             =   5325
            Width           =   390
         End
         Begin VB.TextBox txtDiscountG 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   690
            Left            =   18585
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   5280
            Width           =   2010
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
            Height          =   4305
            Left            =   195
            TabIndex        =   52
            Top             =   510
            Width           =   22185
            _cx             =   39132
            _cy             =   7594
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
            Cols            =   56
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"projectsbill.frx":5CFEF
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
            Left            =   0
            TabIndex        =   53
            Tag             =   "Delete Row"
            Top             =   4830
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   767
            BTYPE           =   3
            TX              =   "Remove"
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
            MICON           =   "projectsbill.frx":5DAB7
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
            Left            =   14490
            TabIndex        =   163
            Top             =   6030
            Width           =   6750
            _ExtentX        =   11906
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AccountVat 
            Bindings        =   "projectsbill.frx":5DAD3
            Height          =   315
            Left            =   0
            TabIndex        =   189
            Top             =   6030
            Visible         =   0   'False
            Width           =   4230
            _ExtentX        =   7461
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
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   465
            Left            =   18765
            TabIndex        =   206
            Tag             =   "Delete Row"
            Top             =   0
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   820
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
            MICON           =   "projectsbill.frx":5DAE8
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
            Height          =   435
            Left            =   780
            TabIndex        =   270
            Tag             =   "Delete Row"
            Top             =   4830
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   767
            BTYPE           =   3
            TX              =   "Add"
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
            MICON           =   "projectsbill.frx":5DB04
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label66 
            Alignment       =   2  'Center
            Caption         =   "Net Invoice"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   13455
            TabIndex        =   276
            Top             =   4950
            Width           =   1935
         End
         Begin VB.Label Label62 
            Alignment       =   2  'Center
            Caption         =   "Discount"
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   1
            Left            =   11700
            TabIndex        =   274
            Top             =   4950
            Width           =   1920
         End
         Begin VB.Label Label63 
            Alignment       =   2  'Center
            Caption         =   "Material Discount"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   16890
            TabIndex        =   258
            Top             =   4950
            Width           =   1560
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "D.P .Vat"
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   1
            Left            =   4125
            RightToLeft     =   -1  'True
            TabIndex        =   256
            Top             =   4950
            Width           =   1710
         End
         Begin VB.Label Label62 
            Alignment       =   2  'Center
            Caption         =   "DP Discount"
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   14670
            TabIndex        =   254
            Top             =   4950
            Width           =   2565
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Manual Ratio"
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
            Height          =   285
            Index           =   148
            Left            =   6660
            TabIndex        =   211
            Top             =   4950
            Width           =   2085
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Form"
            Height          =   255
            Index           =   0
            Left            =   12915
            RightToLeft     =   -1  'True
            TabIndex        =   208
            Top             =   6045
            Width           =   1350
         End
         Begin VB.Label Label60 
            Alignment       =   2  'Center
            Caption         =   "’«›Ì «·›« Ê—…"
            ForeColor       =   &H00C00000&
            Height          =   165
            Left            =   18000
            TabIndex        =   205
            Top             =   5370
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label Label59 
            Alignment       =   2  'Center
            Caption         =   $"projectsbill.frx":5DB20
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   2535
            TabIndex        =   203
            Top             =   4950
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Vat Value"
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   67
            Left            =   6090
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   4950
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   68
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   4950
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Vat"
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   66
            Left            =   8790
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   4950
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   255
            Index           =   22
            Left            =   21300
            TabIndex        =   164
            Top             =   6030
            Width           =   1125
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   180
            Left            =   150
            TabIndex        =   162
            Top             =   6270
            Width           =   1335
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   165
            Left            =   2955
            TabIndex        =   161
            Top             =   6285
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   195
            Index           =   20
            Left            =   3885
            TabIndex        =   160
            Top             =   6270
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   21
            Left            =   1545
            TabIndex        =   159
            Top             =   6240
            Width           =   1350
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            Caption         =   "Total"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   20715
            TabIndex        =   56
            Top             =   4950
            Width           =   1440
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "Discount"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   18615
            TabIndex        =   55
            Top             =   4950
            Width           =   1905
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "Net Invoice"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   9960
            TabIndex        =   54
            Top             =   4950
            Width           =   1950
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frame13 
         Height          =   4560
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1230
         Width           =   22515
         _cx             =   39714
         _cy             =   8043
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
         Begin VB.TextBox txtPerformanceBond 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   14490
            TabIndex        =   292
            Top             =   2220
            Width           =   1995
         End
         Begin VB.CheckBox chkTaxExempt 
            Caption         =   "Tax-exempt"
            Height          =   345
            Left            =   12360
            RightToLeft     =   -1  'True
            TabIndex        =   291
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txt_Currency_rate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5370
            RightToLeft     =   -1  'True
            TabIndex        =   284
            Text            =   "1"
            Top             =   555
            Width           =   765
         End
         Begin VB.TextBox txtBondAmt 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   8190
            RightToLeft     =   -1  'True
            TabIndex        =   279
            Top             =   405
            Width           =   1575
         End
         Begin VB.TextBox txtDiscount3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   6360
            TabIndex        =   271
            Top             =   2415
            Width           =   2055
         End
         Begin VB.TextBox txtDiscountAccountCode 
            Alignment       =   1  'Right Justify
            Height          =   465
            Left            =   10035
            TabIndex        =   259
            Top             =   2895
            Width           =   1875
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   465
            Left            =   1380
            TabIndex        =   227
            Top             =   -150
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "projectsbill.frx":5DB3A
            Left            =   11850
            List            =   "projectsbill.frx":5DB3C
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   225
            Top             =   195
            Width           =   1515
         End
         Begin VB.TextBox TXTOrDer_no2 
            Alignment       =   1  'Right Justify
            Height          =   465
            Left            =   9765
            TabIndex        =   224
            Top             =   165
            Width           =   2085
         End
         Begin VB.TextBox TxtPreVAT 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   5340
            TabIndex        =   196
            Top             =   3960
            Width           =   2265
         End
         Begin VB.TextBox TxtExPercen 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   9720
            TabIndex        =   193
            Top             =   3930
            Width           =   2190
         End
         Begin VB.ComboBox DcbExPercen 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "projectsbill.frx":5DB3E
            Left            =   9720
            List            =   "projectsbill.frx":5DB4B
            TabIndex        =   191
            Top             =   3390
            Width           =   2190
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   690
            Left            =   18225
            TabIndex        =   183
            Top             =   165
            Width           =   2130
         End
         Begin VB.TextBox TxtAccountUnderImp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   450
            TabIndex        =   182
            Top             =   1995
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "›⁄·Ì"
            Height          =   330
            Left            =   1695
            TabIndex        =   181
            Top             =   2700
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0FFFF&
            Caption         =   " ﬁœÌ—Ì"
            Height          =   330
            Left            =   3150
            TabIndex        =   180
            Top             =   2700
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0FFFF&
            Caption         =   " Õ  «· ‰›Ì–"
            Height          =   330
            Left            =   150
            TabIndex        =   179
            Top             =   2700
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.TextBox DcAccount1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   540
            Left            =   5355
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1425
            Visible         =   0   'False
            Width           =   6555
         End
         Begin VB.TextBox DcAccount2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   420
            Left            =   14535
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1425
            Width           =   5835
         End
         Begin VB.ComboBox billto 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "projectsbill.frx":5DB60
            Left            =   14520
            List            =   "projectsbill.frx":5DB6A
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1905
            Width           =   5835
         End
         Begin VB.ComboBox bill_Type 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "projectsbill.frx":5DB86
            Left            =   14535
            List            =   "projectsbill.frx":5DB88
            TabIndex        =   19
            Top             =   885
            Width           =   1830
         End
         Begin VB.TextBox txtid 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   675
            Left            =   18225
            TabIndex        =   18
            Top             =   210
            Width           =   2130
         End
         Begin VB.TextBox txtprojectname 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            DataField       =   "project_name"
            DataSource      =   "Adodc1"
            Height          =   540
            Left            =   5355
            TabIndex        =   17
            Top             =   885
            Width           =   4470
         End
         Begin VB.TextBox TxtRemarks 
            Height          =   1095
            Left            =   210
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   3390
            Width           =   4740
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   18825
            TabIndex        =   15
            Top             =   2700
            Width           =   1590
         End
         Begin VB.ComboBox cboDiscount1 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "projectsbill.frx":5DB8A
            Left            =   17160
            List            =   "projectsbill.frx":5DB97
            TabIndex        =   14
            Top             =   3210
            Width           =   3195
         End
         Begin VB.ComboBox cboDiscount2 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "projectsbill.frx":5DBAC
            Left            =   17160
            List            =   "projectsbill.frx":5DBB9
            TabIndex        =   13
            Top             =   3735
            Width           =   3195
         End
         Begin VB.TextBox txtDiscount1 
            BackColor       =   &H00C0FFFF&
            Height          =   435
            Left            =   14535
            TabIndex        =   12
            Top             =   3210
            Width           =   2505
         End
         Begin VB.TextBox txtDiscount2 
            BackColor       =   &H00C0FFFF&
            Height          =   585
            Left            =   14535
            TabIndex        =   11
            Top             =   3780
            Width           =   2505
         End
         Begin VB.TextBox txtManualNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            DataSource      =   "Adodc1"
            Height          =   675
            Left            =   14490
            TabIndex        =   10
            Top             =   210
            Width           =   1815
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   9855
            TabIndex        =   9
            Top             =   2415
            Width           =   1995
         End
         Begin VB.TextBox advancedPayment 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   525
            Left            =   5370
            TabIndex        =   8
            Top             =   3390
            Width           =   2265
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   9780
            TabIndex        =   23
            Top             =   885
            Width           =   2130
            _ExtentX        =   3757
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
            Height          =   480
            Left            =   18150
            TabIndex        =   24
            Top             =   885
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   847
            _Version        =   393216
            Format          =   214302721
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker dueDate1 
            Height          =   390
            Left            =   5355
            TabIndex        =   25
            Top             =   1965
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   688
            _Version        =   393216
            Format          =   214302721
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Height          =   315
            Left            =   5355
            TabIndex        =   26
            Top             =   165
            Width           =   2145
            _ExtentX        =   3784
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
            Height          =   405
            Left            =   9750
            TabIndex        =   27
            Top             =   1995
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   714
            _Version        =   393216
            Format          =   214302721
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbosubContractor 
            Height          =   315
            Left            =   14490
            TabIndex        =   28
            Top             =   2715
            Width           =   4215
            _ExtentX        =   7435
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
            Height          =   645
            Left            =   4470
            TabIndex        =   198
            Tag             =   "Delete Row"
            Top             =   2490
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   1138
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
            MICON           =   "projectsbill.frx":5DBCE
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
            Height          =   375
            Left            =   1785
            TabIndex        =   200
            Top             =   2100
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   661
            _Version        =   393216
            Format          =   214302721
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcDiscountAccount 
            Height          =   315
            Left            =   6600
            TabIndex        =   260
            Top             =   2895
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   390
            Index           =   1
            Left            =   8505
            TabIndex        =   278
            Top             =   75
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   688
            ButtonPositionImage=   1
            Caption         =   "Modifications "
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
         Begin MSDataListLib.DataCombo DcCurrency 
            Height          =   315
            Left            =   6210
            TabIndex        =   285
            Top             =   540
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Retention Release"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   15945
            TabIndex        =   293
            Top             =   2325
            Width           =   2280
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            Height          =   300
            Index           =   65
            Left            =   6750
            RightToLeft     =   -1  'True
            TabIndex        =   286
            Top             =   570
            Width           =   1080
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            Caption         =   "Other deductions"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   8415
            TabIndex        =   272
            Top             =   2460
            Width           =   1335
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            Caption         =   "deductions account"
            Height          =   330
            Left            =   11535
            RightToLeft     =   -1  'True
            TabIndex        =   261
            Top             =   3000
            Width           =   2325
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Based on"
            Height          =   360
            Index           =   56
            Left            =   13095
            TabIndex        =   226
            Top             =   255
            Width           =   1245
         End
         Begin VB.Label Label58 
            Alignment       =   2  'Center
            Caption         =   " «—ÌŒ »œ«Ì… «·„‘—Ê⁄"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   3570
            TabIndex        =   201
            Top             =   2100
            Width           =   1740
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            Caption         =   "VAT «·œ›⁄… «·„ﬁœ„…"
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   7575
            TabIndex        =   197
            Top             =   3960
            Width           =   2355
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            Caption         =   "ﬁÌ„…"
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   12090
            TabIndex        =   194
            Top             =   3960
            Width           =   2370
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "‰Ê⁄ «·„” Œ·’"
            Height          =   360
            Left            =   11730
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   3390
            Width           =   2130
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   $"projectsbill.frx":5DBEA
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
            Height          =   1515
            Index           =   5
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   345
            Width           =   5205
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Õ Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   8250
            TabIndex        =   46
            Top             =   2070
            Width           =   1170
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
            Height          =   270
            Index           =   4
            Left            =   2955
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   60
            Width           =   1515
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   1650
            Left            =   150
            Top             =   300
            Width           =   5100
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ „ﬁ«Ê· «·»«ÿ‰"
            Height          =   480
            Left            =   11670
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   1425
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            Caption         =   "‰Ê⁄ «·„” Œ·’"
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   16245
            TabIndex        =   43
            Top             =   885
            Width           =   2025
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„” Œ·’ «·Ï"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   19740
            TabIndex        =   42
            Top             =   1890
            Width           =   2580
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "«· «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   20820
            TabIndex        =   41
            Top             =   885
            Width           =   1500
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·›« Ê—… - «·„” Œ·’"
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   19740
            TabIndex        =   40
            Top             =   210
            Width           =   2580
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ «·⁄„Ì· «·‰Â«∆Ì"
            ForeColor       =   &H00000000&
            Height          =   480
            Left            =   19740
            TabIndex        =   39
            Top             =   1425
            Width           =   2580
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "«”„ «·„‘—Ê⁄"
            Height          =   420
            Index           =   0
            Left            =   11670
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   885
            Width           =   2190
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "„·«ÕŸ« "
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1290
            TabIndex        =   37
            Top             =   3030
            Width           =   2355
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   7425
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   210
            Width           =   1215
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   11535
            TabIndex        =   35
            Top             =   2100
            Width           =   2325
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„ﬁ«Ê· «·»«ÿ‰"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   19740
            TabIndex        =   34
            Top             =   2700
            Width           =   2580
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ’„ ÷„«‰ «·«⁄„«·"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   19740
            TabIndex        =   33
            Top             =   3210
            Width           =   2580
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ’„  œ›⁄Â „ﬁœ„…"
            ForeColor       =   &H00000000&
            Height          =   480
            Left            =   19740
            TabIndex        =   32
            Top             =   3735
            Width           =   2580
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   16515
            TabIndex        =   31
            Top             =   210
            Width           =   1515
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "deductions"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   12570
            TabIndex        =   30
            Top             =   2520
            Width           =   1290
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            Caption         =   "«·œ›⁄Â «·„ﬁœ„…"
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   7605
            TabIndex        =   29
            Top             =   3390
            Width           =   2340
         End
      End
      Begin ALLButtonS.ALLButton CMD_language 
         Height          =   1065
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Language  «··€…"
         Top             =   210
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1879
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
         MICON           =   "projectsbill.frx":5DCA5
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
         Height          =   780
         Index           =   0
         Left            =   3675
         TabIndex        =   2
         Top             =   360
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   1376
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
         ButtonImage     =   "projectsbill.frx":5DCC1
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
         Height          =   780
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1376
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
         ButtonImage     =   "projectsbill.frx":5E05B
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
         Height          =   780
         Index           =   1
         Left            =   4485
         TabIndex        =   4
         Top             =   360
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1376
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
         ButtonImage     =   "projectsbill.frx":5E3F5
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
         Height          =   780
         Index           =   3
         Left            =   2955
         TabIndex        =   5
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   1376
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
         ButtonImage     =   "projectsbill.frx":5E78F
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
         Height          =   615
         Index           =   13
         Left            =   29715
         TabIndex        =   146
         Top             =   0
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1085
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
         MICON           =   "projectsbill.frx":5EB29
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
         Height          =   615
         Index           =   14
         Left            =   24585
         TabIndex        =   147
         Top             =   0
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1085
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
         MICON           =   "projectsbill.frx":5EB45
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DCDocTypes 
         Height          =   315
         Left            =   14400
         TabIndex        =   287
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker txtDateRec 
         Height          =   360
         Left            =   0
         TabIndex        =   289
         Top             =   0
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   635
         _Version        =   393216
         Format          =   209649665
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   122
         Left            =   15570
         TabIndex        =   288
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   345
         Index           =   11
         Left            =   150
         TabIndex        =   223
         Top             =   12330
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   315
         Index           =   10
         Left            =   5055
         TabIndex        =   222
         Top             =   12330
         Width           =   1620
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   315
         Index           =   9
         Left            =   8790
         TabIndex        =   221
         Top             =   12330
         Width           =   1620
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Total Retained Amounts"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   8
         Left            =   1920
         TabIndex        =   220
         Top             =   12330
         Width           =   2550
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Current Retained Amounts"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   7
         Left            =   6255
         TabIndex        =   219
         Top             =   12360
         Width           =   2460
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Previous Retained Amounts"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   3
         Left            =   10425
         TabIndex        =   218
         Top             =   12330
         Width           =   2505
      End
      Begin VB.Image ImgFavorites 
         Height          =   585
         Left            =   11640
         Picture         =   "projectsbill.frx":5EB61
         Stretch         =   -1  'True
         Top             =   165
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   12165
         Picture         =   "projectsbill.frx":627C9
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "—ﬁ„ «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   26415
         TabIndex        =   135
         Top             =   4530
         Width           =   2520
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   24810
         TabIndex        =   134
         Top             =   3375
         Width           =   1890
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "„” Œ·’«  «·„‘«—Ì⁄   "
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
         Height          =   1275
         Left            =   -1305
         TabIndex        =   6
         Top             =   0
         Width           =   23970
      End
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   9870
         Picture         =   "projectsbill.frx":63D36
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2985
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "projectsbill.frx":652A3
      Height          =   2892
      Left            =   9240
      TabIndex        =   129
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
      TabIndex        =   143
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
      MICON           =   "projectsbill.frx":652B8
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
      TabIndex        =   144
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
      MICON           =   "projectsbill.frx":652D4
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
      TabIndex        =   145
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
      MICON           =   "projectsbill.frx":652F0
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
      TabIndex        =   136
      Top             =   2640
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label5 
      Caption         =   "Label2"
      Height          =   15
      Left            =   12120
      TabIndex        =   132
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   12720
      TabIndex        =   131
      Top             =   12840
      Width           =   2172
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "·œ›⁄Â „ÕœœÂ"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   28560
      TabIndex        =   130
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
Attribute VB_Name = "projectsbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X                     As Long
        Dim des        As String
        Dim desEN        As String

Dim last_root             As Integer
Dim last_geeral           As Integer
Dim last_branch           As Integer
Dim mod_flad              As String
Dim first_run             As Boolean
Dim rs                    As ADODB.Recordset
Dim RsDev                 As ADODB.Recordset
Dim current_terms         As String
Dim current_opr           As String
Dim NewGrid               As New ClsGrid
Dim expanses_account      As String
Dim AcountGood            As String
Dim maa_rs                As ADODB.Recordset
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
Dim cCompanyInfo          As New ClsCompanyInfo
Dim FlgBillBuy            As Boolean
Dim i                     As Integer
Dim Export As Integer
Dim s As String
Dim zatcaStatus As Integer
Dim isFocus As Boolean
Dim IsFromItemID As Boolean
Sub ReloadContrac(Optional project_no As Double)
    Dim Dcombos As ClsDataCombos
    Dim StrSQL  As String
    Set Dcombos = New ClsDataCombos
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select CusID,CusName From TblCustemers"
    Else
        StrSQL = "Select CusID,CusNamee From TblCustemers"
    End If
    StrSQL = StrSQL & " where CusID in(SELECT     sub_contractor_id"
    StrSQL = StrSQL & " From dbo.projects_des"
    StrSQL = StrSQL & " WHERE     (project_id = " & project_no & "))"
    Dcombos.ClearMyDataCombo DcbosubContractor

    fill_combo Me.DcbosubContractor, StrSQL
End Sub
'ma
Public Sub Search(ID As Integer)

    Set maa_rs = New ADODB.Recordset
    '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 
    StrSQL = " SELECT *  From dbo.project_billl  where  ID=  " & ID & " Order by ID "
    
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

Private Sub ALLButton3_Click()
    
    
Dim X As Integer
Fg_Journal.rows = Fg_Journal.rows + 1
             
   


End Sub

Private Sub DcDiscountAccount_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '   Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", Val(Me.DcboRevenuesTypes.BoundText))
        
        txtDiscountAccountCode.text = getAccountSerial_Code("Account_Serial", "Account_Code", DcDiscountAccount.BoundText)
       
        'If Me.TxtModFlg.Text <> "R" Then
  
    End If

End Sub

Private Sub DcDiscountAccount_KeyUp(KeyCode As Integer, _
   Shift As Integer)

    If KeyCode = vbKeyF3 Then
        DcDiscountAccount.text = ""
        '   Unload Account_search
        Account_search.show
        Account_search.case_id = 2608178
            
    End If

End Sub

Private Sub maaRetrive()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 2
    Dim RsDev  As ADODB.Recordset
    Dim StrSQL As String
    Dim i      As Integer

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
    txtid.text = IIf(IsNull(maa_rs("id").value), 0, val(maa_rs("id").value))

    XPDtbTrans.value = IIf(IsNull(maa_rs("bill_date").value), Date, maa_rs("bill_date").value)
    dueDate.value = IIf(IsNull(maa_rs("dueDate").value), Date, maa_rs("dueDate").value)
    dueDate1.value = IIf(IsNull(maa_rs("dueDate1").value), Date, maa_rs("dueDate1").value)

    DataCombo2.BoundText = IIf(IsNull(rs("project_no").value), "", rs("project_no").value)
    '*************************************************
    DcbosubContractor.BoundText = IIf(IsNull(maa_rs("subContractorId").value), "", maa_rs("subContractorId").value)
    txtDiscount1.text = IIf(IsNull(maa_rs("discount1value").value), 0, (maa_rs("discount1value").value))
    txtDiscount2.text = IIf(IsNull(maa_rs("discount2value").value), 0, (maa_rs("discount2value").value))

    cboDiscount1.ListIndex = IIf(IsNull(maa_rs("discount1ID").value), 0, (maa_rs("discount1ID").value))
    cboDiscount2.ListIndex = IIf(IsNull(maa_rs("discount2ID").value), 0, (maa_rs("discount2ID").value))

    '*************************************************

    txtprojectname.text = IIf(IsNull(maa_rs("project_name").value), "", maa_rs("project_name").value)
    '    DcAccount1.text = IIf(IsNull(rs("Sub_user_name").value), "", rs("Sub_user_name").value)
    '    DcAccount2.text = IIf(IsNull(rs("End_user_name").value), "", rs("End_user_name").value)

    txtendaccount.text = IIf(IsNull(maa_rs("End_user_account").value), "", maa_rs("End_user_account").value)
    txtsubaccount.text = IIf(IsNull(maa_rs("Sub_user_account").value), "", maa_rs("Sub_user_account").value)
    txtrevenue_account.text = IIf(IsNull(maa_rs("revenue_account").value), "", maa_rs("revenue_account").value)

    'DcAccount4.text = IIf(IsNull(Rs("sub_contractor_name").value), "", Rs("sub_contractor_name").value)

    'DcAccount2.text = IIf(IsNull(Rs("End_user_name").value), "", Rs("End_user_name").value)

    billto.ListIndex = IIf(IsNull(maa_rs("bill_to").value), -1, maa_rs("bill_to").value)
    bill_Type.text = IIf(IsNull(maa_rs("bill_type").value), 0, val(maa_rs("bill_type").value))
    Me.note_id.text = IIf(IsNull(maa_rs("note_id").value), "", maa_rs("note_id").value)
    TxtNoteSerial.text = IIf(IsNull(maa_rs("NoteSerial").value), "", maa_rs("NoteSerial").value)
    txtRemarks.text = IIf(IsNull(maa_rs("Remarks").value), "", maa_rs("Remarks").value)
    TxtManualNO.text = IIf(IsNull(maa_rs("ManualNo").value), "", maa_rs("ManualNo").value)

    'rs("Remarks").value = Trim(TxtRemarks.text)
    'rs("ManualNo").value = Trim(txtManualNo.text)

    total.text = Round(IIf(IsNull(maa_rs("total").value), 0, maa_rs("total").value), Decimal_Places)

    'Exit Sub

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "SELECT     item_id,id, project_no, item, cost, exe, percentage, exedate, bill_id,item_unit ,Unit_id,Quantity,Price,Pre_Quantity,Pre_Value,Pre_Percent,Curr_Quantity,Curr_value,curr_Percent,tot_quantity,tot_value,tot_percent "
        StrSQL = StrSQL + " from dbo.project_bill_details "
        StrSQL = StrSQL + " Where bill_id =" & Me.txtid.text
    
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or maa_rs.EOF) Then
            'Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            'Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
    
            RsDev.MoveFirst
    
            With Me.Fg_Journal
                .rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .rows - 1
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
    If val(txtid.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
        Else
            MsgBox "Save Doc First", vbCritical
        End If
      
        Exit Sub
    End If
 
    SendTopost Me.Name, "project_billl", "id", 0, val(dcBranch.BoundText), val(txtid.text), TxtNoteSerial1.text, val(Me.note_id)
    
    rs.Resync
    If SystemOptions.UserInterface = ArabicInterface Then
        Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
    Else
        Accredit.Caption = "Sent To approval "
    End If
    Retrive (val(Me.txtid.text))

    C1Elastic3.Visible = True
End Sub

Function fillapprovData()
    Dim Num       As Integer
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL    As String
    Label11(6).Tag = "1"
    StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
    StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
    StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
    StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
    StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
    StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.txtid.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
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
        Grid2.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        
            Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
            If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
                Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
            Else
                Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
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
                        Label11(6).Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                        Label24.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                        Label11(6).Tag = "Posted"
                    Else
                        Label24.Caption = "Approved"
                        Label11(6).Tag = "Posted"
                    End If
                    Label24.backcolor = &H80FF80
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Label24.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                        Label11(6).Tag = "0"
                    Else
                        Label24.Caption = "Currently required Approve"
                        Label11(6).Tag = "0"
                    End If
                    Label24.backcolor = &HFFFFC0
                End If

            End If

        Next Num
    Else
        Grid2.rows = 1
    End If
    Label11(6).Caption = Label24.Caption
    Label11(6).backcolor = Label24.backcolor
    
    RsDetails.Close

End Function

Private Sub advancedPayment_Change()
    advancedPayment2 = advancedPayment
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
    VSFlexGrid4.rows = 2
    TxtPreBalaNet.text = 0
    TxtPreBalaValue.text = 0
    TxtPreBalaVAT.text = 0
    TxtPreBalaTotal.text = 0
    TxtPreBalaPayed.text = 0
    TxtPreBalaRemain.text = 0
    TxtPreBalaTransPyed.text = 0
    advancedPayment.text = 0
    BillCustomer
    ClcalteOpiningBalance
End Sub
Sub ClcalteOpiningBalance()
    Dim Valu       As Double
    Dim Percetage  As Double
    Dim Percetage2 As Double
    Dim RecDate    As Date
    GetBalanceProject Valu, RecDate
    TxtPreBalaTotal.text = Round(Valu, Decimal_Places)
    If val(TxtPreBalaTotal.text) > 0 Then
        PercentgValueAddedAccount_Transec RecDate, 6, 1, , Percetage
        TxtPreBalaVATYu.text = Percetage
        If Percetage <> 0 Then
            Percetage2 = Percetage / 100 + 1
            TxtPreBalaValue.text = Round(val(TxtPreBalaTotal.text) / Percetage2, Decimal_Places)
            TxtPreBalaVAT.text = Round(val(TxtPreBalaValue.text) * Percetage / 100, Decimal_Places)
        Else
            TxtPreBalaValue.text = Round(val(TxtPreBalaTotal.text), Decimal_Places)
            TxtPreBalaVAT.text = 0
        End If
    End If
    TxtPreBalaPayed.text = Round(GetValue(), Decimal_Places)
    TxtPreBalaRemain.text = Round(val(TxtPreBalaTotal.text) - val(TxtPreBalaPayed.text), Decimal_Places)
End Sub
Sub BillCustomer(Optional Ind As Integer = 0)
    Dim Msg         As String
    Dim mIsFoundRow As Boolean
    If VSFlexGrid4.rows > 1 Then
        If val(VSFlexGrid4.TextMatrix(1, VSFlexGrid4.ColIndex("NoteSerial1"))) <> 0 Then
            mIsFoundRow = True
        Else
            mIsFoundRow = False
        
        End If
    End If

    If val(TXTEnd_user_id.text) = 0 Then
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
        If Me.TxtModFlg.text <> "R" Then
            VSFlexGrid4.Enabled = True
        Else
            '   VSFlexGrid4.Enabled = False
        End If
        If Ind = 0 Then
            'XPTxtVal.Text = 0
        End If
        If Me.TxtModFlg.text = "N" Then
            If val(billto.ListIndex) = 0 Then
                RetriveBillBuy val(TXTEnd_user_id.text)
            ElseIf val(billto.ListIndex) = 1 Then
                RetriveBillBuy val(DcbosubContractor.BoundText)
                    
            End If
        End If
        If Me.TxtModFlg.text = "E" And (FlgBillBuy = True Or Not mIsFoundRow) Then
            '           VSFlexGrid4.Editable = True
        Else
            '   VSFlexGrid4.Editable = False
        End If
        If val(billto.ListIndex) = 0 Then
            RetriveBillBuy val(TXTEnd_user_id.text)
        ElseIf val(billto.ListIndex) = 1 Then
            RetriveBillBuy val(DcbosubContractor.BoundText)
                    
        End If
            
    End If
    
End Sub
Sub RetriveBillBuy(Optional CuID As Double = 0)
    Dim sql As String
    Dim Rs8 As ADODB.Recordset
    Dim i   As Integer
    Set Rs8 = New ADODB.Recordset
    With VSFlexGrid4
        .Clear flexClearScrollable, flexClearEverything
        .rows = 1
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
            .rows = 1
            .rows = .rows + Rs8.RecordCount
            .rows = .FixedRows + Rs8.RecordCount
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
    Dim rs        As ADODB.Recordset
    Dim i         As Double
    Dim StrSQL    As String
    Set RsDetails = New ADODB.Recordset
    StrSQL = "   SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblPayPrePayed.*"
    StrSQL = StrSQL & "  FROM         dbo.TblPayPrePayed LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblPayPrePayed.branch_no = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL & "  Where (dbo.TblPayPrePayed.NoteID1 = " & val(txtid.text) & ")"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid4
        .Clear flexClearScrollable, flexClearEverything
        .rows = .FixedRows
        If Not (RsDetails.BOF Or RsDetails.EOF) Then
            RsDetails.MoveFirst
            .rows = .FixedRows + RsDetails.RecordCount

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
                .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
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
    Dim i      As Integer
    Dim StrSQL As String
    With VSFlexGrid4
        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
                StrSQL = "Update Notes Set  TotalPayed=0 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
        Next i
    End With
End Sub
Function saveBillBuy()
    Dim StrSQL      As String
    Dim i           As Double
    Dim Diff        As Double
    Dim Note_Value1 As Double
    Diff = 0
    Dim RsDetails As ADODB.Recordset
    If Me.TxtModFlg.text = "E" Then
        StrSQL = "Delete From TblPayPrePayed Where NoteID1=" & val(txtid.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblProjePayPrePayed Where  NoteID=" & val(txtid.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     * from dbo.TblPayPrePayed Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid4
        TxtValueTemp.text = val(Label47.Caption)
    
        For i = .FixedRows To .rows - 1
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > 0 Then
                'Dim LineDiscountPercent As Double
                '        LineDiscountPercent = val(.TextMatrix(i, .ColIndex("NetExe"))) / Results.Text
                RsDetails.AddNew
            
                RsDetails("NoteID1").value = val(txtid.text)
                RsDetails("VATLine").value = val(.TextMatrix(i, .ColIndex("VATLine")))
                RsDetails("ValueLine").value = val(.TextMatrix(i, .ColIndex("ValueLine")))
                RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
                RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
                RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
                RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
                Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
                Diff = 0
                If val(TxtValueTemp.text) > 0 Then
                    If val(TxtValueTemp.text) <= Note_Value1 Then
                        Diff = val(TxtValueTemp.text)
                        TxtValueTemp.text = val(TxtValueTemp.text) - Note_Value1
                    Else
                        Diff = Note_Value1
                        TxtValueTemp.text = val(TxtValueTemp.text) - Note_Value1
                    End If
                End If
                .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
                RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
                RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
                RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
                RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
                RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
                RsDetails.update
                If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
                    StrSQL = "Update Notes Set  TotalPayed=1 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                    Cn.Execute StrSQL, , adExecuteNoRecords
                Else
                    StrSQL = "Update Notes Set  TotalPayed=Null Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                    Cn.Execute StrSQL, , adExecuteNoRecords
                End If
            End If
        Next i
    End With
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     * from dbo.TblProjePayPrePayed Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid4
        For i = .FixedRows To .rows - 1
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) Then
                RsDetails.AddNew
                RsDetails("NoteID").value = val(txtid.text)
                RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
                RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
                RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                RsDetails.update
            End If
        Next i
    End With

End Function

Private Sub ALLButton2_Click()
    FillAllBandsToGrid
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
 
            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid4

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    RelineBuy
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
On Error GoTo ErrTrap
  Dim BeginTrans As Boolean
                
    calcnet
    Dim accountdep As String

        
  
        
    If billto.ListIndex = 0 Then
        X = val(TXTEnd_user_id.text)
        'accountdep = txtendaccount.text
    Else

        If billto.ListIndex = 1 Then
            X = val(TXTsub_contractor_id.text)
            '    accountdep = txtsubaccount.text
        End If
    End If
    
        Dim AdvancedAccount  As String
    Dim GuranteeAccount  As String
    Dim mAccountMaterial As String
    
    X = val(TXTEnd_user_id.text)
    '  Dim x As Double
    '  x = get_Customer_id(accountdep)
        
    '  total.text = gettotal(txtid.text)
    
    
        If BeginTrans = False Then
            Cn.BeginTrans
            BeginTrans = True
        End If

    Dim Rs1 As New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=5000 and NoteSerial='" & Me.TxtNoteSerial.text & "' order by NoteID"
    Rs1.Open StrSQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
  
    If TxtModFlg.text = "N" Then
   
        If X = 0 Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "An error in customer Number", vbCritical: Exit Function
            Else
                MsgBox "ÌÊÃœ Œÿ√ ›Ì —ﬁ„ «·⁄„Ì·", vbCritical: Exit Function
            End If
        End If
        note_id.text = CStr(new_id("Notes", "NoteID", "", True))
        txtid.text = CStr(new_id("project_billl", "id", "", True))
            
        rs.AddNew
    
    Else
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(note_id.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From notes  Where NoteID=" & val(note_id.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        StrSQL = "Delete From project_bill_details Where bill_id=" & val(Me.txtid.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Rs1.AddNew
    'branch_id
    If TxtNoteSerial1.text = "" Then
        If billto.ListIndex = 0 Then
            TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 65, 65, , , , , val(billto.ListIndex))
        Else
            TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 84, 84, , , , , val(billto.ListIndex))
        End If
    End If
     
    Rs1("branch_no").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    Rs1("NoteID").value = val(note_id.text)
    Rs1("Note_Value").value = IIf(total.text = "", Null, val(total.text))
    Rs1("CusID").value = X
    Rs1("NoteType").value = 500
    Rs1("NoteType").value = 5000
    Rs1("NoteDate").value = XPDtbTrans.value
    Rs1("UserID").value = user_id
    Rs1("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", Trim(TxtNoteSerial1.text), Null)

        
    rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, val(Dccurrency.BoundText))
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)
    rs("DateRec").value = txtDateRec.value
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs("CIBAN").value = TXTIban.text
    rs("Invoicetype").value = Me.DefaultInvoicetype.ListIndex

    
    
    Rs1("RemarkE").value = IIf(Me.txtRemarks <> "", Trim(txtRemarks.text), Null)
    Rs1("Remark").value = IIf(Me.txtRemarks <> "", Trim(txtRemarks.text), Null)
   
    rs("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", Trim(TxtNoteSerial1.text), Null)
    rs("ExPercen").value = val(TxtExPercen.text)
    rs("DiscountGMater").value = val(txtDiscountGMater.text)
    
    rs("ExPercenID").value = val(DcbExPercen.ListIndex)
    rs("PreVAT").value = val(TxtPreVAT.text)
    rs("FATYou").value = val(TxtFATYou.text)
    rs("FATValue").value = val(TxtFATValue.text)
    rs("TotalValue").value = val(TxtTotalValue.text)
    rs("AccountCodeVat").value = Me.AccountVat.BoundText
    rs("DiscountAccount").value = Me.DcDiscountAccount.BoundText
    

    
    rs("NetValue").value = val(TxtNetValue.text)
    rs("PerforValue").value = val(TxtPerforValue.text)
    ''/////////
    rs("StartDateProje").value = StartDateProje.value
    rs("PreBalaValue").value = val(TxtPreBalaValue.text)
    rs("PreBalaVAT").value = val(TxtPreBalaVAT.text)
    rs("PreBalaTotal").value = val(TxtPreBalaTotal.text)
    rs("PreBalaPayed").value = val(TxtPreBalaPayed.text)
    rs("PreBalaRemain").value = val(TxtPreBalaRemain.text)
    rs("PreBalaTransPyed").value = val(TxtPreBalaTransPyed.text)
    rs("PreBalaNet").value = val(TxtPreBalaNet.text)
    rs("PreBalaVATYu").value = val(TxtPreBalaVATYu.text)
    rs("SumVATLine").value = val(Label57.Caption)
    rs("SumValueLine").value = val(Label56.Caption)
    ''/////
    If Option7.value = True Then
        rs("UnderImp").value = 0
    ElseIf Option6.value = True Then
        rs("UnderImp").value = 1
    ElseIf Option8.value = True Then
        rs("UnderImp").value = 2
    End If
    If TxtManualNO.text = "" Then
        TxtManualNO.text = TxtNoteSerial1.text
    End If


    Rs1("remark").value = "  Project Invoice No  :  " & TxtManualNO & CHR(13) & "  To Project " & txtprojectname.text
 
    '   Rs1("remark").value = "„” Œ·’ —ﬁ„ :     " & txtid & "    " & Chr(13) & "  ··„‘—Ê⁄  " & txtprojectname.text
    
    If TxtNoteSerial = "" Then
        TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
    End If
       
    Rs1("NoteSerial").value = TxtNoteSerial.text
    
    Rs1("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”·
    Rs1("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    '  rs("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
     
    Rs1("sanad_year").value = year(XPDtbTrans.value)
    Rs1("sanad_month").value = Month(XPDtbTrans.value)
    Rs1("note_value_by_characters").value = WriteNo(Format(Me.Results.text, "0.00"), 0, True, ".")
    
    Rs1.update
    
    rs("id").value = Me.txtid.text
    
    rs("bill_date").value = XPDtbTrans.value
    'branch_id
    rs("branch_no").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
    rs("project_no").value = IIf(Not IsNumeric(DataCombo2.BoundText), "", DataCombo2.BoundText)
    rs("project_name").value = txtprojectname.text
    rs("Sub_user_name").value = IIf(IsNull(DcAccount1.text), "", DcAccount1.text)
    rs("End_user_name").value = IIf(IsNull(DcAccount2.text), "", DcAccount2.text)
    rs("End_user_account").value = IIf(IsNull(txtendaccount.text), "", txtendaccount.text)
    rs("Sub_user_account").value = IIf(IsNull(txtsubaccount.text), "", txtsubaccount.text)
    rs("revenue_account").value = IIf(IsNull(txtrevenue_account.text), "", txtrevenue_account.text)
    rs("UserID").value = IIf(DCboUserName.BoundText <> "", val((DCboUserName.BoundText)), Null)
    rs("bill_to").value = billto.ListIndex
    rs("bill_type").value = bill_Type.ListIndex '  IIf(IsNull(bill_Type.text), "", bill_Type.text)
    rs("note_id").value = IIf(IsNull(note_id.text), "", note_id.text)
    rs("NoteSerial").value = IIf(IsNull(TxtNoteSerial.text), "", TxtNoteSerial.text)
    rs("total").value = IIf(Not IsNumeric(total.text), 0, total.text)
    rs("AccountUnderImp").value = TxtAccountUnderImp.text
    
    rs("CBoBasedON").value = CBoBasedON.ListIndex
    rs.Fields("OrDer_no2").value = val(TXTOrDer_no2.text)
    rs.Fields("OrDer_no").value = val(TXTOrDer_no.text)
    
    '26082015
    rs("Discount").value = IIf(Not IsNumeric(txtDiscount.text), 0, txtDiscount.text)
    rs("PerformanceBond").value = IIf(Not IsNumeric(txtPerformanceBond.text), 0, txtPerformanceBond.text)
    rs("AdvancedPayment").value = IIf(Not IsNumeric(advancedPayment.text), 0, advancedPayment.text)
    
    rs("Results").value = IIf(Not IsNumeric(Results.text), 0, Results.text)
    ''///////23 05 2016
    rs("BillNo").value = val(TxtBillNo.text)
    rs("StartDate").value = StartDate.value
    rs("Period").value = val(txtPeriod.text)
    rs("PeriodType").value = val(DcbPeriodType.ListIndex)
    rs("Remarks2").value = TxtRemarks2.text
 
    '26082015

    rs("dueDate").value = dueDate.value
    rs("dueDate1").value = dueDate1.value

    '*************************************************
    rs("subContractorId").value = IIf(Not IsNumeric(DcbosubContractor.BoundText), Null, DcbosubContractor.BoundText)
    rs("discount1ID").value = val(cboDiscount1.ListIndex)
    rs("discount2ID").value = val(cboDiscount2.ListIndex)
    rs("discount1value").value = val(txtDiscount1.text)
    rs("discount2value").value = val(txtDiscount2.text)
    rs("Remarks").value = Trim(txtRemarks.text)
    rs("ManualNo").value = Trim(TxtManualNO.text)
 
    '*************************************************
    
        SaveBillMonthly
   

    rs.update

    'SuppCreat4Acc
    Set RsDev = New ADODB.Recordset
    '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
    RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim LngDevID As Long
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
    accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", X, "Account_code")

    Dim Posted As Integer
    If CheckAprroveScreen(Me.Name) = True Then
        Posted = 1
    Else
        Posted = 0
    End If

    If billto.ListIndex = 0 Then
        Dim lineno As Integer
        lineno = 1
        '    If accountdep = "" Then GoTo ll
        '«·ÿ—› «·„œÌ‰
        RsDev.AddNew
    
        RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
        RsDev("DEV_ID_Line_No").value = lineno
        If Option8.value = True Then
            RsDev("Account_Code").value = Me.TxtAccountUnderImp.text
        Else
            RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
        End If
        
        
        If SystemOptions.CustCreat4Acc Then
            RsDev("Value").value = val(Me.TxtTotalValue.text)
        Else
            RsDev("Value").value = val(Me.total.text) + val(TxtFATValue.text)
        End If
        RsDev("Credit_Or_Debit").value = 0

        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
        RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & " Manual #   " & TxtManualNO & CHR(13) & txtRemarks.text

        RsDev("Notes_ID").value = val(note_id.text)
        RsDev("project_bill_no").value = val(txtid.text)
        RsDev("project_id").value = val(Me.DataCombo2.BoundText)
        RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
        RsDev("UserID").value = user_id
        RsDev("branch_id").value = my_branch
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    
        RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    
        RsDev.update
        'll:
        lineno = lineno + 1
line_no = line_no + 1
        '«·Õ”„Ì« 
        'Account_Code_dynamic1

        AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(val(TXTEnd_user_id.text)), "Account_CodeHi1")
        GuranteeAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(val(TXTEnd_user_id.text)), "Account_CodeAss2")
        mAccountMaterial = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(val(TXTEnd_user_id.text)), "Account_CodeHi2")
        
        
            If val(Me.txtDiscount3.text) > 0 Then
    
                If SystemOptions.SuppCreat4Acc Then
        
'                        RsDev.AddNew
'
'                        RsDev("branch_id").Value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'                        RsDev("Double_Entry_Vouchers_ID").Value = LngDevID
'                        RsDev("DEV_ID_Line_No").Value = lineno
'                        If SystemOptions.SuppCreat4Acc Then
'                            RsDev("Account_Code").Value = accountdep  '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
'                        Else
'                            RsDev("Account_Code").Value = Account_Code_dynamic1 '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
'                        End If
'                        RsDev("Value").Value = val(Me.txtDiscount3.text)
'                        RsDev("Credit_Or_Debit").Value = 1
'
'                        If SystemOptions.UserInterface = ArabicInterface Then
'                            RsDev("Double_Entry_Vouchers_Description").Value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & Chr(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & Chr(13) & TxtRemarks.text
'                        Else
'                            RsDev("Double_Entry_Vouchers_Description").Value = "  Project Invoice No  :  " & TxtNoteSerial1 & Chr(13) & "  To Project " & txtprojectname.text & "  Manual# " & txtManualNo & Chr(13) & TxtRemarks.text
'                        End If
'
'                        RsDev("Notes_ID").Value = val(note_id.text)
'                        RsDev("project_bill_no").Value = val(txtid.text)
'
'                        ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'
'                        RsDev("RecordDate").Value = XPDtbTrans.Value ' DateValue(Now)
'                        RsDev("UserID").Value = user_id
'                        RsDev("branch_id").Value = my_branch
'                        RsDev("Account_Interval_ID").Value = SystemOptions.SysCurrentAccountIntervalID
'                        RsDev("Posted").Value = IIf(Posted = 0, Null, Posted)
'                        RsDev.update
'                        'll:
'                        lineno = lineno + 1
            
                        RsDev.AddNew
                    line_no = line_no + 1
                        RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
                
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = line_no
                        If SystemOptions.SuppCreat4Acc Then
                            RsDev("Account_Code").value = DcDiscountAccount.BoundText   '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
                        Else
                            RsDev("Account_Code").value = Account_Code_dynamic1 '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
                        End If
                        RsDev("Value").value = val(Me.txtDiscount3.text)
                        RsDev("Credit_Or_Debit").value = 0
                
                        
                        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
                        RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  Manual# " & TxtManualNO & CHR(13) & txtRemarks.text
                        RsDev("Notes_ID").value = val(note_id.text)
                        RsDev("project_bill_no").value = val(txtid.text)
                    
                        RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                   
                        RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
                        RsDev("UserID").value = user_id
                        RsDev("branch_id").value = my_branch
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                        RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                        RsDev.update
                        'll:
                        lineno = lineno + 1
                        line_no = line_no + 1
                    End If
                End If
        
   
            line_no = line_no + 1
        If val(Me.txtDiscount.text) > 0 Then
    
            If SystemOptions.SuppCreat4Acc Then
    
                RsDev.AddNew
            
                RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
        
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
                If SystemOptions.SuppCreat4Acc Then
                    RsDev("Account_Code").value = DcDiscountAccount.BoundText  '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
                Else
                    RsDev("Account_Code").value = Account_Code_dynamic1 '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
                End If
                RsDev("Value").value = val(Me.txtDiscount.text)
                RsDev("Credit_Or_Debit").value = 0
        
                
                RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
                RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  Manual# " & TxtManualNO & CHR(13) & txtRemarks.text
        
                RsDev("Notes_ID").value = val(note_id.text)
                RsDev("project_bill_no").value = val(txtid.text)
            
                 RsDev("project_id").value = val(Me.DataCombo2.BoundText)
           
                RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
                RsDev("UserID").value = user_id
                RsDev("branch_id").value = my_branch
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                RsDev.update
                'll:
                lineno = lineno + 1
                line_no = line_no + 1
    
              '  RsDev.AddNew
            
'                RsDev("branch_id").Value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
'
'                RsDev("Double_Entry_Vouchers_ID").Value = LngDevID
'                RsDev("DEV_ID_Line_No").Value = lineno
'                If SystemOptions.SuppCreat4Acc Then
'                    RsDev("Account_Code").Value = DcDiscountAccount.BoundText   '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
'                Else
'                    RsDev("Account_Code").Value = Account_Code_dynamic1 '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
'                End If
'                RsDev("Value").Value = val(Me.txtDiscount.text)
'                RsDev("Credit_Or_Debit").Value = 1
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    RsDev("Double_Entry_Vouchers_Description").Value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & Chr(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & Chr(13) & TxtRemarks.text
'                Else
'                    RsDev("Double_Entry_Vouchers_Description").Value = "  Project Invoice No  :  " & TxtNoteSerial1 & Chr(13) & "  To Project " & txtprojectname.text & "  Manual# " & txtManualNo & Chr(13) & TxtRemarks.text
'                End If
'
'                RsDev("Notes_ID").Value = val(note_id.text)
'                RsDev("project_bill_no").Value = val(txtid.text)
'
'                ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'
'                RsDev("RecordDate").Value = XPDtbTrans.Value ' DateValue(Now)
'                RsDev("UserID").Value = user_id
'                RsDev("branch_id").Value = my_branch
'                RsDev("Account_Interval_ID").Value = SystemOptions.SysCurrentAccountIntervalID
'                RsDev("Posted").Value = IIf(Posted = 0, Null, Posted)
'                RsDev.update
'                'll:
'                lineno = lineno + 1
        
            Else
                RsDev.AddNew
            
                RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
        line_no = line_no + 1
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
                If SystemOptions.SuppCreat4Acc Then
                    RsDev("Account_Code").value = DcDiscountAccount.BoundText   '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
                Else
                    RsDev("Account_Code").value = Account_Code_dynamic1 '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
                End If
                RsDev("Value").value = val(Me.txtDiscount.text)
                RsDev("Credit_Or_Debit").value = 0
                
                RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
                RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  Manual# " & TxtManualNO & CHR(13) & txtRemarks.text
                
                RsDev("Notes_ID").value = val(note_id.text)
                RsDev("project_bill_no").value = val(txtid.text)
            
                 RsDev("project_id").value = val(Me.DataCombo2.BoundText)
           
                RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
                RsDev("UserID").value = user_id
                RsDev("branch_id").value = my_branch
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                RsDev.update
                'll:
                lineno = lineno + 1
                 line_no = line_no + 1
            End If
        End If

        If val(Me.TxtPerforValue.text) > 0 Then
            If Not SystemOptions.SuppCreat4Acc Then
                RsDev.AddNew
                RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
                If SystemOptions.SuppCreat4Acc Then
                    RsDev("Account_Code").value = accountdep ' get_account_code_branch(152, my_branch) 'Õ”«» Õ”‰ «·«œ«¡
                Else
                    RsDev("Account_Code").value = AcountGood ' get_account_code_branch(152, my_branch) 'Õ”«» Õ”‰ «·«œ«¡
                End If
                RsDev("Value").value = val(Me.TxtPerforValue.text)
                RsDev("Credit_Or_Debit").value = 0
                
                
                RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
                RsDev("Double_Entry_Vouchers_Descriptione").value = " Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & " Manual#  " & TxtManualNO & CHR(13) & txtRemarks.text
                RsDev("Notes_ID").value = val(note_id.text)
                RsDev("project_bill_no").value = val(txtid.text)
                 RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
                RsDev("UserID").value = user_id
                RsDev("branch_id").value = my_branch
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                RsDev.update
            End If
            lineno = lineno + 1
             line_no = line_no + 1
            If SystemOptions.SuppCreat4Acc Then
                RsDev.AddNew
                RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
         
                RsDev("Account_Code").value = GuranteeAccount ' get_account_code_branch(152, my_branch) 'Õ”«» Õ”‰ «·«œ«¡
        
                RsDev("Value").value = val(Me.TxtPerforValue.text)
                RsDev("Credit_Or_Debit").value = 0

                    RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
                RsDev("Double_Entry_Vouchers_Descriptione").value = " Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & " Manual#  " & TxtManualNO & CHR(13) & txtRemarks.text
                RsDev("Notes_ID").value = val(note_id.text)
                RsDev("project_bill_no").value = val(txtid.text)
                 RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
                RsDev("UserID").value = user_id
                RsDev("branch_id").value = my_branch
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                RsDev.update
            End If
    
            'll:
            lineno = lineno + 1
             line_no = line_no + 1

        End If
      If val(txtPerformanceBond) <> 0 Then
                des = "«” —Ã«⁄ ÷„«‰ «⁄„«· " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " ··„” Œ·’ —ﬁ„ " & TxtNoteSerial1.text & CHR(13) & txtRemarks.text
                desEN = " Performance Bond " & "   " & txtRemarks & "Inv# " & TxtNoteSerial1 & " manual#" & TxtManualNO & CHR(13) & txtRemarks.text

           
                If GuranteeAccount = "" Then
                    GuranteeAccount = accountdep
                End If
            
'                If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(txtPerformanceBond), 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
'                    GoTo ErrTrap
'                End If
                    
                line_no = line_no + 1
            
                If ModAccounts.AddNewDev(LngDevID, line_no, GuranteeAccount, val(txtPerformanceBond), 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                    
                line_no = line_no + 1
        End If
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        If val(Me.txtDiscountGMater.text) > 0 Then
            lineno = lineno + 1
            line_no = line_no + 1
     
       des = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text

            desEN = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  manual  " & TxtManualNO & CHR(13) & txtRemarks.text
     
       If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(txtDiscountGMater), 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                
                 'll:
            lineno = lineno + 1
     line_no = line_no + 1
     
            If ModAccounts.AddNewDev(LngDevID, line_no, mAccountMaterial, val(txtDiscountGMater), 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                
           
        End If

        '«·œ›⁄«  «·„ﬁœ„…
        'Account_Code_dynamic2
        If val(Me.advancedPayment.text) > 0 Then
            lineno = lineno + 1
             line_no = line_no + 1
            If SystemOptions.SuppCreat4Acc Then
            
            
             line_no = line_no + 1
                des = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                desEN = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  Manuall#  " & TxtManualNO & CHR(13) & txtRemarks.text
     
            If ModAccounts.AddNewDev(LngDevID, line_no, AdvancedAccount, val(advancedPayment), 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , , , , val(txtid.text)) = False Then
                    GoTo ErrTrap
                End If
                
            line_no = line_no + 1
           
            
                lineno = lineno + 1
        
'                RsDev.AddNew
'
'                RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
'
'                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'                RsDev("DEV_ID_Line_No").value = lineno
'                RsDev("Account_Code").value = accountdep  '    Õ”«» œ›⁄«  „ﬁœ„…
'                RsDev("Value").value = val(Me.advancedPayment.text)
'                RsDev("Credit_Or_Debit").value = 1
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo & "œ›⁄«  „ﬁœ„…" & CHR(13) & TxtRemarks.text
'                Else
'                    RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "   Manual   " & txtManualNo & CHR(13) & TxtRemarks.text
'                End If
'
'                RsDev("Notes_ID").value = val(note_id.text)
'                RsDev("project_bill_no").value = val(txtID.text)
'                RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'                RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'                RsDev("UserID").value = user_id
'                RsDev("branch_id").value = my_branch
'                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'                RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
'                RsDev.update
                'll:
                lineno = lineno + 1
                 line_no = line_no + 1
            Else
                RsDev.AddNew
            
                RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
         line_no = line_no + 1
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
                RsDev("Account_Code").value = Account_Code_dynamic2 '    Õ”«» œ›⁄«  „ﬁœ„…
                RsDev("Value").value = val(Me.advancedPayment.text)
                RsDev("Credit_Or_Debit").value = 0
        
                RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  Manuall#  " & TxtManualNO & CHR(13) & "Advance payments " & txtRemarks.text
                RsDev("Notes_ID").value = val(note_id.text)
                RsDev("project_bill_no").value = val(txtid.text)
                 RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
                RsDev("UserID").value = user_id
                RsDev("branch_id").value = my_branch
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                RsDev.update
                'll:
                lineno = lineno + 1
         line_no = line_no + 1
                RsDev.AddNew
            
                RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
        
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
                RsDev("Account_Code").value = accountdep '    Õ”«» œ›⁄«  „ﬁœ„…
                RsDev("Value").value = val(Me.advancedPayment.text)
                RsDev("Credit_Or_Debit").value = 1
        
                
                RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  Manuall#  " & TxtManualNO & CHR(13) & "Advance payments " & txtRemarks.text
                RsDev("Notes_ID").value = val(note_id.text)
                RsDev("project_bill_no").value = val(txtid.text)
                RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
                RsDev("UserID").value = user_id
                RsDev("branch_id").value = my_branch
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                RsDev.update
                'll:
                
                 line_no = line_no + 1
            End If
        End If

        '«·«Ì—œ« 

        '«·ÿ—› «·œ«∆‰
        If Option8.value = False Then
            If Me.txtrevenue_account.text = "" Then Exit Function
    
        Else
            If (accountdep = "") Then Exit Function
        End If
        RsDev.AddNew
        RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
        RsDev("branch_id").value = my_branch
        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
        RsDev("DEV_ID_Line_No").value = line_no
        'If SystemOptions.Revenueowed = True Then
        If Option8.value = True Then
            RsDev("Account_Code").value = accountdep
        Else
            ''
            RsDev("Account_Code").value = Me.txtrevenue_account.text ' Account_Code_dynamic1
 
        End If
        '   Else
        'RsDev("Account_Code").value = Me.txtrevenue_account .text
        '   End If
         If SystemOptions.SuppCreat4Acc Then
        RsDev("Value").value = val(Me.Results.text)   '«·«Ì—«œ« 
        Else
            RsDev("Value").value = val(Me.total.text) + val(advancedPayment.text)  ' val(Me.Results.text)  '«·«Ì—«œ« 
        End If
        RsDev("Credit_Or_Debit").value = 1


        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
        RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & CHR(13) & txtRemarks.text
        RsDev("Notes_ID").value = val(note_id.text)
        RsDev("project_bill_no").value = val(txtid.text)
        If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
              RsDev("project_id").value = val(Me.DataCombo2.BoundText)
        Else
            RsDev("project_id").value = val(Me.DataCombo2.BoundText)
        End If
 
        RsDev("RecordDate").value = XPDtbTrans.value
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
        RsDev.update
         line_no = line_no + 1
        '///////////////////
        If Me.AccountVat.BoundText <> "" And val(Me.TxtFATValue.text) > 0 Then
            RsDev.AddNew
            RsDev("Account_Code").value = Me.AccountVat.BoundText
            RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
            RsDev("branch_id").value = my_branch
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = line_no
           ' RsDev("Value").value = val(Me.TxtFATValue.text) - val(val(Me.TxtPreVAT.text)) '«·«Ì—«œ« 
            RsDev("Value").value = val(Me.TxtFATValue.text)  '«·«Ì—«œ« 
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & txtid & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " Õ”«» «·ﬁÌ„… «·„÷«›…" & CHR(13) & txtRemarks.text
            RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "Account VAT" & CHR(13) & " Value Added Tax (VAT) " & txtRemarks.text
            
            RsDev("Notes_ID").value = val(note_id.text)
            RsDev("project_bill_no").value = val(txtid.text)
            If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
                  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
            Else
                  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
            End If
 
            RsDev("RecordDate").value = XPDtbTrans.value
            RsDev("UserID").value = user_id
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
     
            RsDev.update
             line_no = line_no + 1
            ''///////////////////////›«  «·œ›⁄«  «·„ﬁœ„…

            If SystemOptions.SuppCreat4Acc = False Then
                If val(TxtPreVAT.text) > 0 Then
                    RsDev.AddNew
                    RsDev("Account_Code").value = Me.AccountVat.BoundText
                    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
                    RsDev("branch_id").value = my_branch
                    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                    RsDev("DEV_ID_Line_No").value = line_no
                    RsDev("Value").value = val(Me.TxtPreVAT.text)  '«·«Ì—«œ« 
                    RsDev("Credit_Or_Debit").value = 0
                    
    
                    RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " Õ”«» «·ﬁÌ„… «·„÷«›… œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                    RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "Account VAT" & CHR(13) & " Value Added Tax (VAT) advance payments " & CHR(13) & txtRemarks.text
                    
                    RsDev("Notes_ID").value = val(note_id.text)
                    RsDev("project_bill_no").value = val(txtid.text)
                    If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
                          RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                    Else
                           RsDev("project_id").value = val(Me.DataCombo2.BoundText) 'xxxxxxxxxxxxxx
                    End If
 
                    RsDev("RecordDate").value = XPDtbTrans.value
                    RsDev("UserID").value = user_id
                    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                    RsDev.update
                     line_no = line_no + 1
                    ''//////////////
                    RsDev.AddNew
                    If Option8.value = True Then
                        RsDev("Account_Code").value = Me.TxtAccountUnderImp.text
                    Else
                        RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
                    End If
                    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
                    RsDev("branch_id").value = my_branch
                    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                    RsDev("DEV_ID_Line_No").value = line_no
                    RsDev("Value").value = val(Me.TxtPreVAT.text)  '«·«Ì—«œ« 
                    RsDev("Credit_Or_Debit").value = 1
         
                    RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " Õ”«» «·ﬁÌ„… «·„÷«›… œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                    RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "Account VAT" & CHR(13) & " Value Added Tax (VAT) advance payments " & CHR(13) & txtRemarks.text


                    RsDev("Notes_ID").value = val(note_id.text)
                    RsDev("project_bill_no").value = val(txtid.text)
                    If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
                          RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                    Else
                           RsDev("project_id").value = val(Me.DataCombo2.BoundText) 'xxxxxxxxxxxxxx
                    End If
                    RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                    RsDev("RecordDate").value = XPDtbTrans.value
                    RsDev("UserID").value = user_id
                    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                    RsDev.update
                    lineno = lineno + 1

                End If
            End If
        End If
        ''/////////
    Else
        '
        'If SystemOptions.SubContactorHave3Account = True Then
        Dim Discount1  As Double
        Dim Discount2  As Double
        Dim Discount3  As Double
        Dim netvalue   As Double
        Dim TotalValue As Double
  
        
        If cboDiscount1.ListIndex = 0 Then
            Discount1 = 0
        ElseIf cboDiscount1.ListIndex = 1 Then
            Discount1 = val(txtDiscount1) * val(Me.TxtNetValue.text) / 100
        ElseIf cboDiscount1.ListIndex = 2 Then
            Discount1 = val(txtDiscount1)
        End If
        
        If cboDiscount2.ListIndex = 0 Then
            Discount2 = 0
        ElseIf cboDiscount2.ListIndex = 1 Then
            Discount2 = val(txtDiscount2) * val(TxtNetValue.text) / 100
        ElseIf cboDiscount2.ListIndex = 2 Then
            Discount2 = val(txtDiscount2)
        End If
                            
        Discount3 = val(txtDiscountGMater.text)
               
        If SystemOptions.SuppCreat4Acc Then
            AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_CodeHi1")
            GuranteeAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_CodeAss2")
            mAccountMaterial = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_CodeHi2")
                    
        Else
            AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code2")
            GuranteeAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code1")
            mAccountMaterial = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code1")
        End If
               
        accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code")
               
        line_no = line_no + 1
        If SystemOptions.AllowNoRoudProjectInvoices = True Then
            Discount1 = Round(Discount1, val(cCompanyInfo.NoRoudProjectInvoices))
            Discount2 = Round(Discount2, val(cCompanyInfo.NoRoudProjectInvoices))
            netvalue = Round(val(TxtNetValue.text) - Discount1 - Discount2, val(cCompanyInfo.NoRoudProjectInvoices))
            TotalValue = Round(val(TxtNetValue), val(cCompanyInfo.NoRoudProjectInvoices))
        Else
            Discount1 = Round(Discount1, 2)
            Discount2 = Round(Discount2, 2)
            netvalue = Round(val(TxtNetValue.text) - Discount1 - Discount2 - Discount3, 2)
            TotalValue = Round(val(TxtNetValue), Decimal_Places)
        End If
        If Option8.value = True Then
            des = "„’—Ê›«  «·„‘«—Ì⁄ " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
            desEN = "Advance payments invoices" & "   " & txtRemarks & " Project Invoice No" & TxtNoteSerial1 & " Manual#  " & TxtManualNO & CHR(13) & txtRemarks.text
        Else
            des = "„” Œ·’ «·„‘«—Ì⁄ " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & "Manual#" & TxtManualNO & CHR(13) & txtRemarks.text
            desEN = "Project Invoice " & "   " & txtRemarks & " Project Invoice No" & TxtNoteSerial1 & " Manual#  " & TxtManualNO & CHR(13) & txtRemarks.text
        End If
        
          







        If TotalValue > 0 Then '
                
            If SystemOptions.SuppCreat4Acc Then
                Dim mmexpanses_account As String
                mmexpanses_account = expanses_account
                For i = 1 To Fg_Journal.rows - 1
                    TotalValue = val(Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("NetExe")))
                    '   TotalValue = Round(val(TxtNetValue), Decimal_Places)
                    If TotalValue <> 0 Then
                        expanses_account = Trim(Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("AccountCode")))
                      
                        If ModAccounts.AddNewDev(LngDevID, line_no, expanses_account, TotalValue, 0, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                            GoTo ErrTrap
                        End If
                    
                        line_no = line_no + 1
                    End If
                    
                    If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, TotalValue, 1, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                        GoTo ErrTrap
                    End If
                    
                    line_no = line_no + 1
                    
                Next
                If expanses_account = "" Then expanses_account = mmexpanses_account
                    
            Else
                If Option8.value = True Then
                    If ModAccounts.AddNewDev(LngDevID, line_no, TxtAccountUnderImp.text, TotalValue, 0, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                        GoTo ErrTrap
                    End If
                    
                    line_no = line_no + 1
                Else
                    If ModAccounts.AddNewDev(LngDevID, line_no, expanses_account, TotalValue, 0, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                        GoTo ErrTrap
                    End If
                    
                    line_no = line_no + 1
                End If
                   
            End If
                   
                   
               If SystemOptions.SuppCreat4Acc = False Then
                    line_no = line_no + 1
               LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(TotalValue), 1, Msg & des & "  " & "    " & Me.DcbosubContractor.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "    " & Me.DcbosubContractor.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                    'desEN & "  " & "    " & txtprojectname.text
                line_no = line_no + 1
                End If
            If val(TxtFATValue.text) > 0 Then
               ' If ModAccounts.AddNewDev(LngDevID, line_no, Me.AccountVat.BoundText, val(TxtFATValue.text) - val(val(Me.TxtPreVAT.text)), 0, Msg & "  " & "    " & txtprojectname.text & "    VAT  " & "   INV# " & TxtNoteSerial1.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                If ModAccounts.AddNewDev(LngDevID, line_no, Me.AccountVat.BoundText, val(TxtFATValue.text), 0, Msg & "  " & "    " & txtprojectname.text & "    VAT  " & "   INV# " & TxtNoteSerial1.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , txtprojectname.text & "    VAT  " & "   INV# " & TxtNoteSerial1.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                    
                line_no = line_no + 1
            End If
            '  End If
            '////////////////////////////////////////////////////
            ' End If
         
            If SystemOptions.UserInterface = ArabicInterface Then
                
            Else
                
       
            End If
            des = "Œ’„ ÷„«‰ «⁄„«· " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " ··„” Œ·’ —ﬁ„ " & TxtNoteSerial1.text & CHR(13) & txtRemarks.text
            desEN = " Discount " & "   " & txtRemarks & "Inv# " & TxtNoteSerial1 & " manual#" & TxtManualNO & CHR(13) & txtRemarks.text
            If Discount1 > 0 Then '÷„«‰ «·«⁄„«·
    
                If GuranteeAccount = "" Then
                    GuranteeAccount = accountdep
                End If
            
                If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, Discount1, 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                    
                line_no = line_no + 1
            
                If ModAccounts.AddNewDev(LngDevID, line_no, GuranteeAccount, Discount1, 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                    
                line_no = line_no + 1
  
            End If
            
            
           If val(txtPerformanceBond) <> 0 Then
                des = "«” —Ã«⁄ ÷„«‰ «⁄„«· " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " ··„” Œ·’ —ﬁ„ " & TxtNoteSerial1.text & CHR(13) & txtRemarks.text
                desEN = " Performance Bond " & "   " & txtRemarks & "Inv# " & TxtNoteSerial1 & " manual#" & TxtManualNO & CHR(13) & txtRemarks.text

           
                If GuranteeAccount = "" Then
                    GuranteeAccount = accountdep
                End If
            
                If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(txtPerformanceBond), 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                    
                line_no = line_no + 1
            
                If ModAccounts.AddNewDev(LngDevID, line_no, GuranteeAccount, val(txtPerformanceBond), 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                    
                line_no = line_no + 1
        End If
            
            If SystemOptions.UserInterface = ArabicInterface Then
                
            Else
                
            End If
            des = "Œ’„ œ›⁄«  „ﬁœ„…   " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
            desEN = "Advance Discount " & "   " & txtRemarks & "InvE " & TxtNoteSerial1 & "Manuall#   " & TxtManualNO & CHR(13) & txtRemarks.text
            If Discount2 > 0 Then '
    
                If AdvancedAccount = "" Then
                    AdvancedAccount = accountdep
                End If
            
                '               If ModAccounts.AddNewDev(LngDevID, line_no, AdvancedAccount, Discount2, 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.Text, val(note_id.Text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , DataCombo2.BoundText, , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                '                                        GoTo ErrTrap
                '                                    End If
                '
                '                                    line_no = line_no + 1
  
            End If
         
            '«·œ›⁄«  «·„ﬁœ„…
            'Account_Code_dynamic2
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

            If val(Me.txtDiscountGMater.text) > 0 Then
                lineno = lineno + 1
                
                
                 des = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                desEN = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  manual  " & TxtManualNO & CHR(13) & txtRemarks.text
                 line_no = line_no + 1
            
                If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(txtDiscountGMater), 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                
                
            
                'll:
                lineno = lineno + 1
    
                 des = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                desEN = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  manual  " & TxtManualNO & CHR(13) & txtRemarks.text
    
                If ModAccounts.AddNewDev(LngDevID, line_no, mAccountMaterial, val(txtDiscountGMater), 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                
            
            End If
            If val(txtDiscountG) > 0 Then
            
             line_no = line_no + 1
    
                    des = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "Õ”„Ì«  " & CHR(13) & txtRemarks.text
                    desEN = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  manual  " & TxtManualNO & " Deductions " & CHR(13) & txtRemarks.text
    
                    If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(txtDiscountG), 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                    End If
    
                  
                    'll:
    
                    line_no = line_no + 1
                    des = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "Õ”„Ì«  " & CHR(13) & txtRemarks.text
                    desEN = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  manual  " & TxtManualNO & " Deductions " & CHR(13) & txtRemarks.text
    
                    If ModAccounts.AddNewDev(LngDevID, line_no, DcDiscountAccount.BoundText, val(txtDiscountG), 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                    End If
     
    
            End If
            If val(Me.advancedPayment.text) > 0 Then

                If SystemOptions.SuppCreat4Acc Then
    
                    line_no = line_no + 1
    
    
    
                des = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                desEN = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  manual  " & TxtManualNO & CHR(13) & txtRemarks.text
    
                If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(advancedPayment), 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , , , , val(txtid.text)) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
                 
                 If ModAccounts.AddNewDev(LngDevID, line_no, AdvancedAccount, val(advancedPayment), 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , , , , val(txtid.text)) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
    
                   
                    'll:
    
           
                Else
                
                 If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(advancedPayment), 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , , , , val(txtid.text)) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
                
                  
    
                      If ModAccounts.AddNewDev(LngDevID, line_no, AdvancedAccount, val(advancedPayment), 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "Project    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , , , , val(txtid.text)) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
                   
                    'll:
                  
                    'll:
                    line_no = line_no + 1

                    RsDev.AddNew

                    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

                    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                    RsDev("DEV_ID_Line_No").value = lineno
                    RsDev("Account_Code").value = Me.AccountVat.BoundText   '    Õ”«» œ›⁄«  „ﬁœ„…
                    RsDev("Value").value = val(Me.TxtPreVAT.text)
                    RsDev("Credit_Or_Debit").value = 1

                   
                                 
                    RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
                    RsDev("Double_Entry_Vouchers_Descriptione").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  manual  " & TxtManualNO & " Advance payments" & CHR(13) & txtRemarks.text

    
                    

                    RsDev("Notes_ID").value = val(note_id.text)
                    RsDev("project_bill_no").value = val(txtid.text)
                       RsDev("project_id").value = val(Me.DataCombo2.BoundText)
                    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
                    RsDev("UserID").value = user_id
                    RsDev("branch_id").value = my_branch
                    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
                    RsDev.update
                End If
            End If
 
       
            des = " «⁄„«·" & " ··„” Œ·’ —ﬁ„ " & TxtNoteSerial1.text & CHR(13) & txtRemarks.text
            desEN = " Works" & "  Inv#" & TxtNoteSerial1.text & CHR(13) & txtRemarks.text
            'If val(TxtFATValue.text) - val(TxtPreVAT.text) > 0 Then '
            If val(TxtFATValue.text) > 0 Then  '
    
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                'If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(TxtFATValue.text) - val(TxtPreVAT.text), 1, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, val(TxtFATValue.text), 1, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , desEN & "  " & "    " & txtprojectname.text, setfoxy_Line, , val(Me.DataCombo2.BoundText), , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                    
                line_no = line_no + 1
  
            End If

        End If
    End If

  
  
    Savetemp
    Cn.CommitTrans
    If bill_Type.ListIndex = 0 Then
        If Not chkTaxExempt.value = vbChecked And SystemOptions.ApplyEinvoice Then savenewelectroncic
    End If
If SystemOptions.IsBluee = True And bill_Type.ListIndex = 0 Then
 
   
                MsgBox SENDEINVOICE(Me.XPTxtBillID, True, val(Me.TXTEnd_user_id.text), 1, "project_billl", "ID"), vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  
End If
BeginTrans = False
    Exit Function
  
  
ErrTrap:

    If Err.Number = -2147217900 Then

         Msg = "You can not save data " & CHR(13)
        Msg = Msg + "It has been enter  incorrect data " & CHR(13)
        Msg = Msg + "Make sure of the validity of the data and try again"

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        Cn.RollbackTrans
        BeginTrans = False
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry...error during Saving"
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  
End Function
Private Sub UpdatePre_QuantityCont(ByVal mRow As Long)

    Dim s As String

    s = "Update SubcontractorContract2 Set  Pre_Quantity = Pre_Quantity + "
    s = s & val(Fg_Journal.TextMatrix(mRow, Fg_Journal.ColIndex("quntExc")))
    s = s & " Where project_id = " & val(Fg_Journal.TextMatrix(mRow, Fg_Journal.ColIndex("project_id")))
    s = s & " and oprid = " & val(Fg_Journal.TextMatrix(mRow, Fg_Journal.ColIndex("oprid")))
    s = s & " and bill_id =  " & val(TXTOrDer_no2)
    Cn.Execute s
End Sub
Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    Dim sql As String
    
    If Fg_Journal.rows > 1 Then
        If Fg_Journal.rows = 2 Then
            Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Fg_Journal.rows > 1 Then
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
    
    If (Index = 1 Or Index = 4) And zatcaStatus = 1 And TxtModFlg.text <> "N" Then
    ' If SystemOptions.IsBluee = True Then
           Msg = "·« Ì„ﬂ‰  ⁄œÌ· «Ê Õ–› «Ì „” ‰œ Ì„ﬂ‰ﬂ ⁄„· „” ‰œ ⁄ﬂ”Ì ›ﬁÿ"
                        Msg = Msg & CHR(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
   '     End If

End If


      If mZakamsg <> "" Then
            
        MsgBox mZakamsg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, " ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ «·„—Õ·… «·À«‰Ì… - „—Õ·… «·—»ÿ Ê«· ﬂ«„·"
    End If
    
    
    Select Case Index
        Case 12
            txtid.text = ""
            TxtModFlg.text = "N"

            Fg_Journal.rows = Fg_Journal.rows + 1
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

            TxtModFlg.text = "N"
            clear_all Me
            Accredit.Caption = ""

            Me.DCboUserName.BoundText = user_id
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            
            
            Fg_Journal.rows = 2
            Fg_Journal.Enabled = True
            
            GrdBondHistory.Clear flexClearScrollable, flexClearEverything
          GrdBondHistory.rows = GrdBondHistory.rows + 1
            Command1(1).Enabled = True
            XPDtbTrans.value = DateValue(Now)
            Results.text = 0
            XPDtbTrans.value = Date
            TxtNoteSerial.text = ""
            Me.dcBranch.BoundText = Current_branch
            cboDiscount1.ListIndex = 0
            cboDiscount1.ListIndex = 0
            cboDiscount2.ListIndex = 0
            billto.ListIndex = 0
            
             DefaultInvoicetype.ListIndex = SystemOptions.DefaultInvoicetype
            zatcaStatus = 0
            FlgAproved = 0
            txtDateRec.value = Date
        Case 1
            If bill_Type.ListIndex = 0 Then
                If checkCustomerdata(val(Me.TXTEnd_user_id.text), val(TxtTotalValue2), val(DefaultInvoicetype.ListIndex), Dccurrency.text, Export) = False Then Exit Sub
            End If
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
            If val(txtDiscount.text) > 0 Then
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
    
            If val(TxtPerforValue.text) > 0 Then
        
                If AcountGood = "" Then
                    MsgBox "·„ Ì „ «‰‘«¡ Õ”«» Õ”‰ «·«œ«¡ ·Â–« «·„‘—Ê⁄", vbCritical
        
                    'create_accounts = False
                    Exit Sub

                End If
        
            End If
    
            '«· «ﬂœ „‰ «·œ›⁄«  «·„ﬁœ„…
            If val(advancedPayment.text) > 0 Then
  
                If SystemOptions.CustomerhavethreeAccounts = True Then
                    Account_Code_dynamic2 = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.TXTEnd_user_id.text), "Account_code2")
         
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
                ElseIf SystemOptions.SuppCreat4Acc Then
         
                    Account_Code_dynamic2 = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.TXTEnd_user_id.text), "Account_CodeHi2")
                
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
    
            If val(total.text) <= 0 Then MsgBox "Õœœ  ﬂ·›… «·„„‰›– «Ê·«", vbCritical: Exit Sub
            If Not IsNumeric(dcBranch.BoundText) Then MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical: Exit Sub
            my_branch = val(dcBranch.BoundText)
    
            If TxtNoteSerial.text = "" Then
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
            If val(TxtBillNo.text) > 0 Then
                If val(txtPeriod.text) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ «œŒ«· „œ… «·› —… »Ì‰ «·›Ê« Ì—"
                    Else
                        MsgBox "Please enter Period"
                    End If
                    txtPeriod.SetFocus
                    Exit Sub
                End If
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
            If TxtNoteSerial1.text = "" Then
                'TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 65, 65, , , , , val(billto.ListIndex))
                
                If billto.ListIndex = 0 Then
                    TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 65, 65, , , , , val(billto.ListIndex))
                Else
                    TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 84, 84, , , , , val(billto.ListIndex))
                End If
    
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
                GrdBondHistory.rows = GrdBondHistory.rows + 1
            Dim AccountVATDept As String
            Dim str            As String
            str = "30/05/2017"
            If AccountVat.BoundText = "" And True = True And CheckAnyVAT(XPDtbTrans.value) = True And StartDateProje.value > CDate(str) Then
                MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
                Exit Sub
            End If
                      If Dccurrency.BoundText = "" Then
    
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "«Œ — «·⁄„·… «Ê·« "
    
            End If
            Command1(2).Enabled = True
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Dccurrency.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            '  Cmd(2).Enabled = True
            Exit Sub
            
        End If

    If billto.ListIndex = -1 Then MsgBox "Õœœ «·„” Œ·’ „ﬁœ„  «·Ï „‰ ", vbCritical: Exit Sub
    
    
    
    
    
    If billto.ListIndex = 1 And DcbosubContractor.BoundText = "" Then MsgBox "Õœœ „ﬁ«Ê· «·»«ÿ‰  ", vbCritical: Exit Sub
    If Trim(DcDiscountAccount.text) = "" And val(txtDiscount) <> 0 Then
        MsgBox "Õœœ Õ”«» «·Õ”„Ì«    ", vbCritical: Exit Sub
    End If
    
    
      
    Dim found As Boolean
     
    j = Fg_Journal.FixedRows
    found = False
     
    For j = Fg_Journal.FixedRows To Fg_Journal.rows - 1
        If Fg_Journal.TextMatrix(j, Fg_Journal.ColIndex("item")) <> "" Then
            found = True
        End If
    Next


    If found = False Then

        MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„ ›Ï «·›« Ê—… ", vbCritical: Exit Sub

    End If

    If SystemOptions.SuppCreat4Acc And val(TXTOrDer_no2) <> 0 Then
        j = Fg_Journal.FixedRows
        found = False
         
        For j = Fg_Journal.FixedRows To Fg_Journal.rows - 1
            If Fg_Journal.TextMatrix(j, Fg_Journal.ColIndex("AccountCode")) <> "" Then
                found = True
            End If
        Next
    
        If found = False Then
    
            MsgBox "·«»œ „‰ «œŒ«· «·Õ”«» ›Ï «·ÃœÊ· ", vbCritical: Exit Sub
            Exit Sub
        End If
     
    End If
            If SystemOptions.SuppCreat4Acc Then
                SaveData
            Else
                SaveDataOld
            End If
            
            ''Adodc1.Recordset.Fields!  project_no = DataCombo2.text
        Case 11

            On Error Resume Next

            ShowAttachments txtid.text, "24122020001"

            Exit Sub

            If SystemOptions.UserInterface = EnglishInterface Then
                If txtid.text = "" Then MsgBox "Select Bill firstly": Exit Sub

            Else

                If txtid.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „” Œ·’ «Ê·«": Exit Sub

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

            imaged.SUBJECT_NO = txtid.text
            imaged.txtopeation_type = "„—›ﬁ«  „” Œ·’"

            imaged.Adodc1.CommandType = adCmdText
            imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  „” Œ·’' and subject_no='" & txtid.text & "'"
            imaged.Adodc1.Refresh

            If imaged.Adodc1.Recordset.RecordCount > 0 Then

                imaged.DBPix201.Visible = True
            Else
                imaged.DBPix201.Visible = False
            End If

        Case 3
            Frame15.Enabled = True
            Frame15.Visible = False
            If ScreenAproved(val(txtid.text), Me.Name) = True Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
'                Else
'                    MsgBox "Can not edit.This process associated with approvals"
'                End If
'                Exit Sub
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

            'Dim Msg    As String
            Dim StrSQL As String
 
            Dim RsTemp As New ADODB.Recordset
            StrSQL = "select * From ProjectBillBuy where Bill_id=" & val(txtid.text)
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                Msg = "·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  Â–« «·›« Ê—… " & CHR(13)
                Msg = Msg + "·«‰Â«  „ ⁄·ÌÂ« ⁄„·Ì«  ”œ«œ"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
          
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            Fg_Journal.rows = Fg_Journal.rows + 1
            Fg_Journal.Enabled = True
            Command1(1).Enabled = True

        Case 4

        Case 5

        Case 6
            Undo

        Case 9
            If ScreenAproved(val(txtid.text), Me.Name) = True Then
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
           ' Savetemp
            print_report val(DataCombo2.BoundText)
        Case 18
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.text, , 200, , , , "rpt_Projects2.rpt", DataCombo2.text


        Case 8

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.text, , 200
            
        Case 10
    
            projectsbill_Search.show
        Case 15
            If Me.TxtModFlg.text = "N" Then
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
Function calcnetOld()
TxtNetValue.text = Round(val(Results.text) - val(txtDiscountG), Decimal_Places) ' - val(advancedPayment.text)
If val(cboDiscount1.ListIndex) = 1 Then
TxtPerforValue.text = Round(val(txtDiscount1.text) * val(TxtNetValue.text) / 100, Decimal_Places)
ElseIf val(cboDiscount1.ListIndex) = 2 Then
TxtPerforValue.text = Round(val(txtDiscount1.text), Decimal_Places)
Else
TxtPerforValue.text = 0
End If
total.text = Round(val(TxtNetValue.text) - val(TxtPerforValue.text), Decimal_Places)
Calculte
End Function

Sub SaveBillMonthly()

    rs("TotalBefore").value = val(Me.txtTotalBefore.text)
    rs("Discount4").value = val(Me.txtDiscount4.text)
    rs("Discount3").value = val(Me.txtDiscount3.text)
    
    rs("BondAmt").value = IIf(Not IsNumeric(Me.txtBondAmt.text), 0, Me.txtBondAmt.text)
    
    
    If val(Me.TxtBillNo.text) = 0 Then
        Exit Sub
    
    End If
    Dim RsDevsub As ADODB.Recordset
    Dim StrSQL   As String
    Dim i        As Integer
    StrSQL = "Delete from project_billl_Month where Bill_ID =" & val(Me.txtid.text) & ""
    Cn.Execute StrSQL
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from project_billl_Month Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    For i = 1 To val(Me.TxtBillNo.text)
        If i = 1 Then
            DateTemp.value = StartDate.value
        Else
            If val(DcbPeriodType.ListIndex) = 0 Then
                DateTemp.value = DateAdd("D", val(Me.txtPeriod.text), DateTemp.value)
            ElseIf val(DcbPeriodType.ListIndex) = 1 Then
                DateTemp.value = DateAdd("M", val(Me.txtPeriod.text), DateTemp.value)
            ElseIf val(DcbPeriodType.ListIndex) = 2 Then
                DateTemp.value = DateAdd("YYYY", val(Me.txtPeriod.text), DateTemp.value)
            End If
        End If
        RsDevsub.AddNew
        RsDevsub("Bill_ID").value = val(txtid.text)
        RsDevsub("RecordDate").value = DateTemp.value
        RsDevsub.update
    Next i
End Sub
Function print_report(Optional NoteSerial As Integer, Optional ByVal mType As Integer = 0)
    
    On Error Resume Next
    Dim MySQL          As String
    Dim RsData         As New ADODB.Recordset
    Dim xApp           As New CRAXDRT.Application
    Dim xReport        As CRAXDRT.Report
    Dim CViewer        As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName    As String
    Dim Msg            As String
 
    'new
    



    

   
    MySQL = "SELECT         project_billl.QrCodeImage, project_billl.PerformanceBond, project_billl.PerforValue, dbo.project_billl.id, dbo.project_billl.bill_date, dbo.project_billl.ManualNO, dbo.project_billl.duedate1, dbo.project_billl.discount, dbo.project_billl.dueDate, dbo.project_billl.NoteSerial, dbo.project_billl.total, "
    MySQL = MySQL & "                           dbo.project_billl.Remarks, dbo.project_billl.Results, dbo.project_billl.advancedPayment, dbo.project_billl.discount2value, dbo.project_billl.discount1value, dbo.project_billl.bill_type, dbo.project_billl.project_no,"
    MySQL = MySQL & "                                 dbo.projects.Fullcode, dbo.projects.project_name, dbo.project_billl.End_user_name, dbo.project_billl.Sub_user_name, dbo.project_billl.End_user_account, dbo.project_billl.bill_to, dbo.project_billl.Sub_user_account,"
    MySQL = MySQL & "                                  project_billl.TotalBefore,project_billl.Discount4,project_billl.Discount3,project_billl.BondAmt,"
    MySQL = MySQL & "                                 dbo.project_billl.revenue_account, dbo.project_billl.subContractorId, dbo.TblCustemers.Address, dbo.TblCustemers.VATNO CustemersVATNO, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.project_billl.Branch_NO,"
    MySQL = MySQL & "                                 dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.project_billl.discount1ID, dbo.project_billl.discount2ID, dbo.project_billl.note_id, dbo.project_bill_details.project_no AS project_noDet,"
    MySQL = MySQL & "                                 dbo.project_bill_details.item, dbo.project_bill_details.cost, dbo.project_bill_details.exe, dbo.project_bill_details.percentage * 100 AS percentage, dbo.project_bill_details.exedate, dbo.project_bill_details.bill_id,"
    MySQL = MySQL & "                                 dbo.project_bill_details.line_no, dbo.project_bill_details.item_id, dbo.project_bill_details.Quantity, dbo.project_bill_details.Price, dbo.project_bill_details.Pre_Quantity, dbo.project_bill_details.Pre_Value,"
    MySQL = MySQL & "                                 dbo.project_bill_details.Pre_Percent * 100 AS Pre_Percent, dbo.project_bill_details.Curr_Quantity, dbo.project_bill_details.Curr_value, dbo.project_bill_details.curr_Percent * 100 AS curr_Percent,"
    MySQL = MySQL & "                                 dbo.project_bill_details.tot_quantity, dbo.project_bill_details.tot_value, dbo.project_bill_details.tot_percent * 100 AS tot_percent, dbo.project_bill_details.Unit_id, dbo.TblProcessUnites.UnitName,"
    MySQL = MySQL & "                                 dbo.TblProcessUnites.UnitNamee, dbo.project_bill_details.oprid, dbo.project_bill_details.totEx, dbo.project_bill_details.quntExc, dbo.project_bill_details.net, dbo.project_bill_details.discount AS discountDet,"
    MySQL = MySQL & "                                 dbo.project_bill_details.total AS totalDet, dbo.project_bill_details.qty, dbo.project_bill_details.item_unit, dbo.project_bill_details.discountEXE, dbo.project_bill_details.NetExe,"
    MySQL = MySQL & "                                 dbo.project_bill_details.percentage1 * 100 AS percentage1, dbo.project_bill_details.Pre_Percent1 * 100 AS Pre_Percent1, dbo.project_bill_details.tot_percent1, dbo.project_bill_details.percentage1 AS Expr1,"
    MySQL = MySQL & "                                 dbo.project_bill_details.Pre_Percent1 AS Expr2, dbo.project_bill_details.QtyApprov, dbo.project_bill_details.PriceApprov, dbo.project_bill_details.TotalApprov, dbo.project_bill_details.DiscApprov,"
    MySQL = MySQL & "                                 dbo.project_bill_details.NetApprov, dbo.project_billl.NoteSerial1, dbo.project_billl.FATYou, dbo.project_billl.FATValue, dbo.project_billl.TotalValue, dbo.project_billl.ExPercen, dbo.project_billl.ExPercenID,"
    MySQL = MySQL & "                                 dbo.project_bill_details.ExPercen AS ExPercenDet, dbo.project_billl.PreVAT, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.projects.REVENUE_account_balance, dbo.projects.Project_nameE,Cuss.CusName subContractorName,Cuss.CusNamee subContractorNamee,TblCustemers.Address,TblCustemers.Cus_Phone,TblCustemers.Cus_mobile,TblCustemers.E_mail,TblCustemers.CustGID as CustGID,"
    
    MySQL = MySQL & "                                 Cuss.Address cussAddress  ,cuss.VATNO as VATNO2,cuss.CustGID as CustGID2,  Cuss.AddressE cussAddressE2, Cuss.Cus_Phone  CussCus_Phone,Cuss.Cus_mobile CussCus_mobile,Cuss.E_mail CussE_mail,TblCustemers.AddressE cussAddressE,"
    
    MySQL = MySQL & "                                 TblBranchesData.Commonname,TblBranchesData.OrganizationName,TblBranchesData.industrey,"
    MySQL = MySQL & "                                           TblBranchesData.CityName,TblBranchesData.CitySubdivisionName,TblBranchesData.CountrySubentity,"
    MySQL = MySQL & "                                           TblBranchesData.PostalZone,TblBranchesData.PlotIdentification,TblBranchesData.BuildingNumber,TblBranchesData.AdditionalStreetName,"
    MySQL = MySQL & "                                           TblBranchesData.StreetName BranchesDataStreetName , TblBranchesData.VATRegNo BranchesDataVATRegNo, TblBranchesData.VATNO BranchesDataVATNO, TblBranchesData.Company_Comment"
    MySQL = MySQL & "        FROM            dbo.TblCustemers INNER JOIN"
    MySQL = MySQL & "                                 dbo.projects ON dbo.TblCustemers.CusID = dbo.projects.End_user_id RIGHT OUTER JOIN"
    MySQL = MySQL & "                                dbo.TblProcessUnites RIGHT OUTER JOIN"
    MySQL = MySQL & "                                 dbo.project_bill_details ON dbo.TblProcessUnites.UnitID = dbo.project_bill_details.Unit_id RIGHT OUTER JOIN"
    MySQL = MySQL & "                                 dbo.project_billl ON dbo.project_bill_details.bill_id = dbo.project_billl.id LEFT OUTER JOIN"
    MySQL = MySQL & "                                 dbo.TblBranchesData ON dbo.project_billl.Branch_NO = dbo.TblBranchesData.branch_id ON dbo.projects.id = dbo.project_billl.project_no"
    MySQL = MySQL & "                                 Left outer join TblCustemers Cuss  ON dbo.project_billl.subContractorId = Cuss.CusID "
    
    MySQL = MySQL & " inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "

    MySQL = MySQL & "  Where dbo.project_billl.id  = " & val(txtid.text)
    MySQL = MySQL + " order by project_bill_details.id"

    
    If mType = 1 Then
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Projects2.rpt"
    Else
        
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    xReport.ParameterFields(3).AddCurrentValue WriteNo(Format(Me.TxtTotalValue.text, "0.00"), 0, True, ".", , 1, , , 1)

    'xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.VATRegNo
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    ''//////
    Dim xLogo  As CRAXDRT.OLEObject
    Dim SqlT   As String
    Dim i      As Integer
    Dim EmpIDD As Long
    Dim xWidth As Integer
    Dim Rs4    As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset
    SqlT = " SELECT        TOP (100) PERCENT dbo.TblUsers.Empid"
    SqlT = SqlT + "    FROM            dbo.ApprovalData INNER JOIN"
    SqlT = SqlT + "                      dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
    SqlT = SqlT + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.txtid.text) & ") AND (NOT (ApprovDate IS NULL)) AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
    SqlT = SqlT & " ORDER BY levelorder"
    Rs4.Open SqlT, Cn, adOpenStatic, adLockOptimistic, adCmdText
    xWidth = 300
    For i = 1 To Rs4.RecordCount
        EmpIDD = IIf(IsNull(Rs4("Empid").value), 0, Rs4("Empid").value)
        If Dir(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG") <> "" Then
        
            Set xLogo = xReport.Areas(4).Sections(1).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG", xWidth, 600)
            xLogo.Width = 800
            xLogo.Height = 400
            xLogo.backcolor = vbWhite
            xLogo.BorderColor = 255
            xLogo.CloseAtPageBreak = True
            xWidth = xWidth + 1000
        End If
        Rs4.MoveNext
    Next i
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_report2(Optional NoteSerial As Integer)
    
    On Error Resume Next
    Dim MySQL            As String
    Dim RsData           As New ADODB.Recordset
    Dim xApp             As New CRAXDRT.Application
    Dim xReport          As CRAXDRT.Report
    Dim CViewer          As ClsReportViewer
    Dim StrReportTitle   As String
    Dim StrFileName      As String
    Dim Msg              As String
 
    Dim currentProjectid As Integer
    Dim billlID          As Integer
    billlID = val(txtid.text)
    currentProjectid = val(DataCombo2.BoundText)
    subContractorId = val(DcbosubContractor.BoundText)

    'newwwwwwwwwwwwwwwww

    MySQL = " SELECT OLDTotalwithVat, CurrenttotalWithvat, Totalwitvat, oldPerforValue, totalPerforValue ,        ROUND(dbo.project_bill_details.totEx * (1 + dbo.project_billl.FATYou / 100), 3) AS CurrenttotalWithVatSalim, dbo.project_billl.project_no, dbo.projects.Fullcode, dbo.project_billl.project_name, dbo.TblProcessUnites.UnitName,"
    MySQL = MySQL & "                           dbo.TblProcessUnites.UnitNamee, dbo.projects.id AS ProjectID, dbo.GetPaymentValue(dbo.projects.id, dbo.project_billl.bill_date) AS TotalPayment, dbo.project_bill_details.item, dbo.project_bill_details.exe,"
    MySQL = MySQL & "                           dbo.project_billl.discount AS Sumdiscount, dbo.project_bill_details.discountEXE AS SumdiscountEXE, dbo.project_billl.advancedPayment AS SumadvancedPayment, dbo.project_bill_details.oprid, dbo.project_billl.FATYou,"
    MySQL = MySQL & "                           dbo.project_billl.FATValue, dbo.project_billl.NoteSerial1, dbo.project_bill_details.cost, dbo.project_billl.PreVAT, dbo.project_billl.PreBalaValue, dbo.project_billl.PreBalaVAT, dbo.project_billl.PreBalaTotal,"
    MySQL = MySQL & "                           dbo.project_billl.PreBalaPayed, dbo.project_billl.PreBalaRemain, dbo.project_billl.PreBalaTransPyed, dbo.project_billl.PreBalaNet, dbo.project_billl.SumVATLine, dbo.project_billl.PreBalaVATYu, dbo.project_billl.SumValueLine,"
    MySQL = MySQL & "                           dbo.project_billl.StartDateProje, dbo.project_billl.NetValue, dbo.project_billl.PerforValue, dbo.project_billl.PostedDate, dbo.project_billl.Posted, dbo.project_bill_details.Curr_Quantity, dbo.project_bill_details.Curr_value,"
    MySQL = MySQL & "                           dbo.project_bill_details.tot_quantity, dbo.projects.End_user_name, dbo.projects.sub_contractor_Account, dbo.project_billl.bill_date, dbo.project_billl.NoteSerial, dbo.project_billl.Branch_NO, dbo.TblBranchesData.branch_name,"
    MySQL = MySQL & "                           dbo.TblBranchesData.branch_namee, dbo.project_billl.Remarks, dbo.project_billl.ManualNO, dbo.project_billl.UserID, dbo.project_billl.BillNo, dbo.TblUsers.UserName, dbo.project_billl.StartDate, dbo.project_billl.Period,"
    MySQL = MySQL & "                           dbo.project_billl.Remarks2, dbo.project_billl.subContractorId, dbo.TblCustemers.CusName AS Subcontractorname, dbo.TblCustemers.CusNamee AS Subcontractornamee, dbo.TblCustemers.CusID,"
    MySQL = MySQL & "                           ROUND(dbo.GetTotEx(dbo.project_bill_details.oprid, 9), 2) AS OldtotlEXEValueSalim, ROUND(dbo.GetQuntExc(dbo.project_bill_details.oprid, 9), 2) AS oldtotQtysalim, ROUND(dbo.GetOLDPerforValue(18, 9), 2)"
    MySQL = MySQL & "                           AS tOLDPerforValue, dbo.project_bill_details.qty, ROUND(dbo.GetOLDPerforValuebysubContractorId(18, 9, 328), 2) AS OLDPerforValuebyCONTRACTORiD, dbo.project_bill_details.quntExc, dbo.project_bill_details.totEx,"
    MySQL = MySQL & "                           dbo.project_bill_details.LineDiscountPercent, dbo.project_bill_details.LineDiscount, dbo.project_bill_details.linenetaftermainDiscount, dbo.project_bill_details.linenetaftermainDiscountBeforevat, dbo.project_bill_details.LineVat,"
    MySQL = MySQL & "                           dbo.project_bill_details.linenetaftermainDiscountWithvat, dbo.project_bill_details.PerforVLineDiscount, dbo.project_bill_details.LineFinal, dbo.project_bill_details.qtySubContractor, dbo.project_bill_details.costSubContractor,"
    MySQL = MySQL & "                           dbo.project_bill_details.percentage1, dbo.project_bill_details.Pre_Percent1, dbo.project_bill_details.tot_percent1, dbo.project_bill_details.Pre_Quantity, dbo.project_bill_details.Pre_Value, dbo.project_bill_details.Pre_Percent,"
    MySQL = MySQL & "                           dbo.project_bill_details.curr_Percent, dbo.project_bill_details.percentage, dbo.project_bill_details.exedate, dbo.project_bill_details.Quantity, dbo.project_bill_details.Price, dbo.project_bill_details.tot_value,"
    MySQL = MySQL & "                           dbo.project_bill_details.tot_percent, dbo.project_bill_details.total, dbo.project_bill_details.discount, dbo.project_bill_details.net, dbo.project_bill_details.QtyApprov, dbo.project_bill_details.TotalApprov,"
    MySQL = MySQL & "                           dbo.project_bill_details.PriceApprov , dbo.project_bill_details.DiscApprov, dbo.project_bill_details.NetApprov, dbo.project_bill_details.PrMainDesID"
    MySQL = MySQL & "   FROM            dbo.TblBranchesData RIGHT OUTER JOIN"
    MySQL = MySQL & "                           dbo.project_billl ON dbo.TblBranchesData.branch_id = dbo.project_billl.Branch_NO LEFT OUTER JOIN"
    MySQL = MySQL & "                           dbo.TblUsers ON dbo.project_billl.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
    MySQL = MySQL & "                           dbo.TblCustemers ON dbo.project_billl.subContractorId = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                           dbo.TblProcessUnites RIGHT OUTER JOIN"
    MySQL = MySQL & "                           dbo.project_bill_details ON dbo.TblProcessUnites.UnitID = dbo.project_bill_details.Unit_id ON dbo.project_billl.id = dbo.project_bill_details.bill_id LEFT OUTER JOIN"
    MySQL = MySQL & "                           dbo.projects ON dbo.project_billl.project_no = dbo.projects.id"
    MySQL = MySQL & "  Where (dbo.project_billl.ID = " & billlID & ") And (dbo.Projects.ID = " & currentProjectid & ")     ORDER BY project_bill_details.ID"

    If billto.ListIndex = 0 Then
        ' MySQL = MySQL + " order by project_bill_details.id"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    ''//////
    '  Dim xLogo As CRAXDRT.OLEObject
    '  Dim SqlT As String
    '  Dim i As Integer
    '  Dim EmpIDD As Long
    '  Dim xWidth As Integer
    '  Dim Rs4 As ADODB.Recordset
    '  Set Rs4 = New ADODB.Recordset
    ' SqlT = " SELECT        TOP (100) PERCENT dbo.TblUsers.Empid"
    ' SqlT = SqlT + "    FROM            dbo.ApprovalData INNER JOIN"
    ' SqlT = SqlT + "                      dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
    ' SqlT = SqlT + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.txtid.Text) & ") AND (NOT (ApprovDate IS NULL)) AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
    ' SqlT = SqlT & " ORDER BY levelorder"
    ' Rs4.Open SqlT, Cn, adOpenStatic, adLockOptimistic, adCmdText
    ' xWidth = 300
    ' For i = 1 To Rs4.RecordCount
    ' EmpIDD = IIf(IsNull(Rs4("Empid").value), 0, Rs4("Empid").value)
    '           If Dir(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG") <> "" Then
    '
    '
    '
    '           Set xLogo = xReport.Areas(5).Sections(2).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG", xWidth, 300)
    '           xLogo.Width = 800
    '           xLogo.Height = 800
    '           xLogo.backcolor = vbWhite
    '           xLogo.BorderColor = 255
    '           xLogo.CloseAtPageBreak = True
    '          xWidth = xWidth + 1000
    '         End If
    '   Next i
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub Del_Trans()
    Dim Msg    As String
    Dim StrSQL As String
    'On Error GoTo ErrTrap
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If Me.txtid.text <> "" Then
        StrSQL = "select * From ProjectBillBuy where Bill_id=" & val(txtid.text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·›« Ê—… " & CHR(13)
            Msg = Msg + "·«‰Â«  „ ⁄·ÌÂ« ⁄„·Ì«  ”œ«œ"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtid.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            StrSQL = "Delete  Notes  where NoteSerial ='" & TxtNoteSerial & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete from project_billl_Month where Bill_ID =" & val(Me.txtid.text) & ""
            Cn.Execute StrSQL
            StrSQL = "Delete From TblPayPrePayed Where NoteID1=" & val(Me.txtid.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblProjePayPrePayed Where   NoteID=" & val(Me.txtid.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            DeleteBillBuy
            VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid4.rows = 1
            If Not rs.RecordCount < 1 Then
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.rows = 3
                    Fg_Journal.Enabled = False
           
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
     
        Exit Sub
    End If
 
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Function GET_PROJECT_DATA(Optional IDx As Integer = 0)
    On Error Resume Next

    If DataCombo2.text = "" Then Exit Function
    Dim My_SQL As String

    My_SQL = "select * from projects where id =" & DataCombo2.BoundText
 
    Set Rec = New ADODB.Recordset
    Rec.CursorLocation = adUseClient

    Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If SystemOptions.UserInterface = ArabicInterface Then
        txtprojectname.text = Rec.Fields("Project_name").value
    Else
        txtprojectname.text = Rec.Fields("Project_nameE").value
    End If
    
    If Trim(Rec.Fields("Project_nameE").value & "") = "" Then
        txtprojectname.text = Rec.Fields("Project_name").value
    Else
        txtprojectname.text = Rec.Fields("Project_nameE").value
    End If
    txtsubaccount.text = IIf(IsNull(Rec.Fields("sub_contractor_Account").value), "", Rec.Fields("sub_contractor_Account").value)
    TxtAccountUnderImp.text = IIf(IsNull(Rec.Fields("AccountUnderImp").value), "", Rec.Fields("AccountUnderImp").value)

    DcAccount1.text = IIf(IsNull(Rec.Fields("sub_contractor_name").value), "", Rec.Fields("sub_contractor_name").value)
    txtendaccount.text = IIf(IsNull(Rec.Fields("End_user_Account").value), "", Rec.Fields("End_user_Account").value)
    DcAccount2.text = IIf(IsNull(Rec.Fields("End_user_name").value), "", Rec.Fields("End_user_name").value)
    Dim End_user_id       As Double
    Dim sub_contractor_id As Double
    If Me.TxtModFlg = "N" Then
        StartDateProje.value = IIf(IsNull(Rec.Fields("StartDate").value), Date, Rec.Fields("StartDate").value)
        StartDateProje_Change
    End If
    End_user_id = IIf(IsNull(Rec.Fields("End_user_id").value), 0, Rec.Fields("End_user_id").value)
    sub_contractor_id = IIf(IsNull(Rec.Fields("sub_contractor_id").value), 0, Rec.Fields("sub_contractor_id").value)
    DcAccount2.text = GET_ACCOUNT_name_by_Code(get_Customer_Account(End_user_id))
    DcAccount1.text = GET_ACCOUNT_name_by_Code(get_Customer_Account(sub_contractor_id))
    If SystemOptions.Revenueowed = True Then
        txtrevenue_account.text = IIf(IsNull(Rec.Fields("legal").value), "", Rec.Fields("legal").value) 'Õ”«» «·„” Œ·’« \
    Else
        txtrevenue_account.text = IIf(IsNull(Rec.Fields("REVENUE_account").value), "", Rec.Fields("REVENUE_account").value) 'Õ”«» «·«Ì—«œ« \

    End If
  
    TXTEnd_user_id.text = IIf(IsNull(Rec.Fields("End_user_id").value), "", Rec.Fields("End_user_id").value) '—ﬁ„ «·⁄„Ì· «·‰Â«∆Ì
    TXTsub_contractor_id.text = IIf(IsNull(Rec.Fields("sub_contractor_id").value), "", Rec.Fields("sub_contractor_id").value) '—ﬁ„   „ﬁ«Ê· «·»«ÿ‰

    expanses_account = IIf(IsNull(Rec.Fields("expanses_account").value), "", Rec.Fields("expanses_account").value) 'Õ”«»  «·„’—Ê›« \
    If val(Rec!UnderImp & "") = 2 Then
        AcountGood = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code2")
        'AcountGood = IIf(IsNull(Rec.Fields("AcountGood").value), "", Rec.Fields("AcountGood").value)
    Else
        AcountGood = IIf(IsNull(Rec.Fields("AcountGood").value), "", Rec.Fields("AcountGood").value)
    End If
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
 
        If CBoBasedON.ListIndex <> 1 And val(TXTOrDer_no2.text) <> 0 Then
 
            ReloadContrac (val(DataCombo2.BoundText))
        Else
            Dim Dcombos As ClsDataCombos
        
            Set Dcombos = New ClsDataCombos
     
            Dcombos.GetPersons Me.DcbosubContractor
        End If
    End If
    'My_SQL = "  select net,des from projects_des  where project_id='" & DataCombo2.BoundText & "'"
    'fill_combo DataCombo5, My_SQL

End Function

Private Sub Command10_Click()
    Dim i      As Integer
    Dim StrSQL As String
    If Me.TxtModFlg.text = "E" Then
        DeleteBillBuy
        VSFlexGrid4.Enabled = True
        Check1.Enabled = True
        StrSQL = "Delete From TblPayPrePayed Where NoteID1=" & val(Me.txtid.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblProjePayPrePayed Where   NoteID=" & val(Me.txtid.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
        VSFlexGrid4.rows = 1

        FlgBillBuy = True
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «·€«¡ «·”œ«œ"
        Else
            MsgBox "Done"
        End If
        ALLButton1_Click
        With Me.VSFlexGrid4

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With
    End If
End Sub

Private Sub DataCombo2_Change()
    GET_PROJECT_DATA 1
   
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    GET_PROJECT_DATA 1
    ' ReloadContrac (val(DataCombo2.BoundText))
End Sub

Private Sub DataCombo5_Click(Area As Integer)

    If DataCombo5.BoundText <> "" Then
        Text6.text = DataCombo5.BoundText
        Text9.text = ""
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
            total.text = Round(gettotal(txtid.text), Decimal_Places)

        End If

    End If

End Sub

Function gettotal(X As String) As Double
    Dim My_SQL As String

    My_SQL = "  select Sum(exe) as total  from project_bill_details where bill_id=" & X

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

Private Sub DcbExPercen_Change()
    Label44.Caption = DcbExPercen.text
    ReLineGrid
End Sub

Private Sub DcbExPercen_Click()
    DcbExPercen_Change
End Sub

Private Sub DcbosubContractor_Change()
    Dim fullcode As String
   
    creditlocked = 0
    Dim CPaymentType As Integer
    If Trim(DcbosubContractor.text) = "" Then Exit Sub
    GetCustomersDetail val(DcbosubContractor.BoundText), , fullcode, 3
    Text2.text = fullcode
End Sub

Private Sub DcbosubContractor_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 10
        FrmCompanySearch.show vbModal
           
    End If
        
End Sub

Private Sub Dcbranch_Change()
    If ChekSanNumber(Current_branch, 65) = True Then
        TxtNoteSerial1.text = ""
    End If
    TxtNoteSerial.text = ""
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

                XPTxtSum.text = 0
            End If

        Case 1
            Frame14.Visible = False
            VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("total_salary1")) = val(txt_emp_salary)
            ReLineGrid

    End Select

End Sub
Sub RelineBu22()
    Dim IntCounter   As Integer
    Dim Percetage    As Double
    Dim PercetageAdv As Double
    Dim SumVATLine   As Double
    Dim SumValueLine As Double
    Dim Sm           As Double
    Sm = 0
    SumVATLine = 0
    SumValueLine = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid4
        For i = .FixedRows To .rows - 1
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
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
                If val(TxtFATYou.text) <> 0 And val(.TextMatrix(i, .ColIndex("VAT"))) <> 0 Then
    
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
    
    If SumValueLine = 0 And val(advancedPayment) <> 0 Then
        PercentgValueAddedAccount_Transec XPDtbTrans.value, 6, 1, "", PercetageAdv
        'PercetageAdv = PercetageAdv / 100 + 1
        SumVATLine = val(advancedPayment) * PercetageAdv / 100

    End If
    Label56.Caption = Round(SumValueLine, 2)
    Label57.Caption = Round(SumVATLine, 2)
    Label47.Caption = Sm
    SumVAT
End Sub
Sub SumVAT()
    If Me.TxtModFlg.text <> "R" Then
        Dim val1 As Double
        TxtPreVAT.text = 0
        advancedPayment.text = 0
        If val(txtDiscount2) <> 0 Then
            advancedPayment = txtDiscount2
        Else
            advancedPayment.text = val(Label56.Caption)
        End If
        TxtPreVAT.text = val(Label57.Caption)
        If val(TxtPreBalaTransPyed.text) > 0 Then
            If val(TxtPreBalaVATYu.text) <> 0 Then
                Percetage2 = val(TxtPreBalaVATYu.text) / 100 + 1
                val1 = Round(val(TxtPreBalaTransPyed.text) / Percetage2, 4)
                advancedPayment.text = val(advancedPayment.text) + val1
                TxtPreVAT.text = val(TxtPreVAT.text) + Round(val1 * val(TxtPreBalaVATYu.text) / 100, 4)
            Else
                advancedPayment.text = val(advancedPayment.text) + val(TxtPreBalaTransPyed.text)
                TxtPreVAT.text = 0
            End If
        End If
    End If
End Sub

Private Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
   ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg            As String
    Dim rs             As New ADODB.Recordset
    Dim Rs1            As New ADODB.Recordset
    Dim StrSQL         As String
    Dim ClsAcc         As New ClsAccounts
    Dim LngRow         As Long
    Dim netexe         As Double
    Dim QtyExe         As Double
    
    With Fg_Journal
        If val(.TextMatrix(Row, .ColIndex("ExPercen"))) = 0 Then
            If val(TxtExPercen.text) <> 0 Then
                .TextMatrix(Row, .ColIndex("ExPercen")) = val(TxtExPercen.text)
            Else
                .TextMatrix(Row, .ColIndex("ExPercen")) = 100
            End If
        End If
        Select Case .ColKey(Col)
        
            Case "Account_Serial"
                ' .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                StrSQL = StrSQL & GetAccountByBarnchUser
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    If BolEditOnMainAccounts = False Then
                        'If LastAccount(rs("Account_Code").value) = False Then
                        '    .TextMatrix(Row, Col) = ""
                        '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                        '    Exit Sub
                        'End If
                    End If

                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
                 
xx:
                Else
                    'GetMsgs 130, vbExclamation
                    If Not isFromExcel Then
                        MsgBox "ﬂÊœ Õ”«» Œ«ÿÏ¡", vbCritical
                        .TextMatrix(Row, Col) = ""
                        .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                       
                    Else
                        .TextMatrix(Row, Col) = ""
                        .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                        
                        '                        .Rows = .Rows - 1
                    End If
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing



        Case "item_id"
                If CBoBasedON.ListIndex = 1 And val(TXTOrDer_no2) <> 0 Then
                
'                    StrSQL = " SELECT item_id oprid,item des FROM SubcontractorContract2"
'                    StrSQL = StrSQL + " Where dbo.SubcontractorContract2.bill_id =" & val(Me.TXTOrDer_no2.text)
'                    StrSQL = StrSQL + " order by SubcontractorContract2.item_id"
         
                Else
                    StrSQL = "SELECT  oprid, fullcode,line_no, oprid,des, net, project_id from dbo.projects_des WHERE project_id =" & val(DataCombo2.BoundText)
                    If DcbosubContractor.text <> "" And val(DcbosubContractor.BoundText) <> 0 Then
                        StrSQL = StrSQL & " and  sub_contractor_id =" & val(DcbosubContractor.BoundText) & ""
                    End If
                    If val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) <> 0 Then
                        StrSQL = StrSQL & " and  PrMainDesID =" & val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) & " "
                        StrSQL = StrSQL & " and  Fullcode = " & val(.TextMatrix(Row, .ColIndex("item_id"))) & " "
                    Else
                       Exit Sub
                    End If
                End If
                  
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs.EOF Then
                    .TextMatrix(Row, .ColIndex("oprid")) = val(rs!oprid & "")
                    
                    IsFromItemID = True
                    Fg_Journal_AfterEdit Row, Fg_Journal.ColIndex("item")
                    IsFromItemID = False
                End If

            Case "AccountName"
        
                'sgl = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                'Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                If LngRow <> -1 Then
                    'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    '.TextMatrix(Row, Col) = ""
                    '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Exit Sub
                End If

                Set ClsAcc = New ClsAccounts

                If BolEditOnMainAccounts = False Then
                    'If LastAccount(StrAccountCode) = False Then
                    '    .TextMatrix(Row, Col) = ""
                    '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Else

                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                    'End If
                Else
                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                End If

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
                
                
                If 0 = 0 And Not IsFromItemID Then
                    StrAccountCode = .ComboData
                    LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("oprid"), False, True)
                    .TextMatrix(Row, .ColIndex("oprid")) = StrAccountCode
                Else
                    StrAccountCode = val(.TextMatrix(Row, .ColIndex("oprid")))
                    
                    
                    
                End If
                If StrAccountCode <> "" And val(StrAccountCode) <> 0 Then
                
                    If CBoBasedON.ListIndex = 1 And val(TXTOrDer_no2) <> 0 Then
                        StrSQL = "SELECT   dbo.projects_des.TotalExe, dbo.projects_des.QtyExe, dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, dbo.projects_des.cost, "
                        StrSQL = StrSQL & "        dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id,"
                        StrSQL = StrSQL & "        dbo.projects_des.esQty, dbo.projects_des.Remark, dbo.projects_des.fullcode, dbo.projects_des.PandUnitID, dbo.TblProcessUnites.UnitName,"
                        StrSQL = StrSQL & "        dbo.TblProcessUnites.UnitNamee,"
                        StrSQL = StrSQL & "        qtySubContractor = (Select Top 1 SubcontractorContract2.qtySubContractor from SubcontractorContract2 Where SubcontractorContract2.Bill_id = " & val(TXTOrDer_no2.text) & "  and SubcontractorContract2.project_id =projects_des.project_id and SubcontractorContract2.Item_id = projects_des.oprid) , "
                        StrSQL = StrSQL & "        costSubContractor =( Select Top 1  SubcontractorContract2.costSubContractor from SubcontractorContract2 Where SubcontractorContract2.Bill_id = " & val(TXTOrDer_no2.text) & "  and SubcontractorContract2.project_id =projects_des.project_id and SubcontractorContract2.Item_id = projects_des.oprid)  ,"
                    
                        StrSQL = StrSQL + "  Pre_Quantity_Contr = (Select Sum( CC.quntExc) from project_bill_details CC inner join project_billl On cc.bill_id =project_billl.id  "
                        StrSQL = StrSQL + "  "
                        StrSQL = StrSQL + "  Where cc.oprid ='" & StrAccountCode & "'"
                        StrSQL = StrSQL + "  and CC.project_id =projects_des.project_id"
                        StrSQL = StrSQL + "  and project_billl.OrDer_no2=  " & val(TXTOrDer_no2) & ")"
                    
                        StrSQL = StrSQL & " FROM         dbo.projects_des LEFT OUTER JOIN"
                        StrSQL = StrSQL & "               dbo.TblProcessUnites ON dbo.projects_des.PandUnitID = dbo.TblProcessUnites.UnitID"
                        StrSQL = StrSQL & "  WHERE dbo.projects_des.oprid='" & StrAccountCode & "'"
                        StrSQL = StrSQL & "   and projects_des.project_id In (SELECT project_id FROM SubcontractorContract2 Where SubcontractorContract2.Bill_id = " & val(TXTOrDer_no2.text) & " )"
                
                    Else
                        StrSQL = "SELECT  projects_des.qtySubContractor,projects_des.costSubContractor,  dbo.projects_des.TotalExe, dbo.projects_des.QtyExe, dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, dbo.projects_des.cost, "
                        StrSQL = StrSQL & "        dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id,"
                        StrSQL = StrSQL & "        dbo.projects_des.esQty,Pre_Quantity_Contr = 0, dbo.projects_des.Remark, dbo.projects_des.fullcode, dbo.projects_des.PandUnitID, dbo.TblProcessUnites.UnitName,"
                        StrSQL = StrSQL & "        dbo.TblProcessUnites.UnitNamee"
                        StrSQL = StrSQL & " FROM         dbo.projects_des LEFT OUTER JOIN"
                        StrSQL = StrSQL & "               dbo.TblProcessUnites ON dbo.projects_des.PandUnitID = dbo.TblProcessUnites.UnitID"
                        StrSQL = StrSQL & "  WHERE dbo.projects_des.oprid='" & StrAccountCode & "'"
                        StrSQL = StrSQL & " and dbo.projects_des.project_id =" & val(DataCombo2.BoundText)
                    End If
                    '  StrSQL = "SELECT   line_no, oprid,des, net, project_id  ,[unit] ,[Quantity],[Price] ,[Pre_Quantity] ,[Pre_Value],[Pre_Percent] ,[Curr_Quantity]  ,[Curr_value] ,[curr_Percent] ,[tot_quantity] ,[tot_value] ,[tot_percent]   from dbo.projects_des  WHERE fullcode='" & .ComboData & "'"  ' project_id =" & Val(DataCombo2.BoundText) & "and line_no=" & Val(.ComboItem)
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    .TextMatrix(Row, .ColIndex("qty")) = IIf(IsNull(Rs1("qty").value), 0, Rs1("qty").value)
                    If Me.ChQty.value = vbChecked Then
                        .TextMatrix(Row, .ColIndex("quntExc")) = val(.TextMatrix(Row, .ColIndex("qty")))
                    End If
                    If IsFromItemID Then
                        .TextMatrix(Row, .ColIndex("item")) = IIf(IsNull(Rs1("des").value), "", Rs1("des").value)
                    End If
                    .TextMatrix(Row, .ColIndex("cost")) = IIf(IsNull(Rs1("cost").value), 0, Rs1("cost").value)
                    
                    .TextMatrix(Row, .ColIndex("qtySubContractor")) = IIf(IsNull(Rs1("qtySubContractor").value), 0, Rs1("qtySubContractor").value)
                    .TextMatrix(Row, .ColIndex("costSubContractor")) = IIf(IsNull(Rs1("costSubContractor").value), 0, Rs1("costSubContractor").value)
                    
                    .TextMatrix(Row, .ColIndex("Pre_Quantity_Contr")) = IIf(IsNull(Rs1("Pre_Quantity_Contr").value), 0, Rs1("Pre_Quantity_Contr").value)
                   
                    .TextMatrix(Row, .ColIndex("StillQty")) = (IIf(IsNull(Rs1("qtySubContractor").value), 0, Rs1!qtySubContractor)) - (val(Rs1!Pre_Quantity_Contr & ""))
                    '+ val(Rs1!Pre_Quantity_Contr & ""))
                    
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
                    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                        GetTermsTotals val(.TextMatrix(Row, .ColIndex("oprid"))), val(txtid.text), XPDtbTrans.value, netexe, QtyExe, TxtModFlg.text, , , , , , , val(DataCombo2.BoundText)
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
                   
                    If CBoBasedON.ListIndex = 1 And val(TXTOrDer_no2) <> 0 Then
                        .TextMatrix(Row, .ColIndex("project_id")) = IIf(IsNull(Rs1("project_id").value), "", Rs1("project_id").value)
                    End If
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
            
                If CBoBasedON.ListIndex = 1 And val(TXTOrDer_no2) <> 0 Then
             
                    If val(.TextMatrix(Row, .ColIndex("qtySubContractor"))) - (val(.TextMatrix(Row, .ColIndex("Pre_Quantity_Contr"))) + val(.TextMatrix(Row, .ColIndex("quntExc")))) < 0 Then
                        MsgBox "·« Ì„ﬂ‰ ··ﬂ„Ì… «‰   Ã«Ê“ «·ﬂ„Ì… «·„ »ﬁÌ…"
                        .TextMatrix(Row, .ColIndex("quntExc")) = ""
                        '.TextMatrix(Row, .ColIndex("StillQty")) = val(.TextMatrix(Row, .ColIndex("qtySubContractor"))) - val(.TextMatrix(Row, .ColIndex("Pre_Quantity_Contr")) + val(.TextMatrix(Row, .ColIndex("quntExc"))))
                        Exit Sub
                    End If
             
                    .TextMatrix(Row, .ColIndex("StillQty")) = val(.TextMatrix(Row, .ColIndex("qtySubContractor"))) - (val(.TextMatrix(Row, .ColIndex("Pre_Quantity_Contr"))) + val(.TextMatrix(Row, .ColIndex("quntExc"))))
                    '                If val(.TextMatrix(Row, .ColIndex("quntExc"))) > val(.TextMatrix(Row, .ColIndex("StillQty"))) Then  '- val(.TextMatrix(Row, .ColIndex("qtySubContractor"))) Then
                    '                    MsgBox "·« Ì„ﬂ‰ ··ﬂ„Ì… «‰   Ã«Ê“ «·ﬂ„Ì… «·„ »ﬁÌ…"
                    '                    .TextMatrix(Row, .ColIndex("quntExc")) = ""
                    '                    .TextMatrix(Row, .ColIndex("StillQty")) = val(.TextMatrix(Row, .ColIndex("qtySubContractor"))) - val(.TextMatrix(Row, .ColIndex("Pre_Quantity_Contr")) + val(.TextMatrix(Row, .ColIndex("quntExc"))))
                    '                    Exit Sub
                    '                End If
                End If
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
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With
    
    
    ClculteVAT
   
    
  
    ReLineGrid

End Sub
Sub FillAllBandsToGrid()
    Dim sql As String
    Dim i   As Long
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 2
    sql = " SELECT     dbo.projects_des.PrMainDesID, dbo.ProjectMainDes.ProjectID, dbo.ProjectMainDes.Name, dbo.ProjectMainDes.FullCode, dbo.projects_des.des, "
    sql = sql & "  dbo.projects_des.oprid"
    sql = sql & " FROM         dbo.ProjectMainDes LEFT OUTER JOIN"
    sql = sql & "                      dbo.projects_des ON dbo.ProjectMainDes.ID = dbo.projects_des.PrMainDesID"
    sql = sql & " Where (dbo.ProjectMainDes.ProjectID = " & val(DataCombo2.BoundText) & ")"
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs2.RecordCount > 0 Then
        With Fg_Journal
            rs2.MoveFirst
            .rows = rs2.RecordCount + 1
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("PrMainDesID")) = IIf(IsNull(rs2("PrMainDesID").value), "", rs2("PrMainDesID").value)
                .TextMatrix(i, .ColIndex("MainDes")) = IIf(IsNull(rs2("Name").value), "", rs2("Name").value)
                Fg_Journal_AfterEdit i, .ColIndex("MainDes")
                .TextMatrix(i, .ColIndex("item")) = IIf(IsNull(rs2("des").value), "", rs2("des").value)
                .TextMatrix(i, .ColIndex("oprid")) = IIf(IsNull(rs2("oprid").value), "", rs2("oprid").value)
                Fg_Journal_AfterEdit i, .ColIndex("item")

                rs2.MoveNext
            Next i
        End With
    End If

End Sub

Sub CalCultePers(Optional i As Long)
    With Fg_Journal
        If .TextMatrix(i, .ColIndex("item")) <> "" Then
            If val(.TextMatrix(i, .ColIndex("ExPercen"))) = 0 Then
                If val(TxtExPercen.text) <> 0 Then
                    .TextMatrix(i, .ColIndex("ExPercen")) = val(TxtExPercen.text)
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
               
            Case "Account_Serial"
                .ComboList = ""
            Case "PriceApprov", "QtyApprov"
                '    If SystemOptions.AllowChangePriceApprove = True Then
                .ComboList = ""
                '    Else
                Cancel = True
                '    End If
                
            Case "Pre_Quantity_Contr"
                If Not SystemOptions.SuppCreat4Acc Then
                    Cancel = True
                End If

            Case "qtySubContractor"
                Cancel = True
            Case "costSubContractor"
                Cancel = True
                
            Case "percentage"
                Cancel = True
            Case "net"
                Cancel = True
                
            Case "discount"
                Cancel = True
            Case "LineNo"
                Cancel = True
          
            Case "cost"
                If SystemOptions.AllowChanProjectBillPrice = True Then
                    .ComboList = ""
                Else
                    Cancel = True
                End If
            Case "total"
                Cancel = True
                
            Case "item_id"
                'Cancel = True
            Case "qty"
                Cancel = True
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

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, Shift As Integer)
    With Me.Fg_Journal
        Select Case .ColKey(.Col)
            Case "AccountName", "Account_Serial"
                If KeyCode = vbKeyF3 Then
                    Account_search.show
                    Account_search.case_id = 33350

                End If
        End Select
    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)

    Dim rs             As New ADODB.Recordset
    Dim Rs1            As New ADODB.Recordset
    Dim StrSQL         As String
    Dim StrAccountType As String
    Dim StrComboList   As String
    
    Dim Rs4            As New ADODB.Recordset
    Dim StrComboList_1 As String
    Dim StrSQL_2       As String
     
    Dim Msg            As String

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
                
                If CBoBasedON.ListIndex = 1 And val(TXTOrDer_no2) <> 0 Then
                
                    StrSQL = " SELECT item_id oprid,item des FROM SubcontractorContract2"
                    StrSQL = StrSQL + " Where dbo.SubcontractorContract2.bill_id =" & val(Me.TXTOrDer_no2.text)
                    StrSQL = StrSQL + " order by SubcontractorContract2.item_id"
         
                Else
                    StrSQL = "SELECT  oprid, fullcode,line_no, oprid,des, net, project_id from dbo.projects_des WHERE project_id =" & val(DataCombo2.BoundText)
                    If DcbosubContractor.text <> "" And val(DcbosubContractor.BoundText) <> 0 Then
                        StrSQL = StrSQL & " and  sub_contractor_id =" & val(DcbosubContractor.BoundText) & ""
                    End If
                    If val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) <> 0 Then
                        StrSQL = StrSQL & " and  PrMainDesID =" & val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) & ""
                    End If
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
    TxtModFlg.text = "R"
    Set rs = New ADODB.Recordset
    
    If SystemOptions.AllowEditVaTManulay = True Then
        txtManulaVat.Enabled = True
        txtManulaVat.Visible = True
    Else
        txtManulaVat.Enabled = False
        txtManulaVat.text = 0
        txtManulaVat.Visible = False
    End If


With Me.DefaultInvoicetype
            .Clear
            
             


            .AddItem "Tax Invoice "
            .ItemData(0) = 0
     
            .AddItem "Simplified Tax Invoice"

            .ItemData(1) = 2
         
        End With

    '  StrSQL = "SELECT  dueDate,  branch_no,NoteSerial, id, bill_date, project_no, project_name, Sub_user_name, End_user_name, End_user_account, bill_to, Sub_user_account, bill_type, revenue_account, note_id,"
    '  StrSQL = StrSQL + "total  From dbo.project_billl Order by ID"
    
    'StrSQL = "SELECT  dueDate,  branch_no,NoteSerial, id, bill_date, project_no, project_name, Sub_user_name, End_user_name, End_user_account, bill_to, Sub_user_account, bill_type, revenue_account, note_id,"
    
    StrSQL = StrSQL + "SELECT *  From dbo.project_billl  where 1=1"
    StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ") Order by ID"
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    With Me.CBoBasedON
        .Clear
        '.AddItem "»·«"
        .AddItem "»·«"
        .AddItem "»‰«¡« ⁄·Ï ⁄ﬁœ „ﬁ«Ê·"

    End With


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
         With bill_Type
        .Clear
        .AddItem "«Ì—«œ"
        .AddItem "«Ì—«œ „” Õﬁ"
        End With
    Else
         With Me.CBoBasedON
            .Clear
            '.AddItem "»·«"
            .AddItem "N/A "
            .AddItem "a Contractor Agreement"
    
        End With
        
        
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
    
    
     With bill_Type
        .Clear
        .AddItem "Revenue"
        .AddItem "Accrued Revenue"
    End With
    
    
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
   
    'first_run = True
    Dim My_SQL As String
 
    My_SQL = "  select id,Fullcode from Projects where not (Fullcode is null) and Fullcode <>N'""' "
    My_SQL = My_SQL & "  AND      branch_no in(" & Current_branchSql & ")"
    fill_combo DataCombo2, My_SQL



    My_SQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, My_SQL
    Dim Dcombos As ClsDataCombos
        
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches dcBranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetPersons Me.DcbosubContractor
    Dcombos.GetAccountingCodes Me.DcDiscountAccount, True
    
    Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)
    
    
    If my_language = "E" Then
        CMD_language.ToolTipText = "Change Language"

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
    
    'LoadSettings
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    Set NewGrid.Grid = FG
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
    Set NewGrid.TxtPrice = TxtPrice
    NewGrid.FillGrid
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

        Temp = XPBtnMove(1).left
        XPBtnMove(1).left = XPBtnMove(2).left
        XPBtnMove(2).left = Temp
        Label26.Caption = "Branch"

        Temp = XPBtnMove(0).left
        XPBtnMove(0).left = XPBtnMove(3).left
        XPBtnMove(3).left = Temp
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
Label44.Caption = "Value"
        Frame14.Caption = "Labors Data"
  
        DataGrid1.RightToLeft = False
        CMD_language.Caption = "⁄—»Ì"
'        Frame4.Visible = True
        Frame3.Visible = True

   
        Adodc1.Caption = "move"
  
        With Fg_Journal
        
                    .TextMatrix(0, .ColIndex("Account_serial")) = "Account_Serial"
            .TextMatrix(0, .ColIndex("AccountName")) = "Account Name"
            .TextMatrix(0, .ColIndex("qtySubContractor")) = "Contract Quantity for the Contractor"
            
            .TextMatrix(0, .ColIndex("costSubContractor")) = "Contractor Price"
            .TextMatrix(0, .ColIndex("Pre_Quantity_Contr")) = "Previously Executed Quantity for the Contract"
            .TextMatrix(0, .ColIndex("quntExc")) = "Currently Executed Quantity"
            .TextMatrix(0, .ColIndex("StillQty")) = "Remaining Quantity of the Contract"
            .TextMatrix(0, .ColIndex("exe")) = "Current Executed Price"
            .TextMatrix(0, .ColIndex("percentage")) = "percentage"
            .TextMatrix(0, .ColIndex("totEx")) = "Total Value of Current Works"
            .TextMatrix(0, .ColIndex("discountEXE")) = "discountEXE"
            .TextMatrix(0, .ColIndex("NetExe")) = "NetExe"
            .TextMatrix(0, .ColIndex("percentage1")) = "percentage1"
            .TextMatrix(0, .ColIndex("QtyApprov")) = "QtyApprov"
            .TextMatrix(0, .ColIndex("PriceApprov")) = "PriceApprov"
            .TextMatrix(0, .ColIndex("TotalApprov")) = "TotalApprov"
            .TextMatrix(0, .ColIndex("DiscApprov")) = "DiscApprov"
            .TextMatrix(0, .ColIndex("NetApprov")) = "NetApprov"
            
            .TextMatrix(0, .ColIndex("Pre_Quantity")) = "Pre_Quantity"
            .TextMatrix(0, .ColIndex("Pre_Percent")) = "Pre_Percent"
            .TextMatrix(0, .ColIndex("Pre_Value")) = "Pre_Value"
            .TextMatrix(0, .ColIndex("Pre_Percent1")) = "Pre_Percent1"
            .TextMatrix(0, .ColIndex("tot_quantity")) = "tot_quantity"
            .TextMatrix(0, .ColIndex("tot_percent")) = "tot_percent"
            
            .TextMatrix(0, .ColIndex("tot_percent1")) = "tot_percent1"
            .TextMatrix(0, .ColIndex("OLDTotalwithVat")) = "OLDTotalwithVat"
            .TextMatrix(0, .ColIndex("CurrenttotalWithvat")) = "CurrenttotalWithvat"
            .TextMatrix(0, .ColIndex("Totalwitvat")) = "Totalwitvat"
            
            .TextMatrix(0, .ColIndex("exedate")) = "exedate"



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
  
    SuperLabel2.text = "Search"
    Command1(4).Caption = "By ID"
    Command1(5).Caption = "Search"
    Command1(11).Caption = "Attachement"

    Label29.Caption = "Total"
    Label35.Caption = "Discount"
    Label36.Caption = "Net"
    Label37.Caption = "Advanced"

End Function

Private Sub retrive1(Item_ID As String)
 
    Dim RsDev  As ADODB.Recordset
    Dim StrSQL As String
    Dim i      As Integer
 
    'On Error GoTo ErrTrap
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.rows = 2
    VSFlexGrid2.Enabled = True
    txt_opr_total.text = 0
          
    StrSQL = "select * from terms_operations_project_bill where term_fullcode='" & Item_ID & "' and bill_id=" & val(Me.txtid.text)
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid2
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
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

            Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .rows - 1, .ColIndex("total"))
        
        End With

    End If
          
    ReLineGrid

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
   UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG    As String

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                If SystemOptions.SuppCreat4Acc Then
                    SaveData
                Else
                    SaveDataOld
                End If


            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub GrdBondHistory_AfterEdit(ByVal Row As Long, ByVal Col As Long)
   Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With GrdBondHistory
        .TextMatrix(1, .ColIndex("GuaranteeAmount")) = TxtValue
        Select Case .ColKey(Col)
        
         
           Case "AmountPlus"
            End Select
            If Row > 1 Then
                .TextMatrix(Row, .ColIndex("GuaranteeAmount")) = .TextMatrix(Row - 1, .ColIndex("Total"))
            'ElseIf row > 2 Then
            
            End If
            .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("GuaranteeAmount"))) + val(.TextMatrix(Row, .ColIndex("AmountPlus"))) - val(.TextMatrix(Row, .ColIndex("AmountMin")))
            
            Me.txtTotalBondHistory.text = .TextMatrix(.rows - 2, .ColIndex("Total"))   '.Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .rows - 1, .ColIndex("Total"))
             If Row = .rows - 1 Then
                .rows = .rows + 1
            End If
            txtBondAmt = txtTotalBondHistory
    End With
    
    
    
    
End Sub


Private Sub GrdBondHistory_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  With GrdBondHistory

        If Row > .FixedRows Then
             If val(.TextMatrix(Row, .ColIndex("NoteSerial"))) <> 0 And .ColKey(Col) <> "NoteSerial" Then
                  Cancel = True
                  Exit Sub
              End If
        End If

        Select Case .ColKey(Col)
            Case "Vat", "StillAmount"
                 Cancel = True
            Case "Total", "PayedAmount"
                .ComboList = ""
        Case "Vatyo"
             
            Case "value"
               Cancel = True
              Case "Amount", "MarginNo", "GuaranteeDate", "Serial"
                .ComboList = ""
              Case "ChSameCurrncy"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub GrdBondHistory_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With GrdBondHistory
Select Case .ColKey(Col)
       Case "NoteSerial"
'                                  LngRow = Row
        'ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial"))), , 22004
 
    Case "CreateNote2"
        If val(.TextMatrix(Row, .ColIndex("NoteId2"))) = 0 And val(.TextMatrix(Row, .ColIndex("PayedAmount"))) <> 0 Then
            'CreateEntry row, 1, 0
        End If
    Case "CreateNote"
        If val(.TextMatrix(Row, .ColIndex("NoteId"))) = 0 And val(.TextMatrix(Row, .ColIndex("Total"))) <> 0 Then
            'CreateEntry row, 3, 0
        End If
 End Select
End With

End Sub


Private Sub ImgFavorites_Click()
    AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub


Private Sub ISButton1_Click(Index As Integer)
    
    C1Elastic3.Visible = True
    If Index = 0 Then
        Grid2.Visible = True
        GrdBondHistory.Visible = False
    Else
        Grid2.Visible = False
        GrdBondHistory.Visible = True
        GrdBondHistory.rows = GrdBondHistory.rows + 1

    End If
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
            VSFlexGrid3.rows = 2
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

                XPTxtSum.text = 0
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
   
            .rows = .FixedRows + RsDev.RecordCount
   
            For i = .FixedRows To .rows - 1
            
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
            Me.txt_emp_salary.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .rows - 1, .ColIndex("total"))
            Me.txt_employee_count.text = .Aggregate(flexSTCount, .FixedRows, .ColIndex("total"), .rows - 1, .ColIndex("total"))
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
   
            .rows = .FixedRows + RsDev.RecordCount
   
            For i = .FixedRows To .rows - 1
            
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
            Me.txt_expenses_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
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
            StrSQL = "Delete From terms_operations_project_bill Where term_fullcode ='" & current_terms & "' and bill_id=" & val(Me.txtid.text) ' Val(Me.txt_project_id.text) & "AND item_id=" & current_terms
            Cn.Execute StrSQL, , adExecuteNoRecords
            ' ⁄„·Ì«  «·»‰Êœ
            Set RsDev = New ADODB.Recordset
            RsDev.Open "terms_operations_project_bill", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            Dim i As Integer

            With Me.VSFlexGrid2

                For i = .FixedRows To .rows - 1

                    '
                    If .TextMatrix(i, .ColIndex("fullcode")) <> "" Then

                        RsDev.AddNew
                        RsDev("bill_id").value = val(Me.txtid.text)
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

                XPTxtSum.text = 0
            End If

        Case 1
            Frame10.Visible = False

    End Select

End Sub

Private Sub Retrive2(current_opr As String)
 
    Dim RsDev  As ADODB.Recordset
    Dim StrSQL As String
    Dim i      As Integer
 
    'On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Enabled = True
 
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where    bill_id is null and (payed =1 )  and opr_fullcode='" & current_opr & "' and Transaction_Date<='" & SQLDate(DTPicker1.value) & "'"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = IIf(IsNull(RsDetails("Quantity")), 0, (RsDetails("Quantity").value)) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value))
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Sub

Private Sub StartDateProje_Change()
    Dim str As String
    str = "11/05/2020"
    If StartDateProje.value <= CDate(str) Then
        TxtFATYou.text = 5

        If SystemOptions.AllowEditVaTManulay = True Then
            'If val(txtManulaVat) > 0 Then
            TxtFATYou.text = txtManulaVat
        End If
    Else
        ClculteVAT
    End If
End Sub

Private Sub total_Change()
    Calculte
End Sub

Private Sub TxtBillNo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtBillNo.text, 0)
End Sub

Private Sub txtDiscount_Change()
    txtDiscountG.text = txtDiscount.text
    txtDiscount4.text = txtDiscount3.text
    
    calcnet
    ReLineGrid
    ReLineGrid
End Sub

Function calcnet()
    
If SystemOptions.SuppCreat4Acc = False Then calcnetOld: Exit Function
    
    TxtNetValue.text = Round(val(Results.text) - val(txtDiscountG) - val(advancedPayment2.text) - val(txtDiscountGMater), Decimal_Places)   ' - val(advancedPayment.text)
    If val(cboDiscount1.ListIndex) = 1 Then
        TxtPerforValue.text = Round(val(txtDiscount1.text) * val(TxtNetValue.text) / 100, Decimal_Places)
        TxtPerforValue.text = Round(val(txtDiscount1.text) * val(Results.text) / 100, Decimal_Places)

    ElseIf val(cboDiscount1.ListIndex) = 2 Then
        TxtPerforValue.text = Round(val(txtDiscount1.text), Decimal_Places)
    Else
        TxtPerforValue.text = 0
    End If
    
    txtTotalBefore = Round(val(Results.text) - val(txtDiscountG) - val(advancedPayment2.text) - val(txtDiscountGMater), Decimal_Places)   ' - val(advancedPayment.text)
    total.text = txtTotalBefore - val(txtDiscount3)
    ' ⁄œÌ·
    total.text = Round(val(TxtNetValue.text), Decimal_Places) - val(txtDiscount3)
    Calculte
End Function
Private Sub TxtDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtDiscount.text, 0)
End Sub

Private Sub TxtDiscount1_Change()
    ReLineGrid
     calcnet
    

End Sub

Private Sub txtDiscount2_Change()
    ReLineGrid
    
    advancedPayment = txtDiscount2
    RelineBu22

End Sub

Private Sub txtDiscount3_Change()
    If Me.TxtModFlg.text <> "R" Then
        txtDiscount_Change
    End If
End Sub

Private Sub txtDiscountG_Change()
    calcnet
End Sub

Private Sub txtDiscountGMater_Change()
    calcnet
End Sub

Private Sub TxtExPercen_Change()
    ReLineGrid
End Sub

Private Sub txtId_Change()
    ' "select * from project_bill_details where bill_id=" & Val(txtid.text)

End Sub
Sub ClculteVAT()
    changegridFildssd

    If Me.TxtModFlg.text <> "R" Then
        Dim Percetage As Double
        Dim account   As String
        If val(billto.ListIndex) = 0 Then
            PercentgValueAddedAccount_Transec XPDtbTrans.value, 6, 1, account, Percetage
        ElseIf val(billto.ListIndex) = 1 Then
            PercentgValueAddedAccount_Transec XPDtbTrans.value, 7, 0, account, Percetage
        End If
        TxtFATYou.text = Percetage
    
        If SystemOptions.AllowEditVaTManulay = True Then
            If val(txtManulaVat) <> 0 Or Trim(txtManulaVat) = "00" Then
                TxtFATYou.text = txtManulaVat
            End If
        End If

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

Sub CalculteOld()
If Me.TxtModFlg.text <> "R" Then
If SystemOptions.AllowEditVaTManulay = True Then
    If val(txtManulaVat) > 0 Or Trim(txtManulaVat) = "00" Then TxtFATYou.text = txtManulaVat
End If
If val(TxtFATYou.text) > 0 Then
TxtFATValue.text = Round((val(TxtNetValue.text) * val(TxtFATYou.text)) / 100, Decimal_Places)
'TxtFATValue.text = Round((val(total.text) * val(TxtFATYou.text)) / 100, Decimal_Places)
Else
TxtFATValue.text = 0
End If
TxtTotalValue.text = Round(val(total.text) + val(TxtFATValue.text), Decimal_Places)
End If
End Sub


Sub Calculte()
If SystemOptions.SuppCreat4Acc = False Then CalculteOld: Exit Sub
    
    If Me.TxtModFlg.text <> "R" Then

        If SystemOptions.AllowEditVaTManulay = True Then
            If val(txtManulaVat) > 0 Or Trim(txtManulaVat) = "0" Then TxtFATYou.text = txtManulaVat
        End If
        If val(TxtFATYou.text) > 0 Then
            TxtFATValue.text = Round(((val(TxtNetValue.text) * val(TxtFATYou.text)) / 100), Decimal_Places)
        Else
            TxtFATValue.text = 0
        End If
        'TxtTotalValue.text = Round(val(total.text) + val(TxtFATValue.text) + val(TxtPreVAT2) - val(TxtPerforValue), Decimal_Places)
        'TxtTotalValue.text = Round(val(total.text) + val(TxtFATValue.text) - val(TxtPerforValue) - val(txtDiscount3), Decimal_Places)
        TxtTotalValue.text = Round(val(total.text) + val(TxtFATValue.text) - val(TxtPerforValue), Decimal_Places) + val(txtPerformanceBond)
    End If

    CalcFormat
End Sub

Private Sub CalcFormat()

    Results2 = Format(val(Results.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    txtDiscountG2 = Format(val(txtDiscountG.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    advancedPayment22 = Format(val(advancedPayment2.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    txtTotalBefore = val(total.text) + val(txtDiscount3)
    total2 = Format(val(total.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    TxtFATValue2 = Format(val(TxtFATValue.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    TxtPerforValue2 = Format(val(TxtPerforValue.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    TxtTotalValue2 = Format(val(TxtTotalValue.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

End Sub
Public Sub Retrive(Optional Lngid As Long, Optional note_id As Double = 0)
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 2
    Dim RsDev  As ADODB.Recordset
    Dim StrSQL As String
    Dim i      As Integer

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
    TxtNetValue.text = Round(IIf(IsNull(rs("NetValue").value), IIf(IsNull(rs("total").value), 0, rs("total").value), rs("NetValue").value), Decimal_Places)
    TxtPerforValue.text = Round(IIf(IsNull(rs("PerforValue").value), 0, rs("PerforValue").value), Decimal_Places)
    
    
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    txtDateRec.value = IIf(IsNull(rs("DateRec").value), Date, (rs("DateRec").value))
    zatcaStatus = IIf(IsNull(rs("zatcaStatus").value), 0, rs("zatcaStatus").value)
    TXTIban.text = IIf(IsNull(rs("CIBAN").value), "", (rs("CIBAN").value))
    
    DefaultInvoicetype.ListIndex = IIf(IsNull(rs("Invoicetype").value), 0, rs("Invoicetype").value)
    StartDateProje.value = IIf(IsNull(rs("StartDateProje").value), Date, rs("StartDateProje").value)
    Me.dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    DcDiscountAccount.BoundText = IIf(IsNull(rs("DiscountAccount").value), "", rs("DiscountAccount").value)
    txtid.text = IIf(IsNull(rs("id").value), 0, (rs("id").value))
    TxtPreVAT.text = IIf(IsNull(rs("PreVAT").value), 0, rs("PreVAT").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    XPDtbTrans.value = IIf(IsNull(rs("bill_date").value), Date, rs("bill_date").value)
    dueDate.value = IIf(IsNull(rs("dueDate").value), Date, rs("dueDate").value)
    dueDate1.value = IIf(IsNull(rs("dueDate1").value), Date, rs("dueDate1").value)
       TXTOrDer_no = IIf(IsNull(rs("OrDer_no").value), "", rs("OrDer_no").value)
    TXTOrDer_no2 = IIf(IsNull(rs("OrDer_no2").value), "", rs("OrDer_no2").value)
    
    
    TxtAccountUnderImp.text = IIf(IsNull(rs("AccountUnderImp").value), "", rs("AccountUnderImp").value)

    TxtFATYou.text = IIf(IsNull(rs("FATYou").value), 0, (rs("FATYou").value))
    txtManulaVat.text = IIf(IsNull(rs("FATYou").value), 0, (rs("FATYou").value))
 
    Dim mmm As String
    
    If Not (IsNull(rs("QrCodeImage").value)) Then
        LoadPictureFromDB Picture1, rs, "QrCodeImage", mmm
    Else
        Set Picture1.Picture = Nothing
    End If

    CBoBasedON.ListIndex = IIf(IsNull(rs("CBoBasedON").value), -1, rs("CBoBasedON").value)
 

    TxtFATValue.text = IIf(IsNull(rs("FATValue").value), 0, (rs("FATValue").value))
    TxtTotalValue.text = Round(IIf(IsNull(rs("TotalValue").value), 0, (rs("TotalValue").value)), Decimal_Places)
    Me.AccountVat.BoundText = IIf(IsNull(rs("AccountCodeVat").value), "", (rs("AccountCodeVat").value))
    DataCombo2.BoundText = IIf(IsNull(rs("project_no").value), "", rs("project_no").value)
    '*************************************************
    DcbosubContractor.BoundText = IIf(IsNull(rs("subContractorId").value), "", rs("subContractorId").value)
    txtDiscount1.text = IIf(IsNull(rs("discount1value").value), 0, (rs("discount1value").value))
    txtPerformanceBond = IIf(IsNull(rs("PerformanceBond").value), 0, (rs("PerformanceBond").value))
    txtDiscount2.text = IIf(IsNull(rs("discount2value").value), 0, (rs("discount2value").value))
    
     txtTotalBefore.text = IIf(IsNull(rs("TotalBefore").value), 0, (rs("TotalBefore").value))
     txtDiscount4.text = IIf(IsNull(rs("Discount4").value), 0, (rs("Discount4").value))
     txtDiscount3.text = IIf(IsNull(rs("Discount3").value), 0, (rs("Discount3").value))
txtBondAmt.text = IIf(IsNull(rs("BondAmt").value), "", rs("BondAmt").value)
    
    

    cboDiscount1.ListIndex = IIf(IsNull(rs("discount1ID").value), 0, (rs("discount1ID").value))
    cboDiscount2.ListIndex = IIf(IsNull(rs("discount2ID").value), 0, (rs("discount2ID").value))
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    TxtExPercen.text = IIf(IsNull(rs("ExPercen").value), 0, (rs("ExPercen").value))
    txtDiscountGMater = IIf(IsNull(rs("DiscountGMater").value), 0, (rs("DiscountGMater").value))
    DcbExPercen.ListIndex = IIf(IsNull(rs("ExPercenID").value), -1, (rs("ExPercenID").value))
    '*************************************************
    If Not IsNull(rs("UnderImp").value) Then
        If rs("UnderImp").value = 0 Then
            Option7.value = True
        ElseIf rs("UnderImp").value = 1 Then
            Option6.value = True
        ElseIf rs("UnderImp").value = 2 Then
            Option8.value = True
        End If
    Else
        Option7.value = True
    End If
    '26082015
    txtDiscount.text = Round(IIf(IsNull(rs("Discount").value), 0, (rs("Discount").value)), Decimal_Places)
    Results.text = Round(IIf(IsNull(rs("Results").value), 0, (rs("Results").value)), Decimal_Places)
    advancedPayment.text = Round(IIf(IsNull(rs("advancedPayment").value), 0, (rs("advancedPayment").value)), Decimal_Places)
    ''//////////23 05 2016
    TxtBillNo.text = IIf(IsNull(rs("BillNo").value), 0, rs("BillNo").value)
    txtPeriod.text = IIf(IsNull(rs("Period").value), 0, rs("Period").value)
    Me.DcbPeriodType.ListIndex = IIf(IsNull(rs("PeriodType").value), -1, rs("PeriodType").value)
    Me.TxtRemarks2.text = IIf(IsNull(rs("Remarks2").value), "", rs("Remarks2").value)
    StartDate.value = IIf(IsNull(rs("StartDate").value), Date, rs("StartDate").value)
    ''//
    TxtPreBalaValue.text = IIf(IsNull(rs("PreBalaValue").value), 0, rs("PreBalaValue").value)
    TxtPreBalaVAT.text = IIf(IsNull(rs("PreBalaVAT").value), 0, rs("PreBalaVAT").value)
    TxtPreBalaTotal.text = Round(IIf(IsNull(rs("PreBalaTotal").value), 0, rs("PreBalaTotal").value), Decimal_Places)
    TxtPreBalaPayed.text = IIf(IsNull(rs("PreBalaPayed").value), 0, rs("PreBalaPayed").value)
    TxtPreBalaRemain.text = IIf(IsNull(rs("PreBalaRemain").value), 0, rs("PreBalaRemain").value)
    TxtPreBalaTransPyed.text = IIf(IsNull(rs("PreBalaTransPyed").value), 0, rs("PreBalaTransPyed").value)
    TxtPreBalaNet.text = Round(IIf(IsNull(rs("PreBalaNet").value), 0, rs("PreBalaNet").value), Decimal_Places)
    TxtPreBalaVATYu.text = IIf(IsNull(rs("PreBalaVATYu").value), 0, rs("PreBalaVATYu").value)
    Label57.Caption = IIf(IsNull(rs("SumVATLine").value), 0, rs("SumVATLine").value)
    Label56.Caption = IIf(IsNull(rs("SumValueLine").value), 0, rs("SumValueLine").value)
 
    '26082015
    
    s = "Select  Project_nameE,Project_name from projects where id = " & val(rs!project_no & "")
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        If Trim(rsDummy("Project_nameE").value & "") = "" Then
            txtprojectname.text = IIf(IsNull(rsDummy("Project_name").value), "", rsDummy("Project_name").value)
        Else
            txtprojectname.text = IIf(IsNull(rsDummy("Project_nameE").value), "", rsDummy("Project_nameE").value)
        End If
    Else
        txtprojectname.text = IIf(IsNull(rs("project_name").value), "", rs("project_name").value)
    End If
    '    DcAccount1.text = IIf(IsNull(rs("Sub_user_name").value), "", rs("Sub_user_name").value)
    '    DcAccount2.text = IIf(IsNull(rs("End_user_name").value), "", rs("End_user_name").value)

    txtendaccount.text = IIf(IsNull(rs("End_user_account").value), "", rs("End_user_account").value)
    txtsubaccount.text = IIf(IsNull(rs("Sub_user_account").value), "", rs("Sub_user_account").value)
    txtrevenue_account.text = IIf(IsNull(rs("revenue_account").value), "", rs("revenue_account").value)

    'DcAccount4.text = IIf(IsNull(Rs("sub_contractor_name").value), "", Rs("sub_contractor_name").value)

    'DcAccount2.text = IIf(IsNull(Rs("End_user_name").value), "", Rs("End_user_name").value)

    billto.ListIndex = IIf(IsNull(rs("bill_to").value), -1, rs("bill_to").value)
    If IsNull(rs("bill_type").value) Then
        bill_Type.ListIndex = 0
    Else
        bill_Type.ListIndex = IIf(IsNull(rs("bill_type").value), 0, val(rs("bill_type").value))
    End If
    Me.note_id.text = IIf(IsNull(rs("note_id").value), "", rs("note_id").value)
    TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    txtRemarks.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
    TxtManualNO.text = IIf(IsNull(rs("ManualNo").value), "", rs("ManualNo").value)

    'rs("Remarks").value = Trim(TxtRemarks.text)
    'rs("ManualNo").value = Trim(txtManualNo.text)

    total.text = Round(IIf(IsNull(rs("total").value), 0, rs("total").value), Decimal_Places)

    'Exit Sub

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        If val(TXTOrDer_no2.text) <> 0 And CBoBasedON.ListIndex = 1 Then
            
            StrSQL = "SELECT     TOP 100 PERCENT dbo.TblProcessUnites.UnitName,Accounts.Account_Serial,Accounts.Account_Name, dbo.TblProcessUnites.UnitNamee, dbo.ProjectMainDes.FullCode, dbo.ProjectMainDes.Name,"
            StrSQL = StrSQL + "              dbo.project_bill_details.*,"
            StrSQL = StrSQL + "  Pre_Quantity_Contr = (Select Sum( CC.quntExc) from project_bill_details CC inner join project_billl On cc.bill_id =project_billl.id  "
            StrSQL = StrSQL + "  "
            StrSQL = StrSQL + "  Where cc.oprid =project_bill_details.oprid             "
            StrSQL = StrSQL + "  and CC.project_id =project_bill_details.project_id  "
            StrSQL = StrSQL + "  and project_billl.OrDer_no2=  " & val(TXTOrDer_no2) & "  and  project_billl.Id <> " & val(txtid) & ")"
                    
            StrSQL = StrSQL + "    FROM         dbo.project_bill_details LEFT OUTER JOIN"
            StrSQL = StrSQL + "              dbo.ProjectMainDes ON dbo.project_bill_details.PrMainDesID = dbo.ProjectMainDes.ID LEFT OUTER JOIN"
            StrSQL = StrSQL + "              dbo.TblProcessUnites ON dbo.project_bill_details.Unit_id = dbo.TblProcessUnites.UnitID"
            StrSQL = StrSQL + "              Left outer join Accounts On Accounts.Account_Code = project_bill_details.AccountCode"
                    
            StrSQL = StrSQL + " Where dbo.project_bill_details.bill_id =" & val(Me.txtid.text)
            StrSQL = StrSQL + " order by project_bill_details.id"

        Else
            StrSQL = "SELECT     TOP 100 PERCENT dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee, dbo.ProjectMainDes.FullCode, dbo.ProjectMainDes.Name,Accounts.Account_Serial,Accounts.Account_Name, "
            StrSQL = StrSQL + "              dbo.project_bill_details.*"
            StrSQL = StrSQL + "    FROM         dbo.project_bill_details LEFT OUTER JOIN"
            StrSQL = StrSQL + "              dbo.ProjectMainDes ON dbo.project_bill_details.PrMainDesID = dbo.ProjectMainDes.ID LEFT OUTER JOIN"
            StrSQL = StrSQL + "              dbo.TblProcessUnites ON dbo.project_bill_details.Unit_id = dbo.TblProcessUnites.UnitID"
            StrSQL = StrSQL + "              Left outer join Accounts On Accounts.Account_Code = project_bill_details.AccountCode"
            StrSQL = StrSQL + " Where dbo.project_bill_details.bill_id =" & Me.txtid.text
            StrSQL = StrSQL + " order by project_bill_details.id"
                 
        End If
       
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            'Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            'Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
    
            RsDev.MoveFirst
    
            With Me.Fg_Journal
                .rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .rows - 1
                
                    .TextMatrix(i, .ColIndex("projectName")) = IIf(IsNull(RsDev("projectName").value), 0, RsDev("projectName").value)
                    .TextMatrix(i, .ColIndex("project_id")) = IIf(IsNull(RsDev("project_id").value), 0, RsDev("project_id").value)
                    .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(RsDev("FullCode").value), 0, RsDev("FullCode").value)
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                    .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
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
                    'newwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
                    .TextMatrix(i, .ColIndex("QtyApprov")) = IIf(IsNull(RsDev("QtyApprov").value), 0, RsDev("QtyApprov").value)
                    .TextMatrix(i, .ColIndex("TotalApprov")) = IIf(IsNull(RsDev("TotalApprov").value), 0, RsDev("TotalApprov").value)
                    .TextMatrix(i, .ColIndex("PriceApprov")) = IIf(IsNull(RsDev("PriceApprov").value), 0, RsDev("PriceApprov").value)
                    .TextMatrix(i, .ColIndex("DiscApprov")) = IIf(IsNull(RsDev("DiscApprov").value), 0, RsDev("DiscApprov").value)
                    .TextMatrix(i, .ColIndex("NetApprov")) = IIf(IsNull(RsDev("NetApprov").value), 0, RsDev("NetApprov").value)
                    '''////
                    .TextMatrix(i, .ColIndex("Pre_Quantity_Contr")) = IIf(IsNull(RsDev("Pre_Quantity_Contr").value), 0, RsDev("Pre_Quantity_Contr").value)
                
                    .TextMatrix(i, .ColIndex("discountEXE")) = IIf(IsNull(RsDev("discountEXE").value), 0, RsDev("discountEXE").value)
                    .TextMatrix(i, .ColIndex("NetExe")) = IIf(IsNull(RsDev("NetExe").value), 0, RsDev("NetExe").value)

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
                    .TextMatrix(i, .ColIndex("StillQty")) = val(RsDev!qtySubContractor & "") - val(RsDev!Pre_Quantity_Contr & "") + val(RsDev!quntExc & "")
                    'IIf(IsNull(RsDev("qtySubContractor").value), "", RsDev("qtySubContractor").value) - (IIf(IsNull(RsDev("Pre_Quantity_Contr").value), 0, val(RsDev("Pre_Quantity_Contr").value) & "") + IIf(IsNull(RsDev("quntExc").value), 0, RsDev("quntExc").value))
                    
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
        
              

        
StrSQL = "sELECT * from TBLProjectBillHistory  where bill_id=" & val(Me.txtid.text)
loadgrid StrSQL, GrdBondHistory, True, False


        If GrdBondHistory.rows - 1 > 1 Then
                Me.txtTotalBondHistory.text = GrdBondHistory.Aggregate(flexSTSum, GrdBondHistory.FixedRows, GrdBondHistory.ColIndex("Total"), GrdBondHistory.rows - 1, GrdBondHistory.ColIndex("Total"))
        End If


    RetriveBillBuyData
    ReLineGrid
    ReLineGrid
    
    fillapprovData
    changegridFildssd
    CalcFormat
    Exit Sub
ErrTrap:

End Sub
Function changegridFildssd()
    With Fg_Journal
        If billto.ListIndex = 0 Then
            If Me.TxtModFlg <> "R" Then
                DcbosubContractor.BoundText = 0
            End If
            DcbosubContractor.Visible = False
            Text2.Visible = False
            Label22.Visible = False
    
            .ColHidden(.ColIndex("qty")) = False
            .ColHidden(.ColIndex("cost")) = False
            .ColHidden(.ColIndex("total")) = False
            .ColHidden(.ColIndex("discount")) = False
            .ColHidden(.ColIndex("net")) = False
            .ColHidden(.ColIndex("qtySubContractor")) = True
            .ColHidden(.ColIndex("costSubContractor")) = True
        Else '„ﬁ«Ê·
    
            DcbosubContractor.Visible = True
            Text2.Visible = True
            Label22.Visible = True
    
            .ColHidden(.ColIndex("qty")) = True
            .ColHidden(.ColIndex("cost")) = True
            .ColHidden(.ColIndex("total")) = True
            .ColHidden(.ColIndex("discount")) = True
            .ColHidden(.ColIndex("net")) = True
            .ColHidden(.ColIndex("qtySubContractor")) = False
            .ColHidden(.ColIndex("costSubContractor")) = False
     
        End If
      
    End With
 
End Function
Sub RelineBuy()
    Dim IntCounter As Integer
    Dim Sm         As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid4
        For i = .FixedRows To .rows - 1
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
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
    Dim i          As Long
    Dim sql        As String
    Dim rs         As ADODB.Recordset
    Dim IntCounter As Integer
    Dim XRound     As Integer
    changegridFildssd

    If SystemOptions.AllowNoRoudProjectInvoices = True Then
        XRound = val(cCompanyInfo.NoRoudProjectInvoices)
    Else
        XRound = 2
    End If
    With Fg_Journal

        For i = .FixedRows To .rows - 1

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
 
                Dim netexe                          As Double
                Dim QtyExe                          As Double
                Dim VATPer                          As Double
                Dim oldPerforValue                  As Double
                Dim discountHasmyat                 As Double
                Dim linenetaftermainDiscountWithvat As Double
                linenetaftermainDiscountWithvat = 0
                '    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                GetTermsTotals val(.TextMatrix(i, .ColIndex("oprid"))), val(txtid.text), XPDtbTrans.value, netexe, QtyExe, TxtModFlg.text, billto.ListIndex, val(DcbosubContractor.BoundText), VATPer, oldPerforValue, discountHasmyat, linenetaftermainDiscountWithvat, val(DataCombo2.BoundText)
                '    End If
  
                '************************** cancelled***********************************************
                .TextMatrix(i, .ColIndex("Pre_Quantity")) = QtyExe
                .TextMatrix(i, .ColIndex("Pre_Value")) = netexe
       
                .TextMatrix(i, .ColIndex("OLDTotalwithVat")) = Round(linenetaftermainDiscountWithvat, 2)
   
                If val(Results.text) <> 0 Then
                    LineDiscountPercent = val(.TextMatrix(i, .ColIndex("NetExe"))) / val(Results.text)
                End If
                LineDiscount = (val(txtDiscountG.text)) * LineDiscountPercent
               
                linenetaftermainDiscount = val(.TextMatrix(i, .ColIndex("NetExe"))) - LineDiscount
 
                .TextMatrix(i, .ColIndex("CurrenttotalWithvat")) = Round(((linenetaftermainDiscount - val(.TextMatrix(i, .ColIndex("discountEXE"))))) * (1 + TxtFATYou / 100), 2)
                .TextMatrix(i, .ColIndex("Totalwitvat")) = Round(val(.TextMatrix(i, .ColIndex("CurrenttotalWithvat"))) + .TextMatrix(i, .ColIndex("OLDTotalwithVat")), 2)
          
                lbl(9).Caption = Round(oldPerforValue, 2)
                lbl(10).Caption = Round(val(TxtPerforValue.text), 2)
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

        Me.Results.text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("NetExe"), .rows - 1, .ColIndex("NetExe")), Decimal_Places)
        '  Me.total.Text = val(Me.Results.Text) - val(txtDiscountG.Text)  ' .Aggregate(flexSTSum, .FixedRows, .ColIndex("totEx"), .Rows - 1, .ColIndex("totEx"))
        calcnet
         
    End With

    IntCounter = 0

    With VSFlexGrid2

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
          
                .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("total_expenses1"))) + val(.TextMatrix(i, .ColIndex("total_salary1"))) + val(.TextMatrix(i, .ColIndex("total_items1")))
           
            End If

        Next i

        Me.txt_opr_total.text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .rows - 1, .ColIndex("total")), Decimal_Places)
    End With

    IntCounter = 0

    With VSFlexGrid1

        For i = .FixedRows To .rows - 1

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

    Select Case Me.TxtModFlg.text

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

Public Sub TXTOrDer_no2_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If Me.TxtModFlg = "R" Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim s   As String
    Dim rs2 As New ADODB.Recordset
    ' If CBoBasedON.ListIndex = 0 And val(TXTOrDer_no2.Text) <> 0 Then Exit Sub
    '  Dim s As String

    If CBoBasedON.ListIndex = -1 Then Exit Sub
    'Else
    Dim Dcombos As ClsDataCombos
    Dim StrSQL  As String
    Set Dcombos = New ClsDataCombos
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select CusID,CusName From TblCustemers"
    Else
        StrSQL = "Select CusID,CusNamee From TblCustemers"
    End If
    StrSQL = StrSQL & " Where Type = 3"
              
    'fill_combo Me.DcbosubContractor, StrSQL
    Dcombos.GetPersons Me.DcbosubContractor
    'End If
            
    Dim orderStatus As Integer
     
    MintDone = 0
    Set rs2 = New ADODB.Recordset
    If CBoBasedON.ListIndex = 1 Then
        StrSQL = "select * from SubcontractorContract where NoteSerial1= " & val(TXTOrDer_no2.text) & " "
        Set rs2 = New ADODB.Recordset
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs2.EOF Then
            TXTOrDer_no = val(rs2!ID & "")
        End If
        ' StrSQL = "select * from SubcontractorContract where iD= " & val(TXTOrDer_no.Text) & " "
            
    Else
        TXTOrDer_no2 = ""
        TXTOrDer_no = ""
                
    End If
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs2.RecordCount > 0 Then
                
        TXTOrDer_no = val(rs2!ID & "")
        TxtNetValue.text = Round(IIf(IsNull(rs2("NetValue").value), IIf(IsNull(rs2("total").value), 0, rs2("total").value), rs2("NetValue").value), Decimal_Places)
        TxtPerforValue.text = Round(IIf(IsNull(rs2("PerforValue").value), 0, rs2("PerforValue").value), Decimal_Places)
                
        StartDateProje.value = IIf(IsNull(rs2("StartDateProje").value), Date, rs2("StartDateProje").value)
        Me.dcBranch.BoundText = IIf(IsNull(rs2("branch_no").value), "", rs2("branch_no").value)
        'txtid.Text = IIf(IsNull(rs2("id").value), 0, (rs2("id").value))
        TxtPreVAT.text = IIf(IsNull(rs2("PreVAT").value), 0, rs2("PreVAT").value)
        'Me.TxtNoteSerial1.Text = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
        '  Me.TxtNoteSerial1.Text = IIf(IsNull(rs2("Id").Value), "", rs2("Id").Value)
        XPDtbTrans.value = Date ' IIf(IsNull(rs2("bill_date").value), Date, rs2("bill_date").value)
        dueDate.value = IIf(IsNull(rs2("dueDate").value), Date, rs2("dueDate").value)
        'dueDate1.value = IIf(IsNull(rs2("dueDate1").value), Date, rs2("dueDate1").value)
        'TxtAccountUnderImp.Text = IIf(IsNull(rs2("AccountUnderImp").value), "", rs2("AccountUnderImp").value)
        ' TxtFATYou.Text = IIf(IsNull(rs2("FATYou").Value), 0, (rs2("FATYou").Value))
        '   txtManulaVat.Text = IIf(IsNull(rs2("FATYou").Value), 0, (rs2("FATYou").Value))
                
        ' TxtFATValue.Text = IIf(IsNull(rs2("FATValue").Value), 0, (rs2("FATValue").Value))
        '  TxtTotalValue.Text = Round(IIf(IsNull(rs2("TotalValue").Value), 0, (rs2("TotalValue").Value)), Decimal_Places)
        Me.AccountVat.BoundText = IIf(IsNull(rs2("AccountCodeVat").value), "", (rs2("AccountCodeVat").value))
        DataCombo2.BoundText = IIf(IsNull(rs2("project_no").value), "", rs2("project_no").value)
        '*************************************************
        billto.ListIndex = 1
        DcbosubContractor.BoundText = IIf(IsNull(rs2("subContractorId").value), "", rs2("subContractorId").value)
        txtDiscount1.text = IIf(IsNull(rs2("discount1value").value), 0, (rs2("discount1value").value))
        txtDiscount2.text = IIf(IsNull(rs2("discount2value").value), 0, (rs2("discount2value").value))
                
        cboDiscount1.ListIndex = IIf(IsNull(rs2("discount1ID").value), 0, (rs2("discount1ID").value))
        cboDiscount2.ListIndex = IIf(IsNull(rs2("discount2ID").value), 0, (rs2("discount2ID").value))
        ' DCboUserName.BoundText = IIf(IsNull(rs2("UserID").value), "", rs2("UserID").value)
        'TxtCost.Text = IIf(IsNull(rs2("ExPercen").value), 0, (rs2("ExPercen").value))
        DcbExPercen.ListIndex = IIf(IsNull(rs2("ExPercenID").value), -1, (rs2("ExPercenID").value))
                    
        Results.text = Round(IIf(IsNull(rs2("Results").value), 0, (rs2("Results").value)), Decimal_Places)
        'advancedPayment.Text = Round(IIf(IsNull(rs2("advancedPayment").value), 0, (rs2("advancedPayment").value)), Decimal_Places)
        ''//////////23 05 2016
        TxtBillNo.text = IIf(IsNull(rs2("BillNo").value), 0, rs2("BillNo").value)
                
        Me.DcbPeriodType.ListIndex = IIf(IsNull(rs2("PeriodType").value), -1, rs2("PeriodType").value)
        'Me.TxtRemarks2.Text = IIf(IsNull(rs2("Remarks2").value), "", rs2("Remarks2").value)
        'startDate.value = IIf(IsNull(rs2("StartDate").value), Date, rs2("StartDate").value)
        ''//
        TxtPreBalaValue.text = IIf(IsNull(rs2("PreBalaValue").value), 0, rs2("PreBalaValue").value)
        TxtPreBalaVAT.text = IIf(IsNull(rs2("PreBalaVAT").value), 0, rs2("PreBalaVAT").value)
        TxtPreBalaTotal.text = Round(IIf(IsNull(rs2("PreBalaTotal").value), 0, rs2("PreBalaTotal").value), Decimal_Places)
        TxtPreBalaPayed.text = IIf(IsNull(rs2("PreBalaPayed").value), 0, rs2("PreBalaPayed").value)
        TxtPreBalaRemain.text = IIf(IsNull(rs2("PreBalaRemain").value), 0, rs2("PreBalaRemain").value)
        TxtPreBalaTransPyed.text = IIf(IsNull(rs2("PreBalaTransPyed").value), 0, rs2("PreBalaTransPyed").value)
        TxtPreBalaNet.text = Round(IIf(IsNull(rs2("PreBalaNet").value), 0, rs2("PreBalaNet").value), Decimal_Places)
        TxtPreBalaVATYu.text = IIf(IsNull(rs2("PreBalaVATYu").value), 0, rs2("PreBalaVATYu").value)
        Label57.Caption = IIf(IsNull(rs2("SumVATLine").value), 0, rs2("SumVATLine").value)
        Label56.Caption = IIf(IsNull(rs2("SumValueLine").value), 0, rs2("SumValueLine").value)
                
        '26082015
        txtprojectname.text = IIf(IsNull(rs2("project_name").value), "", rs2("project_name").value)
        '    DcAccount1.text = IIf(IsNull(rs2("Sub_user_name").value), "", rs2("Sub_user_name").value)
        '    DcAccount2.text = IIf(IsNull(rs2("End_user_name").value), "", rs2("End_user_name").value)
                
        txtendaccount.text = IIf(IsNull(rs2("End_user_account").value), "", rs2("End_user_account").value)
        'txtsubaccount.Text = IIf(IsNull(rs2("Sub_user_account").value), "", rs2("Sub_user_account").value)
        txtrevenue_account.text = IIf(IsNull(rs2("revenue_account").value), "", rs2("revenue_account").value)
                
        'DcAccount4.text = IIf(IsNull(Rs("sub_contractor_name").value), "", Rs("sub_contractor_name").value)
                
        'DcAccount2.text = IIf(IsNull(Rs("End_user_name").value), "", Rs("End_user_name").value)
                
        '  billto.ListIndex = IIf(IsNull(rs2("bill_to").Value), -1, rs2("bill_to").Value)
        '                If IsNull(rs2("bill_type").Value) Then
        '                bill_Type.ListIndex = 0
        '                Else
        '                bill_Type.ListIndex = IIf(IsNull(rs2("bill_type").Value), 0, val(rs2("bill_type").Value))
        '                End If
        ' Me.note_id.Text = IIf(IsNull(rs2("note_id").value), "", rs2("note_id").value)
        ' TxtNoteSerial.Text = IIf(IsNull(rs2("NoteSerial").value), "", rs2("NoteSerial").value)
        'TxtRemarks.Text = IIf(IsNull(rs2("Remarks").value), "", rs2("Remarks").value)
        '   txtManualNo.Text = IIf(IsNull(rs2("ManualNo").value), "", rs2("ManualNo").value)

        'rs2("Remarks").value = Trim(TxtRemarks.text)
        'rs2("ManualNo").value = Trim(txtManualNo.text)

        total.text = Round(IIf(IsNull(rs2("total").value), 0, rs2("total").value), Decimal_Places)
                
        Dim X As Integer

        If SystemOptions.UserInterface = EnglishInterface Then
            X = MsgBox("Do you want to list all the terms of the contract?", vbCritical + vbYesNo)
        Else
            X = MsgBox("Â·  —Ìœ «œ—«Ã ﬂ· »‰Êœ «·⁄ﬁœ", vbCritical + vbYesNo)
        End If

        If X = vbNo Then Exit Sub
    
        Dim sql As String
    
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee, dbo.ProjectMainDes.FullCode, dbo.ProjectMainDes.Name,"
        StrSQL = StrSQL + "              dbo.project_bill_details.*,"
        StrSQL = StrSQL + "  Pre_Quantity_Contr = (Select Sum( CC.quntExc) from project_bill_details CC inner join project_billl On cc.bill_id =project_billl.id  "
        StrSQL = StrSQL + "  "
        StrSQL = StrSQL + "  Where cc.oprid =SubcontractorContract2.oprid             "
        StrSQL = StrSQL + "  and CC.project_id =SubcontractorContract2.project_id  "
        StrSQL = StrSQL + "  and project_billl.OrDer_no2=  " & val(TXTOrDer_no2) & ")"
                    
        StrSQL = StrSQL + "    FROM         dbo.project_bill_details LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.ProjectMainDes ON dbo.project_bill_details.PrMainDesID = dbo.ProjectMainDes.ID LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.TblProcessUnites ON dbo.project_bill_details.Unit_id = dbo.TblProcessUnites.UnitID"
        StrSQL = StrSQL + " Where dbo.project_bill_details.bill_id =" & val(Me.txtid.text)
        StrSQL = StrSQL + " order by project_bill_details.id"
    
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee, dbo.ProjectMainDes.FullCode, dbo.ProjectMainDes.Name,Accounts.Account_Serial,Accounts.Account_Name, "
        StrSQL = StrSQL + "              dbo.SubcontractorContract2.*,"
        
        StrSQL = StrSQL + "  Pre_Quantity_Contr = (Select Sum( CC.quntExc) from project_bill_details CC inner join project_billl On cc.bill_id =project_billl.id  "
        StrSQL = StrSQL + "  "
        StrSQL = StrSQL + "  Where cc.oprid =SubcontractorContract2.oprid             "
        StrSQL = StrSQL + "  and CC.project_id =SubcontractorContract2.project_id  "
        StrSQL = StrSQL + "  and project_billl.Id <> " & val(Me.txtid.text) & "   and project_billl.OrDer_no2=  " & val(TXTOrDer_no2) & ")"
                    
        StrSQL = StrSQL + "    FROM         dbo.SubcontractorContract2 LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.ProjectMainDes ON dbo.SubcontractorContract2.PrMainDesID = dbo.ProjectMainDes.ID LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.TblProcessUnites ON dbo.SubcontractorContract2.Unit_id = dbo.TblProcessUnites.UnitID"
        StrSQL = StrSQL + "              Left outer join Accounts On Accounts.Account_Code = SubcontractorContract2.AccountCode"
        StrSQL = StrSQL + " Where dbo.SubcontractorContract2.bill_id =" & val(Me.TXTOrDer_no.text)
        StrSQL = StrSQL + " order by SubcontractorContract2.id"
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            'Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            'Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
    
            RsDev.MoveFirst
    
            With Me.Fg_Journal
                .rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .rows - 1
                    .TextMatrix(i, .ColIndex("ExPercen")) = 0 'IIf(IsNull(RsDev("ExPercen").value), 0, RsDev("ExPercen").value)
                    .TextMatrix(i, .ColIndex("PrMainDesID")) = IIf(IsNull(RsDev("PrMainDesID").value), 0, RsDev("PrMainDesID").value)
                    .TextMatrix(i, .ColIndex("CodeMain")) = IIf(IsNull(RsDev("FullCode").value), "", RsDev("FullCode").value)
                    .TextMatrix(i, .ColIndex("MainDes")) = "ff" 'IIf(IsNull(RsDev("Name").value), "", RsDev("Name").value)
                    .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(RsDev("qty").value), 0, RsDev("qty").value)
                    '.TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(RsDev("qtySubContractor").value), "", RsDev("qtySubContractor").value)
                   
                    '   .TextMatrix(i, .ColIndex("exe")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    .TextMatrix(i, .ColIndex("costSubContractor")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    '.TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    
                    .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("cost").value), "", RsDev("cost").value)
                                
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                    .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
                    
                    ' .TextMatrix(i, .ColIndex("quntExc")) = IIf(IsNull(RsDev("quntExc").value), 0, RsDev("quntExc").value)
                    .TextMatrix(i, .ColIndex("quntExc")) = "" ' IIf(IsNull(RsDev("qtySubContractor").Value), 0, RsDev("qtySubContractor").Value)
                    .TextMatrix(i, .ColIndex("exe")) = IIf(IsNull(RsDev("exe").value), "", RsDev("exe").value)
                
                    .TextMatrix(i, .ColIndex("oprid")) = IIf(IsNull(RsDev("oprid").value), 0, RsDev("oprid").value)
                    .TextMatrix(i, .ColIndex("totEx")) = IIf(IsNull(RsDev("totEx").value), 0, RsDev("totEx").value)
                
                    .TextMatrix(i, .ColIndex("net")) = IIf(IsNull(RsDev("net").value), 0, RsDev("net").value)
                    .TextMatrix(i, .ColIndex("discount")) = IIf(IsNull(RsDev("discount").value), 0, RsDev("discount").value)
                    .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), 0, RsDev("total").value)
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
                    
                    .TextMatrix(i, .ColIndex("Pre_Quantity_Contr")) = IIf(IsNull(RsDev("Pre_Quantity_Contr").value), "", RsDev("Pre_Quantity_Contr").value)
                    
                    .TextMatrix(i, .ColIndex("qtySubContractor")) = IIf(IsNull(RsDev("qtySubContractor").value), "", RsDev("qtySubContractor").value)
                    .TextMatrix(i, .ColIndex("costSubContractor")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    .TextMatrix(i, .ColIndex("StillQty")) = val(.TextMatrix(i, .ColIndex("qtySubContractor"))) - val(.TextMatrix(i, .ColIndex("Pre_Quantity_Contr")))
                    
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
                    
                    '
           
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
                
                    '.TextMatrix(i, .ColIndex("StillQty")) = .TextMatrix(i, .ColIndex("qtySubContractor"))
        
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
            ReLineGrid
        End If

    End If
       
    Exit Sub

Exits:

    TXTOrDer_no = ""
    TXTOrDer_no2 = ""
            
End Sub

Public Sub getDataBill()
    'Dcombos.GetPersons Me.DcbosubContractor
    'End If
            
    Dim orderStatus As Integer
     
    MintDone = 0
    Set rs2 = New ADODB.Recordset
    If CBoBasedON.ListIndex = 1 Then
        StrSQL = "select SubcontractorContract.* from SubcontractorContract "
        'StrSQL = StrSQL + "              Left outer join Accounts On Accounts.Account_Code = SubcontractorContract2.AccountCode"
        StrSQL = StrSQL + "              where NoteSerial1= " & val(TXTOrDer_no2.text) & " "
        Set rs2 = New ADODB.Recordset
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs2.EOF Then
            TXTOrDer_no = val(rs2!ID & "")
        End If
        ' StrSQL = "select * from SubcontractorContract where iD= " & val(TXTOrDer_no.Text) & " "
            
    Else
        TXTOrDer_no2 = ""
        TXTOrDer_no = ""
                
    End If
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs2.RecordCount > 0 Then
                
        TXTOrDer_no = val(rs2!ID & "")
        TxtNetValue.text = Round(IIf(IsNull(rs2("NetValue").value), IIf(IsNull(rs2("total").value), 0, rs2("total").value), rs2("NetValue").value), Decimal_Places)
        TxtPerforValue.text = Round(IIf(IsNull(rs2("PerforValue").value), 0, rs2("PerforValue").value), Decimal_Places)
                
        StartDateProje.value = IIf(IsNull(rs2("StartDateProje").value), Date, rs2("StartDateProje").value)
        Me.dcBranch.BoundText = IIf(IsNull(rs2("branch_no").value), "", rs2("branch_no").value)
        'txtid.Text = IIf(IsNull(rs2("id").value), 0, (rs2("id").value))
        TxtPreVAT.text = IIf(IsNull(rs2("PreVAT").value), 0, rs2("PreVAT").value)
        'Me.TxtNoteSerial1.Text = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
        '  Me.TxtNoteSerial1.Text = IIf(IsNull(rs2("Id").Value), "", rs2("Id").Value)
        XPDtbTrans.value = IIf(IsNull(rs2("bill_date").value), Date, rs2("bill_date").value)
        dueDate.value = IIf(IsNull(rs2("dueDate").value), Date, rs2("dueDate").value)
        'dueDate1.value = IIf(IsNull(rs2("dueDate1").value), Date, rs2("dueDate1").value)
        'TxtAccountUnderImp.Text = IIf(IsNull(rs2("AccountUnderImp").value), "", rs2("AccountUnderImp").value)
        ' TxtFATYou.Text = IIf(IsNull(rs2("FATYou").Value), 0, (rs2("FATYou").Value))
        '   txtManulaVat.Text = IIf(IsNull(rs2("FATYou").Value), 0, (rs2("FATYou").Value))
                
        ' TxtFATValue.Text = IIf(IsNull(rs2("FATValue").Value), 0, (rs2("FATValue").Value))
        '  TxtTotalValue.Text = Round(IIf(IsNull(rs2("TotalValue").Value), 0, (rs2("TotalValue").Value)), Decimal_Places)
        Me.AccountVat.BoundText = IIf(IsNull(rs2("AccountCodeVat").value), "", (rs2("AccountCodeVat").value))
        DataCombo2.BoundText = IIf(IsNull(rs2("project_no").value), "", rs2("project_no").value)
        '*************************************************
        billto.ListIndex = 1
        DcbosubContractor.BoundText = IIf(IsNull(rs2("subContractorId").value), "", rs2("subContractorId").value)
        txtDiscount1.text = IIf(IsNull(rs2("discount1value").value), 0, (rs2("discount1value").value))
        txtDiscount2.text = IIf(IsNull(rs2("discount2value").value), 0, (rs2("discount2value").value))
                
        cboDiscount1.ListIndex = IIf(IsNull(rs2("discount1ID").value), 0, (rs2("discount1ID").value))
        cboDiscount2.ListIndex = IIf(IsNull(rs2("discount2ID").value), 0, (rs2("discount2ID").value))
        DCboUserName.BoundText = IIf(IsNull(rs2("UserID").value), "", rs2("UserID").value)
        'TxtCost.Text = IIf(IsNull(rs2("ExPercen").value), 0, (rs2("ExPercen").value))
        DcbExPercen.ListIndex = IIf(IsNull(rs2("ExPercenID").value), -1, (rs2("ExPercenID").value))
                    
        Results.text = Round(IIf(IsNull(rs2("Results").value), 0, (rs2("Results").value)), Decimal_Places)
        'advancedPayment.Text = Round(IIf(IsNull(rs2("advancedPayment").value), 0, (rs2("advancedPayment").value)), Decimal_Places)
        ''//////////23 05 2016
        TxtBillNo.text = IIf(IsNull(rs2("BillNo").value), 0, rs2("BillNo").value)
                
        Me.DcbPeriodType.ListIndex = IIf(IsNull(rs2("PeriodType").value), -1, rs2("PeriodType").value)
        'Me.TxtRemarks2.Text = IIf(IsNull(rs2("Remarks2").value), "", rs2("Remarks2").value)
        'startDate.value = IIf(IsNull(rs2("StartDate").value), Date, rs2("StartDate").value)
        ''//
        TxtPreBalaValue.text = IIf(IsNull(rs2("PreBalaValue").value), 0, rs2("PreBalaValue").value)
        TxtPreBalaVAT.text = IIf(IsNull(rs2("PreBalaVAT").value), 0, rs2("PreBalaVAT").value)
        TxtPreBalaTotal.text = Round(IIf(IsNull(rs2("PreBalaTotal").value), 0, rs2("PreBalaTotal").value), Decimal_Places)
        TxtPreBalaPayed.text = IIf(IsNull(rs2("PreBalaPayed").value), 0, rs2("PreBalaPayed").value)
        TxtPreBalaRemain.text = IIf(IsNull(rs2("PreBalaRemain").value), 0, rs2("PreBalaRemain").value)
        TxtPreBalaTransPyed.text = IIf(IsNull(rs2("PreBalaTransPyed").value), 0, rs2("PreBalaTransPyed").value)
        TxtPreBalaNet.text = Round(IIf(IsNull(rs2("PreBalaNet").value), 0, rs2("PreBalaNet").value), Decimal_Places)
        TxtPreBalaVATYu.text = IIf(IsNull(rs2("PreBalaVATYu").value), 0, rs2("PreBalaVATYu").value)
        Label57.Caption = IIf(IsNull(rs2("SumVATLine").value), 0, rs2("SumVATLine").value)
        Label56.Caption = IIf(IsNull(rs2("SumValueLine").value), 0, rs2("SumValueLine").value)
                
        '26082015
        txtprojectname.text = IIf(IsNull(rs2("project_name").value), "", rs2("project_name").value)
        '    DcAccount1.text = IIf(IsNull(rs2("Sub_user_name").value), "", rs2("Sub_user_name").value)
        '    DcAccount2.text = IIf(IsNull(rs2("End_user_name").value), "", rs2("End_user_name").value)
                
        txtendaccount.text = IIf(IsNull(rs2("End_user_account").value), "", rs2("End_user_account").value)
        'txtsubaccount.Text = IIf(IsNull(rs2("Sub_user_account").value), "", rs2("Sub_user_account").value)
        txtrevenue_account.text = IIf(IsNull(rs2("revenue_account").value), "", rs2("revenue_account").value)
                
        'DcAccount4.text = IIf(IsNull(Rs("sub_contractor_name").value), "", Rs("sub_contractor_name").value)
                
        'DcAccount2.text = IIf(IsNull(Rs("End_user_name").value), "", Rs("End_user_name").value)
                
        '  billto.ListIndex = IIf(IsNull(rs2("bill_to").Value), -1, rs2("bill_to").Value)
        '                If IsNull(rs2("bill_type").Value) Then
        '                bill_Type.ListIndex = 0
        '                Else
        '                bill_Type.ListIndex = IIf(IsNull(rs2("bill_type").Value), 0, val(rs2("bill_type").Value))
        '                End If
        ' Me.note_id.Text = IIf(IsNull(rs2("note_id").value), "", rs2("note_id").value)
        ' TxtNoteSerial.Text = IIf(IsNull(rs2("NoteSerial").value), "", rs2("NoteSerial").value)
        'TxtRemarks.Text = IIf(IsNull(rs2("Remarks").value), "", rs2("Remarks").value)
        '   txtManualNo.Text = IIf(IsNull(rs2("ManualNo").value), "", rs2("ManualNo").value)

        'rs2("Remarks").value = Trim(TxtRemarks.text)
        'rs2("ManualNo").value = Trim(txtManualNo.text)

        total.text = Round(IIf(IsNull(rs2("total").value), 0, rs2("total").value), Decimal_Places)
                
        Dim X As Integer

        If SystemOptions.UserInterface = EnglishInterface Then
            X = MsgBox("Do you want to list all the terms of the contract?", vbCritical + vbYesNo)
        Else
            X = MsgBox("Â·  —Ìœ «œ—«Ã ﬂ· »‰Êœ «·⁄ﬁœ", vbCritical + vbYesNo)
        End If

        If X = vbNo Then Exit Sub
    
        Dim sql As String
    
        StrSQL = "SELECT     TOP 100 PERCENT "
        StrSQL = StrSQL + "              dbo.project_bill_details.*,"
        StrSQL = StrSQL + "  dbo.TblProcessUnites.UnitName , dbo.TblProcessUnites.UnitNamee, dbo.ProjectMainDes.Fullcode, dbo.ProjectMainDes.Name,"
        StrSQL = StrSQL + "  Pre_Quantity_Contr = (Select Sum( CC.quntExc) from project_bill_details CC inner join project_billl On cc.bill_id =project_billl.id  "
        StrSQL = StrSQL + "  "
        StrSQL = StrSQL + "  Where cc.oprid =SubcontractorContract2.oprid             "
        StrSQL = StrSQL + "  and CC.project_id =SubcontractorContract2.project_id  "
        StrSQL = StrSQL + "  and project_billl.OrDer_no2=  " & val(TXTOrDer_no2) & ")"
                    
        StrSQL = StrSQL + "    FROM         dbo.project_bill_details LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.ProjectMainDes ON dbo.project_bill_details.PrMainDesID = dbo.ProjectMainDes.ID LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.TblProcessUnites ON dbo.project_bill_details.Unit_id = dbo.TblProcessUnites.UnitID"
        StrSQL = StrSQL + " Where dbo.project_bill_details.bill_id =" & val(Me.txtid.text)
        StrSQL = StrSQL + " order by project_bill_details.id"
    
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee, dbo.ProjectMainDes.FullCode, dbo.ProjectMainDes.Name,Accounts.Account_Serial,Accounts.Account_Name  ,"
        StrSQL = StrSQL + "              dbo.SubcontractorContract2.*,"
        
        StrSQL = StrSQL + "  Pre_Quantity_Contr = (Select Sum( CC.quntExc) from project_bill_details CC inner join project_billl On cc.bill_id =project_billl.id  "
        StrSQL = StrSQL + "  "
        StrSQL = StrSQL + "  Where cc.oprid =SubcontractorContract2.oprid             "
        StrSQL = StrSQL + "  and CC.project_id =SubcontractorContract2.project_id  "
        StrSQL = StrSQL + "  and project_billl.OrDer_no2=  " & val(TXTOrDer_no2) & ")"
                    
        StrSQL = StrSQL + "    FROM         dbo.SubcontractorContract2 LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.ProjectMainDes ON dbo.SubcontractorContract2.PrMainDesID = dbo.ProjectMainDes.ID LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.TblProcessUnites ON dbo.SubcontractorContract2.Unit_id = dbo.TblProcessUnites.UnitID"
        StrSQL = StrSQL + "              Left outer join Accounts On Accounts.Account_Code = SubcontractorContract2.AccountCode"
        StrSQL = StrSQL + " Where dbo.SubcontractorContract2.bill_id =" & val(Me.TXTOrDer_no.text)
        StrSQL = StrSQL + " order by SubcontractorContract2.id"
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or RsDev.EOF) Then
            'Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            'Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
            
            RsDev.MoveFirst
    
            With Me.Fg_Journal
                .rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .rows - 1
                    .TextMatrix(i, .ColIndex("ExPercen")) = 0  'IIf(IsNull(RsDev("ExPercen").value), 0, RsDev("ExPercen").value)
                    .TextMatrix(i, .ColIndex("PrMainDesID")) = IIf(IsNull(RsDev("PrMainDesID").value), 0, RsDev("PrMainDesID").value)
                    .TextMatrix(i, .ColIndex("CodeMain")) = IIf(IsNull(RsDev("FullCode").value), "", RsDev("FullCode").value)
                    .TextMatrix(i, .ColIndex("MainDes")) = "ff" 'IIf(IsNull(RsDev("Name").value), "", RsDev("Name").value)
                    .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(RsDev("qty").value), 0, RsDev("qty").value)
                    '.TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(RsDev("qtySubContractor").value), "", RsDev("qtySubContractor").value)

                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                    .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
                                       
                    '   .TextMatrix(i, .ColIndex("exe")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    .TextMatrix(i, .ColIndex("costSubContractor")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    '.TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    
                    .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("cost").value), "", RsDev("cost").value)
                    
                    ' .TextMatrix(i, .ColIndex("quntExc")) = IIf(IsNull(RsDev("quntExc").value), 0, RsDev("quntExc").value)
                    .TextMatrix(i, .ColIndex("quntExc")) = "" ' IIf(IsNull(RsDev("qtySubContractor").Value), 0, RsDev("qtySubContractor").Value)
                    .TextMatrix(i, .ColIndex("exe")) = IIf(IsNull(RsDev("exe").value), "", RsDev("exe").value)
                
                    .TextMatrix(i, .ColIndex("oprid")) = IIf(IsNull(RsDev("oprid").value), 0, RsDev("oprid").value)
                    .TextMatrix(i, .ColIndex("totEx")) = IIf(IsNull(RsDev("totEx").value), 0, RsDev("totEx").value)
                
                    .TextMatrix(i, .ColIndex("net")) = IIf(IsNull(RsDev("net").value), 0, RsDev("net").value)
                    .TextMatrix(i, .ColIndex("discount")) = IIf(IsNull(RsDev("discount").value), 0, RsDev("discount").value)
                    .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), 0, RsDev("total").value)
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
                    
                    .TextMatrix(i, .ColIndex("Pre_Quantity_Contr")) = IIf(IsNull(RsDev("Pre_Quantity_Contr").value), "", RsDev("Pre_Quantity_Contr").value)
                    
                    .TextMatrix(i, .ColIndex("StillQty")) = val(RsDev!qtySubContractor & "") - val(RsDev!Pre_Quantity_Contr & "")
                    
                    .TextMatrix(i, .ColIndex("qtySubContractor")) = IIf(IsNull(RsDev("qtySubContractor").value), "", RsDev("qtySubContractor").value)
                    .TextMatrix(i, .ColIndex("costSubContractor")) = IIf(IsNull(RsDev("costSubContractor").value), "", RsDev("costSubContractor").value)
                    
                    If val(.TextMatrix(i, .ColIndex("qtySubContractor"))) = 0 Then
                        .TextMatrix(i, .ColIndex("qtySubContractor")) = .TextMatrix(i, .ColIndex("qty"))
                    End If
                    ' .TextMatrix(i, .ColIndex("StillQty")) = .TextMatrix(i, .ColIndex("qtySubContractor"))
                    
                    If val(.TextMatrix(i, .ColIndex("costSubContractor"))) = 0 Then
                        .TextMatrix(i, .ColIndex("costSubContractor")) = .TextMatrix(i, .ColIndex("exe"))
                    End If
                                
                    .TextMatrix(i, .ColIndex("OLDTotalwithVat")) = IIf(IsNull(RsDev("OLDTotalwithVat").value), 0, RsDev("OLDTotalwithVat").value)
                    .TextMatrix(i, .ColIndex("CurrenttotalWithvat")) = IIf(IsNull(RsDev("CurrenttotalWithvat").value), 0, RsDev("CurrenttotalWithvat").value)
                    .TextMatrix(i, .ColIndex("Totalwitvat")) = IIf(IsNull(RsDev("Totalwitvat").value), 0, RsDev("Totalwitvat").value)
                                  
                    lbl(9).Caption = IIf(IsNull(RsDev("oldPerforValue").value), 0, RsDev("oldPerforValue").value)
                    lbl(11).Caption = IIf(IsNull(RsDev("totalPerforValue").value), 0, RsDev("totalPerforValue").value)
                    
                    '
           
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
            ReLineGrid
        End If

    End If
       
    Exit Sub

Exits:

    TXTOrDer_no = ""
    TXTOrDer_no2 = ""

End Sub

Private Function GetPre_Quantity_Contr() As Double
    Dim s       As String
    Dim rsDummy As New ADODB.Recordset
    s = "Select Sum()"
    
End Function

Private Sub TXTOrDer_no2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Order_no_search4.show
        Order_no_search4.RetrunType = 83
        Order_no_search4.Label1(2).Caption = "«·»ÕÀ ⁄‰ ⁄ﬁÊœ „ﬁ«Ê·Ì «·»«ÿ‰"

        'If val(Me.DBCboClientName.BoundText) <> 2 Then
        
        '    Order_no_search4.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
        'End If
    End If

End Sub

Private Sub txtPerformanceBond_Change()
 calcnet
    ReLineGrid
End Sub

Private Sub TxtPeriod_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.txtPeriod.text, 0)
End Sub

Private Sub TxtPreBalaTransPyed_Change()
    ClculteVATBalan
    SumVAT
End Sub

Private Sub TxtPreBalaTransPyed_LostFocus()
    If val(TxtPreBalaTransPyed.text) > val(TxtPreBalaRemain.text) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬂ»— „‰ «·„ »ﬁÌ"
        Else
            MsgBox "Paid value is greater than remaining"
        End If
        TxtPreBalaTransPyed.text = 0
        Exit Sub
    End If
    ClculteVATBalan
    SumVAT
End Sub
Sub ClculteVATBalan()
    Dim Percetage2 As Double
    TxtPreBalaNet.text = val(TxtPreBalaRemain.text) - val(TxtPreBalaTransPyed.text)

End Sub
Function GetBalanceProject(Optional ByRef Valu As Double, _
                           Optional ByRef RecDate As Date) As Boolean
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
    sql = sql & " From dbo.project_billl"
    sql = sql & " WHERE     (id <> " & val(txtid.text) & ") AND (project_no = N'" & DataCombo2.BoundText & "')"
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs2.RecordCount > 0 Then
        GetValue = IIf(IsNull(rs2("SumValue").value), 0, rs2("SumValue").value)
    Else
        GetValue = 0
    End If
End Function

Private Sub TxtPreVAT_Change()
    TxtPreVAT2 = TxtPreVAT
End Sub

Private Sub txtprojectname_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmProjectSearch.lblSearchtype.Caption = 8
        FrmProjectSearch.show vbModal
    End If
End Sub

Private Sub TxtTotalValue_Change()
    CalcFormat
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, _
   ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg            As String
    Dim rs             As New ADODB.Recordset
    Dim Rs1            As New ADODB.Recordset
    Dim StrSQL         As String
    Dim ClsAcc         As New ClsAccounts
    Dim LngRow         As Long

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
    
        If Row = .rows - 1 Then
            .rows = .rows + 1
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
   
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If
 
    End With

    ReLineGrid

End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    Dim rs             As New ADODB.Recordset
    Dim Rs1            As New ADODB.Recordset
    Dim StrSQL         As String
    Dim StrAccountType As String
    Dim StrComboList   As String
    Dim Msg            As String

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

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "id='" & val(txtid.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub VSFlexGrid4_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    RelineBuy
    RelineBu22
End Sub

Private Sub VSFlexGrid4_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)
  
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

        If VSFlexGrid4.ColIndex("TransPayedValue") = Col And VSFlexGrid4.cell(flexcpChecked, Row, 4) <> 2 Then
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

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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
    TxtNoteSerial.text = ""
    If Me.TxtModFlg.text <> "R" Then
        If ChekSanNumber(Current_branch, 65) = True Then
            TxtNoteSerial1.text = ""
        End If
        TxtNoteSerial.text = ""
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
        txtRemarks.SetFocus
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

Private Sub CBoBasedON_Change()
    If CBoBasedON.ListIndex = 0 Then
    
        TXTOrDer_no2.Visible = False
        billto.Enabled = True
        DcbosubContractor.text = ""
        DcbosubContractor.Enabled = True
        Text2.Enabled = True
    Else
        billto.Enabled = False
        DcbosubContractor.Enabled = False
        Text2.Enabled = False
    
        TXTOrDer_no2.Visible = True
   
    End If
    
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        
        If TXTOrDer_no2.text <> "" Then
            '     TXTOrDer_no(0).Text = ""
            '     TXTOrDer_no(1).Text = ""
        End If
        
    End If
End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Sub Savetemp()
    Dim i   As Long
  Dim LineDiscountPercent               As Double
    Dim LineDiscount                      As Double
    Dim linenetaftermainDiscount          As Double
    Dim linenetaftermainDiscountBeforevat As Double
    Dim LineVat                           As Double
    Dim linenetaftermainDiscountWithvat   As Double
    Dim OLDTotalwithVat                   As Double
    Dim CurrenttotalWithvat               As Double
    Dim Totalwitvat                       As Double
    Dim oldPerforValue                    As Double
    Dim totalPerforValue                  As Double

    Dim Rs3 As New ADODB.Recordset
    '   Rs3.Open "project_bill_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    
       
     StrSQL = "sELECT * FROM TBLProjectBillHistory Where 1 = -1"
   
    saveGrid StrSQL, GrdBondHistory, "Serial", "", "bill_id", val(Me.txtid.text)
     
     
    
    
   StrSQL = " delete dbo.project_bill_details where bill_id = " & val(Me.txtid.text)
   Cn.Execute StrSQL
    StrSQL = "SELECT     * from dbo.project_bill_details Where (1 = -1)"
    Rs3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Fg_Journal

        For i = .FixedRows To .rows - 1

            '        Dim IntDEV_Type As Integer
            '        Dim SngDEV_Value As Single
            If .TextMatrix(i, .ColIndex("item")) <> "" Then

                Rs3.AddNew
                Rs3("bill_id").value = Me.txtid.text
                
                Rs3("FullCode").value = Trim(.TextMatrix(i, .ColIndex("FullCode")))
                Rs3("project_id").value = IIf(.TextMatrix(i, .ColIndex("project_id")) = "", Null, val(.TextMatrix(i, .ColIndex("project_id"))))
                Rs3("projectName").value = .TextMatrix(i, .ColIndex("projectName"))
                Rs3("AccountCode").value = .TextMatrix(i, .ColIndex("AccountCode"))

                Rs3("ExPercen").value = IIf(.TextMatrix(i, .ColIndex("ExPercen")) = "", Null, val(.TextMatrix(i, .ColIndex("ExPercen"))))
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
                LineDiscountPercent = val(.TextMatrix(i, .ColIndex("NetExe"))) / Results.text
                
                LineDiscount = (val(txtDiscountG.text)) * LineDiscountPercent
         
                PerforVLineDiscount = val(TxtPerforValue.text) * LineDiscountPercent
               
                linenetaftermainDiscount = val(.TextMatrix(i, .ColIndex("NetExe"))) - LineDiscount
                
                Rs3("LineDiscountPercent").value = LineDiscountPercent
                Rs3("LineDiscount").value = LineDiscount
                Rs3("PerforVLineDiscount").value = PerforVLineDiscount
                linenetaftermainDiscountBeforevat = val(.TextMatrix(i, .ColIndex("NetExe"))) - LineDiscount
                  
                Rs3("linenetaftermainDiscountBeforevat").value = linenetaftermainDiscountBeforevat
                 
                LineVat = Rs3("linenetaftermainDiscountBeforevat").value * val(TxtFATYou.text) / 100
                
                Rs3("LineVat").value = LineVat
                linenetaftermainDiscountWithvat = linenetaftermainDiscount + LineVat
                 
                Rs3("linenetaftermainDiscountWithvat").value = linenetaftermainDiscountWithvat
                '     Rs3("linenetaftermainDiscountWithvat").value = linenetaftermainDiscountWithvat - PerforVLineDiscount
                Rs3("LineFinal").value = linenetaftermainDiscountWithvat - PerforVLineDiscount
                  
                'newwwwwwwwwwwwwwwwwwwwwwww
                Rs3("QtyApprov").value = IIf(.TextMatrix(i, .ColIndex("QtyApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("QtyApprov"))))
                Rs3("TotalApprov").value = IIf(.TextMatrix(i, .ColIndex("TotalApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("TotalApprov"))))
                Rs3("PriceApprov").value = IIf(.TextMatrix(i, .ColIndex("PriceApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("PriceApprov"))))
                Rs3("DiscApprov").value = IIf(.TextMatrix(i, .ColIndex("DiscApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("DiscApprov"))))
                Rs3("NetApprov").value = IIf(.TextMatrix(i, .ColIndex("NetApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("NetApprov"))))
                '''//////////////
                Rs3("discountEXE").value = IIf(.TextMatrix(i, .ColIndex("discountEXE")) = "", Null, val(.TextMatrix(i, .ColIndex("discountEXE"))))
                
                Rs3("NetExe").value = IIf(.TextMatrix(i, .ColIndex("NetExe")) = "", Null, val(.TextMatrix(i, .ColIndex("NetExe"))))

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
            If val(TXTOrDer_no2.text) <> 0 And CBoBasedON.ListIndex = 1 Then
            
                UpdatePre_QuantityCont i
            End If
        Next i

    End With

    updateNotesValueAndNobytext val(note_id.text)
    saveBillBuy
                
    If SystemOptions.SuppCreat4Acc Then
                
    SaveQRCode "project_billl", "ID", val(txtid), TxtNoteSerial1.text, (XPDtbTrans.value), _
       (TxtTotalValue.text), Picture1, 0, (TxtFATValue.text), (TxtTotalValue.text)
    Else
        SaveQRCode6 "project_billl", "ID", val(txtid), TxtNoteSerial1.text, (XPDtbTrans.value), _
       val(Results.text) + val(TxtFATValue) - val(TxtPreVAT.text) - val(advancedPayment), Picture1, 0, val(TxtFATValue.text) - val(TxtPreVAT.text), val(Results.text) + val(TxtFATValue) - val(TxtPreVAT.text) - val(advancedPayment), val(dcBranch.BoundText)
       
       
       
       ' SaveQRCode6 "tblEInvoice", "ID", val(RsData!ID & ""), Trim(RsData!invoiceID & ""), RsData!IssueDate & "", _
       ' val(RsData!PayableAmount & ""), Picture1, 0, val(RsData!VATValue & ""), val(RsData!PayableAmount & ""), BranchID
       
       '{adox.advancedPayment}-{adox.prevat}
       'Results.text
    End If
    TxtModFlg.text = "R"
    fillapprovData
       
    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Saved", vbInformation
    Else
        MsgBox " „ Õ›Ÿ «·»Ì«‰« ", vbInformation
  
    End If
     Set Rs3 = New ADODB.Recordset
    '   Rs3.Open "project_bill_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     * from dbo.project_bill_details Where (1 = -1)"
    Rs3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
End Sub


Function SaveDataOld()
    calcnet
    Dim accountdep As String

    If billto.ListIndex = -1 Then MsgBox "Õœœ «·„” Œ·’ „ﬁœ„  «·Ï „‰ ", vbCritical: Exit Function
        
   If billto.ListIndex = 1 And DcbosubContractor.BoundText = "" Then MsgBox "Õœœ „ﬁ«Ê· «·»«ÿ‰  ", vbCritical: Exit Function
   
        
     Dim j As Integer
     Dim found As Boolean
     
   j = Fg_Journal.FixedRows
     found = False
     
  For j = Fg_Journal.FixedRows To Fg_Journal.rows - 1
  If Fg_Journal.TextMatrix(j, Fg_Journal.ColIndex("item")) <> "" Then
        found = True
  End If
Next

If found = False Then

MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„ ›Ï «·›« Ê—… ", vbCritical: Exit Function

End If

        
        
        
        
    If billto.ListIndex = 0 Then
   X = val(TXTEnd_user_id.text)
        'accountdep = txtendaccount.text
    Else

        If billto.ListIndex = 1 Then
        X = val(TXTsub_contractor_id.text)
        '    accountdep = txtsubaccount.text
        End If
    End If
X = val(TXTEnd_user_id.text)
  '  Dim x As Double
  '  x = get_Customer_id(accountdep)
        
    '  total.text = gettotal(txtid.text)
    Dim Rs1 As New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=5000 and NoteSerial='" & Me.TxtNoteSerial.text & "' order by NoteID"
    Rs1.Open StrSQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
  
    If TxtModFlg.text = "N" Then
   
        If X = 0 Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "An error in customer Number", vbCritical: Exit Function
            Else
                MsgBox "ÌÊÃœ Œÿ√ ›Ì —ﬁ„ «·⁄„Ì·", vbCritical: Exit Function
            End If
        End If
          note_id.text = CStr(new_id("Notes", "NoteID", "", True))
            txtid.text = CStr(new_id("project_billl", "id", "", True))
            
        rs.AddNew
    
    Else
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(note_id.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From notes  Where NoteID=" & val(note_id.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        StrSQL = "Delete From project_bill_details Where bill_id=" & val(Me.txtid.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Rs1.AddNew
 'branch_id
'     If TxtNoteSerial1.text = "" Then
'     TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 65, 65, , , , , val(billto.ListIndex))
'     End If
     
        If TxtNoteSerial1.text = "" Then
                If billto.ListIndex = 0 Then
                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 65, 65, , , , , val(billto.ListIndex))
                Else
                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 84, 84, , , , , val(billto.ListIndex))
                End If
        End If
     
    Rs1("branch_no").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    Rs1("NoteID").value = val(note_id.text)
    Rs1("Note_Value").value = IIf(total.text = "", Null, val(total.text))
    Rs1("CusID").value = X
    Rs1("NoteType").value = 500
    Rs1("NoteType").value = 5000
    Rs1("NoteDate").value = XPDtbTrans.value
    Rs1("UserID").value = user_id
   Rs1("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", Trim(TxtNoteSerial1.text), Null)






'
'    Rs1("CBoBasedON").Value = CBoBasedON.ListIndex
'    Rs1.Fields("OrDer_no2").Value = Trim(TXTOrDer_no2.Text)
'    Rs1.Fields("OrDer_no").Value = val(TXTOrDer_no.Text)
'
   Rs1("RemarkE").value = IIf(Me.txtRemarks <> "", Trim(txtRemarks.text), Null)
   Rs1("Remark").value = IIf(Me.txtRemarks <> "", Trim(txtRemarks.text), Null)
   
    rs("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", Trim(TxtNoteSerial1.text), Null)
    rs("ExPercen").value = val(TxtExPercen.text)
    rs("ExPercenID").value = val(DcbExPercen.ListIndex)
    rs("PreVAT").value = val(TxtPreVAT.text)
    rs("FATYou").value = val(TxtFATYou.text)
    rs("FATValue").value = val(TxtFATValue.text)
    rs("TotalValue").value = val(TxtTotalValue.text)
    rs("AccountCodeVat").value = Me.AccountVat.BoundText
    rs("NetValue").value = val(TxtNetValue.text)
    rs("PerforValue").value = val(TxtPerforValue.text)
    ''/////////
    rs("StartDateProje").value = StartDateProje.value
    rs("PreBalaValue").value = val(TxtPreBalaValue.text)
    rs("PreBalaVAT").value = val(TxtPreBalaVAT.text)
    rs("PreBalaTotal").value = val(TxtPreBalaTotal.text)
    rs("PreBalaPayed").value = val(TxtPreBalaPayed.text)
    rs("PreBalaRemain").value = val(TxtPreBalaRemain.text)
    rs("PreBalaTransPyed").value = val(TxtPreBalaTransPyed.text)
    rs("PreBalaNet").value = val(TxtPreBalaNet.text)
    rs("PreBalaVATYu").value = val(TxtPreBalaVATYu.text)
    rs("SumVATLine").value = val(Label57.Caption)
    rs("SumValueLine").value = val(Label56.Caption)
    ''/////
 If Option7.value = True Then
  rs("UnderImp").value = 0
 ElseIf Option6.value = True Then
  rs("UnderImp").value = 1
 ElseIf Option8.value = True Then
  rs("UnderImp").value = 2
End If
If TxtManualNO.text = "" Then
TxtManualNO.text = TxtNoteSerial1.text
End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Rs1("remark").value = "„” Œ·’ —ﬁ„  :  " & TxtManualNO & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text
    Else
        Rs1("remark").value = "  Project Invoice No  :  " & TxtManualNO & CHR(13) & "  To Project " & txtprojectname.text
    End If
 
    '   Rs1("remark").value = "„” Œ·’ —ﬁ„ :     " & txtid & "    " & Chr(13) & "  ··„‘—Ê⁄  " & txtprojectname.text
    
    If TxtNoteSerial = "" Then
        TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
    End If
       
    Rs1("NoteSerial").value = TxtNoteSerial.text
    
    Rs1("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”·
    Rs1("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    '  rs("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
     
    Rs1("sanad_year").value = year(XPDtbTrans.value)
    Rs1("sanad_month").value = Month(XPDtbTrans.value)
    Rs1("note_value_by_characters").value = WriteNo(Format(Me.Results.text, "0.00"), 0, True, ".")
    
    Rs1.update
    
    rs("id").value = Me.txtid.text
    
    rs("bill_date").value = XPDtbTrans.value
  'branch_id
    rs("branch_no").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
    rs("project_no").value = IIf(Not IsNumeric(DataCombo2.BoundText), "", DataCombo2.BoundText)
    rs("project_name").value = txtprojectname.text
    rs("Sub_user_name").value = IIf(IsNull(DcAccount1.text), "", DcAccount1.text)
    rs("End_user_name").value = IIf(IsNull(DcAccount2.text), "", DcAccount2.text)
    rs("End_user_account").value = IIf(IsNull(txtendaccount.text), "", txtendaccount.text)
    rs("Sub_user_account").value = IIf(IsNull(txtsubaccount.text), "", txtsubaccount.text)
    rs("revenue_account").value = IIf(IsNull(txtrevenue_account.text), "", txtrevenue_account.text)
    rs("UserID").value = IIf(DCboUserName.BoundText <> "", val((DCboUserName.BoundText)), Null)
    rs("bill_to").value = billto.ListIndex
    rs("bill_type").value = bill_Type.ListIndex '  IIf(IsNull(bill_Type.text), "", bill_Type.text)
    rs("note_id").value = IIf(IsNull(note_id.text), "", note_id.text)
    rs("NoteSerial").value = IIf(IsNull(TxtNoteSerial.text), "", TxtNoteSerial.text)
    rs("total").value = IIf(Not IsNumeric(total.text), 0, total.text)
    rs("AccountUnderImp").value = TxtAccountUnderImp.text
    
    
    rs("CBoBasedON").value = CBoBasedON.ListIndex
    rs.Fields("OrDer_no2").value = val(TXTOrDer_no2.text)
    rs.Fields("OrDer_no").value = val(TXTOrDer_no.text)
    
    '26082015
    rs("Discount").value = IIf(Not IsNumeric(txtDiscount.text), 0, txtDiscount.text)
    rs("PerformanceBond").value = IIf(Not IsNumeric(txtPerformanceBond.text), 0, txtPerformanceBond.text)
    
    rs("AdvancedPayment").value = IIf(Not IsNumeric(advancedPayment.text), 0, advancedPayment.text)
    
    rs("Results").value = IIf(Not IsNumeric(Results.text), 0, Results.text)
   ''///////23 05 2016
  rs("BillNo").value = val(TxtBillNo.text)
  rs("StartDate").value = StartDate.value
  rs("Period").value = val(txtPeriod.text)
  rs("PeriodType").value = val(DcbPeriodType.ListIndex)
  rs("Remarks2").value = TxtRemarks2.text
  
 
'26082015


         rs("dueDate").value = dueDate.value
rs("dueDate1").value = dueDate1.value


'*************************************************
rs("subContractorId").value = IIf(Not IsNumeric(DcbosubContractor.BoundText), Null, DcbosubContractor.BoundText)
rs("discount1ID").value = val(cboDiscount1.ListIndex)
rs("discount2ID").value = val(cboDiscount2.ListIndex)
rs("discount1value").value = val(txtDiscount1.text)
rs("discount2value").value = val(txtDiscount2.text)
rs("Remarks").value = Trim(txtRemarks.text)
rs("ManualNo").value = Trim(TxtManualNO.text)

 
'*************************************************
If val(Me.TxtBillNo.text) > 0 Then
    SaveBillMonthly
End If

    rs.update

    
    Set RsDev = New ADODB.Recordset
 '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
               StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   Dim LngDevID As Long
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
  accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", X, "Account_code")

 Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If

'If Label11(6).Tag <> "Posted" And Label11(6).Tag <> "1" Then
'    GoTo AfterEntry
'End If
If billto.ListIndex = 0 Then
  Dim lineno As Integer
  lineno = 1
'    If accountdep = "" Then GoTo ll
    '«·ÿ—› «·„œÌ‰
    RsDev.AddNew
    
    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
If Option8.value = True Then
    RsDev("Account_Code").value = Me.TxtAccountUnderImp.text
 Else
    RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
End If
        RsDev("Value").value = val(Me.total.text) + val(TxtFATValue.text)
    RsDev("Credit_Or_Debit").value = 0

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & " Manual #   " & TxtManualNO & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
     RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    
    RsDev.update
'll:
lineno = lineno + 1

'«·Õ”„Ì« 
'Account_Code_dynamic1
If val(Me.txtDiscount.text) > 0 Then
    RsDev.AddNew
    
    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = Account_Code_dynamic1 '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
    RsDev("Value").value = val(Me.txtDiscount.text)
    RsDev("Credit_Or_Debit").value = 0

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  Manual# " & TxtManualNO & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
    
   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
   
   
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
   RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    RsDev.update
'll:
lineno = lineno + 1

End If
If val(Me.TxtPerforValue.text) > 0 Then
    RsDev.AddNew
    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = AcountGood ' get_account_code_branch(152, my_branch) 'Õ”«» Õ”‰ «·«œ«¡
    RsDev("Value").value = val(Me.TxtPerforValue.text)
    RsDev("Credit_Or_Debit").value = 0
    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & " Manual#  " & TxtManualNO & CHR(13) & txtRemarks.text
    End If
    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    RsDev.update
'll:
'lineno = lineno + 1
'
'    RsDev.AddNew
'    RsDev("branch_id").value = IIf(Trim$(Me.DcBranch.BoundText) = "", Null, Trim$(Me.DcBranch.BoundText))
'    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'    RsDev("DEV_ID_Line_No").value = lineno
'    RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
'    RsDev("Value").value = val(Me.TxtPerforValue.Text)
'    RsDev("Credit_Or_Debit").value = 1
'    If SystemOptions.UserInterface = ArabicInterface Then
'        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & txtid & Chr(13) & "  ··„‘—Ê⁄ " & txtprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNo
'    Else
'        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & txtid & Chr(13) & "  To Project " & txtprojectname.Text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNo
'    End If
'    RsDev("Notes_ID").value = val(note_id.Text)
'    RsDev("project_bill_no").value = val(txtid.Text)
'   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
'    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
'    RsDev("UserID").value = user_id
'    RsDev("branch_id").value = my_branch
'    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'    RsDev.update
'll:
lineno = lineno + 1

End If


'«·œ›⁄«  «·„ﬁœ„…
'Account_Code_dynamic2
If val(Me.advancedPayment.text) > 0 Then
    RsDev.AddNew
    
    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = Account_Code_dynamic2 '    Õ”«» œ›⁄«  „ﬁœ„…
    RsDev("Value").value = val(Me.advancedPayment.text)
    RsDev("Credit_Or_Debit").value = 0

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  Manuall#  " & TxtManualNO & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    RsDev.update
'll:
lineno = lineno + 1

  RsDev.AddNew
    
    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = accountdep '    Õ”«» œ›⁄«  „ﬁœ„…
    RsDev("Value").value = val(Me.advancedPayment.text)
    RsDev("Credit_Or_Debit").value = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "   Manual   " & TxtManualNO & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
     RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    RsDev.update
'll:
lineno = lineno + 1

End If


'«·«Ì—œ« 

    '«·ÿ—› «·œ«∆‰
   If Option8.value = False Then
    If Me.txtrevenue_account.text = "" Then Exit Function
    
   Else
   If (accountdep = "") Then Exit Function
  End If
    RsDev.AddNew
    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
    RsDev("branch_id").value = my_branch
    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
 'If SystemOptions.Revenueowed = True Then
 If Option8.value = True Then
 RsDev("Account_Code").value = accountdep
 Else
 ''
 RsDev("Account_Code").value = Me.txtrevenue_account.text ' Account_Code_dynamic1
 
 End If
 '   Else
    'RsDev("Account_Code").value = Me.txtrevenue_account .text
 '   End If
    
    RsDev("Value").value = val(Me.Results.text)  '«·«Ì—«œ« 
    RsDev("Credit_Or_Debit").value = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
 '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
 Else
   RsDev("project_id").value = val(Me.DataCombo2.BoundText)
 End If
 
    RsDev("RecordDate").value = XPDtbTrans.value
    RsDev("UserID").value = user_id
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    RsDev.update
lineno = lineno + 1
'///////////////////
If Me.AccountVat.BoundText <> "" And val(Me.TxtFATValue.text) > 0 Then
    RsDev.AddNew
    RsDev("Account_Code").value = Me.AccountVat.BoundText
    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
    RsDev("branch_id").value = my_branch
    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Value").value = val(Me.TxtFATValue.text)  '«·«Ì—«œ« 
    RsDev("Credit_Or_Debit").value = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & txtid & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " Õ”«» «·ﬁÌ„… «·„÷«›…" & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "Account VAT" & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
    If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
     '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
     Else
     '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
     End If
 
    RsDev("RecordDate").value = XPDtbTrans.value
    RsDev("UserID").value = user_id
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
     
    RsDev.update
    lineno = lineno + 1
''///////////////////////›«  «·œ›⁄«  «·„ﬁœ„…
    If val(TxtPreVAT.text) > 0 Then
        RsDev.AddNew
        RsDev("Account_Code").value = Me.AccountVat.BoundText
        RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
        RsDev("branch_id").value = my_branch
        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
        RsDev("DEV_ID_Line_No").value = lineno
        RsDev("Value").value = val(Me.TxtPreVAT.text)  '«·«Ì—«œ« 
        RsDev("Credit_Or_Debit").value = 0
        If SystemOptions.UserInterface = ArabicInterface Then
            RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " Õ”«» «·ﬁÌ„… «·„÷«›… œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
        Else
            RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "Account VAT" & CHR(13) & txtRemarks.text
        End If
    
        RsDev("Notes_ID").value = val(note_id.text)
        RsDev("project_bill_no").value = val(txtid.text)
        If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
         '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
         Else
        '   RsDev("project_id").value = val(Me.DataCombo2.BoundText) 'xxxxxxxxxxxxxx
         End If
 
        RsDev("RecordDate").value = XPDtbTrans.value
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
        RsDev.update
        lineno = lineno + 1
''//////////////
        RsDev.AddNew
        If Option8.value = True Then
            RsDev("Account_Code").value = Me.TxtAccountUnderImp.text
        Else
            RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
        End If
        RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))
        RsDev("branch_id").value = my_branch
        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
        RsDev("DEV_ID_Line_No").value = lineno
        RsDev("Value").value = val(Me.TxtPreVAT.text)  '«·«Ì—«œ« 
        RsDev("Credit_Or_Debit").value = 1
        If SystemOptions.UserInterface = ArabicInterface Then
            RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " Õ”«» «·ﬁÌ„… «·„÷«›… œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
        Else
            RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "Account VAT" & CHR(13) & txtRemarks.text
        End If

        RsDev("Notes_ID").value = val(note_id.text)
        RsDev("project_bill_no").value = val(txtid.text)
        If SystemOptions.Revenueowed = True Then '«Ì—«œ«  „” Õﬁ…
         '  RsDev("project_id").value = val(Me.DataCombo2.BoundText)
         Else
        '   RsDev("project_id").value = val(Me.DataCombo2.BoundText) 'xxxxxxxxxxxxxx
         End If
 RsDev("project_id").value = val(Me.DataCombo2.BoundText)
        RsDev("RecordDate").value = XPDtbTrans.value
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
        RsDev.update
        lineno = lineno + 1

    End If
End If
''/////////
Else
'
        'If SystemOptions.SubContactorHave3Account = True Then
                Dim Discount1 As Double
                Dim Discount2 As Double
                Dim netvalue As Double
                Dim TotalValue As Double
                Dim AdvancedAccount As String
                Dim GuranteeAccount As String
                Dim line_no As Integer
                Dim des As String
                            If cboDiscount1.ListIndex = 0 Then
                                Discount1 = 0
                            ElseIf cboDiscount1.ListIndex = 1 Then
                                Discount1 = val(txtDiscount1) * val(Me.TxtNetValue.text) / 100
                            ElseIf cboDiscount1.ListIndex = 2 Then
                                Discount1 = val(txtDiscount1)
                            End If
        
                            If cboDiscount2.ListIndex = 0 Then
                                Discount2 = 0
                            ElseIf cboDiscount2.ListIndex = 1 Then
                                Discount2 = val(txtDiscount2) * val(TxtNetValue.text) / 100
                            ElseIf cboDiscount2.ListIndex = 2 Then
                                Discount2 = val(txtDiscount2)
                            End If
               AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code2")
               GuranteeAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code1")
               accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code")
               
               line_no = 1
               If SystemOptions.AllowNoRoudProjectInvoices = True Then
                Discount1 = Round(Discount1, val(cCompanyInfo.NoRoudProjectInvoices))
                Discount2 = Round(Discount2, val(cCompanyInfo.NoRoudProjectInvoices))
               netvalue = Round(val(TxtNetValue.text) - Discount1 - Discount2, val(cCompanyInfo.NoRoudProjectInvoices))
               TotalValue = Round(val(TxtNetValue), val(cCompanyInfo.NoRoudProjectInvoices))
               Else
               Discount1 = Round(Discount1, 2)
                Discount2 = Round(Discount2, 2)
               netvalue = Round(val(TxtNetValue.text) - Discount1 - Discount2, 2)
               TotalValue = Round(val(TxtNetValue), Decimal_Places)
              End If
               If Option8.value = True Then
                              des = "„’—Ê›«  «·„‘«—Ì⁄ " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
                         Else
                         des = "„” Œ·’ «·„‘«—Ì⁄ " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & "Manual#" & TxtManualNO & CHR(13) & txtRemarks.text
                   End If
           If TotalValue > 0 Then '
    
                
            
            If Option8.value = True Then
               If ModAccounts.AddNewDev(LngDevID, line_no, TxtAccountUnderImp.text, TotalValue, 0, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , DataCombo2.BoundText, , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
                   Else
                    If ModAccounts.AddNewDev(LngDevID, line_no, expanses_account, TotalValue, 0, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , DataCombo2.BoundText, , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
                   End If
                 If val(TxtFATValue.text) > 0 Then
                  If ModAccounts.AddNewDev(LngDevID, line_no, Me.AccountVat.BoundText, val(TxtFATValue.text), 0, Msg & "  " & "    " & txtprojectname.text & "    VAT  " & "   INV# " & TxtNoteSerial1.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
                    End If
              '  End If
  '////////////////////////////////////////////////////
        ' End If
         
      If SystemOptions.UserInterface = ArabicInterface Then
               des = "Œ’„ ÷„«‰ «⁄„«· " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & " ··„” Œ·’ —ﬁ„ " & TxtNoteSerial1.text & CHR(13) & txtRemarks.text
       Else
                      des = " Discount " & "   " & txtRemarks & "Inv# " & TxtNoteSerial1 & " manual#" & TxtManualNO & CHR(13) & txtRemarks.text
       
       End If
           If Discount1 > 0 Then '÷„«‰ «·«⁄„«·
    
                If GuranteeAccount = "" Then
                GuranteeAccount = accountdep
                End If
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, GuranteeAccount, Discount1, 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
            
  
         End If
       If SystemOptions.UserInterface = ArabicInterface Then
         des = "Œ’„ œ›⁄«  „ﬁœ„…   " & "   " & txtRemarks & " —ﬁ„ «·”‰œ " & TxtNoteSerial1 & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & CHR(13) & txtRemarks.text
        Else
        des = "Advance Discount " & "   " & txtRemarks & "InvE " & TxtNoteSerial1 & "Manuall#   " & TxtManualNO & CHR(13) & txtRemarks.text
        End If
           If Discount2 > 0 Then '
    
               If AdvancedAccount = "" Then
                AdvancedAccount = accountdep
                End If
            
               If ModAccounts.AddNewDev(LngDevID, line_no, AdvancedAccount, Discount2, 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , DataCombo2.BoundText, , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
              
  
         End If
         
        '«·œ›⁄«  «·„ﬁœ„…
'Account_Code_dynamic2
LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
If val(Me.advancedPayment.text) > 0 Then
    RsDev.AddNew
    
    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = accountdep '    Õ”«» œ›⁄«  „ﬁœ„…
    RsDev("Value").value = val(Me.advancedPayment.text) + val(Me.TxtPreVAT.text)
    RsDev("Credit_Or_Debit").value = 0

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "Manual" & TxtManualNO & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    RsDev.update
'll:
lineno = lineno + 1

  RsDev.AddNew

    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = AdvancedAccount   '    Õ”«» œ›⁄«  „ﬁœ„…
    RsDev("Value").value = val(Me.advancedPayment.text)
    RsDev("Credit_Or_Debit").value = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & "  manual  " & TxtManualNO & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
 '   RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    RsDev.update
'll:
lineno = lineno + 1




  RsDev.AddNew

    RsDev("branch_id").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = Me.AccountVat.BoundText   '    Õ”«» œ›⁄«  „ﬁœ„…
    RsDev("Value").value = val(Me.TxtPreVAT.text)
    RsDev("Credit_Or_Debit").value = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & TxtNoteSerial1 & CHR(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & TxtManualNO & "œ›⁄«  „ﬁœ„…" & CHR(13) & txtRemarks.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & TxtNoteSerial1 & CHR(13) & "  To Project " & txtprojectname.text & " Manual" & TxtManualNO & CHR(13) & txtRemarks.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
 '   RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev("Posted").value = IIf(Posted = 0, Null, Posted)
    RsDev.update

End If

 
        If SystemOptions.UserInterface = ArabicInterface Then
         des = " «⁄„«·" & " ··„” Œ·’ —ﬁ„ " & TxtNoteSerial1.text & CHR(13) & txtRemarks.text
         Else
         des = " Works" & "  Inv#" & TxtNoteSerial1.text & CHR(13) & txtRemarks.text
         End If
         
           If netvalue > 0 Then '
    
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                 If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, netvalue + val(TxtFATValue.text), 1, Msg & des & "  " & "    " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
            
  
         End If
         


End If
End If



AfterEntry:
    Dim Rs3 As New ADODB.Recordset
 '   Rs3.Open "project_bill_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
               StrSQL = "SELECT     * from dbo.project_bill_details Where (1 = -1)"
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

        For i = .FixedRows To .rows - 1

            '        Dim IntDEV_Type As Integer
            '        Dim SngDEV_Value As Single
            If .TextMatrix(i, .ColIndex("item")) <> "" Then

                Rs3.AddNew
                Rs3("bill_id").value = Me.txtid.text
                
                Rs3("FullCode").value = Trim(.TextMatrix(i, .ColIndex("FullCode")))
                Rs3("project_id").value = IIf(.TextMatrix(i, .ColIndex("project_id")) = "", Null, val(.TextMatrix(i, .ColIndex("project_id"))))
                Rs3("projectName").value = .TextMatrix(i, .ColIndex("projectName"))
                

                Rs3("ExPercen").value = IIf(.TextMatrix(i, .ColIndex("ExPercen")) = "", Null, val(.TextMatrix(i, .ColIndex("ExPercen"))))
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
                LineDiscountPercent = val(.TextMatrix(i, .ColIndex("NetExe"))) / Results.text
                
                 LineDiscount = (val(txtDiscountG.text)) * LineDiscountPercent
         
                 PerforVLineDiscount = val(TxtPerforValue.text) * LineDiscountPercent
                 
               
                 linenetaftermainDiscount = val(.TextMatrix(i, .ColIndex("NetExe"))) - LineDiscount
                 
                
                
                
                
                
                Rs3("LineDiscountPercent").value = LineDiscountPercent
                  Rs3("LineDiscount").value = LineDiscount
                   Rs3("PerforVLineDiscount").value = PerforVLineDiscount
                  linenetaftermainDiscountBeforevat = val(.TextMatrix(i, .ColIndex("NetExe"))) - LineDiscount
                  
                 Rs3("linenetaftermainDiscountBeforevat").value = linenetaftermainDiscountBeforevat
                 
                LineVat = Rs3("linenetaftermainDiscountBeforevat").value * val(TxtFATYou.text) / 100
                
                 Rs3("LineVat").value = LineVat
                 linenetaftermainDiscountWithvat = linenetaftermainDiscount + LineVat
                 
                 
                 
                  Rs3("linenetaftermainDiscountWithvat").value = linenetaftermainDiscountWithvat
             '     Rs3("linenetaftermainDiscountWithvat").value = linenetaftermainDiscountWithvat - PerforVLineDiscount
                  Rs3("LineFinal").value = linenetaftermainDiscountWithvat - PerforVLineDiscount
                  
              'newwwwwwwwwwwwwwwwwwwwwwww
              Rs3("QtyApprov").value = IIf(.TextMatrix(i, .ColIndex("QtyApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("QtyApprov"))))
              Rs3("TotalApprov").value = IIf(.TextMatrix(i, .ColIndex("TotalApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("TotalApprov"))))
              Rs3("PriceApprov").value = IIf(.TextMatrix(i, .ColIndex("PriceApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("PriceApprov"))))
              Rs3("DiscApprov").value = IIf(.TextMatrix(i, .ColIndex("DiscApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("DiscApprov"))))
              Rs3("NetApprov").value = IIf(.TextMatrix(i, .ColIndex("NetApprov")) = "", Null, val(.TextMatrix(i, .ColIndex("NetApprov"))))
              '''//////////////
                Rs3("discountEXE").value = IIf(.TextMatrix(i, .ColIndex("discountEXE")) = "", Null, val(.TextMatrix(i, .ColIndex("discountEXE"))))
                
                Rs3("NetExe").value = IIf(.TextMatrix(i, .ColIndex("NetExe")) = "", Null, val(.TextMatrix(i, .ColIndex("NetExe"))))
                


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
            If val(TXTOrDer_no2.text) <> 0 And CBoBasedON.ListIndex = 1 Then
            
                 UpdatePre_QuantityCont i
            End If
        Next i

    End With
updateNotesValueAndNobytext val(note_id.text)
saveBillBuy

Savetemp


 If bill_Type.ListIndex = 0 Then
        If Not chkTaxExempt.value = vbChecked And SystemOptions.ApplyEinvoice Then savenewelectroncic
    End If
If SystemOptions.IsBluee = True And bill_Type.ListIndex = 0 Then
 
   
                MsgBox SENDEINVOICE(Me.XPTxtBillID, True, val(Me.TXTEnd_user_id.text), 1, "project_billl", "ID"), vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  
End If


    TxtModFlg.text = "R"
fillapprovData
    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Saved", vbInformation
    Else
        MsgBox " „ Õ›Ÿ «·»Ì«‰« ", vbInformation
  
    End If
    Exit Function
ErrTrap:
    
    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "error During Saving", vbInformation
    Else
        MsgBox "ÕœÀ Œÿ√ „« «À‰«¡ Õ›Ÿ «·»Ì«‰«  ", vbInformation
  
    End If
End Function



Private Sub DcCurrency_Change()

    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.Dccurrency.BoundText <> "" Then
        txt_Currency_rate.text = get_currency_rate(Me.Dccurrency.BoundText)
    Else
        txt_Currency_rate.text = 1
    End If

End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub








Function savenewelectroncic()
   'vat data
    Dim InvoiceTypeCodeID As Integer
    rs("CIBAN").value = TXTIban.text
    'vat data
      rs("RecTime").value = Time
            
   
   
  If val(DCDocTypes.BoundText) <> 0 Then
  'wAEL
        getDocAccounts val(DCDocTypes.BoundText), , , , , , , , , , , , InvoiceTypeCodeID
  Else
        InvoiceTypeCodeID = 388
  End If
  rs("InvoiceTypeCodeID").value = InvoiceTypeCodeID
 
 
 
 If val(Me.DefaultInvoicetype.ListIndex) = 0 Then
   
   
    If Export = 1 Then
    rs("InvoiceTypeCodename").value = "0100100"
    Else
      rs("InvoiceTypeCodename").value = "0100000"
   End If
   
   
   
   
   Else
    rs("InvoiceTypeCodename").value = "0200000"
   End If

   rs("DocumentCurrencyCode").value = Dccurrency.text
   rs("TaxCurrencyCode").value = Dccurrency.text
  rs("ActualDeliveryDate").value = txtDateRec.value
 rs("LatestDeliveryDate").value = txtDateRec.value
Dim PaymentMeansCode As String
         
            '10 In cash
            '30 Credit
            '42 Payment to bank account
            '48 Bank card
            '1 Instrument not defined(Free text)
            Dim paymentnote
'        If CboPayMentType.ListIndex = 0 Then ' ‰ﬁœ«
'                  PaymentMeansCode = "10"
'                      paymentnote = "Payment by Cash"
'        ElseIf CboPayMentType.ListIndex = 1 Then ' ¬Ã·
'                 PaymentMeansCode = "30"
'                 paymentnote = "Payment by Credit"
'         ElseIf CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then  '  ÕÊÌ· »‰ﬂÌ
'                    If SystemOptions.AllowSalesMultyPayed = True Then
'                     PaymentMeansCode = "48" 'ﬂ«— 
'                      paymentnote = "Payment by Bank Card"
'                    Else
'                    PaymentMeansCode = "42" '»‰ﬂ Õ”«»
'                    paymentnote = "Payment by Bank Account"
'                    End If
'
'         End If
         PaymentMeansCode = "30"
                 paymentnote = "Payment by Credit"
         rs("PaymentMeansCode").value = PaymentMeansCode
      
rs("paymentnote").value = paymentnote
rs.update
End Function


