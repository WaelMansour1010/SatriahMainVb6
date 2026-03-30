VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form projectsbill_Search 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ «·„‘«—Ì⁄"
   ClientHeight    =   8160
   ClientLeft      =   3525
   ClientTop       =   1470
   ClientWidth     =   15930
   Icon            =   "projectsbill_search.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   15930
   Begin VB.Frame Frame7 
      Height          =   855
      Left            =   0
      TabIndex        =   122
      Top             =   7200
      Width           =   15855
      Begin ALLButtonS.ALLButton btnOk 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "„Ê«›ﬁ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
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
         MICON           =   "projectsbill_search.frx":6852
         PICN            =   "projectsbill_search.frx":686E
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
   Begin VB.Frame Frame14 
      Caption         =   "«”„«¡ «·⁄«„·Ì‰ ›Ì «·„‘—Ê⁄"
      Height          =   3615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   109
      Top             =   9840
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox txt_employee_count 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   9960
         TabIndex        =   111
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_emp_salary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   6360
         TabIndex        =   110
         Top             =   3000
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   1860
         Left            =   120
         TabIndex        =   112
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
         FormatString    =   $"projectsbill_search.frx":D0D0
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
         TabIndex        =   113
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
         MICON           =   "projectsbill_search.frx":D2A9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label27 
         Caption         =   "«Ã„«·Ì ⁄œœ «·⁄„·"
         Height          =   255
         Left            =   11640
         TabIndex        =   115
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·"
         Height          =   255
         Left            =   8040
         TabIndex        =   114
         Top             =   3120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame13 
      Height          =   2772
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   99
      Top             =   600
      Width           =   15975
      Begin VB.TextBox ManualNO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   9480
         TabIndex        =   130
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   13200
         TabIndex        =   129
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   13200
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TXTEnd_user_id 
         Height          =   285
         Left            =   18120
         TabIndex        =   116
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtprojectname 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         DataField       =   "project_name"
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   12600
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox bill_Type 
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "projectsbill_search.frx":D2C5
         Left            =   240
         List            =   "projectsbill_search.frx":D2CF
         TabIndex        =   11
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ComboBox billto 
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "projectsbill_search.frx":D2E0
         Left            =   5640
         List            =   "projectsbill_search.frx":D2EA
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   8535
      End
      Begin VB.TextBox DcAccount2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   8535
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   100
         Top             =   3600
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   5640
         TabIndex        =   2
         Top             =   720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbosubContractor 
         Height          =   315
         Left            =   5640
         TabIndex        =   6
         Top             =   1800
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcoItem 
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Top             =   2160
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker todate 
         Height          =   315
         Left            =   5640
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   213450755
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker fromDate 
         Height          =   315
         Left            =   7440
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   213450755
         CurrentDate     =   37140
      End
      Begin ALLButtonS.ALLButton btnSearch 
         Default         =   -1  'True
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "»ÕÀ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
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
         MICON           =   "projectsbill_search.frx":D306
         PICN            =   "projectsbill_search.frx":D322
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker todate1 
         Height          =   315
         Left            =   240
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   213450755
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker fromDate1 
         Height          =   315
         Left            =   2160
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   213450755
         CurrentDate     =   37140
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11040
         TabIndex        =   131
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label190 
         Alignment       =   2  'Center
         Caption         =   "«·Ì"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   128
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "„‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3720
         TabIndex        =   127
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "«·Ì"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6720
         TabIndex        =   124
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "«·»‰œ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14280
         TabIndex        =   120
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "«·„ﬁ«Ê· «·»«ÿ‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14160
         TabIndex        =   119
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   118
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›—⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   117
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "—ﬁ„ «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14160
         TabIndex        =   108
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "«”„ «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   107
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "«”„ «·⁄„Ì· «·‰Â«∆Ì"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14160
         TabIndex        =   106
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "—ﬁ„ «·„” Œ·’"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14160
         TabIndex        =   105
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "„‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8520
         TabIndex        =   104
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "«·„” Œ·’ «·Ï"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14160
         TabIndex        =   103
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "‰Ê⁄ «·„” Œ·’"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   102
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   720
         TabIndex        =   101
         Top             =   3720
         Width           =   1092
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "«·„’—Ê›« "
      Height          =   3615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   94
      Top             =   9840
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
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   2340
         Left            =   240
         TabIndex        =   96
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
         FormatString    =   $"projectsbill_search.frx":13B84
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
         TabIndex        =   97
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
         MICON           =   "projectsbill_search.frx":13C92
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
         TabIndex        =   98
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "„Ê«œ «·⁄„·Ì… —ﬁ„"
      Height          =   3615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   72
      Top             =   9960
      Visible         =   0   'False
      Width           =   15375
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
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1530
      End
      Begin VB.TextBox TxtFillData 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   0
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox XPTxtBillID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   360
         Left            =   720
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   690
         Index           =   2
         Left            =   360
         TabIndex        =   76
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
         Begin VB.ComboBox CboItemCase 
            Height          =   288
            Left            =   6870
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   300
            Width           =   1920
         End
         Begin VB.TextBox TxtQuantity 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2730
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   300
            Width           =   1770
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   4500
            MaxLength       =   20
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   300
            Width           =   2310
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   900
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   300
            Width           =   1755
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   8805
            TabIndex        =   81
            Top             =   300
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboItemsCode 
            Height          =   315
            Left            =   11790
            TabIndex        =   82
            Top             =   300
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdAdd 
            Height          =   375
            Left            =   120
            TabIndex        =   83
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
            ButtonImage     =   "projectsbill_search.frx":13CAE
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
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   31
            Left            =   11985
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   0
            Width           =   2700
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈”„ «·’‰›"
            Height          =   255
            Index           =   30
            Left            =   9150
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   0
            Width           =   2640
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·’‰›"
            Height          =   255
            Index           =   29
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   0
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ì—Ì«·"
            Height          =   255
            Index           =   28
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   0
            Width           =   2205
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì…"
            Height          =   255
            Index           =   27
            Left            =   2925
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   0
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—"
            Height          =   255
            Index           =   26
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   0
            Width           =   1635
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   1905
         Left            =   240
         TabIndex        =   90
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
         FormatString    =   $"projectsbill_search.frx":14048
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
         TabIndex        =   91
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
         MICON           =   "projectsbill_search.frx":14210
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
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·«’‰«›"
         Height          =   255
         Index           =   2
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label LblItemsCount 
         Caption         =   "Label27"
         Height          =   135
         Left            =   240
         TabIndex        =   92
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "⁄„·Ì«  ﬂ· »‰œ"
      Height          =   3615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   9960
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox txt_opr_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   2760
         Width           =   3015
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   2340
         Left            =   120
         TabIndex        =   64
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
         FormatString    =   $"projectsbill_search.frx":1422C
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
         TabIndex        =   65
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
         MICON           =   "projectsbill_search.frx":145AA
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
         TabIndex        =   66
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
         MICON           =   "projectsbill_search.frx":145C6
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
         TabIndex        =   67
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
         MICON           =   "projectsbill_search.frx":145E2
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
         TabIndex        =   68
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
         MICON           =   "projectsbill_search.frx":145FE
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
         TabIndex        =   69
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.TextBox TxtModFlg 
      Height          =   285
      Left            =   14760
      TabIndex        =   61
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox note_id 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2280
      TabIndex        =   60
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtsubaccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   17280
      TabIndex        =   59
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtendaccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   16200
      TabIndex        =   58
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "total"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   1320
      TabIndex        =   56
      Top             =   9840
      Width           =   1095
   End
   Begin VB.TextBox txtdate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   18840
      TabIndex        =   55
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   2400
      TabIndex        =   49
      Top             =   9840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   53
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   52
         Top             =   240
         Width           =   975
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4440
         TabIndex        =   51
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   720
      TabIndex        =   44
      Top             =   10200
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   47
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   46
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   4095
      Left            =   6960
      TabIndex        =   40
      Top             =   9600
      Width           =   1455
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   0
         TabIndex        =   41
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   42
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
            MICON           =   "projectsbill_search.frx":1461A
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
   Begin VB.Frame Frame5 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   120
      TabIndex        =   34
      Top             =   10320
      Width           =   6495
      Begin VB.Line Line1 
         Index           =   7
         X1              =   960
         X2              =   960
         Y1              =   120
         Y2              =   720
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
         TabIndex        =   38
         Top             =   240
         Width           =   1335
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
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4460
         X2              =   4460
         Y1              =   120
         Y2              =   720
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
         TabIndex        =   36
         Top             =   240
         Width           =   975
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
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6180
         X2              =   6180
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   2700
         X2              =   2700
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   960
      TabIndex        =   27
      Top             =   9840
      Width           =   6495
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   5520
         X2              =   5520
         Y1              =   120
         Y2              =   720
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
         TabIndex        =   32
         Top             =   240
         Width           =   1215
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
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3780
         X2              =   3780
         Y1              =   120
         Y2              =   720
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
         TabIndex        =   29
         Top             =   240
         Width           =   735
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
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   840
      TabIndex        =   24
      Top             =   9720
      Visible         =   0   'False
      Width           =   2415
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
         TabIndex        =   39
         Top             =   1920
         Width           =   2415
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
         TabIndex        =   30
         Top             =   1440
         Width           =   1815
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
         TabIndex        =   26
         Top             =   1320
         Width           =   1812
      End
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
         TabIndex        =   25
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   11520
      TabIndex        =   23
      Top             =   9600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   240
      TabIndex        =   20
      Top             =   9720
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
         Caption         =   "M15"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   9840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "last_root"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text5"
      Top             =   9840
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "projectsbill_search.frx":14636
      Height          =   2895
      Left            =   120
      TabIndex        =   17
      Top             =   10200
      Visible         =   0   'False
      Width           =   7575
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
      Height          =   570
      Left            =   -240
      Top             =   9960
      Visible         =   0   'False
      Width           =   1785
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
      Height          =   330
      Left            =   10320
      Top             =   10080
      Visible         =   0   'False
      Width           =   1785
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
      Height          =   330
      Left            =   10320
      Top             =   10440
      Visible         =   0   'False
      Width           =   1785
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
      Height          =   330
      Left            =   10320
      Top             =   10800
      Visible         =   0   'False
      Width           =   1785
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
      Height          =   330
      Left            =   10320
      Top             =   11040
      Visible         =   0   'False
      Width           =   1785
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
      Height          =   330
      Left            =   10320
      Top             =   10080
      Visible         =   0   'False
      Width           =   1785
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
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   33
      ToolTipText     =   "Language  «··€…"
      Top             =   120
      Visible         =   0   'False
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
      MICON           =   "projectsbill_search.frx":1464B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   4080
      Top             =   9840
      Width           =   7695
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   312
      Left            =   5400
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   840
      Width           =   1428
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy/M/d"
      Format          =   184352771
      CurrentDate     =   37140
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   3360
      Width           =   15975
      Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
         Height          =   3540
         Left            =   150
         TabIndex        =   15
         Top             =   240
         Width           =   15720
         _cx             =   27728
         _cy             =   6244
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"projectsbill_search.frx":14667
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   615
            Left            =   840
            TabIndex        =   123
            Top             =   1800
            Visible         =   0   'False
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   1085
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin ALLButtonS.ALLButton CmdRemove 
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Tag             =   "Delete Row"
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         BCOL            =   0
         BCOLO           =   0
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "projectsbill_search.frx":14888
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Image ImgFavoritesdd 
      Height          =   615
      Left            =   15120
      Picture         =   "projectsbill_search.frx":148A4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   885
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "«Ã„«·Ì «·›« Ê—…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   57
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   12480
      TabIndex        =   54
      Top             =   9840
      Width           =   2172
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "»ÕÀ ›Ê« Ì— «·„‘«—Ì⁄                "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   -8280
      TabIndex        =   19
      Top             =   0
      Width           =   24255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   11880
      TabIndex        =   18
      Top             =   9840
      Width           =   855
   End
End
Attribute VB_Name = "projectsbill_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Long
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

Public mIndex As Integer
Private Sub bill_Type_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
btnSearch_Click
End If
End Sub

Private Sub billto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
btnSearch_Click
End If
End Sub

Private Sub BtnOK_Click()
Dim ID As Integer
Dim row As Integer
row = Fg_Journal.row
ID = val(Fg_Journal.TextMatrix(Fg_Journal.row, Fg_Journal.ColIndex("id")))
If mIndex = 0 Then
    projectsbill.Search (ID)
Else
    frmSubcontractorContract.Search (ID)
End If
End Sub
Private Sub btnSearch_Click()
Dim ss As Integer
ProgressBar1.Visible = True
: ProgressBar1.value = 10
   Set rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 
 'strsql = strsql + " select project_billl.* , project_bill_details.item  from project_bill_details , project_billl  where project_billl.id = project_bill_details.bill_id   "
 
 If mIndex = 0 Then
   StrSQL = " SELECT       dbo.project_billl.*, dbo.project_bill_details.item,   dbo.projects.fullcode , dbo.projects.Project_nameE, dbo.projects.sub_contractor_id, dbo.projects.End_user_id, "
StrSQL = StrSQL & "                                  dbo.TblCustemers.CusName AS endUsernamen, dbo.TblCustemers.CusNamee AS endUsernamene, TblCustemers_1.CusName AS subconnamen, CC.FullCode as DcbosubContractorCode,CC.CusName as DcbosubContractor,"
StrSQL = StrSQL & "                                  TblCustemers_1.CusNamee AS subconnamene"

StrSQL = StrSQL & "            FROM            TblCustemers AS TblCustemers_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                                     projects ON TblCustemers_1.CusID = projects.sub_contractor_id RIGHT OUTER JOIN"
StrSQL = StrSQL & "                                     TblCustemers ON projects.End_user_id = TblCustemers.CusID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                                     project_bill_details INNER JOIN"
StrSQL = StrSQL & "                                     project_billl ON project_bill_details.bill_id = project_billl.id ON projects.id = project_billl.project_no LEFT OUTER JOIN"
StrSQL = StrSQL & "                                     TblCustemers AS CC ON project_billl.subContractorId = CC.CusID"

'StrSQL = StrSQL & "            FROM         dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
'StrSQL = StrSQL & "                                  dbo.projects ON TblCustemers_1.CusID = dbo.projects.sub_contractor_id RIGHT OUTER JOIN"
'StrSQL = StrSQL & "                                  dbo.TblCustemers ON dbo.projects.End_user_id = dbo.TblCustemers.CusID RIGHT OUTER JOIN"
'StrSQL = StrSQL & "                                  dbo.project_bill_details INNER JOIN"
'StrSQL = StrSQL & "                                  dbo.project_billl ON dbo.project_bill_details.bill_id = dbo.project_billl.id ON dbo.projects.id = dbo.project_billl.project_no"
'StrSQL = StrSQL & "                                  RIGHT OUTER JOIN dbo.TblCustemers CC ON dbo.project_billl.subContractorId = CC.CusID "
StrSQL = StrSQL & "                 where 1=1"
    
        StrSQL = StrSQL + " and  project_billl.Branch_NO in(" & Current_branchSql & ") "

 If txtid.text <> "" Then
        StrSQL = StrSQL + " and project_billl.NoteSerial1 = '" & txtid.text & "'"
 End If
: ProgressBar1.value = 20
  If DataCombo2.BoundText <> "" Then
  
  ss = val(DataCombo2.BoundText)
         StrSQL = StrSQL + " and project_billl.Project_no = " & ss
  End If
  
  If DcAccount2.text <> "" Then
        StrSQL = StrSQL + " and project_billl.End_user_name like '%" & DcAccount2.text & "%'"
  End If
  
    If ManualNO.text <> "" Then
        StrSQL = StrSQL + " and project_billl.ManualNO like '%" & ManualNO.text & "%'"
  End If
  
  
 
  
    If txtprojectname.text <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
'dbo.projects.Project_nameE='1'

        StrSQL = StrSQL + " and projects.Project_name like '%" & txtprojectname.text & "%'"
Else
   StrSQL = StrSQL + " and projects.Project_nameE like '%" & txtprojectname.text & "%'"
End If
  End If
  
 'txtprojectname
  
: ProgressBar1.value = 30
  If billto.ListIndex <> -1 Then
        StrSQL = StrSQL + " and project_billl.bill_to = " & billto.ListIndex
  End If
  
  If Dcbranch.text <> "" Then
       StrSQL = StrSQL + " and project_billl.branch_no = " & val(Dcbranch.BoundText)
  End If
: ProgressBar1.value = 40
   If bill_Type.text <> "" Then
        StrSQL = StrSQL + " and project_billl.bill_Type = '" & bill_Type.text & "'"
   End If
: ProgressBar1.value = 50
   
    If Not IsNull(Me.fromDate.value) Then
          StrSQL = StrSQL + " and  bill_Date   >= " & SQLDate(Me.fromDate.value, True) & ""
    End If
   
    If Not IsNull(Me.todate.value) Then
          StrSQL = StrSQL + " and  bill_Date   <= " & SQLDate(Me.todate.value, True) & ""
    End If
      If DcbosubContractor.text <> "" Then
                    StrSQL = StrSQL + " and project_billl.subContractorId = " & val(DcbosubContractor.BoundText)
              End If
    
    If Not IsNull(Me.fromDate1.value) Then
          StrSQL = StrSQL + " and  dueDate   >= " & SQLDate(Me.fromDate1.value, True) & ""
    End If
   
    If Not IsNull(Me.todate1.value) Then
          StrSQL = StrSQL + " and  dueDate   <= " & SQLDate(Me.todate1.value, True) & ""
    End If
    
    
 
    
    StrSQL = StrSQL + " order by    project_billl.ID"
    
  Else
  
  
  
              
              StrSQL = " SELECT       dbo.SubcontractorContract.*, dbo.SubcontractorContract2.item,   dbo.projects.fullcode , dbo.projects.Project_nameE, dbo.projects.sub_contractor_id, dbo.projects.End_user_id, "
            StrSQL = StrSQL & "                                  dbo.TblCustemers.CusName AS endUsernamen, dbo.TblCustemers.CusNamee AS endUsernamene, TblCustemers_1.CusName AS subconnamen,"
            StrSQL = StrSQL & "                                  CC.FullCode as DcbosubContractorCode,CC.CusName as DcbosubContractor,"
            StrSQL = StrSQL & "                                  TblCustemers_1.CusNamee AS subconnamene"
            StrSQL = StrSQL & "            FROM         dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
            StrSQL = StrSQL & "                                  dbo.projects ON TblCustemers_1.CusID = dbo.projects.sub_contractor_id RIGHT OUTER JOIN"
            StrSQL = StrSQL & "                                  dbo.TblCustemers ON dbo.projects.End_user_id = dbo.TblCustemers.CusID RIGHT OUTER JOIN"
            
          
          
           
            
            
            StrSQL = StrSQL & "                                  dbo.SubcontractorContract2 INNER JOIN"
            StrSQL = StrSQL & "                                  dbo.SubcontractorContract ON dbo.SubcontractorContract2.bill_id = dbo.SubcontractorContract.id ON dbo.projects.id = dbo.SubcontractorContract.project_no"
             
           StrSQL = StrSQL & "                                  RIGHT OUTER JOIN dbo.TblCustemers CC ON dbo.SubcontractorContract.subContractorId = CC.CusID "
           
            StrSQL = StrSQL & "                 where 1=1"
                
                    StrSQL = StrSQL + " and  SubcontractorContract.Branch_NO in(" & Current_branchSql & ") "
            
             If txtid.text <> "" Then
                    StrSQL = StrSQL + " and SubcontractorContract.NoteSerial1 = '" & txtid.text & "'"
             End If
: ProgressBar1.value = 20
              If DataCombo2.BoundText <> "" Then
              
              ss = val(DataCombo2.BoundText)
                     StrSQL = StrSQL + " and SubcontractorContract.Project_no = " & ss
              End If
              
              If DcAccount2.text <> "" Then
                    StrSQL = StrSQL + " and SubcontractorContract.End_user_name like '%" & DcAccount2.text & "%'"
              End If
              
                If ManualNO.text <> "" Then
                    StrSQL = StrSQL + " and SubcontractorContract.ManualNO like '%" & ManualNO.text & "%'"
              End If
              
                If DcbosubContractor.text <> "" Then
                    StrSQL = StrSQL + " and SubcontractorContract.subContractorId = " & val(DcbosubContractor.BoundText)
              End If
'DcbosubContractor
              
  
              
              
                If txtprojectname.text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
            'dbo.projects.Project_nameE='1'
            
                    StrSQL = StrSQL + " and projects.Project_name like '%" & txtprojectname.text & "%'"
            Else
               StrSQL = StrSQL + " and projects.Project_nameE like '%" & txtprojectname.text & "%'"
            End If
              End If
              
             'txtprojectname
              
: ProgressBar1.value = 30
              If billto.ListIndex <> -1 Then
                    StrSQL = StrSQL + " and SubcontractorContract.bill_to = " & billto.ListIndex
              End If
              
              If Dcbranch.text <> "" Then
                   StrSQL = StrSQL + " and SubcontractorContract.branch_no = " & val(Dcbranch.BoundText)
              End If
: ProgressBar1.value = 40
               If bill_Type.text <> "" Then
                    StrSQL = StrSQL + " and SubcontractorContract.bill_Type = '" & bill_Type.text & "'"
               End If
: ProgressBar1.value = 50
               
                If Not IsNull(Me.fromDate.value) Then
                      StrSQL = StrSQL + " and  bill_Date   >= " & SQLDate(Me.fromDate.value, True) & ""
                End If
               
                If Not IsNull(Me.todate.value) Then
                      StrSQL = StrSQL + " and  bill_Date   <= " & SQLDate(Me.todate.value, True) & ""
                End If
                
                
                If Not IsNull(Me.fromDate1.value) Then
                      StrSQL = StrSQL + " and  dueDate   >= " & SQLDate(Me.fromDate1.value, True) & ""
                End If
               
                If Not IsNull(Me.todate1.value) Then
                      StrSQL = StrSQL + " and  dueDate   <= " & SQLDate(Me.todate1.value, True) & ""
                End If
                
                
             
                
                StrSQL = StrSQL + " order by    SubcontractorContract.ID"
  End If
    
: ProgressBar1.value = 70
   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
: ProgressBar1.value = 80

    If rs.RecordCount < 1 Then
        '    XPTxtCurrent.Caption = 0
        '    XPTxtCount.Caption = 0
            ProgressBar1.Visible = False
            ProgressBar1.value = 0
       Fg_Journal.rows = Fg_Journal.FixedRows
                Exit Sub
    End If
: ProgressBar1.value = 100
    If rs.EOF Or rs.BOF Then
        Exit Sub
         ProgressBar1.Visible = False
         ProgressBar1.value = 0
    Else
    Retrive
   ProgressBar1.Visible = False
   ProgressBar1.value = 0
End If
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
   
 
End Function

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
                If Me.Fg_Journal.row <> Me.Fg_Journal.FixedRows - 1 Then
                    Me.Fg_Journal.RemoveItem (Me.Fg_Journal.row)
                End If
            End If
        End If
    End If
            
    ReLineGrid

End Sub

Private Sub Command1_Click(index As Integer)


    'On Error Resume Next
    Select Case index
Case 12
txtid.text = ""
    TxtModFlg.text = "N"

            Fg_Journal.rows = Fg_Journal.rows + 1
            Fg_Journal.Enabled = True
            Command1(1).Enabled = True
            
        Case 0
 
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.rows = 2
            Fg_Journal.Enabled = True
            Command1(1).Enabled = True
            XPDtbTrans.value = DateValue(Now)
       
            XPDtbTrans.value = Date
            TxtNoteSerial.text = ""
            Me.Dcbranch.BoundText = Current_branch
            cboDiscount1.ListIndex = 0
            cboDiscount1.ListIndex = 0
            cboDiscount2.ListIndex = 0

        Case 1
    
If val(total.text) <= 0 Then MsgBox "Õœœ  ﬂ·›… «·„„‰›– «Ê·«", vbCritical: Exit Sub
If Not IsNumeric(Dcbranch.BoundText) Then MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical: Exit Sub

    
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
GET_PROJECT_DATA
            SaveData

            ''Adodc1.Recordset.Fields!  project_no = DataCombo2.text
        Case 11

            On Error Resume Next

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

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            Dim Msg As String
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

            Fg_Journal.rows = Fg_Journal.rows + 1
            Fg_Journal.Enabled = True
            Command1(1).Enabled = True

        Case 4

        Case 5

        Case 6
            Undo

        Case 9

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

            ShowGL_cc Me.TxtNoteSerial.text, , 200
    End Select

End Sub


Function print_report(Optional NoteSerial As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


    MySQL = "select * from project_billl P,project_bill_details PD Where P.id = pd.bill_id and P.id = " & txtid.text

 
      
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Projects.rpt"
       

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

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        
        
        xReport.ParameterFields(3).AddCurrentValue (txtid.text)
        xReport.ParameterFields(4).AddCurrentValue (DataCombo2.text)
        xReport.ParameterFields(5).AddCurrentValue (DcAccount2.text)
        xReport.ParameterFields(6).AddCurrentValue (billto.text)
        xReport.ParameterFields(7).AddCurrentValue (DcbosubContractor.text)
        xReport.ParameterFields(8).AddCurrentValue (cboDiscount1.text)
        xReport.ParameterFields(9).AddCurrentValue (cboDiscount2.text)
        
        xReport.ParameterFields(10).AddCurrentValue (txtManualNo.text)
        
        
        xReport.ParameterFields(11).AddCurrentValue (Format(XPDtbTrans.value, "yyyy/M/d"))
        
        
        xReport.ParameterFields(12).AddCurrentValue (Dcbranch.text)
        xReport.ParameterFields(13).AddCurrentValue (txtprojectname.text)
        xReport.ParameterFields(14).AddCurrentValue (DcAccount1.text)
        
          xReport.ParameterFields(15).AddCurrentValue (bill_Type.text)
          
          xReport.ParameterFields(16).AddCurrentValue (Format(dueDate1.value, "yyyy/M/d"))
            
          xReport.ParameterFields(17).AddCurrentValue (Format(dueDate.value, "yyyy/M/d"))
          
               xReport.ParameterFields(18).AddCurrentValue (TxtRemarks.text)
               
                  xReport.ParameterFields(19).AddCurrentValue (txtid.text)
          ' xReport.ParameterFields(14).AddCurrentValue (DcAccount1.text)
        
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        
           xReport.ParameterFields(3).AddCurrentValue (txtid.text)
        xReport.ParameterFields(4).AddCurrentValue (DataCombo2.text)
        xReport.ParameterFields(5).AddCurrentValue (DcAccount2.text)
        xReport.ParameterFields(6).AddCurrentValue (billto.text)
        xReport.ParameterFields(7).AddCurrentValue (DcbosubContractor.text)
        xReport.ParameterFields(8).AddCurrentValue (cboDiscount1.text)
        xReport.ParameterFields(9).AddCurrentValue (cboDiscount2.text)
        
        xReport.ParameterFields(10).AddCurrentValue (txtManualNo.text)
        
        
        xReport.ParameterFields(11).AddCurrentValue (Format(XPDtbTrans.value, "yyyy/M/d"))
        
        
        xReport.ParameterFields(12).AddCurrentValue (Dcbranch.text)
        xReport.ParameterFields(13).AddCurrentValue (txtprojectname.text)
        xReport.ParameterFields(14).AddCurrentValue (DcAccount1.text)
        
          xReport.ParameterFields(15).AddCurrentValue (bill_Type.text)
          
          xReport.ParameterFields(16).AddCurrentValue (Format(dueDate1.value, "yyyy/M/d"))
            
          xReport.ParameterFields(17).AddCurrentValue (Format(dueDate.value, "yyyy/M/d"))
          
               xReport.ParameterFields(18).AddCurrentValue (TxtRemarks.text)
               
                  xReport.ParameterFields(19).AddCurrentValue (txtid.text)
       
        
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
       ' StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
'        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
'         xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
'   xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

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



Function GET_PROJECT_DATA()
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
    txtsubaccount.text = IIf(IsNull(Rec.Fields("sub_contractor_Account").value), "", Rec.Fields("sub_contractor_Account").value)
    
    
    DcAccount1.text = IIf(IsNull(Rec.Fields("sub_contractor_name").value), "", Rec.Fields("sub_contractor_name").value)
    txtendaccount.text = IIf(IsNull(Rec.Fields("End_user_Account").value), "", Rec.Fields("End_user_Account").value)
    DcAccount2.text = IIf(IsNull(Rec.Fields("End_user_name").value), "", Rec.Fields("End_user_name").value)
 Dim End_user_id As Double

Dim sub_contractor_id As Double
 
 
 End_user_id = IIf(IsNull(Rec.Fields("End_user_id").value), 0, Rec.Fields("End_user_id").value)
 sub_contractor_id = IIf(IsNull(Rec.Fields("sub_contractor_id").value), 0, Rec.Fields("sub_contractor_id").value)
'DcAccount2.text = GET_ACCOUNT_name_by_Code(get_Customer_Account(End_user_id))
' DcAccount1.text = GET_ACCOUNT_name_by_Code(get_Customer_Account(sub_contractor_id))
 
 
 

 




 If SystemOptions.Revenueowed = True Then
    txtrevenue_account.text = IIf(IsNull(Rec.Fields("legal").value), "", Rec.Fields("legal").value) 'Õ”«» «·„” Œ·’« \
  Else
      txtrevenue_account.text = IIf(IsNull(Rec.Fields("REVENUE_account").value), "", Rec.Fields("REVENUE_account").value) 'Õ”«» «·«Ì—«œ« \

  End If
  
TXTEnd_user_id.text = IIf(IsNull(Rec.Fields("End_user_id").value), "", Rec.Fields("End_user_id").value) '—ﬁ„ «·⁄„Ì· «·‰Â«∆Ì
TXTsub_contractor_id.text = IIf(IsNull(Rec.Fields("sub_contractor_id").value), "", Rec.Fields("sub_contractor_id").value) '—ﬁ„   „ﬁ«Ê· «·»«ÿ‰

 expanses_account = IIf(IsNull(Rec.Fields("expanses_account").value), "", Rec.Fields("expanses_account").value) 'Õ”«»  «·„’—Ê›« \

    'My_SQL = "  select net,des from projects_des  where project_id='" & DataCombo2.BoundText & "'"
    'fill_combo DataCombo5, My_SQL

End Function

Private Sub DataCombo2_Change()
   ' GET_PROJECT_DATA
End Sub

Private Sub DataCombo2_Click(Area As Integer)
  '  GET_PROJECT_DATA
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
            total.text = gettotal(txtid.text)

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

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
btnSearch_Click
End If
End Sub

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

Private Sub DcbosubContractor_Change()
    Dim Fullcode As String
   
     creditlocked = 0
     Dim CPaymentType As Integer
    If Trim(DcbosubContractor.text) = "" Then Exit Sub
    GetCustomersDetail val(DcbosubContractor.BoundText), , Fullcode, 3
    Text2.text = Fullcode

End Sub

Private Sub DcbosubContractor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
btnSearch_Click
End If
End Sub

Private Sub DcbosubContractor_KeyUp(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 6060
           FrmCompanySearch.show vbModal
           
        End If
        
End Sub

Private Sub Dcbranch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
btnSearch_Click
End If
End Sub

Private Sub dcoItem_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
btnSearch_Click
End If
End Sub

Private Sub employee_details_Click(index As Integer)

    Select Case index

        Case 0

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame14.Visible = True

                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.row, VSFlexGrid2.ColIndex("fullcode"))
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
            VSFlexGrid2.TextMatrix(VSFlexGrid2.row, VSFlexGrid2.ColIndex("total_salary1")) = val(txt_emp_salary)
            ReLineGrid

    End Select

End Sub

Private Sub Fg_Journal_AfterEdit(ByVal row As Long, _
                                 ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)
 
            Case "item"
                StrAccountCode = .ComboItem
       
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("item"), False, True)
                .TextMatrix(row, .ColIndex("item")) = StrAccountCode
            
                If StrAccountCode <> "" Then
                    StrSQL = "SELECT   line_no, oprid,des, net, project_id  ,[unit] ,[Quantity],[Price] ,[Pre_Quantity] ,[Pre_Value],[Pre_Percent] ,[Curr_Quantity]  ,[Curr_value] ,[curr_Percent] ,[tot_quantity] ,[tot_value] ,[tot_percent]   from dbo.projects_des  WHERE fullcode='" & .ComboData & "'"  ' project_id =" & Val(DataCombo2.BoundText) & "and line_no=" & Val(.ComboItem)
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
                    .TextMatrix(row, .ColIndex("cost")) = IIf(IsNull(Rs1("net").value), 0, Rs1("net").value)
                    .TextMatrix(row, .ColIndex("exe")) = 0
                    .TextMatrix(row, .ColIndex("percentage")) = 0
                    .TextMatrix(row, .ColIndex("item_id")) = .ComboData
                    
                   '  .TextMatrix(Row, .ColIndex("unit")) = IIf(IsNull(Rs1("unit").value), 0, Rs1("unit").value)
                   '   .TextMatrix(Row, .ColIndex("Quantity")) = IIf(IsNull(Rs1("Quantity").value), 0, Rs1("Quantity").value)
                   '    .TextMatrix(Row, .ColIndex("Price")) = IIf(IsNull(Rs1("Price").value), 0, Rs1("Price").value)
                   '     .TextMatrix(Row, .ColIndex("Pre_Quantity")) = IIf(IsNull(Rs1("Pre_Quantity").value), 0, Rs1("Pre_Quantity").value)
                   '      .TextMatrix(Row, .ColIndex("Pre_Value")) = IIf(IsNull(Rs1("Pre_Value").value), 0, Rs1("Pre_Value").value)
                   '       .TextMatrix(Row, .ColIndex("Pre_Percent")) = IIf(IsNull(Rs1("Pre_Percent").value), 0, Rs1("Pre_Percent").value)
                   '        .TextMatrix(Row, .ColIndex("Curr_Quantity")) = IIf(IsNull(Rs1("Curr_Quantity").value), 0, Rs1("Curr_Quantity").value)
                   ' .TextMatrix(Row, .ColIndex("Curr_value")) = IIf(IsNull(Rs1("Curr_value").value), 0, Rs1("Curr_value").value)
                   ' .TextMatrix(Row, .ColIndex("curr_Percent")) = IIf(IsNull(Rs1("curr_Percent").value), 0, Rs1("curr_Percent").value)
                   ' .TextMatrix(Row, .ColIndex("tot_quantity")) = IIf(IsNull(Rs1("tot_quantity").value), 0, Rs1("tot_quantity").value)
                   ' .TextMatrix(Row, .ColIndex("tot_value")) = IIf(IsNull(Rs1("tot_value").value), 0, Rs1("tot_value").value)
                   ' .TextMatrix(Row, .ColIndex("tot_percent")) = IIf(IsNull(Rs1("tot_percent").value), 0, Rs1("tot_percent").value)
                Else
                    .TextMatrix(row, .ColIndex("cost")) = 0
                    .TextMatrix(row, .ColIndex("exe")) = 0
                    .TextMatrix(row, .ColIndex("percentage")) = 0
                    .TextMatrix(row, .ColIndex("item_id")) = ""
                    
                    
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
 
            Case "exe"
            
            Case "Unit"
               If StrAccountCode <> "" Then
               .TextMatrix(row, .ColIndex("unit_id")) = .ComboData
               End If
        End Select

        '  Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    
        If row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "item"
                .ComboList = ""

            Case "cost"
                .ComboList = ""
        
            Case "exe"
                .ComboList = ""
        
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
        End Select

    End With
 
End Sub

Private Sub Fg_Journal_Click()
   ' current_terms = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("item_id"))
   
   
   With Fg_Journal
    If mIndex = 0 Then
        projectsbill.Retrive val(.TextMatrix(.row, .ColIndex("id")))
    ElseIf mIndex = 1 Then
        frmSubcontractorContract.Retrive val(.TextMatrix(.row, .ColIndex("id")))
    End If
   End With
End Sub

Private Sub Fg_Journal_StartEdit(ByVal row As Long, _
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

            Case "item"
       
                'Full Path Display
                StrSQL = "SELECT   fullcode,line_no, oprid,des, net, project_id from dbo.projects_des WHERE project_id =" & val(DataCombo2.BoundText)
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "des", "fullcode")

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
  '  StrSQL = "SELECT  dueDate,  branch_no,NoteSerial, id, bill_date, project_no, project_name, Sub_user_name, End_user_name, End_user_account, bill_to, Sub_user_account, bill_type, revenue_account, note_id,"
  '  StrSQL = StrSQL + "total  From dbo.project_billl Order by ID"
    
  'StrSQL = "SELECT  dueDate,  branch_no,NoteSerial, id, bill_date, project_no, project_name, Sub_user_name, End_user_name, End_user_account, bill_to, Sub_user_account, bill_type, revenue_account, note_id,"
    
    
    
    StrSQL = StrSQL + "SELECT *  From dbo.project_billl where id = -1 Order by ID"
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

fromDate.value = Date
   fromDate1.value = Date
   
   todate.value = Date
   todate1.value = Date
   
   fromDate.value = Null
   fromDate1.value = Null
   
   todate.value = Null
   todate1.value = Null
   
   '

    'first_run = True
    Dim My_SQL As String
 
    My_SQL = "  select id,Fullcode from Projects"
     My_SQL = My_SQL + " where   branch_no in(" & Current_branchSql & ") "
    fill_combo DataCombo2, My_SQL

Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
  Dcombos.GetBranches Dcbranch
  
    Dcombos.GetPersons Me.DcbosubContractor
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
   ' NewGrid.FillGrid
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    ChangeLang
    
    '    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Command1_Click (0)
    End If

End Sub

Function ChangeLang()

    If SystemOptions.UserInterface = EnglishInterface Then
        Label190.Caption = "To"
'With cboDiscount1
'.Clear
'.AddItem "NA"
'.AddItem "Perce%"
'.AddItem "Value"
''End With

'With cboDiscount2
'.Clear
'.AddItem "NA"
'.AddItem "Perce%"
''.AddItem "Value"
'End With
Label23.Caption = "Manual No."
   fromDate.value = Null
   fromDate1.value = Null
   
   todate.value = Null
   todate1.value = Null
   
With billto
.Clear
.AddItem "End User"
.AddItem "Sub-Contractor"
 
End With

With Me.bill_Type

.Clear
.AddItem "Partial"
.AddItem "Final"
 
End With

btnSearch.Caption = "Search"
 btnOk.Caption = "Ok"
Label7.Caption = "Item"
 
  Label21.Caption = "Due Date"
'Label32.Caption = "deduct adv. Payment"
'Label31.Caption = "deduct ensure business "
Label22.Caption = "Sub-contractor"
 
 
With Fg_Journal
 .TextMatrix(0, .ColIndex("bill_type")) = "Type"
 .TextMatrix(0, .ColIndex("id")) = "Bill No."
 .TextMatrix(0, .ColIndex("ManualNO")) = "Manual No."
 
  
.TextMatrix(0, .ColIndex("no")) = "Project No."
.TextMatrix(0, .ColIndex("project_name")) = "Project Name"
.TextMatrix(0, .ColIndex("End_user_name")) = "End user name"
.TextMatrix(0, .ColIndex("duedate")) = "Due Date"
.TextMatrix(0, .ColIndex("bill_date")) = "Bill Date"

.TextMatrix(0, .ColIndex("DcbosubContractorCode")) = "subContractor Code"
.TextMatrix(0, .ColIndex("DcbosubContractor")) = "subContractor Name"
.TextMatrix(0, .ColIndex("NoteSerial1")) = "Invoice No"


End With
 
 
       ' temp = XPBtnMove(1).left
      '  XPBtnMove(1).left = XPBtnMove(2).left
       ' XPBtnMove(2).left = temp
Label26.Caption = "Branch"

        'temp = XPBtnMove(0).left
        'XPBtnMove(0).left = XPBtnMove(3).left
        'XPBtnMove(3).left = temp
        SetInterface Me
        Me.Caption = "         Project Invoice  search"
        Label9.Caption = Me.Caption

        Label20.Caption = "Bill No."
        Label25.Caption = "Date"

        Label6.Caption = "Project Code"
        Label1.Caption = "Project Name"
         Label15.Caption = "End User"
'        Label23.Caption = "Sub-Contractor"
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
  
'        Command1(0).Caption = "new"
'        Command1(1).Caption = "save"
'        Command1(2).Caption = "Attachments"
''        Command1(3).Caption = "Edit"
'        Command1(6).Caption = "Delete"
  
        SuperLabel2.text = "Search"
        Command1(4).Caption = "By ID"
        Command1(5).Caption = "Search"
  Command1(11).Caption = "Attachement"
        Adodc1.Caption = "move"
  
 '       With Fg_Journal
 '           .TextMatrix(0, .ColIndex("LineNo")) = "Index"
 '           .TextMatrix(0, .ColIndex("Item_ID")) = "Term#"
'
'            .TextMatrix(0, .ColIndex("item")) = "Term Desc."
'            .TextMatrix(0, .ColIndex("cost")) = "cost"
'            .TextMatrix(0, .ColIndex("exe")) = "exe"
'            .TextMatrix(0, .ColIndex("percentage")) = "percentage"
'            .TextMatrix(0, .ColIndex("exedate")) = "exe date"
'
'  .TextMatrix(0, .ColIndex("Unit")) = "Unit"
'  .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
'.TextMatrix(0, .ColIndex("Price")) = "Price"
'.TextMatrix(0, .ColIndex("Pre_Quantity")) = "Pre. Exe. Quantity"
'.TextMatrix(0, .ColIndex("Pre_Value")) = "Pre. exe value "
'.TextMatrix(0, .ColIndex("Pre_Percent")) = "Pre. exe percentage"
'.TextMatrix(0, .ColIndex("Curr_Quantity")) = " Current exe Quantity"
'.TextMatrix(0, .ColIndex("Curr_value")) = " Current exe value"
'.TextMatrix(0, .ColIndex("curr_Percent")) = "Current exe percentage"
'.TextMatrix(0, .ColIndex("tot_quantity")) = "Total Quantity"
'.TextMatrix(0, .ColIndex("tot_value")) = "Total Value"
'.TextMatrix(0, .ColIndex("tot_percent")) = "Total Percent"
'
'
'        End With

        opr_items(0).Caption = "View Term Operations"
        Frame11.Caption = "Term Operaions"
 
        Label27.Caption = "Labors Count"
        Label24.Caption = "Total"
Label190.Caption = "To"
        With VSFlexGrid1
            .TextMatrix(0, .ColIndex("LineNo")) = "Index"
            .TextMatrix(0, .ColIndex("code")) = "Labor Code"
            .TextMatrix(0, .ColIndex("name")) = "name"

            .TextMatrix(0, .ColIndex("jobname")) = "ıJob"
            .TextMatrix(0, .ColIndex("daysalary")) = "ıday salary"
            .TextMatrix(0, .ColIndex("Start")) = "Start"
            .TextMatrix(0, .ColIndex("End")) = "End"
            .TextMatrix(0, .ColIndex("Count")) = "No Of Days"
            .TextMatrix(0, .ColIndex("total")) = "Total"

        End With

        With VSFlexGrid2
            .TextMatrix(0, .ColIndex("LineNo")) = "Index"
            .TextMatrix(0, .ColIndex("fullcode")) = "OPR Code"

            .TextMatrix(0, .ColIndex("name")) = "Operation Desc."
            .TextMatrix(0, .ColIndex("total_items")) = "total items Cost"
            .TextMatrix(0, .ColIndex("total_salary")) = "total salary"
            .TextMatrix(0, .ColIndex("total_expenses")) = "total expenses"
            .TextMatrix(0, .ColIndex("total")) = "total"
            .TextMatrix(0, .ColIndex("total_items1")) = "total items Cost EXE"
            .TextMatrix(0, .ColIndex("total_salary1")) = "total salary EXE"
            .TextMatrix(0, .ColIndex("total_expenses1")) = "total expenses EXE"
            .TextMatrix(0, .ColIndex("total1")) = "total EXE"

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
            .TextMatrix(0, .ColIndex("value")) = "value"

            .TextMatrix(0, .ColIndex("des")) = "des"
 
        End With

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
        Label17.Caption = "From"
        Label19.Caption = "To"
    Else
        billto.Clear
        billto.AddItem "⁄„Ì· ‰Â«∆Ì"
        billto.AddItem "„ﬁ«Ê· »«ÿ‰"
        bill_Type.Clear
        bill_Type.AddItem "Ã“∆Ì"
        bill_Type.AddItem "‰Â«∆Ì"
 
    End If

End Function

Private Sub retrive1(Item_ID As String)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
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
    Dim StrMSG As String

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
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub FromDate_Change()
 btnSearch_Click
 
End Sub

Private Sub fromDate1_Change()
 btnSearch_Click
 
End Sub

Private Sub ManualNO_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
btnSearch_Click
End If
End Sub

Private Sub opr_expenses_Click(index As Integer)

    Select Case index

        Case 0
  
            VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.rows = 2
            VSFlexGrid3.Enabled = True

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame12.Visible = True

                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.row, VSFlexGrid2.ColIndex("fullcode"))
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

Private Sub opr_items_Click(index As Integer)

    Select Case index

        Case 0

            DTPicker1.value = Date

            If Not Fg_Journal.TextMatrix(Fg_Journal.row, Fg_Journal.ColIndex("item_id")) = "" Then
                Frame11.Visible = True
        
                current_terms = Fg_Journal.TextMatrix(Fg_Journal.row, Fg_Journal.ColIndex("item_id"))
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
                        RsDev("project_id").value = DataCombo2.BoundText
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

Private Sub terms_operations_Click(index As Integer)

End Sub

Private Sub Show_items_Click(index As Integer)

    Select Case index

        Case 0

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame10.Visible = True

                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.row, VSFlexGrid2.ColIndex("fullcode"))
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
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
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



Private Sub Text2_Change()
Dim CUSTID As Integer
    GetCustomersDetail CUSTID, , Text2.text, 3
        DcbosubContractor.BoundText = CUSTID
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text2.text, 3
        DcbosubContractor.BoundText = CUSTID
        
    End If


End Sub

Private Sub Text2_Validate(Cancel As Boolean)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
    
        
    End If


End Sub

Private Sub ToDate_Change()
 
btnSearch_Click
 
End Sub

Private Sub todate1_Change()
 
btnSearch_Click
 
End Sub

Private Sub txtId_Change()
    ' "select * from project_bill_details where bill_id=" & Val(txtid.text)

End Sub

Private Sub Retrive()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        '    XPTxtCurrent.Caption = 0
        '    XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
   End If

'*************************************************

'  Dim StrSQL As String
'    StrSQL = "Select CusID,CusName From TblCustemers"
'    StrSQL = StrSQL + " Where (Type=3)"
'    StrSQL = StrSQL + " Order BY CusName"


    '-----------------------------------------------------------------------------
        If Not (rs.BOF Or rs.EOF) Then
        
            rs.MoveFirst
    
            With Me.Fg_Journal
                .rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .rows - 1
                
                
                  .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                  .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                  
                  If mIndex = 0 Then
                  .TextMatrix(i, .ColIndex("ManualNO")) = IIf(IsNull(rs("ManualNO").value), "", rs("ManualNO").value)
                  End If
                  
                    .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(rs("Project_no").value), "", rs("Project_no").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
                    .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs("endUsernamen").value), "", rs("endUsernamen").value)
                    
                    
                    Else
                  .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs("Project_nameE").value), "", rs("Project_nameE").value)
                  .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs("endUsernamene").value), "", rs("endUsernamene").value)
                    End If
                    
                    .TextMatrix(i, .ColIndex("DcbosubContractor")) = IIf(IsNull(rs("DcbosubContractor").value), "", rs("DcbosubContractor").value)
                    .TextMatrix(i, .ColIndex("DcbosubContractorCode")) = IIf(IsNull(rs("DcbosubContractorCode").value), "", rs("DcbosubContractorCode").value)
                    
                    
                    ' .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs("End_user_name").value), "", rs("End_user_name").value)
                     
                    ' .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs("End_user_name").value), "", rs("End_user_name").value)
                      
                        .TextMatrix(i, .ColIndex("duedate")) = IIf(IsNull(rs("duedate").value), "", rs("duedate").value)
                        .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs("bill_date").value), "", rs("bill_date").value)
                        .TextMatrix(i, .ColIndex("bill_type")) = IIf(IsNull(rs("bill_type").value), "", rs("bill_type").value)
                        
                         .TextMatrix(i, .ColIndex("no")) = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
                rs.MoveNext
                Next i
         
            End With

    End If

    '-----------------------------------------------------------------------------
    'XPTxtCurrent.Caption = Rs.AbsolutePosition
    'XPTxtCount.Caption = Rs.RecordCount
GET_PROJECT_DATA
    ReLineGrid
    Exit Sub
ErrTrap:

End Sub

Private Sub ReLineGrid(Optional current_terms As String = "")
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("item")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                .TextMatrix(i, .ColIndex("cost")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("cost"))), 1, .TextMatrix(i, .ColIndex("cost")))
           
                ' sql = "  From terms_operations Where term_fullcode='" & .TextMatrix(I, .ColIndex("fullcode")) & "'"
                sql = "select sum(total1) as total  from terms_operations_project_bill where term_fullcode='" & .TextMatrix(i, .ColIndex("item_id")) & "' and bill_id=" & val(Me.txtid.text)
         
                Set rs = New ADODB.Recordset
                rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If rs.RecordCount > 0 And Not IsNull(rs("total").value) Then
                    .TextMatrix(i, .ColIndex("exe")) = IIf(IsNull(rs("total").value), 0, rs("total").value)
         
                Else
                    .TextMatrix(i, .ColIndex("exe")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("exe"))), 0, .TextMatrix(i, .ColIndex("exe")))
                End If
        
                .TextMatrix(i, .ColIndex("percentage")) = Round(.TextMatrix(i, .ColIndex("exe")) / .TextMatrix(i, .ColIndex("cost")) * 100, 2)
             
                .TextMatrix(i, .ColIndex("exedate")) = IIf(.TextMatrix(i, .ColIndex("exedate")) = "", Date, .TextMatrix(i, .ColIndex("exedate")))
        
            End If

        Next i

        Me.total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("exe"), .rows - 1, .ColIndex("exe"))
         
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

        Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .rows - 1, .ColIndex("total"))
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
        
            'Me.XPBtnMove(0).Enabled = True
            'Me.XPBtnMove(1).Enabled = True
            'Me.XPBtnMove(2).Enabled = True
           ' Me.XPBtnMove(3).Enabled = True
 
            If rs.RecordCount < 1 Then
              ''  Me.XPBtnMove(0).Enabled = False
                'Me.XPBtnMove(1).Enabled = False
              '  Me.XPBtnMove(2).Enabled = False
             '   Me.XPBtnMove(3).Enabled = False
              '  Me.Command1(9).Enabled = False
              '  Me.Command1(3).Enabled = False
            
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
        
           ' Me.XPBtnMove(0).Enabled = False
           ' Me.XPBtnMove(1).Enabled = False
          '  Me.XPBtnMove(2).Enabled = False
            'Me.XPBtnMove(3).Enabled = False
  
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub txtprojectname_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
btnSearch_Click
End If
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal row As Long, _
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
                .TextMatrix(row, .ColIndex("name")) = StrAccountCode
            
                If StrAccountCode <> "" Then
                    StrSQL = "SELECT   * from dbo.terms_operations WHERE  fullcode ='" & .ComboData & "'"
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             
                    .TextMatrix(row, .ColIndex("fullcode")) = IIf(IsNull(Rs1("fullcode").value), 0, Rs1("fullcode").value)
                    .TextMatrix(row, .ColIndex("total_items")) = IIf(IsNull(Rs1("total_items").value), 0, Rs1("total_items").value)
            
                    .TextMatrix(row, .ColIndex("total_salary")) = IIf(IsNull(Rs1("total_salary").value), 0, Rs1("total_salary").value)
                    .TextMatrix(row, .ColIndex("total_expenses")) = IIf(IsNull(Rs1("total_expenses").value), 0, Rs1("total_expenses").value)
                    .TextMatrix(row, .ColIndex("total")) = IIf(IsNull(Rs1("total").value), 0, Rs1("total").value)
                    .TextMatrix(row, .ColIndex("total_items1")) = get_opr_material_total(.ComboData, DTPicker1.value)
                    .TextMatrix(row, .ColIndex("total_expenses1")) = get_opr_expenses_total(.ComboData, DTPicker1.value)
             
                Else
 
                    .TextMatrix(row, .ColIndex("fullcode")) = ""
             
                End If

        End Select

        '  Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    
        If row = .rows - 1 Then
            .rows = .rows + 1
        End If
    
    End With

    ReLineGrid
End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid2
        .ComboList = ""

        Select Case .ColKey(Col)
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid2_Click()
    current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.row, VSFlexGrid2.ColIndex("fullcode"))

    With VSFlexGrid2
   
        If row = .rows - 1 Then
            .rows = .rows + 1
        End If
 
    End With

    ReLineGrid

End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal row As Long, _
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

Private Sub XPBtnMove_Click(index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case index

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

Private Sub ¡¡¡¡_Click(index As Integer)

End Sub

Private Sub XPDtbTrans_Change()
    TxtNoteSerial.text = ""
 
End Sub
