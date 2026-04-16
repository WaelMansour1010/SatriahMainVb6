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
Begin VB.Form FrmLC_Search 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ «·«⁄ „«œ« "
   ClientHeight    =   8160
   ClientLeft      =   3525
   ClientTop       =   1470
   ClientWidth     =   13485
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   13485
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame7 
      Height          =   855
      Left            =   0
      TabIndex        =   117
      Top             =   7200
      Width           =   13215
      Begin ALLButtonS.ALLButton btnOk 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   12945
         _ExtentX        =   22834
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
         MICON           =   "FrmLC_search.frx":0000
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
      TabIndex        =   104
      Top             =   9840
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox txt_employee_count 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   9960
         TabIndex        =   106
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_emp_salary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   6360
         TabIndex        =   105
         Top             =   3000
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   1860
         Left            =   120
         TabIndex        =   107
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
         FormatString    =   ""
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
         TabIndex        =   108
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
         MICON           =   "FrmLC_search.frx":001C
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
         TabIndex        =   110
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·"
         Height          =   255
         Left            =   8040
         TabIndex        =   109
         Top             =   3120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame13 
      Height          =   2772
      Left            =   -90
      RightToLeft     =   -1  'True
      TabIndex        =   95
      Top             =   450
      Width           =   13305
      Begin VB.TextBox TXTEnd_user_id 
         Height          =   285
         Left            =   14280
         TabIndex        =   111
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtname 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         DataField       =   "project_name"
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   7920
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox PrimaryInvoiceNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   4932
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   96
         Top             =   3600
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DcVendor 
         Height          =   288
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   3972
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcoCurrency 
         Height          =   288
         Left            =   5640
         TabIndex        =   3
         Top             =   1800
         Width           =   4932
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcoBank 
         Height          =   288
         Left            =   5640
         TabIndex        =   4
         Top             =   2280
         Width           =   4932
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   312
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   1428
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   124780547
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   312
         Left            =   2640
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   1548
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   124780547
         CurrentDate     =   37140
      End
      Begin ALLButtonS.ALLButton btnSearch 
         Height          =   372
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   4092
         _ExtentX        =   7223
         _ExtentY        =   661
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
         MICON           =   "FrmLC_search.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker LastParcilDate 
         Height          =   312
         Left            =   2640
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1548
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   124780547
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker CloseDate 
         Height          =   312
         Left            =   240
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1428
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   124780547
         CurrentDate     =   37140
      End
      Begin MSDataListLib.DataCombo CountryID 
         Height          =   288
         Left            =   5640
         TabIndex        =   124
         Top             =   720
         Width           =   4932
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo LCTyperId 
         Height          =   288
         Left            =   5640
         TabIndex        =   125
         Top             =   1440
         Width           =   4968
         _ExtentX        =   8758
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
      Begin MSDataListLib.DataCombo DcBranch 
         Height          =   288
         Left            =   240
         TabIndex        =   126
         Top             =   1080
         Width           =   3972
         _ExtentX        =   7011
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
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›—⁄"
         Height          =   300
         Index           =   0
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   127
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ «·«‰ Â«¡"
         Height          =   252
         Index           =   21
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   120
         Top             =   1920
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "  «Œ— ‘Õ‰"
         Height          =   252
         Index           =   22
         Left            =   3972
         RightToLeft     =   -1  'True
         TabIndex        =   119
         Top             =   1920
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "«·»‰ﬂ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   115
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ «·⁄„·…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   114
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ «·› Õ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   4200
         TabIndex        =   113
         Top             =   1560
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Ê—œ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   112
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«·œÊ·…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   103
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«”„"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   4200
         TabIndex        =   102
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·›« Ê—… «·„»œ∆Ì…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   101
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "«·—ﬁ„ "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   100
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   " «—ÌŒ «·€·ﬁ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   1560
         TabIndex        =   99
         Top             =   1560
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "«·‰Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   98
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   720
         TabIndex        =   97
         Top             =   3720
         Width           =   1092
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "«·„’—Ê›« "
      Height          =   3615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   90
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
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   2340
         Left            =   240
         TabIndex        =   92
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
         FormatString    =   ""
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
         TabIndex        =   93
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
         MICON           =   "FrmLC_search.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·„’—Ê›« "
         Height          =   255
         Index           =   6
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   94
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
      TabIndex        =   68
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   690
         Index           =   2
         Left            =   360
         TabIndex        =   72
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
            Top             =   300
            Width           =   2310
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   900
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   300
            Width           =   1755
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   8805
            TabIndex        =   77
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
            TabIndex        =   78
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
            TabIndex        =   79
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
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   31
            Left            =   11985
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   0
            Width           =   2700
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈”„ «·’‰›"
            Height          =   255
            Index           =   30
            Left            =   9150
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   0
            Width           =   2640
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·’‰›"
            Height          =   255
            Index           =   29
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   0
            Width           =   1725
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ì—Ì«·"
            Height          =   255
            Index           =   28
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   0
            Width           =   2205
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì…"
            Height          =   255
            Index           =   27
            Left            =   2925
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   0
            Width           =   1515
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—"
            Height          =   255
            Index           =   26
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   0
            Width           =   1635
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   1905
         Left            =   240
         TabIndex        =   86
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
         FormatString    =   ""
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
         TabIndex        =   87
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
         MICON           =   "FrmLC_search.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·«’‰«›"
         Height          =   255
         Index           =   2
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label LblItemsCount 
         Caption         =   "Label27"
         Height          =   135
         Left            =   240
         TabIndex        =   88
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
      TabIndex        =   58
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
         TabIndex        =   59
         Top             =   2760
         Width           =   3015
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   2340
         Left            =   120
         TabIndex        =   60
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
         FormatString    =   ""
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
         TabIndex        =   61
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
         MICON           =   "FrmLC_search.frx":008C
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
         TabIndex        =   62
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
         MICON           =   "FrmLC_search.frx":00A8
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
         TabIndex        =   63
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
         MICON           =   "FrmLC_search.frx":00C4
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
         TabIndex        =   64
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
         MICON           =   "FrmLC_search.frx":00E0
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
         TabIndex        =   65
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.TextBox TxtModFlg 
      Height          =   285
      Left            =   14760
      TabIndex        =   57
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
      TabIndex        =   56
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
      TabIndex        =   55
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
      TabIndex        =   54
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
      TabIndex        =   52
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
      TabIndex        =   51
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   2400
      TabIndex        =   45
      Top             =   9840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4440
         TabIndex        =   47
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   720
      TabIndex        =   40
      Top             =   10200
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   43
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   4095
      Left            =   6960
      TabIndex        =   36
      Top             =   9600
      Width           =   1455
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   0
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   38
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
            MICON           =   "FrmLC_search.frx":00FC
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
            TabIndex        =   39
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
      TabIndex        =   30
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
      TabIndex        =   23
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   25
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
         TabIndex        =   24
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
      TabIndex        =   20
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
         TabIndex        =   35
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
         TabIndex        =   26
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   11520
      TabIndex        =   19
      Top             =   9600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   240
      TabIndex        =   16
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
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
      Height          =   2895
      Left            =   120
      TabIndex        =   13
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
      TabIndex        =   29
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
      MICON           =   "FrmLC_search.frx":0118
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
      TabIndex        =   116
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
      Format          =   179830787
      CurrentDate     =   37140
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   3360
      Width           =   13275
      Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
         Height          =   3450
         Left            =   0
         TabIndex        =   11
         Top             =   300
         Width           =   13170
         _cx             =   23230
         _cy             =   6085
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmLC_search.frx":0134
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   615
            Left            =   840
            TabIndex        =   118
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
         TabIndex        =   67
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
         MICON           =   "FrmLC_search.frx":0306
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   312
      Left            =   2640
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1548
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy/M/d"
      Format          =   179765251
      CurrentDate     =   37140
   End
   Begin VB.Image ImgFavoritesdd 
      Height          =   615
      Left            =   11160
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
      TabIndex        =   53
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   12480
      TabIndex        =   50
      Top             =   9840
      Width           =   2172
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "»ÕÀ «·«⁄ „«œ«                 "
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
      TabIndex        =   15
      Top             =   0
      Width           =   20655
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   11880
      TabIndex        =   14
      Top             =   9840
      Width           =   855
   End
End
Attribute VB_Name = "FrmLC_Search"
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


Private Sub BtnOK_Click()
Dim ID As Integer
Dim row As Integer
row = Fg_Journal.row
ID = val(Fg_Journal.TextMatrix(Fg_Journal.row, Fg_Journal.ColIndex("TblLCID")))
FrmLC.Search (ID)
End Sub
Private Sub btnSearch_Click()
ProgressBar1.Visible = True
: ProgressBar1.value = 10
   Set rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 
 
 
 StrSQL = StrSQL + " select * from tbllc where 1 =1   "
   
   
  If DCVendor.BoundText <> "" Then
        StrSQL = StrSQL + "  and  tbllc.vendorid =   " & val(DCVendor.BoundText)
  End If
   
 If txtID.text <> "" Then
        StrSQL = StrSQL + " and LCNO Like N'%" + txtID.text & "%'"
 End If
 
: ProgressBar1.value = 20
  
  
  If CountryID.SelectedItem <> -1 Then
         StrSQL = StrSQL & " and CountryID = " & val(CountryID.BoundText)
  End If
  
  If PrimaryInvoiceNo.text <> "" Then
        StrSQL = StrSQL + " and PrimaryInvoiceNo  = '" & DcAccount2.text & "'"
  End If
  
: ProgressBar1.value = 30

  If LCTyperId.BoundText <> "" Then
        StrSQL = StrSQL + " and LCTyperId = " & val(LCTyperId.BoundText)
  End If
  
  If dcoCurrency.BoundText <> "" Then
       StrSQL = StrSQL + " and tbllc.CurrencyId  = " & val(dcoCurrency.BoundText)
  End If
: ProgressBar1.value = 40

   If dcoBank.BoundText <> "" Then
        StrSQL = StrSQL + " and tbllc.BankId =  " & dcoBank.BoundText
   End If
: ProgressBar1.value = 50
   
  If TxtName.text <> "" Then
        StrSQL = StrSQL + " and name like '%" & TxtName.text & "%'"
  End If
  
   If Dcbranch.BoundText <> "" Then
        StrSQL = StrSQL + " and tbllc.BranchID =  " & Dcbranch.BoundText
   End If
   
   '///////////////////////////////////
'    If Not IsNull(Me.FromDate.value) Then
'          StrSQL = StrSQL + " and  FromDate >= " & SQLDate(Me.FromDate.value, True) & ""
'    End If
'
'    If Not IsNull(Me.FromDate.value) Then
'          StrSQL = StrSQL + " and  FromDate = " & SQLDate(Me.FromDate.value, True) & ""
'    End If
'
'
'   '///////////////////////////////////
'    If Not IsNull(Me.ToDate.value) Then
'          StrSQL = StrSQL + " and  ToDate  >= " & SQLDate(Me.ToDate.value, True) & ""
'    End If
'
'    If Not IsNull(Me.ToDate.value) Then
'          StrSQL = StrSQL + " and  ToDate = " & SQLDate(Me.ToDate.value, True) & ""
'    End If
    
       '///////////////////////////////////
'    If Not IsNull(Me.CloseDate.value) Then
'          StrSQL = StrSQL + " and  CloseDate  >= " & SQLDate(Me.CloseDate.value, True) & ""
'    End If
'
'    If Not IsNull(Me.CloseDate.value) Then
'          StrSQL = StrSQL + " and  CloseDate = " & SQLDate(Me.CloseDate.value, True) & ""
'    End If
'
'
'
'           '///////////////////////////////////
'    If Not IsNull(Me.LastParcilDate.value) Then
'          StrSQL = StrSQL + " and  LastParcilDate  >= " & SQLDate(Me.LastParcilDate.value, True) & ""
'    End If
'
'    If Not IsNull(Me.LastParcilDate.value) Then
'          StrSQL = StrSQL + " and  LastParcilDate = " & SQLDate(Me.LastParcilDate.value, True) & ""
'    End If
'
    
    
    
    
    
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

Private Sub btnSearch1_Click()

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


Private Sub DcbosubContractor_KeyUp(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 10
           FrmCompanySearch.show vbModal
           
        End If
        
End Sub







Private Sub Dcbranch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
 Else
      KeyAscii = 0
 End If
End Sub

Private Sub dcoBank_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyBack Then
 Else
      KeyAscii = 0
 End If
End Sub


Private Sub dcoCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyBack Then
 Else
      KeyAscii = 0
 End If
End Sub

Private Sub DcVendor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
 Else
      KeyAscii = 0
 End If
End Sub

Private Sub Fg_Journal_DblClick()
BtnOK_Click
End Sub

Private Sub Form_Load()
    On Error Resume Next
    TxtModFlg.text = "R"
    Set rs = New ADODB.Recordset

       Set Dcombos = New ClsDataCombos
      Dcombos.GetCustomersSuppliers 0, Me.DCVendor, True
    'Dcombos.GetItemsNames dcitems
    Dcombos.GetBanks Me.dcoBank
    Dcombos.GetCountriesNames Me.CountryID
     ' Dcombos.GetLCTypesName Me.LCTyperId
    Dcombos.GetLCTypesName Me.LCTyperId
    Dcombos.GetCUrrencyNames Me.dcoCurrency
    Dcombos.GetBranches Dcbranch
    FromDate.value = Date
    ToDate.value = Date
    LastParcilDate.value = Date
    CloseDate.value = Date

    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If
    
    dcoBank.BoundText = ""
    Dcbranch.BoundText = ""
    DCVendor.BoundText = ""
 

End Sub



Private Sub ChangeLang()
  
   
With Fg_Journal
 .TextMatrix(0, .ColIndex("LCNO")) = "No."
 .TextMatrix(0, .ColIndex("name")) = "Name"
.TextMatrix(0, .ColIndex("value")) = "Value"
.TextMatrix(0, .ColIndex("NoOfParcil")) = "Parcil No"
.TextMatrix(0, .ColIndex("PrimaryInvoiceNo")) = "Primary Inv. No"
.TextMatrix(0, .ColIndex("FromDate")) = "From Date"
.TextMatrix(0, .ColIndex("Todate")) = "To Date"
.TextMatrix(0, .ColIndex("CloseDate")) = "Close Date"
.TextMatrix(0, .ColIndex("LastParcilDate")) = "last ship. date"
End With
 
lbl(0).Caption = "Branch"
    Label20.Caption = "No."
    Label6.Caption = "Country"
    Label15.Caption = "Begining Inv. No."
    Label18.Caption = "Type"
    Label22.Caption = "Currency"
    Label7.Caption = "Bank"
    Label1.Caption = "Name"
    Label26.Caption = "Vendor"
    Label21.Caption = "From Date"
    Label25.Caption = "To Date"
    lbl(22).Caption = "last shipment date"
    lbl(21).Caption = "End Date"
    btnSearch.Caption = "Search"
    btnOk.Caption = "OK"
    Me.Caption = "LC Search"
    Label9.Caption = "LC Search"

    '
End Sub

Private Sub Retrive()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 2
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
     
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
   End If

'*************************************************


    '-----------------------------------------------------------------------------
        If Not (rs.BOF Or rs.EOF) Then
        
            rs.MoveFirst
    
            With Me.Fg_Journal
                .rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .rows - 1
                
                
                  .TextMatrix(i, .ColIndex("TblLCID")) = IIf(IsNull(rs("TblLCID").value), "", rs("TblLCID").value)
                    .TextMatrix(i, .ColIndex("LCNO")) = IIf(IsNull(rs("LCNO").value), "", rs("LCNO").value)
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                     .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(rs("value").value), "", rs("value").value)
                      .TextMatrix(i, .ColIndex("NoOfParcil")) = IIf(IsNull(rs("NoOfParcil").value), "", rs("NoOfParcil").value)
                      .TextMatrix(i, .ColIndex("PrimaryInvoiceNo")) = IIf(IsNull(rs("PrimaryInvoiceNo").value), "", rs("PrimaryInvoiceNo").value)
                        .TextMatrix(i, .ColIndex("fromdate")) = IIf(IsNull(rs("fromdate").value), "", rs("fromdate").value)
                        .TextMatrix(i, .ColIndex("todate")) = IIf(IsNull(rs("todate").value), "", rs("todate").value)
                        .TextMatrix(i, .ColIndex("closedate")) = IIf(IsNull(rs("closedate").value), "", rs("closedate").value)
                         .TextMatrix(i, .ColIndex("LastParcilDate")) = IIf(IsNull(rs("LastParcilDate").value), "", rs("LastParcilDate").value)
                rs.MoveNext
                Next i
         
            End With

    End If

 
    Exit Sub
ErrTrap:

End Sub

Private Sub LCTyperId_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyBack Then
 Else
      KeyAscii = 0
 End If
End Sub

Private Sub PrimaryInvoiceNo_KeyPress(KeyAscii As Integer)
  If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then

    ElseIf KeyAscii = vbKeyBack Then

    Else
      KeyAscii = 0
    End If
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)

'    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
'
'    ElseIf KeyAscii = vbKeyBack Then
'
'    Else
'      KeyAscii = 0
'    End If

End Sub
