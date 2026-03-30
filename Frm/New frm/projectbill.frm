VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form projectsbill 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›Ê« Ì—  «·„‘«—Ì⁄"
   ClientHeight    =   9780
   ClientLeft      =   3525
   ClientTop       =   1470
   ClientWidth     =   20355
   Icon            =   "projectbill.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   20355
   Begin VB.Frame Frame14 
      Caption         =   "«”„«¡ «·⁄«„·Ì‰ ›Ì «·„‘—Ê⁄"
      Height          =   3615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   128
      Top             =   9840
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox txt_employee_count 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   9960
         TabIndex        =   130
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_emp_salary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   6360
         TabIndex        =   129
         Top             =   3000
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   1860
         Left            =   120
         TabIndex        =   131
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
         FormatString    =   $"projectbill.frx":000C
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
         TabIndex        =   132
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
         MICON           =   "projectbill.frx":01E5
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
         TabIndex        =   134
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·"
         Height          =   255
         Left            =   8040
         TabIndex        =   133
         Top             =   3120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame13 
      Height          =   3252
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   102
      Top             =   600
      Width           =   20295
      Begin VB.TextBox txtManualNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   13560
         TabIndex        =   153
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDiscount2 
         Height          =   285
         Left            =   13560
         TabIndex        =   145
         Top             =   2880
         Width           =   1812
      End
      Begin VB.TextBox txtDiscount1 
         Height          =   285
         Left            =   13560
         TabIndex        =   146
         Top             =   2520
         Width           =   1812
      End
      Begin VB.ComboBox cboDiscount2 
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "projectbill.frx":0201
         Left            =   15480
         List            =   "projectbill.frx":020E
         TabIndex        =   152
         Top             =   2880
         Width           =   3012
      End
      Begin VB.ComboBox cboDiscount1 
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "projectbill.frx":0223
         Left            =   15480
         List            =   "projectbill.frx":0230
         TabIndex        =   151
         Top             =   2520
         Width           =   3012
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   17640
         TabIndex        =   150
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TXTsub_contractor_id 
         Height          =   375
         Left            =   13080
         TabIndex        =   138
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TXTEnd_user_id 
         Height          =   285
         Left            =   13800
         TabIndex        =   137
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "projectbill.frx":0245
         Left            =   13560
         List            =   "projectbill.frx":024F
         RightToLeft     =   -1  'True
         TabIndex        =   136
         Top             =   1800
         Visible         =   0   'False
         Width           =   4932
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   855
         Left            =   0
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   126
         Top             =   2160
         Width           =   11052
      End
      Begin VB.TextBox txtprojectname 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataField       =   "project_name"
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   5040
         TabIndex        =   109
         Top             =   600
         Width           =   6012
      End
      Begin VB.TextBox txtid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   360
         Left            =   17280
         TabIndex        =   108
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox bill_Type 
         DataSource      =   "Adodc1"
         Height          =   288
         ItemData        =   "projectbill.frx":0273
         Left            =   8760
         List            =   "projectbill.frx":027D
         TabIndex        =   107
         Top             =   1320
         Width           =   2292
      End
      Begin VB.ComboBox billto 
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "projectbill.frx":0290
         Left            =   13560
         List            =   "projectbill.frx":029A
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   1440
         Width           =   4932
      End
      Begin VB.TextBox DcAccount2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   13560
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   1080
         Width           =   4932
      End
      Begin VB.TextBox DcAccount1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   960
         Visible         =   0   'False
         Width           =   6012
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   103
         Top             =   3600
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   13560
         TabIndex        =   110
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   312
         Left            =   9624
         TabIndex        =   111
         Top             =   240
         Width           =   1428
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker dueDate1 
         Height          =   312
         Left            =   5160
         TabIndex        =   112
         Top             =   1320
         Width           =   2268
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Height          =   312
         Left            =   6600
         TabIndex        =   139
         Top             =   240
         Width           =   2412
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker dueDate 
         Height          =   312
         Left            =   8784
         TabIndex        =   142
         Top             =   1680
         Width           =   2268
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcbosubContractor 
         Height          =   285
         Left            =   13560
         TabIndex        =   147
         Top             =   2160
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14640
         TabIndex        =   154
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’„  œ›⁄Â „ﬁœ„…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   18600
         TabIndex        =   149
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’„ ÷„«‰ «·«⁄„«·"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   18600
         TabIndex        =   148
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„ﬁ«Ê· «·»«ÿ‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   18960
         TabIndex        =   144
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   10680
         TabIndex        =   143
         Top             =   1680
         Width           =   1812
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›—⁄"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   140
         Top             =   240
         Width           =   372
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "·œ›⁄Â „ÕœœÂ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   18960
         TabIndex        =   135
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "„·«ÕŸ« "
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   11760
         TabIndex        =   127
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   $"projectbill.frx":02B8
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1068
         Index           =   5
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   600
         Width           =   2892
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   18840
         TabIndex        =   124
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   10440
         TabIndex        =   123
         Top             =   720
         Width           =   2052
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·⁄„Ì· «·‰Â«∆Ì"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   18720
         TabIndex        =   122
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·„” Œ·’"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   18600
         TabIndex        =   121
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "«· «—ÌŒ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   10680
         TabIndex        =   120
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„” Œ·’ «·Ï"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   18960
         TabIndex        =   119
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ «·„” Œ·’"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   10680
         TabIndex        =   118
         Top             =   1440
         Width           =   1812
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ „ﬁ«Ê· «·»«ÿ‰"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   10200
         TabIndex        =   117
         Top             =   1080
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   1212
         Left            =   120
         Top             =   480
         Width           =   4812
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
         Height          =   252
         Index           =   4
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   120
         Width           =   1272
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   720
         TabIndex        =   114
         Top             =   3720
         Width           =   1092
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ï  «—ÌŒ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   7320
         TabIndex        =   113
         Top             =   1320
         Width           =   1332
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
         FormatString    =   $"projectbill.frx":0373
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
         MICON           =   "projectbill.frx":0481
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
            ButtonImage     =   "projectbill.frx":049D
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
         FormatString    =   $"projectbill.frx":0837
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
         MICON           =   "projectbill.frx":09FF
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
      TabIndex        =   60
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
         TabIndex        =   61
         Top             =   2760
         Width           =   3015
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   2340
         Left            =   120
         TabIndex        =   62
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
         FormatString    =   $"projectbill.frx":0A1B
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
         TabIndex        =   63
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
         MICON           =   "projectbill.frx":0D99
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
         TabIndex        =   64
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
         MICON           =   "projectbill.frx":0DB5
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
         TabIndex        =   65
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
         MICON           =   "projectbill.frx":0DD1
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
         TabIndex        =   66
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
         MICON           =   "projectbill.frx":0DED
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
         TabIndex        =   67
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.TextBox TxtModFlg 
      Height          =   285
      Left            =   3240
      TabIndex        =   59
      Top             =   120
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtrevenue_account 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   0
      TabIndex        =   55
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtsubaccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   5760
      TabIndex        =   54
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtendaccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   4680
      TabIndex        =   53
      Top             =   240
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
      TabIndex        =   51
      Top             =   9840
      Width           =   1095
   End
   Begin VB.TextBox txtdate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   7320
      TabIndex        =   50
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   7800
      TabIndex        =   39
      Top             =   9000
      Width           =   11295
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   1
         Left            =   7560
         TabIndex        =   40
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         MICON           =   "projectbill.frx":0E09
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
         Height          =   375
         Index           =   2
         Left            =   -1560
         TabIndex        =   41
         Top             =   120
         Width           =   1455
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
         MICON           =   "projectbill.frx":0E25
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
         Height          =   375
         Index           =   0
         Left            =   9960
         TabIndex        =   42
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         MICON           =   "projectbill.frx":0E41
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
         Height          =   375
         Index           =   5
         Left            =   -1560
         TabIndex        =   44
         Top             =   120
         Width           =   1095
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
         MICON           =   "projectbill.frx":0E5D
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
         Height          =   375
         Index           =   3
         Left            =   8760
         TabIndex        =   57
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         MICON           =   "projectbill.frx":0E79
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
         Height          =   375
         Index           =   6
         Left            =   6360
         TabIndex        =   58
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         MICON           =   "projectbill.frx":0E95
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
         Height          =   375
         Index           =   7
         Left            =   1320
         TabIndex        =   99
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÿ»«⁄Â «·›« Ê—…"
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
         MICON           =   "projectbill.frx":0EB1
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
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   100
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         MICON           =   "projectbill.frx":0ECD
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
         Height          =   375
         Index           =   9
         Left            =   5160
         TabIndex        =   101
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         MICON           =   "projectbill.frx":0EE9
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
         Height          =   375
         Index           =   10
         Left            =   3960
         TabIndex        =   125
         Top             =   120
         Width           =   1095
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
         MICON           =   "projectbill.frx":0F05
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
         Height          =   375
         Index           =   11
         Left            =   2520
         TabIndex        =   141
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "projectbill.frx":0F21
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
         TabIndex        =   43
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   2400
      TabIndex        =   34
      Top             =   9840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4440
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   720
      TabIndex        =   29
      Top             =   9720
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   4095
      Left            =   6960
      TabIndex        =   25
      Top             =   9600
      Width           =   1455
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   0
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   27
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
            MICON           =   "projectbill.frx":0F3D
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
            TabIndex        =   28
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
      TabIndex        =   19
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
      TabIndex        =   12
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   14
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
         TabIndex        =   13
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
      TabIndex        =   9
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
         TabIndex        =   24
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
         TabIndex        =   15
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   11520
      TabIndex        =   8
      Top             =   9600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   240
      TabIndex        =   5
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
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
      Bindings        =   "projectbill.frx":0F59
      Height          =   2895
      Left            =   120
      TabIndex        =   2
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
      TabIndex        =   18
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
      MICON           =   "projectbill.frx":0F6E
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
      Left            =   1665
      TabIndex        =   45
      Top             =   120
      Width           =   495
      _ExtentX        =   873
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
      ButtonImage     =   "projectbill.frx":0F8A
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
      Left            =   600
      TabIndex        =   46
      Top             =   120
      Width           =   495
      _ExtentX        =   873
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
      ButtonImage     =   "projectbill.frx":1324
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
      Left            =   2190
      TabIndex        =   47
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ButtonImage     =   "projectbill.frx":16BE
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
      Height          =   345
      Index           =   3
      Left            =   1125
      TabIndex        =   48
      Top             =   120
      Width           =   495
      _ExtentX        =   873
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
      ButtonImage     =   "projectbill.frx":1A58
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
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
   Begin VB.Frame Frame2 
      Caption         =   "»‰Êœ «·„‘—Ê⁄"
      Height          =   5172
      Left            =   -360
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   3840
      Width           =   20655
      Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
         Height          =   4380
         Left            =   -600
         TabIndex        =   69
         Top             =   240
         Width           =   21120
         _cx             =   37253
         _cy             =   7726
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
         FormatString    =   $"projectbill.frx":1DF2
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
         Index           =   0
         Left            =   18000
         TabIndex        =   70
         Top             =   4680
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "projectbill.frx":21D7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         MICON           =   "projectbill.frx":21F3
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
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "«Ã„«·Ì «·›« Ê—…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   52
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   12480
      TabIndex        =   49
      Top             =   9840
      Width           =   2172
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "›Ê« Ì—  «·„‘«—Ì⁄  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   20295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   11880
      TabIndex        =   3
      Top             =   9840
      Width           =   855
   End
End
Attribute VB_Name = "projectsbill"
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
    
    Dim accountdep As String

    If billto.ListIndex = -1 Then MsgBox "Õœœ «·„” Œ·’ „ﬁœ„  «·Ï „‰ ", vbCritical: Exit Function
        
        
     Dim J As Integer
     Dim found As Boolean
     
   J = Fg_Journal.FixedRows
     found = False
     
  For J = Fg_Journal.FixedRows To Fg_Journal.Rows - 1
  If Fg_Journal.TextMatrix(J, Fg_Journal.ColIndex("item")) <> "" Then
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
    Rs1("branch_no").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))

    Rs1("NoteID").value = val(note_id.text)
    Rs1("Note_Value").value = IIf(total.text = "", Null, val(total.text))
    Rs1("CusID").value = X
    Rs1("NoteType").value = 500
    Rs1("NoteType").value = 5000
    Rs1("NoteDate").value = XPDtbTrans.value
    Rs1("UserID").value = user_id

    If SystemOptions.UserInterface = ArabicInterface Then
        Rs1("remark").value = "„” Œ·’ —ﬁ„  :  " & txtid & Chr(13) & "  ··„‘—Ê⁄ " & txtprojectname.text
    Else
        Rs1("remark").value = "  Project Invoice No  :  " & txtid & Chr(13) & "  To Project " & txtprojectname.text
    End If
 
    '   Rs1("remark").value = "„” Œ·’ —ﬁ„ :     " & txtid & "    " & Chr(13) & "  ··„‘—Ê⁄  " & txtprojectname.text
    
    If TxtNoteSerial = "" Then
        TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
    End If
       
    Rs1("NoteSerial").value = TxtNoteSerial.text
    
    Rs1("NoteSerial1").value = Trim$(Me.txtid.text) '„”·”·
    Rs1("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    '  rs("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
     
    Rs1("sanad_year").value = year(XPDtbTrans.value)
    Rs1("sanad_month").value = Month(XPDtbTrans.value)
    Rs1("note_value_by_characters").value = WriteNo(Format(Me.total.text, "0.00"), 0, True, ".")
    
    Rs1.update
    
    rs("id").value = Me.txtid.text
    
    rs("bill_date").value = XPDtbTrans.value
  'branch_id
    rs("branch_no").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
    rs("project_no").value = IIf(Not IsNumeric(DataCombo2.BoundText), "", DataCombo2.BoundText)
    rs("project_name").value = IIf(IsNull(txtprojectname.text), "", txtprojectname.text)
    rs("Sub_user_name").value = IIf(IsNull(DcAccount1.text), "", DcAccount1.text)
    rs("End_user_name").value = IIf(IsNull(DcAccount2.text), "", DcAccount2.text)
    rs("End_user_account").value = IIf(IsNull(txtendaccount.text), "", txtendaccount.text)
    rs("Sub_user_account").value = IIf(IsNull(txtsubaccount.text), "", txtsubaccount.text)
    rs("revenue_account").value = IIf(IsNull(txtrevenue_account.text), "", txtrevenue_account.text)
    rs("bill_to").value = billto.ListIndex
    rs("bill_type").value = IIf(IsNull(bill_Type.text), "", bill_Type.text)
    rs("note_id").value = IIf(IsNull(note_id.text), "", note_id.text)
    rs("NoteSerial").value = IIf(IsNull(TxtNoteSerial.text), "", TxtNoteSerial.text)
    rs("total").value = IIf(Not IsNumeric(total.text), 0, total.text)
    '
         rs("dueDate").value = dueDate.value
rs("dueDate1").value = dueDate1.value


'*************************************************
rs("subContractorId").value = IIf(Not IsNumeric(DcbosubContractor.BoundText), Null, DcbosubContractor.BoundText)
rs("discount1ID").value = val(cboDiscount1.ListIndex)
rs("discount2ID").value = val(cboDiscount2.ListIndex)
rs("discount1value").value = val(txtDiscount1.text)
rs("discount2value").value = val(txtDiscount2.text)
rs("Remarks").value = Trim(TxtRemarks.text)
rs("ManualNo").value = Trim(txtManualNo.text)

 
'*************************************************


    rs.update

    
    Set RsDev = New ADODB.Recordset
 '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
               StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   Dim LngDevID As Long
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
  accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", X, "Account_code")
If billto.ListIndex = 0 Then
   
'    If accountdep = "" Then GoTo ll
    '«·ÿ—› «·„œÌ‰
    RsDev.AddNew
    
    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = 1
    RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
    RsDev("Value").value = val(Me.total.text)
    RsDev("Credit_Or_Debit").value = 0

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & txtid & Chr(13) & "  ··„‘—Ê⁄ " & txtprojectname.text
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & txtid & Chr(13) & "  To Project " & txtprojectname.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
   ' RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value ' DateValue(Now)
    RsDev("UserID").value = user_id
    RsDev("branch_id").value = my_branch
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev.update
'll:

    '«·ÿ—› «·œ«∆‰
    If Me.txtrevenue_account.text = "" Then Exit Function
    RsDev.AddNew
    RsDev("branch_id").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))
    RsDev("branch_id").value = my_branch
    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = 2
 'If SystemOptions.Revenueowed = True Then
    RsDev("Account_Code").value = Me.txtrevenue_account.text
 '   Else
    'RsDev("Account_Code").value = Me.txtrevenue_account .text
 '   End If
    
    RsDev("Value").value = val(Me.total.text)
    RsDev("Credit_Or_Debit").value = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & txtid & Chr(13) & "  ··„‘—Ê⁄ " & txtprojectname.text & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & txtid & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & txtid & Chr(13) & "  To Project " & txtprojectname.text
    End If

    RsDev("Notes_ID").value = val(note_id.text)
    RsDev("project_bill_no").value = val(txtid.text)
   RsDev("project_id").value = val(Me.DataCombo2.BoundText)
    RsDev("RecordDate").value = XPDtbTrans.value
    RsDev("UserID").value = user_id
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev.update

Else
'
        'If SystemOptions.SubContactorHave3Account = True Then
                Dim Discount1 As Double
                Dim Discount2 As Double
                Dim netvalue As Double
                Dim totalvalue As Double
                Dim AdvancedAccount As String
                Dim GuranteeAccount As String
                Dim line_no As Integer
                Dim des As String
                            If cboDiscount1.ListIndex = 0 Then
                                Discount1 = 0
                            ElseIf cboDiscount1.ListIndex = 1 Then
                                Discount1 = val(txtDiscount1) * val(total.text) / 100
                            ElseIf cboDiscount1.ListIndex = 2 Then
                                Discount1 = val(txtDiscount1)
                            End If
        
                            If cboDiscount2.ListIndex = 0 Then
                                Discount2 = 0
                            ElseIf cboDiscount2.ListIndex = 1 Then
                                Discount2 = val(txtDiscount2) * val(total.text) / 100
                            ElseIf cboDiscount2.ListIndex = 2 Then
                                Discount2 = val(txtDiscount2)
                            End If
               AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code2")
               GuranteeAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code1")
               accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbosubContractor.BoundText), "Account_code")
               line_no = 1
               Discount1 = Round(Discount1, 2)
                Discount2 = Round(Discount2, 2)
               netvalue = Round(val(total.text) - Discount1 - Discount2, 2)
               totalvalue = Round(val(total), 2)
               
                              des = "„’—Ê›«  «·„‘«—Ì⁄ " & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & txtid & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo
           If totalvalue > 0 Then '
    
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, expanses_account, totalvalue, 0, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Dcbranch.BoundText)) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
              
  '////////////////////////////////////////////////////
       '  End If
         
         
               des = "Œ’„ ÷„«‰ «⁄„«· " & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & txtid & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo
           If Discount1 > 0 Then '÷„«‰ «·«⁄„«·
    
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, GuranteeAccount, Discount1, 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Dcbranch.BoundText)) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
            
  
         End If
         
         des = "Œ’„ œ›⁄«  „ﬁœ„…   " & "   " & TxtRemarks & " —ﬁ„ «·”‰œ " & txtid & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & txtManualNo
           If Discount2 > 0 Then '
    
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, AdvancedAccount, Discount2, 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Dcbranch.BoundText)) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
              
  
         End If
         
         des = " «⁄„«·"
           If netvalue > 0 Then '
    
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, netvalue, 1, Msg & des & "  " & "··„‘—Ê⁄   " & txtprojectname.text, val(note_id.text), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Dcbranch.BoundText)) = False Then
                                        GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
               
  
         End If
         


End If
End If

    Dim Rs3 As New ADODB.Recordset
 '   Rs3.Open "project_bill_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
               StrSQL = "SELECT     * from dbo.project_bill_details Where (1 = -1)"
   Rs3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
          
    Dim i As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            '        Dim IntDEV_Type As Integer
            '        Dim SngDEV_Value As Single
            If .TextMatrix(i, .ColIndex("item")) <> "" Then

                Rs3.AddNew
                Rs3("bill_id").value = Me.txtid.text
                Rs3("item").value = IIf(.TextMatrix(i, .ColIndex("item")) = "", Null, .TextMatrix(i, .ColIndex("item")))
                Rs3("item_id").value = IIf(.TextMatrix(i, .ColIndex("item_id")) = "", Null, .TextMatrix(i, .ColIndex("item_id")))
                Rs3("cost").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("cost"))), 0, .TextMatrix(i, .ColIndex("cost")))
                Rs3("exe").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("exe"))), 0, .TextMatrix(i, .ColIndex("exe")))
                Rs3("percentage").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("percentage"))), 0, .TextMatrix(i, .ColIndex("percentage")))
                Rs3("exedate").value = IIf(Not IsDate(.TextMatrix(i, .ColIndex("exedate"))), Date, .TextMatrix(i, .ColIndex("exedate")))
                
                
                
               'Rs3("unit").value = IIf(.TextMatrix(i, .ColIndex("unit")) = "", Null, .TextMatrix(i, .ColIndex("unit")))
               Rs3("item_unit").value = IIf(.TextMatrix(i, .ColIndex("unit")) = "", Null, .TextMatrix(i, .ColIndex("unit")))
               Rs3("Unit_id").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Unit_id"))), 0, .TextMatrix(i, .ColIndex("Unit_id")))
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

    TxtModFlg.text = "R"

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

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
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
 
        Case 0
 
            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
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
                If DcAccount1.text = "" And billto.ListIndex = 1 Then MsgBox "this project have no subcontractor", vbCritical: Exit Sub

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

            imaged.Show

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

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            Dim Msg As String
            Dim StrSQL As String
 
            Dim RsTemp As New ADODB.Recordset
            StrSQL = "select * From ProjectBillBuy where Bill_id=" & val(txtid.text)
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                Msg = "·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  Â–« «·›« Ê—… " & Chr(13)
                Msg = Msg + "·«‰Â«  „ ⁄·ÌÂ« ⁄„·Ì«  ”œ«œ"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
          
            TxtModFlg.text = "E"

            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            Command1(1).Enabled = True

        Case 4

        Case 5

        Case 6
            Undo

        Case 9

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 7 'ÿ»«⁄Â «·›« Ê—…

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If
            print_report val(DataCombo2.BoundText)

        Case 8

            If DoPremis(Do_Print, Me.name, True) = False Then
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
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
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·›« Ê—… " & Chr(13)
            Msg = Msg + "·«‰Â«  „ ⁄·ÌÂ« ⁄„·Ì«  ”œ«œ"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (txtid.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            StrSQL = "Delete  Notes  where NoteSerial ='" & TxtNoteSerial & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
 
            If Not rs.RecordCount < 1 Then
                rs.Delete
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
     
        Exit Sub
    End If
 
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Private Sub Command2_Click()
End Sub

Function GET_PROJECT_DATA()
    On Error Resume Next

    If DataCombo2.text = "" Then Exit Function
    Dim My_SQL As String

    My_SQL = "select * from projects where id =" & DataCombo2.BoundText
 
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    txtprojectname.text = rec.Fields("Project_name").value
    txtsubaccount.text = IIf(IsNull(rec.Fields("sub_contractor_Account").value), "", rec.Fields("sub_contractor_Account").value)
    DcAccount1.text = IIf(IsNull(rec.Fields("sub_contractor_name").value), "", rec.Fields("sub_contractor_name").value)
    txtendaccount.text = IIf(IsNull(rec.Fields("End_user_Account").value), "", rec.Fields("End_user_Account").value)
    DcAccount2.text = IIf(IsNull(rec.Fields("End_user_name").value), "", rec.Fields("End_user_name").value)
 If SystemOptions.Revenueowed = True Then
    txtrevenue_account.text = IIf(IsNull(rec.Fields("legal").value), "", rec.Fields("legal").value) 'Õ”«» «·„” Œ·’« \
  Else
      txtrevenue_account.text = IIf(IsNull(rec.Fields("REVENUE_account").value), "", rec.Fields("REVENUE_account").value) 'Õ”«» «·«Ì—«œ« \

  End If
  
TXTEnd_user_id.text = IIf(IsNull(rec.Fields("End_user_id").value), "", rec.Fields("End_user_id").value) '—ﬁ„ «·⁄„Ì· «·‰Â«∆Ì
TXTsub_contractor_id.text = IIf(IsNull(rec.Fields("sub_contractor_id").value), "", rec.Fields("sub_contractor_id").value) '—ﬁ„   „ﬁ«Ê· «·»«ÿ‰

 expanses_account = IIf(IsNull(rec.Fields("expanses_account").value), "", rec.Fields("expanses_account").value) 'Õ”«»  «·„’—Ê›« \

    'My_SQL = "  select net,des from projects_des  where project_id='" & DataCombo2.BoundText & "'"
    'fill_combo DataCombo5, My_SQL

End Function

Private Sub DataCombo2_Change()
    GET_PROJECT_DATA
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    GET_PROJECT_DATA
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
            Adodc7.Recordset.Delete
            DataGrid2.Refresh
            Command1_Click (1)
            total.text = gettotal(txtid.text)

        End If

    End If

End Sub

Function gettotal(X As String) As Double
    Dim My_SQL As String

    My_SQL = "  select Sum(exe) as total  from project_bill_details where bill_id=" & X

    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    gettotal = IIf(IsNull(rec.Fields("total").value), 0, rec.Fields("total").value)

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
             FrmProjectSearch.Show vbModal
           
        End If
        
        
End Sub

Private Sub DcbosubContractor_KeyUp(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 10
           FrmCompanySearch.Show vbModal
           
        End If
        
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

    With Fg_Journal

        Select Case .ColKey(Col)
 
            Case "item"
                StrAccountCode = .ComboItem
       
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("item"), False, True)
                .TextMatrix(Row, .ColIndex("item")) = StrAccountCode
            
                If StrAccountCode <> "" Then
                    StrSQL = "SELECT   line_no, oprid,des, net, project_id  ,[unit] ,[Quantity],[Price] ,[Pre_Quantity] ,[Pre_Value],[Pre_Percent] ,[Curr_Quantity]  ,[Curr_value] ,[curr_Percent] ,[tot_quantity] ,[tot_value] ,[tot_percent]   from dbo.projects_des  WHERE fullcode='" & .ComboData & "'"  ' project_id =" & Val(DataCombo2.BoundText) & "and line_no=" & Val(.ComboItem)
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
                    .TextMatrix(Row, .ColIndex("cost")) = IIf(IsNull(Rs1("net").value), 0, Rs1("net").value)
                    .TextMatrix(Row, .ColIndex("exe")) = 0
                    .TextMatrix(Row, .ColIndex("percentage")) = 0
                    .TextMatrix(Row, .ColIndex("item_id")) = .ComboData
                    
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
 
            Case "exe"
            
            Case "Unit"
               If StrAccountCode <> "" Then
               .TextMatrix(Row, .ColIndex("unit_id")) = .ComboData
               End If
        End Select

        '  Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

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
    StrSQL = StrSQL + "SELECT *  From dbo.project_billl Order by ID"
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    '

    'first_run = True
    Dim My_SQL As String
 
    My_SQL = "  select id,Fullcode from Projects"
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
    Set NewGrid.txttotal = XPTxtSum
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
    NewGrid.fillgrid
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    ChangeLang
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Command1_Click (0)
    End If

End Sub

Function ChangeLang()

    If SystemOptions.UserInterface = EnglishInterface Then
        temp = XPBtnMove(1).left
        XPBtnMove(1).left = XPBtnMove(2).left
        XPBtnMove(2).left = temp
Label26.Caption = "Branch"

        temp = XPBtnMove(0).left
        XPBtnMove(0).left = XPBtnMove(3).left
        XPBtnMove(3).left = temp
        SetInterface Me
        Me.Caption = "Project Invoice"
        Label9.Caption = Me.Caption

        Label20.Caption = "Bill No."
        Label25.Caption = "Date"
lbl(5).Caption = "Is the extract of the project, including implementation bill provides a total value, or a statement of what has been implemented in detail at the level of the terms of each individual process"
        Label6.Caption = "Project Code"
        Label1.Caption = "Project Name"
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
  
        Command1(0).Caption = "new"
        Command1(1).Caption = "save"
        Command1(2).Caption = "Attachments"
        Command1(3).Caption = "Edit"
        Command1(6).Caption = "Delete"
  
        SuperLabel2.text = "Search"
        Command1(4).Caption = "By ID"
        Command1(5).Caption = "Search"
  Command1(11).Caption = "Attachement"
        Adodc1.Caption = "move"
  
        With Fg_Journal
            .TextMatrix(0, .ColIndex("LineNo")) = "Index"
            .TextMatrix(0, .ColIndex("Item_ID")) = "Term#"

            .TextMatrix(0, .ColIndex("item")) = "Term Desc."
            .TextMatrix(0, .ColIndex("cost")) = "cost"
            .TextMatrix(0, .ColIndex("exe")) = "exe"
            .TextMatrix(0, .ColIndex("percentage")) = "percentage"
            .TextMatrix(0, .ColIndex("exedate")) = "exe date"

  .TextMatrix(0, .ColIndex("Unit")) = "Unit"
  .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
.TextMatrix(0, .ColIndex("Price")) = "Price"
.TextMatrix(0, .ColIndex("Pre_Quantity")) = "Pre. Exe. Quantity"
.TextMatrix(0, .ColIndex("Pre_Value")) = "Pre. exe value "
.TextMatrix(0, .ColIndex("Pre_Percent")) = "Pre. exe percentage"
.TextMatrix(0, .ColIndex("Curr_Quantity")) = " Current exe Quantity"
.TextMatrix(0, .ColIndex("Curr_value")) = " Current exe value"
.TextMatrix(0, .ColIndex("curr_Percent")) = "Current exe percentage"
.TextMatrix(0, .ColIndex("tot_quantity")) = "Total Quantity"
.TextMatrix(0, .ColIndex("tot_value")) = "Total Value"
.TextMatrix(0, .ColIndex("tot_percent")) = "Total Percent"


        End With

        opr_items(0).Caption = "View Term Operations"
        Frame11.Caption = "Term Operaions"
 
        Label27.Caption = "Labors Count"
        Label24.Caption = "Total"

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

        With Me.VSFlexGrid3
            .TextMatrix(0, .ColIndex("LineNo")) = "Index"
            .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Names"
            .TextMatrix(0, .ColIndex("value")) = "value"

            .TextMatrix(0, .ColIndex("des")) = "des"
 
        End With
Label21.Caption = "Achievement Date"
Label32.Caption = "deduct advance Payment"
Label31.Caption = "deduct ensure business "
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
    VSFlexGrid2.Rows = 2
    VSFlexGrid2.Enabled = True
    txt_opr_total.text = 0
          
    StrSQL = "select * from terms_operations_project_bill where term_fullcode='" & Item_ID & "' and bill_id=" & val(Me.txtid.text)
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

            Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        
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
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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
            Me.txt_emp_salary.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
            Me.txt_employee_count.text = .Aggregate(flexSTCount, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
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
            Me.txt_expenses_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
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

                For i = .FixedRows To .Rows - 1

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

Private Sub terms_operations_Click(Index As Integer)

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
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Enabled = True
 
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where    bill_id is null and (payed =1 )  and opr_fullcode='" & current_opr & "' and Transaction_Date<='" & SQLDate(DTPicker1.value) & "'"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = IIf(IsNull(RsDetails("Quantity")), 0, (RsDetails("Quantity").value)) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value))
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Sub

Private Sub txtId_Change()
    ' "select * from project_bill_details where bill_id=" & Val(txtid.text)

End Sub

Private Sub Retrive()
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
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
    Else

        If Lngid <> 0 Then
            rs.find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
Me.Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    txtid.text = IIf(IsNull(rs("id").value), 0, val(rs("id").value))

    XPDtbTrans.value = IIf(IsNull(rs("bill_date").value), Date, rs("bill_date").value)
dueDate.value = IIf(IsNull(rs("dueDate").value), Date, rs("dueDate").value)
dueDate1.value = IIf(IsNull(rs("dueDate1").value), Date, rs("dueDate1").value)

    DataCombo2.BoundText = IIf(IsNull(rs("project_no").value), "", rs("project_no").value)
'*************************************************
DcbosubContractor.BoundText = IIf(IsNull(rs("subContractorId").value), "", rs("subContractorId").value)
txtDiscount1.text = IIf(IsNull(rs("discount1value").value), 0, (rs("discount1value").value))
txtDiscount2.text = IIf(IsNull(rs("discount2value").value), 0, (rs("discount2value").value))

cboDiscount1.ListIndex = IIf(IsNull(rs("discount1ID").value), 0, (rs("discount1ID").value))
cboDiscount2.ListIndex = IIf(IsNull(rs("discount2ID").value), 0, (rs("discount2ID").value))

'*************************************************


    txtprojectname.text = IIf(IsNull(rs("project_name").value), "", rs("project_name").value)
    DcAccount1.text = IIf(IsNull(rs("Sub_user_name").value), "", rs("Sub_user_name").value)
    DcAccount2.text = IIf(IsNull(rs("End_user_name").value), "", rs("End_user_name").value)

    txtendaccount.text = IIf(IsNull(rs("End_user_account").value), "", rs("End_user_account").value)
    txtsubaccount.text = IIf(IsNull(rs("Sub_user_account").value), "", rs("Sub_user_account").value)
    txtrevenue_account.text = IIf(IsNull(rs("revenue_account").value), "", rs("revenue_account").value)

    'DcAccount4.text = IIf(IsNull(Rs("sub_contractor_name").value), "", Rs("sub_contractor_name").value)

    'DcAccount2.text = IIf(IsNull(Rs("End_user_name").value), "", Rs("End_user_name").value)

    billto.ListIndex = IIf(IsNull(rs("bill_to").value), -1, rs("bill_to").value)
    bill_Type.text = IIf(IsNull(rs("bill_type").value), "", rs("bill_type").value)
    Me.note_id.text = IIf(IsNull(rs("note_id").value), "", rs("note_id").value)
    TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    TxtRemarks.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
        txtManualNo.text = IIf(IsNull(rs("ManualNo").value), "", rs("ManualNo").value)
        

'rs("Remarks").value = Trim(TxtRemarks.text)
'rs("ManualNo").value = Trim(txtManualNo.text)

    total.text = IIf(IsNull(rs("total").value), 0, rs("total").value)

    'Exit Sub

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "SELECT     item_id,id, project_no, item, cost, exe, percentage, exedate, bill_id,item_unit ,Unit_id,Quantity,Price,Pre_Quantity,Pre_Value,Pre_Percent,Curr_Quantity,Curr_value,curr_Percent,tot_quantity,tot_value,tot_percent "
        StrSQL = StrSQL + " from dbo.project_bill_details "
        StrSQL = StrSQL + " Where bill_id =" & Me.txtid.text
    
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
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

        For i = .FixedRows To .Rows - 1

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

        Me.total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("exe"), .Rows - 1, .ColIndex("exe"))
         
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

        Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
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

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "id='" & val(txtid.text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub ¡¡¡¡_Click(Index As Integer)

End Sub

Private Sub XPDtbTrans_Change()
    TxtNoteSerial.text = ""
 
End Sub
