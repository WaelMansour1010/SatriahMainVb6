VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRegisterHoliday 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12495
   Icon            =   "frmRegisterHoliday.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12495
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   1935
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   80
      Top             =   5520
      Width           =   12465
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
         Height          =   1320
         Left            =   120
         TabIndex        =   81
         Top             =   480
         Width           =   12315
         _cx             =   21722
         _cy             =   2328
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegisterHoliday.frx":000C
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
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄Âœ «·⁄Ì‰Ì…"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   14
         Left            =   11040
         TabIndex        =   82
         Top             =   120
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   2535
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   480
      Width           =   12495
      Begin VB.TextBox Txtjopstatusid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox Contract_period11 
         Height          =   315
         ItemData        =   "frmRegisterHoliday.frx":014A
         Left            =   13680
         List            =   "frmRegisterHoliday.frx":014C
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox dcjopstatus1 
         Height          =   315
         ItemData        =   "frmRegisterHoliday.frx":014E
         Left            =   6480
         List            =   "frmRegisterHoliday.frx":0150
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox TxtOther 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox TxtTelephone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         Locked          =   -1  'True
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TxtLogConract 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox TXTCode 
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
         Height          =   315
         Left            =   10080
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox Contract_period 
         Height          =   315
         ItemData        =   "frmRegisterHoliday.frx":0152
         Left            =   13200
         List            =   "frmRegisterHoliday.frx":015C
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   8160
         TabIndex        =   37
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   93782017
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcBranches 
         Bindings        =   "frmRegisterHoliday.frx":016A
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
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
      Begin MSDataListLib.DataCombo DcNational 
         Height          =   315
         Left            =   3360
         TabIndex        =   39
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCEmP 
         Height          =   315
         Left            =   6480
         TabIndex        =   45
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· «Œ Ì«— «·„‰œÊ»"
         Top             =   720
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDepartMen 
         Height          =   315
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DateSatrContrac 
         Height          =   315
         Left            =   4200
         TabIndex        =   50
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   93782017
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker DateEndContrac 
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   93782017
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcbJobs 
         Height          =   315
         Left            =   6480
         TabIndex        =   58
         Top             =   1215
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDirctManger 
         Height          =   315
         Left            =   120
         TabIndex        =   60
         Top             =   1215
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbMangment 
         Height          =   315
         Left            =   3360
         TabIndex        =   62
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTP_Date 
         Height          =   315
         Left            =   4200
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   93782019
         CurrentDate     =   37140
      End
      Begin MSDataListLib.DataCombo dctype 
         Height          =   315
         Left            =   10080
         TabIndex        =   83
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ—Ï"
         Height          =   285
         Index           =   9
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   285
         Index           =   23
         Left            =   10080
         TabIndex        =   69
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—« » «·Õ«·Ì"
         Height          =   285
         Index           =   29
         Left            =   11190
         TabIndex        =   68
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ   —ﬂ «·Œœ„…"
         Height          =   285
         Index           =   0
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2160
         Width           =   1650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«‘⁄«—"
         Height          =   285
         Index           =   5
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«œ«—…"
         Height          =   285
         Index           =   10
         Left            =   5370
         TabIndex        =   63
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œÌ— «·„»«‘— "
         Height          =   285
         Index           =   9
         Left            =   2310
         TabIndex        =   61
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…"
         Height          =   285
         Index           =   8
         Left            =   8910
         TabIndex        =   59
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Â« ›"
         Height          =   285
         Index           =   7
         Left            =   11310
         TabIndex        =   57
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œ… «·⁄ﬁœ"
         Height          =   405
         Index           =   6
         Left            =   3360
         TabIndex        =   55
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Â«Ì… «·⁄ﬁœ"
         Height          =   285
         Index           =   5
         Left            =   1230
         TabIndex        =   53
         Top             =   1695
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œ«Ì… «·⁄„·"
         Height          =   285
         Index           =   3
         Left            =   5370
         TabIndex        =   51
         Top             =   1695
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ”„"
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   48
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Õ«·Â"
         Height          =   285
         Index           =   4
         Left            =   11505
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   285
         Index           =   1
         Left            =   11385
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   720
         Width           =   930
      End
      Begin VB.Label lblbr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»"
         Height          =   285
         Index           =   4
         Left            =   11310
         TabIndex        =   42
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         Height          =   285
         Index           =   2
         Left            =   5370
         TabIndex        =   41
         Top             =   705
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   9030
         TabIndex        =   40
         Top             =   255
         Width           =   1005
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2685
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3000
      Width           =   12525
      Begin VB.TextBox TxtRemarkss 
         Alignment       =   1  'Right Justify
         Height          =   825
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   72
         Top             =   1680
         Width           =   11055
      End
      Begin VB.TextBox TxtDes 
         Alignment       =   1  'Right Justify
         Height          =   1425
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   70
         Top             =   120
         Width           =   5415
      End
      Begin VB.ComboBox CmbType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmRegisterHoliday.frx":017F
         Left            =   2280
         List            =   "frmRegisterHoliday.frx":018F
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   -450
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox TxtSerialx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4425
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   -330
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.TextBox Txt_NotEndWork 
         Alignment       =   1  'Right Justify
         Height          =   1425
         Left            =   6480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   120
         Width           =   4695
      End
      Begin VB.TextBox TxtSerial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   -960
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo DCJob 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·ÊŸÌ€…"
         Top             =   -360
         Visible         =   0   'False
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ« "
         Height          =   285
         Index           =   8
         Left            =   11400
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ﬁ —«Õ« "
         Height          =   645
         Index           =   7
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·⁄„·Ì…"
         Height          =   195
         Index           =   3
         Left            =   5685
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   -570
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”»«» «·«” ﬁ«·…"
         Height          =   285
         Index           =   6
         Left            =   11370
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   1050
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   12435
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   4
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483624
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„” Œœ„"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   13
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Text            =   "modflag"
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRegisterHoliday.frx":01A8
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRegisterHoliday.frx":0542
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRegisterHoliday.frx":08DC
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRegisterHoliday.frx":0C76
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRegisterHoliday.frx":1010
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRegisterHoliday.frx":13AA
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRegisterHoliday.frx":1744
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRegisterHoliday.frx":1CDE
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmRegisterHoliday.frx":2078
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   7
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmRegisterHoliday.frx":2412
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmRegisterHoliday.frx":27AC
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmRegisterHoliday.frx":2B46
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ·» «‰Â«¡ Œœ„…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   5070
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2040
         Picture         =   "frmRegisterHoliday.frx":2EE0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   780
      Left            =   15
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7425
      Width           =   12480
      _cx             =   22013
      _cy             =   1376
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
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   1
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
      Begin ImpulseButton.ISButton btnNew 
         Height          =   330
         Left            =   8055
         TabIndex        =   20
         Top             =   315
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÃœÌœ"
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
         ButtonImage     =   "frmRegisterHoliday.frx":6B48
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   6510
         TabIndex        =   21
         Top             =   315
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ›Ÿ"
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
         ButtonImage     =   "frmRegisterHoliday.frx":6EE2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   7275
         TabIndex        =   22
         Top             =   315
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ⁄œÌ·"
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
         ButtonImage     =   "frmRegisterHoliday.frx":727C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   5745
         TabIndex        =   23
         Top             =   315
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " —«Ã⁄"
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
         ButtonImage     =   "frmRegisterHoliday.frx":7616
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   4980
         TabIndex        =   24
         Top             =   315
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–›"
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
         ButtonImage     =   "frmRegisterHoliday.frx":79B0
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   3240
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   315
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
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
         ButtonImage     =   "frmRegisterHoliday.frx":7F4A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   11685
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   105
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
         BackColor       =   14871017
         FontSize        =   9.75
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "frmRegisterHoliday.frx":82E4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   4125
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   315
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
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
         ButtonImage     =   "frmRegisterHoliday.frx":867E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   2385
         TabIndex        =   28
         Top             =   315
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Œ—ÊÃ"
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
         ButtonImage     =   "frmRegisterHoliday.frx":8A18
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9120
         TabIndex        =   77
         Top             =   240
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ… : "
         Height          =   270
         Index           =   11
         Left            =   11505
         TabIndex        =   78
         Top             =   315
         Width           =   900
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   105
         Width           =   540
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   930
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   105
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   2505
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   105
         Width           =   975
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Grid 
      Height          =   3405
      Left            =   15
      TabIndex        =   33
      Top             =   9210
      Width           =   7965
      _cx             =   14049
      _cy             =   6006
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRegisterHoliday.frx":8DB2
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
End
Attribute VB_Name = "FrmRegisterHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim StrSQL As String
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
 MySQL = "SELECT     dbo.TBLRegisterHoliday.id, dbo.TBLRegisterHoliday.EmpID, dbo.TBLRegisterHoliday.EndWork, dbo.TBLRegisterHoliday.Notsstkala, "
MySQL = MySQL & "                      dbo.TBLRegisterHoliday.jopstatusid, dbo.jopstatus.name, dbo.jopstatus.namee, dbo.TBLRegisterHoliday.BranchID, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TBLRegisterHoliday.NationID, dbo.Nationality.name AS Nationaliname, dbo.Nationality.namee AS NationalinameE,"
MySQL = MySQL & "                      dbo.TBLRegisterHoliday.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TBLRegisterHoliday.DirctMangerID,"
MySQL = MySQL & "                      dbo.TBLRegisterHoliday.DepartMentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
MySQL = MySQL & "                      dbo.TBLRegisterHoliday.MangmentID, dbo.TblSection.name AS Mangname, dbo.TblSection.namee AS MangnameE, dbo.TBLRegisterHoliday.Telephone,"
MySQL = MySQL & "                      dbo.TBLRegisterHoliday.LogConract, dbo.TBLRegisterHoliday.Other, dbo.TBLRegisterHoliday.Jopstatus1, dbo.TBLRegisterHoliday.DateSatrContrac,"
MySQL = MySQL & "                      dbo.TBLRegisterHoliday.DateEndContrac, dbo.TBLRegisterHoliday.Remarkss, dbo.TBLRegisterHoliday.Des, dbo.TBLRegisterHoliday.RecordDate,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name2, TblEmployee_1.Emp_Name3,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name4, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, TblEmployee_1.Emp_Namee1, TblEmployee_1.Emp_Namee2,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee3, TblEmployee_1.Emp_Namee4, TblEmployee_2.Emp_Name AS MangerEmp_Name,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Namee AS MangerEmp_NameE, dbo.TBLRegisterHoliday.Salary"
MySQL = MySQL & " FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TBLRegisterHoliday ON TblEmployee_1.Emp_ID = dbo.TBLRegisterHoliday.EmpID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.TBLRegisterHoliday.DirctMangerID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblSection ON dbo.TBLRegisterHoliday.MangmentID = dbo.TblSection.Id ON"
MySQL = MySQL & "                      dbo.TblEmpDepartments.DeparmentID = dbo.TBLRegisterHoliday.DepartMentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.jopstatus ON dbo.TBLRegisterHoliday.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TBLRegisterHoliday.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Nationality ON dbo.TBLRegisterHoliday.NationID = dbo.Nationality.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TBLRegisterHoliday.BranchID = dbo.TblBranchesData.branch_id"
 MySQL = MySQL & "  Where (dbo.TBLRegisterHoliday.id =" & val(TxtVac_ID.text) & ")"
 


  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\" & "EndserviceWork.rpt"
        Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\" & "EndserviceWorkE.rpt"
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
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(lbl(23).Caption), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function
Private Sub BtnCancel_Click()
    Unload Me
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & "ﬂÊœ  «·„ÊŸ›" & TxtCode.text & Chr(13) & "   «”„ «·„ÊŸ› " & DCEmP & Chr(13) & "  «·Õ«·… " & dctype.text & Chr(13) & "   «·”»» " & Txt_NotEndWork
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Sales Person Code" & TxtCode.text & Chr(13) & "    Sales Person Name   " & DCEmP & Chr(13) & "  Status " & dctype & Chr(13) & "   Reasons " & Txt_NotEndWork
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
    
End Function

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
If CheckEndService = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–›  „ ⁄„· ‰Â«Ì… Œœ„…"
Else
MsgBox "Can Not Delete this is Requet Already in End Service"
End If
Exit Sub
End If
    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
        MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        End If

        If MSGType = vbYes Then
      '  If val(Txtjopstatusid.text) = 0 Or val(Txtjopstatusid.text) = 1 Then
      '   StrSQL = "update TblEmployee Set   jopstatusid=1 ,workstate=1  where Emp_ID=" & val(DcEmp.BoundText)
      '    Cn.Execute StrSQL
    ' Else
    '    StrSQL = "update TblEmployee Set   jopstatusid=" & val(Txtjopstatusid.text) & " ,workstate=0  where Emp_ID=" & val(DcEmp.BoundText)
    '      Cn.Execute StrSQL
    ' End If
    Cn.Execute "delete from  dbo.TBLRegisterHolidayDet Where RegID=" & val(TxtVac_ID.text) & ""
    
            RsSavRec.find "id=" & val(TxtVac_ID.text), , adSearchForward, 1
            StrSQL = "update TblEmployee Set ChekStkala=0 , EndWork=Null,Notsstkala=  Null where Emp_ID=" & val(DCEmP.BoundText)
            Cn.Execute StrSQL

            CuurentLogdata ("D")
            RsSavRec.delete
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Record Deleted", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    FiLLTXT

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String
If CheckEndService = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·  „ ⁄„· ‰Â«Ì… Œœ„…"
Else
MsgBox "Can Not Edit this is Requet Already in End Service"
End If
Exit Sub
End If

    If DoPremis(Do_Edit, Me.name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.text <> "" Then
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
        Frm2.Enabled = True
    DcBranches.BoundText = Current_branch
        CuurentLogdata
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If
XPDtbTrans.value = Date
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    '-----------------------------------
    Me.TxtVac_ID.text = ""
 
    Me.DCEmP.BoundText = ""
   ' Me.dcjopstatus.BoundText = ""
    Me.Txt_NotEndWork.text = ""
         clear_all Me
         DTP_Date.value = Date
     Me.DCboUserName.BoundText = user_id
    '-----------------------------------
    TxtModFlg.text = "N"
DcBranches.BoundText = Current_branch
    My_SQL = "TBLRegisterHoliday"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0
 
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext

        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious

    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrint_Click()
print_report
End Sub

Private Sub btnQuery_Click()
Unload FrmSerachRegisterEndService
Load FrmSerachRegisterEndService
FrmSerachRegisterEndService.ind = 0
FrmSerachRegisterEndService.show
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next
    
If val(DCEmP.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «”„ «·„ÊŸ›"
Else
MsgBox "Please Entere Employee"
End If
DCEmP.SetFocus
Exit Sub
End If
If val(dctype.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«—  Õ«·… «·„ÊŸ›"
Else
MsgBox "Please Select State Employee"
End If
dctype.SetFocus
Exit Sub
End If

    '------------------------------ check if Empcode exist ----------------------
 
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title

End Sub
 
Private Sub BtnUndo_Click()
If val(TxtVac_ID.text) <> 0 Then
    FindRec val(TxtVac_ID.text)
    Else
    RsSavRec.MoveFirst
    FiLLTXT
    End If
    Me.TxtModFlg.text = "R"
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click

    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub dcEmp_Change()

    If val(Me.DCEmP.BoundText) = 0 Then Exit Sub
    Me.TxtCode.text = get_EMPLOYEE_Data(val(Me.DCEmP.BoundText), "Emp_Code")
    Dcemp_Click (0)
    
End Sub

Private Sub Dcemp_Click(Area As Integer)
    On Error Resume Next
       If val(DCEmP.BoundText) = 0 Then Exit Sub
        
   If Me.TxtModFlg = "R" Then Exit Sub
   
   
    Dim StrSQL As String

        Dim IssueDate As Date
        Dim depid As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_Code  As String
        Dim Balance As String
        Dim ProjectID As Integer
 Dim endiqama As String
        Dim national As String
        Dim endContractPerMonth As Double
       Dim BignDateWork As Date
       Dim JobTypeName As String
       Dim JobTypeIDIQ As Integer
       Dim Contract_period As Integer
     Dim Contract_periodno As Integer
   Dim dcjopstatus As Integer
   Dim mangerid As Integer
 Dim LastDate As Date
 Dim RegionID As Integer
 Dim Emp_Phone As String
 Dim Contract_date As Date
        get_employee_information val(Me.DCEmP.BoundText), IssueDate, depid, specid, JobTypeID, gradeID, Account_code2, Account_Code, endContractPerMonth, national, mangerid, , ProjectID, , , , , endiqama, , BignDateWork, LastDate, JobTypeName, Contract_period, Contract_periodno, , dcjopstatus, JobTypeIDIQ, , , , , Emp_Phone, Contract_date, RegionID
       lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DCEmP.BoundText), "", 0)
       DcbDepartMen.BoundText = depid
       DcbJobs.BoundText = JobTypeID
       DcNational.text = national
       If Me.TxtModFlg.text = "N" Then
       Txtjopstatusid.text = dcjopstatus
       End If
     
       Me.DcbDirctManger.BoundText = mangerid
       DcbMangment.BoundText = RegionID
       TxtTelephone.text = Emp_Phone
       DateSatrContrac.value = BignDateWork
       Me.Contract_period11.ListIndex = Contract_period
      Me.TxtLogConract.text = Contract_periodno & "     " & Me.Contract_period11.text
      If Me.Contract_period11.ListIndex = 0 Then
      DateEndContrac.value = DateAdd("m", Contract_periodno, Contract_date)
      ElseIf Me.Contract_period11.ListIndex = 1 Then
     DateEndContrac.value = DateAdd("yyyy", Contract_periodno, Contract_date)
     RtriverAsse val(Me.DCEmP.BoundText)
    End If


End Sub




Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    My_SQL = "TBLRegisterHoliday"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.DcBranches
    Dcombos.GetEmployees Me.DCEmP
    Dcombos.GetEmpJobsTypes Me.DcbJobs
    Dcombos.GETNationality Me.DcNational
    Dcombos.GetEmpDepartments Me.DcbDepartMen
    Dcombos.GetSection Me.DcbMangment
    Dcombos.GetEmployees Me.DcbDirctManger
    Dcombos.GetUsers Me.DCboUserName
 Dcombos.GetJobEndService dctype
 
   ' If SystemOptions.UserInterface = ArabicInterface Then
  '      My_SQL = "  select  id,name  from jopstatus   where id>1"
  '  Else
   '     My_SQL = "  select  id,namee  from jopstatus where id>1  "
   ' End If
   ' fill_combo dcjopstatus, My_SQL
If SystemOptions.UserInterface = ArabicInterface Then
dcjopstatus1.AddItem "«” ﬁ«·…"
dcjopstatus1.AddItem "⁄œ„ «·—€»… ›Ì «· ÃœÌœ"
dcjopstatus1.AddItem "«”»«» „—÷Ì…"
dcjopstatus1.AddItem "«Œ—Ï"
Contract_period11.AddItem "‘Â—"
Contract_period11.AddItem "”‰Â"
Else
dcjopstatus1.AddItem "Resigantion"
dcjopstatus1.AddItem "Non-Renewal of Contract"
dcjopstatus1.AddItem "Sick Leave"
dcjopstatus1.AddItem "Other"
Contract_period11.AddItem "Month"
Contract_period11.AddItem "Year"
End If

    

    Set cSearch = New clsDCboSearch
    Set cSearch.Client = Me.DCEmP

    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("EmpName"), Me.DCEmP

    FillGridWithData

    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("DiscountValue")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        
    End If

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    XPLbl(14).Caption = "Covenant"
lbl(4).Caption = "No.Req"
lbl(1).Caption = "Date"
lblBr.Caption = "Branch"
lbl(7).Caption = "Telephone"
lbl(2).Caption = "Nationality"
lbl(0).Caption = "Department"
lbl(8).Caption = "Job"
lbl(10).Caption = "Manage"
Label1(5).Caption = "Notice"
lbl(3).Caption = "Start Work"
lbl(9).Caption = "Manager"
lbl(29).Caption = "Salary"
lbl(6).Caption = "Contract period"
Label1(9).Caption = "Other"
Label1(7).Caption = "Suggestion"
Label1(8).Caption = "Remarks"
lbl(11).Caption = "By"
BtnPrint.Caption = "Print"
btnQuery.Caption = "Search"
lbl(5).Caption = "End contract"
    Me.Caption = "Register End Of Service"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(1).Caption = "Emp Name"
    Label1(4).Caption = " Status"
    Label1(0).Caption = "End Date"
    Label1(6).Caption = "Reason"

    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO. Recordes"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
        With Me.VSFlexGrid2
       .TextMatrix(0, .ColIndex("Serial")) = "Serial"
       .TextMatrix(0, .ColIndex("AsCode")) = "No"
       .TextMatrix(0, .ColIndex("mofrd")) = "Name"
         .TextMatrix(0, .ColIndex("ReciveDate")) = "ReciveDate"
       .TextMatrix(0, .ColIndex("DeliverDate")) = "DeliverDate"
       .TextMatrix(0, .ColIndex("Emp_NameTo")) = "Recipient Name"

    End With
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "Ser"
 
        .TextMatrix(0, .ColIndex("EmpName")) = "Name"
        .TextMatrix(0, .ColIndex("EndDate")) = "End Date"
        .TextMatrix(0, .ColIndex("Des")) = "Remarks"
 
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing

End Sub
Function CheckEndService() As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim Sql As String
Sql = "Select * from TBLRegisterHoliday   where FlagPayed=1 and id =" & val(XPTxtID.text) & " "
Rs8.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckEndService = True
Else
CheckEndService = False
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If

        RsSavRec.Close
        Set RsSavRec = Nothing
    End If

    Set cSearch = Nothing
ErrTrap:
End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TBLRegisterHoliday", "id", "")
    RsSavRec.AddNew
    RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Sub RtriverAsse(Optional empid As Integer = 0)
Dim Sql As String
Dim i As Integer
Dim RsDev As ADODB.Recordset
Sql = " SELECT     TOP 100 PERCENT dbo.TblAssestes.AsID, dbo.TblAssestes.AsName, dbo.TblAssestes.AsCode, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode,"
Sql = Sql & "                      TblEmployee_2.Emp_Namee, dbo.TblEmpAsest.ToEmId, TblEmployee_1.Emp_Name AS Emp_NameTo, TblEmployee_1.Fullcode AS FullcodeTo,"
Sql = Sql & "                       TblEmployee_1.Emp_Namee AS Emp_NameToE, dbo.TblEmpAsest.DeliverDate, dbo.TblEmpAsest.PostedDate, dbo.TblEmpAsestDetails.Qunt,"
Sql = Sql & "                       dbo.TblEmpAsestDetails.DIFF , dbo.TblEmpAsestDetails.FlagAs, dbo.TblEmpAsest.TypeAsset, dbo.TblEmpAsest.EmpAsestID"
Sql = Sql & "  FROM         dbo.TblEmpAsest LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblEmpAsestDetails ON dbo.TblEmpAsest.EmpAsID = dbo.TblEmpAsestDetails.IDAseset LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblAssestes ON dbo.TblEmpAsestDetails.AsID = dbo.TblAssestes.AsID LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblEmployee TblEmployee_1 ON dbo.TblEmpAsest.ToEmId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblEmployee TblEmployee_2 ON dbo.TblEmpAsest.EmpAsestID = TblEmployee_2.Emp_ID"
Sql = Sql & "  Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsest.EmpAsestID =" & empid & ")"
Set RsDev = New ADODB.Recordset
       RsDev.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
           VSFlexGrid2.Rows = 1
    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid2
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(RsDev("AsID").value), "", RsDev("AsID").value)
            
                .TextMatrix(i, .ColIndex("AsCode")) = IIf(IsNull(RsDev("AsCode").value), "", RsDev("AsCode").value)
                .TextMatrix(i, .ColIndex("DeliverDate")) = IIf(IsNull(RsDev("DeliverDate").value), "", RsDev("DeliverDate").value)
                .TextMatrix(i, .ColIndex("ReciveDate")) = IIf(IsNull(RsDev("PostedDate").value), "", RsDev("PostedDate").value)
            
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(RsDev("ToEmId").value), "", RsDev("ToEmId").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_NameTo").value), "", RsDev("Emp_NameTo").value)
                Else
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_NameToE").value), "", RsDev("Emp_NameToE").value)
                End If
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
End Sub
Public Sub FiLLRec()
    'On Error GoTo ErrTrap

    RsSavRec.Fields("Telephone").value = IIf((TxtTelephone.text) <> "", TxtTelephone.text, Null)
    RsSavRec.Fields("EmpID").value = IIf(val(Me.DCEmP.BoundText) <> 0, val(Me.DCEmP.BoundText), Null)
    RsSavRec.Fields("BranchID").value = IIf(val(Me.DcBranches.BoundText) <> 0, val(Me.DcBranches.BoundText), Null)
    RsSavRec.Fields("NationID").value = IIf(val(Me.DcNational.BoundText) <> 0, val(Me.DcNational.BoundText), Null)
    RsSavRec.Fields("JobID").value = IIf(val(Me.DcbJobs.BoundText) <> 0, val(Me.DcbJobs.BoundText), Null)
    RsSavRec.Fields("DirctMangerID").value = IIf(val(Me.DcbDirctManger.BoundText) <> 0, val(Me.DcbDirctManger.BoundText), Null)
    RsSavRec.Fields("DepartMentID").value = IIf(val(Me.DcbDepartMen.BoundText) <> 0, val(Me.DcbDepartMen.BoundText), Null)
    RsSavRec.Fields("MangmentID").value = IIf(val(Me.DcbMangment.BoundText) <> 0, val(Me.DcbMangment.BoundText), Null)
    RsSavRec.Fields("UserID").value = IIf(val(Me.DCboUserName.BoundText) <> 0, val(Me.DCboUserName.BoundText), Null)
    RsSavRec("EndWork").value = Me.DTP_Date.value
    RsSavRec.Fields("LogConract").value = IIf((TxtLogConract.text) <> "", TxtLogConract.text, Null)
    RsSavRec.Fields("Salary").value = IIf((lbl(23).Caption) <> "", val(lbl(23).Caption), Null)
    RsSavRec.Fields("Other").value = IIf((Me.TxtOther.text) <> "", TxtOther.text, Null)
    If Me.TxtModFlg.text = "N" Then
    RsSavRec.Fields("jopstatusid2").value = IIf((Me.Txtjopstatusid.text) <> "", Txtjopstatusid.text, Null)
    End If
        RsSavRec("Jopstatus1").value = val(Me.dcjopstatus1.ListIndex)
   
    If val(Me.dctype.BoundText) = 0 Then
        RsSavRec("jopstatusid").value = Null
    Else
        RsSavRec("jopstatusid").value = val(Me.dctype.BoundText)
    End If
RsSavRec("RecordDate").value = Me.XPDtbTrans.value
RsSavRec("DateSatrContrac").value = Me.DateSatrContrac.value
RsSavRec("DateEndContrac").value = Me.DateEndContrac.value

    RsSavRec("Notsstkala").value = IIf(Txt_NotEndWork.text = "", "", Trim(Txt_NotEndWork.text))
    RsSavRec("Remarkss").value = IIf(TxtRemarkss.text = "", "", Trim(TxtRemarkss.text))
    RsSavRec("Des").value = IIf(TxtDes.text = "", "", Trim(TxtDes.text))

    '

    RsSavRec.update
    CuurentLogdata
    Dim i As Integer
    If Me.TxtModFlg.text = "E" Then
    Cn.Execute "delete from  dbo.TBLRegisterHolidayDet Where RegID=" & val(TxtVac_ID.text) & ""
    End If
        ''/////////////
       Dim RsDev As ADODB.Recordset
    
             Set RsDev = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TBLRegisterHolidayDet Where (1 = -1)"
         RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      With Me.VSFlexGrid2

        For i = 1 To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("MofrdID"))) <> 0 Then
                RsDev.AddNew
                RsDev("RegID").value = val(Me.TxtVac_ID.text)
                RsDev("MofrdID").value = val(.TextMatrix(i, .ColIndex("MofrdID")))
                RsDev("EmpID").value = val(.TextMatrix(i, .ColIndex("EmpID")))
                RsDev("DeliverDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("DeliverDate"))), .TextMatrix(i, .ColIndex("DeliverDate")), Null)
                RsDev("ReciveDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("ReciveDate"))), .TextMatrix(i, .ColIndex("ReciveDate")), Null)
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
    
  '   StrSQL = "update TblEmployee Set   jopstatusid=" & val(dcjopstatus.BoundText) & " ,workstate=0  where Emp_ID=" & val(DcEmp.BoundText)
  '  Cn.Execute StrSQL
    
  '   StrSQL = "update TblEmployee Set ChekStkala=1 , EndWork='" & SQLDate(DTP_Date.value) & "',Notsstkala='" & Txt_NotEndWork.text & "' where Emp_ID=" & val(DcEmp.BoundText)
  '  Cn.Execute StrSQL
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
    MsgBox "Saved SuccesFully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    Dim i As Integer
   VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPTxtID.text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    Me.DCEmP.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    DTP_Date.value = IIf(IsNull(RsSavRec("EndWork").value), Date, RsSavRec("EndWork").value)
    Me.dctype.BoundText = IIf(IsNull(RsSavRec("jopstatusid").value), "", RsSavRec("jopstatusid").value)
    Txt_NotEndWork.text = IIf(IsNull(RsSavRec("Notsstkala").value), "", Trim(RsSavRec("Notsstkala").value))
''//
Me.DcBranches.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
Me.DcNational.BoundText = IIf(IsNull(RsSavRec.Fields("NationID").value), "", RsSavRec.Fields("NationID").value)
Me.DcbJobs.BoundText = IIf(IsNull(RsSavRec.Fields("JobID").value), "", RsSavRec.Fields("JobID").value)
Me.DcbDirctManger.BoundText = IIf(IsNull(RsSavRec.Fields("DirctMangerID").value), "", RsSavRec.Fields("DirctMangerID").value)
Me.DcbDepartMen.BoundText = IIf(IsNull(RsSavRec.Fields("DepartMentID").value), "", RsSavRec.Fields("DepartMentID").value)
Me.DcbMangment.BoundText = IIf(IsNull(RsSavRec.Fields("MangmentID").value), "", RsSavRec.Fields("MangmentID").value)
Me.TxtTelephone.text = IIf(IsNull(RsSavRec("Telephone").value), "", Trim(RsSavRec("Telephone").value))
Me.TxtLogConract.text = IIf(IsNull(RsSavRec("LogConract").value), "", Trim(RsSavRec("LogConract").value))
Me.TxtOther.text = IIf(IsNull(RsSavRec("Other").value), "", Trim(RsSavRec("Other").value))
Me.TxtRemarkss.text = IIf(IsNull(RsSavRec("Remarkss").value), "", Trim(RsSavRec("Remarkss").value))
Me.TxtDes.text = IIf(IsNull(RsSavRec("Des").value), "", Trim(RsSavRec("Des").value))
 Me.dcjopstatus1.ListIndex = IIf(IsNull(RsSavRec("Jopstatus1").value), -1, RsSavRec("Jopstatus1").value)
 DateSatrContrac.value = IIf(IsNull(RsSavRec("DateSatrContrac").value), Date, RsSavRec("DateSatrContrac").value)
 DateEndContrac.value = IIf(IsNull(RsSavRec("DateEndContrac").value), Date, RsSavRec("DateEndContrac").value)
 XPDtbTrans.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
 Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)
 lbl(23).Caption = IIf(IsNull(RsSavRec("Salary").value), 0, RsSavRec("Salary").value)
 Txtjopstatusid.text = IIf(IsNull(RsSavRec("jopstatusid2").value), 0, RsSavRec("jopstatusid2").value)
''/
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:
filgrid
End Sub

Public Sub EditRec(StrTable As String, _
                   RecID As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("EmpID")))
ErrTrap:
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption

End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
   Dim empid As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtCode.text, empid
        Me.DCEmP.BoundText = empid
    End If
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "id=" & RecID, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TBLRegisterHoliday order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
           
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("EmpName")) = get_EMPLOYEE_Data(val(.TextMatrix(i, .ColIndex("EmpID"))), "Emp_Name")
                Else
                    .TextMatrix(i, .ColIndex("EmpName")) = get_EMPLOYEE_Data(val(.TextMatrix(i, .ColIndex("EmpID"))), "Emp_NameE")
                End If
            
                .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(rs.Fields("EndWork").value), "", rs.Fields("EndWork").value)
            
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(rs.Fields("Notsstkala").value), "", rs.Fields("Notsstkala").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Sub filgrid()
Dim RsDev As ADODB.Recordset
Set RsDev = New ADODB.Recordset
Dim i As Integer
Dim StrSQL As String
StrSQL = " SELECT     dbo.TBLRegisterHolidayDet.RegID, dbo.TBLRegisterHolidayDet.DeliverDate, dbo.TBLRegisterHolidayDet.ID, dbo.TBLRegisterHolidayDet.ReciveDate,"
StrSQL = StrSQL & "                       dbo.TBLRegisterHolidayDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TBLRegisterHolidayDet.MofrdID,"
StrSQL = StrSQL & "                      dbo.TblAssestes.AsName , dbo.TblAssestes.AsCode, dbo.TblAssestes.AsestName"
StrSQL = StrSQL & " FROM         dbo.TBLRegisterHolidayDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAssestes ON dbo.TBLRegisterHolidayDet.MofrdID = dbo.TblAssestes.AsID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TBLRegisterHolidayDet.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " Where (dbo.TBLRegisterHolidayDet.RegID = " & val(TxtVac_ID.text) & ")"
RsDev.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDev.RecordCount > 0 Then
       With Me.VSFlexGrid2
        .Rows = .FixedRows + RsDev.RecordCount
        For i = .FixedRows To .Rows - 1
                 .TextMatrix(i, .ColIndex("Serial")) = i
                 .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(RsDev("MofrdID").value), "", RsDev("MofrdID").value)
                 .TextMatrix(i, .ColIndex("AsCode")) = IIf(IsNull(RsDev("AsCode").value), "", RsDev("AsCode").value)
                 .TextMatrix(i, .ColIndex("DeliverDate")) = IIf(IsNull(RsDev("DeliverDate").value), "", RsDev("DeliverDate").value)
                 .TextMatrix(i, .ColIndex("ReciveDate")) = IIf(IsNull(RsDev("ReciveDate").value), "", RsDev("ReciveDate").value)
                 .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(RsDev("EmpID").value), "", RsDev("EmpID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                Else
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                End If
            RsDev.MoveNext
         
        Next i
End With
End If
End Sub
'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If

    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If

    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If

    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If

    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If

    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If

    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If

    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If

    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If

    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If

    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If

    'End If

    Exit Sub
ErrTrap:
End Sub

Private Function CheckDelCountry(Lngid As Long) As Boolean
    'Dim Rs As ADODB.Recordset
    'Dim StrSQL As String
    'StrSQL = "Select * From TblEmployee Where GovernmentID=" & Lngid & ""
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '    CheckDelCountry = False
    'Else
    '    CheckDelCountry = True
    'End If
    'Rs.Close
    'Set Rs = Nothing
End Function



