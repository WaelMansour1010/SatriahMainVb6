VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmRegDevelopment 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   ‘«‘… „ «»⁄… «·„Â«„"
   ClientHeight    =   8445
   ClientLeft      =   3165
   ClientTop       =   2700
   ClientWidth     =   14175
   Icon            =   "FrmRegDevelopment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   14175
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ÊÊﬁ  «·„Â„Â"
      Height          =   2235
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   3120
      Width           =   6135
      Begin VB.TextBox TxtNoDayEnd 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   72
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtNoDaySatart 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   67
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker StartDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   59
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93913089
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker EndActDate 
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93913089
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker StartTime 
         Height          =   315
         Left            =   3240
         TabIndex        =   65
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93913090
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker EndActTIme 
         Height          =   315
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93913090
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   68
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   93913089
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   315
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   93913089
         CurrentDate     =   41640
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " √ŒÌ— «·»œ«Ì… »«·«Ì«„"
         Height          =   285
         Index           =   17
         Left            =   4560
         TabIndex        =   74
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " √ŒÌ— «·«‰ Â«¡ »«·«Ì«„"
         Height          =   285
         Index           =   15
         Left            =   1680
         TabIndex        =   73
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·»œ«Ì… "
         Height          =   285
         Index           =   12
         Left            =   4560
         TabIndex        =   71
         Top             =   255
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·«‰ Â«¡ "
         Height          =   285
         Index           =   5
         Left            =   1590
         TabIndex        =   70
         Top             =   255
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «·«‰ Â«¡ «·›⁄·Ì"
         Height          =   285
         Index           =   16
         Left            =   1560
         TabIndex        =   64
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «·»œ«Ì… «·›⁄·Ì"
         Height          =   285
         Index           =   14
         Left            =   4560
         TabIndex        =   63
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·«‰ Â«¡ «·›⁄·Ì"
         Height          =   285
         Index           =   13
         Left            =   1590
         TabIndex        =   62
         Top             =   735
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·»œ«Ì… «·›⁄·Ì"
         Height          =   285
         Index           =   9
         Left            =   4560
         TabIndex        =   60
         Top             =   735
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   6615
      Left            =   -360
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   480
      Width           =   14535
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«Ê·ÊÌ…"
         Enabled         =   0   'False
         Height          =   795
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1800
         Width           =   6135
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   375
            Index           =   0
            Left            =   3000
            TabIndex        =   56
            Top             =   240
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "⁄«œÌ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   57
            Top             =   240
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "„Â„"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   1275
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   600
         Width           =   6135
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3840
            TabIndex        =   79
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3840
            TabIndex        =   52
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox TxtCustomer 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            TabIndex        =   49
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSDataListLib.DataCombo DcbCustomer 
            Height          =   315
            Left            =   240
            TabIndex        =   50
            Top             =   1200
            Visible         =   0   'False
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbManager 
            Height          =   315
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Height          =   315
            Left            =   240
            TabIndex        =   80
            Top             =   720
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œÌ— «·„Â„Â "
            Height          =   285
            Index           =   3
            Left            =   4950
            TabIndex        =   81
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   285
            Index           =   10
            Left            =   4800
            TabIndex        =   54
            Top             =   1200
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”∆Ê· «·⁄„·Ì…"
            Height          =   285
            Index           =   0
            Left            =   4830
            TabIndex        =   51
            Top             =   720
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   5955
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   600
         Width           =   10905
         Begin VB.ComboBox DcbPand 
            Height          =   315
            ItemData        =   "FrmRegDevelopment.frx":038A
            Left            =   120
            List            =   "FrmRegDevelopment.frx":038C
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   3120
            Width           =   6255
         End
         Begin VB.ComboBox DcbProcess 
            Height          =   315
            ItemData        =   "FrmRegDevelopment.frx":038E
            Left            =   120
            List            =   "FrmRegDevelopment.frx":0390
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   2400
            Width           =   6255
         End
         Begin VB.TextBox TxtDesOp 
            Alignment       =   1  'Right Justify
            Height          =   1635
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   240
            Width           =   6255
         End
         Begin VB.TextBox TxtAnlysOp 
            Alignment       =   1  'Right Justify
            Height          =   2115
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   3600
            Width           =   6255
         End
         Begin MSDataListLib.DataCombo DcbTypeVisit1 
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   2040
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbDes 
            Height          =   315
            Left            =   120
            TabIndex        =   75
            Top             =   2760
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·„Â„Â"
            Height          =   285
            Index           =   21
            Left            =   6720
            TabIndex        =   83
            Top             =   2400
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·⁄„·Ì…"
            Height          =   285
            Index           =   20
            Left            =   6480
            TabIndex        =   82
            Top             =   3120
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„·Ì…"
            Height          =   285
            Index           =   19
            Left            =   6600
            TabIndex        =   76
            Top             =   2760
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„Â„Â"
            Height          =   285
            Index           =   2
            Left            =   6720
            TabIndex        =   48
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·„ÊŸ›"
            Height          =   645
            Index           =   29
            Left            =   6480
            TabIndex        =   45
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·«œ«—…"
            Height          =   1125
            Index           =   18
            Left            =   6480
            TabIndex        =   41
            Top             =   4080
            Width           =   1245
         End
      End
      Begin VB.ComboBox Contract_period 
         Height          =   315
         ItemData        =   "FrmRegDevelopment.frx":0392
         Left            =   18840
         List            =   "FrmRegDevelopment.frx":039C
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   11490
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmRegDevelopment.frx":03AA
         Height          =   315
         Left            =   600
         TabIndex        =   33
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
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
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   8880
         TabIndex        =   42
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   93913089
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker RecordTime 
         Height          =   315
         Left            =   6120
         TabIndex        =   84
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   93913090
         CurrentDate     =   41640
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "»Ì«‰«  «· ÿÊÌ—"
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
         Height          =   1500
         Index           =   23
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   4920
         Width           =   5895
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1575
         Left            =   480
         Top             =   4920
         Width           =   6015
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «·«œŒ«·"
         Height          =   285
         Index           =   22
         Left            =   7560
         TabIndex        =   85
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   -1320
         TabIndex        =   38
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·«œŒ«·"
         Height          =   285
         Index           =   1
         Left            =   10230
         TabIndex        =   36
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Õ—ﬂ…"
         Height          =   285
         Index           =   4
         Left            =   13230
         TabIndex        =   35
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblbr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13410
      TabIndex        =   27
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   14310
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   14190
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   19470
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   14190
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14205
      _cx             =   25056
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "          „ «»⁄… «·„Â«„ Ê«·⁄„·Ì«    "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1185
         TabIndex        =   1
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmRegDevelopment.frx":03BF
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
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmRegDevelopment.frx":0759
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
         Height          =   375
         Index           =   1
         Left            =   1710
         TabIndex        =   3
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmRegDevelopment.frx":0AF3
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
         Height          =   375
         Index           =   3
         Left            =   645
         TabIndex        =   4
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmRegDevelopment.frx":0E8D
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   3960
         Picture         =   "FrmRegDevelopment.frx":1227
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2400
         TabIndex        =   20
         Top             =   0
         Width           =   2205
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   2910
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7860
      Width           =   8745
      _cx             =   15425
      _cy             =   953
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   7230
         TabIndex        =   6
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   6375
         TabIndex        =   7
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   2
         Left            =   5535
         TabIndex        =   8
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   9
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   4
         Left            =   3825
         TabIndex        =   10
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   11
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   855
         TabIndex        =   12
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "„”«⁄œ…"
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   19
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   9
         Left            =   1920
         TabIndex        =   22
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â"
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   10020
      TabIndex        =   13
      Top             =   7440
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13470
      TabIndex        =   28
      Top             =   3570
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   13830
      TabIndex        =   29
      Top             =   1920
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·‘Œ’ «·„”∆Ê·"
      Height          =   285
      Index           =   28
      Left            =   4080
      TabIndex        =   44
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   13080
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–… «·‘«‘…  ﬁÊ„ » ”ÃÌ· ÿ·» ”›… ‰ﬁœÌ… ÊÌ „ «Õ ”«» ﬁÌ„… «·œ›⁄ «·Ì«"
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
      Height          =   660
      Index           =   25
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4170
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   4080
      Width           =   6015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   12645
      TabIndex        =   18
      Top             =   7515
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   4950
      TabIndex        =   17
      Top             =   7500
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   3210
      TabIndex        =   16
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2490
      TabIndex        =   15
      Top             =   7500
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   4260
      TabIndex        =   14
      Top             =   7500
      Width           =   615
   End
End
Attribute VB_Name = "FrmRegDevelopment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Public bol As Boolean
Public novalue As Boolean
'Private Sub Accredit_Click()
'    Dim BeginTrans As Boolean
'
'    Cn.BeginTrans
'    BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
'        rs("PostedDate") = Time
'    Else
'        rs("Posted") = Null
'       rs("PostedDate") = Time
'    End If
'
'    rs.update
' If SystemOptions.UserInterface = ArabicInterface Then
'    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If

'    Cn.CommitTrans
'    BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub

'Private Sub bClose_Click()
'Frame6.Visible = False
'If Me.ChekAccept.value = xtpChecked Then
'Frame2.Visible = True
'End If
'If Me.ChekContracted.value = xtpChecked Then
'Frame5.Visible = True
'End If
'End Sub

'Private Sub ChekAccept_Click()
'If Me.ChekAccept.value = vbChecked Then
'Me.CHekNotAccept.value = vbUnchecked
'Me.ChekContracted.value = vbUnchecked
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'Me.Frame2.Visible = True
'Me.Frame5.Visible = False
'Else
'Me.Frame2.Visible = False
'End If
'End Sub
'Private Sub RemoveGridRow()
'
'    With Me.Fg
'
'        If .Row <= 0 Then Exit Sub
'        .RemoveItem .Row
'    End With
'
'    ReLineGrid
'End Sub
'Private Sub RemoveGridRow2()
'
'    With Me.fg2
'
'        If .Row <= 0 Then Exit Sub
'        .RemoveItem .Row
'    End With
'
'    ReLineGrid
'End Sub

'Private Sub ChekContracted_Click()
'If Me.ChekContracted.value = xtpChecked Then
'Me.CHekNotAccept.value = xtpUnchecked
'Me.ChekAccept.value = xtpUnchecked
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'Me.Frame2.Visible = False
'Frame5.Visible = True
'Else
'Me.Frame5.Visible = False
'End If
'
'End Sub

'Private Sub CHekNotAccept_Click()
'If Me.CHekNotAccept.value = vbChecked Then
'Me.Frame2.Visible = False
'Me.Frame5.Visible = False
'lbl(36).Visible = True
'Me.txtnotAccept.Visible = True
'Me.ChekAccept.value = vbUnchecked
''Me.ChekContracted.value = vbUnchecked
'Else
'Me.Frame2.Visible = True
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'End If
'End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0
          

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            ' Me.DcbTO1.BoundText = 0
            '   Me.DcbTO2.BoundText = 0
            clear_all Me
          DcbProcess.ListIndex = 0
          DcbPand.ListIndex = 0
            Me.DCboUserName.BoundText = user_id
        '    TxtPaymentCounts.text = 1
Dcbranch.BoundText = Current_branch
RecordTime.value = Time
            'XPDtbTrans.SetFocus
            
           ' Accredit.Enabled = True
          '      If SystemOptions.UserInterface = ArabicInterface Then
          '                                          Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
          '                                        Else
          '                                          Accredit.Caption = " send to Approval   "
          '                                     End If
          '
          '
          ''
        Case 1
'Fg.Rows = Fg.Rows + 1
'Fg.Enabled = True
'fg2.Rows = fg2.Rows + 1
'fg2.Enabled = True
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.Dcbranch.BoundText

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
        
             Load FemSearchDevelopment
             FemSearchDevelopment.show vbModal

        Case 6
            Unload Me

        Case 7
           ' ShowGL_cc Me.txtNoteSerial.text, , 200

        Case 8
            
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.xptxtid.Text) <> 0 Then
                print_report val(Me.xptxtid.Text)
        
        
            End If
        
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
'  MySQL = "SELECT     dbo.TblRegDevelopment.Id, dbo.TblRegDevelopment.RecordDate, dbo.TblRegDevelopment.StrDate, dbo.TblRegDevelopment.EndExptedDate, "
'  MySQL = MySQL & "                    dbo.TblRegDevelopment.EndActDate, dbo.TblRegDevelopment.UserID, dbo.TblRegDevelopment.Important, dbo.TblRegDevelopment.MoDay,"
'  MySQL = MySQL & "                    dbo.TblRegDevelopment.CusID, dbo.TblRegDevelopment.DesOp, dbo.TblRegDevelopment.AnlysOp, dbo.TblRegDevelopment.TimeReq,"
'  MySQL = MySQL & "                    dbo.TblRegDevelopment.StartTime, dbo.TblRegDevelopment.EndExptedTime, dbo.TblRegDevelopment.EndActTIme, dbo.TblRegDevelopment.StatusProcess,"
'  MySQL = MySQL & "                    dbo.TblRegDevelopment.StatusPand, dbo.TblRegDevelopment.NoDaySatart, dbo.TblRegDevelopment.NoDayEnd, dbo.TblRegDevelopment.FromDate,"
'  MySQL = MySQL & "                    dbo.TblRegDevelopment.ToDate, dbo.TblRegDevelopment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
'  MySQL = MySQL & "                    dbo.TblRegDevelopment.EmpID, TblEmployee_1.Emp_Name, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.TblRegDevelopment.MangID,"
'  MySQL = MySQL & "                    TblEmployee_1.Emp_Name AS MangEmp_Name, TblEmployee_1.Fullcode AS MangFullcode, TblEmployee_1.Emp_Namee AS MangEmp_NameE,"
'  MySQL = MySQL & "                    dbo.TblRegDevelopment.OpType, dbo.TblProceeDevelper.Name, dbo.TblProceeDevelper.NameE, dbo.TblRegDevelopment.DesID, dbo.TblProceeDevelperDet.Des,"
'  MySQL = MySQL & "                    dbo.TblRegDevelopment.RecordTime , dbo.TblProceeDevelperDet.DesE"
'  MySQL = MySQL & "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblRegDevelopment LEFT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblProceeDevelperDet ON dbo.TblRegDevelopment.DesID = dbo.TblProceeDevelperDet.ID LEFT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblProceeDevelper ON dbo.TblRegDevelopment.OpType = dbo.TblProceeDevelper.ID LEFT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDevelopment.MangID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblEmployee TblEmployee_2 ON dbo.TblRegDevelopment.EmpID = TblEmployee_2.Emp_ID ON"
'  MySQL = MySQL & "                    dbo.TblBranchesData.branch_id = dbo.TblRegDevelopment.BranchId"
  MySQL = " SELECT     dbo.TblRegDevelopment.Id, dbo.TblRegDevelopment.RecordDate, dbo.TblRegDevelopment.StrDate, dbo.TblRegDevelopment.EndExptedDate,"
  MySQL = MySQL & "                               dbo.TblRegDevelopment.EndActDate, dbo.TblRegDevelopment.UserID, dbo.TblRegDevelopment.Important, dbo.TblRegDevelopment.MoDay,"
  MySQL = MySQL & "                               dbo.TblRegDevelopment.CusID, dbo.TblRegDevelopment.DesOp, dbo.TblRegDevelopment.AnlysOp, dbo.TblRegDevelopment.TimeReq,"
  MySQL = MySQL & "                               dbo.TblRegDevelopment.StartTime, dbo.TblRegDevelopment.EndExptedTime, dbo.TblRegDevelopment.EndActTIme, dbo.TblRegDevelopment.StatusProcess,"
  MySQL = MySQL & "                               dbo.TblRegDevelopment.StatusPand, dbo.TblRegDevelopment.NoDaySatart, dbo.TblRegDevelopment.NoDayEnd, dbo.TblRegDevelopment.FromDate,"
  MySQL = MySQL & "                               dbo.TblRegDevelopment.ToDate, dbo.TblRegDevelopment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
  MySQL = MySQL & "                               dbo.TblRegDevelopment.EmpID, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee, dbo.TblRegDevelopment.MangID,"
  MySQL = MySQL & "                               TblEmployee_1.Emp_Name AS MangEmp_Name, TblEmployee_1.Fullcode AS MangFullcode, TblEmployee_1.Emp_Namee AS MangEmp_NameE,"
  MySQL = MySQL & "                               dbo.TblRegDevelopment.OpType, dbo.TblProceeDevelper.Name, dbo.TblProceeDevelper.NameE, dbo.TblRegDevelopment.DesID, dbo.TblProceeDevelperDet.Des,"
  MySQL = MySQL & "                               dbo.TblRegDevelopment.RecordTime , dbo.TblProceeDevelperDet.DesE"
  MySQL = MySQL & "         FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
  MySQL = MySQL & "                               dbo.TblRegDevelopment LEFT OUTER JOIN"
   MySQL = MySQL & "                              dbo.TblProceeDevelperDet ON dbo.TblRegDevelopment.DesID = dbo.TblProceeDevelperDet.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                               dbo.TblProceeDevelper ON dbo.TblRegDevelopment.OpType = dbo.TblProceeDevelper.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                               dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDevelopment.MangID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                               dbo.TblEmployee TblEmployee_2 ON dbo.TblRegDevelopment.EmpID = TblEmployee_2.Emp_ID ON"
  MySQL = MySQL & "                               dbo.TblBranchesData.branch_id = dbo.TblRegDevelopment.BranchID"
'Where (dbo.TblRegDevelopment.ID = 30)

  
 MySQL = MySQL & " WHERE     (dbo.TblRegDevelopment.Id =" & val(Me.xptxtid.Text) & ") "
  
  
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDevelopment.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDevelopmentE.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
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
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(lbl(23).Caption), "0.00"), 0, True, ".")
'        xReport.ParameterFields(11).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(11).AddCurrentValue WriteNo(Format(val(lbl(23).Caption), "0.00"), 0, True, ".")
'  xReport.ParameterFields(12).AddCurrentValue WriteNo(Format(val(lbl(31).Caption), "0.00"), 0, True, ".")
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

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

'Private Sub Command2_Click()
'FillGridDetails
'Frame6.Visible = True
'Frame5.Visible = False
'Frame2.Visible = False
'End Sub
'Sub FillGridDetails()
'Dim StrSQL As String
'Dim i As Integer
'Dim RsDetails As ADODB.Recordset
'Set RsDetails = New ADODB.Recordset
'StrSQL = " SELECT     dbo.TblRegDateDelgate.Id, TblEmployee_1.Emp_ID, dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.DelgID, TblEmployee_1.Emp_Code AS Emp_CodeD, "
'StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name AS Emp_NameD, TblEmployee_1.Nationality AS NationalityD, TblEmployee_1.Fullcode AS FullcodeD,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.CustomerName, dbo.TblRegDateDelgate.Remark, dbo.TblRegDateDelgate.VisitID, TblTypeVisit_1.name, TblTypeVisit_1.namee,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.VisitID2, dbo.TblRegDateDelgate.SpAsID, dbo.TblSpeciaAsement.name AS nameSp, dbo.TblSpeciaAsement.namee AS nameeSp,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Accept, dbo.TblRegDateDelgate.VisitDate, dbo.TblRegDateDelgate.Remark2, dbo.TblRegDateDelgate.PersonConc,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile, dbo.TblRegDateDelgate.Email, dbo.TblRegDateDelgate.JobID, dbo.TblRegDateDelgate.LongTime,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.VisitDate1, TblTypeVisit_2.name AS name2, TblTypeVisit_2.namee AS namee2, dbo.TblRegDateDelgate.Entry, dbo.TblRegDateDelgate.Map,"
''StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Adress, dbo.TblRegDateDelgate.NotAcept, dbo.TblRegDateDelgate.BillNo, TblEmployee_1.Emp_Namee, dbo.TblRegDateDelgate.CustomerID,"
'StrSQL = StrSQL & "                         dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblRegDateDelgate.ToTime1, dbo.TblRegTimeDelgate.name AS ToTime11,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.FromTime1, TblRegTimeDelgate_2.name AS FromTime11, dbo.TblRegDateDelgate.FromTime2, TblRegTimeDelgate_3.name AS FromTime22,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.ToTime2, TblRegTimeDelgate_1.name AS ToTime22"
'StrSQL = StrSQL & "    FROM         dbo.TblRegTimeDelgate RIGHT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate TblRegTimeDelgate_1 RIGHT OUTER JOIN"
' StrSQL = StrSQL & "                        dbo.TblRegDateDelgate ON TblRegTimeDelgate_1.Id = dbo.TblRegDateDelgate.ToTime2 LEFT OUTER JOIN"
'' StrSQL = StrSQL & "                        dbo.TblRegTimeDelgate TblRegTimeDelgate_3 ON dbo.TblRegDateDelgate.FromTime2 = TblRegTimeDelgate_3.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate TblRegTimeDelgate_2 ON dbo.TblRegDateDelgate.FromTime1 = TblRegTimeDelgate_2.Id ON"
'StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate.Id = dbo.TblRegDateDelgate.ToTime1 LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblCustemers ON dbo.TblRegDateDelgate.CustomerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblTypeVisit TblTypeVisit_2 ON dbo.TblRegDateDelgate.VisitID2 = TblTypeVisit_2.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblTypeVisit TblTypeVisit_1 ON dbo.TblRegDateDelgate.VisitID = TblTypeVisit_1.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblSpeciaAsement ON dbo.TblRegDateDelgate.SpAsID = dbo.TblSpeciaAsement.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDateDelgate.DelgID = TblEmployee_1.Emp_ID"
'StrSQL = StrSQL & "    Where (dbo.TblRegDateDelgate.customerid =" & val(Me.DcbCustomer.BoundText) & ")"
'RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
'    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
'    VSFlexGrid1.Rows = VSFlexGrid1.FixedRows
'With VSFlexGrid1
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        .Rows = .FixedRows + RsDetails.RecordCount
'
'        For i = .FixedRows To .Rows - 1
'        .TextMatrix(i, .ColIndex("Serial")) = i
'        .TextMatrix(i, .ColIndex("PersonConc")) = IIf(IsNull(RsDetails("PersonConc").value), "", RsDetails("PersonConc").value) ' RsDetails("remark").value
'           ' .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CustomerName").value), "", RsDetails("CustomerName").value) 'RsDetails("fullcode").value
'            If SystemOptions.UserInterface = EnglishInterface Then
'           .TextMatrix(i, .ColIndex("Emp_NameD")) = IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value) 'RsDetails("Emp_Namee").value
'            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value)
'           Else
'           .TextMatrix(i, .ColIndex("Emp_NameD")) = IIf(IsNull(RsDetails("Emp_NameD").value), "", RsDetails("Emp_NameD").value) ' RsDetails("emp_name").value
'            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value)
'           End If
'            .TextMatrix(i, .ColIndex("Mobile")) = IIf(IsNull(RsDetails("Mobile").value), "", RsDetails("Mobile").value)
'             .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(RsDetails("JobID").value), "", RsDetails("JobID").value)
'              .TextMatrix(i, .ColIndex("Tel")) = IIf(IsNull(RsDetails("Tel").value), "", RsDetails("Tel").value)
'               .TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(RsDetails("Email").value), "", RsDetails("Email").value)
'                .TextMatrix(i, .ColIndex("FromTim11")) = IIf(IsNull(RsDetails("FromTime11").value), "", RsDetails("FromTime11").value)
 '                .TextMatrix(i, .ColIndex("ToTime11")) = IIf(IsNull(RsDetails("ToTime11").value), "", RsDetails("ToTime11").value)
'                  .TextMatrix(i, .ColIndex("Adress")) = IIf(IsNull(RsDetails("Adress").value), "", RsDetails("Adress").value)
'                  .TextMatrix(i, .ColIndex("VisitDate1")) = IIf(IsNull(RsDetails("VisitDate1").value), "", RsDetails("VisitDate1").value)
''                  DcbTypeVisit1.BoundText = val(IIf(IsNull(RsDetails("VisitID").value), "", RsDetails("VisitID").value))
 '                 .TextMatrix(i, .ColIndex("VisitID")) = DcbTypeVisit1.text
 '               If RsDetails("Accept").value = 0 Then
 '               .TextMatrix(i, .ColIndex("Accept")) = ""
 '               End If
 '                If RsDetails("Accept").value = 1 Then
 '               .TextMatrix(i, .ColIndex("Accept")) = " „ «·“Ì«—…"
 '               End If
 '                If RsDetails("Accept").value = 2 Then
 '               .TextMatrix(i, .ColIndex("Accept")) = " „ «· ⁄«ﬁœ"
 '               End If
 ''                If RsDetails("Accept").value = 3 Then
  '              .TextMatrix(i, .ColIndex("Accept")) = "≈·€«¡ «·“Ì«—…"
  '              End If
  '               .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(RsDetails("Remark").value), "", RsDetails("Remark").value)
                
                
  '          RsDetails.MoveNext
  '      Next i

  '  End If
'End With
'    RsDetails.Close
'    Set RsDetails = Nothing
'End Sub
'Private Sub DateVisit1_KeyUp(KeyCode As Integer, Shift As Integer)
'If TxtModFlg.text <> "R" Then
'If val(Me.DcboEmpName.BoundText) = 0 Then
'MsgBox "ÌÃ»  ÕœÌœ «”„ «·„‰œÊ» «Ê·«"
'Exit Sub
'Else
'fileFgtim val(Me.DcboEmpName.BoundText), 0
'refiltimdetails val(Me.DcboEmpName.BoundText), 0
'End If
'End If
'End Sub



'Private Sub DcbCustomer_Change()
'If Me.TxtModFlg.text <> "R" Then
''Me.TxtCustomer.text = ""
'retInfoCustomer

'End If
'End Sub

'Private Sub DcbFrom1_Change()
'If TxtModFlg.text <> "R" Then
'If val(Me.DcboEmpName.BoundText) = 0 Then
'MsgBox "«Œ Ì«— «·„ÊŸ› «Ê·«"
'Exit Sub
'End If
'If Me.DcbFrom1.text <> "" And Me.DcbTO1.text <> "" Then
'If val(Me.DcbFrom1.text) >= val(Me.DcbTO1.text) Then
'MsgBox "ÌÃ» «‰ ÌﬂÊ‰ «·Êﬁ  «·«ŒÌ— «ﬂ»— „‰ Êﬁ  «·»œ«ÌÂ"
''DcbTO2.SetFocus
'Exit Sub

'Else
'chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom1.BoundText), val(Me.DcbTO1.BoundText), 2
'fileFgtim val(Me.DcboEmpName.BoundText), 0
'refiltimdetails val(Me.DcboEmpName.BoundText), 0
'End If
'End If
'End If
'End Sub

'Private Sub DcboEmpName_Change()
'If TxtModFlg.text <> "R" Then
'fileFgtim val(Me.DcboEmpName.BoundText), 0
'refiltimdetails val(Me.DcboEmpName.BoundText), 0
'End If
'End Sub

'Private Sub DcboEmpName_Change()
'DcboEmpName_Click (0)

'End Sub






 




 

'Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
'                             Shift As Integer)
'
'    If KeyCode = vbKeyF3 Then
'        FrmEmployeeSearch.lbltype = 8
'       ' Set FrmEmployeeSearch.RetrunFrm = Me
'
'        FrmEmployeeSearch.Show
'
'    End If

Private Sub DcbCustomer_Change()
DcbCustomer_Click (0)
End Sub

Private Sub DcbCustomer_Click(Area As Integer)
  If val(DcbCustomer.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , DcbCustomer.BoundText, EmpCode
    Me.TxtCustomer.Text = EmpCode
End Sub

Private Sub DcbDes_Change()
DcbDes_Click (0)
End Sub

Private Sub DcbDes_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
If val(DcbTypeVisit1.BoundText) <> 0 Then
If val(Me.DcbDes.BoundText) <> 0 Then
RetriveInfoProcee val(DcbTypeVisit1.BoundText), val(Me.DcbDes.BoundText)
FromDate_Change
ToDate_Change
End If
End If
End If
End Sub

Private Sub DcbManager_Change()
DcbManager_Click (0)
End Sub

Private Sub DcbManager_Click(Area As Integer)
 If val(DcbManager.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcbManager.BoundText, EmpCode
    Me.Text1.Text = EmpCode
End Sub

Private Sub DcboEmpName_Change()
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    Me.Text2.Text = EmpCode
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
DcboEmpName_Change
End Sub

Private Sub DcbTypeVisit1_Change()

If val(DcbTypeVisit1.BoundText) <> 0 Then
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetDevelopProcessPand Me.DcbDes, val(DcbTypeVisit1.BoundText)
    End If

End Sub

Private Sub DcbTypeVisit1_Click(Area As Integer)
DcbTypeVisit1_Change
End Sub
Sub RetriveInfoProcee(Optional DevlOpID As Double = 0, Optional ID As Double = 0)
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "SELECT     dbo.TblProceeDevelper.ID, dbo.TblProceeDevelperDet.DevlOpID, dbo.TblProceeDevelperDet.EmpID, dbo.TblProceeDevelperDet.Des, "
sql = sql & "                      dbo.TblProceeDevelperDet.StartDate, dbo.TblProceeDevelperDet.NoDay, dbo.TblProceeDevelperDet.EndDate, dbo.TblProceeDevelperDet.Priority,"
sql = sql & "                       dbo.TblProceeDevelper.empID1 , dbo.TblProceeDevelper.remark, dbo.TblProceeDevelper.description , dbo.TblProceeDevelperDet.ID AS IDPand"
sql = sql & "  FROM         dbo.TblProceeDevelper LEFT OUTER JOIN"
sql = sql & "                       dbo.TblProceeDevelperDet ON dbo.TblProceeDevelper.ID = dbo.TblProceeDevelperDet.DevlOpID"
sql = sql & "  WHERE     (dbo.TblProceeDevelperDet.DevlOpID = " & DevlOpID & ") and (dbo.TblProceeDevelperDet.ID =" & ID & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
DcbManager.BoundText = IIf(IsNull(Rs8("empID1").value), 0, Rs8("empID1").value)
DcboEmpName.BoundText = IIf(IsNull(Rs8("EmpID").value), 0, Rs8("EmpID").value)
FromDate.value = IIf(IsNull(Rs8("StartDate").value), Date, Rs8("StartDate").value)
ToDate.value = IIf(IsNull(Rs8("EndDate").value), Date, Rs8("EndDate").value)
If Not (IsNull(Rs8("Priority").value)) Then
If Rs8("Priority").value = 2 Then
Opt(1).value = True
Else
Opt(0).value = True
End If
End If
End If
End Sub
'End Sub

Private Sub EndActDate_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoDayEnd.Text = DateDiff("d", ToDate.value, EndActDate.value)
End If
End Sub

Private Sub FromDate_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoDaySatart.Text = DateDiff("d", FromDate.value, StartDate.value)
End If
End Sub

'Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
'       If val(DcboEmpName.BoundText) = 0 Then Exit Sub


'    Dim EmpCode  As String
 
'    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
  '  TxtSearchCode.text = EmpCode
    
        'txtFile.text = EmpCode
        
'   If Me.TxtModFlg = "R" Then Exit Sub
'
'
'    Dim StrSQL As String
'
'
'        GetEmployeeSalaryAccordingToComponentAll val(Me.DcboEmpName.BoundText)
'
'        Dim IssueDate As Date
'        Dim depid As Double
'        Dim specid As Double
'        Dim JobTypeID As Double
'        Dim gradeID As Double
'        Dim Account_code2 As String
'           Dim Account_Code  As String
'        Dim Balance As String
'        Dim projectid As Integer
' Dim endiqama As String
'        Dim national As String
'        Dim endContractPerMonth As Double
'       Dim BignDateWork As Date
'       Dim JobTypeName As String
'       Dim JobTypeIDIQ As Integer
'       Dim iqama As String
'       Dim Contract_period As Integer
'     Dim Contract_periodno As Integer
'   Dim dcjopstatus As Integer
'Dim LastDate As Date
'        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, depid, specid, JobTypeID, gradeID, Account_code2, Account_Code, endContractPerMonth, national, , , projectid, , iqama, , , endiqama, , BignDateWork, LastDate, JobTypeName, Contract_period, Contract_periodno, , dcjopstatus, JobTypeIDIQ
        
'          WriteCustomerBalPublic Account_code2, Balance
          
'  lbl(22).Caption = val(Balance)
'Me.Contract_period.ListIndex = Contract_period
'Me.Txtlong.text = Contract_periodno & "     " & Me.Contract_period.text
'          WriteCustomerBalPublic Account_Code, Balance
      '  TxtNuWork.text = JobTypeName
'  lbl(21).Caption = val(Balance)
 ' lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
       ' DBIssueDate.value = issuedate
      '  DcboEmpDepartments.BoundText = depid
     ' DcProject.BoundText = projectid
      '  DcboSpecifications.BoundText = gradeID
'        DcboJobsType.BoundText = JobTypeIDIQ
'        lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "", 0)
'        lbl(31).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "", 1)
       ' Txtincrease.text = GetEmployeeSalaryAccordingToComponentName(val(Me.DcboEmpName.BoundText), "", 0)
    '  TxtOther.text = GetEmployeeSalaryAccordingToComponentName(val(Me.DcboEmpName.BoundText), "", 1)
    '    DcNational.text = national
  ' Me.DBEndDate.value = (endiqama)
'Me.dcjopstatus.BoundText = dcjopstatus
     '   Me.IssueDate.value = BignDateWork
       ' Me.TxtIqamaNo.text = iqama
 

'End Sub

' Sub GetEmployeeSalaryAccordingToComponentAll(Emp_id As Integer)
'
'  Dim sql As String
'    Dim mofrad_name As String
'    Dim valuee As Double
'    Dim rs As New ADODB.Recordset
'    Dim Balance As Double
'    Dim Mofradd As String
'    Dim i As Integer
'    Mofradd = ""
'
'    sql = "SELECT     dbo.EmpSalaryComponent.[Value],dbo.mofrdat.mofrad_name,dbo.mofrdat.mofrad_type "
''    sql = sql & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
 '   sql = sql & " dbo.mofrdat ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
 '   sql = sql & " WHERE   (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
 '
 '     rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'  With Me.Fg
'  .Rows = rs.RecordCount + 1
'      For i = 1 To rs.RecordCount
'       .TextMatrix(i, .ColIndex("Serial")) = i
'      .TextMatrix(i, .ColIndex("mofrdID")) = IIf(IsNull(rs("mofrad_type").value), 0, rs("mofrad_type").value)
'       .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
' .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(rs("value").value), 0, rs("value").value)
'
'
' rs.MoveNext
'      Next i
' End With
'     End If
     
     

'    rs.Close
    
'End Sub






'Private Sub DcbTO1_Change()
'If TxtModFlg.text <> "R" Then
'If val(Me.DcboEmpName.BoundText) = 0 Then
'MsgBox "«Œ Ì«— «·„ÊŸ› «Ê·«"
'Exit Sub
'End If
'If Me.DcbFrom1.text <> "" And Me.DcbTO1.text <> "" Then
'If val(Me.DcbFrom1.text) >= val(Me.DcbTO1.text) Then
'MsgBox "ÌÃ» «‰ ÌﬂÊ‰ «·Êﬁ  «·«ŒÌ— «ﬂ»— „‰ Êﬁ  «·»œ«ÌÂ"
''DcbTO2.SetFocus
'Exit Sub
'
'Else
'chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom1.BoundText), val(Me.DcbTO1.BoundText), 2
'fileFgtim val(Me.DcboEmpName.BoundText), 0
'refiltimdetails val(Me.DcboEmpName.BoundText), 0
'End If
'End If
'End If
'End Sub






















'Private Sub DcbTO2_Change()
'If TxtModFlg.text <> "R" Then
'If val(Me.DcboEmpName.BoundText) = 0 Then
'MsgBox "«Œ Ì«— «·„ÊŸ› «Ê·«"
''Exit Sub
'End If
'If Me.DcbFrom2.text <> "" And Me.DcbTO2.text <> "" Then
'If val(Me.DcbFrom2.text) >= val(Me.DcbTO2.text) Then
'MsgBox "ÌÃ» «‰ ÌﬂÊ‰ «·Êﬁ  «·«ŒÌ— «ﬂ»— „‰ Êﬁ  «·»œ«ÌÂ"
''DcbTO2.SetFocus
'Exit Sub
'
'Else
'chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom2.BoundText), val(Me.DcbTO2.BoundText), 1
'End If
'End If
'End If
'End Sub

'Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'Dim StrAccountCode As String
'Dim StrAccountCode1 As String
'    Dim Msg As String
'    Dim rs As New ADODB.Recordset
'    Dim StrSQL As String
'    Dim ClsAcc As New ClsAccounts
'    Dim LngRow As Long
'Dim StrComboList As String
'Dim bol As Boolean
'Dim Tye As Integer
'    With Fg
'
'
'
'        Select Case .ColKey(Col)
'
'            Case "empname"
'
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("code"), False, True)
'                ChekRepeat val(StrAccountCode), Row, bol
'                If bol = False Then
'               ' If StrAccountCode <> "" Then
'                .TextMatrix(Row, .ColIndex("empid")) = val(StrAccountCode)
'                StrSQL = " select Fullcode from  TblEmployee where Emp_ID=" & StrAccountCode & ""
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                If rs.RecordCount > 0 Then
'                .TextMatrix(Row, .ColIndex("code")) = rs("Fullcode").value
''                End If
 '               Else
 '               MsgBox "·«Ì„ﬂ‰ «Œ Ì«— «·„‰œÊ» Ê·«Ì„ﬂ‰ «· ﬂ—«—"
 ''               .TextMatrix(Row, .ColIndex("empname")) = ""
  '              .TextMatrix(Row, .ColIndex("code")) = ""
  ''              Exit Sub
   '             End If
   '           If TxtModFlg.text <> "R" Then
'fileFgtim val(StrAccountCode), Tye
'refiltimdetails val(StrAccountCode), Tye
'If Tye = 1 Then
'MsgBox .TextMatrix(Row, .ColIndex("empname")) & "·«Ì„ﬂ‰ «Œ Ì«—"
'.TextMatrix(Row, .ColIndex("empname")) = ""
'                .TextMatrix(Row, .ColIndex("code")) = ""
'              '  Exit Sub
'End If
'End If
'          Case "code"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("code"), False, True)
'               If StrAccountCode <> "" Then
'                .TextMatrix(Row, .ColIndex("empid")) = StrAccountCode
'                 StrSQL = " select * from  TblEmployee where Emp_ID=" & StrAccountCode & ""
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                If SystemOptions.UserInterface = ArabicInterface Then
'                .TextMatrix(Row, .ColIndex("empname")) = rs("Emp_Name").value
'                Else
'                .TextMatrix(Row, .ColIndex("empname")) = rs("Emp_Namee").value
'                End If
'                End If
'                   End Select
'
'        If Row = .Rows - 1 Then
'
'            .Rows = .Rows + 1
'        End If
'
        ' ReLineGrid
'    End With

'    ReLineGrid
'End Sub


     
    



Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub



 

 

'Private Sub menue_Click(Index As Integer)
'If Index = 2 Then
' Load FrmCustemers
'            FrmCustemers.Show
'            End If
'End Sub

 
'Private Sub XPDtbTrans_Change()
'If Me.TxtModFlg.text <> "R" Then
     
'         XPDtbTransH.value = ToHijriDate(XPDtbTrans.value)
       
'End If
'    If Trim(TxtNoteSerial1.text) <> "" Then
'        oldtxtNoteSerial1.text = TxtNoteSerial1.text
'    End If
'
'    TxtNoteSerial.text = ""
'    TxtNoteSerial1.text = ""

'End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.Text = ""
    txtNoteSerial1.Text = ""
End Sub
Sub ReloadDcb()
Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetDevelopProcess Me.DcbTypeVisit1
    Dcombos.GetDevelopProcessPand Me.DcbDes, val(DcbTypeVisit1.BoundText)
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.DcbManager
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetFileCustomer Me.DcbCustomer
End Sub
Private Sub Form_Load()
    
    Dim StrSQL As String
    Dim My_SQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
    If SystemOptions.UserInterface = ArabicInterface Then
    DcbProcess.AddItem " Õ  «· ‰›Ì–"
    DcbProcess.AddItem "„⁄·ﬁ"
    DcbProcess.AddItem " „ «·«‰ Â«¡"
    DcbPand.AddItem " Õ  «· ‰›Ì–"
    DcbPand.AddItem "„⁄·ﬁ"
    DcbPand.AddItem " „ «·«‰ Â«¡"
    Else
    DcbPand.AddItem "Under Execution"
    DcbPand.AddItem "Pending"
    DcbPand.AddItem "Completed"
    DcbProcess.AddItem "Under Execution"
    DcbProcess.AddItem "Pending"
    DcbProcess.AddItem "Completed"
    End If
ReloadDcb
If SystemOptions.Allowrank = True Then
TxtAnlysOp.locked = False
Else
TxtAnlysOp.locked = True
End If
'Frame6.Visible = False
    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    

    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If


    SetDtpickerDate Me.XPDtbTrans
  '  YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblRegDevelopment     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.Text = "R"
            If SystemOptions.UserInterface = EnglishInterface Then
           ' DcbMonDay.AddItem "Hour"
           ' DcbMonDay.AddItem "Day"
        SetInterface Me
        ChangeLang
        Else
       ' DcbMonDay.AddItem "”«⁄Â"
       ' DcbMonDay.AddItem "ÌÊ„"
    End If
    Me.Opt(0).value = False
     Me.Opt(1).value = False
    Retrive
   

 

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    Exit Sub

ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
  '  Label1.Visible = False

lbl(23).Caption = "Data of Development"

lbl(22).Caption = "Time"
Frame5.Caption = "Date & Time Tasks"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"


    Me.Caption = "Tasks Follow-Up "
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblBr.Caption = "Branch"
   lbl(3).Caption = "Employee"
   lbl(0).Caption = "Manger"
   'lbl(10).Caption = "Customer"
   'Frame4.Caption = "Data"
   lbl(12).Caption = "Start Date"
   lbl(5).Caption = "End Date"
   lbl(9).Caption = "From Actual Date"
   lbl(13).Caption = "To Actual Date"
   lbl(14).Caption = "From Actual Time"
   lbl(16).Caption = "To Actual Tim"
   lbl(18).Caption = "Management Comments"
   Frame2.Caption = "Priority"
   lbl(17).Caption = "Delay Start"
   lbl(15).Caption = "Delay End"
   Opt(0).RightToLeft = False
   lbl(29).Caption = "Employee Comments "
      Opt(1).RightToLeft = False
        Opt(0).Caption = "Normal"
     Opt(1).Caption = "Important"
      'Frame3.Caption = "Data of Process"
lbl(19).Caption = "Process"
 lbl(8).Caption = "By"
 lbl(2).Caption = "Task"
 lbl(21).Caption = "Task Status"
 lbl(20).Caption = "Process Status"
        lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
  
End Sub

' Private Sub YearMonth()

'    Dim i As Integer
'    Dim IntDefIndex As Integer

  '  CmbMonth.Clear

 '   For i = 1 To 12
    '    CmbMonth.AddItem MonthName(i)
   ' Next

   ' CmbMonth.ListIndex = Month(Date) - 1
   ' CboYear.Clear

  '  For i = 2010 To 2050
  '      CboYear.AddItem i
'
'        If i = year(Date) Then
'            IntDefIndex = CboYear.NewIndex
'        End If

'    Next

'    CboYear.ListIndex = IntDefIndex
'End Sub

Private Sub Form_Paint()
    TTD.Destroy
End Sub

Private Sub Form_Resize()
    TTD.Destroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub



Private Sub STARTDATE_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoDaySatart.Text = DateDiff("d", FromDate.value, StartDate.value)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text1.Text, EmpID
        DcbManager.BoundText = EmpID
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text2.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub

Private Sub ToDate_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoDayEnd.Text = DateDiff("d", ToDate.value, EndActDate.value)
End If
End Sub

Private Sub TxtCustomer_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode TxtCustomer.Text, EmpID
        DcbCustomer.BoundText = EmpID
    End If
End Sub

'Public Sub retInfoCustomer()
' Dim EmpID As Integer
'Dim name As String
'Dim mobile As String
'Dim phone As String
'Dim boxmail As String
'Dim fax As String
'Dim mail As String
'Dim adress As String
'Dim ZipCode As String
'Dim DigCus As String
'    Dim fullcode As String
'    Dim map As String
'Dim entry As String
'Dim ResponsibleContact As String
'    Dim jobname As String
'        GetCustomerIDFromCode Me.TxtCustomer.text, EmpID, , fullcode, Me.DcbCustomer.text, name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus, jobname, entry, map, ResponsibleContact
'       '  Me.TxtCustomer = fullcode
       '  Me.TxtPersonCont.text = ResponsibleContact
'        Me.DcbCustomer.BoundText = EmpID
'      ' Me.TxtMobi.text = mobile
'        Me.TxtTel.text = phone
'       Me.TxtMap.text = map
'        Me.TxtEnter.text = entry
'        Me.DcbJobID.text = jobname
'        Me.Txtemail.text = mail
'        Me.TxtAdres.text = adress
'        'Me.txtboxzip.text = ZipCode
'
'        'Me.TxtTypeCustomer.text = val(DigCus) + 1
       ' DcboEmpName.BoundText = EmpID
    
'End Sub

'Private Sub TxtCustomer_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then
''Me.DcbDelegate.BoundText = ""
'retInfoCustomer
'End If
'End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
        Frame1.Enabled = False
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  "
            'Me.menue(2).Enabled = True
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
          '  TxtAdvanceValue.Locked = True
            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False
            RecordTime.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
        Frame1.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            RecordTime.Enabled = True
            XPDtbTrans.value = Date
            RecordTime.value = Time

        Case "E"
        Frame1.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            RecordTime.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

 



Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim ContactTime As Date
    Dim i As Integer
    Dim StrSQL As String
    Me.Opt(0).value = False
     Me.Opt(1).value = False
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
  If Not IsNull(rs("StartTime").value) Then
      ContactTime = FormatDateTime(rs("StartTime").value, vbShortTime)
        Me.StartTime.value = ContactTime
   
    End If

  If Not IsNull(rs("RecordTime").value) Then
      ContactTime = FormatDateTime(rs("RecordTime").value, vbShortTime)
        Me.RecordTime.value = ContactTime
   
    End If
      If Not IsNull(rs("EndActTIme").value) Then
      ContactTime = FormatDateTime(rs("EndActTIme").value, vbShortTime)
        Me.EndActTIme.value = ContactTime
   
    End If
    xptxtid.Text = IIf(IsNull(rs("Id").value), "", val(rs("Id").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    StartDate.value = IIf(IsNull(rs("StrDate").value), Date, rs("StrDate").value)
' EndExptedDate.value = IIf(IsNull(rs("EndExptedDate").value), Date, rs("EndExptedDate").value)
 EndActDate.value = IIf(IsNull(rs("EndActDate").value), Date, rs("EndActDate").value)
 DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
 Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
 DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
 DcbManager.BoundText = IIf(IsNull(rs("MangID").value), "", rs("MangID").value)
 ' DcbMonDay.ListIndex = val(IIf(IsNull(rs("MoDay").value), -1, rs("MoDay").value))
    Me.DcbCustomer.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    DcbTypeVisit1.BoundText = val(IIf(IsNull(rs("OpType").value), 0, rs("OpType").value))
Me.TxtDesOp.Text = IIf(IsNull(rs("DesOp").value), "", rs("DesOp").value)
Me.TxtAnlysOp.Text = IIf(IsNull(rs("AnlysOp").value), "", rs("AnlysOp").value)

   ' Me.TxtTimeReq.text = IIf(IsNull(rs("TimeReq").value), "", rs("TimeReq").value)
    FromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
     ToDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
     Me.TxtNoDaySatart.Text = IIf(IsNull(rs("NoDaySatart").value), 0, rs("NoDaySatart").value)
     Me.TxtNoDayEnd.Text = IIf(IsNull(rs("NoDayEnd").value), 0, rs("NoDayEnd").value)
     Me.DcbDes.BoundText = IIf(IsNull(rs("DesID").value), 0, rs("DesID").value)
     Me.DcbPand.ListIndex = IIf(IsNull(rs("StatusPand").value), -1, rs("StatusPand").value)
     Me.DcbProcess.ListIndex = IIf(IsNull(rs("StatusProcess").value), -1, rs("StatusProcess").value)
     If Not (IsNull((rs("Important").value))) Then
 If val(rs("Important").value) = 0 Then
Me.Opt(0).value = True
ElseIf val(rs("Important").value) = 1 Then
Me.Opt(1).value = True
End If
End If
'       If IsNull(rs("posted").value) Then
'                                                   If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                  Else
'                                                    Accredit.Caption = " send to Approval   "
'                                               End If
'                                               Accredit.Enabled = True
'  Else
''                                                   If SystemOptions.UserInterface = ArabicInterface Then
 '                                                   Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
 '                                                 Else
 ''                                                   Accredit.Caption = " sent to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = False
  ' End If
  '
  '
  '  Set RsDetails = New ADODB.Recordset
 'StrSQL = " SELECT     dbo.TblRegDateDelgateDails.Id, dbo.TblRegDateDelgateDails.DelgID, dbo.TblRegDateDelgateDails.EmpID, dbo.TblEmployee.Emp_Code, "
'StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode, dbo.TblRegDateDelgateDails.remark,"
'StrSQL = StrSQL & "                      dbo.TblRegDateDelgateDails.Type"
'StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.Type = 0) And (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.text) & " )"
'StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.text) & " )"


' RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    Fg.Clear flexClearScrollable, flexClearEverything
'    Fg.Rows = Fg.FixedRows
'
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        Fg.Rows = Fg.FixedRows + RsDetails.RecordCount

'        For i = Me.Fg.FixedRows To Fg.Rows - 1
'        Fg.TextMatrix(i, Fg.ColIndex("Serial")) = i
'        Fg.TextMatrix(i, Fg.ColIndex("remarks")) = IIf(IsNull(RsDetails("remark").value), "", RsDetails("remark").value) ' RsDetails("remark").value
'            Fg.TextMatrix(i, Fg.ColIndex("code")) = IIf(IsNull(RsDetails("fullcode").value), "", RsDetails("fullcode").value) 'RsDetails("fullcode").value
'            If SystemOptions.UserInterface = EnglishInterface Then
'           Fg.TextMatrix(i, Fg.ColIndex("empname")) = IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value) 'RsDetails("Emp_Namee").value
'           Else
'           Fg.TextMatrix(i, Fg.ColIndex("empname")) = IIf(IsNull(RsDetails("emp_name").value), "", RsDetails("emp_name").value) ' RsDetails("emp_name").value
'           End If
'            Fg.TextMatrix(i, Fg.ColIndex("empid")) = RsDetails("EmpID").value
'            RsDetails.MoveNext
'        Next i
'
'    End If

'    RsDetails.Close
'    Set RsDetails = Nothing
   '''''''''''''///////////////////////
'   Set RsDetails1 = New ADODB.Recordset
' StrSQL = "SELECT     dbo.TblRegDateDelgateDails.Id, dbo.TblRegDateDelgateDails.DelgID, dbo.TblRegDateDelgateDails.EmpID, dbo.TblRegDateDelgateDails.remark, "
'StrSQL = StrSQL & "                      dbo.TblRegDateDelgateDails.Type , dbo.TblCompo.name, dbo.TblCompo.namee, dbo.TblRegDateDelgateDails.Quantity"
'StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
'  StrSQL = StrSQL & "                    dbo.TblCompo ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblCompo.Id"
'
'StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.Type = 1) And (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.text) & " )"



' RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
'    fg2.Clear flexClearScrollable, flexClearEverything
'    fg2.Rows = fg2.FixedRows
'
'    If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
'        RsDetails1.MoveFirst
'        fg2.Rows = fg2.FixedRows + RsDetails1.RecordCount
'
'        For i = Me.fg2.FixedRows To fg2.Rows - 1
'        fg2.TextMatrix(i, fg2.ColIndex("Serial")) = i
'        fg2.TextMatrix(i, fg2.ColIndex("remarks")) = IIf(IsNull(RsDetails1("remark").value), "", RsDetails1("remark").value) ' RsDetails1("remark").value
'            fg2.TextMatrix(i, fg2.ColIndex("code")) = IIf(IsNull(RsDetails1("quantity").value), "", RsDetails1("quantity").value) 'RsDetails1("fullcode").value
'            If SystemOptions.UserInterface = EnglishInterface Then
'           fg2.TextMatrix(i, fg2.ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value) 'RsDetails1("Emp_Namee").value
''           Else
 '          fg2.TextMatrix(i, fg2.ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value) ' RsDetails1("emp_name").value
 '          End If
 '           fg2.TextMatrix(i, fg2.ColIndex("empid")) = RsDetails1("EmpID").value
 '           RsDetails1.MoveNext
 '       Next i
'
'    End If

'    RsDetails1.Close
'    Set RsDetails1 = Nothing
   
   
   
   
   
 '  fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub


Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ „œÌ— «·⁄„·Ì…..!! "
            Else
            Msg = "Please Select Manager"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            DcboEmpName.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
   If Me.DcbTypeVisit1.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ   «·⁄„·Ì…..!! "
            Else
            Msg = "Please Select Process"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcbTypeVisit1.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
        

   
     Dim RsTest As New ADODB.Recordset

        Cn.BeginTrans
        BeginTrans = True
        
              If TxtModFlg.Text = "N" Then


        '”·› ”«»ﬁ…
   


            xptxtid.Text = CStr(new_id("TblRegDevelopment", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
       ' ElseIf Me.TxtModFlg.text = "E" Then
       '     StrSQL = "Delete From TblRegDateDelgateDails Where DelgID=" & val(Me.XPTxtID.text)
       '     Cn.Execute StrSQL, , adExecuteNoRecords

        End If
           rs("ID").value = val(xptxtid.Text)
    
         rs("RecordDate").value = XPDtbTrans.value
        rs("EndActDate").value = EndActDate.value
        rs("StrDate").value = StartDate.value
        
          rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
          rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
           rs("EmpID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
           rs("MangID").value = IIf(DcbManager.BoundText = "", Null, DcbManager.BoundText)
           rs("CusID").value = IIf(DcbCustomer.BoundText = "", Null, DcbCustomer.BoundText)
        rs("OpType").value = IIf(DcbTypeVisit1.BoundText = "", Null, DcbTypeVisit1.BoundText)
    rs("StartTime").value = FormatDateTime(Me.StartTime.value, vbShortTime)
  ' rs("EndExptedTime").value = FormatDateTime(Me.EndExptedTime.value, vbShortTime)
    rs("EndActTIme").value = FormatDateTime(Me.EndActTIme.value, vbShortTime)
    rs("DesOp").value = IIf(Me.TxtDesOp.Text = "", "", Me.TxtDesOp.Text)
    rs("AnlysOp").value = IIf(Me.TxtAnlysOp.Text = "", "", Me.TxtAnlysOp.Text)
    rs("RecordTime").value = FormatDateTime(Me.RecordTime.value, vbShortTime)
   '  rs("TimeReq").value = IIf(Me.TxtTimeReq.text = "", "", Me.TxtTimeReq.text)
   '  rs("MoDay").value = val(IIf(Me.DcbMonDay.ListIndex = -1, -1, Me.DcbMonDay.ListIndex))
     rs("FromDate").value = FromDate.value
     rs("ToDate").value = ToDate.value
     rs("NoDaySatart").value = IIf(Me.TxtNoDaySatart.Text = "", 0, val(Me.TxtNoDaySatart.Text))
     rs("NoDayEnd").value = IIf(Me.TxtNoDayEnd.Text = "", 0, val(Me.TxtNoDayEnd.Text))
     rs("DesID").value = IIf(Me.DcbDes.Text = "", Null, val(DcbDes.BoundText))
     rs("StatusPand").value = IIf(Me.DcbPand.ListIndex = -1, Null, val(DcbPand.ListIndex))
     rs("StatusProcess").value = IIf(Me.DcbProcess.ListIndex = -1, Null, val(DcbProcess.ListIndex))
     
     If Opt(0).value = True Then
     rs("Important").value = 0
End If
     If Opt(1).value = True Then
     rs("Important").value = 1
End If
        rs.update
   '        Set RsDetails = New ADODB.Recordset
   '    StrSQL = "SELECT     *  from dbo.TblRegDateDelgateDails Where (1 = -1)"
   'RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    
   '     For i = Me.Fg.FixedRows To Fg.Rows - 1
   '    If val(Fg.TextMatrix(i, Fg.ColIndex("EmpID"))) <> 0 Then
   '         RsDetails.AddNew
   '         RsDetails("DelgID").value = val(XPTxtID.text)
   '         RsDetails("Type").value = 0
   '        RsDetails("remark").value = Fg.TextMatrix(i, Fg.ColIndex("remarks"))
   '         RsDetails("EmpID").value = val(Fg.TextMatrix(i, Fg.ColIndex("empid")))
   '
   '         RsDetails.update
   '     End If
   '     Next i
  ''///////////'''''''''''''''''''''''''''''''
   '     Set RsDetails1 = New ADODB.Recordset
   ''    StrSQL = "SELECT     *  from dbo.TblRegDateDelgateDails Where (1 = -1)"
 '  RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    
 '       For i = Me.fg2.FixedRows To fg2.Rows - 1
 '      If val(fg2.TextMatrix(i, fg2.ColIndex("EmpID"))) <> 0 Then
 '           RsDetails1.AddNew
 '           RsDetails1("DelgID").value = val(XPTxtID.text)
 '           RsDetails1("Type").value = 1
 '          RsDetails1("remark").value = fg2.TextMatrix(i, fg2.ColIndex("remarks"))
 '           RsDetails1("EmpID").value = val(fg2.TextMatrix(i, fg2.ColIndex("empid")))
 '   RsDetails1("quantity").value = val(fg2.TextMatrix(i, fg2.ColIndex("code")))
 '           RsDetails1.update
 '       End If
 '       Next i
'        Dim NoteID As Long
'        Dim line_no As Integer
'        Dim RsNotes As New ADODB.Recordset
'        RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
'        If detect_employee_work_type = 1 Then
        
'            If Me.TxtModFlg.text = "E" Then
 
'                StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords

'            End If

'            RsNotes.AddNew
'            NoteID = CStr(TxtNoteID.text)
'            RsNotes("NoteID").value = CStr(TxtNoteID.text)
'            RsNotes("NoteType").value = 8032
'            RsNotes("NoteDate").value = XPDtbTrans.value
'            RsNotes("UserID").value = user_id
'            RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
'            RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”· «–‰ «·’—›
'            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
'            RsNotes("numbering_type1").value = sand_numbering_type(32) ' ”ÃÌ· «·”·›'‰Ê⁄  —ﬁÌ„    
'            RsNotes("sanad_year").value = year(XPDtbTrans.value)
'            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
'            RsNotes("note_value_by_characters").value = WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
            '     RsNotes("remark").value = txtRemarks.text & bankDes
'            RsNotes("Branch_no").value = val(Me.Dcbranch.BoundText)
                
'            RsNotes.update
                
'            line_no = 1
        
'            Msg = "”·› „ÊŸ›Ì‰ —ﬁ„ " & val(Me.XPTxtID.text)
'            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'
'            Employee_account = get_EMPLOYEE_Account(val(Me.DcboEmpName.BoundText), "Account_Code")
'            StrAccountCode = Employee_account
'
            '        StrAccountCode = "a1a3a4" 'Õ”«» “„„ «·„ÊŸ›Ì‰
'            If ModAccounts.AddNewDev(LngDevID, 1, StrAccountCode, val(Me.TxtAdvanceValue.text), 0, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If

'            StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

'            If ModAccounts.AddNewDev(LngDevID, 2, StrAccountCode, val(Me.TxtAdvanceValue.text), 1, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
        
'        End If
    
        Cn.CommitTrans
     '   BeginTrans = False
     '   RsDetails.Close
        
     '   Set RsDetails = Nothing
     '   RsDetails1.Close
     '   Set RsDetails1 = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
Else
Msg = "This Record Alredy Saved"
Msg = Msg & "You Need To Enter another Record "
End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                MsgBox " Saved SuccessFully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select

        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID='" & val(xptxtid.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                
                Retrive
                Exit Sub
            End If

            
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
Dim StrSQL1 As String
    On Error GoTo ErrTrap

    If xptxtid.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
     Else
     Msg = "Confirm Delete"
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblRegDevelopment Where ID=" & val(Me.xptxtid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
              '  StrSQL1 = "Delete From TblDefinDetails Where IDDef=" & val(Me.XPTxtID.text)
 'Cn.Execute StrSQL1, , adExecuteNoRecords
                If rs.RecordCount < 1 Then
                    clear_all Me
                    ' Fg.Clear flexClearScrollable, flexClearEverything
           ' Fg.Rows = 2
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "this operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



'Function FillApprovedTable()
' Dim RSApproval  As New ADODB.Recordset
'   Set RSApproval = New ADODB.Recordset
'   Dim currentdate As Date
'   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'
' Dim sql As String
'  Dim rs1 As New ADODB.Recordset
' Dim i As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
'  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
'  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
'  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
'  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
'sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs1.RecordCount > 0 Then
'            currentdate = Now
'            For i = 1 To rs1.RecordCount
'              RSApproval.AddNew
'                RSApproval("ScreenName").value = Me.name
'                RSApproval("levelo").value = IIf(IsNull(rs1("levelo").value), Null, rs1("levelo").value)
'               RSApproval("EmpID").value = IIf(IsNull(rs1("EmpID").value), Null, rs1("EmpID").value)
'                RSApproval("levelorder").value = IIf(IsNull(rs1("levelorder").value), Null, rs1("levelorder").value)
'                 RSApproval("currorder").value = IIf(IsNull(rs1("currorder").value), Null, rs1("currorder").value)
'                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
'                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
'                RSApproval("Transaction_Date").value = Date
                
'                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
'               RSApproval("SendTime").value = currentdate
'
'                 If i = 1 Then
'                        RSApproval("Currcursor").value = 1
'                         RSApproval("FromUser").value = user_name
'                End If
'
'                RSApproval.update
'                rs1.MoveNext
'            Next i
'
'    End If
    
    

'End Function



'Function fillapprovData()
'Dim Num As Integer
' Dim RsDetails As New ADODB.Recordset
' Dim StrSQL As String
'
'
' StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
'StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
'StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
'StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
'StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"
'
'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
' If Not (RsDetails.EOF Or RsDetails.BOF) Then
'        GRID2.Rows = RsDetails.RecordCount + 1
'
'
'        For Num = 1 To RsDetails.RecordCount
'
'       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
'    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
'   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
'   Else
'    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
'    End If
'
'        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
'           If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
'          Else
'             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
'          End If
'            If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            Else
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            End If
'            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
'          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
'
'
'RsDetails.MoveNext
'If Num = RsDetails.RecordCount Then
'
'        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
'                                 Else
'                                       Label11.Caption = "Approved"
'                                 End If
'                            Label11.BackColor = &H80FF80
'        Else
'                             If SystemOptions.UserInterface = ArabicInterface Then
'                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
'                            Else
'                                     Label11.Caption = "Currently required Approve"
'                            End If
'                 Label11.BackColor = &HFFFFC0
'        End If
'
'End If

'        Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close
'
'End Function
'Private Sub ChekRepeat(Optional ind As Integer, Optional Row As Long, Optional ByRef bo As Boolean)
'    Dim i As Integer
'
'
'    With fg2
' bo = False
'        For i = .FixedRows To .Rows - 1
'If i <> Row Then
'            If val(.TextMatrix(i, .ColIndex("empid"))) = val(ind) Then
'             bo = True
'   End If
'            End If
'            Next i
'            End With
'        With Fg
' bo = False
'        For i = .FixedRows To .Rows - 1
'If i <> Row Then
'            If val(.TextMatrix(i, .ColIndex("empid"))) = val(ind) Then
'             bo = True
'             End If
'             Else
             
'            If val(ind) = val(Me.DcboEmpName.BoundText) Then
'              bo = True
'              End If
'   End If
'
'            Next i
'            End With
'        End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
                    'SystemOptions.AllowIndirectCost
                    
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "    ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "     ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "     ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "     ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "     ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "     ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "     ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ‘«‘… „ «»⁄… «· ÿÊÌ—  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

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

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

 
'Private Sub XPDtbTransH_LostFocus()
'If Me.TxtModFlg.text <> "R" Then
'
'      XPDtbTrans.value = ToGregorianDate(XPDtbTransH.value)
'
'End If
'End Sub
