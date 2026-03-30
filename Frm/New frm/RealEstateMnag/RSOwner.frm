VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form RSOwner 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  »Ì«‰«  «·„·«ﬂ "
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   HelpContextID   =   60
   Icon            =   "RSOwner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   10530
   Begin VB.TextBox TxtVATNO 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6270
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   126
      Top             =   6180
      Width           =   2985
   End
   Begin VB.TextBox txtBankName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3390
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   121
      Top             =   4980
      Width           =   1665
   End
   Begin VB.TextBox txtBankAccount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   150
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   120
      Top             =   4980
      Width           =   1905
   End
   Begin VB.TextBox TxtBankIBAN 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   150
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   119
      Top             =   5340
      Width           =   1905
   End
   Begin VB.TextBox TxtBankCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3390
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   118
      Top             =   5340
      Width           =   1665
   End
   Begin VB.CheckBox chkSendMessage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«„ﬂ«‰Ì… «—”«· «·—”«∆·"
      Height          =   375
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   117
      Top             =   8010
      Width           =   1935
   End
   Begin VB.TextBox PasswordTxt 
      Alignment       =   1  'Right Justify
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   6240
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   116
      Top             =   4440
      Width           =   3045
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—’Ìœ «·«ÌÃ«—«  «·„” Õﬁ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1305
      Index           =   8
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   107
      Top             =   3600
      Width           =   2865
      Begin VB.TextBox TxtOpenBalance1 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   111
         Top             =   510
         Width           =   1365
      End
      Begin VB.OptionButton OptType1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "€Ì— „Õœœ"
         Height          =   255
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   210
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton OptType1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œ«∆‰"
         Height          =   255
         Index           =   1
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   109
         Top             =   210
         Width           =   765
      End
      Begin VB.OptionButton OptType1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÌ‰"
         Height          =   255
         Index           =   0
         Left            =   1950
         RightToLeft     =   -1  'True
         TabIndex        =   108
         Top             =   210
         Width           =   765
      End
      Begin MSComCtl2.DTPicker Dtp1 
         Height          =   330
         Left            =   120
         TabIndex        =   112
         Top             =   870
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   234160131
         CurrentDate     =   38718
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·—’Ìœ "
         Height          =   255
         Index           =   47
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   114
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   285
         Index           =   48
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   113
         Top             =   930
         Width           =   1215
      End
   End
   Begin VB.CheckBox locked 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ìﬁ«› «· ⁄«„·"
      Height          =   255
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—’Ìœ «·√›  «ÕÏ ··„«·ﬂ"
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
      Height          =   1335
      Index           =   0
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   88
      Top             =   3600
      Width           =   3075
      Begin VB.TextBox TxtOpenBalance 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÌ‰"
         Height          =   255
         Index           =   0
         Left            =   1950
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   210
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œ«∆‰"
         Height          =   255
         Index           =   1
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   210
         Width           =   915
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "€Ì— „Õœœ"
         Height          =   255
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   210
         Width           =   915
      End
      Begin MSComCtl2.DTPicker Dtp 
         Height          =   330
         Left            =   150
         TabIndex        =   93
         Top             =   900
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   234160131
         CurrentDate     =   38718
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·—’Ìœ "
         Height          =   345
         Index           =   5
         Left            =   1770
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   510
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   315
         Index           =   6
         Left            =   1770
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   900
         Width           =   1125
      End
   End
   Begin VB.CheckBox chkCustomerandVendor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄„Ì· Ê„Ê—œ"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   84
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8070
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   82
      Top             =   600
      Width           =   1125
   End
   Begin VB.TextBox txtopening_balance_voucher_id 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   -1560
      RightToLeft     =   -1  'True
      TabIndex        =   81
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox XPTxtCusNamee 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   120
      MaxLength       =   50
      TabIndex        =   77
      Top             =   960
      Width           =   3645
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   3840
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox XPMTxtRemarks2 
         Alignment       =   1  'Right Justify
         Height          =   795
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   74
         Top             =   480
         Width           =   5145
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”»» «·«Ìﬁ«›"
         Height          =   285
         Index           =   32
         Left            =   1950
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.TextBox txtcode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3405
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ’Ê„«  Œ«’… ··„Ê—œ ›Ï ›Ê« Ì— «·‘—«¡"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1005
      Index           =   6
      Left            =   10500
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   2490
      Width           =   8535
      Begin VB.TextBox TxtDiscountValuePur 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   600
         Width           =   1425
      End
      Begin VB.ComboBox CboDiscountTypePur 
         Height          =   315
         Left            =   3390
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   690
         Width           =   195
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·Œ’„"
         Height          =   285
         Index           =   28
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   660
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·Œ’„"
         Height          =   285
         Index           =   29
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ﬁ— «·„«·ﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2025
      Index           =   5
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   1500
      Width           =   4215
      Begin VB.TextBox TxtAddress 
         Alignment       =   1  'Right Justify
         Height          =   585
         Left            =   30
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   1140
         Width           =   2985
      End
      Begin MSDataListLib.DataCombo DcboCountryID 
         Height          =   315
         Left            =   450
         TabIndex        =   7
         Top             =   150
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboGovernmentID 
         Height          =   315
         Left            =   450
         TabIndex        =   52
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCityID 
         Height          =   315
         Left            =   450
         TabIndex        =   53
         Top             =   810
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·œÊ·…"
         Height          =   225
         Index           =   22
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   210
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Õ«›Ÿ…"
         Height          =   225
         Index           =   24
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   510
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œÌ‰…"
         Height          =   225
         Index           =   25
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   840
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄‰Ê«‰ »«· ›’Ì·"
         Height          =   585
         Index           =   26
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1140
         Width           =   765
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ’Ê„«  Œ«’… ··⁄„Ì· ›Ï ›Ê« Ì— «·»Ì⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Index           =   4
      Left            =   14340
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1170
      Width           =   5925
      Begin VB.ComboBox CboDiscountType 
         Height          =   315
         Left            =   3390
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox TxtDiscountValue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   630
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·Œ’„"
         Height          =   285
         Index           =   19
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·Œ’„"
         Height          =   285
         Index           =   20
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   690
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   3180
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   720
         Width           =   195
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  „Õ«”»Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2205
      Index           =   1
      Left            =   13050
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3540
      Width           =   5895
      Begin VB.TextBox TxtDepitInterval 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtCreditInterval 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox dcDepitIntervalID 
         Height          =   315
         ItemData        =   "RSOwner.frx":038A
         Left            =   120
         List            =   "RSOwner.frx":038C
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox dcCreditIntervalID 
         Height          =   315
         ItemData        =   "RSOwner.frx":038E
         Left            =   120
         List            =   "RSOwner.frx":0390
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtCreditlimitCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2910
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox TxtCreditLimit 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2910
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÂ «·«∆ „«‰"
         Height          =   285
         Index           =   30
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÂ «·«∆ „«‰"
         Height          =   285
         Index           =   31
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœ «·√∆ „«‰(œ«∆‰)"
         Height          =   285
         Index           =   11
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   570
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœ «·√∆ „«‰(„œÌ‰)"
         Height          =   285
         Index           =   10
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   210
         Width           =   1515
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·√ ’«·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2685
      Index           =   3
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   1530
      Width           =   4245
      Begin VB.TextBox txtauthorizationNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   2280
         Width           =   2925
      End
      Begin VB.TextBox TXTauthorizationname 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   1920
         Width           =   2925
      End
      Begin VB.TextBox TxtResponsibleContact 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   210
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   210
         Width           =   2805
      End
      Begin VB.TextBox TxtFaxNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   990
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1260
         Width           =   2025
      End
      Begin VB.TextBox TxtE_mail 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1590
         Width           =   2925
      End
      Begin VB.TextBox XPTxtMobile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   990
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   915
         Width           =   2025
      End
      Begin VB.TextBox XPTxtPhone 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   990
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   570
         Width           =   2025
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Êﬂ«·Â"
         Height          =   285
         Index           =   14
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊﬂÌ·"
         Height          =   285
         Index           =   13
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”∆Ê· «·≈ ’«·"
         Height          =   315
         Index           =   23
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·›«ﬂ”"
         Height          =   285
         Index           =   7
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1290
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·»—Ìœ «·≈·ﬂ —Ê‰Ï"
         Height          =   285
         Index           =   12
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1590
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« ›"
         Height          =   285
         Index           =   3
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÃÊ«·"
         Height          =   285
         Index           =   2
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   945
         Width           =   1155
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—’Ìœ «·„«·ﬂ «·Õ«·Ï"
      ForeColor       =   &H00000080&
      Height          =   1335
      Index           =   2
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   6540
      Width           =   2895
      Begin ImpulseButton.ISButton Cmd 
         Height          =   435
         Index           =   8
         Left            =   480
         TabIndex        =   32
         Top             =   510
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "⁄—÷  ﬁ—Ì— ﬂ‘› Õ”«»"
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
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton CmdPriceList 
         Height          =   255
         Left            =   600
         TabIndex        =   100
         Top             =   960
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         ButtonPositionImage=   1
         Caption         =   "ﬁ«∆„… ⁄ﬁ—«  «·„«·ﬂ"
         BackColor       =   14737632
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "RSOwner.frx":0392
         ColorButton     =   14737632
         ColorHighlight  =   16777215
         ColorHoverText  =   255
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledText=   255
         ColorToggledHoverText=   255
         ColorTextShadow =   -2147483637
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   255
         Index           =   8
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   315
         Index           =   9
         Left            =   1650
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   660
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   540
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox XPTxtComID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox XPMTxtRemark 
      Alignment       =   1  'Right Justify
      Height          =   555
      Left            =   6270
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   5010
      Width           =   3015
   End
   Begin VB.TextBox XPTxtComName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5250
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   990
      Width           =   3915
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   10515
      _cx             =   18547
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "  »Ì«‰«  «·„·«ﬂ "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
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
         Height          =   345
         Index           =   0
         Left            =   1185
         TabIndex        =   11
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
         ButtonImage     =   "RSOwner.frx":072C
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
         Left            =   120
         TabIndex        =   13
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
         ButtonImage     =   "RSOwner.frx":0AC6
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
         Left            =   1710
         TabIndex        =   10
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
         ButtonImage     =   "RSOwner.frx":0E60
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
         Left            =   645
         TabIndex        =   12
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
         ButtonImage     =   "RSOwner.frx":11FA
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
         Left            =   4200
         Picture         =   "RSOwner.frx":1594
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Image Img 
         Height          =   480
         Left            =   2280
         Picture         =   "RSOwner.frx":51FC
         Top             =   0
         Width           =   480
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   9165
      TabIndex        =   22
      Top             =   8400
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   8445
      TabIndex        =   23
      Top             =   8400
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   7725
      TabIndex        =   24
      Top             =   8400
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   6945
      TabIndex        =   25
      Top             =   8400
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
      Left            =   6105
      TabIndex        =   26
      Top             =   8400
      Width           =   795
      _ExtentX        =   1402
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
      Index           =   5
      Left            =   5190
      TabIndex        =   27
      Top             =   8400
      Width           =   885
      _ExtentX        =   1561
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
      Index           =   6
      Left            =   120
      TabIndex        =   28
      Top             =   8400
      Width           =   705
      _ExtentX        =   1244
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   4380
      TabIndex        =   29
      Top             =   8400
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
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
      Left            =   1830
      TabIndex        =   30
      Top             =   8400
      Width           =   825
      _ExtentX        =   1455
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
   Begin MSDataListLib.DataCombo DboParentAccount 
      Height          =   315
      Left            =   3960
      TabIndex        =   79
      Top             =   5760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   6720
      TabIndex        =   83
      Top             =   600
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   120
      TabIndex        =   85
      Top             =   600
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   12600
      TabIndex        =   97
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·”»»"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "RSOwner.frx":5EC6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DcCustomerType 
      Height          =   315
      Left            =   12600
      TabIndex        =   98
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   1680
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   9
      Left            =   2760
      TabIndex        =   105
      Top             =   8400
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â  ’œÌﬁ «·€—›…"
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
      Index           =   11
      Left            =   960
      TabIndex        =   106
      Top             =   8400
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«·„—›ﬁ« "
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
      Caption         =   "—ﬁ„ «· ”ÃÌ· VAT"
      Height          =   345
      Index           =   40
      Left            =   9300
      RightToLeft     =   -1  'True
      TabIndex        =   127
      Top             =   6180
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ"
      Height          =   315
      Index           =   17
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   125
      Top             =   5010
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ”«» «·»‰ﬂ"
      Height          =   285
      Index           =   16
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   124
      Top             =   5010
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·«Ì»«‰"
      Height          =   285
      Index           =   36
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   123
      Top             =   5370
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—„“ «·»‰ﬂ"
      Height          =   315
      Index           =   37
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   122
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·»«”Ê—œ"
      Height          =   285
      Index           =   15
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   115
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„Ê—œ"
      Height          =   285
      Index           =   2
      Left            =   10890
      RightToLeft     =   -1  'True
      TabIndex        =   99
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4185
      TabIndex        =   86
      Top             =   600
      Width           =   690
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ”«» «·—∆Ì”Ì"
      Height          =   315
      Index           =   33
      Left            =   9330
      RightToLeft     =   -1  'True
      TabIndex        =   80
      Top             =   5760
      Width           =   1155
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ «·«‰Ã·Ì“Ì"
      Height          =   255
      Index           =   4
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„«·ﬂ"
      Height          =   315
      Index           =   2
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   8010
      Width           =   615
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   1020
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   8010
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   0
      Left            =   2010
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8010
      Width           =   1635
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      Height          =   315
      Index           =   4
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   8010
      Width           =   255
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·„Ê—œ"
      Height          =   315
      Index           =   1
      Left            =   12450
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   660
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   165
      Index           =   1
      Left            =   9330
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   5130
      Width           =   975
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„«·ﬂ"
      Height          =   315
      Index           =   0
      Left            =   9210
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1005
      Width           =   1185
   End
End
Attribute VB_Name = "RSOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim ComReport As ClsCompanyReport
Dim Dcombo As ClsDataCombos
Dim cSearch(2) As clsDCboSearch
Dim FirstPeriodDateInthisYear  As Date
Public index As Integer

Private Sub ALLButton1_Click()
    Frame2.Visible = True
End Sub

Private Sub CboDiscountType_Change()
    Me.lbl(21).Visible = (Me.CboDiscountType.ListIndex = 2)

    If CboDiscountType.ListIndex = 0 Then
        lbl(20).Visible = False
        TxtDiscountValue.Visible = False
        lbl(21).Visible = False
    Else
        lbl(20).Visible = True
        TxtDiscountValue.Visible = True
        lbl(21).Visible = True
    End If

End Sub

Private Sub CboDiscountType_Click()
    CboDiscountType_Change
End Sub

Private Sub CboDiscountTypePur_Change()
    Me.lbl(27).Visible = (Me.CboDiscountTypePur.ListIndex = 2)

    If CboDiscountTypePur.ListIndex = 0 Then
        lbl(28).Visible = False
        TxtDiscountValuePur.Visible = False
        lbl(27).Visible = False
    Else
        lbl(28).Visible = True
        TxtDiscountValuePur.Visible = True
        lbl(27).Visible = True
    End If

End Sub

Private Sub CboDiscountTypePur_Click()
    CboDiscountTypePur_Change
End Sub

Function DeleteOpeningBalance()
    Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.text = 0
    Cmd_Click (2)

End Function

Private Sub Cmd_Click(index As Integer)
    Dim Msg As String
'    On Error GoTo ErrTrap

    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    Me.Dtp.value = FirstPeriodDateInthisYear
 
    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
 
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
            Me.dcBranch.BoundText = Current_branch
            '        XPTxtComID.text = CStr(new_id("TblCustemers", "CusID", "", True))
            '        txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
            '   XPTxtComName.SetFocus
            Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(47, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··„·«ﬂ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
            OptType(2).value = True
OptType1(2).value = True
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            '        If XPTxtComID.text = 1 Then
            '            Msg = "·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  Â–« «·”Ã·"
            '            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '            Exit Sub
            '        End If
            TxtModFlg.text = "E"

        Case 2
 
            Dim currentcode As String

            If txtID.text = "" Then
                currentcode = get_coding(branch_id, "TblCustemers", 16, Me.DCPreFix.text)

                If currentcode = "miniError" Then
                    MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Exit Sub
            
                ElseIf currentcode = "Manual" Then
                    MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                    Exit Sub
                Else
                    txtID = currentcode
                End If
            End If

            SaveData

        Case 3
            Call Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If XPTxtComID.text = 1 Then
                Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·”Ã·"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
            Dim s As String
            s = " SELECT ownerid from TblContract WHERE "
            s = s & " ownerid In (SELECT CusID From TblCustemers WHERE Fullcode = '" & txtID & "' and Type=57)"
            Dim rsCheck As ADODB.Recordset
            Set rsCheck = New ADODB.Recordset
            rsCheck.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Not rsCheck.EOF Then
                Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·”Ã· ·ÊÃÊœ ⁄ﬁÊœ „”Ã·… »«”„ «·„«·ﬂ"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            
            End If
            Del_Company

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
            
  FrmCompanySearch.lblSearchtype.Caption = 2122014
           FrmCompanySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            printingReport

        Case 6
            Unload Me

        Case 8
            '     If DoPremis(Do_Print, "ReportSuppliers", True) = False Then
            '         Exit Sub
            '     End If
            '     ShowCusBalance
            Dim FirstPeriod As Date
            getFirstPeriodDateInthisYear FirstPeriod
            ShowReport IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), XPTxtComName.text, FirstPeriod, Date

               Case 11
            On Error Resume Next
ShowAttachments DCPreFix.text & txtID.text, "270120151"
 
 
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdPriceList_Click()
If val(XPTxtComID.text) <> 0 Then
FrmAqarListOfOwner.loadgrid val(XPTxtComID.text)
Load FrmAqarListOfOwner
FrmAqarListOfOwner.show
End If
End Sub

Private Sub DboParentAccount_KeyUp(KeyCode As Integer, _
                                   Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 2122014
    End If

End Sub

Private Sub DcboCityID_Change()
    LoadDataCombos False, False, True
End Sub

Private Sub DcboCityID_Click(Area As Integer)
    DcboCityID_Change
End Sub

Private Sub DcboCityID_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        LoadDataCombos
    End If

End Sub

Private Sub DcboCountryID_Change()
    LoadDataCombos True, False, False
End Sub

Private Sub DcboCountryID_Click(Area As Integer)

    If val(Me.DcboCountryID.BoundText) <> 0 Then
        DcboCountryID_Change
    End If

End Sub

Private Sub DcboCountryID_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        LoadDataCombos
    End If

End Sub

Private Sub DcboGovernmentID_Change()
    LoadDataCombos False, True, False
End Sub

Private Sub DcboGovernmentID_Click(Area As Integer)
    DcboGovernmentID_Change
End Sub
 
Private Sub DcboGovernmentID_KeyUp(KeyCode As Integer, _
                                   Shift As Integer)

    If KeyCode = vbKeyF5 Then
        LoadDataCombos
    End If

End Sub

Private Sub Form_Activate()
    'XPTxtComID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
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

Private Sub Form_Load()
    Dim StrSQL As String

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    On Error GoTo ErrTrap
    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "»Ì«‰«  «·„·«ﬂ "
    LogTexte = " Open Window " & " Owners Data"
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    LoadDataCombos

    With CboDiscountType
        .Clear
        .AddItem "·«ÌÊÃœ Œ’„"
        .AddItem "Œ’„ »ﬁÌ„…"
        .AddItem "Œ’„ »‰”»…"
    End With

    With CboDiscountTypePur
        .Clear
        .AddItem "·«ÌÊÃœ Œ’„"
        .AddItem "Œ’„ »ﬁÌ„…"
        .AddItem "Œ’„ »‰”»…"
    End With

    With Me.dcDepitIntervalID
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÌÊ„"
            .AddItem "‘Â—"
            .AddItem "”‰…"
        Else
            .AddItem "day"
            .AddItem "month"
            .AddItem "year"
        End If

    End With

    With Me.dcCreditIntervalID
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÌÊ„"
            .AddItem "‘Â—"
            .AddItem "”‰…"
        Else
            .AddItem "day"
            .AddItem "month"
            .AddItem "year"
        End If

    End With

    Me.Dtp.value = Date
    'Resize_Form Me
    AddTip
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetAccountingCodes Me.DboParentAccount, False, True

    Dcombos.GetCodeing Me.DCPreFix, 5

    StrSQL = "select * From TblCustemers where Type=57"
            If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
        End If
        
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "   «·Œ—ÊÃ „‰ " & "»Ì«‰«  «·„·«ﬂ "
    LogTexte = " Exit Window " & " Owners Data"
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Set ComReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·„·«ﬂ " _
       & CHR(13) & " ﬂÊœ «·„Ê—œ  " & DCPreFix & txtID.text _
       & CHR(13) & "«·«”„ ⁄—»Ì  " & XPTxtComName _
       & CHR(13) & "   „”∆Ê· «·« ’«·   " & TxtResponsibleContact _
       & CHR(13) & " —ﬁ„ «·Â« ›     " & xptxtphone _
       & CHR(13) & " —ﬁ„ «·ÃÊ«·     " & XPTxtmobile _
       & CHR(13) & " —ﬁ„ «·›«ﬂ”     " & TxtFaxNumber _
       & CHR(13) & "  «·»—Ìœ «·«·ﬂ —Ê‰Ì       " & TxtE_mail _
       & CHR(13) & " «·œÊ·Â   " & DcboCountryID.text _
       & CHR(13) & " «·„Õ«›Ÿ…   " & DcboGovernmentID.text _
       & CHR(13) & "  «·„œÌ‰…  " & DcboCityID.text _
       & CHR(13) & "  «·⁄‰Ê«‰ »«· ›’Ì· " & TxtAddress _
       & CHR(13) & " „·«ÕŸ«   " & XPMTxtRemark.text _
       & CHR(13) & " ‰Ê⁄ «·Œ’„ ··„»Ì⁄«    " & CboDiscountType.text _
       & CHR(13) & "   ﬁÌ„Â «·Œ’„  " & TxtDiscountValue _
       & CHR(13) & " ‰Ê⁄ «·Œ’„ ··„‘ —Ì«    " & CboDiscountTypePur.text _
       & CHR(13) & "   ﬁÌ„Â «·Œ’„  " & TxtDiscountValuePur _
       & CHR(13) & "  ‰Ê⁄ «·„Ê—œ  " & DcCustomerType.text _
       & CHR(13) & " Õœ «·«∆ „«‰ „œÌ‰  " & TxtCreditLimit _
       & CHR(13) & " „œ… «·«∆ „«‰     " & TxtDepitInterval.text & "   " & dcDepitIntervalID.text _
       & CHR(13) & " Õœ «·«∆ „«‰ œ«∆‰   " & TxtCreditlimitCredit _
       & CHR(13) & " „œ… «·«∆ „«‰      " & TxtCreditInterval.text & "   " & dcCreditIntervalID.text _

       LogTextA = LogTextA & CHR(13) & " «·„·«ﬂ ø       "

    If chkCustomerandVendor.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„"
    Else
        LogTextA = LogTextA & "·«"
    End If

    LogTextA = LogTextA & CHR(13) & "«Ìﬁ«› «· ⁄«„·   ø     "

    If locked.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„"
        LogTextA = LogTextA & CHR(13) & "  ”»» «·«Ìﬁ«›   "
        LogTextA = LogTextA & CHR(13) & XPMTxtRemarks2
    Else
        LogTextA = LogTextA & "·«"
    End If

    LogTextA = LogTextA & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTextA = LogTextA & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTextA = LogTextA & "   œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTextA = LogTextA & "€Ì— „Õœœ"
    End If

    LogTextA = LogTextA & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ     " & TxtOpenBalance
    LogTextA = LogTextA & CHR(13) & "«·Õ”«» «·—∆Ì”Ì    " & DboParentAccount

    LogTexte = "  Õ›Ÿ ‘«‘… " & " Customers Data  " _
       & CHR(13) & "  Code  " & DCPreFix & txtID.text _
       & CHR(13) & "Name " & XPTxtCusNamee _
       & CHR(13) & " Contact Person" & TxtResponsibleContact _
       & CHR(13) & " Tel " & xptxtphone _
       & CHR(13) & "Mob " & XPTxtmobile _
       & CHR(13) & " Fax  " & TxtFaxNumber _
       & CHR(13) & "  Email   " & TxtE_mail _
       & CHR(13) & " Contry   " & DcboCountryID.text _
       & CHR(13) & " City   " & DcboGovernmentID.text _
       & CHR(13) & "  Town  " & DcboCityID.text _
       & CHR(13) & " Address " & TxtAddress _
       & CHR(13) & " Remarks  " & XPMTxtRemark _
       & CHR(13) & " Sales Discount  type  " & CboDiscountType.text _
       & CHR(13) & " Discount Value " & TxtDiscountValue _
       & CHR(13) & " Purchase Discount type " & CboDiscountTypePur.text _
       & CHR(13) & "  Discount Value" & TxtDiscountValuePur _
       & CHR(13) & "  Supplier . Type " & DcCustomerType.text _
       & CHR(13) & "The limit for debit  " & TxtCreditLimit _
       & CHR(13) & " Period     " & TxtDepitInterval.text & "   " & dcDepitIntervalID.text _
       & CHR(13) & "The limit for Credit   " & TxtCreditlimitCredit _
       & CHR(13) & " Period " & TxtCreditInterval.text & "   " & dcCreditIntervalID.text _

       LogTexte = LogTexte & CHR(13) & "Customer & Supplier ?  "

    If chkCustomerandVendor.value = vbChecked Then
        LogTexte = LogTexte & " Yes "
    Else
        LogTexte = LogTexte & " No "
    End If

    LogTexte = LogTexte & CHR(13) & "Locked"

    If locked.value = vbChecked Then
        LogTexte = LogTexte & "Yes "
        LogTexte = LogTexte & CHR(13) & "  Reasons  "
        LogTexte = LogTexte & CHR(13) & XPMTxtRemarks2
    Else
        LogTexte = LogTexte & "No "
    End If

    LogTexte = LogTexte & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTexte = LogTexte & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTexte = LogTexte & "œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTexte = LogTexte & "€Ì— „Õœœ"
    End If

    LogTexte = LogTexte & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
    LogTexte = LogTexte & CHR(13) & "  Parent Acc. " & DboParentAccount
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If

End Function

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Label2_Click()
    Frame2.Visible = False
End Sub

Private Sub OptType_Click(index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.text)
End Sub

Private Sub TxtCreditLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCreditLimit.text, 1)
End Sub

Private Sub TxtCreditlimitCredit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCreditlimitCredit.text, 1)
End Sub

Private Sub TxtModFlg_Change()

    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            DboParentAccount.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„·«ﬂ"
            Else
                Me.Caption = "Suppliers Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.XPTxtComID.locked = True
            Me.XPTxtComName.locked = True
            Me.xptxtphone.locked = True
            Me.XPTxtmobile.locked = True
            Me.XPMTxtRemark.locked = True

            If XPTxtComID.text <> "" Then
                If XPTxtComID.text = 1 Then
                    CmdPriceList.Enabled = False
                Else
                    CmdPriceList.Enabled = True
                End If
            End If

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
                CmdPriceList.Enabled = False
            End If

            Fra(0).Enabled = False

            '        Me.Dtp.Enabled = True
        Case "N"
            DboParentAccount.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„·«ﬂ( ÃœÌœ )"
            Else
                Me.Caption = "Suppliers Data(Enter New Record)."
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtComID.locked = True
            Me.XPTxtComName.locked = False
            Me.xptxtphone.locked = False
            Me.XPMTxtRemark.locked = False
            Me.XPTxtmobile.locked = False
            CmdPriceList.Enabled = False
            Fra(0).Enabled = True

            '        Me.Dtp.Enabled = True
        Case "E"
            DboParentAccount.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„·«ﬂ(  ⁄œÌ· )"
            Else
                Me.Caption = "Suppliers Data(Edit Record)."
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtComID.locked = True
            Me.XPTxtComName.locked = False
            Me.XPTxtmobile.locked = False
            Me.xptxtphone.locked = False
            Me.XPMTxtRemark.locked = False
            CmdPriceList.Enabled = False
            Fra(0).Enabled = True
            '        Me.Dtp.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)

    Dim SngCusBegainAccount As Single

    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "CusID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    Me.txtID.text = IIf(IsNull(rs("code").value), "", rs("code").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
    XPTxtComID.text = IIf(IsNull(rs("CusID").value), "", val(rs("CusID").value))
    Me.TxtCode = IIf(IsNull(rs("c1").value), "", rs("c1").value)
    XPTxtComName.text = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
    Me.TxtResponsibleContact.text = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
    xptxtphone.text = IIf(IsNull(rs("Cus_Phone").value), "", Trim(rs("Cus_Phone").value))
    
    txtauthorizationname.text = IIf(IsNull(rs("authorizationname").value), "", Trim(rs("authorizationname").value))
    txtauthorizationNo.text = IIf(IsNull(rs("authorizationNo").value), "", Trim(rs("authorizationNo").value))
   
    XPTxtmobile.text = IIf(IsNull(rs("Cus_mobile").value), "", Trim(rs("Cus_mobile").value))
    XPMTxtRemark.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    
    XPTxtCusNamee.text = IIf(IsNull(rs("CusNamee")), "", Trim(rs("CusNamee")))
    XPMTxtRemarks2.text = IIf(IsNull(rs("Remark2")), "", Trim(rs("Remark2")))
    locked.value = IIf(rs("locked") = True, 1, 0)
    If IsNull(rs("PassWord").value) Then
    chkSendMessage.value = vbUnchecked
    Else
    
    chkSendMessage.value = IIf(rs("SendMessage") = 1, 1, 0)
    
    End If
    
      
    Me.DboParentAccount.BoundText = IIf(IsNull(rs("parent_account")), "", rs("parent_account"))
    Me.DcCustomerType.BoundText = IIf(IsNull(rs("CustomerTypeID")), "", rs("CustomerTypeID"))
    PasswordTxt.text = IIf(IsNull(rs("PassWord").value), "", Trim(rs("PassWord").value))
    
    If rs("CustomerandVendor").value = True Then
        chkCustomerandVendor.value = vbChecked
    Else
        chkCustomerandVendor.value = vbUnchecked
    End If

    If XPTxtComID.text = 1 Then
        CmdPriceList.Enabled = False
    Else
        CmdPriceList.Enabled = True
    End If

    If Not (IsNull(rs("OpenBalanceDate").value)) Then
        Me.Dtp.value = rs("OpenBalanceDate").value
    Else
        Me.Dtp.value = Date
        Me.Dtp.Enabled = False
    End If

    If Not IsNull(rs("OpenBalanceType").value) Then
        Me.TxtOpenBalance.text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))
        If rs("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf rs("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
    Else
        Me.TxtOpenBalance.text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If

    If Not IsNull(rs("OpenBalanceType1").value) Then
        Me.TxtOpenBalance1.text = IIf(IsNull(rs("OpenBalance1")), "", Trim(rs("OpenBalance1")))
        If rs("OpenBalanceType1").value = 0 Then
            OptType1(0).value = True
            OptType1_Click 0
        ElseIf rs("OpenBalanceType1").value = 1 Then
            OptType1(1).value = True
            OptType1_Click 1
        End If
    Else
        Me.TxtOpenBalance1.text = 0
        Me.OptType1(2).value = True
        OptType1_Click 2
    End If
    
    TxtCreditLimit.text = IIf(IsNull(rs("CreditLimit").value), "0", rs("CreditLimit").value)
    Me.TxtCreditlimitCredit.text = IIf(IsNull(rs("CreditlimitCredit").value), "0", rs("CreditlimitCredit").value)
    Me.TxtFaxNumber.text = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
    Me.TxtE_mail.text = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
    'SngCusBegainAccount = GetCustomerAccount(val(XPTxtComID.text), True)

    Dim balanceString As String
    WriteCustomerBalPublic IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), , balanceString
    lbl(8).Caption = balanceString

    'If SngCusBegainAccount < 0 Then
    '    Me.lbl(8).Caption = Abs(SngCusBegainAccount)
    '    Me.lbl(9).Caption = "„œÌ‰"
    'ElseIf SngCusBegainAccount > 0 Then
    '    Me.lbl(8).Caption = Abs(SngCusBegainAccount)
    '    Me.lbl(9).Caption = "œ«∆‰"
    'Else
    '    Me.lbl(8).Caption = 0
    '    Me.lbl(9).Caption = ""
    'End If

    If IsNull(rs("Trans_DiscountType").value) Then
        Me.CboDiscountType.ListIndex = 0
        Me.TxtDiscountValue.text = 0
    ElseIf rs("Trans_DiscountType").value = 0 Then
        Me.CboDiscountType.ListIndex = 0
        Me.TxtDiscountValue.text = 0
    ElseIf rs("Trans_DiscountType").value = 1 Then
        Me.CboDiscountType.ListIndex = 1
        Me.TxtDiscountValue.text = IIf(IsNull(rs("Trans_Discount").value), "", rs("Trans_Discount").value)
    ElseIf rs("Trans_DiscountType").value = 2 Then
        Me.CboDiscountType.ListIndex = 2
        Me.TxtDiscountValue.text = IIf(IsNull(rs("Trans_Discount").value), "", rs("Trans_Discount").value)
    End If

    If IsNull(rs("Trans_DiscountTypePur").value) Then
        Me.CboDiscountTypePur.ListIndex = 0
        Me.TxtDiscountValuePur.text = 0
    ElseIf rs("Trans_DiscountTypePur").value = 0 Then
        Me.CboDiscountTypePur.ListIndex = 0
        Me.TxtDiscountValuePur.text = 0
    ElseIf rs("Trans_DiscountTypePur").value = 1 Then
        Me.CboDiscountTypePur.ListIndex = 1
        Me.TxtDiscountValuePur.text = IIf(IsNull(rs("Trans_DiscountPur").value), "", rs("Trans_DiscountPur").value)
    ElseIf rs("Trans_DiscountTypePur").value = 2 Then
        Me.CboDiscountTypePur.ListIndex = 2
        Me.TxtDiscountValuePur.text = IIf(IsNull(rs("Trans_DiscountPur").value), "", rs("Trans_DiscountPur").value)
    End If

    Me.DcboCountryID.BoundText = IIf(IsNull(rs("CountryID")), "", rs("CountryID"))
    Me.DcboGovernmentID.BoundText = IIf(IsNull(rs("GovernmentID")), "", rs("GovernmentID"))
    Me.DcboCityID.BoundText = IIf(IsNull(rs("CityID")), "", rs("CityID"))
    Me.TxtAddress.text = IIf(IsNull(rs("Address")), "", Trim(rs("Address")))
    TxtDepitInterval.text = IIf(IsNull(rs("DepitInterval")), 0, rs("DepitInterval"))
    TxtCreditInterval.text = IIf(IsNull(rs("CreditInterval")), 0, rs("CreditInterval"))
    
    dcDepitIntervalID.ListIndex = IIf(IsNull(rs("DepitIntervalID")), -1, rs("DepitIntervalID"))
    dcCreditIntervalID.ListIndex = IIf(IsNull(rs("CreditIntervalID")), -1, rs("CreditIntervalID"))

    txtBankAccount.text = IIf(IsNull(rs("BankAccount").value), "", rs("BankAccount").value)
    TXTBankName.text = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
    TxtVATNO.text = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
    TxtBankCode.text = IIf(IsNull(rs("BankCode").value), "", rs("BankCode").value)
    TxtBankIBAN.text = IIf(IsNull(rs("BankIBAN").value), "", rs("BankIBAN").value)
     
   



    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub ShowCusBalance()
    Dim cReport As ClsCustemerReport
    Dim LngCusID As Long
    LngCusID = val(XPTxtComID.text)
    ShowCusBalDailog LngCusID, 0
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.text, 0)
End Sub

Private Sub XPBtnMove_Click(index As Integer)
    On Error GoTo ErrTrap

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

Private Sub OptType1_Click(index As Integer)
    Me.TxtOpenBalance1.Enabled = Not OptType1(2).value
    Me.TxtOpenBalance1.text = IIf(OptType1(2).value = True, 0, Me.TxtOpenBalance1.text)

End Sub

Private Sub SaveData()

    On Error GoTo ErrTrap

    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim ParentAccount As String
    
    If Me.TxtModFlg.text <> "R" Then
        If XPTxtComName.text = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„«·ﬂ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtComName.SetFocus
            Exit Sub
        End If
    
        If Me.OptType(2).value = False Then
            If val(Me.TxtOpenBalance.text) = 0 Then
                Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ...!!!"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    If TxtOpenBalance.Enabled = True Then
                        TxtOpenBalance.SetFocus
                    End If
                    Exit Sub
            End If
        End If

        If Me.OptType1(2).value = False Then
            If val(Me.TxtOpenBalance1.text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  «œŒ· —’Ìœ «·«ÌÃ«—«  «·„” Õﬁ… ...!!!"
                Else
                    Msg = "Enter  Due  Balance for Checks"
                End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    If TxtOpenBalance1.Enabled = True Then
                        TxtOpenBalance1.SetFocus
                    End If
                    Exit Sub
                End If
            End If
        
            If Me.CboDiscountType.ListIndex = -1 Or Me.CboDiscountType.ListIndex = 0 Then
                Me.TxtDiscountValue.text = 0
            ElseIf Me.CboDiscountType.ListIndex = 1 Then
                If val(Me.TxtDiscountValue.text) = 0 Then
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·„«·ﬂ...!!!"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtDiscountValue.SetFocus
                    Exit Sub
                End If
            ElseIf Me.CboDiscountType.ListIndex = 2 Then
                If val(Me.TxtDiscountValue.text) = 0 Then
                    Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… »«·„«·ﬂ...!!!"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtDiscountValue.SetFocus
                    Exit Sub
                ElseIf val(Me.TxtDiscountValue.text) > 100 Then
                    Msg = "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ‰”»… «·Œ’„ «ﬂ»— „‰ 100 ...!!!"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtDiscountValue.SetFocus
                    Exit Sub
                End If
            End If
            
            If Me.CboDiscountTypePur.ListIndex = -1 Or Me.CboDiscountTypePur.ListIndex = 0 Then
                Me.TxtDiscountValuePur.text = 0
            ElseIf Me.CboDiscountTypePur.ListIndex = 1 Then
                If val(Me.TxtDiscountValuePur.text) = 0 Then
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·„«·ﬂ ›Ï ›Ê« Ì— «·‘—«¡...!!!"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtDiscountValuePur.SetFocus
                    Exit Sub
                End If
            ElseIf Me.CboDiscountTypePur.ListIndex = 2 Then
                If val(Me.TxtDiscountValuePur.text) = 0 Then
                    Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… »«·„«·ﬂ ›Ï ›Ê« Ì— «·‘—«¡..!!!"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtDiscountValuePur.SetFocus
                    Exit Sub
                ElseIf val(Me.TxtDiscountValuePur.text) > 100 Then
                    Msg = "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ‰”»… «·Œ’„ «ﬂ»— „‰ 100 ...!!!"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtDiscountValuePur.SetFocus
                    Exit Sub
                End If
            End If

            Select Case Me.TxtModFlg.text
                Case "N"
                    XPTxtComID.text = CStr(new_id("TblCustemers", "CusID", "", True))
                    StrSQL = "select * From  TblCustemers where Type=57 And CusName='" & Trim(XPTxtComName.text) & "'"
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If RsTemp.RecordCount > 0 Then
                        Msg = "ÌÊÃœ „«·ﬂ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ê—œ"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtComName.SetFocus
                        Exit Sub
                    End If
                    
                    If (TxtCode.text) = "" Then
                        GoTo xx
                    End If
                    
                    RsTemp.Close
                    StrSQL = "select * From  TblCustemers where Type=57 And c1=" & val(Me.TxtCode.text)
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If RsTemp.RecordCount > 0 Then
                        Msg = "ÌÊÃœ „«·ﬂ „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·ﬂÊœ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ—ﬂÊœ «·„Ê—œ"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtComName.SetFocus
                        Exit Sub
                    End If
xx:
                Case "E"
                    StrSQL = "select * From  TblCustemers where  Type=57 And CusName='" & Trim(XPTxtComName.text) & "'"
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If RsTemp.RecordCount > 0 Then
                        If RsTemp("CusID").value <> val(XPTxtComID.text) Then
                            Msg = "ÌÊÃœ „«·ﬂ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ê—œ"
                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            XPTxtComName.SetFocus
                            Exit Sub
                        End If
                    End If
                    
                    RsTemp.Close
                    
                    If (TxtCode.text) = "" Then
                        GoTo ll
                    End If
                    
                    StrSQL = "select * From  TblCustemers where  Type=57 And c1= " & val(TxtCode.text)
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    
                    If RsTemp.RecordCount > 0 Then
                        If RsTemp("CusID").value <> val(XPTxtComID.text) Then
                            Msg = "ÌÊÃœ „«·ﬂ „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·ﬂÊœ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“«·ﬂÊœ "
                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            XPTxtComName.SetFocus
                            Exit Sub
                        End If
                    End If
ll:
            End Select

            Cn.BeginTrans
            BeginTrans = True
            
            If Me.TxtModFlg.text = "N" Then
                Dim Account_Code_dynamic As String
                Account_Code_dynamic = Me.DboParentAccount.BoundText
                rs.AddNew
                rs("CusID").value = val(XPTxtComID.text)
            ElseIf Me.TxtModFlg.text = "E" Then
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
            
            rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        
            If val(TxtOpenBalance.text) = 0 Then
                txtopening_balance_voucher_id = 0
            End If
       
            If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
                If val(Me.txtopening_balance_voucher_id.text) = 0 Then
                    txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
                End If
            End If
         
            rs("code").value = txtID.text
            rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtID.text) = "", Null, txtID.text)
            rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)
            rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
            rs("c1").value = Me.TxtCode.text
            rs("CusName").value = Trim(XPTxtComName.text)
            rs("authorizationname").value = IIf(txtauthorizationname.text = "", "", Trim(txtauthorizationname.text))
            rs("authorizationno").value = IIf(txtauthorizationNo.text = "", "", Trim(txtauthorizationNo.text))
            rs("Cus_Phone").value = IIf(xptxtphone.text = "", "", Trim(xptxtphone.text))
            rs("Cus_mobile").value = IIf(XPTxtmobile.text = "", "", Trim(XPTxtmobile.text))
            rs("Remark").value = IIf(XPMTxtRemark.text = "", "", Trim(XPMTxtRemark.text))
            rs("Remark2").value = IIf(XPMTxtRemarks2.text = "", "", Trim(XPMTxtRemarks2.text))
            rs("parent_account").value = IIf(Me.DboParentAccount.BoundText = "", Null, Me.DboParentAccount.BoundText)
        
        
            rs("BankCode").value = Trim(TxtBankCode.text)
            rs("BankIBAN").value = Trim(TxtBankIBAN.text)
            rs("BankAccount").value = IIf(txtBankAccount.text = "", "", Trim(txtBankAccount.text))
            rs("BankName").value = IIf(TXTBankName.text = "", "", Trim(TXTBankName.text))
            rs("VATNO").value = TxtVATNO.text


            If locked.value = vbChecked Then
                rs("locked").value = 1
            Else
                rs("locked").value = 0
            End If
    
    If chkSendMessage.value = vbChecked Then
              rs("SendMessage").value = 1
            Else
                rs("SendMessage").value = 0
            End If
            
            If Trim(XPTxtCusNamee.text) = "" Then XPTxtCusNamee.text = Trim(XPTxtComName)
                rs("CusNamee").value = Trim(XPTxtCusNamee.text)
                If chkCustomerandVendor.value = vbChecked Then
                    rs("CustomerandVendor").value = 1
                Else
                    rs("CustomerandVendor").value = 0
                End If
                rs("Type").value = 57
                
                If Me.OptType(2).value = True Then
                    rs("OpenBalance").value = 0
                    rs("OpenBalanceType").value = Null
                ElseIf Me.OptType(0).value = True Then
                    rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
                    rs("OpenBalanceType").value = 0
                ElseIf Me.OptType(1).value = True Then
                    rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
                    rs("OpenBalanceType").value = 1
                End If

                If Me.OptType1(2).value = True Then
                    rs("OpenBalance1").value = 0
                    rs("OpenBalanceType1").value = Null
                ElseIf Me.OptType1(0).value = True Then
                    rs("OpenBalance1").value = val(Me.TxtOpenBalance1.text)
                    rs("OpenBalanceType1").value = 0
                ElseIf Me.OptType1(1).value = True Then
                    rs("OpenBalance1").value = val(Me.TxtOpenBalance1.text)
                    rs("OpenBalanceType1").value = 1
                End If
                
                rs("OpenBalanceDate").value = Me.Dtp.value
                rs("CreditLimit").value = val(Me.TxtCreditLimit.text)
                rs("CreditlimitCredit").value = val(Me.TxtCreditlimitCredit.text)
                rs("FaxNumber").value = IIf(Trim$(Me.TxtFaxNumber.text) = "", Null, Trim$(Me.TxtFaxNumber.text))
                rs("E_mail").value = IIf(Trim$(Me.TxtE_mail.text) = "", Null, Trim$(Me.TxtE_mail.text))

                If Me.CboDiscountType.ListIndex = -1 Or Me.CboDiscountType.ListIndex = 0 Then
                    rs("Trans_DiscountType").value = 0
                    rs("Trans_Discount").value = 0
                ElseIf Me.CboDiscountType.ListIndex = 1 Then
                    rs("Trans_DiscountType").value = 1
                    rs("Trans_Discount").value = val(Me.TxtDiscountValue.text)
                ElseIf Me.CboDiscountType.ListIndex = 2 Then
                    rs("Trans_DiscountType").value = 2
                    rs("Trans_Discount").value = val(Me.TxtDiscountValue.text)
                End If
                
                If Me.CboDiscountTypePur.ListIndex = -1 Or Me.CboDiscountTypePur.ListIndex = 0 Then
                    rs("Trans_DiscountTypePur").value = 0
                    rs("Trans_DiscountPur").value = 0
                ElseIf Me.CboDiscountTypePur.ListIndex = 1 Then
                    rs("Trans_DiscountTypePur").value = 1
                    rs("Trans_DiscountPur").value = val(Me.TxtDiscountValuePur.text)
                ElseIf Me.CboDiscountTypePur.ListIndex = 2 Then
                    rs("Trans_DiscountTypePur").value = 2
                    rs("Trans_DiscountPur").value = val(Me.TxtDiscountValuePur.text)
                End If
            
            
                 
                If Me.TxtModFlg.text = "N" Then
                
                
                 
                    If SystemOptions.Create2account4Supp = False Then
                        rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtComName.text), True, False, XPTxtCusNamee.text, , , , , , , , , , 1, 1, 1, 0, 0)
                    Else
                        ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtComName.text, False, False, XPTxtCusNamee.text)
                        rs("ParentAccount").value = ParentAccount
                        rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtComName.text), True, False, XPTxtCusNamee.text)
                        rs("Account_Code1").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtComName.text) & "   «ÌÃ«—«   „” Õﬁ…   ", True, False, XPTxtCusNamee.text & "  Dur Rent ")
                    End If
                Else
                    If SystemOptions.Create2account4Supp = False Then
                        If Not IsNull(rs("Account_Code").value) And Not (rs("Account_Code").value) = "" Then
                            ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtComName.text, Me.XPTxtCusNamee.text, , , , , , , , , 1, 1, 1, 0, 0, , , , True
                        Else
                            If Account_Code_dynamic = "" Then Account_Code_dynamic = Me.DboParentAccount.BoundText
                            rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtComName.text), True, False, XPTxtCusNamee.text, , , , , , , , , , 1, 1, 1, 0, 0)
                        End If
                    Else
                        If Not IsNull(rs("ParentAccount").value) And Not (rs("ParentAccount").value) = "" Then
                            ModAccounts.EditAccount rs("ParentAccount").value, Me.XPTxtComName.text, Trim(XPTxtCusNamee.text), , , , , , , , , , , , , , , , , False
                        Else
                            ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtComName.text, False, False, XPTxtCusNamee.text)
                            rs("ParentAccount").value = ParentAccount
                        End If
                        
                        If Not IsNull(rs("Account_Code").value) And Not (rs("Account_Code").value) = "" Then
                            ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtComName.text, Me.XPTxtCusNamee.text, , , , , , , , , 1, 1, 1, 0, 0, , , , True
                        Else
                            rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtComName.text), True, False, XPTxtCusNamee.text, , , , , , , , , , 1, 1, 1, 0, 0)
                        End If
                            
                        If Not IsNull(rs("Account_Code1").value) And Not (rs("Account_Code1").value) = "" Then
                            ModAccounts.EditAccount rs("Account_Code1").value, Me.XPTxtComName.text & " «ÌÃ«—«  „” Õﬁ…", Me.XPTxtCusNamee.text & "Due rent", , , , , , , , , 1, 1, 1, 0, 0, , , , True
                        Else
                            rs("Account_Code1").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtComName.text) & " «ÌÃ«—«  „” Õﬁ…", True, False, XPTxtCusNamee.text & "Due rent", , , , , , , , , , 1, 1, 1, 0, 0)
                        End If
                    End If
                End If
            End If
            
               Account_Code_dynamic = Trim(rs("ParentAccount").value & "")
                    
                 If Me.TxtModFlg.text = "N" Then
                    rs("Account_VAT").value = ModAccounts.AddNewAccount(Account_Code_dynamic, "«·ﬁÌ„… «·„÷«›… " & Trim$(Me.XPTxtComName.text), True, False, XPTxtComName.text, , , , , , , , , , 1, 1, 1, 0, 0)
                    
                 Else
                        If Trim(rs("Account_VAT").value & "" = "") Then
                            rs("Account_VAT").value = ModAccounts.AddNewAccount(Account_Code_dynamic, "«·ﬁÌ„… «·„÷«›… " & Trim$(Me.XPTxtComName.text), True, False, XPTxtComName.text, , , , , , , , , , 1, 1, 1, 0, 0)
                        End If
                        ModAccounts.EditAccount Trim(rs("Account_VAT").value & ""), Me.XPTxtComName.text & " «·ﬁÌ„… «·„÷«›…", Me.XPTxtCusNamee.text & "Due rent", , , , , , , , , 1, 1, 1, 0, 0, , , , True
                      '  ModAccounts.EditAccount Trim(Account_Code_dynamic), Me.XPTxtComName.Text & " «·ﬁÌ„… «·„÷«›…", Me.XPTxtCusNamee.Text & "Due rent", , , , , , , , , 1, 1, 1, 0, 0, , , , True
                 End If
            
            rs("CountryID").value = IIf(val(Me.DcboCountryID.BoundText) = 0, Null, val(Me.DcboCountryID.BoundText))
            rs("GovernmentID").value = IIf(val(Me.DcboGovernmentID.BoundText) = 0, Null, val(Me.DcboGovernmentID.BoundText))
            rs("CityID").value = IIf(val(Me.DcboCityID.BoundText) = 0, Null, val(Me.DcboCityID.BoundText))
            rs("ResponsibleContact").value = Trim$(Me.TxtResponsibleContact.text)
            rs("Address").value = Trim$(Me.TxtAddress.text)
            rs("CustomerTypeID").value = IIf(val(Me.DcCustomerType.BoundText) = 0, Null, val(Me.DcCustomerType.BoundText))
            rs("DepitInterval").value = val(TxtDepitInterval.text)
            rs("CreditInterval").value = val(TxtCreditInterval.text)
            rs("DepitIntervalID").value = val(dcDepitIntervalID.ListIndex)
            rs("CreditIntervalID").value = val(dcCreditIntervalID.ListIndex)
            rs("PassWord").value = Trim$(Me.PasswordTxt.text)
        rs.update
    
        Dim StrDes As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtComName.text) & " "
        Else
            StrDes = " Opening Balance For: " & Trim(Me.XPTxtCusNamee.text) & " "
        End If
        
        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        
                Dim LngDevID As Long
                Dim LngOpenID As Long
                Dim Account_Code_dynamic1 As String
        
                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
                If Me.OptType(0).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(60, my_branch)
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else
                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                ElseIf Me.OptType(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(60, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else
                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
            End If
        End If

'2222222222222222222222222222222222222222222222222222222
GoTo endNew
        If SystemOptions.Create2account4Supp = True Then
            If Me.OptType1(0).value = True Or Me.OptType1(1).value = True Then
                If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
                    LngOpenID = 1
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
                    If Me.OptType1(0).value = True Then
                        Account_Code_dynamic1 = get_account_code_branch(124, my_branch)
                        If Account_Code_dynamic1 = "NO branch" Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                            GoTo ErrTrap
                        Else
                            If Account_Code_dynamic1 = "NO account" Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ   ··«ÌÃ«—«  «·„” Õﬁ… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                GoTo ErrTrap
                            End If
                        End If
                        
                        If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code1").value, Round(Me.TxtOpenBalance1.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtComName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
                        
                        If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance1.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtComName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
                
                    ElseIf Me.OptType1(1).value = True Then
                        Account_Code_dynamic1 = get_account_code_branch(124, my_branch)
                        If Account_Code_dynamic1 = "NO branch" Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                            GoTo ErrTrap
                        Else
                            If Account_Code_dynamic1 = "NO account" Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ   ··«ÌÃ«—«  «·„” Õﬁ…  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                GoTo ErrTrap
                            End If
                        End If
                
                        If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance1.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtComName.text) & "  " & Trim$(Me.XPTxtComName.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If

                        If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code1").value, Round(Me.TxtOpenBalance1.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtComName.text) & "  " & Trim$(Me.XPTxtComName.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
                    End If
                End If
            End If
'22222222222222222222222222222222222222222222222222222222
endNew:
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata

        Select Case Me.TxtModFlg.text
            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„«·ﬂ" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Done, do you want new Owners"
                End If
            
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
                
                
                 If index = 1 Then
 RSAkar.loadcombo
            RSAkar.dcsupplier.BoundText = val(XPTxtComID.text)
            End If
            
            
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox " Update Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
        End Select
        TxtModFlg.text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "CusID='" & val(XPTxtComID.text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Company()
    Dim Msg As String
    Dim IntRes As Integer
    Dim BegainTrans As Boolean
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtComID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„«·ﬂ   " & CHR(13)
        Msg = Msg + (XPTxtComName.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
    
        If IntRes = vbYes Then
            If Not rs.RecordCount < 1 Then
                DeleteOpeningBalance
                Cn.BeginTrans
                BegainTrans = True
                'StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & Val(Me.XPTxtComID.text)
                'Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
             
                '   update_account_opening_balance get_account_code_branch(19, my_branch)
             
                Dim StrAccountCode As String
                StrAccountCode = rs("Account_Code").value
                '     If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                '         rs.delete
                '     Else
                '         Exit Sub
                '     End If
            
                StrSQL = "Delete From Accounts Where Account_Code='" & StrAccountCode & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords

If SystemOptions.Create2account4Supp = True Then
                    If Not IsNull(rs("Account_Code1").value) Then
                   StrSQL = StrSQL & " or   Account_Code='" & rs("Account_Code1").value & "'"
                   End If
        
        
        
   End If
   
   
                CuurentLogdata ("D")
                rs.delete
                Cn.CommitTrans
                BegainTrans = False
                Msg = " „  ⁄„·Ì… «·Õ–›."
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPBtnMove_Click 2

                If rs.RecordCount < 1 Then
                    clear_all Me
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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate

    If BegainTrans = True Then
        Cn.RollbackTrans
        BegainTrans = False
    End If

    'End If
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „Ê—œ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·„Ê—œ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„Ê—œ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  „Ê—œ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „Ê—œ" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„·«ﬂ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Add New Supplier Data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current Supplier data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit the current Supplier data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the current editing or Save the new Supplier data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the adding new record" & Wrap & "OR undo editing current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current Supplier data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search" & Wrap & "Search for a Supplier..." & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Owners Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Show Help File", BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub printingReport()
    Dim CusReport As ClsCustemerReport
    On Error GoTo ErrTrap

    If XPTxtComID.text <> "" Then
        Set CusReport = New ClsCustemerReport
        CusReport.CustemerData XPTxtComID.text, 2
    End If

    Exit Sub
ErrTrap:

    'On Error GoTo ErrTrap
    'If XPTxtComID.text <> "" Then
    '    Set ComReport = New ClsCompanyReport
    '    ComReport.CompanyData XPTxtComID.text, 2
    'End If
    'Exit Sub
    'ErrTrap:
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

Private Sub ChangeLang()
    'With Me.CboSaleType
    '    .Clear
  
    '        .AddItem "Retail"
    '        .AddItem "WholeSale"
 
    'End With
    Label3.Caption = "Branch"
chkSendMessage.Caption = "Allow SMS"

    With CboDiscountType
        .Clear
        .AddItem "No"
        .AddItem "Value"
        .AddItem "percentage"
    End With

    With CboDiscountTypePur
        .Clear
        .AddItem "no"
        .AddItem "Value"
        .AddItem "percentage"
    End With

    lbl(23).Caption = "Contact person"
    lbl(19).Caption = " type"
    lbl(20).Caption = "Value"

    lbl(29).Caption = " type"
    lbl(28).Caption = "Value"
    lbl(22).Caption = "State"
    lbl(24).Caption = "Province"
    lbl(25).Caption = " City "
    lbl(26).Caption = "Address"
    Fra(5).Caption = "Work Address"
    Fra(4).Caption = "Discounts sales invoices"
    Fra(6).Caption = "Discounts purchase invoices"

    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Owners Data"
    EleHeader.Caption = Me.Caption
    XPLbl(1).Caption = "ID"
    XPLbl(2).Caption = "Code"

    XPLbl(0).Caption = "Owner Name"
    XPLbl(4).Caption = "English Name"
    lbl(3).Caption = "Phone"
    lbl(2).Caption = "Mobile"
    lbl(1).Caption = "Remarks"
    lbl(0).Caption = "Current Record"
    'lbl(4).Caption = "NO. Recordes"
    lbl(7).Caption = "Fax NO."
    lbl(10).Caption = "Credit Limit(Debit)"
    lbl(11).Caption = "Credit Limit(Credit)"
    lbl(12).Caption = "E-Mail."
    lbl(33).Caption = "Parent Acc"
    Me.Fra(1).Caption = "Open Balance"
    Me.Fra(0).Caption = "Open Balance State"
    OptType(0).Caption = "Debit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Un Sign"
    
    lbl(5).Caption = "Balance Value"
    lbl(6).Caption = "Record Date"
    Fra(3).Caption = "Contact Info."
    
    chkCustomerandVendor.Caption = "Owners"
    Label1(2).Caption = "Type"
    Me.Fra(2).Caption = "Current Balance State"
    Me.Cmd(8).Caption = "Customer Balance Report"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    Me.CmdPriceList.Caption = "Owners Price List"

    locked.Caption = "locked"
    ALLButton1.Caption = "Reason"
    lbl(32).Caption = "reason"
    lbl(30).Caption = "period"
    lbl(31).Caption = "period"

End Sub

Private Sub LoadDataCombos(Optional BolExceptCountries As Boolean = False, _
                           Optional BolExceptGovern As Boolean = False, _
                           Optional BolExceptCities As Boolean = False)
    Set Dcombo = New ClsDataCombos

    If BolExceptCountries = False Then
        Dcombo.GetCountriesNames Me.DcboCountryID
        Set cSearch(0) = New clsDCboSearch
        Set cSearch(0).Client = Me.DcboCountryID
    End If

    If BolExceptGovern = False Then
        Dcombo.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID.BoundText)
        Set cSearch(1) = New clsDCboSearch
        Set cSearch(1).Client = Me.DcboGovernmentID
    End If

    If BolExceptCities = False Then
        Dcombo.GetCountriesGovernCities Me.DcboCityID, val(Me.DcboCountryID.BoundText), val(Me.DcboGovernmentID.BoundText)
        Set cSearch(2) = New clsDCboSearch
        Set cSearch(2).Client = Me.DcboCityID
    End If

    Dcombo.GetCustomerType Me.DcCustomerType
    Dcombo.GetBranches dcBranch
  
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

End Sub

Private Sub XPTxtComName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtCusNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
