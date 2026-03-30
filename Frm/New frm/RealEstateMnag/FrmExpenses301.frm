VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmExpenses301 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  ’›Ì… «·⁄Âœ…  «·«„·«ﬂ"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18390
   HelpContextID   =   280
   Icon            =   "FrmExpenses301.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   18390
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ›«’Ì· «·⁄ÂœÂ"
      Height          =   1575
      Left            =   -750
      RightToLeft     =   -1  'True
      TabIndex        =   135
      Top             =   780
      Width           =   6135
      Begin ImpulseAniLabel.ISAniLabel LblLink 
         Height          =   315
         Left            =   -9720
         TabIndex        =   136
         Top             =   720
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   556
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmExpenses301.frx":038A
         BackColor       =   12648447
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel ISAniLabel1 
         Height          =   315
         Left            =   -4800
         TabIndex        =   137
         Top             =   2160
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   556
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmExpenses301.frx":04EC
         BackColor       =   12648447
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLink1 
         Height          =   315
         Left            =   120
         TabIndex        =   138
         Top             =   120
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   556
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmExpenses301.frx":064E
         BackColor       =   12648447
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel ISAniLabel2 
         Height          =   315
         Left            =   120
         TabIndex        =   139
         Top             =   960
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   556
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmExpenses301.frx":07B0
         BackColor       =   12648447
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·—’Ìœ «·Õ«·Ï:"
         Height          =   315
         Index           =   28
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   145
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„Â «·⁄Âœ…"
         Height          =   315
         Index           =   29
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   144
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì «·„‰’—›"
         Height          =   315
         Index           =   30
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   143
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ »ﬁÌ"
         Height          =   315
         Index           =   31
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   142
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„»·€ «·„ÿ·Ê» ··«” ⁄«÷…"
         Height          =   315
         Index           =   32
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   33
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   140
         Top             =   480
         Width           =   3960
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "„·«ÕŸ… Â«„…"
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   2400
      Width           =   5355
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   " ”  Œœ„ Â–… «·‘«‘… · ”ÂÌ·  ”ÊÌ… «·⁄Âœ"
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
         Height          =   420
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.TextBox XPTxtValView 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   7530
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   2925
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   750
      Width           =   12945
      Begin VB.TextBox txt_general_des 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   180
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   134
         Top             =   2280
         Width           =   11295
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         Left            =   6810
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   130
         Top             =   450
         Width           =   1785
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   126
         Top             =   1440
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox TxtAccount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   240
            Width           =   705
         End
         Begin MSDataListLib.DataCombo DcbAccount 
            Height          =   315
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«»"
            Height          =   285
            Index           =   91
            Left            =   5400
            TabIndex        =   129
            Top             =   210
            Width           =   675
         End
      End
      Begin VB.Frame FraNote 
         BackColor       =   &H00E2E9E9&
         Height          =   1455
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   113
         Top             =   720
         Width           =   6075
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   120
            Width           =   705
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   120
            Width           =   705
         End
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   780
            Width           =   4365
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   2850
            TabIndex        =   118
            Top             =   1080
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            Format          =   240975873
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   90
            TabIndex        =   119
            Top             =   450
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   90
            TabIndex        =   120
            Top             =   120
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ"
            Height          =   285
            Index           =   22
            Left            =   6750
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            Height          =   285
            Index           =   19
            Left            =   4380
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   18
            Left            =   4380
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   750
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·»‰ﬂ"
            Height          =   285
            Index           =   17
            Left            =   4380
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·⁄Âœ…"
            Height          =   285
            Index           =   16
            Left            =   4350
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1815
         Left            =   -1740
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   2970
         Visible         =   0   'False
         Width           =   5175
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   330
            Width           =   2655
         End
         Begin VB.TextBox txtto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -1170
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   109
            Top             =   300
            Width           =   2655
         End
         Begin VB.TextBox txt_ORDER_NO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   195
            Index           =   26
            Left            =   -810
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ›« Ê—… «·„Ê—œ"
            Height          =   285
            Index           =   0
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1230
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   990
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10050
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox CboPaymentType1 
         Height          =   315
         Left            =   10050
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   450
         Width           =   1455
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   -240
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text7 
         DataField       =   "id"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   960
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   1110
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TXT_A_NoteID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Text            =   "Text8"
         Top             =   3150
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   7290
         TabIndex        =   0
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   240975873
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   -210
         TabIndex        =   61
         Top             =   3390
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin MSDataListLib.DataCombo dcproject 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   870
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCostCenter 
         Bindings        =   "FrmExpenses301.frx":0912
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   4935
         _ExtentX        =   8705
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
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmExpenses301.frx":0927
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
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
      Begin MSDataListLib.DataCombo DCVendor 
         Height          =   315
         Left            =   0
         TabIndex        =   132
         Top             =   30
         Visible         =   0   'False
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·‘—Õ «·⁄«„"
         Height          =   285
         Index           =   20
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   146
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Ê—œ"
         Height          =   285
         Index           =   34
         Left            =   4440
         TabIndex        =   133
         Top             =   0
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         Height          =   195
         Index           =   15
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   131
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Index           =   0
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         Height          =   285
         Index           =   4
         Left            =   11490
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›« "
         Height          =   285
         Index           =   3
         Left            =   -510
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   8940
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   135
         Width           =   555
      End
      Begin VB.Image ImgNote 
         Height          =   240
         Left            =   -240
         Picture         =   "FrmExpenses301.frx":093C
         Top             =   750
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‘—Ê⁄"
         Height          =   255
         Index           =   14
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   870
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
         Height          =   255
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   405
         Index           =   21
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   1950
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «· ”ÊÌ…"
         Height          =   285
         Index           =   23
         Left            =   11490
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   510
         Width           =   1275
      End
   End
   Begin VB.OptionButton OptSort 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   1
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   240
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkLastAccount 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9300
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   8190
      Width           =   1905
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
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
      Height          =   1035
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   10980
      Width           =   6465
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   35
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   37
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   12
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   13
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   11
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   10
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   9
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4320
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   645
      Left            =   23160
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   210
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   7650
      Width           =   2145
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Index           =   0
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   18375
      _cx             =   32411
      _cy             =   1349
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
      Picture         =   "FrmExpenses301.frx":0EC6
      Caption         =   "    ’›Ì… «·⁄Âœ…  «·«„·«ﬂ"
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
      PicturePos      =   0
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
      Begin VB.CheckBox ChkPurchaseFixedAssets 
         Alignment       =   1  'Right Justify
         Caption         =   "›« Ê—… ‘—«¡ «’·"
         Height          =   195
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox oldTxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   11
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
         ButtonImage     =   "FrmExpenses301.frx":1BA0
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
         Left            =   630
         TabIndex        =   12
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
         ButtonImage     =   "FrmExpenses301.frx":1F3A
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
         Left            =   2220
         TabIndex        =   13
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
         ButtonImage     =   "FrmExpenses301.frx":22D4
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
         Left            =   1155
         TabIndex        =   14
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
         ButtonImage     =   "FrmExpenses301.frx":266E
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   4680
         Top             =   0
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
         Left            =   2640
         Top             =   0
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   12240
         Picture         =   "FrmExpenses301.frx":2A08
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Label LblShortcutKeys 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
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
         Height          =   195
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   510
         Width           =   5445
      End
   End
   Begin MSDataListLib.DataCombo XPCboExpensesType 
      Height          =   315
      Left            =   22920
      TabIndex        =   6
      Top             =   2760
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   12540
      TabIndex        =   18
      Top             =   8190
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   16470
      TabIndex        =   24
      Top             =   7620
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   15570
      TabIndex        =   25
      Top             =   7620
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   14760
      TabIndex        =   26
      Top             =   7620
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   3
      Left            =   13965
      TabIndex        =   27
      Top             =   7620
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   4
      Left            =   13050
      TabIndex        =   28
      Top             =   7620
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   6
      Left            =   9570
      TabIndex        =   29
      Top             =   7620
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdHelp 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   10410
      TabIndex        =   30
      Top             =   7620
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   5
      Left            =   12240
      TabIndex        =   31
      Top             =   7620
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   6060
      TabIndex        =   42
      Top             =   8040
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "„—«ﬂ“ «· ﬂ·›…"
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
      MICON           =   "FrmExpenses301.frx":6670
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   8
      Left            =   11310
      TabIndex        =   44
      Top             =   7620
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   873
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   9
      Left            =   5640
      TabIndex        =   45
      Top             =   10200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ALLButtonS.ALLButton CmdRemove 
      Height          =   375
      Left            =   17430
      TabIndex        =   46
      Tag             =   "Delete Row"
      Top             =   7620
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–› ”ÿ—"
      ENAB            =   0   'False
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
      MICON           =   "FrmExpenses301.frx":668C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   10
      Left            =   8010
      TabIndex        =   50
      Top             =   8190
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   4470
      TabIndex        =   77
      Top             =   8160
      Width           =   915
      _ExtentX        =   1614
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
   Begin ALLButtonS.ALLButton CMDRemoveAll 
      Height          =   375
      Left            =   17430
      TabIndex        =   78
      Tag             =   "Delete Row"
      Top             =   8100
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–› «·ﬂ·"
      ENAB            =   0   'False
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
      MICON           =   "FrmExpenses301.frx":66A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   3870
      Left            =   0
      TabIndex        =   79
      Top             =   3660
      Width           =   18450
      _cx             =   32544
      _cy             =   6826
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "«·„ﬁ»Ê÷« |Õ«·…«·«⁄ „«œ"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "FrmExpenses301.frx":66C4
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3405
         Left            =   45
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   45
         Width           =   18360
         _cx             =   32385
         _cy             =   6006
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
         Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
            Height          =   3225
            Left            =   0
            TabIndex        =   81
            Top             =   60
            Width           =   18180
            _cx             =   32067
            _cy             =   5689
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   28
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmExpenses301.frx":6A5E
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
            Begin VB.PictureBox PicDes 
               BorderStyle     =   0  'None
               Height          =   1635
               Left            =   240
               RightToLeft     =   -1  'True
               ScaleHeight     =   1635
               ScaleWidth      =   2925
               TabIndex        =   82
               Top             =   960
               Visible         =   0   'False
               Width           =   2925
               Begin VB.TextBox TxtDes 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   1125
                  Left            =   30
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   83
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2115
               End
               Begin VB.Label LblDes 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000C&
                  Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                  ForeColor       =   &H0000C8FF&
                  Height          =   315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   0
                  Width           =   2445
               End
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   3285
            Left            =   -870
            TabIndex        =   85
            Top             =   30
            Width           =   18870
            _cx             =   33285
            _cy             =   5794
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   49
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmExpenses301.frx":6E79
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
            Begin VB.Frame Frame3 
               Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
               Height          =   1215
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   6720
               Visible         =   0   'False
               Width           =   4215
               Begin VB.CommandButton Command5 
                  Caption         =   "‰”Œ"
                  Height          =   255
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   255
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   3915
               Left            =   2550
               RightToLeft     =   -1  'True
               ScaleHeight     =   3915
               ScaleWidth      =   9405
               TabIndex        =   86
               Top             =   0
               Visible         =   0   'False
               Width           =   9405
               Begin VB.TextBox TxtDese 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   1485
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   90
                  Top             =   2040
                  Width           =   8955
               End
               Begin VB.TextBox txtcodesub 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   3600
                  Width           =   855
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "Add des"
                  Height          =   255
                  Left            =   7440
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   3600
                  Width           =   1350
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "Call des"
                  Height          =   255
                  Left            =   6240
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   3600
                  Width           =   1095
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   3900
                  Left            =   120
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   10905
                  _cx             =   19235
                  _cy             =   6879
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial (Arabic)"
                     Size            =   20.25
                     Charset         =   178
                     Weight          =   700
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
                  Caption         =   ""
                  Align           =   0
                  AutoSizeChildren=   0
                  BorderWidth     =   6
                  ChildSpacing    =   4
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
                  PicturePos      =   7
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
                  Begin VB.TextBox Text3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1605
                     Left            =   0
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   92
                     Top             =   480
                     Visible         =   0   'False
                     Width           =   8955
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H8000000C&
                     Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                     ForeColor       =   &H0000C8FF&
                     Height          =   315
                     Left            =   6840
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   0
                     Width           =   2445
                  End
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Code"
                  Height          =   495
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   3480
                  Width           =   735
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Height          =   495
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   1200
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Code"
                  Height          =   255
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   1320
                  Width           =   735
               End
            End
            Begin VDSCOMBOLibCtl.SmartCombo CboDes 
               Height          =   315
               Left            =   0
               TabIndex        =   101
               ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
               Top             =   0
               Visible         =   0   'False
               Width           =   2955
               _cx             =   1973752924
               _cy             =   1973748268
               Alignment       =   0
               Appearance      =   3
               AutoSearch      =   0   'False
               BackColor       =   -2147483624
               BackgroundColor =   -2147483633
               BorderColor     =   0
               BorderVisible   =   -1  'True
               Caption         =   "SmartCombo1"
               CaptionAlignment=   4
               CaptionBackColor=   -2147483633
               BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionForeColor=   -2147483630
               CaptionHeight   =   15
               CaptionOnTop    =   0   'False
               CaptionMultiLine=   0
               Checkbox3D      =   0   'False
               CheckboxAlignment=   5
               CheckboxBackColor=   16777215
               CheckboxSize    =   13
               CheckboxValue   =   0
               BrowsePictureAlignment=   5
               BrowsePictureStretchH=   0
               BrowsePictureStretchV=   0
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
               ForeColor       =   0
               Gap             =   0
               HideSelection   =   -1  'True
               Locked          =   0   'False
               MaxLength       =   0
               MultiLine       =   0
               OnFocus         =   3
               PasswordChar    =   ""
               Picture         =   "FrmExpenses301.frx":759A
               PictureAlignment=   5
               PictureBackColor=   -2147483624
               PictureStretchH =   0
               PictureStretchV =   0
               Redraw          =   -1  'True
               ScrollBar       =   0
               Style           =   0
               Text            =   ""
               UnderLine       =   0   'False
               Enabled0        =   -1  'True
               Position0       =   0
               Tip0            =   "Caption"
               Visible0        =   0   'False
               Width0          =   90
               Enabled1        =   -1  'True
               Position1       =   1
               Tip1            =   ""
               Visible1        =   -1  'True
               Width1          =   32
               Enabled2        =   -1  'True
               Position2       =   2
               Tip2            =   "Check Box (Space, Ctrl + Space)"
               Visible2        =   0   'False
               Width2          =   16
               Enabled3        =   -1  'True
               Position3       =   3
               Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
               Visible3        =   -1  'True
               Width3          =   145
               Enabled4        =   -1  'True
               Position4       =   4
               Tip4            =   "Left Spinner (Alt + Left)"
               Visible4        =   0   'False
               Width4          =   16
               Enabled5        =   -1  'True
               Position5       =   5
               Tip5            =   "Right Spinner (Alt + Right)"
               Visible5        =   0   'False
               Width5          =   16
               Enabled6        =   -1  'True
               Position6       =   6
               Tip6            =   "Up Spinner (Ctrl + Up)"
               Visible6        =   0   'False
               Width6          =   16
               Enabled7        =   -1  'True
               Position7       =   7
               Tip7            =   "Down Spinner (Ctrl + Down)"
               Visible7        =   0   'False
               Width7          =   16
               Enabled8        =   -1  'True
               Position8       =   8
               Tip8            =   "Browse (Alt + Enter)"
               Visible8        =   0   'False
               Width8          =   16
               Enabled9        =   -1  'True
               Position9       =   9
               Tip9            =   " (Alt + Down, F4)"
               Visible9        =   -1  'True
               Width9          =   16
               Enabled10       =   -1  'True
               Position10      =   10
               Tip10           =   "Right Arrow (Alt + >)"
               Visible10       =   0   'False
               Width10         =   16
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   3405
         Left            =   19095
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   45
         Width           =   18360
         _cx             =   32385
         _cy             =   6006
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
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   2775
            Left            =   120
            TabIndex        =   103
            Tag             =   "1"
            Top             =   0
            Width           =   18180
            _cx             =   32067
            _cy             =   4895
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
            FormatString    =   $"FrmExpenses301.frx":7B34
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
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   165
            Left            =   6510
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   2820
            Width           =   3315
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   180
            Left            =   6510
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   4230
            Visible         =   0   'False
            Width           =   3315
         End
      End
   End
   Begin ImpulseButton.ISButton Accredit 
      Height          =   390
      Left            =   3090
      TabIndex        =   106
      Top             =   8130
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   688
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   24
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   2160
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
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
      Height          =   75
      Index           =   27
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   8040
      Width           =   5835
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   300
      Index           =   8
      Left            =   14325
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   8205
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   255
      Left            =   11340
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   8220
      Width           =   735
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7680
      Width           =   6015
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   1140
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   9330
      Width           =   555
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   9330
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      Height          =   435
      Index           =   6
      Left            =   930
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   9330
      Width           =   165
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
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
      Height          =   435
      Index           =   7
      Left            =   1530
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8130
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ã„«·Ì"
      Height          =   285
      Index           =   2
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   7650
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«„—"
      Height          =   285
      Index           =   5
      Left            =   16440
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   1515
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmExpenses301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim branch_no  As String
Dim RsNotes As ADODB.Recordset
Dim BolEditOnMainAccounts As Boolean
Dim Balance As String
Dim balanceString As String
Public LngRow As Double
Public LngCol As Double

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
'If Me.TxtModFlg.Text <> "R" Then
        If CboPayMentType.ListIndex = 4 Then
            Me.DcboCreditSide.BoundText = DcbAccount.BoundText
        End If
' End If
End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 20190719
    End If


End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.text)
End Sub

Sub HidFat()
With Me.VSFlexGrid1
If True = True Then
.ColHidden(.ColIndex("Vat")) = False
.ColHidden(.ColIndex("Vatyo")) = False
Else
.ColHidden(.ColIndex("Vat")) = True
.ColHidden(.ColIndex("Vatyo")) = True
End If
End With
With Fg_Journal
If True = True Then
.ColHidden(.ColIndex("Vat")) = False
.ColHidden(.ColIndex("Vatyo")) = False
Else
.ColHidden(.ColIndex("Vat")) = True
.ColHidden(.ColIndex("Vatyo")) = True
End If
End With
End Sub
Function saveChequeBoxContents1(NoteID As Double)

    If SystemOptions.banks_Accounts3 = False Then Exit Function
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords
 
 '   rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     * from dbo.TblChecqueBoxContent1 Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    If CboPayMentType.ListIndex = 1 Then
        rs.AddNew
        rs("noteid").value = NoteID
     
        rs("RecordDate").value = XPDtbTrans.value
        rs("DueDate").value = DtpChequeDueDate.value
        rs("BankID").value = val(DcboBankName.BoundText)
        rs("BankName").value = DcboBankName.text
        
        rs("ChequeNo").value = TxtChequeNumber.text
        rs("ChequeValue").value = val(XPTxtVal.text)
    
        rs("Remarks").value = Me.DcboDebitSide.text
        rs("Payed").value = 0
       
        rs("DepitAccount").value = (DcboDebitSide.BoundText)
        rs("notes_all").value = NoteID
      
        rs.update
    End If

    rs.Close
End Function

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
 
    SendTopost Me.Name, "notes_all", "NoteID", 0, val(dcBranch.BoundText), val(XPTxtID.text), TxtSerial1.text
  '' RsNetes.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
fillapprovData
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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
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
 Grid2.rows = 1
    End If
RsDetails.Close

End Function
Private Sub ALLButton1_Click()
    On Error GoTo ErrTrap

    If DcCostCenter.BoundText <> "" Then

        MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        Exit Sub
    End If

    Dim opr_id As Double

    If Not IsNumeric(Text1.text) Then Exit Sub
    'If Me.TxtModFlg.text = "N" Then
    opr_id = val(Me.Text1.text)
    'Else
    'opr_id = TxtDEV_NO.text
    'End If

If CboPaymentType1.ListIndex = 0 Then

    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
        If Not val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE"))) = 0 Then

            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If
            
    End If

Else
    If Not VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountCode")) = "" Then
        If Not val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("VALUE"))) = 0 Then

            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("LineNo1"))
        
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If
            
    End If

End If

    marakes_taklefa_tawze3.opr_type = "›« Ê—… „«·Ì…"
    marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
    marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
    marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
 If CboPaymentType1.ListIndex = 0 Then
    marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
 Else
 marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountCode")) & "' and  line_no=" & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("LineNo1"))
 End If
    marakes_taklefa_tawze3.Adodc3.Refresh
    '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdRemoveAll_Click()
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.rows = 3
            Fg_Journal.Enabled = True
          
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 3
            VSFlexGrid1.Enabled = True
End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub

Private Sub CboPayMentType_Change()

    FraNote.Visible = False
 Frame12(2).Visible = False
 lbl(34).Visible = False
    If Me.TxtModFlg.text = "E" Then
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.DcboBox.text = ""
        DCVendor.text = ""
    End If

    If Me.CboPayMentType.ListIndex = 0 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        DcbAccount.BoundText = ""
        FraNote.Visible = True
    ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Me.DCVendor.Enabled = False
        DcbAccount.BoundText = ""
     FraNote.Visible = True
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
    
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Me.DcboBox.Enabled = False
        Me.DCVendor.Enabled = True
        lbl(34).Visible = True
        FraNote.Visible = True
        ElseIf Me.CboPayMentType.ListIndex = 4 Then
     
        DcbAccount_Change
        Frame12(2).Visible = True
        FraNote.Visible = False
         '   Cmd(13).Enabled = False
        Me.DcboBox.Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
          DcboBankName.BoundText = ""
        'TxtChequeNumber.text = ""

FraNote.Visible = False
    Else
        FraNote.Visible = False
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Function setfoxy()
    Text1.text = CStr(new_id("foxy", "id", "", True))

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id").value = Text1.text
 
    rs.update
    
End Function

Private Sub CboPaymentType1_Change()
VSFlexGrid1.Visible = True
'Fg_Journal.Visible = False

    If Me.CboPaymentType1.ListIndex = 0 Then
        Fg_Journal.Enabled = True
        VSFlexGrid1.Visible = False
        Fg_Journal.Visible = True
        Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Fg_Journal.rows = 3
          
       ' Fg_Journal.Visible = True

    ElseIf Me.CboPaymentType1.ListIndex = 1 Then

        Fg_Journal.Enabled = False
        VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        VSFlexGrid1.rows = 3
        VSFlexGrid1.Visible = True
        
        Fg_Journal.Visible = False
        
        
            ElseIf Me.CboPaymentType1.ListIndex = 2 Then
 
        'VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        'VSFlexGrid1.Rows = 3
        'VSFlexGrid1.Enabled = True
        '
       
      '  VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
      '  VSFlexGrid1.Rows = 3
      '  VSFlexGrid1.Enabled = True
        
        
    End If

End Sub

Private Sub CboPaymentType1_Click()
    CboPaymentType1_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            DcCostCenter.text = ""
            Grid2.Clear flexClearScrollable, flexClearEverything
            Grid2.rows = 1
            ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
Accredit.Caption = ""
            Me.DCboUserName.BoundText = user_id
            '        XPDtbTrans.SetFocus
'            Fg_Journal.Visible = False
'            VSFlexGrid1.Visible = False

            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.rows = 3
            Fg_Journal.Enabled = True
          
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 3
            VSFlexGrid1.Enabled = True
          
            DtpChequeDueDate.value = Date
            setfoxy
            CBoBasedON.ListIndex = 0
            CboPaymentType1.ListIndex = 0
            Me.dcBranch.BoundText = branch_id
CboPaymentType1.ListIndex = 1
            CboPayMentType.ListIndex = 0
            
            DcboBox_Change

        Case 1
                    
             If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
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
              
            Dim Msg As String

            If SystemOptions.banks_Accounts3 = True Then
                If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
    
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
         '   Me.DCboUserName.BoundText = user_id
            Fg_Journal.rows = Fg_Journal.rows + 1
            Fg_Journal.Enabled = True
            VSFlexGrid1.rows = VSFlexGrid1.rows + 1
            VSFlexGrid1.Enabled = True
        
            CuurentLogdata

        Case 2
        If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If CBoBasedON.ListIndex > 0 And Trim(TXT_order_no.text) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify NO For"
                Else
                    Msg = "Õœœ —ﬁ„ "
                End If

                Msg = Msg & "  " & CBoBasedON.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TXT_order_no.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText
            DcboBox_Change
            DcboBankName_Change
            DCVendor_Click (0)

            SaveData
           
        Case 3
            Undo

        Case 4
            If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not delete.This process associated with approvals"
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

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

          Load FrmNotesSearch
          FrmNotesSearch.SearchType = 360
         FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            ViewDataList

        Case 8
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report TxtSerial.text, DCVendor.text

        Case 9
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_Cheque TxtChequeNumber.text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtSerial.text

        Case 10
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

          '  ShowGL_cc TxtSerial.Text, , 200
    ShowGL_cc TxtSerial.text, , 350, , TxtSerial1.text
    End Select

    Exit Sub
ErrTrap:
End Sub

Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
    hide_logo = True
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Expanses_Order  where ChqueNum='" & ChqueNum & "' and noteserial='" & TxtSerial & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    Else
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
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

    Dim cCompanyInfo As New ClsCompanyInfo
    'MsgBox ToHijriDate(Date)

    xReport.ParameterFields(5).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 1, 2)
    xReport.ParameterFields(6).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 4, 2)
    xReport.ParameterFields(7).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 9, 2)

    xReport.ParameterFields(8).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
    xReport.ParameterFields(9).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
    xReport.ParameterFields(10).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
    xReport.ParameterFields(11).AddCurrentValue CStr(txtto.text)
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.text)
    xReport.ParameterFields(13).AddCurrentValue CStr(Me.XPMTxtRemarks.text)
    xReport.ParameterFields(14).AddCurrentValue CStr(LblValue.Caption)
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Function print_report(Optional NoteSerial As String, Optional VendorName As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = " SELECT     dbo.TblExpensesDet301.ID, dbo.TblExpensesDet301.ExpID, dbo.TblExpensesDet301.FlgVat, dbo.TblExpensesDet301.ForcedFlg, dbo.TblExpensesDet301.CurrRow, "
MySQL = MySQL & "                      dbo.TblExpensesDet301.BrnchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExpensesDet301.[Value],"
MySQL = MySQL & "                      dbo.TblExpensesDet301.Vatyo, dbo.TblExpensesDet301.Vat, dbo.TblExpensesDet301.Des, dbo.TblExpensesDet301.billno, dbo.TblExpensesDet301.Unitss,"
MySQL = MySQL & "                      dbo.TblExpensesDet301.StrUnit, dbo.TblExpensesDet301.Departementid, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
MySQL = MySQL & "                      dbo.TblExpensesDet301.NEmpid, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblExpensesDet301.Aqarid,"
MySQL = MySQL & "                      dbo.TblAqar.aqarname, dbo.TblExpensesDet301.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblExpensesDet301.UnitNo,"
MySQL = MySQL & "                      TblAqarDetai_1.unitno AS UnitnoName55, dbo.TblExpensesDet301.AccountCode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial,"
MySQL = MySQL & "                      dbo.ACCOUNTS.Account_NameEng, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.ChqueNum,"
MySQL = MySQL & "                      dbo.notes_all.DueDate, dbo.notes_all.NoteHijriDate, dbo.notes_all.BoxID, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE, dbo.notes_all.NoteSerial1,"
MySQL = MySQL & "                      dbo.notes_all.general_des, dbo.notes_all.note_value_by_characters, dbo.TblExp301UnitNo.Valu, dbo.TblExp301UnitNo.UnitID, TblAqarDetai_1.unitno AS UnitNoName,"
MySQL = MySQL & "                       dbo.TblExpensesDet301.projectid, dbo.projects.Fullcode AS ProjectFullcode, dbo.projects.Project_name, dbo.TblExpensesDet301.pandid,"
MySQL = MySQL & "                      dbo.projects_des.des AS Panddes, dbo.TblExpensesDet301.operid, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE"
MySQL = MySQL & " FROM         dbo.TblProcessDEF RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblExpensesDet301 ON dbo.TblProcessDEF.TblProcessDEFID = dbo.TblExpensesDet301.operid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects_des ON dbo.TblExpensesDet301.pandid = dbo.projects_des.oprid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects ON dbo.TblExpensesDet301.projectid = dbo.projects.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblExp301UnitNo LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAqarDetai TblAqarDetai_1 ON dbo.TblExp301UnitNo.UnitID = TblAqarDetai_1.Id ON"
MySQL = MySQL & "                      dbo.TblExpensesDet301.ID = dbo.TblExp301UnitNo.ExpDetails RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBoxesData RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.notes_all ON dbo.TblBoxesData.BoxID = dbo.notes_all.BoxID ON dbo.TblExpensesDet301.ExpID = dbo.notes_all.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.ACCOUNTS ON dbo.TblExpensesDet301.AccountCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAqarDetai TblAqarDetai_2 ON dbo.TblExpensesDet301.UnitNo = TblAqarDetai_2.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAkarUnit ON dbo.TblExpensesDet301.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAqar ON dbo.TblExpensesDet301.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblExpensesDet301.NEmpid = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments ON dbo.TblExpensesDet301.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblExpensesDet301.BrnchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblExpensesDet301.ExpID = " & val(XPTxtID.text) & ")"
'MySQL = MySQL + "order by dev_id_line_no"

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\New Reports\ohdaIaqr.rpt"
        Else
            StrFileName = App.path & "\Reports\New Reports\ohdaIaqr.rpt"
        End If

'     If CboPaymentType1.ListIndex = 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\Reports\New Reports\" & "ohda1.rpt"
'        Else
'            StrFileName = App.path & "\Reports\New Reports\" & "ohda1.rpt"
'        End If
'
'    Else
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\Reports\New Reports\" & "ohda.rpt"
'        Else
'            StrFileName = App.path & "\Reports\New Reports\" & "ohda.rpt"
'        End If
'
'    End If
'
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
    xReport.ParameterFields(6).AddCurrentValue VendorName
    xReport.ParameterFields(7).AddCurrentValue balanceString
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

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1, "0612201404"

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
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

    sql = "Delete  marakes_taklefa_temp where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
    Cn.Execute sql, , adExecuteNoRecords
    
    If CboPaymentType1.ListIndex = 0 Then
        If Fg_Journal.rows > 1 Then
            If Fg_Journal.rows = 2 Then
                Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.Fg_Journal.rows > 1 Then
                    If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                                        
                        With Me.Fg_Journal
Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                            If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
                            LogTextA = "  Õ–› «·„’—Ê›   " & .cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                            LogTexte = "  Delete  Expensen   " & .cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                                                         
                            AddToLogFile CInt(user_id), 350, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
                        End With
                                                        
                        
                    End If
                End If
            End If
        End If
            
        With Fg_Journal
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With

    ElseIf CboPaymentType1.ListIndex = 1 Then

        If VSFlexGrid1.rows > 1 Then
            If VSFlexGrid1.rows = 2 Then
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.VSFlexGrid1.rows > 1 Then
                    If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                         Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                        With Me.VSFlexGrid1

                            If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
                            LogTextA = "  Õ–› «·Õ”«»   " & .cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                            LogTexte = "  Delete  Account   " & .cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                                                         
                            AddToLogFile CInt(user_id), 350, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
                        End With
                        
                       
                    End If
                End If
            End If
        End If
            
        With VSFlexGrid1
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With
             
    Else
 
        Exit Sub
    End If

End Sub
Private Sub DcboBankName_Change()
    On Error Resume Next

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
    
        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

        If CboPayMentType.ListIndex = 3 Then
                     
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

    End If

End Sub

Private Sub DcboBox_Change()
Dim acc As String

    
    If DcboBox.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    
    acc = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    WriteCustomerBalPublic acc, Balance, balanceString
    LblLink1.Caption = balanceString
    End If

    
End Sub

Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub DcboCreditSide_Change()

    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
End Sub



Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub Fg_Journal_KeyPress(KeyAscii As Integer)
' SendKeys "{F4}"
'  SendKeys "{BACKSPACE}"
'  SendKeys CHR(KeyAscii)
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub

Private Sub LblLink_Click()
 
    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboCreditSide.BoundText, DcboCreditSide.text, FirstPeriod, Date

End Sub

Private Sub LblLink_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
 
    If SystemOptions.UserInterface = ArabicInterface Then
        LblLink.ToolTipText = "—’Ìœ «·⁄Âœ…:" & WriteNo(Balance, 0, True)
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        LblLink.ToolTipText = "Credit Balance:" & WriteNo(Balance, 0, True)
    End If
 
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtSerial.text = ""
    TxtSerial1.text = ""
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 3
    End If

End Sub

Private Sub DCVendor_Click(Area As Integer)

    If DCVendor.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If

    Text2.text = Me.DCVendor.BoundText
End Sub
Sub DeleteGridCurrRow(Optional CurrRow As Long)
Dim i As Integer
With Fg_Journal
i = .rows
Do
i = i - 1
If val(.TextMatrix(i, .ColIndex("CurrRow"))) = CurrRow Then
.RemoveItem i
End If
Loop While i > 1
End With
End Sub
Sub DeleteGridCurrRowExp(Optional CurrRow As Long)
Dim i As Integer
With VSFlexGrid1
i = .rows
Do
i = i - 1
If val(.TextMatrix(i, .ColIndex("CurrRow"))) = CurrRow Then
.RemoveItem i
End If
Loop While i > 1
End With
End Sub
Sub AddVATExp(Optional Row As Long)
If True = True Then
Dim ForcedFlg As Integer
Dim valuee As Double
Dim AccountVATDept As String
Dim i As Integer
Dim k As Integer
Dim ClsAcc  As New ClsAccounts

Dim flg As Integer
With VSFlexGrid1
valuee = val(.TextMatrix(Row, .ColIndex("Value")))
.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(.TextMatrix(Row, .ColIndex("BrnchID"))), ForcedFlg)
.TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg
.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)
GetValueAddedAccount XPDtbTrans.value, AccountVATDept


.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(.TextMatrix(Row, .ColIndex("BrnchID"))), ForcedFlg)
If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccounProject(XPDtbTrans.value, flg, val(.TextMatrix(Row, .ColIndex("BrnchID"))), ForcedFlg)
End If
.TextMatrix(Row, .ColIndex("Rate")) = val(.TextMatrix(Row, .ColIndex("Vatyo"))) / 100 + 1
If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) > 0 And val(.TextMatrix(Row, .ColIndex("Rate"))) > 0 Then
.TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("PriceTotal"))) / val(.TextMatrix(Row, .ColIndex("Rate"))), 2)
End If
valuee = val(.TextMatrix(Row, .ColIndex("Value")))
.TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg
.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)


''/////////////
If val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
   If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
    DeleteGridCurrRowExp Row
   For i = 1 To 1
         .AddItem " ", .Row + i
  k = .Row + i
.TextMatrix(k, .ColIndex("CurrRow")) = Row
 
If i = 1 Then
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(AccountVATDept)
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_Name(, AccountVATDept)
.TextMatrix(k, .ColIndex("AccountCode")) = AccountVATDept
.TextMatrix(k, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Vat"))
Else
.TextMatrix(k, .ColIndex("AccountCode")) = DcboCreditSide.BoundText
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_Name(, DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Vat"))
End If
.TextMatrix(k, .ColIndex("projectid")) = .TextMatrix(Row, .ColIndex("projectid"))
.TextMatrix(k, .ColIndex("BrnchID")) = .TextMatrix(Row, .ColIndex("BrnchID"))
.TextMatrix(k, .ColIndex("Des")) = .TextMatrix(Row, .ColIndex("Des")) & " " & " ﬁÌ„… „÷«›…"
'.TextMatrix(k, .ColIndex("operid")) = .TextMatrix(Row, .ColIndex("operid"))
'.TextMatrix(k, .ColIndex("pandid")) = .TextMatrix(Row, .ColIndex("pandid"))
'.TextMatrix(k, .ColIndex("ProjectCode")) = .TextMatrix(Row, .ColIndex("ProjectCode"))
.TextMatrix(k, .ColIndex("branch_name")) = .TextMatrix(Row, .ColIndex("branch_name"))
'.TextMatrix(k, .ColIndex("project")) = .TextMatrix(Row, .ColIndex("project"))
'.TextMatrix(k, .ColIndex("pand")) = .TextMatrix(Row, .ColIndex("pand"))
'.TextMatrix(k, .ColIndex("oper")) = .TextMatrix(Row, .ColIndex("oper"))
.TextMatrix(k, .ColIndex("FixedAsset")) = .TextMatrix(Row, .ColIndex("FixedAsset"))
.TextMatrix(k, .ColIndex("Departementid")) = .TextMatrix(Row, .ColIndex("Departementid"))
'.TextMatrix(k, .ColIndex("fixedid")) = .TextMatrix(Row, .ColIndex("fixedid"))
.TextMatrix(k, .ColIndex("FixedAssetId")) = .TextMatrix(Row, .ColIndex("FixedAssetId"))
.TextMatrix(k, .ColIndex("NEmpid")) = .TextMatrix(Row, .ColIndex("NEmpid"))
.TextMatrix(k, .ColIndex("NEmpName")) = .TextMatrix(Row, .ColIndex("NEmpName"))
'.TextMatrix(k, .ColIndex("Aqarid")) = .TextMatrix(Row, .ColIndex("Aqarid"))
'.TextMatrix(k, .ColIndex("UnitType")) = .TextMatrix(Row, .ColIndex("UnitType"))
'.TextMatrix(k, .ColIndex("aqarname")) = .TextMatrix(Row, .ColIndex("aqarname"))
'.TextMatrix(k, .ColIndex("name")) = .TextMatrix(Row, .ColIndex("name"))
'.TextMatrix(k, .ColIndex("UnitNo")) = .TextMatrix(Row, .ColIndex("UnitNo"))
.TextMatrix(k, .ColIndex("Departement")) = .TextMatrix(Row, .ColIndex("Departement"))
'.TextMatrix(k, .ColIndex("unitnoName")) = .TextMatrix(Row, .ColIndex("unitnoName"))
.TextMatrix(k, .ColIndex("FlgVat")) = 1
Next i
End If
End If
End With
End If
End Sub
Sub AddVAT(Optional Row As Long)
If True = True Then
Dim ForcedFlg As Integer
Dim valuee As Double
Dim AccountVATDept As String
Dim i As Integer
Dim k As Integer
Dim ClsAcc  As New ClsAccounts
With Fg_Journal

.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(.TextMatrix(Row, .ColIndex("BrnchID"))), ForcedFlg)
.TextMatrix(Row, .ColIndex("Rate")) = val(.TextMatrix(Row, .ColIndex("Vatyo"))) / 100 + 1
If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) > 0 And val(.TextMatrix(Row, .ColIndex("Rate"))) > 0 Then
    .TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("PriceTotal"))) / val(.TextMatrix(Row, .ColIndex("Rate"))), 2)
End If
valuee = val(.TextMatrix(Row, .ColIndex("value")))
.TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg
.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)
GetValueAddedAccount XPDtbTrans.value, AccountVATDept
If AccountVATDept = "" And val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «·Õ”«» «·„œÌ‰ ›Ì ‘«‘… «⁄œ«œ  «·›« "
Else
MsgBox "Please Enter Account In VAT Settings"
End If
.TextMatrix(Row, .ColIndex("Vat")) = 0
.TextMatrix(Row, .ColIndex("Vatyo")) = 0
Exit Sub
End If


If val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
   If Not .TextMatrix(Fg_Journal.Row, .ColIndex("AccountCode")) = "" Then
    DeleteGridCurrRow Row
   For i = 1 To 1
         .AddItem " ", Fg_Journal.Row + i
  k = .Row + i
.TextMatrix(k, .ColIndex("CurrRow")) = Row
 
If i = 1 Then
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(AccountVATDept)
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_Name(, AccountVATDept)
.TextMatrix(k, .ColIndex("AccountCode")) = AccountVATDept
.TextMatrix(k, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("Vat"))
Else
.TextMatrix(k, .ColIndex("AccountCode")) = DcboCreditSide.BoundText
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_Name(, DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("Vat"))
End If
.TextMatrix(k, .ColIndex("projectid2")) = .TextMatrix(Row, .ColIndex("projectid2"))
.TextMatrix(k, .ColIndex("BrnchID")) = .TextMatrix(Row, .ColIndex("BrnchID"))
.TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " ﬁÌ„… „÷«›…"
'.TextMatrix(k, .ColIndex("pandid2")) = .TextMatrix(Row, .ColIndex("pandid2"))
'.TextMatrix(k, .ColIndex("operid2")) = .TextMatrix(Row, .ColIndex("operid2"))
.TextMatrix(k, .ColIndex("ExpensesID")) = .TextMatrix(Row, .ColIndex("ExpensesID"))
.TextMatrix(k, .ColIndex("branch_name")) = .TextMatrix(Row, .ColIndex("branch_name"))
.TextMatrix(k, .ColIndex("opr_fullcode")) = .TextMatrix(Row, .ColIndex("opr_fullcode"))
.TextMatrix(k, .ColIndex("Order_No")) = .TextMatrix(Row, .ColIndex("Order_No"))
'.TextMatrix(k, .ColIndex("ProjectCode")) = .TextMatrix(Row, .ColIndex("ProjectCode"))
'.TextMatrix(k, .ColIndex("project")) = .TextMatrix(Row, .ColIndex("project"))
'.TextMatrix(k, .ColIndex("pand")) = .TextMatrix(Row, .ColIndex("pand"))
.TextMatrix(k, .ColIndex("fixedid")) = .TextMatrix(Row, .ColIndex("fixedid"))
'.TextMatrix(k, .ColIndex("oper")) = .TextMatrix(Row, .ColIndex("oper"))
.TextMatrix(k, .ColIndex("Fixes")) = .TextMatrix(Row, .ColIndex("Fixes"))
.TextMatrix(k, .ColIndex("FlgVat")) = 1
Next i
End If
End If
End With
End If
End Sub
Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim Rs3 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal
    If val(.TextMatrix(Row, .ColIndex("BrnchID"))) = 0 Then
.TextMatrix(Row, .ColIndex("BrnchID")) = val(dcBranch.BoundText)
.TextMatrix(Row, .ColIndex("branch_name")) = dcBranch.text
End If

        Select Case .ColKey(Col)
        
 Case "Vatyo"
        If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
        .TextMatrix(Row, .ColIndex("Vat")) = 0
        If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) <> 0 Then
        .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("PriceTotal")))
        End If
        If .rows > Row Then
        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
        .RemoveItem Row + 1
        End If
        End If
        End If
        
                Case "PriceTotal"
                AddVAT Row

        
                 Case "branch_name"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("BrnchID")) = StrAccountCode
                AddVAT Row
         Case "Fixes"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("fixedid")) = StrAccountCode
                
         Case "project"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("projectid2")) = StrAccountCode
               If val(.TextMatrix(Row, .ColIndex("projectid2"))) <> 0 Then
               StrSQL = "Select Fullcode from  Projects where ID =" & val(.TextMatrix(Row, .ColIndex("projectid2"))) & ""
               Set Rs3 = New ADODB.Recordset
               Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If Rs3.RecordCount > 0 Then
               .TextMatrix(Row, .ColIndex("ProjectCode")) = IIf(IsNull(Rs3("Fullcode").value), "", Rs3("Fullcode").value)
               Else
               .TextMatrix(Row, .ColIndex("ProjectCode")) = ""
               End If
               End If
   Case "Vatyo"
        If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
        .TextMatrix(Row, .ColIndex("Vat")) = 0
        If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) <> 0 Then
        .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("PriceTotal")))
        End If
        If .rows > Row Then
        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
        .RemoveItem Row + 1
        End If
        End If
        End If
         Case "ProjectCode"
               If .TextMatrix(Row, .ColIndex("ProjectCode")) <> "" Then
               StrSQL = "select * from  Projects where Fullcode ='" & .TextMatrix(Row, .ColIndex("ProjectCode")) & "'"
                Set Rs3 = New ADODB.Recordset
               Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If Rs3.RecordCount > 0 Then
               .TextMatrix(Row, .ColIndex("projectid2")) = IIf(IsNull(Rs3("ID").value), "", Rs3("ID").value)
               If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(Rs3("Project_name").value), "", Rs3("Project_name").value)
               Else
               .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(Rs3("Project_nameE").value), "", Rs3("Project_nameE").value)
               End If
               Else
               .TextMatrix(Row, .ColIndex("projectid2")) = 0
               .TextMatrix(Row, .ColIndex("project")) = ""
               End If
               End If
         Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("pandid2")) = StrAccountCode
         Case "oper"
                 StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("operid2")) = StrAccountCode

         Case "ExpensesID"
              
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                .TextMatrix(Row, .ColIndex("Order_No")) = TXT_order_no.text
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                   .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                   .TextMatrix(Row, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If
                AddVAT Row
Case "Account_Serial"
If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Serial='" & .TextMatrix(Row, .ColIndex("Account_Serial")) & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Serial='" & .TextMatrix(Row, .ColIndex("Account_Serial")) & "'"
                End If
                  rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                    .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    End If
                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                  .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                    .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(.TextMatrix(Row, .ColIndex("AccountCode")))
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                .TextMatrix(Row, .ColIndex("Order_No")) = TXT_order_no.text
                      If CheckAccountHaveDestributions(.TextMatrix(Row, .ColIndex("AccountCode"))) = True Then
             
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " Â–« «·„’—Ê› ·Â ŒÿÂ  Ê“Ì⁄  ⁄·Ï «·›—Ê⁄ Â·  —Ìœ «· Ê“Ì⁄  " & CHR(13)
                        Msg = Msg + "‰⁄„ «„ ·« "
                          
                    Else
                        Msg = " This Expenses Have Destribution Plan Do you want  Destribute  " & CHR(13)
                        Msg = Msg + "Yes Or No"
                    
                    End If
                                 
                    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                        .TextMatrix(Row, .ColIndex("Destribute")) = 1
         
                    Else
                        .TextMatrix(Row, .ColIndex("Destribute")) = 0
                    End If
            
                End If
            
                End If
                AddVAT Row
            Case "value", "opr_fullcode"
                Dim sgl As String
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
               
                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
                sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
        AddVAT Row
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

    With Me.Fg_Journal

        If Me.TxtModFlg <> "E" Then Exit Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
            LogTextA = "   ⁄œÌ· «·„’—Ê› «·Ï " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Account To " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Value") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„…  «·Ï " & .cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " ··„’—Ê›   " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change value" & .cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " To Expenses " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Des") Then
            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï " & .cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " ··„’—Ê›   " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Des " & .cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Expenses " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        End If

        AddToLogFile CInt(user_id), 350, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
    End With

End Sub

Function calcnets()

    If Me.CboPaymentType1.ListIndex = 0 Then

        With Fg_Journal
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With

    Else

        With Me.VSFlexGrid1
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With

    End If

End Function

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

If val(.TextMatrix(Row, .ColIndex("FlgVat"))) <> 0 Then
   Cancel = True
Else
 Select Case .ColKey(Col)
        Case "Vat"
                 Cancel = True
 Case "PriceTotal"
                .ComboList = ""
        Case "Vatyo"
              If val(.TextMatrix(Row, .ColIndex("ForcedFlg"))) = 1 Then
                 Cancel = True
              Else
              .ComboList = ""
              End If
        Case "ProjectCode"
                .ComboList = ""
     Case "LineNo"
                .ComboList = ""
     
     
            Case "Account_Serial"
                .ComboList = ""
                   Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
                '  Cancel = True
            
            Case "Order_No"
                .ComboList = ""
        End Select
End If
    End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
           CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.cell(flexcpData, r, c)) <> "String" Then
            TxtDes.text = ""
        Else
            '
            TxtDes.text = Fg_Journal.cell(flexcpData, r, c)
        End If

        ' show new note
       CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
      CboDes.Visible = True
       CboDes.ZOrder 0
       CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                
                    Order_no_search.show
                    Order_no_search.RetrunType = 4
                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.show
                    FrmExpensesSearch.RetrunType = 350
                End If
 
        End Select

    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
        Case "Fixes"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    ISEQUP=1 or  id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE   ISEQUP=1 or  id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
       Case "branch_name"
             StrSQL = "  SELECT     branch_id, branch_name, branch_namee"
             StrSQL = StrSQL & " From dbo.TblBranchesData"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_name", "branch_id")
            Else
                   StrComboList = Fg_Journal.BuildComboList(rs, "branch_namee", "branch_id")
           End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
                
                    Case "project"

               
                StrSQL = " SELECT  LTRIM(RTRIM( Project_name )) as Project_name,Project_nameE , id From dbo.Projects  "

         
             If SystemOptions.UserInterface = ArabicInterface Then
    
        StrSQL = StrSQL & " where  not (Project_name is null)and Project_name<>N'""'"
    Else
       
        StrSQL = StrSQL & " where  not (Project_nameE is null)and Project_nameE<>N'""'"
    End If
    StrSQL = StrSQL & " and (Not (Fullcode Is Null))"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by  Project_name"
    Else
    StrSQL = StrSQL & " order by  Project_nameE"
    End If
    
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = Fg_Journal.BuildComboList(rs, "Project_name", "id")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
             Case "pand"
             If .TextMatrix(Row, .ColIndex("projectid2")) = "" Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Exit Sub
             End If

                StrSQL = " SELECT     des, oprid From projects_des "
                 StrSQL = StrSQL & "    Where (project_id =" & val(.TextMatrix(Row, .ColIndex("projectid2"))) & ")"
           
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = Fg_Journal.BuildComboList(rs, "des", "oprid")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                  Case "oper"
                   
If .TextMatrix(Row, .ColIndex("projectid2")) = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
If .TextMatrix(Row, .ColIndex("pandid2")) = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·»‰œ «Ê·«"
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
           
                If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               Else
               StrSQL = "SELECT     dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEF"
                End If
               StrSQL = StrSQL & "    Where (ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("pandid2"))) & ") And (project_id = " & val(.TextMatrix(Row, .ColIndex("projectid2"))) & ")"
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "ProcessName", "TblProcessDEFID")
                    Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "ProcessNameE", "TblProcessDEFID")
                    End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        
        
        
        

            Case "AccountName"
                 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts order by Account_Name"
                Else
                    StrSQL = "select * from Expenses_accounts_eng order by Account_Nameeng"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

           '     If SystemOptions.UserInterface = ArabicInterface Then
           '         StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
           '     Else
           '         StrComboList = Fg_Journal.BuildComboList(rs, "Account_NameEng", "Account_Code")
           '     End If
          If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "parent_account,account_serial,*Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "parent_account,account_serial,*Account_NameEng", "Account_Code")
                End If
                
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                  
            Case "opr_fullcode"
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")

                If SystemOptions.Items_or_operation = 1 Then
                    StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,name", "fullcode")
                ElseIf SystemOptions.Items_or_operation = 0 Then
                    StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,des", "fullcode")
         
                End If

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim StrComboList  As String
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap
    ScreenNameArabic = " ’›Ì… «·⁄Âœ…  "
    ScreenNameEnglish = "Era Exchange"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 350
HidFat
'    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
'    fill_combo Me.DcCostCenter, StrSQL
                      If SystemOptions.SpecialVersion = True Then
                     Label3.Visible = False
TxtSerial.Visible = False
Cmd(10).Visible = False
    End If
   
 


    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("FillData").Picture
    Resize_Form Me
    AddTip
    SetDtpickerDate XPDtbTrans
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetExpensesType XPCboExpensesType
    Dcombos.GetBranches Me.dcBranch
        Dcombos.GetAccountingCodes Me.DcbAccount, True, False
'    Dim Dcombos As ClsDataCombos
'Set Dcombos = New ClsDataCombos
Dcombos.GetCostCenter DcCostCenter

    
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.XPCboExpensesType

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide


    
        With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem " ÕÊÌ· »‰ﬂÌ"
        .AddItem "  ‘Ìﬂ „”œœ"
        .AddItem "Õ”«»"
    End With


    With Me.CboPaymentType1
        .Clear
        .AddItem "„’«—Ì›"
        .AddItem "Õ”«»« "
    '    .AddItem "„’«—Ì› ÊÕ”«»« "
    End With

    With Me.CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "√„— ‘—¡"
        .AddItem "›« Ê—… „»œ∆ÌÂ"
        .AddItem " «„— «‰ «Ã  "
    
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    StrSQL = " select expanses_account,Project_name from projects  where   Project_name<>N'""' and not (Project_name is null)  and not(expanses_account is null)"
    fill_combo dcproject, StrSQL

    'StrSQL = " select  CusID, CusName from TblCustemers  where Type=2"
    'fill_combo Me.DCVendor, StrSQL

    Dcombos.GetCustomersSuppliers 3, Me.DCVendor
With Me.VSFlexGrid1
           
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '   End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                    End If
                
                Else
                
           '         StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name   from ACCOUNTS   Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '     End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If
                
                End If
                Set rs = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
End With


    
    Set rs = New ADODB.Recordset
    StrSQL = "select * From notes_all where notetype=350 and bill_Type<>2"
   StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
   
        If SystemOptions.FixedCustomer = 1 Then
                              StrSQL = StrSQL & " and  UserID = " & user_id
                               End If
                               
        If SystemOptions.usertype <> UserAdminAll Then
       ' StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
    
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 350
    hide_logo = False

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

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

    If ButtonID = vdsDownArrow Then
        If CboDes.IsDropped = False Then
            If PicHeight > 0 Then
                PicDes.Height = PicHeight
                PicDes.Width = PicWidth
            Else
                PicDes.Width = CboDes.Width - 10
                PicDes.Height = CboDes.Height * 8
            End If

            Debug.Print PicHeight
            Debug.Print PicWidth
            TxtDes.Visible = True
            TxtDes.text = Fg_Journal.cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            CboDes.DropDown PicDes.hWnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Sendkeys "{F4}"
    End If

End Sub

Private Sub LblLink1_Click()
  If SystemOptions.SpecialVersion = True Then
        Exit Sub
End If
    
    
  Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboCreditSide.BoundText, DcboCreditSide.text, FirstPeriod, Date

End Sub

Private Sub PicDes_Resize()

    With PicDes
        LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If CBoBasedON.ListIndex = 3 Then
        If KeyCode = vbKeyF3 Then
            Order_no_search2.show
            Order_no_search2.RetrunType = 2
         
        End If

    Else

        If KeyCode = vbKeyF3 Then
            Order_no_search.show
            Order_no_search.RetrunType = 1
        End If

    End If

End Sub

Private Sub TxtDes_LostFocus()
    PicHeight = PicDes.Height
    PicWidth = PicDes.Width
    CboDes.CloseUp
    CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
        PutData
        CboDes.CloseUp
    End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
        
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = False
            Frame1.Enabled = False
        
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            CmdRemove.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            XPTxtVal.locked = True
            '        XPCboProfLevel.Locked = True
            '        XPTxtProfMail.Locked = True
            '        XPTxtPhone.Locked = True
            '        XPTxtMobile.Locked = True
            XPMTxtRemarks.locked = True
            XPCboExpensesType.locked = True
            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False
     CMDRemoveAll.Enabled = False
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            
            End If

        Case "N"
        
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            CmdRemove.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
        
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True
       
            CmdRemove.Enabled = True
            CMDRemoveAll.Enabled = True
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
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    'check_cost_center
    
    
    
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
    Dim Rs3 As ADODB.Recordset
    With VSFlexGrid1
If val(.TextMatrix(Row, .ColIndex("BrnchID"))) = 0 Then
.TextMatrix(Row, .ColIndex("BrnchID")) = val(dcBranch.BoundText)
.TextMatrix(Row, .ColIndex("branch_name")) = dcBranch.text
End If
        Select Case .ColKey(Col)
                     
        Case "PriceTotal"
                AddVATExp Row
         Case "Vatyo"
        If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
        .TextMatrix(Row, .ColIndex("Vat")) = 0
        If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) <> 0 Then
        .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("PriceTotal")))
        End If
        If .rows > Row Then
        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
        .RemoveItem Row + 1
        End If
        End If
        End If
            Case "branch_name"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("BrnchID")) = StrAccountCode
                AddVATExp Row
          Case "project"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("projectid")) = StrAccountCode
                If val(.TextMatrix(Row, .ColIndex("projectid"))) <> 0 Then
               StrSQL = "Select Fullcode from  Projects where ID =" & val(.TextMatrix(Row, .ColIndex("projectid"))) & ""
               Set Rs3 = New ADODB.Recordset
               Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If Rs3.RecordCount > 0 Then
               .TextMatrix(Row, .ColIndex("ProjectCode")) = IIf(IsNull(Rs3("Fullcode").value), "", Rs3("Fullcode").value)
               Else
               .TextMatrix(Row, .ColIndex("ProjectCode")) = ""
               End If
               End If
         Case "ProjectCode"
               If .TextMatrix(Row, .ColIndex("ProjectCode")) <> "" Then
               StrSQL = "select * from  Projects where Fullcode ='" & .TextMatrix(Row, .ColIndex("ProjectCode")) & "'"
                Set Rs3 = New ADODB.Recordset
               Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If Rs3.RecordCount > 0 Then
               .TextMatrix(Row, .ColIndex("projectid")) = IIf(IsNull(Rs3("ID").value), "", Rs3("ID").value)
               If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(Rs3("Project_name").value), "", Rs3("Project_name").value)
               Else
               .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(Rs3("Project_nameE").value), "", Rs3("Project_nameE").value)
               End If
               Else
               .TextMatrix(Row, .ColIndex("projectid")) = 0
               .TextMatrix(Row, .ColIndex("project")) = ""
               End If
               End If
                  Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("pandid")) = StrAccountCode
                  Case "oper"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("operid")) = StrAccountCode
                
           Case "aqarname"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("Aqarid")) = StrAccountCode
         Case "name"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("unittype")) = StrAccountCode
         Case "unitnoName"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("unitno")) = StrAccountCode
                
     Case "NEmpName"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("NEmpid")) = StrAccountCode
                
                Case "Departement"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("Departementid")) = StrAccountCode
        
        

        
                   Case "FixedAsset"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("FixedAssetId")) = StrAccountCode



            Case "Value"
               .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
                AddVATExp Row
            Case "DebitValue", "CreditValue"

                'remove destribution
     
                ' sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                ' Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    ' Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                    '    Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    ' Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '     Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))

                If .ColKey(Col) = "DebitValueE" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                    End If

                    '
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If
 
                    
            
            Case "Account_Serial"
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                StrSQL = StrSQL & GetAccountByBarnchUser
                 StrSQL = StrSQL & GetAccountCodeHiding
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
                    
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    Dim rs2 As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
                    AddVATExp Row
xx:
                Else
                    'GetMsgs 130, vbExclamation
                    MsgBox "ﬂÊœ Õ”«» Œ«ÿÏ¡", vbCritical
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

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
              AddVATExp Row
                Set ClsAcc = Nothing
            
                StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
                StrSQL = StrSQL & GetAccountByBarnchUser
                 StrSQL = StrSQL & GetAccountCodeHiding
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
            
                    'Dim rs2 As ADODB.Recordset
                    'Dim My_SQL As String
                    If IsNull(rs("currenct_code").value) Then
                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo ll
                    End If

                    My_SQL = "  select * from currency WHERE id=" & rs("currenct_code").value

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
                    AddVATExp Row
ll:
                End If

        End Select

        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ReLineGrid

    End With

    With Me.VSFlexGrid1

        If Me.TxtModFlg <> "E" Then Exit Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
            LogTextA = "   ⁄œÌ· «·Õ”«» «·Ï " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Account To " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Value") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„…  «·Ï " & .cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " ··Õ”«»   " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change value" & .cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " To Account " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Des") Then
            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï " & .cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " ··Õ”«»   " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Des " & .cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Account " & .cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        End If

        AddToLogFile CInt(user_id), 350, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If
If val(.TextMatrix(Row, .ColIndex("FlgVat"))) <> 0 Then
   Cancel = True
Else
 Select Case .ColKey(Col)
  Case "PriceTotal"
                .ComboList = ""
        Case "Vat"
                 Cancel = True
        Case "Vatyo"
              If val(.TextMatrix(Row, .ColIndex("ForcedFlg"))) = 1 Then
                 Cancel = True
              Else
              .ComboList = ""
              End If
     Case "LineNo", "Unitss"
                .ComboList = ""
     Case "ProjectCode"
                .ComboList = ""
                
            Case "Value"
                .ComboList = ""
                 Case "billno"
                .ComboList = ""

            Case "Account_Serial"
                .ComboList = ""
            Case "Des"
                .ComboList = ""
            Case "billno"
                .ComboList = ""
                '  Cancel = True
            
        End Select
    End If
 End With

End Sub

Private Sub VSFlexGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Me.VSFlexGrid1
Select Case .ColKey(Col)
   Case "unitnoName"
      If val(.TextMatrix(Row, .ColIndex("UnitType"))) <> 0 And val(.TextMatrix(Row, .ColIndex("Aqarid"))) <> 0 Then
           LngRow = Row
           LngCol = Col
           
          FrmIqarUnitNo.TypIndex = 2
           Load FrmIqarUnitNo
           FrmIqarUnitNo.TypIndex = 2
           FrmIqarUnitNo.show vbModal
           
      Else
        If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄ﬁ«— Ê«·‰Ê⁄"
        Else
         MsgBox "Please Select Real Estate"
       End If
       Exit Sub
        End If
End Select
End With
End Sub

Private Sub VSFlexGrid1_KeyPress(KeyAscii As Integer)
'  SendKeys "{F4}"
'  SendKeys "{BACKSPACE}"
'  SendKeys CHR(KeyAscii)
End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
       Account_search.show
       Account_search.case_id = 350350

    End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim Rs3 As ADODB.Recordset
    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)
              Case "unitnoName"
                 .ColComboList(.ColIndex("unitnoName")) = "..."
                Case "branch_name"
                StrSQL = "SELECT     branch_id, branch_name, branch_namee"
                StrSQL = StrSQL & "     From dbo.TblBranchesData"
                Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                  If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "branch_name", "branch_id")
                    Else
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "branch_namee", "branch_id")
                  End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
             Case "project"

               
                StrSQL = " SELECT  LTRIM(RTRIM( Project_name )) as Project_name ,Project_nameE, id From dbo.Projects  "
                    If SystemOptions.UserInterface = ArabicInterface Then
    
        StrSQL = StrSQL & " where  not (Project_name is null)and Project_name<>N'""'"
    Else
        
        StrSQL = StrSQL & " where  not (Project_nameE is null)and Project_nameE<>N'""'"
    End If
    StrSQL = StrSQL & " and (Not (Fullcode Is Null))"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by  Project_name"
    Else
    StrSQL = StrSQL & " order by  Project_nameE"
    End If
    
                Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "Project_name", "id")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
             Case "pand"
             If .TextMatrix(Row, .ColIndex("projectid")) = "" Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Exit Sub
             End If

                StrSQL = " SELECT     des, oprid From projects_des "
                 StrSQL = StrSQL & "    Where (project_id =" & val(.TextMatrix(Row, .ColIndex("projectid"))) & ")"
                Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "des", "oprid")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                  Case "oper"
                   
If .TextMatrix(Row, .ColIndex("projectid")) = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
If .TextMatrix(Row, .ColIndex("pandid")) = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·»‰œ «Ê·«"
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
           
                If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               Else
               StrSQL = "SELECT     dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEF"
                End If
               StrSQL = StrSQL & "    Where (ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("pandid"))) & ") And (project_id = " & val(.TextMatrix(Row, .ColIndex("projectid"))) & ")"
               Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "ProcessName", "TblProcessDEFID")
                    Else
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "ProcessNameE", "TblProcessDEFID")
                    End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
  

Case "FixedAsset"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE  ISEQUP=1 or    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE   ISEQUP=1 or  id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
         Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "Name", "id")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                


            Case "Departement"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     TOP 100 PERCENT DeparmentID, DepartmentName FROM         dbo.TblEmpDepartments ORDER BY DepartmentName  "
                Else
                    StrSQL = " SELECT     TOP 100 PERCENT DeparmentID, DepartmentNamee FROM         dbo.TblEmpDepartments ORDER BY DepartmentNamee   "
                End If
               Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "DepartmentNamee", "DeparmentID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            Case "NEmpName"


                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "   SELECT     TOP 100 PERCENT Emp_ID, Emp_Name from dbo.TblEmployee ORDER BY Emp_Name "
                Else
                    StrSQL = "   SELECT     TOP 100 PERCENT Emp_ID, Emp_Namee from dbo.TblEmployee ORDER BY Emp_Namee "
                End If
         Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "Emp_Namee", "Emp_ID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "aqarname"

                StrSQL = " SELECT     TOP 100 PERCENT Aqarid, aqarname from TblAqar ORDER BY aqarname "
                Set Rs3 = New ADODB.Recordset
                 Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "aqarname", "Aqarid")
            

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
           Case "unitnoName"
           Dim aqr As Integer
           Dim unt As Integer
           
           If val(.TextMatrix(Row, .ColIndex("Aqarid"))) <> 0 Then
If val(.TextMatrix(Row, .ColIndex("unittype"))) <> 0 Then
aqr = val(.TextMatrix(Row, .ColIndex("Aqarid")))
unt = val(.TextMatrix(Row, .ColIndex("unittype")))
                StrSQL = " SELECT     TOP 100 PERCENT id, unitno from TblAqarDetai where (Aqarid =" & aqr & ") and (unittype=" & unt & ") ORDER BY unitno "
     Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs3.RecordCount > 0 Then
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "unitno", "id")
            End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                
                Else
                .ComboList = ""
                MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·ÊÕœÂ «Ê·«"
                StrComboList = ""
              Exit Sub
                End If
                Else
                .ComboList = ""
                MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄ﬁ«—"
                StrComboList = ""
                Exit Sub
                End If
                .ComboList = StrComboList
              Case "name"


                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "   SELECT     TOP 100 PERCENT Id, name from dbo.TblAkarUnit ORDER BY name "
                Else
                    StrSQL = "   SELECT     TOP 100 PERCENT Id, namee from dbo.TblAkarUnit ORDER BY namee "
                End If
               Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "name", "ID")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "namee", "ID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "AccountName"
               ' Exit Sub
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, SUBSTRING(ACCOUNTS.Account_NameEng, 0, 20)  As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
StrSQL = StrSQL & GetAccountByBarnchUser
StrSQL = StrSQL & GetAccountCodeHiding
                    '   End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                    End If
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, SUBSTRING(ACCOUNTS.Account_Name, 0, 20) As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                 
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                 StrSQL = StrSQL & GetAccountByBarnchUser
                 StrSQL = StrSQL & GetAccountCodeHiding
                 
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If
                
                End If
                Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 
              StrComboList = VSFlexGrid1.BuildComboList(Rs3, "RootName,ParentName,*FirstName", "Account_Code")
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 3
          
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 2
          
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        'Lngid
        '  If XPTxtID.text <> 0 Then
        '      Rs.find "NoteID=" & XPTxtID.text, , adSearchForward, adBookmarkFirst
        '      If Rs.EOF Or Rs.BOF Then
        '          Exit Sub
        '      End If
        '  End If
  
        If Lngid <> 0 Then
            rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If

    End If

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If

    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.TXT_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)

    TXT_A_NoteID.text = IIf(IsNull(rs("A_NoteID").value), "", val(rs("A_NoteID").value))

    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    txtto.text = IIf(IsNull(rs("too").value), "", rs("too").value)
    txt_general_des.text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)
DcbAccount.BoundText = IIf(IsNull(rs("AccountPaym").value), "", rs("AccountPaym").value)
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    XPCboExpensesType.BoundText = IIf(IsNull(rs("ExpensesID").value), "", rs("ExpensesID").value)

    If (rs("bill_Type").value) = 0 Then
        Me.CboPaymentType1.ListIndex = 0
    ElseIf (rs("bill_Type").value) = 1 Then
        Me.CboPaymentType1.ListIndex = 1
    Else
        Me.CboPaymentType1.ListIndex = 0
    End If

    CboPaymentType1_Change

    If Not IsNull(rs("BasedONID").value) Then
        Me.CBoBasedON.ListIndex = rs("BasedONID").value
    Else
        Me.CBoBasedON.ListIndex = 0
 
    End If

 If IsNull(rs("NoteCashingType").value) Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPayMentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
 
    ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPayMentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPayMentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        ElseIf rs("NoteCashingType").value = 4 Then
        Me.CboPayMentType.ListIndex = 4
    End If


    CboPayMentType_Change

    'ÿMe.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), "", Rs("BoxID").value)
    'DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))

    If rs("NoteCashingType").value = 0 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    ElseIf rs("NoteCashingType").value = 1 Or rs("NoteCashingType").value = 3 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
    ElseIf rs("NoteCashingType").value = 2 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt_Numorder.text = IIf(IsNull(rs("NumOrderInpot").value), "", rs("NumOrderInpot").value)
    Me.TxtSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)

    Me.TxtSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)

    Me.oldTxtSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)

    Me.dcproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)
Dim acc As String
    acc = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    WriteCustomerBalPublic acc, Balance, balanceString
    LblLink1.Caption = balanceString
    
    If CboPaymentType1.ListIndex = 1 Then 'Õ”«Ì« 

       ' StrSQL = "SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, "
       ' StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.UserID , dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[value],dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description"
       ' StrSQL = StrSQL + " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
       ' StrSQL = StrSQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
       ' StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
       ' StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

'StrSQL = " SELECT     TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, "
'StrSQL = StrSQL + "                      dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.UserID,"
'StrSQL = StrSQL + "                      dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
'StrSQL = StrSQL + "                      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.DOUBLE_ENTREY_VOUCHERS.NEmpid,"
'StrSQL = StrSQL + "                      dbo.DOUBLE_ENTREY_VOUCHERS.Departementid, dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId, dbo.FixedAssets.Name AS fixedassetname,"
'StrSQL = StrSQL + "                      dbo.FixedAssets.namee AS fixedassetnamee, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
'StrSQL = StrSQL + "                      dbo.TblEmployee.Emp_Name AS NEmpName, dbo.TblEmployee.Emp_Namee AS NEmpNamee, dbo.DOUBLE_ENTREY_VOUCHERS.Billno,"
'StrSQL = StrSQL + "                      dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.DOUBLE_ENTREY_VOUCHERS.unitno,"
'StrSQL = StrSQL + "                      dbo.TblAqarDetai.unitno AS unitnoName, dbo.DOUBLE_ENTREY_VOUCHERS.unittype, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee,"
'StrSQL = StrSQL + "                      dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.projects.Project_name, dbo.DOUBLE_ENTREY_VOUCHERS.pandid, dbo.projects_des.des,"
'StrSQL = StrSQL + "                      dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id,"
'StrSQL = StrSQL + "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee ,dbo.projects.Fullcode as ProjectCode"
'StrSQL = StrSQL + "  , dbo.DOUBLE_ENTREY_VOUCHERS.Vat , dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo , dbo.DOUBLE_ENTREY_VOUCHERS.FlgVat ,dbo.DOUBLE_ENTREY_VOUCHERS.CurrRow"
'StrSQL = StrSQL + " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
'StrSQL = StrSQL + "                      dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.TblProcessDEF ON dbo.DOUBLE_ENTREY_VOUCHERS.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.projects_des ON dbo.DOUBLE_ENTREY_VOUCHERS.pandid = dbo.projects_des.oprid and dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS.projectid = dbo.projects.id LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.TblAkarUnit ON dbo.DOUBLE_ENTREY_VOUCHERS.unittype = dbo.TblAkarUnit.id LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.TblAqarDetai ON dbo.DOUBLE_ENTREY_VOUCHERS.unitno = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.TblAqar ON dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.TblEmployee ON dbo.DOUBLE_ENTREY_VOUCHERS.NEmpid = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.TblEmpDepartments ON dbo.DOUBLE_ENTREY_VOUCHERS.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
'StrSQL = StrSQL + "                      dbo.FixedAssets ON dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = dbo.FixedAssets.id"
'StrSQL = StrSQL + " Where ( hideline is null and dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(RS("A_NoteID").value) & ")"
'StrSQL = StrSQL & "   ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
StrSQL = " SELECT    DISTINCT dbo.TblExpensesDet301.ID, dbo.TblExpensesDet301.ExpID, dbo.TblExpensesDet301.FlgVat, dbo.TblExpensesDet301.ForcedFlg, dbo.TblExpensesDet301.CurrRow, "
StrSQL = StrSQL & "                      dbo.TblExpensesDet301.BrnchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExpensesDet301.[Value],"
StrSQL = StrSQL & "                      dbo.TblExpensesDet301.Vatyo, dbo.TblExpensesDet301.Vat, dbo.TblExpensesDet301.Des, dbo.TblExpensesDet301.billno, dbo.TblExpensesDet301.Unitss,"
StrSQL = StrSQL & "                      dbo.TblExpensesDet301.StrUnit, dbo.TblExpensesDet301.Departementid, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
StrSQL = StrSQL & "                      dbo.TblExpensesDet301.NEmpid, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblExpensesDet301.Aqarid,"
StrSQL = StrSQL & "                      dbo.TblAqar.aqarname, dbo.TblExpensesDet301.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblExpensesDet301.UnitNo,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.unitno AS UnitNoName, dbo.TblExpensesDet301.AccountCode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial,"
StrSQL = StrSQL & "                      dbo.ACCOUNTS.Account_NameEng, dbo.TblExpensesDet301.projectid, dbo.projects.Project_name, dbo.projects.Fullcode AS ProjectFullcode,"
StrSQL = StrSQL & "                      dbo.TblExpensesDet301.operid, dbo.projects_des.des AS Panddes, dbo.TblExpensesDet301.pandid, dbo.TblProcessDEF.ProcessName,"
StrSQL = StrSQL & "                      dbo.TblProcessDEF.ProcessNameE"
StrSQL = StrSQL & " FROM         dbo.TblExpensesDet301 LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProcessDEF ON dbo.TblExpensesDet301.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.projects_des ON dbo.TblExpensesDet301.pandid = dbo.projects_des.oprid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.projects ON dbo.TblExpensesDet301.projectid = dbo.projects.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.ACCOUNTS ON dbo.TblExpensesDet301.AccountCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblExpensesDet301.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblExpensesDet301.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblExpensesDet301.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblExpensesDet301.NEmpid = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblExpensesDet301.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblExpensesDet301.BrnchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " Where (dbo.TblExpensesDet301.ExpID = " & val(XPTxtID.text) & ")"


'
'StrSQL = " SELECT     TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.Carid,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, "
'StrSQL = StrSQL + "                       dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.UserID,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.projects.Fullcode AS ProjectFullcode,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.DOUBLE_ENTREY_VOUCHERS.NEmpid,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Departementid, dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId, "
'StrSQL = StrSQL + "                        dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
'StrSQL = StrSQL + "                       dbo.TblEmployee.Emp_Name AS NEmpName, dbo.TblEmployee.Emp_Namee AS NEmpNamee, dbo.DOUBLE_ENTREY_VOUCHERS.Billno,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.DOUBLE_ENTREY_VOUCHERS.unitno,"
'StrSQL = StrSQL + "                       dbo.TblAqarDetai.unitno AS unitnoName, dbo.DOUBLE_ENTREY_VOUCHERS.unittype, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.projects.Project_name, dbo.DOUBLE_ENTREY_VOUCHERS.pandid, dbo.projects_des.des ,dbo.projects_des.des  AS Panddes,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id,"
'StrSQL = StrSQL + "                       dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode AS ProjectCode, dbo.DOUBLE_ENTREY_VOUCHERS.Vat,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo, dbo.DOUBLE_ENTREY_VOUCHERS.FlgVat, dbo.DOUBLE_ENTREY_VOUCHERS.CurrRow,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Rate2, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierName, dbo.DOUBLE_ENTREY_VOUCHERS.CusVATNO,"
'StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.PriceTotal, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID"
'
'StrSQL = " SELECT   DISTINCT  dbo.TblExpensesDet301.ID, dbo.TblExpensesDet301.ExpID, dbo.TblExpensesDet301.FlgVat, dbo.TblExpensesDet301.ForcedFlg, dbo.TblExpensesDet301.CurrRow, "
'StrSQL = StrSQL & "                      dbo.TblExpensesDet301.BrnchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExpensesDet301.[Value],"
'StrSQL = StrSQL & "                      dbo.TblExpensesDet301.Vatyo, dbo.TblExpensesDet301.Vat, dbo.TblExpensesDet301.Des, dbo.TblExpensesDet301.billno, dbo.TblExpensesDet301.Unitss,"
'StrSQL = StrSQL & "                      dbo.TblExpensesDet301.StrUnit, dbo.TblExpensesDet301.Departementid, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
'StrSQL = StrSQL & "                      dbo.TblExpensesDet301.NEmpid, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblExpensesDet301.Aqarid,"
'StrSQL = StrSQL & "                      dbo.TblAqar.aqarname, dbo.TblExpensesDet301.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblExpensesDet301.UnitNo,"
'StrSQL = StrSQL & "                      dbo.TblAqarDetai.unitno AS UnitNoName, dbo.TblExpensesDet301.AccountCode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial,"
'StrSQL = StrSQL & "                      dbo.ACCOUNTS.Account_NameEng, dbo.TblExpensesDet301.projectid, dbo.projects.Project_name, dbo.projects.Fullcode AS ProjectFullcode,"
'StrSQL = StrSQL & "                      dbo.TblExpensesDet301.operid, dbo.projects_des.des AS Panddes, dbo.TblExpensesDet301.pandid, dbo.TblProcessDEF.ProcessName,"
'StrSQL = StrSQL & "                      dbo.TblProcessDEF.ProcessNameE,dbo.DOUBLE_ENTREY_VOUCHERS.PriceTotal"
'StrSQL = StrSQL + "  FROM         dbo.DOUBLE_ENTREY_VOUCHERS RIGHT  Outer JOIN"
'
'StrSQL = StrSQL + "                        TblExpensesDet301 On TblExpensesDet301.ExpID = dbo.DOUBLE_ENTREY_VOUCHERS.notes_id "
'StrSQL = StrSQL + "                        LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblProcessDEF ON dbo.TblExpensesDet301.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.projects_des ON dbo.TblExpensesDet301.pandid = dbo.projects_des.oprid LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.projects ON dbo.TblExpensesDet301.projectid = dbo.projects.id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.ACCOUNTS ON dbo.TblExpensesDet301.AccountCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblExpensesDet301.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblExpensesDet301.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblExpensesDet301.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblExpensesDet301.NEmpid = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblExpensesDet301.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblExpensesDet301.BrnchID = dbo.TblBranchesData.branch_id"
'
''StrSQL = StrSQL & "  Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = 15596)"
'StrSQL = StrSQL & " Where (dbo.TblExpensesDet301.ExpID = " & val(XPTxtID.text) & ")"
'StrSQL = StrSQL + " AND  (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
        
        
StrSQL = " SELECT DISTINCT E.ID, E.ExpID, E.FlgVat, E.ForcedFlg, E.CurrRow, "
StrSQL = StrSQL & " E.BrnchID, B.branch_name, B.branch_namee, E.[Value], "
StrSQL = StrSQL & " E.Vatyo, E.Vat, E.Des, E.billno, E.Unitss, "
StrSQL = StrSQL & " E.StrUnit, E.Departementid, D.DepartmentName, D.DepartmentNamee, "
StrSQL = StrSQL & " E.NEmpid, Emp.Emp_Name, Emp.Fullcode, Emp.Emp_Namee, E.Aqarid, "
StrSQL = StrSQL & " Aqar.aqarname, E.UnitType, AU.name, AU.namee, E.UnitNo, "
StrSQL = StrSQL & " AD.unitno AS UnitNoName, E.AccountCode, Acc.Account_Name, Acc.Account_Serial, "
StrSQL = StrSQL & " Acc.Account_NameEng, E.projectid, Proj.Project_name, Proj.Fullcode AS ProjectFullcode, "
StrSQL = StrSQL & " E.operid, PD.des AS Panddes, E.pandid, ProcDef.ProcessName, "
StrSQL = StrSQL & " ProcDef.ProcessNameE, VoucherData.PriceTotal "

StrSQL = StrSQL & " FROM dbo.TblExpensesDet301 AS E "

' «” Œœ«„ OUTER APPLY ·Ã·» ﬁÌ„… Ê«Õœ… ›ﬁÿ „‰ «·ﬁÌÊœ ·„‰⁄ «· ﬂ—«—
StrSQL = StrSQL & " OUTER APPLY (SELECT TOP 1 PriceTotal FROM dbo.DOUBLE_ENTREY_VOUCHERS WHERE notes_all = E.ExpID) AS VoucherData "

StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblProcessDEF AS ProcDef ON E.operid = ProcDef.TblProcessDEFID "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.projects_des AS PD ON E.pandid = PD.oprid "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.projects AS Proj ON E.projectid = Proj.id "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.ACCOUNTS AS Acc ON E.AccountCode = Acc.Account_Code "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblAqarDetai AS AD ON E.UnitNo = AD.Id "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblAkarUnit AS AU ON E.UnitType = AU.id "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblAqar AS Aqar ON E.Aqarid = Aqar.Aqarid "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblEmployee AS Emp ON E.NEmpid = Emp.Emp_ID "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblEmpDepartments AS D ON E.Departementid = D.DeparmentID "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblBranchesData AS B ON E.BrnchID = B.branch_id "

StrSQL = StrSQL & " WHERE (E.ExpID = " & val(XPTxtID.text) & ")"


        Set RsDev = New ADODB.Recordset
         RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsDev.RecordCount > 0 Then
            RsDev.MoveFirst
        End If
    
        With Me.VSFlexGrid1
 
            .rows = .FixedRows + RsDev.RecordCount
 
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(RsDev("Project_name").value), "", RsDev("Project_name").value)
                .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(RsDev("Panddes").value), "", RsDev("Panddes").value)
                .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                .TextMatrix(i, .ColIndex("projectid")) = IIf(IsNull(RsDev("projectid").value), "", RsDev("projectid").value)
                .TextMatrix(i, .ColIndex("operid")) = IIf(IsNull(RsDev("operid").value), "", RsDev("operid").value)
                .TextMatrix(i, .ColIndex("pandid")) = IIf(IsNull(RsDev("pandid").value), "", RsDev("pandid").value)
                .TextMatrix(i, .ColIndex("ProjectCode")) = IIf(IsNull(RsDev("ProjectFullcode").value), "", RsDev("ProjectFullcode").value)
                .TextMatrix(i, .ColIndex("StrUnit")) = IIf(IsNull(RsDev("StrUnit").value), "", RsDev("StrUnit").value)
                .TextMatrix(i, .ColIndex("Unitss")) = IIf(IsNull(RsDev("Unitss").value), "", RsDev("Unitss").value)
                .TextMatrix(i, .ColIndex("BrnchID")) = IIf(IsNull(RsDev("BrnchID").value), "", RsDev("BrnchID").value)
                .TextMatrix(i, .ColIndex("FlgVat")) = IIf(IsNull(RsDev("FlgVat").value), 0, RsDev("FlgVat").value)
                .TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(RsDev("Vatyo").value), 0, RsDev("Vatyo").value)
                .TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(RsDev("Vat").value), 0, RsDev("Vat").value)
                .TextMatrix(i, .ColIndex("CurrRow")) = IIf(IsNull(RsDev("CurrRow").value), 0, RsDev("CurrRow").value)
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
                .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
                .TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(RsDev("Aqarid").value), "", RsDev("Aqarid").value)
                .TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(RsDev("aqarname").value), "", RsDev("aqarname").value)
                .TextMatrix(i, .ColIndex("UnitType")) = IIf(IsNull(RsDev("UnitType").value), "", RsDev("UnitType").value)
                .TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(RsDev("UnitNo").value), "", RsDev("UnitNo").value)
                .TextMatrix(i, .ColIndex("unitnoName")) = IIf(IsNull(RsDev("UnitNoName").value), "", RsDev("UnitNoName").value)
                .TextMatrix(i, .ColIndex("billno")) = IIf(IsNull(RsDev("Billno").value), "", RsDev("Billno").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("NEmpName")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Departement")) = IIf(IsNull(RsDev("DepartmentName").value), "", RsDev("DepartmentName").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDev("branch_name").value), "", RsDev("branch_name").value)
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                Else
                .TextMatrix(i, .ColIndex("NEmpName")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Departement")) = IIf(IsNull(RsDev("DepartmentNamee").value), "", RsDev("DepartmentNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDev("branch_namee").value), "", RsDev("branch_namee").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                End If
                .TextMatrix(i, .ColIndex("PriceTotal")) = IIf(IsNull(RsDev("PriceTotal").value), 0, RsDev("PriceTotal").value)
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(RsDev("Des").value), "", RsDev("Des").value)
                .TextMatrix(i, .ColIndex("NEmpid")) = IIf(IsNull(RsDev("NEmpid").value), "", RsDev("NEmpid").value)
                .TextMatrix(i, .ColIndex("Departementid")) = IIf(IsNull(RsDev("Departementid").value), "", RsDev("Departementid").value)
                RsDev.MoveNext
            Next i
   ' .AutoSize 1, .Cols - 1, False
        End With

        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        ReLineGrid
       fillapprovData
        
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then '«·„—Ê›« 
        '   StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
        '   StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        ' StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"

        'StrSQL = "SELECT   dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
        'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.notes_all =" & Val(Me.XPTxtID.text) & ") "
        'StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
          StrSQL = "SELECT     TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, "
          StrSQL = StrSQL + "            dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name,"
          StrSQL = StrSQL + "            dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
          StrSQL = StrSQL + "            dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID,"
          StrSQL = StrSQL + "            dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.ORDER_NO,"
          StrSQL = StrSQL + "            dbo.Notes.ProjectID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.Notes.Pand, dbo.projects_des.des, dbo.Notes.Oper,"
          StrSQL = StrSQL + "            dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE, dbo.Notes.fixedid, dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.Fullcode,"
          StrSQL = StrSQL + "            dbo.FixedAssets.namee, dbo.ACCOUNTS.Account_Serial, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, dbo.TblBranchesData.branch_name,"
          StrSQL = StrSQL + "            dbo.TblBranchesData.branch_namee ,dbo.projects.Fullcode as ProjectCode ,dbo.DOUBLE_ENTREY_VOUCHERS.FlgVat,dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo,dbo.DOUBLE_ENTREY_VOUCHERS.Vat ,dbo.DOUBLE_ENTREY_VOUCHERS.CurrRow,dbo.DOUBLE_ENTREY_VOUCHERS.PriceTotal"
          StrSQL = StrSQL + "     FROM         dbo.ACCOUNTS INNER JOIN"
          StrSQL = StrSQL + "            dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
          StrSQL = StrSQL + "            dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
          StrSQL = StrSQL + "            dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
          StrSQL = StrSQL + "            dbo.FixedAssets ON dbo.Notes.fixedid = dbo.FixedAssets.id LEFT OUTER JOIN"
          StrSQL = StrSQL + "            dbo.TblProcessDEF ON dbo.Notes.Oper = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
          StrSQL = StrSQL + "            dbo.projects_des ON dbo.Notes.Pand = dbo.projects_des.oprid and dbo.projects_des.oprid<>0  LEFT OUTER JOIN"
          StrSQL = StrSQL + "            dbo.projects ON dbo.Notes.ProjectID = dbo.projects.id"
          StrSQL = StrSQL + " Where ( hideline is null and dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.text) & ")"
          StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
    
        Set RsDev = New ADODB.Recordset
         RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

  ' StrSQL = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* FROM         dbo.DOUBLE_ENTREY_VOUCHERS WHERE     (Double_Entry_Vouchers_ID = - 1)"
  ' RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i
    
            RsDev.MoveFirst
    
            With Me.Fg_Journal

                If Me.dcproject.BoundText = "" Then
                    .rows = .FixedRows + RsDev.RecordCount
                Else
                    .rows = .FixedRows + RsDev.RecordCount - 1
                End If

                For i = .FixedRows To .rows - 1
                    .TextMatrix(i, .ColIndex("ProjectCode")) = IIf(IsNull(RsDev("ProjectCode").value), "", RsDev("ProjectCode").value)
                     .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
                    .TextMatrix(i, .ColIndex("BrnchID")) = IIf(IsNull(RsDev("branch_id").value), "", RsDev("branch_id").value)
                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
                    .TextMatrix(i, .ColIndex("ExpensesID")) = get_Expenses_id(IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value))
                    .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
                    .TextMatrix(i, .ColIndex("FlgVat")) = IIf(IsNull(RsDev("FlgVat").value), 0, RsDev("FlgVat").value)
                    .TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(RsDev("Vatyo").value), 0, RsDev("Vatyo").value)
                    .TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(RsDev("Vat").value), 0, RsDev("Vat").value)
                    .TextMatrix(i, .ColIndex("CurrRow")) = IIf(IsNull(RsDev("CurrRow").value), 0, RsDev("CurrRow").value)
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                    .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(RsDev("Account_Serial").value), "", RsDev("Account_Serial").value)
                    .TextMatrix(i, .ColIndex("PriceTotal")) = IIf(IsNull(RsDev("PriceTotal").value), 0, RsDev("PriceTotal").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDev("branch_name").value), "", RsDev("branch_name").value)
                        .TextMatrix(i, .ColIndex("Fixes")) = IIf(IsNull(RsDev("Name").value), "", RsDev("Name").value)
                        .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(RsDev("Project_name").value), "", RsDev("Project_name").value)
                        .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                    Else
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDev("branch_namee").value), "", RsDev("branch_namee").value)
                    .TextMatrix(i, .ColIndex("Fixes")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                         .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(RsDev("Project_nameE").value), "", RsDev("Project_nameE").value)
                         .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessNameE").value), "", RsDev("ProcessNameE").value)
                         .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                    End If
                    .TextMatrix(i, .ColIndex("fixedid")) = IIf(IsNull(RsDev("fixedid").value), "", RsDev("fixedid").value)
                    
                    .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)

                    'Double_Entry_Vouchers_Description
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
            
                    '    .TextMatrix(I, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), _
                    '        "", RsDev("Account_Name").value)
        
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
            ''//Notes
                    .TextMatrix(i, .ColIndex("projectid2")) = IIf(IsNull(RsDev("ProjectID").value), "", RsDev("ProjectID").value)
                    .TextMatrix(i, .ColIndex("pandid2")) = IIf(IsNull(RsDev("Pand").value), "", RsDev("Pand").value)
                    .TextMatrix(i, .ColIndex("operid2")) = IIf(IsNull(RsDev("Oper").value), "", RsDev("Oper").value)
                    
            ''/
            
                    .TextMatrix(i, .ColIndex("Order_No")) = IIf(IsNull(RsDev("Order_No").value), "", RsDev("Order_No").value)
 
                    RsDev.MoveNext
                Next i

                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
                ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
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
    ReLineGrid
    fillapprovData
    Exit Sub
ErrTrap:
End Sub
Sub SaveUnitNo(Optional ID As Long, Optional i As Integer)
   Dim RsDetails As ADODB.Recordset
   Dim astrSplit2tems2() As String
   Dim astrSplitItems() As String
   Dim sql As String
   Dim j As Integer
    Dim st As String
    Dim nElements As Integer
    
      If VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("StrUnit")) <> "" Then
          st = VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("StrUnit"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         sql = "Select * from TblExp301UnitNo where 1=-1"
         Set RsDetails = New ADODB.Recordset
         RsDetails.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
         For j = 0 To nElements - 1
          RsDetails.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         RsDetails("ExpID").value = val(XPTxtID.text)
         RsDetails("ExpDetails").value = ID
         RsDetails("UnitID").value = val(astrSplit2tems2(1))
         RsDetails("Valu").value = val(astrSplit2tems2(2))
         RsDetails.update
         Next j
          End If
End Sub
Private Sub SaveData()
    Dim Msg As String
    Dim brnchid As Integer
    Dim total_value As Double
    Dim BranchID As Integer
    Dim BranchID2 As Integer
    Dim DeptSide As String
    Dim credit_side As String
    Dim OtherInformation As New ClsGLOther
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim Priod As Integer
    Dim Posted As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim sql As String
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim j As Integer
    Dim st As String
    Dim des As String
    Dim nElements As Integer
    
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then


   If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccountTimes(Me.DcboBox.BoundText, XPDtbTrans.value, Priod) = False Then
                    MsgBox " „ «‰ Â«¡ «·„œ… «·„Õœœ… ··«” ⁄«÷… " & " »„œ… " & Priod & "  ÌÊ„ ", vbCritical
                        Exit Sub
                    End If
                End If
                
           If Me.CboPayMentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
                Msg = "Select Payment method ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Exit Sub
        End If
    
        If Me.CboPayMentType.ListIndex = 2 Then
            If Trim(Me.DCVendor.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Msg = "ÌÃ» ≈Œ Ì«— «·„Ê—œ..!!"
                Else
                    'Msg = "Select vendor..!!"
                End If

              '  MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              '  DCVendor.SetFocus
              '  SendKeys "{F4}"
              '  Exit Sub
            End If

        End If
    
        If Me.CboPayMentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·⁄Âœ…..!!"
                Else
                    Msg = "Select Box..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                'SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Or Me.CboPayMentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" And Me.CboPayMentType.ListIndex <> 2 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
        ElseIf Me.CboPayMentType.ListIndex = 4 Then
        If DcbAccount.BoundText = "" Or DcbAccount.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«»"
        Else
        MsgBox "Please Select Account"
        End If
        DcbAccount.SetFocus
        Exit Sub
        End If
        ElseIf Me.CboPayMentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ   «·»‰ﬂ...!!"
             Else
             Msg = " Specify Bank    ...!!"
             End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
               Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ —ﬁ„ «·ÕÊ«·Â...!!"
             Else
             Msg = " Define Transfer No#    ...!!"
             End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
            End If

            '     If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '                 If SystemOptions.UserInterface = ArabicInterface Then
            '                     Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '                 Else
            '                 Msg = "Cheque Due Date Not Valid...!!"
            '
            '                 End If
            '         MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '         DtpChequeDueDate.SetFocus
            '         SendKeys "{F4}"
            '         Exit Sub
            '     End If
        End If
                
        If Me.CboPaymentType1.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·›« Ê—… ...!!!"
            Else
                Msg = "Select Bill Type ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Exit Sub
        End If
    
        If Me.CboPayMentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
                Msg = "Select Payment method ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Exit Sub
        End If
    
        If Me.CboPayMentType.ListIndex = 2 Then
            If Trim(Me.DCVendor.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·„Ê—œ..!!"
                Else
                    Msg = "Select vendor..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCVendor.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        
        End If
    
        If Me.CboPayMentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·⁄Âœ…..!!"
                Else
                    Msg = "Select Box..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                'SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If

            '     If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '                 If SystemOptions.UserInterface = ArabicInterface Then
            '                     Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '                 Else
            '                 Msg = "Cheque Due Date Not Valid...!!"
            '
            '                 End If
            '         MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '         DtpChequeDueDate.SetFocus
            '         SendKeys "{F4}"
            '         Exit Sub
            '     End If
        End If
    
        Dim i As Integer

        If CboPaymentType1.ListIndex = 0 Then

            With Fg_Journal

                For i = .FixedRows To .rows - 2

                    If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·« ÌÊÃœ Õ”«» ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Select Expenses in line no" & i, vbCritical
                        End If

                        Exit Sub
              
                    End If
        
                Next i

            End With

            With Fg_Journal

                For i = .FixedRows To .rows - 2

                    If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«Ì ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Enter Value in line no" & i, vbCritical
                        End If
               
                        Exit Sub
                    End If
        
                Next i

            End With

        End If

        'Õ”«»« 
        If Me.CboPaymentType1.ListIndex = 1 Then
      
            With Me.VSFlexGrid1

                For i = .FixedRows To .rows - 2

                    If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«  ÌÊÃœ Õ”«» ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Select Expenses in line no" & i, vbCritical
                        End If

                        Exit Sub
              
                    End If
        
                Next i

            End With
   
            With Me.VSFlexGrid1

                For i = .FixedRows To .rows - 2

                    If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«  ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Enter Value in line no" & i, vbCritical
                        End If
               
                        Exit Sub
                    End If
        
                Next i

            End With
 
        End If
      Dim ISVAT As Boolean
    ISVAT = False
With Fg_Journal
    For i = .FixedRows To .rows - 1
      If val(.TextMatrix(i, .ColIndex("Vat"))) > 0 Then
      ISVAT = True
      End If
     Next i
 End With
 With VSFlexGrid1
    For i = .FixedRows To .rows - 1
      If val(.TextMatrix(i, .ColIndex("Vat"))) > 0 Then
      ISVAT = True
      End If
     Next i
 End With
 
Dim AccountVATDept As String
If ISVAT = True And True = True Then
If GetValueAddedAccount(XPDtbTrans.value, AccountVATDept) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If

        If Me.TxtModFlg.text = "N" Then
            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value) = False Then
                        Exit Sub
                    End If
                End If
            End If

        ElseIf Me.TxtModFlg.text = "E" Then

            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value, , , val(Me.XPTxtID.text)) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    
        Dim xrow As Integer

        With Fg_Journal

            For xrow = .rows - 1 To 2 Step -1

                If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                    .rows = .rows - 1
                End If

            Next xrow

        End With
    
        With Me.VSFlexGrid1

            For xrow = .rows - 1 To 2 Step -1

                If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                    .rows = .rows - 1
                End If

            Next xrow

        End With

        calcnets

        '-------------------------------------------------------------------------------------------
 
        '-------------------------------------------------------------------------------------------
        If TxtSerial1.text = "" Then
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 35, 350) = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ”ÊÌ… ⁄Âœ… ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                    MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                End If

            Else
         
                If Voucher_coding(val(my_branch), XPDtbTrans.value, 35, 350) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                    End If

                Else
                    TxtSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 35, 350)
                End If
            End If
        End If
    
        If TxtSerial.text = "" Then
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                    MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
                End If

            Else
         
                If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        MsgBox "You must Define JE Coding ": Exit Sub
                    End If

                Else
                    TxtSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
                End If
            End If
        End If
    
        Cn.BeginTrans
        BeginTrans = True
    
        '///////////////NOTESALL
        Dim A_NoteID As Long

        If TxtModFlg.text = "N" Then
            XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=350"))
            rs.AddNew
   
            Me.oldTxtSerial1.text = Trim$(Me.TxtSerial1.text)
 
        ElseIf Me.TxtModFlg.text = "E" Then
    
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where notes_all=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             Cn.Execute " Delete from TblExp301UnitNo where  ExpID =" & val(XPTxtID.text)
             Cn.Execute " Delete from TblExpensesDet301 where  ExpID =" & val(XPTxtID.text)
            StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
       
            If DcCostCenter.BoundText <> "" Then
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
        
             StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
        End If
    
        '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("branch_no").value = val(Me.dcBranch.BoundText)
        rs("NoteID").value = val(XPTxtID.text)
        rs("bill_Type").value = Me.CboPaymentType1.ListIndex
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("foxy_no").value = val(Text1.text)
        rs("order_no").value = TXT_order_no.text
        rs("AccountPaym").value = IIf(Trim(DcbAccount.BoundText) = "", Null, DcbAccount.BoundText)
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
        rs("note_value_by_characters").value = WriteNo(Format(val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0)
        
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("too").value = IIf(txtto.text = "", "", Trim(txtto.text))
        rs("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
      
        If CBoBasedON.ListIndex > -1 Then
            rs("BasedONID").value = CBoBasedON.ListIndex
        Else
            rs("BasedONID").value = 0
        End If
    
        rs("CusID").value = Null
        rs("NoteType").value = 350
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteHijriDate").value = ToHijriDate(XPDtbTrans.value)
        rs("UserID").value = user_id
        rs("ExpensesID").value = IIf(XPCboExpensesType.text = "", Null, XPCboExpensesType.BoundText)
  '-------------------5555555555
        If Me.CboPayMentType.ListIndex = 0 Then
            rs("BoxID").value = val(DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = 0
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 1
    
        ElseIf Me.CboPayMentType.ListIndex = 3 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 3
        
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            rs("NoteCashingType").value = 2
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
        ElseIf Me.CboPayMentType.ListIndex = 4 Then
              rs("NoteCashingType").value = 4
        End If
   '-------------5555555555
    
        rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
        rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
        rs("Buy").value = "0"
        rs("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
        rs("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
        rs("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
 
        rs("OldNoteSerial1").value = Trim$(Me.oldTxtSerial1.text) '
     
        rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(35) '‰Ê⁄  —ﬁÌ„  ‰’›Ì… ⁄Âœ…
     
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
'rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)

        If dcproject.BoundText <> "" Then
        '    rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
        Else
       '     rs("note_value_by_characters").value = WriteNo(Format(val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0)
        End If

        If Me.TxtModFlg.text = "N" Then
            A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
            TXT_A_NoteID.text = A_NoteID
        Else
            A_NoteID = val(TXT_A_NoteID.text)
        End If
    
        rs("A_NoteID").value = val(A_NoteID)
     
        rs.update
    
        '/////////////////////Õ”«»«  ⁄«„Â
        Dim line_no  As Integer

        If Me.CboPaymentType1.ListIndex = 1 Then
      
            Set RsNotes = New ADODB.Recordset
       '     RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If TxtModFlg.text = "N" Then
           
            ElseIf Me.TxtModFlg.text = "E" Then
           '   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.Text)
           '     Cn.Execute StrSQL, , adExecuteNoRecords
        
            End If
    
            '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
            ' rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
            ' rs("foxy_no").value = Val(Text1.text)

            'Õ”«»« 
            RsNotes.AddNew
            RsNotes("NoteID").value = A_NoteID
             RsNotes.update
            RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
            RsNotes("order_no").value = TXT_order_no.text
            RsNotes("notes_all").value = Me.XPTxtID.text
            RsNotes("Note_Value").value = IIf(Not IsNumeric(XPTxtVal.text), 0, val(XPTxtVal.text))
            'RsNotes("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
            RsNotes("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
            RsNotes("too").value = IIf(txtto.text = "", "", Trim(txtto.text))
            '    RsNotes("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
 '-------------55555555555
            If Me.CboPayMentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPayMentType.ListIndex = 1 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 1
        
            ElseIf Me.CboPayMentType.ListIndex = 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 3
           
            ElseIf Me.CboPayMentType.ListIndex = 2 Then
                       rs("NoteCashingType").value = 2
                    rs("BoxID").value = Null
                    rs("BankID").value = val(Me.DcboBankName.BoundText)
                    rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                    rs("DueDate").value = Me.DtpChequeDueDate.value
                    rs("ChequeBoxID").value = Null
            End If
    '--------------------555555555
            RsNotes("NoteType").value = 350
            RsNotes("NoteDate").value = XPDtbTrans.value
            RsNotes("UserID").value = user_id
    
            'rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
            'rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
            RsNotes("Buy").value = "0"
            ' RsNotes("Remark").value = XPMTxtRemarks.text
            RsNotes("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
            RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
            RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
            RsNotes("numbering_type1").value = sand_numbering_type(35) '‰Ê⁄  —ﬁÌ„      ‰’›Ì… ⁄Âœ…
     
            RsNotes("sanad_year").value = year(XPDtbTrans.value)
            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
            RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
            RsNotes.update
             sql = "Select * from TblExpensesDet301 where 1=-1"
             rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
            '„œÌ‰ Õ”«»« 
            Dim LineNo1 As Double
            With VSFlexGrid1
                line_no = 1
 LineNo1 = 0
                For i = .FixedRows To .rows - 1
    
                    Dim project_id As Integer
    
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
       
                       '  project_id = get_project_id(dcproject.BoundText, "expanses_account")
                 
                    '''''
                    rs2.AddNew
                    
                    rs2("ExpID").value = val(XPTxtID.text)
                    rs2("projectid").value = IIf(IsNull(.TextMatrix(i, .ColIndex("projectid")) = ""), 0, val(.TextMatrix(i, .ColIndex("projectid"))))
                    rs2("operid").value = IIf(IsNull(.TextMatrix(i, .ColIndex("operid")) = ""), 0, val(.TextMatrix(i, .ColIndex("operid"))))
                    rs2("pandid").value = IIf(IsNull(.TextMatrix(i, .ColIndex("pandid")) = ""), 0, val(.TextMatrix(i, .ColIndex("pandid"))))
                    
                    rs2("FlgVat").value = IIf(IsNull(.TextMatrix(i, .ColIndex("FlgVat")) = ""), 0, val(.TextMatrix(i, .ColIndex("FlgVat"))))
                    rs2("ForcedFlg").value = IIf(IsNull(.TextMatrix(i, .ColIndex("ForcedFlg")) = ""), 0, val(.TextMatrix(i, .ColIndex("ForcedFlg"))))
                    rs2("CurrRow").value = IIf(IsNull(.TextMatrix(i, .ColIndex("CurrRow")) = ""), 0, val(.TextMatrix(i, .ColIndex("CurrRow"))))
                    rs2("BrnchID").value = IIf(IsNull(.TextMatrix(i, .ColIndex("BrnchID")) = ""), 0, val(.TextMatrix(i, .ColIndex("BrnchID"))))
                    rs2("Value").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Value")) = ""), 0, val(.TextMatrix(i, .ColIndex("Value"))))
                    rs2("Vatyo").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Vatyo")) = ""), 0, val(.TextMatrix(i, .ColIndex("Vatyo"))))
                    rs2("Vat").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Vat")) = ""), 0, val(.TextMatrix(i, .ColIndex("Vat"))))
                    rs2("Des").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Des")) = ""), 0, .TextMatrix(i, .ColIndex("Des")))
                    rs2("Departementid").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Departementid")) = ""), 0, val(.TextMatrix(i, .ColIndex("Departementid"))))
                    rs2("NEmpid").value = IIf(IsNull(.TextMatrix(i, .ColIndex("NEmpid")) = ""), 0, val(.TextMatrix(i, .ColIndex("NEmpid"))))
                    rs2("Aqarid").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Aqarid")) = ""), 0, val(.TextMatrix(i, .ColIndex("Aqarid"))))
                    rs2("UnitType").value = IIf(IsNull(.TextMatrix(i, .ColIndex("UnitType")) = ""), 0, val(.TextMatrix(i, .ColIndex("UnitType"))))
                    rs2("UnitNo").value = IIf(IsNull(.TextMatrix(i, .ColIndex("UnitNo")) = ""), 0, val(.TextMatrix(i, .ColIndex("UnitNo"))))
                    rs2("billno").value = IIf(IsNull(.TextMatrix(i, .ColIndex("billno")) = ""), 0, .TextMatrix(i, .ColIndex("billno")))
                    rs2("Unitss").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Unitss")) = ""), 0, .TextMatrix(i, .ColIndex("Unitss")))
                    rs2("StrUnit").value = IIf(IsNull(.TextMatrix(i, .ColIndex("StrUnit")) = ""), 0, .TextMatrix(i, .ColIndex("StrUnit")))
                    rs2("AccountCode").value = IIf(IsNull(.TextMatrix(i, .ColIndex("AccountCode")) = ""), 0, .TextMatrix(i, .ColIndex("AccountCode")))
                    rs2.update
                    SaveUnitNo rs2("id").value, i
                   brnchid = val(.TextMatrix(i, .ColIndex("BrnchID")))
                    OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                    OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                    OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                    OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                    OtherInformation.UnitString = .TextMatrix(i, .ColIndex("StrUnit"))
                    OtherInformation.Unitss = .TextMatrix(i, .ColIndex("Unitss"))
                    OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
      If .TextMatrix(i, .ColIndex("StrUnit")) <> "" Then
          st = .TextMatrix(i, .ColIndex("StrUnit"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         des = ""
         des = .TextMatrix(i, .ColIndex("Des"))
         des = des & " "
         des = des & .TextMatrix(i, .ColIndex("aqarname")) & "\ "
         des = des & " "
         des = des & .TextMatrix(i, .ColIndex("name")) & "\ "
         des = des & astrSplit2tems2(0) & " "
         
         
         If val(astrSplit2tems2(2)) <> 0 Then
             
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
Dim Material_account As String
LineNo1 = LineNo1 + 1

                      If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(astrSplit2tems2(2)), 0, des, A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , LineNo1, val(Me.XPTxtID.text), project_id, , , , val(.TextMatrix(i, .ColIndex("FixedassetId"))), , , brnchid, , , , , , , val(.TextMatrix(i, .ColIndex("Departementid"))), val(.TextMatrix(i, .ColIndex("NEmpid"))), , val(.TextMatrix(i, .ColIndex("Aqarid"))), val(.TextMatrix(i, .ColIndex("UnitType"))), val(astrSplit2tems2(1)), .TextMatrix(i, .ColIndex("billno")), val(.TextMatrix(i, .ColIndex("projectid"))), val(.TextMatrix(i, .ColIndex("pandid"))), val(.TextMatrix(i, .ColIndex("operid"))), , , , , , , , , Posted, , OtherInformation) = False Then
                        
                            GoTo ErrTrap
                    
                         End If
LineNo1 = LineNo1 + 1

'*****************************************************Ã«—Ì
                   BranchID = val(Me.dcBranch.BoundText)
            
                BranchID2 = brnchid

                                  DeptSide = getBranchCurrentAccount(BranchID)
                                                 credit_side = getBranchCurrentAccount(BranchID2)
                                      
       If BranchID <> BranchID2 Then
 total_value = val(astrSplit2tems2(2))
line_no = line_no + 1
                                               If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, total_value, 0, des, A_NoteID, , , , XPDtbTrans.value, user_id, , , , , , , , , LineNo1, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , 1, , , , , , , Posted) = False Then
                                                                   
                                                              End If
                                                              
                                                             line_no = line_no + 1
                                                             LineNo1 = LineNo1 + 1
                                                        '????
                                                              If ModAccounts.AddNewDev(LngDevID, line_no, DeptSide, total_value, 1, des, A_NoteID, , , , XPDtbTrans.value, user_id, , , , , , , , , LineNo1, , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , 1, , , , , , , Posted) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                line_no = line_no + 1
        

       End If
       '*****************************************************Ã«—Ì
           End If
           
         Next j
     Else

    
         If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
             LineNo1 = LineNo1 + 1
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                      If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(.TextMatrix(i, .ColIndex("Value"))), 0, .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , LineNo1, val(Me.XPTxtID.text), project_id, , , , val(.TextMatrix(i, .ColIndex("FixedassetId"))), , , brnchid, , , , , , , val(.TextMatrix(i, .ColIndex("Departementid"))), val(.TextMatrix(i, .ColIndex("NEmpid"))), , val(.TextMatrix(i, .ColIndex("Aqarid"))), val(.TextMatrix(i, .ColIndex("UnitType"))), val(.TextMatrix(i, .ColIndex("UnitNo"))), .TextMatrix(i, .ColIndex("billno")), val(.TextMatrix(i, .ColIndex("projectid"))), val(.TextMatrix(i, .ColIndex("pandid"))), val(.TextMatrix(i, .ColIndex("operid"))), , , , , , , , , Posted, , OtherInformation) = False Then
                        
                            GoTo ErrTrap
                    
                      End If
                      
'*****************************************************Ã«—Ì
                   BranchID = val(Me.dcBranch.BoundText)
            
                BranchID2 = brnchid

                                  DeptSide = getBranchCurrentAccount(BranchID)
                                  credit_side = getBranchCurrentAccount(BranchID2)
                                      
       If BranchID <> BranchID2 Then
       LineNo1 = LineNo1 + 1
 total_value = val(.TextMatrix(i, .ColIndex("Value")))
line_no = line_no + 1
                                               If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, total_value, 0, .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , , XPDtbTrans.value, user_id, , , , , , , , , LineNo1, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , 1, , , , , , , Posted) = False Then
                                                                   
                                                              End If
                                                              
                                                             line_no = line_no + 1
                                                             LineNo1 = LineNo1 + 1
                                                        '????
                                                              If ModAccounts.AddNewDev(LngDevID, line_no, DeptSide, total_value, 1, .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , , XPDtbTrans.value, user_id, , , , , , , , , LineNo1, , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , 1, , , , , , , Posted) = False Then
                                                                   
                                                              End If
                                                                       
                                                   line_no = line_no + 1
        

       End If
       '*****************************************************Ã«—Ì
         End If
         End If
'  If project_id <> 0 And SystemOptions.gldetails_or_gl_general = 1 Then
'                        line_no = line_no + 1
'                        OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
'                        OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
'                        OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
'                        OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
'                              Material_account = get_project_Account(project_id, "expanses_account")
'                              If ModAccounts.AddNewDev(LngDevID, line_no, Material_account, .TextMatrix(i, .ColIndex("Value")), 0, .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , val(.TextMatrix(i, .ColIndex("LineNo1"))), val(Me.XPTxtID.Text), , , , , val(.TextMatrix(i, .ColIndex("FixedassetId"))), , , brnchid, , , , , , , val(.TextMatrix(i, .ColIndex("Departementid"))), val(.TextMatrix(i, .ColIndex("NEmpid"))), , val(.TextMatrix(i, .ColIndex("Aqarid"))), val(.TextMatrix(i, .ColIndex("UnitType"))), val(.TextMatrix(i, .ColIndex("UnitNo"))), .TextMatrix(i, .ColIndex("billno")), val(.TextMatrix(i, .ColIndex("projectid"))), val(.TextMatrix(i, .ColIndex("pandid"))), val(.TextMatrix(i, .ColIndex("operid"))), , 1, , , , , , , Posted, , OtherInformation) = False Then
'
'                            GoTo ErrTrap
'
'                        End If
'
'
'                        line_no = line_no + 1
'
'
'
'                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("Value")), 1, .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , val(.TextMatrix(i, .ColIndex("LineNo1"))), val(Me.XPTxtID.Text), , , , , val(.TextMatrix(i, .ColIndex("FixedassetId"))), , , brnchid, , , , , , , val(.TextMatrix(i, .ColIndex("Departementid"))), val(.TextMatrix(i, .ColIndex("NEmpid"))), , val(.TextMatrix(i, .ColIndex("Aqarid"))), val(.TextMatrix(i, .ColIndex("UnitType"))), val(.TextMatrix(i, .ColIndex("UnitNo"))), .TextMatrix(i, .ColIndex("billno")), val(.TextMatrix(i, .ColIndex("projectid"))), val(.TextMatrix(i, .ColIndex("pandid"))), val(.TextMatrix(i, .ColIndex("operid"))), , 1, , , , , , , Posted, , OtherInformation) = False Then
'
'                            GoTo ErrTrap
'
'                        End If
'
'  End If
            
                    End If

                Next i

            End With

            'œ«∆‰ Õ”«»« 
    
            Dim IntDEV_Type As Integer
            Dim SngDEV_Value As Single
            line_no = line_no + 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            OtherInformation.FlgVat = 0
            OtherInformation.Vat = 0
            OtherInformation.Vatyo = 0
            OtherInformation.CurrRow = 0
                       
            If ModAccounts.AddNewDev(LngDevID, line_no, DcboCreditSide.BoundText, IIf(Not IsNumeric(XPTxtVal.text), 0, val(XPTxtVal.text)), 1, txt_general_des.text, A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.text), , , , , , , , brnchid, , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                    
            End If
        
            ' TxtModFlg.text = "R"
            GoTo ll
      
        End If
    
        '  «·„’—Ê›«  „œÌ‰
    
        '//////////////////////////////////////Notes////////////////////////////////////
        Set RsNotes = New ADODB.Recordset
     '   RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
       
            Set RsDev = New ADODB.Recordset
      '      RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     
    StrSQL = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* FROM         dbo.DOUBLE_ENTREY_VOUCHERS WHERE     (Double_Entry_Vouchers_ID = - 1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
       
            '«·ÿ—› «·„œÌ‰
 
            Dim ExpensesID As Double
 
            Dim NoteID As String

            With Fg_Journal

                line_no = 1
       
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                
                For i = .FixedRows To .rows - 1
   
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                        '////////////////////////////////////////notes
                
                        If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                            Else
                                MsgBox "Cant save no value in line no:  " & i - 1, vbCritical: GoTo ErrTrap
                            End If
               
                        End If

                        RsNotes.AddNew
                        NoteID = CStr(new_id("Notes", "NoteID", "", True))
                        RsNotes("NoteID").value = CStr(NoteID)
                         RsNotes.update
                        RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
                        RsNotes("Note_Value").value = .TextMatrix(i, .ColIndex("value"))
                        RsNotes("ExpensesRemark").value = .TextMatrix(i, .ColIndex("des"))
                        RsNotes("ProjectID").value = val(.TextMatrix(i, .ColIndex("projectid2")))
                        RsNotes("Pand").value = val(.TextMatrix(i, .ColIndex("pandid2")))
                        RsNotes("Oper").value = val(.TextMatrix(i, .ColIndex("operid2")))

                        '  RsNotes("Remark").value = .TextMatrix(I, .ColIndex("des"))
                
                        RsNotes("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))

                        RsNotes("foxy_no").value = val(Text1.text)
                        '  If Me.CboPayMentType.ListIndex = 0 Then
                        '     rsnotes("BoxID").value = Val(DcboBox.BoundText)
                        '     rsnotes("BankID").value = Null
                        '     rsnotes("ChqueNum").value = Null
                        '     rsnotes("DueDate").value = Null
                        '     rsnotes("NoteCashingType").value = 0
                        ' ElseIf Me.CboPayMentType.ListIndex = 1 Then
                        '     rsnotes("BoxID").value = Null
                        '     rsnotes("BankID").value = Val(Me.DCboBankName.BoundText)
                        '     rsnotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                        '     rsnotes("DueDate").value = Me.DtpChequeDueDate.value
                        '     rsnotes("NoteCashingType").value = 1
                        ' End If
               
                        If TXT_order_no.text <> "" Then
                            RsNotes("order_no").value = TXT_order_no.text
                        Else
                            RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                        End If
            
                        RsNotes("CusID").value = Null
                        RsNotes("NoteType").value = 350
                        RsNotes("NoteDate").value = XPDtbTrans.value
                        RsNotes("UserID").value = user_id
                        RsNotes("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
                        RsNotes("fixedid").value = val(.TextMatrix(i, .ColIndex("fixedid")))
                        RsNotes("notes_all").value = Me.XPTxtID.text
                        RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
                        RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
                        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                        RsNotes("numbering_type1").value = sand_numbering_type(35) '‰Ê⁄  —ﬁÌ„    ’›Ì… ⁄Âœ…
                
                        RsNotes("sanad_year").value = year(XPDtbTrans.value)
                        RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                        RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
                
                        RsNotes.update
                        brnchid = .TextMatrix(i, .ColIndex("BrnchID"))
               project_id = val(Me.dcproject.BoundText)
                If project_id = 0 Then
project_id = val(.TextMatrix(i, .ColIndex("projectid2")))
End If

                        '////////////////////////////////////////notes
 
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                        OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                        OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                        OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                        OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(.TextMatrix(i, .ColIndex("value"))), 0, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , val(.TextMatrix(i, .ColIndex("fixedid"))), , , brnchid, , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                       
                            '   GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
                        
                        
                        
                        '*****************************************************Ã«—Ì
                   BranchID = val(Me.dcBranch.BoundText)
            
                BranchID2 = brnchid

                                  DeptSide = getBranchCurrentAccount(BranchID)
                                                 credit_side = getBranchCurrentAccount(BranchID2)
                                      
       If BranchID <> BranchID2 Then
 total_value = val(.TextMatrix(i, .ColIndex("Value")))
line_no = line_no + 1
                                               If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, total_value, 0, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , , XPDtbTrans.value, user_id, , , , , , , , , val(.TextMatrix(i, .ColIndex("LineNo1"))), val(Me.XPTxtID.text), , , , , , , , BranchID, , , , , , , , , , , , , , , , , , 1, , , , , , , Posted) = False Then
                                                                   
                                                              End If
                                                              
                                                             line_no = line_no + 1
                                                        '????
                                                              If ModAccounts.AddNewDev(LngDevID, line_no, DeptSide, total_value, 1, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , , XPDtbTrans.value, user_id, , , , , , , , , val(.TextMatrix(i, .ColIndex("LineNo1"))), val(Me.XPTxtID.text), , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , 1, , , , , , , Posted) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                line_no = line_no + 1
        

       End If
       '*****************************************************Ã«—Ì

                        
                        
   If project_id <> 0 And dcproject.BoundText = "" Then
 
       Material_account = get_project_Account(project_id, "expanses_account")
       If Material_account = "" Then
       Material_account = get_project_Account(project_id, "AccountUnderImp")
       End If
       If SystemOptions.gldetails_or_gl_general = 1 Then
                        OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                        OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                        OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                        OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                         OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                   If ModAccounts.AddNewDev(LngDevID, line_no, Material_account, val(.TextMatrix(i, .ColIndex("value"))), 0, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.text), , .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , val(.TextMatrix(i, .ColIndex("fixedid"))), , , brnchid, , , , , , , , , , , , , , , , , , 1, , , , , , , Posted, , OtherInformation) = False Then
                            '   GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
                        
                   If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(.TextMatrix(i, .ColIndex("value"))), 1, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.text), , .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , val(.TextMatrix(i, .ColIndex("fixedid"))), , , brnchid, , , , , , , , , , , , , , , , , , 1, , , , , , , Posted, , OtherInformation) = False Then
                            '   GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
                       End If
 End If
 
                        
                        
        
                    End If

                Next i

            End With
    
            ' «·„’—Ê›«  «·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
            RsNotes.AddNew
            NoteID = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("NoteID").value = CStr(NoteID)
             RsNotes.update
            RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
 
            RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0)
            RsNotes("Remark").value = Me.txt_general_des
            RsNotes("foxy_no").value = val(Text1.text)

        '-------555555555555

            If Me.CboPayMentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPayMentType.ListIndex = 1 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 1
            ElseIf Me.CboPayMentType.ListIndex = 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 3
                            
            ElseIf Me.CboPayMentType.ListIndex = 2 Then
                rs("NoteCashingType").value = 2
                rs("BoxID").value = Null
                rs("BankID").value = val(Me.DcboBankName.BoundText)
                rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                rs("DueDate").value = Me.DtpChequeDueDate.value
                rs("ChequeBoxID").value = Null
            End If
'-------555555555555

            ' RsNotes("order_no").value = txt_ORDER_NO.text
            '              RsNotes("CusID").value = Null
            RsNotes("NoteType").value = 350
            RsNotes("NoteDate").value = XPDtbTrans.value
            RsNotes("UserID").value = user_id
            ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
            RsNotes("notes_all").value = Me.XPTxtID.text
            RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
            RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
            RsNotes("numbering_type1").value = sand_numbering_type(35) '‰Ê⁄  —ﬁÌ„ ‰’›Ì… ⁄Âœ…
            RsNotes("sanad_year").value = year(XPDtbTrans.value)
            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
            RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
            RsNotes.update
    
            '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("DEV_ID_Line_No").value = line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = DcboCreditSide.BoundText
            RsDev("Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text  ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
                       
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
            '   RsDev("project_id").value = project_id
                        
            RsDev.update
     
            'GoTo ll
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 
            line_no = line_no + 1

            If Me.dcproject.BoundText <> "" Then
                '«·ÿ—› «·„œÌ‰   „’—Ê›«  «·„‘—Ê⁄
                RsNotes.AddNew
                NoteID = CStr(new_id("Notes", "NoteID", "", True))
                RsNotes("NoteID").value = CStr(NoteID)
                 RsNotes.update
                RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
          
                RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0)
                RsNotes("Remark").value = txt_general_des.text 'txtto.text

     '--------------555555555
            If Me.CboPayMentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPayMentType.ListIndex = 1 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 1
            ElseIf Me.CboPayMentType.ListIndex = 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 3
                            
            ElseIf Me.CboPayMentType.ListIndex = 2 Then
                rs("NoteCashingType").value = 2
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
    
            End If
'--------------555555555
                        
                ' RsNotes("order_no").value = txt_ORDER_NO.text
                'RsNotes("CusID").value = Null
                RsNotes("NoteType").value = 350
                RsNotes("NoteDate").value = XPDtbTrans.value
                RsNotes("UserID").value = user_id
                ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
                RsNotes("notes_all").value = Me.XPTxtID.text
                RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
                RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
                RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                RsNotes("numbering_type1").value = sand_numbering_type(35) '‰Ê⁄  —ﬁÌ„   ‰’›Ì… ⁄Âœ…
                RsNotes("sanad_year").value = year(XPDtbTrans.value)
                RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                
                RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
                RsNotes.update
                
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                Set RsDev = New ADODB.Recordset
                
         '       RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    StrSQL = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* FROM         dbo.DOUBLE_ENTREY_VOUCHERS WHERE     (Double_Entry_Vouchers_ID = - 1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

                RsDev.AddNew
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
                RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                
                RsDev("Account_Code").value = dcproject.BoundText
                RsDev("Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text ' .TextMatrix(I, .ColIndex("des"))
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)5
                       
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("notes_all").value = Me.XPTxtID.text
                RsDev("project_id").value = project_id
                        
                RsDev.update
                    
                line_no = line_no + 1

                With Fg_Journal

                    For i = .FixedRows To .rows - 1
        
                        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                            '////////////////////////////////////////notes
                
                            If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                                Else
                                    MsgBox "Cant save enter value in line :  " & i - 1, vbCritical: GoTo ErrTrap
                                End If
               
                            End If

                            project_id = get_project_id(dcproject.BoundText, "expanses_account")
                            brnchid = .TextMatrix(i, .ColIndex("BrnchID"))
   If project_id = 0 Then
project_id = .TextMatrix(i, .ColIndex("projectid"))
End If


                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                        OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                        OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                        OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                        OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                            If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")), 1, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , setfoxy_Line, val(Me.XPTxtID.text), project_id, , , , , , , brnchid, , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                                GoTo ErrTrap
                    
                            End If

                            line_no = line_no + 1
        
                        End If

                    Next i

                End With

                
                sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text) & " and notetype=350" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
                sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text) & " and notetype=350" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
            End If

            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            LblDevID.Caption = LngDevID
            lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If

ll:
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata
    
        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved... " & CHR(13)
                    Msg = Msg + "Do you want to enter another operation?"
        
                End If

                Fg_Journal.Enabled = False

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
                End If

                lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)
        
                Fg_Journal.Enabled = False
        End Select
     
        DcboCreditSide_Change
        'LblLink.Caption = balanceString

        '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
   
        '     If Me.DcCostCenter.BoundText <> "" Then
        save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "  ›« Ê—… „«·Ì…", Me.XPDtbTrans.value
        '     End If
        save_cost_center
        'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„’«—Ì›
     
        ' If saveExpensesDetails(1, TxtSerial.text, TxtSerial1.text, TXT_order_no.text, XPDtbTrans.value) = True Then
        ' End If
    
        'Õ›Ÿ »Ì«‰«  «·‘Ìﬂ« 
        saveChequeBoxContents1 (val(Me.XPTxtID.text))
    Dim acc As String
    acc = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    WriteCustomerBalPublic acc, Balance, balanceString
    LblLink1.Caption = balanceString
    fillapprovData
        TxtModFlg.text = "R"
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
            Msg = "cant save " & CHR(13)
            Msg = Msg + "Invalid entry value " & CHR(13)
            Msg = Msg + "Check data and try again"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorr.... Error during saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.text) Then Exit Function
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.text
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteSerial").value = TxtSerial.text
        rs("Remark").value = "   ›« Ê—… „«·Ì… —ﬁ„ " & TxtSerial1 & "    " & Me.txt_general_des
 
        rs.update
        rs.MoveNext
    Next i

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND  kedno =" & val(Text1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If

   ' rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   
  StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If CboPaymentType1.ListIndex = 0 Then
    With Fg_Journal
 
        .rows = .rows + 1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = .TextMatrix(i, .ColIndex("value"))
                rs("depit_or_credit").value = "„œÌ‰"
                rs("opr_id").value = Me.Text1.text
                rs("kedno").value = Me.Text1.text
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs("general_des").value = 1
                rs.update
        
            End If

        Next i

    End With
Else

    With VSFlexGrid1
 
        .rows = .rows + 1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = .TextMatrix(i, .ColIndex("value"))
                rs("depit_or_credit").value = "„œÌ‰"
                rs("opr_id").value = Me.Text1.text
                rs("kedno").value = Me.Text1.text
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs("general_des").value = 1
                rs.update
        
            End If

        Next i

    End With


End If
    rs.Close
End Function

Private Sub Undo()
    Dim sgl As String
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            sgl = "delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)
         
        Case "E"
            sgl = "delete  marakes_taklefa_temp  where ok is null and  kedno =" & val(Text1.text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            rs.Find "NoteID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If SystemOptions.banks_Accounts3 = True Then
        If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
            Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–… «·⁄„·Ì…"
            Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    End If
    
    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
         '   StrSQL = "Delete From notes Where NoteID=" & val(TXT_A_NoteID.text)
            StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.text)
            
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & val(TxtSerial1.text) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             Cn.Execute " Delete from TblExp301UnitNo where  ExpID =" & val(XPTxtID.text)
             Cn.Execute " Delete from TblExpensesDet301 where  ExpID =" & val(XPTxtID.text)
    
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
       
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.rows = 3
                    Fg_Journal.Enabled = False
                
                    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                    VSFlexGrid1.rows = 2
                    VSFlexGrid1.Enabled = False
                
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Function FillGridWithData()

End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

    IntCounter = 0

    With Me.VSFlexGrid1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Private Sub PutData()

    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
    With Fg_Journal

        If Len(TxtDes.text) > 0 Then
            .cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.text
            .cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        Else
            .cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        End If

    End With

End Sub

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As String
    auto_sanad_no = ""
    departement_name = 1
    branch_no = 1
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=1"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at
        end_at = numbering.Recordset.Fields!end_at

    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=350 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
        detect_no.Refresh

        If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
            If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
            If detect_no.Recordset.Fields!last_sand_no >= end_at Then
                sand_numbering = "error"
                Exit Function
            End If
        End If

    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=350 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(Now, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=350 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
                        sand_numbering = "error"
                        Exit Function
                    End If
                End If
 
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
                    auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & mId(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

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
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    sand_numbering = auto_sanad_no

    'MsgBox auto_sanad_no

End Function

Function setfoxy_Line() As Double
    
    Dim X As Double
    X = CStr(new_id("foxy", "id1", "", True))
    setfoxy_Line = X
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = X ' last_line_id
 
    rs.update
    
End Function

Private Sub CBoBasedON_Change()

    'n
    With Me.Fg_Journal

        If Me.CBoBasedON.ListIndex = 0 Then

        ElseIf Me.CBoBasedON.ListIndex = 1 Then

            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·«„—"
            Else
                lbl(21).Caption = "  Order No"
            End If

        ElseIf Me.CBoBasedON.ListIndex = 2 Then

            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·›« Ê—… «·„»œ∆ÌÂ"
            Else
                lbl(21).Caption = "Performa Invoice NO"
            End If

        End If

        .TextMatrix(0, .ColIndex("order_no")) = lbl(21).Caption

    End With

End Sub

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·›« Ê—… " & TxtSerial1.text & CHR(13) & "   «· «—ÌŒ  " & XPDtbTrans & CHR(13) & "   «·›—⁄ " & dcBranch & CHR(13) & "   „—ﬂ“ «· ﬂ·›… «·⁄«„  " & DcCostCenter & CHR(13) & "   ÿ—Ìﬁ… «·œ›⁄  " & CboPayMentType & CHR(13) & "   «·„‘—Ê⁄  " & dcproject & CHR(13) & "   «·„Ê—œ " & DCVendor & CHR(13) & "   «·Œ“Ì‰… " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "   —ﬁ„ ›« Ê—… «·„Ê—œ " & txtto & CHR(13) & "   »‰«¡ ⁄·Ï  " & CBoBasedON & "  »—ﬁ„  " & TXT_order_no & CHR(13) & "   «·‘—Õ «·⁄«„  " & txt_general_des & CHR(13) & "   «Ã„«·Ì «·›« Ê—…    " & XPTxtValView
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill No " & TxtSerial1.text & CHR(13) & "   Date  " & XPDtbTrans & CHR(13) & "   Branch " & dcBranch & CHR(13) & "   CC  " & DcCostCenter & CHR(13) & "  Payment Type  " & CboPayMentType & CHR(13) & "   Project  " & dcproject & CHR(13) & "   Supplier " & DCVendor & CHR(13) & "   Box " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No:   " & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & "  Supplier Bill No " & txtto & CHR(13) & "   Based On  " & CBoBasedON & "  No:  " & TXT_order_no & CHR(13) & "  Remarks  " & txt_general_des & CHR(13) & "   Bill Total   " & XPTxtValView
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 350, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtSerial, TxtSerial1
    Else
        AddToLogFile CInt(user_id), 350, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtSerial, TxtSerial1
    End If
    
End Function

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
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If BolRtl = True Then

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    Else

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "Add New Record..." & Wrap & "Shortcut Key F12 OR Enter" & Wrap & "OR Alt+N", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the Current Record..." & Wrap & "Shortcut Key F11 " & Wrap & "OR Alt+E", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save the New Record OR Save the Editing in the Current Record..." & Wrap & "Shortcut Key F10 " & Wrap & "OR Alt+S", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & "Shortcut Key F9 " & Wrap & "OR Alt+U", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete the Current Record..." & Wrap & "Shortcut Key F8 " & Wrap & "OR Alt+D", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Close this Screen" & Wrap & "OR Alt+X", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Display Help for this Screen" & Wrap & "Shortcut Key F1" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
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

Private Sub XPCboExpensesType_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", val(Me.XPCboExpensesType.BoundText))
    End If

End Sub

Private Sub XPDtbTrans_Change()

    If Me.TxtModFlg = "E" Then
        If Month(rs("NoteDate").value) = Month(XPDtbTrans.value) And year(rs("NoteDate").value) = year(XPDtbTrans.value) Then Exit Sub
    End If

    If Trim(TxtSerial1.text) <> "" Then
        oldTxtSerial1.text = TxtSerial1.text
    End If

    TxtSerial.text = ""
    TxtSerial1.text = ""

End Sub

Private Sub XPTxtVal_Change()
    'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0)
    XPTxtValView.text = Format(val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 0)

    Else

        'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 1)

    End If
    
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Sub XPTxtVal_Validate(Cancel As Boolean)
    'If Val(XPTxtVal.Text) = 0 Then
    '    Set TTD = New clstooltipdemand
    '    TTD.Style = TTBalloon
    '    TTD.Icon = TTIconWarning
    '    TTD.Centered = True
    '    TTD.RightToLeft = True
    '    TTD.VisibleTime = 600
    '    TTD.BackColor = 0
    '    TTD.Title = "ﬁÌ„… «·„’—Ê›« "
    '    TTD.TipText = "»—Ã«¡ ﬂ «»… ﬁÌ„… «·„’—Ê›« "
    '    TTD.PopupOnDemand = True
    '    TTD.CreateToolTip XPTxtVal.hwnd
    '    TTD.Show 0, XPTxtVal.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
    '    Cancel = True
    'Else
    '    TTD.Destroy
    'End If
End Sub

Private Sub ViewDataList()
   Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    'Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

    With FG
        .Cols = 18
        .RowHeightMin = 320
        .ExplorerBar = flexExSortShowAndMove
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
        StrSQL = StrSQL + " Order By NoteID"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
        '------------------------------------
        '
        '
        '
        '
    
        '------------------------------------
        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        'Rs.Close
        'Set Rs = Nothing
        .AutoSize 0, .Cols - 1, False
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
    FrmView.vsfGroup1.sql = StrSQL
    FrmView.vsfGroup1.ShowTreeGroups = True
    FrmView.vsfGroup1.update
    FrmView.SetDblClickRetrun Me, "NoteID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
    FrmView.show
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    'LblValue.Visible = False
    CmdAttach.Caption = "Attachments"
    
    lbl(24).Caption = "Hint"
    lbl(25).Caption = "This Window Allow To Refister Financial Invoice"
    lbl(23).Caption = "Invoice Type"
    Label3.Caption = "GL No."
    lbl(14).Caption = "Project#"
    'Label1.Caption = "Manual #"
    Me.ALLButton1.Caption = "Cost Center"
    lbl(15).Caption = "Payment Method"
    lbl(16).Caption = "Box Name"
    lbl(20).Caption = "General Des"
    lbl(21).Caption = "Order No:"

    Label8.Caption = "General C. C."
   ' Label1.Caption = "Branch"
    lbl(26).Caption = "Based ON"

    With Me.CBoBasedON
        .Clear
        .AddItem "Without"
        .AddItem "Purchase Invoices"
        .AddItem "Performa Invoices"
        .AddItem "Production Order"
    End With

    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Credit"
        .AddItem "P Cheque "
    End With

    With Me.CboPaymentType1
        .Clear
        .AddItem "Expenses"
        .AddItem "Accounts"
     
    End With

    CmdRemove.Caption = "Delete Row"
    Me.Caption = "Petty Cash Settlement "
  '  Me.Ele.Caption = Me.Caption
Frame4.Caption = "Details"
lbl(28).Caption = "Current Balance"
lbl(29).Caption = "Petty Cash Value"

lbl(30).Caption = "Total Expenses"
lbl(31).Caption = "Remains"
Frame5.Caption = "Notes"


lbl(32).Caption = "Settlement  Value"

    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.lbl(4).Caption = " Vchr#"
    Me.lbl(1).Caption = " Date"
    Me.lbl(3).Caption = "Expenses Type"
    Me.lbl(2).Caption = "Total"
    Me.lbl(0).Caption = "Vendor Bill#"
    Me.lbl(5).Caption = "Remarks"
    Me.lbl(8).Caption = "Issued By."
    Me.lbl(7).Caption = "Current Record."

   ' Fra.Caption = "GL"
    lbl(11).Caption = "GL#"
    lbl(13).Caption = "interval"
    lbl(9).Caption = "Depit"
    lbl(10).Caption = "Credit"
    lbl(17).Caption = "Bank"
    lbl(18).Caption = "Cheque#"
    lbl(19).Caption = "Due Date"
    lbl(22).Caption = "Vendor"

    Me.Cmd(0).Caption = "&New"
    Me.Cmd(1).Caption = "&Edit"
    Me.Cmd(2).Caption = "&Save"
    Me.Cmd(3).Caption = "&Undo"
    Me.Cmd(4).Caption = "&Delete"
    Me.Cmd(5).Caption = "Sear&ch"
    Me.Cmd(6).Caption = "E&xit"
    Me.Cmd(7).Caption = "&Table View"
    Cmd(8).Caption = "Print"
    Cmd(9).Caption = "Cheque Print"
    Cmd(10).Caption = "GL Print "

    Me.CmdHelp.Caption = "&Help"

    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Expenses Name"
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Expenses Code"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("des")) = "description"
        .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("ProjectCode")) = "Project Code"
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("pand")) = "Des"
        .TextMatrix(0, .ColIndex("oper")) = "Process"
        .TextMatrix(0, .ColIndex("Vatyo")) = "VAT %"
        .TextMatrix(0, .ColIndex("Vat")) = "VAT"
        '########## khaled ##########
        .TextMatrix(0, .ColIndex("AccountCode")) = "Account No."
        .TextMatrix(0, .ColIndex("project")) = "Project"
        
        .TextMatrix(0, .ColIndex("oper")) = "Process"
        .TextMatrix(0, .ColIndex("Fixes")) = "Equipment"

    End With
CMDRemoveAll.Caption = "Delete All"
    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Account Name"
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Account Code  "
        .TextMatrix(0, .ColIndex("Value")) = "value"
        .TextMatrix(0, .ColIndex("ProjectCode")) = "Project Code"
        .TextMatrix(0, .ColIndex("aqarname")) = "Real Estate"
        .TextMatrix(0, .ColIndex("name")) = "Unit"
        .TextMatrix(0, .ColIndex("unitnoName")) = "Unit No."
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("pand")) = "Des"
        .TextMatrix(0, .ColIndex("oper")) = "Process"
        .TextMatrix(0, .ColIndex("Vatyo")) = "VAT %"
        .TextMatrix(0, .ColIndex("Vat")) = "VAT"
        '########### khaled ###########
        .TextMatrix(0, .ColIndex("des")) = "Description"
        .TextMatrix(0, .ColIndex("FixedAsset")) = "Equipment"
        .TextMatrix(0, .ColIndex("Departement")) = "Departement"
        .TextMatrix(0, .ColIndex("NEmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("billno")) = "Inv No."
    End With

End Sub

