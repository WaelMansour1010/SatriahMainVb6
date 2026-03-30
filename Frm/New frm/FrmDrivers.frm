VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Begin VB.Form FrmDriversx 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·”«∆ﬁÌ‰"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   HelpContextID   =   70
   Icon            =   "FrmDrivers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9735
   ScaleWidth      =   7320
   Begin VB.CommandButton Command9 
      Caption         =   "Tools2"
      Height          =   315
      Left            =   8520
      TabIndex        =   267
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
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
      Height          =   3525
      Index           =   7
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   240
      Top             =   4800
      Width           =   6855
      Begin VB.TextBox txtopening_balance_voucher_id 
         Height          =   735
         Left            =   960
         TabIndex        =   266
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—’Ìœ «·√›  «ÕÏ «ÃÊ— „Œ’’« "
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
         Index           =   10
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   257
         Top             =   1680
         Width           =   3075
         Begin VB.TextBox TxtOpenBalance2 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   261
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton OptType2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œÌ‰"
            Height          =   255
            Index           =   0
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   260
            Top             =   210
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton OptType2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ«∆‰"
            Height          =   255
            Index           =   1
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   259
            Top             =   210
            Width           =   915
         End
         Begin VB.OptionButton OptType2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "€Ì— „Õœœ"
            Height          =   255
            Index           =   2
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   258
            Top             =   210
            Width           =   915
         End
         Begin MSComCtl2.DTPicker Dtp2 
            Height          =   330
            Left            =   150
            TabIndex        =   262
            Top             =   900
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CustomFormat    =   "yyyy/M/d"
            Format          =   94765059
            CurrentDate     =   38718
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·—’Ìœ "
            Height          =   345
            Index           =   18
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   264
            Top             =   510
            Width           =   1125
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”ÃÌ·"
            Height          =   315
            Index           =   17
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   263
            Top             =   900
            Width           =   1125
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—’Ìœ «·√›  «ÕÏ «ÃÊ— „” Õﬁ…"
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
         Index           =   9
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   249
         Top             =   240
         Width           =   3075
         Begin VB.OptionButton OptType1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "€Ì— „Õœœ"
            Height          =   255
            Index           =   2
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   253
            Top             =   210
            Width           =   915
         End
         Begin VB.OptionButton OptType1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ«∆‰"
            Height          =   255
            Index           =   1
            Left            =   1110
            RightToLeft     =   -1  'True
            TabIndex        =   252
            Top             =   210
            Width           =   915
         End
         Begin VB.OptionButton OptType1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œÌ‰"
            Height          =   255
            Index           =   0
            Left            =   1950
            RightToLeft     =   -1  'True
            TabIndex        =   251
            Top             =   210
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.TextBox TxtOpenBalance1 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   250
            Top             =   480
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker Dtp1 
            Height          =   330
            Left            =   150
            TabIndex        =   254
            Top             =   900
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CustomFormat    =   "yyyy/M/d"
            Format          =   94765059
            CurrentDate     =   38718
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”ÃÌ·"
            Height          =   315
            Index           =   16
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   256
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·—’Ìœ "
            Height          =   345
            Index           =   15
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   255
            Top             =   510
            Width           =   1125
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—’Ìœ «·√›  «ÕÏ –„„"
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
         Index           =   8
         Left            =   3660
         RightToLeft     =   -1  'True
         TabIndex        =   241
         Top             =   240
         Width           =   3075
         Begin VB.TextBox TxtOpenBalance 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   245
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
            TabIndex        =   244
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
            TabIndex        =   243
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
            TabIndex        =   242
            Top             =   210
            Width           =   915
         End
         Begin MSComCtl2.DTPicker Dtp 
            Height          =   330
            Left            =   150
            TabIndex        =   246
            Top             =   900
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CustomFormat    =   "yyyy/M/d"
            Format          =   94765059
            CurrentDate     =   38718
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·—’Ìœ "
            Height          =   345
            Index           =   14
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   248
            Top             =   510
            Width           =   1125
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”ÃÌ·"
            Height          =   315
            Index           =   13
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   247
            Top             =   900
            Width           =   1125
         End
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "⁄‰ «œŒ«· »Ì«‰«  „ÊŸ› Ì „ › Õ 3 Õ”«»«  «·Ì… ·Â ÊÂÌ Õ”«» «·«ÃÊ— «·„” Õﬁ… Ê –„„ Ê „Œ’’« "
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
         Height          =   615
         Index           =   6
         Left            =   0
         TabIndex        =   268
         Top             =   1680
         Width           =   3525
      End
   End
   Begin MSDataListLib.DataCombo DcCostCenter 
      Height          =   315
      Left            =   7440
      TabIndex        =   227
      Top             =   1920
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   495
      Index           =   8
      Left            =   9720
      TabIndex        =   228
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Text            =   "„—ﬂ“ «· ﬂ·›…"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   495
      Index           =   5
      Left            =   13200
      TabIndex        =   122
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Text            =   "«·„‘—Ê⁄"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5070
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   238
      Top             =   600
      Width           =   765
   End
   Begin VB.TextBox XPTxtEmpNamee 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   -1200
      TabIndex        =   237
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   235
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   234
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3840
      MaxLength       =   50
      TabIndex        =   233
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   232
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  ’«Õ» «·⁄„·"
      ForeColor       =   &H000000C0&
      Height          =   1815
      Index           =   6
      Left            =   120
      TabIndex        =   133
      Top             =   7560
      Width           =   3195
      Begin VB.TextBox txtkafeladd 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   1320
         Width           =   1845
      End
      Begin VB.TextBox txtkafeltel 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         MaxLength       =   30
         TabIndex        =   42
         Top             =   960
         Width           =   1845
      End
      Begin VB.TextBox txtKafelName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         MaxLength       =   50
         TabIndex        =   41
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtKafelID 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         MaxLength       =   30
         TabIndex        =   40
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   285
         Index           =   33
         Left            =   2040
         TabIndex        =   160
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· ·Ì›Ê‰"
         Height          =   285
         Index           =   32
         Left            =   2010
         TabIndex        =   159
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—ﬁ„"
         Height          =   285
         Index           =   26
         Left            =   2010
         TabIndex        =   135
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„"
         Height          =   285
         Index           =   25
         Left            =   2070
         TabIndex        =   134
         Top             =   660
         Width           =   1005
      End
   End
   Begin MSDataListLib.DataCombo DCBranch 
      Height          =   315
      Left            =   10800
      TabIndex        =   230
      Top             =   1800
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Frame framx 
      BackColor       =   &H8000000A&
      Caption         =   "„›—œ«  «·—« »"
      Height          =   3495
      Left            =   -1080
      RightToLeft     =   -1  'True
      TabIndex        =   161
      Top             =   -840
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command7 
         Caption         =   "Õ”«»"
         Height          =   315
         Left            =   2760
         TabIndex        =   221
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check13 
         Height          =   195
         Left            =   2640
         TabIndex        =   220
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check12 
         Height          =   195
         Left            =   2640
         TabIndex        =   219
         Top             =   2040
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check11 
         Height          =   195
         Left            =   2640
         TabIndex        =   218
         Top             =   1680
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check10 
         Height          =   195
         Left            =   2640
         TabIndex        =   217
         Top             =   1320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check9 
         Height          =   195
         Left            =   2640
         TabIndex        =   216
         Top             =   960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check8 
         Height          =   195
         Left            =   240
         TabIndex        =   215
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check7 
         Height          =   195
         Left            =   240
         TabIndex        =   214
         Top             =   2040
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check6 
         Height          =   195
         Left            =   240
         TabIndex        =   213
         Top             =   1680
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check5 
         Height          =   195
         Left            =   240
         TabIndex        =   212
         Top             =   1320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check4 
         Height          =   195
         Left            =   240
         TabIndex        =   211
         Top             =   840
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check3 
         Height          =   195
         Left            =   240
         TabIndex        =   210
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox TXTMANG 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   208
         Top             =   1950
         Width           =   1335
      End
      Begin VB.TextBox TXTMANGM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   207
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TXTMOB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   205
         Top             =   1590
         Width           =   1335
      End
      Begin VB.TextBox TXTMOBM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   204
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Height          =   195
         Left            =   2640
         TabIndex        =   203
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "«Œ›«¡"
         Height          =   315
         Left            =   600
         TabIndex        =   195
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   182
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   181
         Top             =   4800
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   180
         Top             =   4440
         Width           =   615
      End
      Begin VB.Frame Frame5 
         Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
         Height          =   3255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   178
         Top             =   4800
         Visible         =   0   'False
         Width           =   4935
         Begin VB.CommandButton Command6 
            Caption         =   "„Ê«›ﬁ"
            Height          =   315
            Left            =   120
            TabIndex        =   199
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   1560
            Width           =   4695
         End
         Begin VB.Frame Frame6 
            Height          =   495
            Left            =   240
            TabIndex        =   188
            Top             =   600
            Width           =   1935
            Begin VB.CommandButton Command1 
               Caption         =   "/"
               Height          =   315
               Index           =   5
               Left            =   480
               TabIndex        =   193
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "*"
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   192
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "-"
               Height          =   315
               Index           =   2
               Left            =   1200
               TabIndex        =   191
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "+"
               Height          =   315
               Index           =   3
               Left            =   1560
               TabIndex        =   190
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "="
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   189
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Caption         =   "«Õ„«·Ì «·„ﬁ«„"
               ForeColor       =   &H000000FF&
               Height          =   15
               Left            =   0
               TabIndex        =   194
               Top             =   2520
               Width           =   1935
            End
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   1200
            Width           =   4695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "”‰ÊÌ"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   185
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "‘Â—Ì"
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   184
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "FrmDrivers.frx":038A
            Left            =   2160
            List            =   "FrmDrivers.frx":0391
            RightToLeft     =   -1  'True
            TabIndex        =   179
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   44
            Left            =   1680
            TabIndex        =   198
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰ ÌÃ…"
            Height          =   285
            Index           =   43
            Left            =   3600
            TabIndex        =   197
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
            Height          =   285
            Index           =   42
            Left            =   3480
            TabIndex        =   186
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   285
            Index           =   41
            Left            =   3720
            TabIndex        =   183
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   177
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtanotherm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   176
         Top             =   2370
         Width           =   1335
      End
      Begin VB.TextBox txtfoodm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   175
         Top             =   1170
         Width           =   1335
      End
      Begin VB.TextBox txtbusm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   174
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtsaknm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   173
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtanother 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   170
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtfood 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   168
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtbus 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   165
         Top             =   750
         Width           =   1335
      End
      Begin VB.TextBox txtsakn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   163
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· ≈‘—«›"
         Height          =   285
         Index           =   45
         Left            =   4200
         TabIndex        =   209
         Top             =   2070
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· ÃÊ«·"
         Height          =   285
         Index           =   36
         Left            =   4200
         TabIndex        =   206
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„… «·‘Â—Ì…"
         Height          =   285
         Index           =   40
         Left            =   480
         TabIndex        =   172
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„… «·”‰ÊÌ…"
         Height          =   285
         Index           =   39
         Left            =   2640
         TabIndex        =   171
         Top             =   120
         Width           =   1515
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ·«  «Œ—Ï"
         Height          =   285
         Index           =   38
         Left            =   4200
         TabIndex        =   169
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· ÿ⁄«„"
         Height          =   285
         Index           =   37
         Left            =   4200
         TabIndex        =   167
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· „Ê«’·« "
         Height          =   285
         Index           =   35
         Left            =   4200
         TabIndex        =   164
         Top             =   870
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· «·”ﬂ‰"
         Height          =   285
         Index           =   34
         Left            =   3840
         TabIndex        =   162
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Tools"
      Height          =   315
      Left            =   7560
      TabIndex        =   225
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataListLib.DataCombo dcproject 
      Height          =   315
      Left            =   10800
      TabIndex        =   223
      Top             =   2280
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox TXT_WORK_PLACE 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox TxtAccountCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   10
      TabIndex        =   200
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtSalary 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3950
      MaxLength       =   10
      TabIndex        =   7
      Text            =   "0"
      Top             =   1680
      Width           =   1815
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Index           =   7
      Left            =   15480
      TabIndex        =   158
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "⁄Âœ… «·„ÊŸ›"
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
      MICON           =   "FrmDrivers.frx":0398
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   -1  'True
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Index           =   6
      Left            =   10800
      TabIndex        =   157
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "   «·⁄ﬁÊœ"
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
      MICON           =   "FrmDrivers.frx":03B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   156
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "„›—œ«  «·—« »"
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
      MICON           =   "FrmDrivers.frx":03D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Index           =   3
      Left            =   7440
      TabIndex        =   155
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·„·› «·’ÕÌ"
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
      MICON           =   "FrmDrivers.frx":03EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Index           =   2
      Left            =   10800
      TabIndex        =   154
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«· ﬁÌÌ„"
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
      MICON           =   "FrmDrivers.frx":0408
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   153
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«· «»⁄Ì‰"
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
      MICON           =   "FrmDrivers.frx":0424
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   152
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·„ƒÂ·«  Ê «·Œ»—« "
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
      MICON           =   "FrmDrivers.frx":0440
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "ÿ»«⁄… «·’Ê—…"
      Height          =   255
      Left            =   12360
      TabIndex        =   150
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "„⁄·Ê„«  œŒÊ· «·ÕœÊœ"
      Height          =   1335
      Left            =   7320
      TabIndex        =   144
      Top             =   2400
      Width           =   3375
      Begin VB.TextBox txthdomnfaz 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   146
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txthdodno 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   145
         Top             =   240
         Width           =   1935
      End
      Begin Dynamic_Byte.NourHijriCal txthdoddate 
         Height          =   255
         Left            =   120
         TabIndex        =   224
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "„‰›– «·œŒÊ·  "
         Height          =   285
         Index           =   31
         Left            =   2160
         TabIndex        =   149
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   " «—ÌŒ «·œŒÊ·  "
         Height          =   285
         Index           =   30
         Left            =   2160
         TabIndex        =   148
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "—ﬁ„  œŒÊ· «·ÕœÊœ"
         Height          =   285
         Index           =   29
         Left            =   2160
         TabIndex        =   147
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "«” ⁄·«„« "
      Height          =   975
      Left            =   7320
      TabIndex        =   137
      Top             =   3720
      Width           =   3375
      Begin VB.CommandButton CommandÛQRY 
         Caption         =   "«” ⁄·«„  "
         Height          =   315
         Left            =   240
         TabIndex        =   141
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptExpirLinc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "«‰ Â«¡ «·—Œ’…"
         Height          =   255
         Left            =   1800
         TabIndex        =   140
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptExpirPas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "«‰ Â«¡ ÃÊ«“ «·”›—"
         Height          =   255
         Left            =   1560
         TabIndex        =   139
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptExpirEkama 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "«‰ Â«¡ «·«ﬁ«„…"
         Height          =   255
         Left            =   360
         TabIndex        =   138
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3840
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4850
      MaxLength       =   50
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   9960
      TabIndex        =   130
      Top             =   8640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÿ»«⁄Â"
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
      MICON           =   "FrmDrivers.frx":045C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmDrivers.frx":0478
      Left            =   10800
      List            =   "FrmDrivers.frx":048E
      TabIndex        =   129
      Top             =   8640
      Width           =   3135
   End
   Begin DBPIXLib.DBPix20 DBPix201 
      Height          =   855
      Left            =   12480
      TabIndex        =   124
      Top             =   2640
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   1508
      _StockProps     =   1
      BackColor       =   12632256
      _Image          =   "FrmDrivers.frx":0507
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   495
      Index           =   0
      Left            =   7320
      TabIndex        =   117
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      Text            =   "«·„ƒÂ·«  Ê«·Œ»—« "
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin VB.CheckBox Chk_EndWork 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›’·"
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Chk_Stkala 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton CmdEstkala 
      Caption         =   "«·”»»"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   3480
      TabIndex        =   110
      Top             =   5280
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox Txt_NotEndWork 
         Alignment       =   1  'Right Justify
         Height          =   1545
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   226
         Top             =   120
         Width           =   3855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   315
         Left            =   120
         TabIndex        =   111
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÃÊ«“ «·”›—"
      ForeColor       =   &H000000C0&
      Height          =   1815
      Index           =   5
      Left            =   3480
      TabIndex        =   105
      Top             =   7200
      Width           =   3675
      Begin VB.TextBox txtpasplace 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   360
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1320
         Width           =   1845
      End
      Begin VB.TextBox Txt_NumPasp 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   360
         MaxLength       =   30
         TabIndex        =   31
         Top             =   240
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker Txt_DateExpPasp 
         Height          =   315
         Left            =   360
         TabIndex        =   32
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker Txt_DatePasp 
         Height          =   315
         Left            =   360
         TabIndex        =   33
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ﬂ«‰ «·«’œ«—"
         Height          =   405
         Index           =   24
         Left            =   2160
         TabIndex        =   136
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ ÃÊ«“ «·”›— "
         Height          =   285
         Index           =   23
         Left            =   2250
         TabIndex        =   108
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " - «·«’œ«—"
         Height          =   285
         Index           =   22
         Left            =   2070
         TabIndex        =   107
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " - «·«‰ Â«¡"
         Height          =   405
         Index           =   21
         Left            =   2100
         TabIndex        =   106
         Top             =   960
         Width           =   1005
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ›ÌŸ…"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Index           =   4
      Left            =   120
      TabIndex        =   95
      Top             =   6240
      Width           =   3195
      Begin Dynamic_Byte.NourHijriCal Txt_DateExppoketH 
         Height          =   255
         Left            =   360
         TabIndex        =   265
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin Dynamic_Byte.NourHijriCal Txt_DateEndpoketH 
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin VB.TextBox Tet_NumPoket 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         MaxLength       =   30
         TabIndex        =   38
         Top             =   240
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker Txt_DateExppoket 
         Height          =   315
         Left            =   360
         TabIndex        =   103
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker Txt_DateEndpoket 
         Height          =   315
         Left            =   360
         TabIndex        =   104
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " - «·«‰ Â«¡"
         Height          =   405
         Index           =   20
         Left            =   2100
         TabIndex        =   98
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " - «·«’œ«—"
         Height          =   285
         Index           =   19
         Left            =   2070
         TabIndex        =   97
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Õ›ÌŸ…"
         Height          =   285
         Index           =   18
         Left            =   2010
         TabIndex        =   96
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«ﬁ«„…"
      ForeColor       =   &H000000C0&
      Height          =   1815
      Index           =   3
      Left            =   3480
      TabIndex        =   90
      Top             =   5280
      Width           =   3675
      Begin Dynamic_Byte.NourHijriCal Txt_DateEndekamah 
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin Dynamic_Byte.NourHijriCal Txt_DateExpEkamaH 
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin VB.TextBox Txt_placEkama 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         MaxLength       =   50
         TabIndex        =   27
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox Txt_NumEkama 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         MaxLength       =   30
         TabIndex        =   28
         Top             =   600
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker Txt_DateExpEkama 
         Height          =   315
         Left            =   360
         TabIndex        =   99
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker Txt_DateEndekama 
         Height          =   315
         Left            =   360
         TabIndex        =   100
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " -«·«‰ Â«¡"
         Height          =   405
         Index           =   17
         Left            =   2100
         TabIndex        =   94
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ﬂ«‰ «·«ﬁ«„…"
         Height          =   285
         Index           =   16
         Left            =   2010
         TabIndex        =   93
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·«ﬁ«„…"
         Height          =   285
         Index           =   15
         Left            =   2100
         TabIndex        =   92
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " -«·«’œ«—"
         Height          =   405
         Index           =   14
         Left            =   2100
         TabIndex        =   91
         Top             =   960
         Width           =   1005
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—Œ’…"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Index           =   2
      Left            =   120
      TabIndex        =   86
      Top             =   4920
      Width           =   3195
      Begin Dynamic_Byte.NourHijriCal Txt_DateEndLincH 
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin Dynamic_Byte.NourHijriCal Txt_DateExpLincH 
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin VB.TextBox Txt_NumLicn 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         MaxLength       =   30
         TabIndex        =   35
         Top             =   240
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker Txt_DateExpLinc 
         Height          =   315
         Left            =   360
         TabIndex        =   101
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker Txt_DateEndLinc 
         Height          =   315
         Left            =   360
         TabIndex        =   102
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·—Œ’…"
         Height          =   285
         Index           =   13
         Left            =   2010
         TabIndex        =   89
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " - «·«’œ«—"
         Height          =   285
         Index           =   12
         Left            =   2070
         TabIndex        =   88
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " - «·«‰ Â«¡"
         Height          =   285
         Index           =   11
         Left            =   2100
         TabIndex        =   87
         Top             =   960
         Width           =   1005
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  Œ«’… »«·⁄„·"
      ForeColor       =   &H000000C0&
      Height          =   1635
      Index           =   1
      Left            =   60
      TabIndex        =   80
      Top             =   2100
      Width           =   3195
      Begin VB.TextBox TxtRegion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         TabIndex        =   85
         Top             =   1620
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo DcboEmpDepartments 
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   270
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboJobsType 
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboSpecifications 
         Height          =   315
         Left            =   60
         TabIndex        =   15
         Top             =   930
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94765057
         CurrentDate     =   38784
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ »œ¡ «·⁄„·"
         Height          =   285
         Index           =   9
         Left            =   1860
         TabIndex        =   131
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‰ÿﬁ…"
         Height          =   225
         Index           =   10
         Left            =   2400
         TabIndex        =   84
         Top             =   1650
         Width           =   735
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· Œ’’"
         Height          =   225
         Index           =   9
         Left            =   2280
         TabIndex        =   83
         Top             =   930
         Width           =   885
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·ÊŸÌ›…"
         Height          =   225
         Index           =   8
         Left            =   2280
         TabIndex        =   82
         Top             =   630
         Width           =   885
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·ﬁ”„"
         Height          =   225
         Index           =   7
         Left            =   2280
         TabIndex        =   81
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· √„Ì‰« "
      ForeColor       =   &H000000C0&
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   74
      Top             =   3540
      Width           =   3195
      Begin VB.TextBox TxtOtherDiscounts 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         TabIndex        =   24
         Top             =   960
         Width           =   1965
      End
      Begin VB.TextBox TxtInsurValue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         TabIndex        =   21
         Top             =   600
         Width           =   1965
      End
      Begin VB.ComboBox CboInsuranceState 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   270
         Width           =   1965
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Œ’Ê„«  √Œ—Ì"
         Height          =   405
         Index           =   6
         Left            =   2100
         TabIndex        =   77
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„»·€ «· √„Ì‰"
         Height          =   285
         Index           =   5
         Left            =   2070
         TabIndex        =   76
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ«·… «· √„Ì‰"
         Height          =   285
         Index           =   4
         Left            =   2010
         TabIndex        =   75
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.TextBox TxtEmpProfitCom 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3540
      MaxLength       =   50
      TabIndex        =   20
      Top             =   4080
      Width           =   2295
   End
   Begin VB.ComboBox CboWorkState 
      Enabled         =   0   'False
      Height          =   315
      Left            =   15660
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   2700
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox TxtEmp_Comm 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3540
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3660
      Width           =   2295
   End
   Begin VB.TextBox TxtEmp_Code 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4620
      MaxLength       =   50
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox XPTxtPhone 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3300
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2355
      Width           =   2535
   End
   Begin VB.TextBox XPTxtmobile 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3300
      MaxLength       =   50
      TabIndex        =   12
      Top             =   2700
      Width           =   2535
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   705
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      Top             =   8100
      Width           =   2535
   End
   Begin VB.TextBox XPTxtProfMail 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3300
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1995
      Width           =   2535
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   -120
      Width           =   7305
      _cx             =   12885
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
      Caption         =   "»Ì«‰«  «·”«∆ﬁÌ‰"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
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
      Begin VB.TextBox Contract_ID 
         Height          =   285
         Left            =   5640
         TabIndex        =   229
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox XPTxtEmpID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   3450
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   2460
         TabIndex        =   78
         Top             =   150
         Visible         =   0   'False
         Width           =   855
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1185
         TabIndex        =   49
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
         ButtonImage     =   "FrmDrivers.frx":051F
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
         TabIndex        =   50
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
         ButtonImage     =   "FrmDrivers.frx":08B9
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
         Left            =   1680
         TabIndex        =   51
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
         ButtonImage     =   "FrmDrivers.frx":0C53
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
         TabIndex        =   52
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
         ButtonImage     =   "FrmDrivers.frx":0FED
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image Img 
         Height          =   480
         Left            =   4320
         Picture         =   "FrmDrivers.frx":1387
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   6435
      TabIndex        =   62
      ToolTipText     =   "⁄‰œ «‰‘«¡ „ÊŸ› Ì „ «‰‘«¡ 3 Õ”«»«  «·Ì… ·Â ÊÂ„  Õ”«» «·–„„ Ê Õ”«» «·«ÃÊ— «·„” Õﬁ… ÊÕ”«» «·„Œ’’« "
      Top             =   9360
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
      Left            =   5715
      TabIndex        =   63
      Top             =   9360
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
      Left            =   4995
      TabIndex        =   64
      Top             =   9360
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
      Left            =   4275
      TabIndex        =   65
      Top             =   9360
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   3465
      TabIndex        =   44
      Top             =   9360
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
      Left            =   2640
      TabIndex        =   66
      Top             =   9360
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   30
      TabIndex        =   67
      Top             =   9360
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
      Left            =   1800
      TabIndex        =   68
      Top             =   9360
      Width           =   795
      _ExtentX        =   1402
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
      Height          =   255
      Left            =   13020
      TabIndex        =   69
      Top             =   9600
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
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
   Begin VB.TextBox XPTxtEmpName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5220
      MaxLength       =   50
      TabIndex        =   45
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DtDate 
      Height          =   315
      Left            =   4740
      TabIndex        =   25
      Top             =   4920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94765057
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd1 
      Height          =   375
      Left            =   840
      TabIndex        =   115
      Top             =   9360
      Width           =   945
      _ExtentX        =   1667
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
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   615
      Index           =   1
      Left            =   7320
      TabIndex        =   118
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      Text            =   "«· «»⁄ÌÌ‰"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   8655
      Index           =   2
      Left            =   7320
      TabIndex        =   119
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   15266
      Text            =   "«· ﬁÌÌ„"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   615
      Index           =   3
      Left            =   14040
      TabIndex        =   120
      Top             =   6960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      Text            =   "«·„·› «·’ÕÌ"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   615
      Index           =   4
      Left            =   7440
      TabIndex        =   121
      Top             =   8280
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      Text            =   "„›—œ«  «·„— »"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   495
      Index           =   6
      Left            =   14160
      TabIndex        =   123
      Top             =   7800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Text            =   "«·„” ‰œ«  Ê «·⁄ﬁÊœ"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   12480
      TabIndex        =   125
      Top             =   3600
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«œ—«Ã ’Ê—… «·„ÊŸ›"
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
   Begin DBPIXLib.DBPix20 DBPix202 
      Height          =   855
      Left            =   10680
      TabIndex        =   126
      Top             =   2640
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   1508
      _StockProps     =   1
      BackColor       =   12632256
      _Image          =   "FrmDrivers.frx":2051
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
   End
   Begin ImpulseButton.ISButton ISButton2 
      Height          =   375
      Left            =   10680
      TabIndex        =   127
      Top             =   3600
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«œ—«Ã  ÊﬁÌ⁄ «·„ÊŸ›"
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   3300
      TabIndex        =   14
      Top             =   3000
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94765057
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DCNationality 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo Dcdean 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   495
      Index           =   7
      Left            =   15360
      TabIndex        =   151
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Text            =   "⁄Âœ… «·„ÊŸ›"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin MSDataListLib.DataCombo dcjopstatus 
      Height          =   315
      Left            =   3300
      TabIndex        =   16
      Top             =   3360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCreditSide 
      Height          =   315
      Left            =   3480
      TabIndex        =   201
      Top             =   9000
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin SuperLablel.SuperLabel lblB 
      Height          =   495
      Index           =   9
      Left            =   12120
      TabIndex        =   231
      Top             =   1680
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      Text            =   "«·›—⁄"
      BackColor       =   12632256
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   4320
      TabIndex        =   239
      Top             =   600
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   285
      Index           =   47
      Left            =   5880
      TabIndex        =   236
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ﬂ«‰ «·⁄„·"
      Height          =   285
      Index           =   46
      Left            =   2280
      TabIndex        =   222
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—› œ«∆‰"
      Height          =   285
      Index           =   31
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   202
      Top             =   9000
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—« » «·«”«”Ì"
      Height          =   285
      Index           =   2
      Left            =   5940
      TabIndex        =   166
      Top             =   1710
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·œÌ«‰…"
      Height          =   225
      Index           =   28
      Left            =   1320
      TabIndex        =   143
      Top             =   600
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ã‰”Ì…"
      Height          =   225
      Index           =   27
      Left            =   3420
      TabIndex        =   142
      Top             =   630
      Width           =   765
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ  «·„Ì·«œ"
      Height          =   285
      Index           =   12
      Left            =   5820
      TabIndex        =   132
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " „«–Ã Â«„… ··ÃÊ«“« "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12240
      TabIndex        =   128
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   9495
      Left            =   10560
      TabIndex        =   116
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "%"
      Height          =   255
      Left            =   3360
      TabIndex        =   114
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "%"
      Height          =   255
      Left            =   3360
      TabIndex        =   113
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " —ﬂ «·⁄„· "
      Height          =   285
      Index           =   11
      Left            =   5880
      TabIndex        =   112
      Top             =   4560
      Width           =   1275
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   10
      Left            =   6360
      TabIndex        =   109
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰”»… «·⁄„Ê·… ⁄·Ï ’«›Ï «·„»Ì⁄« "
      Height          =   405
      Index           =   8
      Left            =   6000
      TabIndex        =   73
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·⁄„·"
      Height          =   285
      Index           =   7
      Left            =   5760
      TabIndex        =   72
      Top             =   3420
      Width           =   1275
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰”»… «·⁄„Ê·… ⁄·Ï ≈Ã„«·Ï «·„»Ì⁄« "
      Height          =   405
      Index           =   5
      Left            =   6000
      TabIndex        =   71
      Top             =   3750
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·»—Ìœ «·«·ﬂ —Ê‰Ì"
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   70
      Top             =   2025
      Width           =   1275
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   30
      TabIndex        =   61
      Top             =   9120
      Width           =   345
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   1440
      TabIndex        =   60
      Top             =   8880
      Width           =   375
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   59
      Top             =   8880
      Width           =   1065
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   420
      TabIndex        =   58
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·Â« ›"
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   56
      Top             =   2385
      Width           =   1275
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÃÊ«·"
      Height          =   285
      Index           =   2
      Left            =   5880
      TabIndex        =   55
      Top             =   2730
      Width           =   1275
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   54
      Top             =   8250
      Width           =   675
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„ÊŸ›"
      Height          =   225
      Index           =   1
      Left            =   5880
      TabIndex        =   53
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   57
      Top             =   945
      Width           =   1275
   End
End
Attribute VB_Name = "FrmDriversx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim EmpReport As ClsEmployeeReport
Dim xReport As New CRAXDRT.Report
Dim NO As Double

Private objScript As Object
Dim case_id As Integer
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
Dim Account_Code_dynamic3 As String
Dim Account_Code_dynamic4 As String
Dim Account_Code_dynamic5 As String
Dim FirstPeriodDateInthisYear  As Date

Function SHOWPIC(PICNAME As String)
    Dim xLogo As CRAXDRT.OLEObject
    StrFileName = App.path & "\Images\" & PICNAME & ".JPG"

    Set xLogo = xReport.Areas(3).Sections(1).AddPictureObject(StrFileName, 4000, 300)
    xLogo.Width = 1700
    xLogo.Height = 1700
    xLogo.backcolor = vbWhite
    xLogo.BorderColor = 255
    xLogo.CloseAtPageBreak = True
    '  xLogo.HyperlinkText = "BYTE"
    '  xLogo.HyperlinkType = crHyperlinkWebsite
    '  rep.Areas(1).Sections(1).SuppressIfBlank = True
    '  rep.Areas(1).Sections(1).Height = xLogo.Height + 250
 
End Function

Private Sub ALLButton1_Click()
    Dim x As String
    On Error Resume Next

    'ALLButton1.Enabled = True
    Select Case Combo1.ListIndex + 1

        Case 1

            Dim xApp As New CRAXDRT.Application

            Dim rs As New ADODB.Recordset

            If SystemOptions.UserInterface = EnglishInterface Then

                x = InputBox("Specify No Of Month ")
            Else
                x = InputBox("Õœœ ⁄œœ ‘ÂÊ— «·«ﬁ«„…", " ÕœÌœ „œ… «·«ﬁ«„… »«·‘ÂÊ—")
            End If

            'If x = 0 Then MsgBox "·«»œ „‰  ÕœÌœ ⁄œœ «·‘ÂÊ— ÊÌﬂÊ‰ «—ﬁ«„ ": Exit Sub

            'Form3.Show
            'Form3.case_id = 1
            'Form3.noofmonth = x
            ' Form3.SHOWPICTURE.Caption = Me.Check1.value
            'Form3.TxtEmp_Code = Me.TxtEmp_Code.text

            sql = "SELECT * from emp_all_details WHERE emp_code='" & FrmEmployee.TxtEmp_Code.text & "'"
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT1.rpt")
            xReport.Database.SetDataSource rs
 
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.TxtPath = (system_path & "\reports\emp\REPORT1.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            xReport.reporttitle = x
            SendKeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

        Case 2
            'Dim xApp As New CRAXDRT.Application
            'Dim Rs As New ADODB.Recordset

            If SystemOptions.UserInterface = EnglishInterface Then

                x = InputBox("Specify No Of Month ")
            Else
                x = InputBox("Õœœ ⁄œœ ‘ÂÊ— «·«ﬁ«„…", " ÕœÌœ „œ… «·«ﬁ«„… »«·‘ÂÊ—")
            End If

            'If x = 0 Then MsgBox "·«»œ „‰  ÕœÌœ ⁄œœ «·‘ÂÊ— ÊÌﬂÊ‰ «—ﬁ«„ ": Exit Sub

            sql = "SELECT * from emp_all_details WHERE emp_code='" & FrmEmployee.TxtEmp_Code.text & "'"
     
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT2.rpt")
            xReport.Database.SetDataSource rs
 
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.TxtPath = (system_path & "\reports\emp\REPORT2.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            xReport.reporttitle = x
            SendKeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

            'Form3.Show
            'Form3.case_id = 2
            'Form3.noofmonth = x
            ' Form3.SHOWPICTURE.Caption = Me.Check1.value
            'Form3.TxtEmp_Code = Me.TxtEmp_Code.text

        Case 3

            sql = "SELECT * from emp_all_details WHERE emp_code='" & FrmEmployee.TxtEmp_Code.text & "'"
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT3.rpt")
            xReport.Database.SetDataSource rs
 
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.TxtPath = (system_path & "\reports\emp\REPORT3.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            '    xReport.ReportTitle = X
            SendKeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

            'Form3.Show
            'Form3.case_id = 3
            ' Form3.SHOWPICTURE.Caption = Me.Check1.value
            'Form3.TxtEmp_Code = Me.TxtEmp_Code.text
        Case 4
            sql = "SELECT * from emp_all_details WHERE emp_code='" & FrmEmployee.TxtEmp_Code.text & "'"
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT4.rpt")
            xReport.Database.SetDataSource rs
 
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.TxtPath = (system_path & "\reports\emp\REPORT4.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            ' xReport.ReportTitle = X
            SendKeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

        Case 5
            outform.show
            outform.Check7.value = Me.Check1.value
            outform.TxtEmp_Code = Me.XPTxtEmpID.text

        Case 6

            If SystemOptions.UserInterface = EnglishInterface Then

                x = InputBox("Specify Date ")
            Else
                x = InputBox("Õœœ  «—ÌŒ «·Â—Ê» ", "10/02/1432")
            End If
    
            If Len(x) <> 10 Then
                If SystemOptions.UserInterface = EnglishInterface Then
        
                    MsgBox "wrong date  ex 11/02/1432 "
                Else
                    MsgBox "Õœœ  «—ÌŒ «·Â—Ê»  «·’ÕÌÕ " & "10/02/1432"
         
                End If

                Exit Sub
            End If

            sql = "SELECT * from emp_all_details WHERE emp_code='" & FrmEmployee.TxtEmp_Code.text & "'"
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT8.rpt")
            xReport.Database.SetDataSource rs
 
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.TxtPath = (system_path & "\reports\emp\REPORT8.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            '    xReport.ReportTitle = X
            xReport.ParameterFields(1).AddCurrentValue Mid(x, 1, 2)
            xReport.ParameterFields(2).AddCurrentValue Mid(x, 4, 2)
            xReport.ParameterFields(3).AddCurrentValue Mid(x, 9, 2)
  
            SendKeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

            'ALLButton1.Enabled = False
    End Select

End Sub

Private Sub ALLButton2_Click(Index As Integer)

    Select Case Index

        Case 0
            mO2AHELAT.show

        Case 1
            TABE3.show

        Case 2
            TAKEEM.show

        Case 3
            SEHY.show

        Case 4

            If Me.TxtModFlg.text = "N" Then
                'If SystemOptions.UserInterface = ArabicInterface Then
                'MsgBox "«Õ›Ÿ »Ì«‰«  «·„ÊŸ› «·«”«”Ì… «Ê·« "
                'Else
                'MsgBox "Save Employee Basic Information Firstly!"
                frmEmpSalaryComponent.show
                frmEmpSalaryComponent.Contract_ID = Me.Contract_ID
                frmEmpSalaryComponent.Emp_id = val(XPTxtEmpID.text)
                frmEmpSalaryComponent.emp_code = TxtEmp_Code.text
                frmEmpSalaryComponent.emp_name(0) = Text1.text
                frmEmpSalaryComponent.emp_name(1) = Text2.text
                frmEmpSalaryComponent.emp_name(2) = Text3.text
                frmEmpSalaryComponent.emp_name(3) = Text4.text
                frmEmpSalaryComponent.Departement.text = DcboEmpDepartments.text
                frmEmpSalaryComponent.job.text = DcboJobsType.text
                frmEmpSalaryComponent.Issue_date.value = DTPicker1.value
                frmEmpSalaryComponent.Basic_salary.text = val(TxtSalary.text)
                frmEmpSalaryComponent.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                frmEmpSalaryComponent.VSFlexGrid1.Rows = 1
                frmEmpSalaryComponent.Cmd_Click (1)

            End If
 
            If val(XPTxtEmpID.text) <> 0 Then
                frmEmpSalaryComponent.show
                frmEmpSalaryComponent.Contract_ID = Me.Contract_ID
                frmEmpSalaryComponent.Retrive val(XPTxtEmpID.text)
            Else

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "ﬂÊœ „ÊŸ› Œÿ√"
                Else
                    MsgBox "Invalid Employee Code"
                End If

            End If

            'Me.framx.Visible = True
            'RWATEB.Show
        Case 5

        Case 6
            frmEmpContract.show
            frmEmpContract.Retrive , val(Me.XPTxtEmpID.text)

            ' Cmd1_Click
        Case 7
            'OHDA.Show
    End Select

End Sub

Private Sub CboInsuranceState_Change()

    If Me.CboInsuranceState.ListIndex = 0 Then
        Me.TxtInsurValue.text = ""
        Me.TxtInsurValue.Enabled = False
    Else
        Me.TxtInsurValue.Enabled = True
    End If

End Sub

Private Sub CboInsuranceState_Click()
    CboInsuranceState_Change
End Sub

Private Sub Chk_EndWork_Click()
    On Error GoTo Errtrp
    '......................................

    If Chk_EndWork.value = Checked Or Me.Chk_Stkala.value = Checked Then
        If Me.TxtModFlg.text = "N" Then
            '                XPTxtValue(1).text = ""
            DtDate.value = Date
        End If

        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            DtDate.Enabled = True
            '               CmdEstkala.Enabled = True

        Else
            DtDate.Enabled = False
            '               CmdEstkala.Enabled = False
        End If

        '            Me.ChkInstall.Enabled = True
    Else
        DtDate.Enabled = False
        '               CmdEstkala.Enabled = False
    End If

    '......................................

Errtrp:

End Sub

Private Sub Chk_Stkala_Click()

    On Error GoTo Errtrp
    '......................................

    If Chk_Stkala.value = Checked Or Chk_EndWork.value = Checked Then
        If Me.TxtModFlg.text = "N" Then
            '                XPTxtValue(1).text = ""
            DtDate.value = Date
        End If

        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            DtDate.Enabled = True
            '               CmdEstkala.Enabled = True

        Else
            DtDate.Enabled = False
            '               CmdEstkala.Enabled = False
        End If

        '            Me.ChkInstall.Enabled = True
    Else
        DtDate.Enabled = False
        '               CmdEstkala.Enabled = False
    End If

    '......................................

Errtrp:

End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            Text1.SetFocus
            Txt_DateExpEkamaH.value = ToHijriDate(Date)
            Txt_DateEndekamah.value = ToHijriDate(Date)
            Txt_DateExpLincH.value = ToHijriDate(Date)
            Txt_DateEndLincH.value = ToHijriDate(Date)
            txthdoddate.value = ToHijriDate(Date)
            Txt_DateExppoketH.value = ToHijriDate(Date)
            Txt_DateEndpoketH.value = ToHijriDate(Date)
        
            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
            Me.Dtp1 = FirstPeriodDateInthisYear
            Me.Dtp2 = FirstPeriodDateInthisYear
   
            OptType(2).value = True
            TxtSalary.text = 0
    
            OptType1(2).value = True
            OptType2(2).value = True
    
        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
            Me.Dtp1 = FirstPeriodDateInthisYear
            Me.Dtp2 = FirstPeriodDateInthisYear

            TxtModFlg.text = "E"

        Case 2
    
            Dim currentcode As String

            If txtid.text = "" Then
                currentcode = get_coding(branch_id, "TblEmployee", 6, Me.DCPreFix.text)

                If currentcode = "miniError" Then
                    MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Exit Sub
            
                ElseIf currentcode = "Manual" Then
                    MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                    Exit Sub
                Else
                    txtid = currentcode
                End If
            End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_ProfData

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            FrmEmployeeSearch.show ' vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            printingReport
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Cmd1_Click()
    On Error Resume Next
 
    If TxtEmp_Code.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „ÊŸ› «Ê·«": Exit Sub

    imaged.show
    imaged.Label9.Caption = "„—›ﬁ«  «·„ÊŸ› —ﬁ„"
    imaged.Caption = "„—›ﬁ«  «·„ÊŸ›  "
    imaged.txtopeation_type = "„—›ﬁ«  „ÊŸ›"
    imaged.SUBJECT_NO = TxtEmp_Code.text
    imaged.Label6.Caption = "ﬂÊœ «·„ÊŸ›"
    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  „ÊŸ›' and subject_no='" & TxtEmp_Code.text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Sub

Private Sub CmdExit_Click()
    Frame1.Visible = False
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdEstkala_Click()
    Frame1.Visible = True
End Sub

Function clear_color()

End Function

Private Sub Combo2_Change()

    If IsNumeric(Combo2.text) Then
        NO = Combo2.text
    Else
        '  MsgBox "·«»œ „‰ ﬂ «»… «—ﬁ«„ ›ﬁÿ"
    End If

End Sub

Private Sub Combo2_Click()

    If Combo2.text = "«·—« » «·«”«”Ì" Then
        NO = IIf(TxtSalary.text = "", 0, val(TxtSalary.text))
    Else

        If Combo2.text = "»œ· ”ﬂ‰" Then
            NO = IIf(txtsakn.text = "", 0, val(txtsakn.text))
        Else

            If Combo2.text = "»œ· „Ê«’·« " Then
                NO = IIf(txtbus.text = "", 0, val(txtbus.text))
 
            Else

                If Combo2.text = "»œ· ÿ⁄«„" Then
                    NO = IIf(txtfood.text = "", 0, val(txtfood.text))

                Else

                    If Combo2.text = "»œ·«  «Œ—Ì" Then
                        NO = IIf(txtanother.text = "", 0, val(txtanother.text))

                    Else

                        If IsNumeric(Combo2.text) Then
                            NO = Combo2.text
                        Else
                            MsgBox "·«»œ „‰ ﬂ «»… «—ﬁ«„ ›ﬁÿ"
                        End If

                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next

    If Command1(Index).Caption <> "=" Then
        Text14.text = Text14.text & NO & Command1(Index).Caption
        Text15.text = Text15.text & Combo2.text & Command1(Index).Caption

    Else

        Text14.text = Text14.text & NO
        Text15.text = Text15.text & Combo2.text
 
        Set objScript = CreateObject("MSScriptControl.ScriptControl")
        objScript.Language = "VBScript"

        XPLbl(44).Caption = objScript.Eval(Text14.text)

    End If

End Sub

Private Sub Command10_Click()

    case_id = 1 = 2
    Combo2.Clear
    Combo2.AddItem "«·—« » «·«”«”Ì"
    Combo2.AddItem "»œ· ”ﬂ‰"
    Combo2.AddItem "»œ· ÿ⁄«„"
    Combo2.AddItem "»œ·«  «Œ—Ì"

End Sub

Private Sub Command2_Click()

    case_id = 1
    Frame5.Visible = True
    Text14.text = ""
    Combo2.text = ""

    Combo2.Clear
    Combo2.AddItem "«·—« » «·«”«”Ì"
    Combo2.AddItem "»œ· „Ê«’·« "
    Combo2.AddItem "»œ· ÿ⁄«„"
    Combo2.AddItem "»œ·«  «Œ—Ì"

End Sub

Private Sub Command3_Click()

    case_id = 3
    Combo2.Clear
    Combo2.AddItem "«·—« » «·«”«”Ì"
    Combo2.AddItem "»œ· ”ﬂ‰"
    Combo2.AddItem "»œ· „Ê«’·« "
    Combo2.AddItem "»œ·«  «Œ—Ì"

End Sub

Private Sub Command4_Click()

    case_id = 4
    Combo2.Clear
    Combo2.AddItem "«·—« » «·«”«”Ì"
    Combo2.AddItem "»œ· ”ﬂ‰"
    Combo2.AddItem "»œ· „Ê«’·« "
    Combo2.AddItem "»œ· ÿ⁄«„"
  
End Sub

Private Sub Command5_Click()
    Me.framx.Visible = False
End Sub

Private Sub Command6_Click()

    Select Case case_id

        Case 1

            If Option1(0).value = True Then
                txtsaknm.text = val(XPLbl(44).Caption)
            Else
                txtsakn.text = val(XPLbl(44).Caption)
            End If

    End Select

End Sub

Private Sub Command8_Click()

    If create_accounts = False Then
        Exit Sub
    End If
        
    'delete old employee account if found
    Dim RsTemp1 As New ADODB.Recordset
    Dim where_str As String
    my_branch = 1
    StrSQL = "select * From branches where  branch_id= " & val(my_branch)
    RsTemp1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim a7 As String
    Dim a29 As String
    Dim a30 As String

    If RsTemp1.RecordCount <> 0 Then
        If IsNull(RsTemp1("a7").value) Or RsTemp1("a7").value = "" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ··„ÊŸ›Ì‰   ·–·ﬂ ·« Ì„ﬂ‰ «·‰⁄œÌ· «·«‰"
            Exit Sub
        Else
            a7 = RsTemp1("a7").value
        End If

        If IsNull(RsTemp1("a29").value) Or RsTemp1("a7").value = "" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ··„ÊŸ›Ì‰   ·–·ﬂ ·« Ì„ﬂ‰ «·‰⁄œÌ· «·«‰"
            Exit Sub
        Else
            a29 = RsTemp1("a29").value

        End If

        If IsNull(RsTemp1("a30").value) Or RsTemp1("a7").value = "" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ··„ÊŸ›Ì‰   ·–·ﬂ ·« Ì„ﬂ‰ «·‰⁄œÌ· «·«‰"
            Exit Sub
        Else
            a30 = RsTemp1("a30").value

        End If

        where_str = "where account_code like'" & a7 & "_%'  or account_code like'" & a29 & "_%'  or account_code like'" & a30 & "_%'"

    End If

    StrSQL = "delete accountS " & where_str
    Cn.Execute StrSQL

    Dim RsTemp As New ADODB.Recordset
 
    StrSQL = "select * From TblEmployee "
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To RsTemp.RecordCount
        RsTemp("Account_Code").value = ModAccounts.AddNewAccount(a7, RsTemp("Emp_Name").value, True, False, IIf(IsNull(RsTemp("Emp_Namee").value), RsTemp("Emp_Name").value, RsTemp("Emp_Namee").value))
        RsTemp("Account_Code1").value = ModAccounts.AddNewAccount(a29, RsTemp("Emp_Name").value & "  «ÃÊ— „” Õ›…", True, False, IIf(IsNull(RsTemp("Emp_Namee").value), RsTemp("Emp_Name").value & "  Salary   ", RsTemp("Emp_Namee").value & " Salary "))
        RsTemp("Account_Code2").value = ModAccounts.AddNewAccount(a30, RsTemp("Emp_Name").value & "Ò „Œ’’«  ", True, False, IIf(IsNull(RsTemp("Emp_Namee").value), RsTemp("Emp_Name").value & "  Reserved  ", RsTemp("Emp_Namee").value & " Reserved"))

        RsTemp.update
        RsTemp.MoveNext
    Next i

    MsgBox " „"

End Sub

Private Sub Command9_Click()
    Dim sql  As String
    sql = "DELETE EmpSalaryComponent  WHERE AccountCode=1"
    Cn.Execute sql
 
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    Dim RsDev As ADODB.Recordset
 
    sql = "select * from TblEmployee "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Set RsDev = New ADODB.Recordset
    RsDev.Open "EmpSalaryComponent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount

            If Not IsNull(rs("Emp_Salary").value) Then
                RsDev.AddNew
                RsDev("Emp_id").value = IIf(IsNull((rs("Emp_ID").value)), 0, rs("Emp_ID").value)
                RsDev("Accountcode").value = 1
                RsDev("value").value = IIf(IsNull((rs("Emp_Salary").value)), 0, rs("Emp_Salary").value)
                RsDev("mofrad_type").value = 1
                RsDev("ModDate").value = Date
                RsDev("Monthly").value = 1
                RsDev("is_fixed").value = 2
                RsDev("Contract_ID").value = 0
                RsDev("specific_value").value = 0
                RsDev("Monthly").value = 1
                RsDev("assurance").value = 0
                RsDev("percentage").value = 0
               
                RsDev.update
            End If
 
            rs.MoveNext
 
        Next i
 
    End If
 
    rs.Close
    MsgBox " „"

End Sub

Private Sub CommandÛQRY_Click()

    'FrmEmpExpir.Show
    If OptExpirEkama.value = True Then
        FrmEmpExpir2.show
    End If

    If OptExpirLinc.value = True Then
        FrmEmpExpir3.show
    End If

    If OptExpirPas.value = True Then
        FrmEmpExpir1.show
    End If

End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 7
    End If

End Sub

Private Sub dcjopstatus_Click(Area As Integer)

    If val(Me.dcjopstatus.BoundText) = 1 Then
        CboWorkState.ListIndex = 0
   
    Else
        CboWorkState.ListIndex = 0
        '      Rs("workstate").Value = 1
    End If
    
End Sub

Private Sub Form_Activate()
    ShowDynamicHelp Me.HelpContextID
End Sub

Private Sub Form_Load()
    system_path = App.path
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from Nationality  "
    Else
        My_SQL = "  select  id,namee  from Nationality  "
    End If

    fill_combo DcNationality, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from dean  "
    Else
        My_SQL = "  select  id,namee  from dean  "
    End If

    fill_combo Dcdean, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from jopstatus  "
    Else
        My_SQL = "  select  id,namee  from jopstatus  "
    End If

    fill_combo dcjopstatus, My_SQL
 
    My_SQL = " select id,Project_name from projects"
 
    fill_combo dcproject, My_SQL

    My_SQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, My_SQL

    Dim Msg As String

    'Dim Dcombos As ClsDataCombos
    '
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        Msg = "Sales Commission"
        Msg = Msg & Chr(13) & "enter the values of the sales commission"
        Msg = Msg & " and it will calculted From the Sales Total Vaules and"
        Msg = Msg & " the Sales net profit values"
    
    Else
        Msg = "ﬁÌ„… «·⁄„Ê·… ⁄·Ï «·„»Ì⁄« "
        Msg = Msg & Chr(13) & " ≈–« ﬂ«‰ «·„ÊŸ› ÌÕ’· ⁄·Ï ⁄„Ê·… ⁄·Ï «·„»Ì⁄« "
        Msg = Msg & " ›√œŒ· ﬁÌ„… Â–Â «·⁄„Ê·… „⁄ «·√Œ– ›Ï «·√⁄ »«—"
        Msg = Msg & " √‰ «·»—‰«„Ã ÌÕ”» Â–Â «·⁄„Ê·… „‰ ﬁÌ„… ’«›Ï —»Õ «·›« Ê—…"
        Msg = Msg & "Ê√Ì÷« „‰ ≈Ã„«·Ï ﬁÌ„… «·›« Ê—…"
    End If

    'Me.lbl(6).Caption = Msg
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    Dcombos.GetEmpJobsTypes Me.DcboJobsType

    Dcombos.GetEmpSpecifications Me.DcboSpecifications
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches Me.dcBranch

    With Me.CboWorkState
        .Clear
        .AddItem "⁄·Ï ﬁÊ… «·⁄„·"
        .AddItem "›’· „‰ «·⁄„·"
    End With

    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.CboInsuranceState
            .Clear
            .AddItem "€Ì— „ƒ„‰ ⁄·ÌÂ"
            .AddItem "„ƒ„‰ ⁄·ÌÂ"
        End With

    Else

        With Me.CboInsuranceState
            .Clear
            .AddItem "Not have"
            .AddItem "  Have Insurance"
        End With

    End If

    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    'rs.Open "[TblEmployee]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Dim StrSQL As String
    StrSQL = "select * from  TblEmployee where not(DriverId is null) order by CAST(Emp_Code AS int)"
    'CAST(Emp_Code AS int)

    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
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
    Set EmpReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub ISButton1_Click()

    If XPTxtEmpID.text = "" Then Exit Sub
    x = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

    If x = vbYes Then
        DBPix201.ImageLoad

        DoEvents
        MsgBox " „  Õ„Ì· «·’Ê—…"
    Else

        If x = vbNo Then
            DBPix201.TWAINAcquire
            MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"

            DoEvents
        Else

            Exit Sub
        End If
    End If

    DBPix201.ImageSaveFile (system_path & "\images\" & XPTxtEmpID.text & ".JPG")
End Sub

Private Sub ISButton2_Click()

    If XPTxtEmpID.text = "" Then Exit Sub
    x = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

    If x = vbYes Then
        DBPix202.ImageLoad

        DoEvents
        MsgBox " „  Õ„Ì· «·’Ê—…"
    Else

        If x = vbNo Then
            DBPix202.TWAINAcquire
            MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"

            DoEvents
        Else

            Exit Sub
        End If
    End If

    DBPix202.ImageSaveFile (system_path & "\images\sign" & XPTxtEmpID.text & ".JPG")
End Sub

Private Sub NourHijriCal1_GotFocus()
    MsgBox NourHijriCal1.value
End Sub

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.text)
End Sub

Private Sub OptType1_Click(Index As Integer)
    Me.TxtOpenBalance1.Enabled = Not OptType1(2).value
    Me.TxtOpenBalance1.text = IIf(OptType1(2).value = True, 0, Me.TxtOpenBalance1.text)

End Sub

Private Sub OptType2_Click(Index As Integer)
    Me.TxtOpenBalance2.Enabled = Not OptType2(2).value
    Me.TxtOpenBalance2.text = IIf(OptType2(2).value = True, 0, Me.TxtOpenBalance2.text)

End Sub

Private Sub TxtEmp_Comm_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtEmp_Comm.text, 0)
End Sub

Private Sub TxtEmpProfitCom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtEmpProfitCom.text, 0)
End Sub

Private Sub TxtInsurValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtInsurValue.text, 0)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„ÊŸ›Ì‰"
            Else
                Me.Caption = "Employees Data"
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
            XPTxtEmpName.locked = True
            '        XPCboProfLevel.Locked = True
            XPTxtProfMail.locked = True
            XPTxtPhone.locked = True
            TxtSalary.locked = True
            TxtEmp_Code.locked = True
            XPTxtmobile.locked = True
            XPMTxtRemarks.locked = True
            Me.Txt_placEkama.locked = True
            Me.Txt_DateEndLinc.Enabled = False
            Me.Txt_DateEndekama.Enabled = False
            Me.Txt_DateEndpoket.Enabled = False
            Me.Txt_DateExpEkama.Enabled = False
            Me.Txt_DateExpLinc.Enabled = False
            Me.Txt_DateExppoket.Enabled = False
            Me.Txt_NumEkama.locked = True
            Me.Txt_NumLicn.locked = True
            Me.Txt_placEkama.locked = True
            Me.Tet_NumPoket.locked = True
            Me.Txt_NumPasp.locked = True
            Me.Txt_DateExpPasp.Enabled = False
            Me.Txt_DatePasp.Enabled = False
            Me.Chk_EndWork.Enabled = False
            Me.Chk_Stkala.Enabled = False
            '            Me.CmdEstkala.Enabled = False
            Me.DtDate.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            
            End If

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„ÊŸ›Ì‰( ”ÃÌ· ”Ã· ÃœÌœ)"
            Else
                Me.Caption = "Employees Data(Enter New Record)"
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
        
            XPTxtEmpName.locked = False
            '        XPCboProfLevel.Locked = False
            XPTxtProfMail.locked = False
            XPTxtPhone.locked = False
            XPTxtmobile.locked = False
            TxtSalary.locked = False
            XPMTxtRemarks.locked = False
            TxtEmp_Code.locked = False
            Me.Txt_placEkama.locked = False
            Me.Txt_DateEndLinc.Enabled = True
            Me.Txt_DateEndekama.Enabled = True
            Me.Txt_DateEndpoket.Enabled = True
            Me.Txt_DateExpEkama.Enabled = True
            Me.Txt_DateExpLinc.Enabled = True
            Me.Txt_DateExppoket.Enabled = True
            Me.Txt_NumEkama.locked = False
            Me.Txt_NumLicn.locked = False
            Me.Txt_placEkama.locked = False
            Me.Tet_NumPoket.locked = False
            Me.Txt_NumPasp.locked = False
            Me.Txt_DateExpPasp.Enabled = True
            Me.Txt_DatePasp.Enabled = True
            Me.Chk_EndWork.Enabled = True
            Me.Chk_Stkala.Enabled = True
            '            Me.CmdEstkala.Enabled = False
            Me.DtDate.Enabled = False

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„ÊŸ›Ì‰(  ⁄œÌ· )"
            Else
                Me.Caption = "Employees Data(Edit Current Record)"
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
            TxtSalary.locked = False
            XPTxtEmpName.locked = False
            '        XPCboProfLevel.Locked = False
            XPTxtProfMail.locked = False
            XPTxtPhone.locked = False
            XPTxtmobile.locked = False
            XPMTxtRemarks.locked = False
            TxtEmp_Code.locked = False
            Me.Txt_NumPasp.locked = False
            Me.Txt_DateExpPasp.Enabled = True
            Me.Txt_DatePasp.Enabled = True

            Me.Txt_placEkama.locked = False
            Me.Txt_DateEndLinc.Enabled = True
            Me.Txt_DateEndekama.Enabled = True
            Me.Txt_DateEndpoket.Enabled = True
            Me.Txt_DateExpEkama.Enabled = True
            Me.Txt_DateExpLinc.Enabled = True
            Me.Txt_DateExppoket.Enabled = True
            Me.Txt_NumEkama.locked = False
            Me.Txt_NumLicn.locked = False
            Me.Txt_placEkama.locked = False
            Me.Tet_NumPoket.locked = False
            Me.Chk_EndWork.Enabled = True
            Me.Chk_Stkala.Enabled = True
            '            Me.CmdEstkala.Enabled = False
            Me.DtDate.Enabled = False

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtOtherDiscounts_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOtherDiscounts.text, 0)
End Sub

Private Sub TxtSalary_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSalary.text, 0)
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

Function create_accounts() As Boolean

    If detect_employee_work_type = 0 Then
        create_accounts = True
        Exit Function
    End If

    Account_Code_dynamic = get_account_code_branch(7, my_branch)
        
    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  –„„ «·„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            create_accounts = False: Exit Function
         
        End If
    End If
        
    Account_Code_dynamic1 = get_account_code_branch(29, my_branch)
        
    If Account_Code_dynamic1 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic1 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·«ÃÊ— «·„” Õﬁ… ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
         
            create_accounts = False: Exit Function
        End If
    End If
        
    Account_Code_dynamic2 = get_account_code_branch(30, my_branch)
        
    If Account_Code_dynamic2 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic2 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „” Œ·’«  ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
         
            create_accounts = False: Exit Function
        End If
    End If
        
    Account_Code_dynamic3 = get_account_code_branch(65, my_branch)
        
    If Account_Code_dynamic3 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic3 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „œ›Ê⁄«  „ﬁœ„Â  «·„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            create_accounts = False: Exit Function
         
        End If
    End If
        
    If SystemOptions.CreateDriverBox = True Then
    
        Account_Code_dynamic4 = get_account_code_branch(6, my_branch)
                
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            create_accounts = False: Exit Function
        Else

            If Account_Code_dynamic = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··’‰«œÌﬁ   ›Ì «·›—⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                create_accounts = False: Exit Function
            End If
        End If
    End If
 
    If SystemOptions.CreateDriverEra = True Then
        Account_Code_dynamic5 = get_account_code_branch(35, my_branch)
                
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            create_accounts = False: Exit Function
        Else

            If Account_Code_dynamic = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··⁄Âœ   ›Ì «·›—⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                create_accounts = False: Exit Function
                 
            End If
        End If

    End If
        
    create_accounts = True
End Function

Public Sub Retrive(Optional Lngid As Long = 0)
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "Emp_ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtEmpID.text = IIf(IsNull(rs("Emp_ID").value), "", val(rs("Emp_ID").value))

    'OPENINGBALNCESDATA
    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)

    If Not (IsNull(rs("OpenBalanceDate").value)) Then
        Me.Dtp.value = rs("OpenBalanceDate").value
        Me.Dtp1.value = rs("OpenBalanceDate").value
        Me.Dtp2.value = rs("OpenBalanceDate").value
        ' Me.Dtp.Enabled = True
    Else
    
        Me.Dtp.value = Date
        '   Me.Dtp.Enabled = False
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

    If Not IsNull(rs("OpenBalanceType2").value) Then
        Me.TxtOpenBalance2.text = IIf(IsNull(rs("OpenBalance2")), "", Trim(rs("OpenBalance2")))

        If rs("OpenBalanceType2").value = 0 Then
            OptType2(0).value = True
            OptType2_Click 0
        ElseIf rs("OpenBalanceType2").value = 1 Then
            OptType2(1).value = True
            OptType2_Click 1
        End If
    
    Else
        Me.TxtOpenBalance2.text = 0
        Me.OptType2(2).value = True
        OptType2_Click 2
    End If

    If XPTxtEmpID.text <> "" Then
        DBPix201.ImageClear
        DBPix202.ImageClear

        If Dir(system_path & "\images\" & XPTxtEmpID.text & ".JPG") <> "" Then
            DBPix201.ImageLoadFile (system_path & "\images\" & XPTxtEmpID.text & ".JPG")
        End If

        If Dir(system_path & "\images\sign" & XPTxtEmpID.text & ".JPG") <> "" Then
            DBPix202.ImageLoadFile (system_path & "\images\sign" & XPTxtEmpID.text & ".JPG")
        End If
 
    End If

    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    Me.txtid.text = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)

    TxtEmp_Code.text = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
    XPTxtEmpName.text = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))
    Text1.text = IIf(IsNull(rs("Emp_Name1").value), "", Trim(rs("Emp_Name1").value))
    Text2.text = IIf(IsNull(rs("Emp_Name2").value), "", Trim(rs("Emp_Name2").value))
    Text3.text = IIf(IsNull(rs("Emp_Name3").value), "", Trim(rs("Emp_Name3").value))
    Text4.text = IIf(IsNull(rs("Emp_Name4").value), "", Trim(rs("Emp_Name4").value))

    XPTxtEmpNamee.text = IIf(IsNull(rs("Emp_Namee").value), "", Trim(rs("Emp_Namee").value))
    Text5.text = IIf(IsNull(rs("Emp_Namee1").value), "", Trim(rs("Emp_Namee1").value))
    Text6.text = IIf(IsNull(rs("Emp_Namee2").value), "", Trim(rs("Emp_Namee2").value))
    Text7.text = IIf(IsNull(rs("Emp_Namee3").value), "", Trim(rs("Emp_Namee3").value))
    Text8.text = IIf(IsNull(rs("Emp_Namee4").value), "", Trim(rs("Emp_Namee4").value))

    TxtAccountCode.text = IIf(IsNull(rs("Account_code").value), "", Trim(rs("Account_code").value))
    DcboCreditSide.BoundText = IIf(IsNull(rs("Account_code1").value), "", Trim(rs("Account_code1").value))

    txthdodno.text = IIf(IsNull(rs("hdodno").value), "", Trim(rs("hdodno").value))

    txthdomnfaz.text = IIf(IsNull(rs("hdomnfaz").value), "", Trim(rs("hdomnfaz").value))

    TxtSalary.text = IIf(IsNull(rs("Emp_Salary").value), "", Trim(rs("Emp_Salary").value))
    TXT_WORK_PLACE.text = IIf(IsNull(rs("placeWORK").value), "", Trim(rs("placeWORK").value))

    txtsaknm.text = IIf(IsNull(rs("Emp_Salary_sakn").value), "", Trim(rs("Emp_Salary_sakn").value))
    txtbusm.text = IIf(IsNull(rs("Emp_Salary_bus").value), "", Trim(rs("Emp_Salary_bus").value))

    txtfoodm.text = IIf(IsNull(rs("Emp_Salary_food").value), "", Trim(rs("Emp_Salary_food").value))
    TXTMOBM.text = IIf(IsNull(rs("Emp_Salary_mob").value), "", Trim(rs("Emp_Salary_mob").value))
    TXTMANGM.text = IIf(IsNull(rs("Emp_Salary_mang").value), "", Trim(rs("Emp_Salary_mang").value))
    txtanotherm.text = IIf(IsNull(rs("Emp_Salary_others").value), "", Trim(rs("Emp_Salary_others").value))

    txtsakn.text = IIf(IsNull(rs("Emp_Salary_sakn1").value), "", Trim(rs("Emp_Salary_sakn1").value))
    txtbus.text = IIf(IsNull(rs("Emp_Salary_bus1").value), "", Trim(rs("Emp_Salary_bus1").value))

    txtfood.text = IIf(IsNull(rs("Emp_Salary_food1").value), "", Trim(rs("Emp_Salary_food1").value))
    TXTMOB.text = IIf(IsNull(rs("Emp_Salary_mob1").value), "", Trim(rs("Emp_Salary_mob1").value))
    TXTMANG.text = IIf(IsNull(rs("Emp_Salary_mang1").value), "", Trim(rs("Emp_Salary_mang1").value))
    txtanother.text = IIf(IsNull(rs("Emp_Salary_others1").value), "", Trim(rs("Emp_Salary_others1").value))

    XPTxtProfMail.text = IIf(IsNull(rs("Emp_Mail").value), "", Trim(rs("Emp_Mail").value))
    XPTxtPhone.text = IIf(IsNull(rs("Emp_Phone").value), "", Trim(rs("Emp_Phone").value))
    XPTxtmobile.text = IIf(IsNull(rs("Emp_mobile").value), "", Trim(rs("Emp_mobile").value))
    XPMTxtRemarks.text = IIf(IsNull(rs("Emp_Remark").value), "", Trim(rs("Emp_Remark").value))
    TxtEmp_Comm.text = IIf(IsNull(rs("Emp_Comm").value), "", Trim(rs("Emp_Comm").value))
    TxtEmpProfitCom.text = IIf(IsNull(rs("EmpProfitCom").value), "", Trim(rs("EmpProfitCom").value))
    Txt_placEkama.text = IIf(IsNull(rs("placeEkama").value), "", Trim(rs("placeEkama").value))
    Txt_NumEkama.text = IIf(IsNull(rs("NumEkama").value), "", Trim(rs("NumEkama").value))
    Txt_NumLicn.text = IIf(IsNull(rs("NumLicn").value), "", Trim(rs("NumLicn").value))
    Tet_NumPoket.text = IIf(IsNull(rs("NumPoket").value), "", Trim(rs("NumPoket").value))

    Txt_NumPasp.text = IIf(IsNull(rs("NumPasp").value), "", Trim(rs("NumPasp").value))
    txtKafelID.text = IIf(IsNull(rs("KafelID").value), "", Trim(rs("KafelID").value))
    txtKafelName.text = IIf(IsNull(rs("KafelName").value), "", Trim(rs("KafelName").value))

    txtkafeltel.text = IIf(IsNull(rs("kafeltel").value), "", Trim(rs("kafeltel").value))

    txtkafeladd.text = IIf(IsNull(rs("kafeladd").value), "", Trim(rs("kafeladd").value))

    txtpasplace.text = IIf(IsNull(rs("pasplace").value), "", Trim(rs("pasplace").value))
    DcNationality.text = IIf(IsNull(rs("Nationality").value), "", Trim(rs("Nationality").value))
    Dcdean.text = IIf(IsNull(rs("dean").value), "", Trim(rs("dean").value))

    Txt_NotEndWork.text = IIf(IsNull(rs("Notsstkala").value), "", Trim(rs("Notsstkala").value))
   
    If rs("ChekStkala").value = True Then
        Chk_Stkala.value = Checked
    Else
        Chk_Stkala.value = Unchecked

    End If

    If rs("ChekEndWork").value = True Then

        Chk_EndWork.value = Checked
    Else
        Chk_EndWork.value = Unchecked
    End If

    DtDate.value = IIf(IsNull(rs("EndWork").value), Date, rs("EndWork").value)
    DTPicker1.value = IIf(IsNull(rs("BignDateWork").value), Date, rs("BignDateWork").value)
    DTPicker2.value = IIf(IsNull(rs("DOB").value), Date, rs("DOB").value)

    If IsNull(rs("workstate").value) Then
        Me.CboWorkState.ListIndex = -1
    Else

        If rs("workstate").value = 1 Then
            Me.CboWorkState.ListIndex = 0
        ElseIf rs("workstate").value = 0 Then
            Me.CboWorkState.ListIndex = 1
        End If
    End If

    Me.DcboEmpDepartments.BoundText = IIf(IsNull(rs("DepartmentID").value), "", rs("DepartmentID").value)
    Me.DcCostCenter.BoundText = IIf(IsNull(rs("cost_center_id").value), "", rs("cost_center_id").value)
    Me.dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    Me.DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
    Me.DcboSpecifications.BoundText = IIf(IsNull(rs("SpecificationID").value), "", rs("SpecificationID").value)
    Me.TxtRegion.text = IIf(IsNull(rs("Region").value), "", rs("Region").value)

    Me.dcjopstatus.BoundText = IIf(IsNull(rs("jopstatusid").value), "", rs("jopstatusid").value)
    Me.dcproject.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)

    If IsNull(rs("InsuranceState").value) Then
        Me.CboInsuranceState.ListIndex = 0
    Else
        Me.CboInsuranceState.ListIndex = rs("InsuranceState").value
    End If

    Me.TxtInsurValue.text = IIf(IsNull(rs("InsuranceValue").value), "", rs("InsuranceValue").value)
    Me.TxtOtherDiscounts.text = IIf(IsNull(rs("OtherDiscounts").value), "", rs("OtherDiscounts").value)

    Txt_DateExpLinc.value = IIf(IsNull(rs("DateExpLinc").value), Date, rs("DateExpLinc").value)
    Txt_DateEndLinc.value = IIf(IsNull(rs("DateEndLinc").value), Date, rs("DateEndLinc").value)

    Txt_DateExppoket.value = IIf(IsNull(rs("Dateexppoket").value), Date, rs("Dateexppoket").value)
    Txt_DateEndpoket.value = IIf(IsNull(rs("dateendpoket").value), Date, rs("dateendpoket").value)

    Txt_DateExpEkama.value = IIf(IsNull(rs("DateExpoekama").value), Date, rs("DateExpoekama").value)
    Txt_DateEndekama.value = IIf(IsNull(rs("DateEndekama").value), Date, rs("DateEndekama").value)

    Txt_DateExpEkamaH.value = IIf(IsNull(rs("DateExpoekamah").value), Date, rs("DateExpoekamah").value)
    Txt_DateEndekamah.value = IIf(IsNull(rs("DateEndekamah").value), Date, rs("DateEndekamah").value)

    Txt_DateExpLincH.value = IIf(IsNull(rs("DateExpLincH").value), Date, rs("DateExpLincH").value)
    Txt_DateEndLincH.value = IIf(IsNull(rs("DateEndLincH").value), Date, rs("DateEndLincH").value)

    Txt_DateExppoketH.value = IIf(IsNull(rs("Dateexppoketh").value), Date, rs("Dateexppoketh").value)
    Txt_DateEndpoketH.value = IIf(IsNull(rs("dateendpoketh").value), Date, rs("dateendpoketh").value)

    Txt_DateExpPasp.value = IIf(IsNull(rs("DateExpPasp").value), Date, rs("DateExpPasp").value)
    Txt_DatePasp.value = IIf(IsNull(rs("DateEndPasp").value), Date, rs("DateEndPasp").value)
    txthdoddate.value = IIf(IsNull(rs("hdoddate").value), Date, rs("hdoddate").value)

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Function CreateBoxAndEra(Boxname As String, BoxnameE As String, BoxAccount As String, EraAccount As String, EmpID As Integer, BranchId As Integer)
 
    Dim BoxID As Integer

    If Me.TxtModFlg.text = "N" Then

        If SystemOptions.CreateDriverBox = True Then
            BoxID = CStr(new_id("tblBoxesData", "BoxID", "", True))
            sql = "insert into  TblBoxesData (BoxID,BoxName,Account_Code,Type,empid,BranchId,BoxNameE,ChequeBox,driverid)"
            sql = sql & "Values(" & BoxID & ",'" & "Œ“Ì‰… " & Boxname & "','" & BoxAccount & "',0," & EmpID & "," & BranchId & ",'" & BoxnameE & "-Box" & "',0," & EmpID & ")"
                
            Cn.Execute sql
        End If

        If SystemOptions.CreateDriverEra = True Then
        
            BoxID = CStr(new_id("tblBoxesData", "BoxID", "", True))
            sql = "insert into  TblBoxesData (BoxID,BoxName,Account_Code,Type,empid,BranchId,BoxNameE,ChequeBox,driverid)"
            sql = sql & "Values(" & BoxID & ",'" & "⁄Âœ… " & Boxname & "','" & EraAccount & "',1," & EmpID & "," & BranchId & ",'" & BoxnameE & "-Era" & "',0," & EmpID & ")"
                
            Cn.Execute sql
        End If

    Else

        If SystemOptions.CreateDriverBox = True Then
                        
            sql = "update TblBoxesData set BoxName='" & "Œ“Ì‰… " & Boxname & "',BranchId=" & BranchId & ",BoxNameE='" & BoxnameE & "-Box" & "'"
            sql = sql & " where Account_Code='" & BoxAccount & "'"
                        
            Cn.Execute sql
        End If

        If SystemOptions.CreateDriverEra = True Then
            sql = "update TblBoxesData set BoxName='" & "⁄Âœ… " & Boxname & "',BranchId=" & BranchId & ",BoxNameE='" & BoxnameE & "-Era" & "'"
            sql = sql & " where Account_Code='" & EraAccount & "'"
                        
            Cn.Execute sql
        End If

    End If

End Function

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    'On Error GoTo ErrTrap
    XPTxtEmpName = Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text)

    If Me.TxtModFlg.text <> "R" Then

        '  If Text1.text = "" Then
        '      Msg = "ÌÃ» «œŒ«· «”„ «·„ÊŸ› "
        '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '      Text1.SetFocus
        '      SelectText Text1
        '      Exit Sub
        '     End If
    
        '      If Text2.text = "" Then
        '      Msg = "ÌÃ» «œŒ«· «”„ «·«» "
        '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '      Text2.SetFocus
        '      SelectText Text2
        '      Exit Sub
        '     End If
    
        '      If Text3.text = "" Then
        'If SystemOptions.UserInterface = ArabicInterface Then
        '       Msg = "ÌÃ» «œŒ«· «”„ «·Ãœ "
        'Else
        'Msg = "Enter Grand Father Name"
        'End If
        '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '       Text3.SetFocus
        '       SelectText Text3
        '       Exit Sub
        '      End If
        '
        '    If Text4.text = "" Then
        '      Msg = "ÌÃ» «œŒ«· «”„ «·⁄«∆·… "
        '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '      Text4.SetFocus
        '      SelectText Text4
        '      Exit Sub
        '     End If
    
        '    If TxtEmp_Code.text = "" Then
        '  If SystemOptions.UserInterface = ArabicInterface Then
        '        Msg = "ÌÃ» «œŒ«· ﬂÊœ «·„ÊŸ› "
        '    Else
        '    Msg = "Enter employee Code "
        '    End If
        '        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        ''        TxtEmp_Code.SetFocus
        '        SelectText TxtEmp_Code
        '        Exit Sub
        '    End If
    
        If Not IsNumeric(TxtSalary.text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» «œŒ«· «·—« » «·«”«”Ì ··„ÊŸ›  "
            Else
                Msg = " Enter Basic Salary Value  "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtSalary.SetFocus
            SelectText TxtSalary
            Exit Sub
        End If
    
        If DcboEmpDepartments.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡  ÕœÌœ «·ﬁ”„ «·–Ì Ì »⁄Â «·„ÊŸ›"
            Else
                Msg = " Specify Branch"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '  CboWorkState.SetFocus
        
            DcboEmpDepartments.SetFocus
            SendKeys "{F4}"
        
            Exit Sub
        End If
    
        If DcboJobsType.BoundText = "" Then
    
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡  ÕœÌœ ÊŸÌ›… «·„ÊŸ›"
            Else
                Msg = " Specify Job type"
            End If
        
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
            DcboJobsType.SetFocus
            SendKeys "{F4}"
        
            Exit Sub
        End If
    
        If CboWorkState.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡  ÕœÌœ Õ«·… «·„ÊŸ› (Â· ⁄·Ï ﬁÊ… «·⁄„· √Ê  „ ›’·Â „‰ «·⁄„·)"
            Else
                Msg = " Specify Job Status"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '  CboWorkState.SetFocus
        
            dcjopstatus.SetFocus
            SendKeys "{F4}"
        
            Exit Sub
        End If
    
        If val(Me.TxtEmp_Comm.text) > 0 Then
            If val(Me.TxtEmp_Comm.text) >= 100 Or val(Me.TxtEmp_Comm.text) < 0 Then
                Msg = "ﬁÌ„… ⁄„Ê·… «·„ÊŸ› €Ì— ’ÕÌÕ…..!!"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtEmp_Comm.SetFocus
                SelectText TxtEmp_Comm
                Exit Sub
            End If
        End If

        StrVacCode = IsRecExist("TblEmployee", "Emp_code", Trim(TxtEmp_Code.text), "Emp_Name", "Emp_ID<>" & val(XPTxtEmpID.text))

        If StrVacCode <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·ﬂÊœ „‰ ﬁ»·"
            Else
                Msg = " Emp Code Already Exist"
            End If
        
            MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
            '        TxtEmp_Code.SetFocus
            SelectText TxtEmp_Code
            Exit Sub
        End If

        '    If Txtsalary.text <> "" Then
        '        If Not (IsNumeric(Txtsalary.text)) Then
        '
        '
        '                If SystemOptions.UserInterface = ArabicInterface Then
        '                Msg = "«·„— » ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì… "
        '            Else
        '            Msg = " Emp  Salary Not Correct  "
        '            End If
        '
        '            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '            Txtsalary.SetFocus
        '            SelectText Txtsalary
        '            Exit Sub
        '        End If
        '    End If
 
        If detect_employee_work_type = 1 Then
            If Me.OptType(2).value = False Then
                If val(Me.TxtOpenBalance.text) = 0 Then
                    Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ...!!!"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

                    If TxtOpenBalance.Enabled = True Then
                        TxtOpenBalance.SetFocus
                    End If

                    Exit Sub
                End If
            End If
    
        End If

        Select Case TxtModFlg.text

            Case "N"

                '   StrSQL = "select * From TblEmployee where Emp_Name='" & Trim(XPTxtEmpName.text) & "'"
                '   RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                '   If RsTemp.RecordCount > 0 Then
                '       Msg = "ÌÊÃœ „ÊŸ› „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                '       Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & Chr(13)
                '       Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                '       MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                '       Exit Sub
                '   End If
            Case "E"
                '   StrSQL = "select * From TblEmployee where Emp_Name='" & Trim(XPTxtEmpName.text) & "'"
                '   RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                '   If RsTemp.RecordCount > 0 Then
                '       If RsTemp("Emp_ID").value <> Val(XPTxtEmpID) Then
                '           Msg = "ÌÊÃœ „ÊŸ› „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                '           Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & Chr(13)
                '           Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                '           MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                '           Exit Sub
                '       End If
                '   End If
        End Select

        If create_accounts = False Then
            Exit Sub
        End If
     
        Cn.BeginTrans
        BeginTrans = True
     
        If TxtModFlg.text = "N" Then
     
            XPTxtEmpID.text = CStr(new_id("TblEmployee", "Emp_ID", "", True))
            Me.TxtEmp_Code.text = CStr(new_id("TblEmployee", "Emp_Code", "", True))
        
            rs.AddNew
            rs("Emp_ID").value = val(XPTxtEmpID.text)
            rs("DriverId").value = val(XPTxtEmpID.text)
        
            If detect_employee_work_type = 1 Then
   
                rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "- –„„ ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & "  ") '–„„
          
                rs("Account_Code1").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "- «ÃÊ— „” Õﬁ…", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & "- Salary  ") '–„„) '
                '«ÃÊ— „” Õﬁ…
                rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic2, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’«  ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Reserved ") '–„„) '„Œ’’« 
        
                rs("Account_Code3").value = ModAccounts.AddNewAccount(Account_Code_dynamic3, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„œ›Ê⁄«  „ﬁœ„…  ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Adv. Payments ") '   „œ›Ê⁄«  „ﬁœ„Â
               
                If SystemOptions.CreateDriverBox = True Then
                    rs("Account_Code4").value = ModAccounts.AddNewAccount(Account_Code_dynamic4, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "  Œ“Ì‰…   ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Box ") '   Œ“‰
                End If

                If SystemOptions.CreateDriverEra = True Then
                    rs("Account_Code5").value = ModAccounts.AddNewAccount(Account_Code_dynamic5, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "  ⁄Âœ…   ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Era ") '     ⁄Âœ…
                End If
                     
            End If
        
        Else
        
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             
            '  Rs("Account_Code").value = ModAccounts.AddNewAccount("a1a2a6", Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text), True, False)
        End If
    
        rs("Emp_Code").value = txtid.text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtid.text) = "", Null, txtid.text)
        rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)

        '  rs("Emp_Code").value = IIf(TxtEmp_Code.text = "", Null, Trim(TxtEmp_Code.text))
        '   Rs("Emp_Name").value = Trim(XPTxtEmpName.text)
        rs("Emp_Name").value = Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text)
    
        rs("Emp_Name1").value = Trim(Text1.text)
        rs("Emp_Name2").value = Trim(Text2.text)
        rs("Emp_Name3").value = Trim(Text3.text)
        rs("Emp_Name4").value = Trim(Text4.text)
    
        rs("Emp_Namee").value = Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text)
    
        rs("Emp_Namee1").value = Trim(Text5.text)
        rs("Emp_Namee2").value = Trim(Text6.text)
        rs("Emp_Namee3").value = Trim(Text7.text)
        rs("Emp_Namee4").value = Trim(Text8.text)

        If detect_employee_work_type = 1 Then
            If IsNull(rs("Account_Code").value) Or rs("Account_Code").value = "" Then
                rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "     ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Adv. Payments ") '   „œ›Ê⁄«  „ﬁœ„Â
          
            Else
            
                ModAccounts.EditAccount rs("Account_Code").value, rs("Emp_Name").value, rs("Emp_Namee").value, , , , , , , , , , , , , , , , , True
            End If
            
            If IsNull(rs("Account_Code1").value) Or rs("Account_Code1").value = "" Then
                rs("Account_Code1").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "«ÃÊ— „” Õﬁ…    ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Adv. Payments ") '   „œ›Ê⁄«  „ﬁœ„Â
          
            Else
                ModAccounts.EditAccount rs("Account_Code1").value, rs("Emp_Name").value & "  «ÃÊ— „”‰Õﬁ… ", rs("Emp_Namee").value & " Salary ", , , , , , , , , , , , , , , , , True
            End If
            
            If IsNull(rs("Account_Code2").value) Or rs("Account_Code2").value = "" Then
                rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’«      ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Adv. Payments ") '   „œ›Ê⁄«  „ﬁœ„Â
          
            Else
            
                ModAccounts.EditAccount rs("Account_Code2").value, rs("Emp_Name").value & "  „Œ’’« ", rs("Emp_Namee").value & "  Reserved ", , , , , , , , , , , , , , , , , True
            End If
            
            If IsNull(rs("Account_Code3").value) Or rs("Account_Code3").value = "" Then
                rs("Account_Code3").value = ModAccounts.AddNewAccount(Account_Code_dynamic3, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„œ›Ê⁄«  „ﬁœ„…  ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Adv. Payments ") '   „œ›Ê⁄«  „ﬁœ„Â
      
            Else
                ModAccounts.EditAccount rs("Account_Code3").value, rs("Emp_Name").value & "  „œ›Ê⁄«  „ﬁœ„Â ", rs("Emp_Namee").value & "  Adv. Payments ", , , , , , , , , , , , , , , , , True
            End If
            
            If SystemOptions.CreateDriverBox = True Then
                          
                If IsNull(rs("Account_Code4").value) Or rs("Account_Code4").value = "" Then
                    rs("Account_Code4").value = ModAccounts.AddNewAccount(Account_Code_dynamic4, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "Œ“Ì‰…    ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & "-Box") '   Œ“Ì‰…
                    
                Else
                    ModAccounts.EditAccount rs("Account_Code4").value, rs("Emp_Name").value & "  Œ“Ì‰…   ", rs("Emp_Namee").value & "  Box ", , , , , , , , , , , , , , , , , True
                End If
            
            End If
            
            If SystemOptions.CreateDriverEra = True Then
   
                If IsNull(rs("Account_Code5").value) Or rs("Account_Code5").value = "" Then
                    rs("Account_Code5").value = ModAccounts.AddNewAccount(Account_Code_dynamic5, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "⁄Âœ…    ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & "-Era") '   ⁄Âœ…
                        
                Else
                    ModAccounts.EditAccount rs("Account_Code5").value, rs("Emp_Name").value & "  ⁄Âœ…   ", rs("Emp_Namee").value & "  Era ", , , , , , , , , , , , , , , , , True
                End If
            
            End If
              
        End If
   
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
    
        If Me.OptType2(2).value = True Then
            rs("OpenBalance2").value = 0
            rs("OpenBalanceType2").value = Null
        ElseIf Me.OptType2(0).value = True Then
            rs("OpenBalance2").value = val(Me.TxtOpenBalance2.text)
            rs("OpenBalanceType2").value = 0
        ElseIf Me.OptType2(1).value = True Then
            rs("OpenBalance2").value = val(Me.TxtOpenBalance2.text)
            rs("OpenBalanceType2").value = 1
        End If
    
        rs("OpenBalanceDate").value = Me.Dtp.value
            
        If detect_employee_work_type = 1 Then
    
            If val(TxtOpenBalance.text) <> 0 Or val(TxtOpenBalance1.text) <> 0 Or val(TxtOpenBalance2.text) <> 0 Then
                txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
                rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
            Else
                rs("opening_balance_voucher_id").value = Null
            End If
    
        End If

        '    Rs("Account_Code1").value = DcboCreditSide.BoundText
    
        rs("hdodno").value = Trim(txthdodno.text)
    
        rs("hdoddate").value = txthdoddate.value
    
        rs("hdomnfaz").value = Trim(txthdomnfaz.text)
    
        rs("Emp_Salary").value = IIf(TxtSalary.text = "", Null, Trim(TxtSalary.text))
        rs("placeWORK").value = IIf(TXT_WORK_PLACE.text = "", Null, Trim(TXT_WORK_PLACE.text))
    
        rs("Emp_Salary_sakn").value = IIf(txtsaknm.text = "", Null, val(txtsaknm.text))
        rs("Emp_Salary_bus").value = IIf(txtbusm.text = "", Null, val(txtbusm.text))
    
        rs("Emp_Salary_food").value = IIf(txtfoodm.text = "", Null, val(txtfoodm.text))
        rs("Emp_Salary_mob").value = IIf(TXTMOBM.text = "", Null, val(TXTMOBM.text))
        rs("Emp_Salary_mang").value = IIf(TXTMANGM.text = "", Null, val(TXTMANGM.text))
        rs("Emp_Salary_others").value = IIf(txtanotherm.text = "", Null, val(txtanotherm.text))
    
        rs("Emp_Salary_sakn1").value = IIf(txtsakn.text = "", Null, val(txtsakn.text))
        rs("Emp_Salary_bus1").value = IIf(txtbus.text = "", Null, val(txtbus.text))
    
        rs("Emp_Salary_food1").value = IIf(txtfood.text = "", Null, val(txtfood.text))
        rs("Emp_Salary_mob1").value = IIf(TXTMOB.text = "", Null, val(TXTMOB.text))
        rs("Emp_Salary_mang1").value = IIf(TXTMANG.text = "", Null, val(TXTMANG.text))
        rs("Emp_Salary_others1").value = IIf(txtanother.text = "", Null, val(txtanother.text))
    
        'Emp_Salary_sakn
    
        rs("Emp_Mail").value = IIf(XPTxtProfMail.text = "", "", Trim(XPTxtProfMail.text))
        rs("Emp_Phone").value = IIf(XPTxtPhone.text = "", "", Trim(XPTxtPhone.text))
        rs("Emp_mobile").value = IIf(XPTxtmobile.text = "", "", Trim(XPTxtmobile.text))
        rs("Emp_Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("Emp_Comm").value = IIf(TxtEmp_Comm.text = "", 0, val(TxtEmp_Comm.text))
        rs("EmpProfitCom").value = IIf(TxtEmpProfitCom.text = "", 0, val(TxtEmpProfitCom.text))
        rs("placeEkama").value = IIf(Txt_placEkama.text = "", Null, Trim(Txt_placEkama.text))
        rs("NumEkama").value = IIf(Txt_NumEkama.text = "", Null, Trim(Txt_NumEkama.text))
    
        rs("DateExpoekamah").value = Txt_DateExpEkamaH.value
    
        rs("DateEndekamah").value = Txt_DateEndekamah.value
    
        rs("DateExpoekama").value = ToGregorianDate(Txt_DateExpEkamaH.value)
        rs("DateEndekama").value = ToGregorianDate(Txt_DateEndekamah.value)
    
        rs("DateExpLincH").value = Txt_DateExpLincH.value
        rs("DateEndLincH").value = Txt_DateEndLincH.value
    
        rs("Dateexppoketh").value = Txt_DateExppoketH.value
        rs("dateendpoketh").value = Txt_DateEndpoketH.value
     
        rs("Dateexppoket").value = ToGregorianDate(Txt_DateExppoketH.value) ' Txt_DateExppoket.value
        rs("dateendpoket").value = ToGregorianDate(Txt_DateEndpoketH.value) 'Txt_DateEndpoket.value
     
        rs("DateExpLinc").value = ToGregorianDate(Txt_DateExpLincH.value)
        rs("DateEndLinc").value = ToGregorianDate(Txt_DateEndLincH.value)
    
        rs("NumLicn").value = IIf(Txt_NumLicn.text = "", Null, Trim(Txt_NumLicn.text))
        'Rs("DateExpLinc").value = Txt_DateExpLinc.value
        'Rs("DateEndLinc").value = Txt_DateEndLinc.value
        rs("NumPoket").value = IIf(Tet_NumPoket.text = "", Null, Trim(Tet_NumPoket.text))

        rs("NumPasp").value = IIf(Txt_NumPasp.text = "", Null, Trim(Txt_NumPasp.text))
        rs("KafelID").value = IIf(txtKafelID.text = "", Null, Trim(txtKafelID.text))
        rs("KafelName").value = IIf(txtKafelName.text = "", Null, Trim(txtKafelName.text))
     
        rs("kafeltel").value = IIf(txtkafeltel.text = "", Null, Trim(txtkafeltel.text))
        rs("kafeladd").value = IIf(txtkafeladd.text = "", Null, Trim(txtkafeladd.text))
       
        rs("pasplace").value = IIf(txtpasplace.text = "", Null, Trim(txtpasplace.text))
        rs("Nationality").value = IIf(DcNationality.text = "", Null, Trim(DcNationality.text))
        rs("dean").value = IIf(Dcdean.text = "", Null, Trim(Dcdean.text))
        rs("project_id").value = IIf(dcproject.text = "", Null, dcproject.BoundText)
        rs("BranchId").value = IIf(Me.dcBranch.text = "", Null, Me.dcBranch.BoundText)
   
        rs("DateExpPasp").value = Txt_DateExpPasp.value
        rs("DateEndPasp").value = Txt_DatePasp.value
        '    Rs("Notsstkala").Value = IIf(Txt_NotEndWork.text = "", "", Trim(Txt_NotEndWork.text))
        rs("Notsstkala").value = IIf(Txt_NotEndWork.text = "", "", Trim(Txt_NotEndWork.text))

        If Chk_Stkala.value = Checked Then
            rs("ChekStkala").value = 1
        Else
            rs("ChekStkala").value = 0
        End If
    
        If Chk_EndWork.value = Checked Then
            rs("ChekEndWork").value = 1
        Else
            rs("ChekEndWork").value = 0
        End If
    
        If Chk_Stkala.value = Checked Or Chk_EndWork.value = Checked Then
            rs("EndWork").value = DtDate.value
        Else
            rs("EndWork").value = Null
        End If
 
        rs("BignDateWork").value = DTPicker1.value
        rs("DOB").value = DTPicker2.value

        '  If Me.CboWorkState.ListIndex = 0 Then
        '      Rs("workstate").value = 1
        '  ElseIf Me.CboWorkState.ListIndex = 1 Then
        '      Rs("workstate").value = 0
        '  End If
    
        If val(Me.dcjopstatus.BoundText) = 1 Then
            rs("workstate").value = 1
   
        Else
            rs("workstate").value = 0
        End If
    
        If val(Me.dcjopstatus.BoundText) = 0 Then
            rs("jopstatusid").value = Null
        Else
            rs("jopstatusid").value = val(Me.dcjopstatus.BoundText)
        End If
    
        If val(Me.DcboEmpDepartments.BoundText) = 0 Then
            rs("DepartmentID").value = Null
        Else
            rs("DepartmentID").value = val(Me.DcboEmpDepartments.BoundText)
        End If
    
        If Me.DcCostCenter.BoundText = "" Then
            rs("cost_center_id").value = Null
        Else
            rs("cost_center_id").value = Me.DcCostCenter.BoundText
        End If
    
        If val(Me.DcboJobsType.BoundText) = 0 Then
            rs("JobTypeID").value = Null
        Else
            rs("JobTypeID").value = val(Me.DcboJobsType.BoundText)
        End If

        If val(Me.DcboSpecifications.BoundText) = 0 Then
            rs("SpecificationID").value = Null
        Else
            rs("SpecificationID").value = val(Me.DcboSpecifications.BoundText)
        End If

        rs("Region").value = Trim$(Me.TxtRegion.text)

        If Me.CboInsuranceState.ListIndex = 0 Or Me.CboInsuranceState.ListIndex = -1 Then
            rs("InsuranceState").value = 0
        ElseIf Me.CboInsuranceState.ListIndex = 1 Then
            rs("InsuranceState").value = 1
        End If

        rs("InsuranceValue").value = val(Me.TxtInsurValue.text)
        rs("OtherDiscounts").value = val(Me.TxtOtherDiscounts.text)
    
        '    If Dir(system_path & "\images\" & XPTxtEmpID.text & ".JPG") <> "" Then
        '     Rs("ItemPhoto").value = DBPix201.ImageLoadFile(system_path & "\images\" & XPTxtEmpID.text & ".JPG")
 
        '    End If

        'OPENING Balance Voucher
        Dim StrDes As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtEmpName.text) & " "
        Else
            StrDes = " Opening Balance For: " & Trim(Me.XPTxtEmpNamee.text) & " "
        End If
        
        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
                Dim LngDevID As Long
                Dim LngOpenID As Long

                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
                If Me.OptType(0).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(61, my_branch)
        
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                     
                        End If
                    End If
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                ElseIf Me.OptType(1).value = True Then
            
                    Account_Code_dynamic1 = get_account_code_branch(61, my_branch)
                 
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                  
                        End If
                    End If
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                End If
                 
            End If
        End If

        If Me.OptType1(0).value = True Or Me.OptType1(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
                If Me.OptType1(0).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(61, my_branch)
        
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                     
                        End If
                    End If
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code1").value, val(Me.TxtOpenBalance1.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance1.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                ElseIf Me.OptType1(1).value = True Then
            
                    Account_Code_dynamic1 = get_account_code_branch(61, my_branch)
                 
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                  
                        End If
                    End If
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance1.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code1").value, val(Me.TxtOpenBalance1.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                End If
                 
            End If
        End If
  
        '33333333333333333333
        If Me.OptType2(0).value = True Or Me.OptType2(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
                If Me.OptType2(0).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(61, my_branch)
        
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                     
                        End If
                    End If
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code2").value, val(Me.TxtOpenBalance2.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance2.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                ElseIf Me.OptType2(1).value = True Then
            
                    Account_Code_dynamic1 = get_account_code_branch(61, my_branch)
                 
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                  
                        End If
                    End If
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance2.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code2").value, val(Me.TxtOpenBalance2.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                        GoTo ErrTrap
                    End If
                End If
                 
            End If
        End If
  
        CreateBoxAndEra rs("Emp_Name").value, rs("Emp_NameE").value, IIf(IsNull(rs("Account_Code4").value), "", rs("Account_Code4").value), IIf(IsNull(rs("Account_Code5").value), "", rs("Account_Code5").value), val(XPTxtEmpID), val(Me.dcBranch.BoundText)

        rs.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

    End If

    ' ⁄œÌ· „—ﬂ“ «· ﬂ·›…

    If Me.DcCostCenter.BoundText = "" Then
        Dim x As Boolean
        x = UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code"), False, 0, "")
        x = UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code1"), False, 0, "")
        x = UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code2"), False, 0, "")
   
    Else

        If UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code"), True, 1, DcCostCenter.BoundText) = False Then
             
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õ”«»  " & "  –„„ «·„ÊŸ›Ì‰  " & "·Â–« «·„ÊŸ› €Ì— „ÊÃÊœ Ê·„ Ì „  ⁄œÌ· „—ﬂ“ «· ﬂ·›… ·Â  "
            Else
                Msg = "staff Accounts  " & " Account " & "Not Defined to this Employee and cost center not be updated"
            End If

            MsgBox Msg, vbCritical
            Exit Sub
            
        End If
            
        If UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code1"), True, 1, DcCostCenter.BoundText) = False Then
             
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õ”«»" & "  «·«ÃÊ— «·„” Õﬁ…  " & "·Â–« «·„ÊŸ› €Ì— „ÊÃÊœ Ê·„ Ì „  ⁄œÌ· „—ﬂ“ «· ﬂ·›… ·Â"
            Else
                Msg = "Due salaries Acc  " & "Account" & "  Not Defined to this Employee and cost center not be updated"
            End If

            MsgBox Msg, vbCritical
            Exit Sub
            
        End If

        If UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code2"), True, 1, DcCostCenter.BoundText) = False Then
             
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õ”«»" & " «·„Œ’’«  " & "·Â–« «·„ÊŸ› €Ì— „ÊÃÊœ Ê·„ Ì „  ⁄œÌ· „—ﬂ“ «· ﬂ·›… ·Â"
            Else
                Msg = "Apportionment " & "Account" & " Not Defined to this Employee and cost center not be updated"
            End If

            MsgBox Msg, vbCritical
            Exit Sub
            
        End If
       
    End If

    ' ÕœÌÀ «·—« » «·«”«”Ì ›Ì ⁄ﬁœ «·„ÊŸ›
    updateEmployeeSalaryComponent val(Me.XPTxtEmpID.text), Me.TxtSalary.text
    ' ÕœÌÀ «·„›—œ«  «·«·Ì…
    addSalaryComponentToEmployee val(Me.XPTxtEmpID.text)
 
    Select Case Me.TxtModFlg.text

        Case "N"
            updateResults

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„ÊŸ› " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
      
            Else
                Msg = " This Employee Data Was Saved" & Chr(13)
                Msg = Msg + "Do you want To enter Another Employee"
            End If
  
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub

            End If
        
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Amendments have been saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

            updateResults
 
    End Select

    rs.Close
    'rs.Open "[TblEmployee]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "select * from  TblEmployee where not(DriverId is null) order by CAST(Emp_Code AS int)"
    'CAST(Emp_Code AS int)

    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    Me.Retrive Me.XPTxtEmpID
       
    TxtModFlg.text = "R"
    
    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        Cn.RollbackTrans
        BeginTrans = False
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Function updateResults()
    '           rs.Close
    '        rs.Open "[TblEmployee]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '       Me.Retrive Me.XPTxtEmpID
End Function
       
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "Emp_ID='" & val(XPTxtEmpID.text) & "'", , adSearchForward, adBookmarkFirst

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

Public Function updateEmployeeSalaryComponent(Emp_id As Integer, _
                                              salary As Double)
    Exit Function
    Dim sql As String
    Dim rs As ADODB.Recordset
    sql = "update Contract set Basic_salary=" & salary & "where Emp_id =" & Emp_id
    Cn.Execute sql
    sql = "Select * From EmpSalaryComponent where emp_ID=" & Emp_id
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount

            If rs("is_fixed") = 1 Then
                rs("value") = val(rs("specific_value"))
            Else
                rs("value") = cal_value(rs("eq_text"))
            
            End If
                 
            If rs("value") < val(rs("min_val")) And val(rs("min_val")) > 0 Then
                rs("value") = rs("min_val")
            ElseIf rs("value") > val(rs("max_val")) And val(rs("max_val")) > 0 Then
                rs("value") = rs("max_val")
            End If

            rs.update
            rs.MoveNext
        Next i

    End If

End Function

Public Function get_value(operand As String) As Double
    operand = Replace$(operand, "A", "")
    Dim value As Double
    Dim mofrad_count As Integer
    mofrad_count = 0

    If operand = 1 Then
        If IsNumeric(Me.TxtSalary.text) Then
            get_value = 1 * val(TxtSalary.text)
            Exit Function
        Else
            get_value = 0
            MsgBox "·„ Ì „  ÕœÌœ ﬁÌ„… «·—« » «·«”«”Ì »—Ã«¡  ÕœÌœÂ«"
            Exit Function
        End If

    End If

    Dim sql As String
    Dim rs As ADODB.Recordset
 
    sql = "Select * From EmpSalaryComponent where emp_ID=" & val(Me.XPTxtEmpID)
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount

            If rs("AccountCode").value = operand Then
                mofrad_count = mofrad_count + 1
              
            End If

            rs.MoveNext
        Next i

    End If

    If mofrad_count = 0 Then
        MsgBox "«·„›—œ €Ì— „ÊÃÊœ"
        Exit Function
    ElseIf mofrad_count > 1 Then
        MsgBox "«·„›—œ    „Õœœ «ﬂÀ— „‰ „—…"
        Exit Function
    End If

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount

            If rs("AccountCode").value = operand Then
                get_value = rs("value").value
                Exit Function
              
            End If

            rs.MoveNext
        Next i

    End If
 
End Function

Public Function cal_value(src As String) As Double
    'On Error GoTo errortrap
    Dim new_pos As Integer
    Dim last_pos As Integer
    Dim cuttent_operand As String
    Dim new_str As String
    Dim objScript As Object
    last_pos = 1
    new_str = ""

    For i = 1 To Len(src)

        If Mid(src, i, 1) = "+" Or Mid(src, i, 1) = "-" Or Mid(src, i, 1) = "*" Or Mid(src, i, 1) = "/" Or Mid(src, i, 1) = "=" Then
            new_pos = i
            cuttent_operand = Mid(src, last_pos, new_pos - last_pos)

            If InStr(cuttent_operand, "A") > 0 Then
                cuttent_operand = get_value(cuttent_operand)
            End If

            new_str = new_str & cuttent_operand & Mid(src, i, 1)

            If i < Len(src) Then
                last_pos = new_pos + 1
            Else
                GoTo ll
            End If
        End If
 
    Next i

ll:
    new_str = Replace$(new_str, "=", "")

    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"
 
    cal_value = objScript.Eval(new_str)
    Exit Function

errortrap:
    cal_value = 0

End Function

Function DeleteOpeningBalance()
    Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.text = 0

    'OptType1(2).value = True
    'TxtOpenBalance1.text = 0

    'OptType2(2).value = True
    'TxtOpenBalance2.text = 0

    Cmd_Click (2)

End Function

Private Sub Del_ProfData()

    Dim Msg As String
    Dim StrSQL As String

    'On Error GoTo ErrTrap
    DeleteOpeningBalance
    StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
             
    If check_employee_transations(val(XPTxtEmpID)) = False Then

        Exit Sub

    End If

    If XPTxtEmpID.text <> "" Then

        Msg = "”Ì „ Õ–› »Ì«‰«  «·„ÊŸ› —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtEmpID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
             
                rs.delete
                rs.MoveFirst
            
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„ÊŸ› "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Function check_employee_transations(Emp_id As String) As Boolean
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    check_employee_transations = True
 
    StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code") & "' or  Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code1") & "' or Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code2") & "' or Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code3") & "' or Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code4") & "' or Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code5") & "'"
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & Chr(13)
        Msg = Msg + "·«‰… „”Ã· ›Ì »⁄÷ «·ﬁÌÊœ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        check_employee_transations = False
        Exit Function
    End If
    
    If ModAccounts.DeleteAccount(get_EMPLOYEE_Account(Emp_id, "Account_Code")) = True Then
        check_employee_transations = True

    End If

    If ModAccounts.DeleteAccount(get_EMPLOYEE_Account(Emp_id, "Account_Code1")) = True Then
        check_employee_transations = True

    End If

    If ModAccounts.DeleteAccount(get_EMPLOYEE_Account(Emp_id, "Account_Code2")) = True Then
        check_employee_transations = True

    End If
            
    If ModAccounts.DeleteAccount(get_EMPLOYEE_Account(Emp_id, "Account_Code3")) = True Then
        check_employee_transations = True

    End If
            
    If ModAccounts.DeleteAccount(get_EMPLOYEE_Account(Emp_id, "Account_Code4")) = True Then
        check_employee_transations = True
        Cn.Execute "Delete TblBoxesData where  Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code4") & "'"
    End If
            
    If ModAccounts.DeleteAccount(get_EMPLOYEE_Account(Emp_id, "Account_Code5")) = True Then
        check_employee_transations = True
        Cn.Execute "Delete TblBoxesData where  Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code5") & "'"
    End If
            
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            KeyCode = 0
            SendKeys "{TAB}"
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

    If Shift = VBRUN.ShiftConstants.vbShiftMask Then
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
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „ÊŸ› ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·„ÊŸ›" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„ÊŸ› «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  „ÊŸ›" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „ÊŸ›" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap, True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New Record ..." & Wrap & "Click here to add a new employee" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the current employee data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the new record or " & Wrap & "save the edit in the " & Wrap & "current record", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo" & Wrap & "Undo in the adding new record" & Wrap & "Or undo in the current editing" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current employee data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Search for an employee" & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist Record" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next" & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last" & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help" & Wrap & "Show the Help File" & Wrap & "" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub printingReport()
    Dim sql As String
    
    'Dim Rs As ADODB.Recordset
    Dim xReport As New CRAXDRT.Report
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim reportpatath As String

    sql = "select * From emp_all_details ORDER BY CAST(Emp_Code AS integer) ASC "
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
   
    If SystemOptions.UserInterface = ArabicInterface Then
        reportpatath = system_path & "\reports\emp\REPORT9.rpt"
    Else
        reportpatath = system_path & "\reports\emp\REPORT9e.rpt"
    End If

    Set xReport = xApp.OpenReport(reportpatath)
    xReport.Database.SetDataSource rs
 
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
  
    FrmReport.CRViewer.viewReport
    FrmReport.TxtPath = reportpatath
    FrmReport.show
    Screen.MousePointer = vbDefault
    '      xReport.ReportTitle = X
    SendKeys "{RIGHT}"

    'Dim Msg As String
    'On Error GoTo ErrTrap
    'If XPTxtEmpID.text <> "" Then
    '    Set EmpReport = New ClsEmployeeReport
    '    EmpReport.EmpData XPTxtEmpID.text
    'Else
    '    Msg = "⁄„·Ì… «·ÿ»«⁄… €Ì— „ «Õ… Õ«·Ì«"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
     
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub ChangeLang()
    XPLbl(46).Caption = "Work Place"
    SuperLabel1(5).text = "Project"
    lblb(9).text = "Branch"
    lbl(31).Caption = "Credit"
    lbl(6).Visible = False

    Dim XPic As IPictureDisp
    Fra(8).Caption = "OB Credit Account"
    Fra(9).Caption = "OB Entitlements Salary"
    Fra(10).Caption = "OB  Allocations "
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Fra(7).Caption = "Accounting data"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Me.framx.Caption = "Salary component"
    XPLbl(39).Caption = "Yearly Value"
    XPLbl(40).Caption = "Monthly Value"
    XPLbl(34).Caption = "Housing"
    XPLbl(35).Caption = "Transport"
    XPLbl(37).Caption = "Food"
    XPLbl(36).Caption = "Mobile"
    XPLbl(45).Caption = "Supervision"
    XPLbl(38).Caption = "Others"
    Command5.Caption = "Hide"

    XPLbl(27).Caption = "Nationality"
    XPLbl(28).Caption = "Religion"
    lbl(12).Caption = "DOB"
    lbl(11).Caption = "Leaving"
    Chk_Stkala.Caption = "Resignation"
    Chk_EndWork.Caption = "Separation"
    lbl(10).Caption = "Date"
    CmdEstkala.Caption = "Reason"
    Fra(3).Caption = "Accommodation"
    XPLbl(16).Caption = "Place"
    XPLbl(15).Caption = "No"
    XPLbl(14).Caption = "Start"
    XPLbl(17).Caption = "End"
    Fra(5).Caption = "Passport"
    XPLbl(23).Caption = "No"
    XPLbl(22).Caption = "Start"
    XPLbl(21).Caption = "End"
    XPLbl(24).Caption = "place"
    Fra(1).Caption = "Work Data"
    SuperLabel1(8).text = "Cost Center"
    XPLbl(7).Caption = "Dep"
    XPLbl(8).Caption = "Job"
    XPLbl(9).Caption = "Spec"
    lbl(9).Caption = "Start Date"
    Fra(0).Caption = "Insurance"
    XPLbl(4).Caption = "status"
    XPLbl(5).Caption = "value"
    XPLbl(6).Caption = "other"
    Fra(2).Caption = "Licence"
    XPLbl(13).Caption = "No"
    XPLbl(12).Caption = "Start"
    XPLbl(11).Caption = "End"
    Fra(4).Caption = "Saudi ID"
    Fra(3).Caption = "IQama No"

    Fra(8).Caption = "Opening Balance "
    OptType(0).Caption = "Depit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "NA"
    lbl(14).Caption = "Balance"
    lbl(13).Caption = "Date"

    Fra(9).Caption = "Opening Balance "
    OptType1(0).Caption = "Depit"
    OptType1(1).Caption = "Credit"
    OptType1(2).Caption = "NA"
    lbl(15).Caption = "Balance"
    lbl(16).Caption = "Date"

    Fra(10).Caption = "Opening Balance "
    OptType2(0).Caption = "Depit"
    OptType2(1).Caption = "Credit"
    OptType2(2).Caption = "NA"
    lbl(18).Caption = "Balance"
    lbl(17).Caption = "Date"

    XPLbl(18).Caption = "No"
    XPLbl(19).Caption = "Start"
    XPLbl(20).Caption = "End"
    CmdExit.Caption = "Exit"
    Fra(6).Caption = "sponsor"
    XPLbl(26).Caption = "NO"
    XPLbl(25).Caption = "Name"
    XPLbl(32).Caption = "Tel"
    XPLbl(33).Caption = "ADD"
    Cmd1.Caption = "Attachments"

    ALLButton2(0).Caption = "Qualifications"
    ALLButton2(1).Caption = "Personnel"
    ALLButton2(2).Caption = "Evaluation"
    ALLButton2(3).Caption = "Health file"

    ALLButton2(4).Caption = "Salary component"

    ALLButton2(6).Caption = "Contract"
    ALLButton2(7).Caption = ""
    ISButton1.Caption = "Insert Imagew"
    ISButton2.Caption = "Insert Signature"
    Frame2.Caption = "Query"
    OptExpirLinc.Caption = "License"
    OptExpirEkama.Caption = "Residence"
    OptExpirPas.Caption = "Passport"
    CommandÛQRY.Caption = "Query"

    Frame3.Caption = "Entry Data"

    XPLbl(29).Caption = "NO"
    XPLbl(30).Caption = "Date"
    XPLbl(31).Caption = "Port"
    Label4.Caption = "Forms"
    Check1.Caption = "Print Image"
    ALLButton1.Caption = "Print"
    Combo1.Clear
    Combo1.AddItem "new Residence"
    Combo1.AddItem "Renew Residence"
    Combo1.AddItem "Residence Replacement"
    Combo1.AddItem "Residence Damaged"
    Combo1.AddItem "Visa"
    Combo1.AddItem "Absence Form"

    Me.Caption = "Drivers Data"
    EleHeader.Caption = Me.Caption
    XPLbl(1).Caption = "Employee Code"
    XPLbl(0).Caption = " Name AR"
    XPLbl(47).Caption = " Name ENG"
    XPLbl(2).Caption = "Employee Salary"
    XPLbl(3).Caption = "Employee Email"
    lbl(3).Caption = "Phone"
    lbl(2).Caption = "Mobile"
    lbl(1).Caption = "Remarks"
    lbl(0).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"
    lbl(5).Caption = "Sales Commission"
    lbl(7).Caption = "Work State"
    lbl(8).Caption = "Commission On Sales Profit"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"

End Sub

