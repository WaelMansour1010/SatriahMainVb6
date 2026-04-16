VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmSearStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13680
   Icon            =   "FrmSearStudent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   13680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame16 
      BackColor       =   &H00E2E9E9&
      Height          =   2235
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   189
      Top             =   4530
      Width           =   13425
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4005
         TabIndex        =   191
         Top             =   1020
         Width           =   1680
      End
      Begin VB.TextBox Text13 
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
         Left            =   4635
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   190
         Top             =   180
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo DcbBranch10 
         Height          =   315
         Left            =   7350
         TabIndex        =   192
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbCustomers 
         Height          =   315
         Left            =   120
         TabIndex        =   193
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   180
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «„— «·«’·«Õ"
         Height          =   300
         Index           =   52
         Left            =   5865
         TabIndex        =   196
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„Ì·"
         Height          =   300
         Index           =   23
         Left            =   5865
         RightToLeft     =   -1  'True
         TabIndex        =   195
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   25
         Left            =   11880
         TabIndex        =   194
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E2E9E9&
      Height          =   2235
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   162
      Top             =   4680
      Width           =   13425
      Begin VB.TextBox TxtStudentEmail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   181
         Top             =   1020
         Width           =   2160
      End
      Begin VB.TextBox TxtUqma9 
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
         Left            =   10215
         TabIndex        =   179
         Top             =   1020
         Width           =   1680
      End
      Begin VB.TextBox TxtPhone9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7350
         TabIndex        =   178
         Top             =   1020
         Width           =   1680
      End
      Begin VB.TextBox TxtMobile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4005
         TabIndex        =   173
         Top             =   1020
         Width           =   1680
      End
      Begin VB.TextBox Text7 
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
         Left            =   11085
         TabIndex        =   170
         Top             =   600
         Width           =   810
      End
      Begin VB.TextBox Text5 
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
         Left            =   4635
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   168
         Top             =   600
         Width           =   1050
      End
      Begin VB.TextBox Text6 
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
         Left            =   4635
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   165
         Top             =   240
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo DcbBranch9 
         Height          =   315
         Left            =   7350
         TabIndex        =   163
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbEmployee 
         Height          =   315
         Left            =   120
         TabIndex        =   166
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbStudent9 
         Height          =   315
         Left            =   7350
         TabIndex        =   171
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCompany9 
         Height          =   315
         Left            =   120
         TabIndex        =   180
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbShiftType 
         Height          =   315
         Left            =   7320
         TabIndex        =   183
         Top             =   1605
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbRestsTypes 
         Height          =   315
         Left            =   3060
         TabIndex        =   184
         Top             =   1590
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·—«Õ…"
         Height          =   285
         Index           =   48
         Left            =   5730
         TabIndex        =   186
         Top             =   1605
         Width           =   1170
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ã„Ê⁄… «·—«Õ« "
         Height          =   240
         Index           =   47
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   185
         Top             =   1665
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÂÊÌ…"
         Height          =   285
         Index           =   18
         Left            =   11880
         TabIndex        =   177
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« ›"
         Height          =   300
         Index           =   46
         Left            =   9075
         TabIndex        =   176
         Top             =   1050
         Width           =   1230
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÃÊ«·"
         Height          =   300
         Index           =   45
         Left            =   5865
         TabIndex        =   175
         Top             =   930
         Width           =   1230
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·»—Ìœ «·«·ﬂ —Ê‰Ì"
         Height          =   300
         Index           =   43
         Left            =   2475
         TabIndex        =   174
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ œ—»"
         Height          =   285
         Index           =   17
         Left            =   11880
         TabIndex        =   172
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„Ì·"
         Height          =   300
         Index           =   16
         Left            =   5865
         RightToLeft     =   -1  'True
         TabIndex        =   169
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÊŸ›"
         Height          =   300
         Index           =   19
         Left            =   5865
         RightToLeft     =   -1  'True
         TabIndex        =   167
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   20
         Left            =   11880
         TabIndex        =   164
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   149
      Top             =   4320
      Width           =   13425
      Begin VB.TextBox Text4 
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
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   153
         Top             =   600
         Width           =   1050
      End
      Begin VB.TextBox Text1 
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
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   152
         Top             =   240
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo DcbBranch8 
         Height          =   315
         Left            =   7350
         TabIndex        =   150
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbInstrucor 
         Height          =   315
         Left            =   120
         TabIndex        =   154
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbGroup 
         Height          =   315
         Left            =   120
         TabIndex        =   155
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCurs 
         Height          =   315
         Left            =   7350
         TabIndex        =   158
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbHall 
         Height          =   315
         Left            =   7350
         TabIndex        =   159
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   960
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„«œ…"
         Height          =   285
         Index           =   15
         Left            =   11880
         TabIndex        =   161
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ«⁄…"
         Height          =   285
         Index           =   14
         Left            =   11880
         TabIndex        =   160
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ—»"
         Height          =   195
         Index           =   13
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   157
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄…"
         Height          =   195
         Index           =   10
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   156
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   11
         Left            =   11880
         TabIndex        =   151
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   133
      Top             =   4320
      Width           =   13425
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   375
         Index           =   0
         Left            =   10560
         TabIndex        =   145
         Top             =   960
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "„‰ ÂÌ…"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox TxtGroupName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6990
         RightToLeft     =   -1  'True
         TabIndex        =   140
         Top             =   600
         Width           =   4545
      End
      Begin VB.TextBox TxtGroupCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   134
         Top             =   240
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DcbBranch7 
         Height          =   315
         Left            =   270
         TabIndex        =   135
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDoplom 
         Height          =   315
         Left            =   270
         TabIndex        =   142
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   375
         Index           =   1
         Left            =   9480
         TabIndex        =   146
         Top             =   960
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "€Ì— „‰ ÂÌ…"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   375
         Index           =   2
         Left            =   8400
         TabIndex        =   147
         Top             =   960
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ«·… «·„Ã„Ê⁄…"
         Height          =   285
         Index           =   42
         Left            =   11880
         TabIndex        =   144
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·œÊ—…"
         Height          =   285
         Index           =   9
         Left            =   5280
         TabIndex        =   143
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   7
         Left            =   5280
         TabIndex        =   141
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„"
         Height          =   285
         Index           =   8
         Left            =   11880
         TabIndex        =   137
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
         Height          =   285
         Index           =   44
         Left            =   11880
         TabIndex        =   136
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   121
      Top             =   4320
      Width           =   13425
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   124
         Top             =   120
         Width           =   4275
         Begin VB.TextBox TxtNoStusCFrom1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox TxtNoStusCTo1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·„ﬁ»Ê·Ì‰ „‰"
            Height          =   315
            Index           =   39
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   315
            Index           =   38
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.TextBox Text2 
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
         Left            =   10440
         MaxLength       =   50
         TabIndex        =   123
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtCandidacyID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   122
         Top             =   240
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DcbCompany6 
         Height          =   315
         Left            =   4560
         TabIndex        =   129
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbBranch6 
         Height          =   315
         Left            =   7350
         TabIndex        =   130
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   4
         Left            =   11880
         TabIndex        =   139
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «· —‘ÌÕ"
         Height          =   285
         Index           =   41
         Left            =   6000
         TabIndex        =   138
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘—ﬂ…"
         Height          =   285
         Index           =   6
         Left            =   11880
         TabIndex        =   131
         Top             =   675
         Width           =   1365
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   107
      Top             =   4320
      Width           =   13425
      Begin VB.TextBox TxtContCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   114
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text3 
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
         Left            =   10440
         MaxLength       =   50
         TabIndex        =   113
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   108
         Top             =   120
         Width           =   4275
         Begin VB.TextBox TxtNoStusCTo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox TxtNoStusCFrom 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   315
            Index           =   35
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·„—‘ÕÌ‰ „‰"
            Height          =   315
            Index           =   32
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   240
            Width           =   1260
         End
      End
      Begin MSDataListLib.DataCombo DcbCompany5 
         Height          =   315
         Left            =   4560
         TabIndex        =   115
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbBranch5 
         Height          =   315
         Left            =   7350
         TabIndex        =   118
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   3
         Left            =   11880
         TabIndex        =   119
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄ﬁœ"
         Height          =   285
         Index           =   36
         Left            =   6120
         TabIndex        =   117
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘—ﬂ…"
         Height          =   285
         Index           =   1
         Left            =   11880
         TabIndex        =   116
         Top             =   675
         Width           =   1365
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   86
      Top             =   4320
      Width           =   13425
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   600
         Width           =   4275
         Begin VB.TextBox FromValue 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox ToValue 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·⁄ﬁœ „‰"
            Height          =   315
            Index           =   34
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   315
            Index           =   33
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.TextBox TxtCode 
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
         TabIndex        =   98
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TxtSudCode 
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
         TabIndex        =   95
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TxtFullcode4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   240
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DcbBranch4 
         Height          =   315
         Left            =   240
         TabIndex        =   88
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton ContType 
         Height          =   255
         Index           =   0
         Left            =   5970
         TabIndex        =   92
         Top             =   240
         Width           =   1050
         _Version        =   786432
         _ExtentX        =   1852
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "⁄ﬁœ „ œ—»"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton ContType 
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   93
         Top             =   240
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   " ⁄ﬁœ ‘—ﬂ« "
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton ContType 
         Height          =   255
         Index           =   2
         Left            =   7320
         TabIndex        =   94
         Top             =   240
         Width           =   1050
         _Version        =   786432
         _ExtentX        =   1852
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbStudent 
         Height          =   315
         Left            =   4680
         TabIndex        =   96
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   1080
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCompany4 
         Height          =   315
         Left            =   4680
         TabIndex        =   99
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   720
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘—ﬂ…"
         Height          =   285
         Index           =   0
         Left            =   12000
         TabIndex        =   100
         Top             =   675
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÿ«·»"
         Height          =   285
         Index           =   12
         Left            =   12000
         TabIndex        =   97
         Top             =   1050
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·⁄ﬁœ"
         Height          =   285
         Index           =   40
         Left            =   12000
         TabIndex        =   91
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   37
         Left            =   3240
         TabIndex        =   90
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·⁄ﬁœ"
         Height          =   285
         Index           =   31
         Left            =   8640
         TabIndex        =   89
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   4320
      Width           =   13425
      Begin VB.TextBox TxtExperience 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   79
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TxtPhone3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   10080
         TabIndex        =   63
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TxtName3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox TxtUQama3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox TxtFullcode3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   240
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   120
         TabIndex        =   69
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbNationality 
         Height          =   315
         Left            =   120
         TabIndex        =   71
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbQuali3 
         Height          =   315
         Left            =   10080
         TabIndex        =   73
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker FrmGradDate 
         Height          =   330
         Left            =   6600
         TabIndex        =   74
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93192195
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker ToGradDate 
         Height          =   330
         Left            =   4560
         TabIndex        =   75
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93192195
         CurrentDate     =   38887
      End
      Begin XtremeSuiteControls.RadioButton TypeTrain 
         Height          =   255
         Index           =   0
         Left            =   9825
         TabIndex        =   81
         Top             =   1320
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "›—œÌ"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton TypeTrain 
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   82
         Top             =   1320
         Width           =   1425
         _Version        =   786432
         _ExtentX        =   2514
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "„‰ ÂÌ »«· ÊŸÌ›"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton TypeTrain 
         Height          =   255
         Index           =   2
         Left            =   11160
         TabIndex        =   83
         Top             =   1320
         Width           =   810
         _Version        =   786432
         _ExtentX        =   1429
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·ÿ·»"
         Height          =   285
         Index           =   30
         Left            =   12000
         TabIndex        =   84
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Œ»—« "
         Height          =   285
         Index           =   29
         Left            =   2880
         TabIndex        =   80
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· Œ—Ã"
         Height          =   195
         Index           =   26
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰ "
         Height          =   315
         Index           =   27
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï "
         Height          =   315
         Index           =   28
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         Height          =   285
         Index           =   25
         Left            =   2880
         TabIndex        =   72
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   24
         Left            =   2880
         TabIndex        =   70
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« ›"
         Height          =   285
         Index           =   23
         Left            =   12000
         TabIndex        =   68
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÂÊÌ…"
         Height          =   285
         Index           =   22
         Left            =   8640
         TabIndex        =   67
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·„ œ—»"
         Height          =   285
         Index           =   21
         Left            =   12000
         TabIndex        =   66
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„"
         Height          =   285
         Index           =   17
         Left            =   8760
         TabIndex        =   65
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ƒÂ· «·œ—«”Ì"
         Height          =   285
         Index           =   16
         Left            =   12000
         TabIndex        =   64
         Top             =   960
         Width           =   1365
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Height          =   1575
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox TxtCode2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtUQama2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox TxtName2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox TxtPhone2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   10080
         TabIndex        =   48
         Top             =   600
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DcbSpecial 
         Height          =   315
         Left            =   4560
         TabIndex        =   56
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· Œ’’"
         Height          =   285
         Index           =   12
         Left            =   8895
         TabIndex        =   57
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„"
         Height          =   285
         Index           =   20
         Left            =   8760
         TabIndex        =   55
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·„œ—»"
         Height          =   285
         Index           =   19
         Left            =   12000
         TabIndex        =   54
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÂÊÌ…"
         Height          =   285
         Index           =   18
         Left            =   2880
         TabIndex        =   53
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« ›"
         Height          =   285
         Index           =   15
         Left            =   12000
         TabIndex        =   52
         Top             =   630
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4440
      Width           =   13425
      Begin VB.ComboBox DcbTypeContract1 
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox DcbStuts1 
         Height          =   315
         Left            =   10080
         TabIndex        =   44
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtSuperVisorName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text15 
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
         Left            =   7815
         MaxLength       =   50
         TabIndex        =   39
         Top             =   960
         Width           =   960
      End
      Begin VB.ComboBox DcbStuts 
         Height          =   315
         Left            =   10080
         TabIndex        =   37
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox TxtStudentPhone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   10080
         TabIndex        =   32
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox DcbTypeContract 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   630
         Width           =   2655
      End
      Begin VB.TextBox TxtName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox TxtUQama 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox TxtFullcode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DcbQuali 
         Height          =   315
         Left            =   4560
         TabIndex        =   33
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   630
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCompany 
         Height          =   315
         Left            =   4560
         TabIndex        =   40
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   960
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„‘—›"
         Height          =   285
         Index           =   11
         Left            =   2880
         TabIndex        =   43
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘—ﬂ…"
         Height          =   285
         Index           =   5
         Left            =   8760
         TabIndex        =   41
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label XPLbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ«·… «·„ œ—»"
         Height          =   285
         Index           =   0
         Left            =   12000
         TabIndex        =   38
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« ›"
         Height          =   285
         Index           =   7
         Left            =   12000
         TabIndex        =   36
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·⁄ﬁœ"
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   35
         Top             =   660
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ƒÂ· «·œ—«”Ì"
         Height          =   285
         Index           =   0
         Left            =   8760
         TabIndex        =   34
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÂÊÌ…"
         Height          =   285
         Index           =   10
         Left            =   2880
         TabIndex        =   26
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·„ œ—»"
         Height          =   285
         Index           =   9
         Left            =   12000
         TabIndex        =   25
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„"
         Height          =   285
         Index           =   8
         Left            =   8760
         TabIndex        =   24
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   0
      Width           =   13665
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ «·„ œ—»Ì‰"
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
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   5400
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   12360
         Picture         =   "FrmSearStudent.frx":6852
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   3015
      Left            =   0
      TabIndex        =   20
      Top             =   720
      Width           =   13665
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2865
         Left            =   0
         TabIndex        =   29
         Top             =   690
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearStudent.frx":15141
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
         Height          =   2865
         Left            =   60
         TabIndex        =   46
         Top             =   450
         Visible         =   0   'False
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         FormatString    =   $"FrmSearStudent.frx":1534C
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
         Height          =   2865
         Left            =   0
         TabIndex        =   58
         Top             =   780
         Visible         =   0   'False
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearStudent.frx":1546D
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
         Height          =   2865
         Left            =   -360
         TabIndex        =   85
         Top             =   570
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearStudent.frx":1565C
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid5 
         Height          =   2865
         Left            =   0
         TabIndex        =   106
         Top             =   630
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearStudent.frx":157B0
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid6 
         Height          =   2865
         Left            =   -30
         TabIndex        =   120
         Top             =   630
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearStudent.frx":158CB
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid7 
         Height          =   2865
         Left            =   0
         TabIndex        =   132
         Top             =   750
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         FormatString    =   $"FrmSearStudent.frx":159EB
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid8 
         Height          =   2865
         Left            =   0
         TabIndex        =   148
         Top             =   600
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         FormatString    =   $"FrmSearStudent.frx":15B25
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid10 
         Height          =   2865
         Left            =   210
         TabIndex        =   182
         Top             =   600
         Visible         =   0   'False
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearStudent.frx":15C51
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid9 
         Height          =   2865
         Left            =   240
         TabIndex        =   187
         Top             =   60
         Visible         =   0   'False
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         FormatString    =   $"FrmSearStudent.frx":15DF3
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
      Begin VSFlex8UCtl.VSFlexGrid GrdCars 
         Height          =   2865
         Left            =   240
         TabIndex        =   188
         Top             =   30
         Width           =   13395
         _cx             =   23627
         _cy             =   5054
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearStudent.frx":15FAD
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   6630
      Width           =   13455
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   10
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   285
         Index           =   2
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   7290
      Width           =   13455
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   13
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
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
         BackStyle       =   0
         ButtonImage     =   "FrmSearStudent.frx":1609F
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         LowerToggledContent=   0   'False
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   14
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "„”Õ"
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
         ButtonImage     =   "FrmSearStudent.frx":1C901
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
      Begin ImpulseButton.ISButton Cmd 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
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
         BackStyle       =   0
         ButtonImage     =   "FrmSearStudent.frx":23163
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
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3720
      Width           =   7035
      Begin VB.TextBox TxtIDTO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ "
         Height          =   195
         Index           =   14
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   6
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   5
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3720
      Width           =   6375
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93192195
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93192195
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ "
         Height          =   195
         Index           =   13
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰ "
         Height          =   315
         Index           =   4
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï "
         Height          =   315
         Index           =   3
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmSearStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
'Dim DCboSearch As FrmGeneralFundReceipt
Public inde As Integer
Sub relod1()
      VSFlexGrid1.Visible = True
      Frame4.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ «·ÿ·«»"
      lbl(13).Caption = " «—ÌŒ «·„Ì·«œ"
   Else
      Label1(2).Caption = "Search Students"
      lbl(13).Caption = "Brith Date"
   End If
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetStudentQualification Me.DcbQuali
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany
   If SystemOptions.UserInterface = ArabicInterface Then
   With DcbStuts
   .Clear
   .AddItem "„” „— ›Ì «·œ—«”…"
   .AddItem "„›’Ê· ⁄‰  «·œ—«”…"
   End With
 With DcbStuts1
   .Clear
   .AddItem "„” „— ›Ì «·œ—«”…"
   .AddItem "„›’Ê· ⁄‰  «·œ—«”…"
   End With
   With DcbTypeContract
    .Clear
   .AddItem "⁄ﬁœ ÿ«·»"
   .AddItem "⁄ﬁœ ‘—ﬂ…"
   End With
      With DcbTypeContract1
    .Clear
   .AddItem "⁄ﬁœ ÿ«·»"
   .AddItem "⁄ﬁœ ‘—ﬂ…"
   End With
   Else
      With DcbStuts
   .Clear
   .AddItem "Continues to Study"
   .AddItem "Terminate"
   End With
   With DcbTypeContract
   .Clear
   .AddItem "Student Contract"
   .AddItem "Company Contract"
   End With
   
    With DcbStuts1
   .Clear
   .AddItem "Continues to Study"
   .AddItem "Terminate"
   End With
   With DcbTypeContract1
   .Clear
   .AddItem "Student Contract"
   .AddItem "Company Contract"
   End With
   End If
End Sub



Private Sub DcbCompany_Change()
DcbCompany_Click (0)
End Sub

Private Sub DcbCompany_Click(Area As Integer)
  If val(DcbCompany.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany.BoundText, EmpCode
    Me.Text15.Text = EmpCode
End Sub

Sub Relod2()
 Dim Dcombos As New ClsDataCombos
  lbreg.Visible = False
      VSFlexGrid2.Visible = True
      Frame5.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ «·„œ—»Ì‰"
      lbl(13).Caption = " «—ÌŒ «·„Ì·«œ"
   Else
      Label1(2).Caption = "Search Instructors"
      lbl(13).Caption = "Brith Date"
   End If
   Dcombos.GetStudentTeachers Me.DcbSpecial
End Sub

Private Sub DcbCompany4_Change()
DcbCompany4_Click (0)
End Sub

Private Sub DcbCompany4_Click(Area As Integer)
  If val(DcbCompany4.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany4.BoundText, EmpCode
    Me.TxtCode.Text = EmpCode
End Sub

Private Sub DcbCompany5_Change()
DcbCompany5_Click (0)
End Sub

Private Sub DcbCompany5_Click(Area As Integer)
  If val(DcbCompany5.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany5.BoundText, EmpCode
    Me.Text3.Text = EmpCode
End Sub

Private Sub DcbCompany6_Change()
DcbCompany6_Click (0)
End Sub

Private Sub DcbCompany6_Click(Area As Integer)
 If val(DcbCompany6.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany6.BoundText, EmpCode
    Me.Text2.Text = EmpCode
End Sub

Private Sub DcbCompany9_Click(Area As Integer)
  If val(DcbCompany9.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany9.BoundText, EmpCode
    Me.Text5.Text = EmpCode
End Sub

Private Sub DcbEmployee_Change()
DcbEmployee_Click (0)
End Sub

Private Sub DcbEmployee_Click(Area As Integer)
If val(Me.DcbEmployee.BoundText) = 0 Then Exit Sub
           Me.Text6.Text = get_EMPLOYEE_Data(val(Me.DcbEmployee.BoundText), "Fullcode")
End Sub

Private Sub DcbGroup_Change()
DcbGroup_Click (0)
End Sub

Private Sub DcbGroup_Click(Area As Integer)
  If val(DcbGroup.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetInstudentGroupCode val(DcbGroup.BoundText), EmpCode, 0
    Me.Text1.Text = EmpCode
End Sub

Private Sub DcbInstrucor_Change()
DcbInstrucor_Click (0)
End Sub

Private Sub DcbInstrucor_Click(Area As Integer)
  If val(DcbInstrucor.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetInstructorCode val(DcbInstrucor.BoundText), EmpCode, 0
    Me.Text4.Text = EmpCode
End Sub

Private Sub DcbStudent_Change()
DcbStudent_Click (0)
End Sub

Private Sub DcbStudent_Click(Area As Integer)
  If val(DcbStudent.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetStudentCode DcbStudent.BoundText, EmpCode
    Me.TxtSudCode.Text = EmpCode
End Sub

Private Sub DcbStudent9_Change()
DcbStudent9_Click (0)
End Sub

Private Sub DcbStudent9_Click(Area As Integer)
  If val(DcbStudent9.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetStudentCode val(DcbStudent9.BoundText), EmpCode
    Me.Text7.Text = EmpCode
End Sub

Private Sub FromValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.FromValue.Text, 0)
End Sub

Private Sub GrdCars_Click()
If inde = 11 Then
    FrmItemsClass.FindRec val(GrdCars.TextMatrix(GrdCars.Row, 1))
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetInstudentGroupCode EmpID, Text1.Text, 1
        DcbGroup.BoundText = EmpID
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID
        DcbCompany.BoundText = EmpID
    End If
End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
   lbl(5).Caption = "From"
   lbl(6).Caption = "To"
   lbl(14).Caption = "No#"
   lbl(13).Caption = "Date"
   lbl(4).Caption = "From"
   lbl(3).Caption = "To"
   lbl(7).Caption = "Telephone"
   lbl(0).Caption = "Qualification"
   lbl(1).Caption = "Type Contract"
   lbl(11).Caption = "Supervisor"
   Label1(5).Caption = "Company"
   XPLbl(0).Caption = "Status"
   lbl(2).Caption = "Total"
   Cmd(0).Caption = "Search"
   Cmd(1).Caption = "Clear"
   Cmd(2).Caption = "Exit"
   lbl(9).Caption = "Code"
   lbl(8).Caption = "Name"
   lbl(10).Caption = "ID"
      With VSFlexGrid1
   .TextMatrix(0, .ColIndex("Serial")) = "Serial"
   .TextMatrix(0, .ColIndex("DateBrithH")) = "Brith Date"
   .TextMatrix(0, .ColIndex("DateBrith")) = "Brith Date"
   .TextMatrix(0, .ColIndex("FullCode")) = "Code"
   .TextMatrix(0, .ColIndex("Name")) = "Name"
   .TextMatrix(0, .ColIndex("UQama")) = "ID"
   .TextMatrix(0, .ColIndex("QuliName")) = "Qualification"
   .TextMatrix(0, .ColIndex("StudentPhone")) = "Student Phone"
   .TextMatrix(0, .ColIndex("StutsID")) = "Status"
   .TextMatrix(0, .ColIndex("TypeContract")) = "Type Contract"
   .TextMatrix(0, .ColIndex("CusName")) = "Company"
   .TextMatrix(0, .ColIndex("SuperVisorName")) = "SuperVisor"
   End With
   '''//////222
   lbl(19).Caption = "Code"
   lbl(18).Caption = "ID"
   lbl(20).Caption = "Name"
   lbl(15).Caption = "Telephone"
   lbl(12).Caption = "Specialization"
   With VSFlexGrid2
   .TextMatrix(0, .ColIndex("Serial")) = "Serial"
   .TextMatrix(0, .ColIndex("FullCode")) = "Code"
   .TextMatrix(0, .ColIndex("Name")) = "Name"
   .TextMatrix(0, .ColIndex("UQama")) = "ID"
   .TextMatrix(0, .ColIndex("SpecName")) = "Specialization"
   .TextMatrix(0, .ColIndex("Phone")) = "Telephone"
   .TextMatrix(0, .ColIndex("Addres")) = "Address"
   End With
 '''////////3
 lbl(21).Caption = "Code"
 lbl(17).Caption = "Name"
 lbl(24).Caption = "Branch"
lbl(25).Caption = "Nationality"
lbl(29).Caption = "Experience"
lbl(23).Caption = "Telephone"
lbl(22).Caption = "ID"
lbl(16).Caption = "Qualification"
lbl(26).Caption = "Graduation Date"
lbl(27).Caption = "From"
lbl(28).Caption = "To"
TypeTrain(2).RightToLeft = False
TypeTrain(1).RightToLeft = False
TypeTrain(0).RightToLeft = False
TypeTrain(0).Caption = "Personal"
TypeTrain(2).Caption = "All"
TypeTrain(1).Caption = "Employment"
lbl(30).Caption = "Type Trining"
      With VSFlexGrid3
   .TextMatrix(0, .ColIndex("Serial")) = "Serial"
   .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
   .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
   .TextMatrix(0, .ColIndex("TypeTrain")) = "Type Trining"
   .TextMatrix(0, .ColIndex("GradDate")) = "Graduation Date"
   .TextMatrix(0, .ColIndex("FullCode")) = "Code"
   .TextMatrix(0, .ColIndex("Name")) = "Name"
   .TextMatrix(0, .ColIndex("UQama")) = "ID"
   .TextMatrix(0, .ColIndex("QuliName")) = "Qualification"
   .TextMatrix(0, .ColIndex("Phone")) = "Telephone"
   .TextMatrix(0, .ColIndex("Experience")) = "Experience"
   .TextMatrix(0, .ColIndex("Nationlname")) = "Nationality"
   End With
  ''////////////444
  lbl(40).Caption = "Code"
  Label1(0).Caption = "Company"
  Label1(12).Caption = "Student"
  lbl(31).Caption = "Type Contract"
  ContType(0).RightToLeft = False
  ContType(1).RightToLeft = False
  ContType(2).RightToLeft = False
  ContType(0).Caption = "Employee"
  ContType(1).Caption = "Company"
  ContType(2).Caption = "All"
  lbl(37).Caption = "Branch"
  lbl(33).Caption = "To"
  lbl(34).Caption = "Value From"
  With VSFlexGrid4
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("FullCode")) = "Code"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("TypeContID")) = "Type"
  .TextMatrix(0, .ColIndex("CusName")) = "Company"
  .TextMatrix(0, .ColIndex("Name")) = "Student"
  .TextMatrix(0, .ColIndex("Price")) = "Value"
  .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
  End With
''///55555
Label1(3).Caption = "Branch"
lbl(36).Caption = "Nomination No."
lbl(39).Caption = "No.Approved"
lbl(38).Caption = "To"
Label1(1).Caption = "Company"
  With VSFlexGrid6
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("ID")) = "No"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("CandidacyID")) = "Nomination No."
  .TextMatrix(0, .ColIndex("CusName")) = "Company"
  .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
  .TextMatrix(0, .ColIndex("NoStudACNow")) = "No.Approved"
  End With
''/////6666
Label1(4).Caption = "Branch"
lbl(41).Caption = "No.Nominees"
lbl(32).Caption = "No. Nominees"
lbl(35).Caption = "To"
Label1(6).Caption = "Company"
  With VSFlexGrid5
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("ID")) = "No"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("NoStusCandid")) = "No. Nominees"
  .TextMatrix(0, .ColIndex("CusName")) = "Company"
  .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
  .TextMatrix(0, .ColIndex("ContCode")) = "No.of Contract"
  End With
''/////777
lbl(44).Caption = "Code"
Label1(8).Caption = "Name"
Label1(7).Caption = "Branch"
Label1(9).Caption = "Diploma"
Rd(0).RightToLeft = False
Rd(1).RightToLeft = False
Rd(2).RightToLeft = False
Rd(0).Caption = "End"
Rd(1).Caption = "Active"
Rd(2).Caption = "All"
lbl(42).Caption = "Group State"
  With VSFlexGrid7
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("ID")) = "No"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("FullCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Group Name"
  .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
  .TextMatrix(0, .ColIndex("DolpName")) = "Diploma"
  .TextMatrix(0, .ColIndex("FlgEnd")) = "End"
  End With
 ''/////////////8888888
 Label1(11).Caption = "Branch"
 Label1(15).Caption = "Subject"
 Label1(14).Caption = "Hall"
 Label1(10).Caption = "Group"
 Label1(13).Caption = "Instructor"
   With VSFlexGrid8
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("ID")) = "No"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
  .TextMatrix(0, .ColIndex("Name")) = "Group Name"
  .TextMatrix(0, .ColIndex("InsName")) = "Instructor "
  .TextMatrix(0, .ColIndex("CurName")) = "Subject"
  .TextMatrix(0, .ColIndex("HallName")) = "Hall"
  End With
''///////99999
Label1(19).Caption = "Caller"
Label1(16).Caption = "Company"
lbl(45).Caption = "Mobile"
lbl(43).Caption = "Email"
lbl(46).Caption = "Telephone"
Label1(20).Caption = "Branch"
Label1(17).Caption = "Student"
Label1(18).Caption = "ID"
   With VSFlexGrid9
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("ID")) = "No"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
  .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
  .TextMatrix(0, .ColIndex("Emp_Name")) = "Caller"
  .TextMatrix(0, .ColIndex("CusName")) = "Company "
  .TextMatrix(0, .ColIndex("Name")) = "Student"
  .TextMatrix(0, .ColIndex("UQama")) = "ID"
  .TextMatrix(0, .ColIndex("Phone")) = "Telephone"
  .TextMatrix(0, .ColIndex("Mobile")) = "Mobile"
  .TextMatrix(0, .ColIndex("Email")) = "Email"
  
  End With
ErrTrap:
End Sub
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic

Frame16.Visible = False
GrdCars.Visible = False
VSFlexGrid9.Visible = False
Frame15.Visible = False
VSFlexGrid8.Visible = False
Frame14.Visible = False
VSFlexGrid7.Visible = False
Frame13.Visible = False
VSFlexGrid6.Visible = False
Frame11.Visible = False
VSFlexGrid5.Visible = False
Frame9.Visible = False
VSFlexGrid1.Visible = False
Frame4.Visible = False
VSFlexGrid2.Visible = False
Frame5.Visible = False
lbreg.Visible = True
Frame6.Visible = False
VSFlexGrid3.Visible = False
VSFlexGrid4.Visible = False
VSFlexGrid10.Visible = False
 Frame7.Visible = False
If inde = 1 Or inde = 101 Or inde = 102 Or inde = 103 Then
 relod1
 ElseIf inde = 2 Or inde = 201 Or inde = 202 Or inde = 203 Then
   Relod2
 ElseIf inde = 3 Then
   Relod3
 ElseIf inde = 4 Or inde = 401 Then
   Relod4
  ElseIf inde = 5 Or inde = 501 Then
   Relod5
  ElseIf inde = 6 Then
   Relod6
ElseIf inde = 7 Or inde = 701 Or inde = 702 Then
   Relod7
ElseIf inde = 8 Then
   Relod8
ElseIf inde = 9 Or inde = 20 Then
   Relod9
ElseIf inde = 10 Then
Relod10
Frame3.Visible = True
GetDataShift

ElseIf inde = 11 Then
Relod11
Frame16.Visible = True
GetDataCars


 '''''''''''''''''''
End If
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
       
      Set GrdBack = New ClsBackGroundPic

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
     SetDtpickerDate Me.DtpDateFrom
     SetDtpickerDate Me.DtpDateTo
      SetDtpickerDate Me.FrmGradDate
     SetDtpickerDate Me.ToGradDate
      
   End Sub
   Private Sub Cmd_Click(Index As Integer)
    Select Case Index
    
        Case 0
        If inde = 1 Or inde = 101 Or inde = 102 Or inde = 103 Then
        GetData
        ElseIf inde = 2 Or inde = 201 Or inde = 202 Or inde = 203 Then
        GetDataInstructor
        ElseIf inde = 3 Then
        GetDataTrining
        ElseIf inde = 4 Or inde = 401 Then
        GetDataContract
        ElseIf inde = 5 Or inde = 501 Then
        GetDataStuCandidacy
         ElseIf inde = 6 Then
        GetDataCandidacyAccept
        ElseIf inde = 7 Or inde = 701 Or inde = 702 Then
        GetDataGroups
        ElseIf inde = 8 Then
        GetDataAttendance
        ElseIf inde = 9 Or inde = 20 Then
        GetDataStudCalling
        ElseIf inde = 10 Then
        GetDataShift
        ElseIf inde = 11 Then
        GetDataCars
        End If
        Case 1
        clear_all Me
        DtpDateFrom.value = ""
        DtpDateTo.value = ""
        FrmGradDate.value = ""
        ToGradDate.value = ""
        TypeTrain(2).value = True
        ContType(2).value = True
        Rd(2).value = True
                If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbll(0).Caption = "Search Results"
            End If
            Case 2
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
 '   PutFormOnTop Me.hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    'Set DCboSearch = Nothing
End Sub
Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblStudent.ID, dbo.TblStudent.Name, dbo.TblStudent.NameE, dbo.TblStudent.FullCode, dbo.TblStudent.UQama, dbo.TblStudent.TypeContract, "
    sql = sql & "                   dbo.TblStudent.SexID, dbo.TblStudent.StudentEmail, dbo.TblStudent.StudentPhone, dbo.TblStudent.DateBrithH, dbo.TblStudent.DateBrith,"
    sql = sql & "                  dbo.TblStudent.StudentAddres, dbo.TblStudent.SuperVisorName, dbo.TblStudent.SuperPhone, dbo.TblStudent.Remarks, dbo.TblStudent.StutsID,"
    sql = sql & "                  dbo.TblStudent.CompID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblStudent.DcbQualiID, dbo.TblStudentQualification.Name AS QuliName,"
    sql = sql & "                  dbo.TblStudentQualification.NameE AS QuliNameE , dbo.TblStudent.BranchID"
    sql = sql & " FROM         dbo.TblStudent LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblStudentQualification ON dbo.TblStudent.DcbQualiID = dbo.TblStudentQualification.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers ON dbo.TblStudent.CompID = dbo.TblCustemers.CusID"
       BolBegine = True
    sql = sql & "  where  (dbo.TblStudent.BranchID=0 or dbo.TblStudent.BranchID is null or         dbo.TblStudent.BranchID in(" & Current_branchSql & "))"
    
    
       StrWhere = ""
  If Me.TxtName.Text <> "" Then
        If BolBegine = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.Name like '%" & Me.TxtName.Text & "%'"
            Else
            StrWhere = StrWhere & "AND  dbo.TblStudent.NameE like '%" & Me.TxtName.Text & "%'"
            End If
        Else
            BolBegine = True
            If SystemOptions.UserInterface = ArabicInterface Then
            StrWhere = " Where dbo.TblStudent.Name like '%" & Me.TxtName.Text & "%'"
            Else
            StrWhere = " Where dbo.TblStudent.NameE like '%" & Me.TxtName.Text & "%'"
            End If
        End If
    End If
      If Me.TxtFullcode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.FullCode like '%" & Me.TxtFullcode.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.FullCode like '%" & Me.TxtFullcode.Text & "%'"
        End If
    End If
      If Me.TxtUQama.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.UQama like '%" & Me.TxtUQama.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.UQama like '%" & Me.TxtUQama.Text & "%'"
        End If
    End If
    If Me.TxtStudentPhone.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.StudentPhone like '%" & Me.TxtStudentPhone.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.StudentPhone like '%" & Me.TxtStudentPhone.Text & "%'"
        End If
    End If
        If Me.DcbQuali.Text <> "" And val(DcbQuali.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.DcbQualiID = " & val(DcbQuali.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.DcbQualiID = " & val(DcbQuali.BoundText) & ""
        End If
    End If
      If Me.DcbTypeContract.Text <> "" And val(DcbTypeContract.ListIndex) <> -1 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.TypeContract = " & val(DcbTypeContract.ListIndex) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.TypeContract = " & val(DcbTypeContract.ListIndex) & ""
        End If
    End If
          If Me.DcbStuts.Text <> "" And val(DcbStuts.ListIndex) <> -1 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.StutsID = " & val(DcbStuts.ListIndex) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.StutsID = " & val(DcbStuts.ListIndex) & ""
        End If
    End If
   ''//////////
           If Me.DcbCompany.Text <> "" And val(DcbCompany.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.CompID = " & val(DcbCompany.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.CompID = " & val(DcbCompany.BoundText) & ""
        End If
    End If
   ''//////
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblStudent.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblStudent.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStudent.DateBrith >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.DateBrith>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStudent.DateBrith <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblStudent.DateBrith<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    If Me.TxtSuperVisorName.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.SuperVisorName like '%" & Me.TxtSuperVisorName.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.SuperVisorName like '%" & Me.TxtSuperVisorName.Text & "%'"
        End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblStudent.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("DateBrith").value)) Then
                .TextMatrix(i, .ColIndex("DateBrith")) = Format(rs("DateBrith").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("DateBrithH")) = IIf(IsNull(rs("DateBrithH").value), "", rs("DateBrithH").value)
                .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
                .TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(rs("UQama").value), "", rs("UQama").value)
                .TextMatrix(i, .ColIndex("StudentPhone")) = IIf(IsNull(rs("StudentPhone").value), "", rs("StudentPhone").value)
                 DcbStuts1.ListIndex = IIf(IsNull(rs("StutsID").value), -1, rs("StutsID").value)
                .TextMatrix(i, .ColIndex("StutsID")) = DcbStuts1.Text
                 DcbTypeContract1.ListIndex = IIf(IsNull(rs("TypeContract").value), -1, rs("TypeContract").value)
                .TextMatrix(i, .ColIndex("TypeContract")) = DcbTypeContract1.Text
                .TextMatrix(i, .ColIndex("SuperVisorName")) = IIf(IsNull(rs("SuperVisorName").value), "", rs("SuperVisorName").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, .ColIndex("QuliName")) = IIf(IsNull(rs("QuliName").value), "", rs("QuliName").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                .TextMatrix(i, .ColIndex("QuliName")) = IIf(IsNull(rs("QuliNameE").value), "", rs("QuliNameE").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataGroups()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblStuGroup.ID, dbo.TblStuGroup.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStuGroup.RecordDateH, "
    sql = sql & "                   dbo.TblStuGroup.RecordDate, dbo.TblStuGroup.Name, dbo.TblStuGroup.NameE, dbo.TblStuGroup.StartDate, dbo.TblStuGroup.FlgEnd, dbo.TblStuGroup.Fullcode,"
    sql = sql & "                   dbo.TblStuGroup.DoplomID, dbo.TblStudentTypeCurs.Name AS DolpName, dbo.TblStudentTypeCurs.NameE AS DolpNameE, dbo.TblStuGroup.Remarks"
    sql = sql & "   FROM         dbo.TblStuGroup LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblStudentTypeCurs ON dbo.TblStuGroup.DoplomID = dbo.TblStudentTypeCurs.ID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblBranchesData ON dbo.TblStuGroup.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = True
    sql = sql & "  where  (dbo.TblStuGroup.BranchID=0 or dbo.TblStuGroup.BranchID is null or         dbo.TblStuGroup.BranchID in(" & Current_branchSql & "))"
    
    
       StrWhere = ""
  If Me.TxtGroupName.Text <> "" Then
        If BolBegine = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrWhere = StrWhere & "AND  dbo.TblStuGroup.Name like '%" & Me.TxtGroupName.Text & "%'"
            Else
            StrWhere = StrWhere & "AND  dbo.TblStuGroup.NameE like '%" & Me.TxtGroupName.Text & "%'"
            End If
        Else
            BolBegine = True
            If SystemOptions.UserInterface = ArabicInterface Then
            StrWhere = " Where dbo.TblStuGroup.Name like '%" & Me.TxtGroupName.Text & "%'"
            Else
            StrWhere = " Where dbo.TblStuGroup.NameE like '%" & Me.TxtGroupName.Text & "%'"
            End If
        End If
    End If
      If Me.TxtGroupCode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuGroup.FullCode like '%" & Me.TxtGroupCode.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuGroup.FullCode like '%" & Me.TxtGroupCode.Text & "%'"
        End If
    End If

        If Me.DcbBranch7.Text <> "" And val(DcbBranch7.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuGroup.BranchID = " & val(DcbBranch7.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuGroup.BranchID = " & val(DcbBranch7.BoundText) & ""
        End If
    End If

           If Me.DcbDoplom.Text <> "" And val(DcbDoplom.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuGroup.DoplomID = " & val(DcbDoplom.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuGroup.DoplomID = " & val(DcbDoplom.BoundText) & ""
        End If
    End If
   ''//////
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuGroup.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuGroup.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblStuGroup.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblStuGroup.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStuGroup.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuGroup.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStuGroup.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblStuGroup.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    If Rd(0).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuGroup.FlgEnd =1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuGroup.FlgEnd =1"
        End If
    End If
        If Rd(1).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND ( dbo.TblStuGroup.FlgEnd =0 or dbo.TblStuGroup.FlgEnd is null) "
        Else
            BolBegine = True
            StrWhere = " Where ( dbo.TblStuGroup.FlgEnd =0 or dbo.TblStuGroup.FlgEnd is null) "
        End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblStuGroup.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid7
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("FlgEnd")) = IIf(IsNull(rs("FlgEnd").value), "", rs("FlgEnd").value)
                .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("DolpName")) = IIf(IsNull(rs("DolpName").value), "", rs("DolpName").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("DolpName")) = IIf(IsNull(rs("DolpNameE").value), "", rs("DolpNameE").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataAttendance()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.TblAttendance.ID, dbo.TblAttendance.RecordDateH, dbo.TblAttendance.RecordDate, dbo.TblAttendance.Remarks, dbo.TblAttendance.BranchID, "
    sql = sql & "                   dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAttendance.GroupID, dbo.TblStuGroup.Name, dbo.TblStuGroup.NameE,"
    sql = sql & "                  dbo.TblStuGroup.Fullcode, dbo.TblAttendance.CursID, dbo.TblStudentCurs.Name AS CurName, dbo.TblStudentCurs.NameE AS CurNameE, dbo.TblAttendance.HallID,"
    sql = sql & "                  dbo.TblStudentClassRooms.Name AS HallName, dbo.TblStudentClassRooms.NameE AS HallNameE, dbo.TblAttendance.InstrcID, dbo.TblInstructors.Name AS InsName,"
    sql = sql & "                   dbo.TblInstructors.NameE AS InsNameE, dbo.TblInstructors.FullCode AS InsFullCode"
    sql = sql & "    FROM         dbo.TblAttendance LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblInstructors ON dbo.TblAttendance.InstrcID = dbo.TblInstructors.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblStudentClassRooms ON dbo.TblAttendance.HallID = dbo.TblStudentClassRooms.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblStudentCurs ON dbo.TblAttendance.CursID = dbo.TblStudentCurs.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblStuGroup ON dbo.TblAttendance.GroupID = dbo.TblStuGroup.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblAttendance.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = True
       StrWhere = ""

  sql = sql & "  where  (dbo.TblAttendance.BranchID=0 or dbo.TblAttendance.BranchID is null or         dbo.TblAttendance.BranchID in(" & Current_branchSql & "))"

    If Me.DcbBranch8.Text <> "" And val(DcbBranch8.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblAttendance.BranchID = " & val(DcbBranch8.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAttendance.BranchID = " & val(DcbBranch8.BoundText) & ""
        End If
    End If

    If Me.DcbGroup.Text <> "" And val(DcbGroup.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblAttendance.GroupID = " & val(DcbGroup.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAttendance.GroupID = " & val(DcbGroup.BoundText) & ""
        End If
    End If
    If Me.DcbInstrucor.Text <> "" And val(DcbInstrucor.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblAttendance.InstrcID = " & val(DcbInstrucor.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAttendance.InstrcID = " & val(DcbInstrucor.BoundText) & ""
        End If
    End If
    
    If Me.DcbCurs.Text <> "" And val(DcbCurs.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblAttendance.CursID = " & val(DcbCurs.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAttendance.CursID = " & val(DcbCurs.BoundText) & ""
        End If
    End If
    
        If Me.DcbHall.Text <> "" And val(DcbHall.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblAttendance.HallID = " & val(DcbHall.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAttendance.HallID = " & val(DcbHall.BoundText) & ""
        End If
    End If
   ''//////
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblAttendance.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAttendance.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblAttendance.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblAttendance.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAttendance.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAttendance.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAttendance.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblAttendance.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblAttendance.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid8
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("InsName")) = IIf(IsNull(rs("InsName").value), "", rs("InsName").value)
                .TextMatrix(i, .ColIndex("HallName")) = IIf(IsNull(rs("HallName").value), "", rs("HallName").value)
                .TextMatrix(i, .ColIndex("CurName")) = IIf(IsNull(rs("CurName").value), "", rs("CurName").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(i, .ColIndex("InsName")) = IIf(IsNull(rs("InsNameE").value), "", rs("InsNameE").value)
                .TextMatrix(i, .ColIndex("HallName")) = IIf(IsNull(rs("HallNameE").value), "", rs("HallNameE").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("CurName")) = IIf(IsNull(rs("CurNameE").value), "", rs("CurNameE").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataStudCalling()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblStudCalling.ID, dbo.TblStudCalling.RecordDateH, dbo.TblStudCalling.RecordDate, dbo.TblStudCalling.Remarks, dbo.TblStudCalling.EnterDateH, "
    sql = sql & "                   dbo.TblStudCalling.EnterDate,TblCustemers.CusID , dbo.TblStudCalling.EnterTime, dbo.TblStudCalling.Mobile, dbo.TblStudCalling.Phone, dbo.TblStudCalling.Email,"
    sql = sql & "                  dbo.TblStudCalling.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblStudCalling.CompID,"
    sql = sql & "                  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblStudCalling.StudID, dbo.TblStudent.Name,"
    sql = sql & "                  dbo.TblStudent.NameE, dbo.TblStudent.FullCode AS StudFullCode, dbo.TblStudCalling.BranchID, dbo.TblBranchesData.branch_name,"
    sql = sql & "                  dbo.TblBranchesData.branch_nameE , dbo.TblStudent.UQama"
    sql = sql & "   FROM         dbo.TblStudCalling LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblStudCalling.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblStudent ON dbo.TblStudCalling.StudID = dbo.TblStudent.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers ON dbo.TblStudCalling.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee ON dbo.TblStudCalling.EmpID = dbo.TblEmployee.Emp_ID"
       BolBegine = True
    sql = sql & "  where  (dbo.TblStudCalling.BranchID=0 or dbo.TblStudCalling.BranchID is null or         dbo.TblStudCalling.BranchID in(" & Current_branchSql & "))"
   
       StrWhere = ""
         If Me.TxtStudentEmail.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudCalling.Email like '%" & Me.TxtStudentEmail.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.Email like '%" & Me.TxtStudentEmail.Text & "%'"
        End If
    End If
    
    If Me.TxtMobile.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudCalling.Mobile like '%" & Me.TxtMobile.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.Mobile like '%" & Me.TxtMobile.Text & "%'"
        End If
    End If
    
      If Me.TxtPhone9.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudCalling.Phone like '%" & Me.TxtPhone9.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.Phone like '%" & Me.TxtPhone9.Text & "%'"
        End If
    End If
      If Me.TxtUqma9.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudent.UQama like '%" & Me.TxtUqma9.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudent.UQama like '%" & Me.TxtUqma9.Text & "%'"
        End If
    End If
    

    If Me.DcbBranch9.Text <> "" And val(DcbBranch9.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudCalling.BranchID = " & val(DcbBranch9.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.BranchID = " & val(DcbBranch9.BoundText) & ""
        End If
    End If

    If Me.DcbStudent9.Text <> "" And val(DcbStudent9.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudCalling.StudID = " & val(DcbStudent9.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.StudID = " & val(DcbStudent9.BoundText) & ""
        End If
    End If
    If Me.DcbEmployee.Text <> "" And val(DcbEmployee.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudCalling.EmpID = " & val(DcbEmployee.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.EmpID = " & val(DcbEmployee.BoundText) & ""
        End If
    End If
    
    If Me.DcbCompany9.Text <> "" And val(DcbCompany9.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudCalling.CompID = " & val(DcbCompany9.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.CompID = " & val(DcbCompany9.BoundText) & ""
        End If
    End If
    
   ''//////
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStudCalling.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblStudCalling.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblStudCalling.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStudCalling.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStudCalling.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStudCalling.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblStudCalling.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblStudCalling.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid9
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
                .TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(rs("Email").value), "", rs("Email").value)
                .TextMatrix(i, .ColIndex("Mobile")) = IIf(IsNull(rs("Mobile").value), "", rs("Mobile").value)
                .TextMatrix(i, .ColIndex("Phone")) = IIf(IsNull(rs("Phone").value), "", rs("Phone").value)
                .TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(rs("UQama").value), "", rs("UQama").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataCars()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
       Dim MySQL As String
       

    MySQL = " SELECT TblRowsEstimated.ID,"
    MySQL = MySQL & "         TblRowsEstimated2.ItemID,"
    MySQL = MySQL & "                 TblRowsEstimated2.UnitID,"
    MySQL = MySQL & "                 TblRowsEstimated2.ShowQty,"
    MySQL = MySQL & "                 TblRowsEstimated2.UnitPrice,"
    MySQL = MySQL & "                 TblRowsEstimated2.Total,"
    MySQL = MySQL & "                 TblRowsEstimated2.Discount,"
    MySQL = MySQL & "                 TblRowsEstimated2.TotalAfterDisc,"
        
    MySQL = MySQL & "                 TblRowsEstimated2.VatValue Vat,"
    MySQL = MySQL & "                 TblRowsEstimated2.VatValue,"
    MySQL = MySQL & "                 TblRowsEstimated2.Net,"
    MySQL = MySQL & "                 TblRowsEstimated2.FullCode,"
    MySQL = MySQL & "                 TblRowsEstimated2.Rem,"
    MySQL = MySQL & "                 TblRowsEstimated2.MasterID,"
    MySQL = MySQL & "                 TblRowsEstimated2.GroupID,"
    MySQL = MySQL & "                 U.UnitName,"
    MySQL = MySQL & "                 G.GroupName,"
    MySQL = MySQL & "                 TblRowsEstimated.PlateNo,"
    MySQL = MySQL & "                 TblRowsEstimated.ClientName,"
    MySQL = MySQL & "                 TblRowsEstimated.Shaseh,"
    MySQL = MySQL & "                 TblCarModels.Model,"
    MySQL = MySQL & "                 TBLCarTypes.Name,"
    MySQL = MySQL & "                 t.ItemCode,"
    MySQL = MySQL & "                 t.ItemName                 AS ItemName,"
    MySQL = MySQL & "                 U.UnitNamee,"
    MySQL = MySQL & "                 TblRowsEstimated.YearFact,"
    MySQL = MySQL & "                 TblRowsEstimated.AuthoOrder,"
    MySQL = MySQL & "                 TblRowsEstimated.CarModelID,"
    MySQL = MySQL & "                 TblRowsEstimated.CarTypeID,"
    MySQL = MySQL & "                 TblRowsEstimated.RecordDate,"
    MySQL = MySQL & "                 TblRowsEstimated.ID        AS Expr2,"
    MySQL = MySQL & "                 TblRowsEstimated.UserID,"
    MySQL = MySQL & "                 TblRowsEstimated.discValue,"
    MySQL = MySQL & "                 TblRowsEstimated.DiscPercent,"
    MySQL = MySQL & "                 TblRowsEstimated.TotalAfterDiscount,"
    MySQL = MySQL & "                 TblRowsEstimated.Vatyo,"
    MySQL = MySQL & "                 TblRowsEstimated.Vat2"
    
    MySQL = MySQL & "                 From TblCarModels"
    MySQL = MySQL & "                        RIGHT OUTER JOIN TblCarsData"
    MySQL = MySQL & "                             ON  TblCarModels.CarID = TblCarsData.id"
    MySQL = MySQL & "                        RIGHT OUTER JOIN TblRowsEstimated"
    MySQL = MySQL & "                        LEFT OUTER JOIN TBLCarTypes"
    MySQL = MySQL & "                             ON  TblRowsEstimated.CarTypeID = TBLCarTypes.id"
    MySQL = MySQL & "                             ON  TblCarsData.id = TblRowsEstimated.CarTypeID"
    MySQL = MySQL & "                             AND TblCarModels.Id = TblRowsEstimated.CarModelID"
    MySQL = MySQL & "                        LEFT OUTER JOIN TblRowsEstimated2"
    MySQL = MySQL & "                        LEFT OUTER JOIN TblUnites   AS U"
    MySQL = MySQL & "                             ON  U.UnitID = TblRowsEstimated2.UnitID"
    MySQL = MySQL & "                        LEFT OUTER JOIN TblItems    AS t"
    MySQL = MySQL & "                             ON  t.ItemID = TblRowsEstimated2.ItemID"
    MySQL = MySQL & "                        LEFT OUTER JOIN Groups      AS G"
    MySQL = MySQL & "                             ON  G.GroupID = TblRowsEstimated2.GroupID"
    MySQL = MySQL & "                             ON  TblRowsEstimated.ID = TblRowsEstimated2.MasterID"
    
    MySQL = MySQL & "          Where 1 = 1 "
   
       
       
       BolBegine = True
   
   
       
       
    
    If Me.cmbCustomers.Text <> "" And val(cmbCustomers.BoundText) <> 0 Then
            StrWhere = StrWhere & "AND  dbo.TblRowsEstimated.CusID = " & val(cmbCustomers.BoundText) & ""
       
    End If
    
    If Me.DcbBranch10.Text <> "" And val(DcbBranch10.BoundText) <> 0 Then
            StrWhere = StrWhere & "AND  dbo.TblRowsEstimated.BranchID = " & val(DcbBranch10.BoundText) & ""
       
    End If
     
    If val(Me.Text11.Text) <> 0 Then
       
            StrWhere = StrWhere & "AND  dbo.TblRowsEstimated.AuthoOrder =" & val(Me.Text11.Text) & ""
      
     
    End If
    
   ''//////
    If val(Me.TxtIDFrom.Text) <> 0 Then
       
            StrWhere = StrWhere & "AND  dbo.TblRowsEstimated.ID >=" & val(Me.TxtIDFrom.Text) & ""
      
     
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
    
          StrWhere = StrWhere & " AND dbo.TblRowsEstimated.ID <=" & val(Me.TxtIDTO.Text) & ""
   
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        
            StrWhere = StrWhere & " AND dbo.TblRowsEstimated.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
       
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
      
            StrWhere = StrWhere & " AND dbo.TblRowsEstimated.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If
      VSFlexGrid9.Visible = False
    VSFlexGrid8.Visible = False
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MySQL = MySQL & StrWhere
    'sql = sql & " Order By dbo.tblRestsSiftTrans.ID "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
     '   MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
     '   MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
  
    GrdCars.Visible = True
        With Me.GrdCars
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
                .TextMatrix(i, .ColIndex("AuthoOrder")) = IIf(IsNull(rs("AuthoOrder").value), "", rs("AuthoOrder").value)
           
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If

End Sub
Public Sub GetDataShift()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
       Dim s As String
    s = " SELECT tblRestsSiftTrans.ID, TbLSheft.SheftName          AS ShiftTypeName,tblRestsSiftTrans.RecordDate,"
    s = s & "        tblRestsTypes.Name RestsTypesName,"
    s = s & "        TblEmployee.Emp_Name     EmpName,DayName = DATENAME(  week ,tblRestsSiftTrans2.fromdate) +2  "
    's = s & "        tblRestsSiftTrans2.*"
    s = s & " From tblRestsSiftTrans2"
    s = s & "        Left Outer JOIN TbLSheft"
    
    s = s & "        "
    s = s & "             ON  TbLSheft.SeftCode = tblRestsSiftTrans2.ShiftTypeID"
    s = s & "        Left Outer JOIN tblRestsSiftTrans"
    s = s & "        On tblRestsSiftTrans2.MasterId = tblRestsSiftTrans.ID "
    s = s & "        Left Outer join TblEmployee"
    s = s & "             ON  TblEmployee.Emp_Id = tblRestsSiftTrans2.EmpId"
    s = s & "        Left Outer join tblRestsTypes"
    s = s & "             ON  tblRestsTypes.Id = tblRestsSiftTrans2.RestsTypesID"
    s = s & " Where 1 = 1 "
       
       
       BolBegine = True
   
   
       
       
    
    If Me.DcbEmployee.Text <> "" And val(DcbEmployee.BoundText) <> 0 Then
            StrWhere = StrWhere & "AND  dbo.tblRestsSiftTrans2.EmpID = " & val(DcbEmployee.BoundText) & ""
       
    End If
    
    If Me.cmbShiftType.Text <> "" And val(cmbShiftType.BoundText) <> 0 Then
  
        StrWhere = StrWhere & " AND  dbo.TbLSheft.SeftCode = " & val(cmbShiftType.BoundText) & ""
   
    End If
    
   ''//////
    If val(Me.TxtIDFrom.Text) <> 0 Then
       
            StrWhere = StrWhere & "AND  dbo.tblRestsSiftTrans.ID >=" & val(Me.TxtIDFrom.Text) & ""
      
     
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
    
          StrWhere = StrWhere & " AND dbo.tblRestsSiftTrans.ID <=" & val(Me.TxtIDTO.Text) & ""
   
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        
            StrWhere = StrWhere & " AND dbo.tblRestsSiftTrans.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
       
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
      
            StrWhere = StrWhere & " AND dbo.tblRestsSiftTrans.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If
      VSFlexGrid9.Visible = False
    VSFlexGrid8.Visible = False
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    s = s & StrWhere
    'sql = sql & " Order By dbo.tblRestsSiftTrans.ID "
    Set rs = New ADODB.Recordset
    rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
     '   MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
     '   MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
  
    
        With Me.VSFlexGrid10
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("SheftName")) = IIf(IsNull(rs("ShiftTypeName").value), "", rs("ShiftTypeName").value)
                .TextMatrix(i, .ColIndex("RestsTypesName")) = IIf(IsNull(rs("RestsTypesName").value), "", rs("RestsTypesName").value)
                '.TextMatrix(i, .ColIndex("Mobile")) = IIf(IsNull(rs("Mobile").value), "", rs("Mobile").value)
                '.TextMatrix(i, .ColIndex("Phone")) = IIf(IsNull(rs("Phone").value), "", rs("Phone").value)
                '.TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(rs("UQama").value), "", rs("UQama").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("EmpName").value), "", rs("EmpName").value)
                '.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                '.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                '.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                '.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
               ' .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
               ' .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               ' .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub

Public Sub GetDataContract()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblContrStudent.ID, dbo.TblContrStudent.RecordDateH, dbo.TblContrStudent.RecordDate, dbo.TblContrStudent.CompID, dbo.TblCustemers.CusName, "
    sql = sql & "                   dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblContrStudent.StudeID, dbo.TblStudent.Name, dbo.TblStudent.NameE,"
    sql = sql & "                  dbo.TblStudent.FullCode AS StudFullCode, dbo.TblContrStudent.Price, dbo.TblContrStudent.TypeContID, dbo.TblContrStudent.BranchID,"
    sql = sql & "                  dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE, dbo.TblContrStudent.fullcode , dbo.TblContrStudent.ContType"
    sql = sql & "  FROM         dbo.TblContrStudent LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblContrStudent.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblStudent ON dbo.TblContrStudent.StudeID = dbo.TblStudent.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers ON dbo.TblContrStudent.CompID = dbo.TblCustemers.CusID"
       BolBegine = True
       StrWhere = ""
    sql = sql & "  where  (dbo.TblContrStudent.BranchID=0 or dbo.TblContrStudent.BranchID is null or         dbo.TblContrStudent.BranchID in(" & Current_branchSql & "))"
      If Me.TxtFullcode4.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblContrStudent.FullCode like '%" & Me.TxtFullcode4.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblContrStudent.FullCode like '%" & Me.TxtFullcode4.Text & "%'"
        End If
    End If
        If Me.DcbBranch4.Text <> "" And val(DcbBranch4.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblContrStudent.BranchID = " & val(DcbBranch4.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblContrStudent.BranchID = " & val(DcbBranch4.BoundText) & ""
        End If
    End If
    
      If Me.DcbCompany4.Text <> "" And val(DcbCompany4.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblContrStudent.CompID = " & val(DcbCompany4.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblContrStudent.CompID = " & val(DcbCompany4.BoundText) & ""
        End If
    End If
          If Me.DcbStudent.Text <> "" And val(DcbStudent.BoundText) <> -0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblContrStudent.StudeID = " & val(DcbStudent.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblContrStudent.StudeID = " & val(DcbStudent.BoundText) & ""
        End If
    End If

   ''//////
    If val(Me.FromValue.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblContrStudent.Price >=" & val(Me.FromValue.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblContrStudent.Price >=" & val(Me.FromValue.Text) & ""
        End If
    End If
    If val(Me.ToValue.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblContrStudent.Price <=" & val(Me.ToValue.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblContrStudent.Price <=" & val(Me.ToValue.Text) & ""
       End If
    End If
    ''/////////////
      If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblContrStudent.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblContrStudent.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblContrStudent.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblContrStudent.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
    
       If ContType(0).value = True Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblContrStudent.ContType =0"
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblContrStudent.ContType =0"
       End If
    End If
       If ContType(1).value = True Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblContrStudent.ContType =1"
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblContrStudent.ContType =1"
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblContrStudent.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblContrStudent.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblContrStudent.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblContrStudent.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
 If inde <> 4 Then
      If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblContrStudent.ContType =1"
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblContrStudent.ContType =1"
       End If
 End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblContrStudent.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid4
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                 .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
                 If IsNull(rs("ContType").value) Or rs("ContType").value = 0 Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("TypeContID")) = "⁄ﬁœ ÿ«·»"
                 Else
                 .TextMatrix(i, .ColIndex("TypeContID")) = "Contract Student"
                 End If
                 ElseIf rs("ContType").value = 1 Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("TypeContID")) = "⁄ﬁœ ‘—ﬂ…"
                 Else
                 .TextMatrix(i, .ColIndex("TypeContID")) = "Contract Company"
                 End If
                 End If
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs("Price").value), 0, rs("Price").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataTrining()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblTrainingRequest.ID, dbo.TblTrainingRequest.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
    sql = sql & "                   dbo.TblTrainingRequest.RecordDateH, dbo.TblTrainingRequest.RecordDate, dbo.TblTrainingRequest.TypeTrain, dbo.TblTrainingRequest.FullCode,"
    sql = sql & "                   dbo.TblTrainingRequest.Name, dbo.TblTrainingRequest.NameE, dbo.TblTrainingRequest.UQama, dbo.TblTrainingRequest.NationalID,"
    sql = sql & "                   dbo.Nationality.name AS Nationlname, dbo.Nationality.namee AS NationlnameE, dbo.TblTrainingRequest.QualiID, dbo.TblStudentQualification.Name AS QuliName,"
    sql = sql & "                   dbo.TblStudentQualification.NameE AS QuliNameE, dbo.TblTrainingRequest.Jeha, dbo.TblTrainingRequest.GradDateH, dbo.TblTrainingRequest.GradDate,"
    sql = sql & "                   dbo.TblTrainingRequest.SexID, dbo.TblTrainingRequest.Experience, dbo.TblTrainingRequest.Remarks, dbo.TblTrainingRequest.JehaAccept,"
    sql = sql & "                   dbo.TblTrainingRequest.Phone, dbo.TblTrainingRequest.Email, dbo.TblTrainingRequest.DateBrithH, dbo.TblTrainingRequest.DateBrith,"
    sql = sql & "                   dbo.TblTrainingRequest.Aprove, dbo.TblTrainingRequest.AproveDateH, dbo.TblTrainingRequest.AproveDate, dbo.TblTrainingRequest.RemarkAprove,"
    sql = sql & "                   dbo.TblTrainingRequest.Address"
    sql = sql & "       FROM         dbo.TblTrainingRequest LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblStudentQualification ON dbo.TblTrainingRequest.QualiID = dbo.TblStudentQualification.ID LEFT OUTER JOIN"
    sql = sql & "                   dbo.Nationality ON dbo.TblTrainingRequest.NationalID = dbo.Nationality.id LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblBranchesData ON dbo.TblTrainingRequest.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = True
    sql = sql & "  where  (dbo.TblTrainingRequest.BranchID =0 or dbo.TblTrainingRequest.BranchID  is null or         dbo.TblTrainingRequest.BranchID  in(" & Current_branchSql & "))"
    
       StrWhere = ""
  If Me.TxtName3.Text <> "" Then
        If BolBegine = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.Name like '%" & Me.TxtName3.Text & "%'"
            Else
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.NameE like '%" & Me.TxtName3.Text & "%'"
            End If
        Else
            BolBegine = True
            If SystemOptions.UserInterface = ArabicInterface Then
            StrWhere = " Where dbo.TblTrainingRequest.Name like '%" & Me.TxtName3.Text & "%'"
            Else
            StrWhere = " Where dbo.TblTrainingRequest.NameE like '%" & Me.TxtName3.Text & "%'"
            End If
        End If
    End If
      If Me.TxtFullcode3.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.FullCode like '%" & Me.TxtFullcode3.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.FullCode like '%" & Me.TxtFullcode3.Text & "%'"
        End If
    End If
      If Me.TxtUQama3.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.UQama like '%" & Me.TxtUQama3.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.UQama like '%" & Me.TxtUQama3.Text & "%'"
        End If
    End If
    If Me.TxtPhone3.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.Phone like '%" & Me.TxtPhone3.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.Phone like '%" & Me.TxtPhone3.Text & "%'"
        End If
    End If
        If Me.DcbNationality.Text <> "" And val(DcbNationality.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.NationalID = " & val(DcbNationality.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.NationalID = " & val(DcbNationality.BoundText) & ""
        End If
    End If
      If Me.DcbQuali3.Text <> "" And val(DcbQuali3.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.QualiID = " & val(DcbQuali3.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.QualiID = " & val(DcbQuali3.BoundText) & ""
        End If
    End If
    If TypeTrain(0).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.TypeTrain = 0"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.TypeTrain =0"
    End If
    End If
       If TypeTrain(1).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.TypeTrain = 1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.TypeTrain =1"
    End If
  End If
   ''//////
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblTrainingRequest.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblTrainingRequest.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblTrainingRequest.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblTrainingRequest.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblTrainingRequest.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    ''''''''''
       If Not IsNull(Me.FrmGradDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblTrainingRequest.GradDate >=" & SQLDate(Me.FrmGradDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.GradDate>=" & SQLDate(Me.FrmGradDate.value, True) & ""
        End If
    End If
    If Not IsNull(Me.ToGradDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblTrainingRequest.GradDate <=" & SQLDate(Me.ToGradDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblTrainingRequest.GradDate<=" & SQLDate(Me.ToGradDate.value, True) & ""
        End If
    End If
    ''''
    If Me.TxtExperience.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTrainingRequest.Experience like '%" & Me.TxtExperience.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTrainingRequest.Experience like '%" & Me.TxtExperience.Text & "%'"
        End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblTrainingRequest.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid3
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                 If Not (IsNull(rs("GradDate").value)) Then
                .TextMatrix(i, .ColIndex("GradDate")) = Format(rs("GradDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
                .TextMatrix(i, .ColIndex("Phone")) = IIf(IsNull(rs("Phone").value), "", rs("Phone").value)
                .TextMatrix(i, .ColIndex("Experience")) = IIf(IsNull(rs("Experience").value), "", rs("Experience").value)
                .TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(rs("UQama").value), "", rs("UQama").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("QuliName")) = IIf(IsNull(rs("QuliName").value), "", rs("QuliName").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("Nationlname")) = IIf(IsNull(rs("Nationlname").value), "", rs("Nationlname").value)
                Else
                .TextMatrix(i, .ColIndex("Nationlname")) = IIf(IsNull(rs("NationlnameE").value), "", rs("NationlnameE").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("QuliName")) = IIf(IsNull(rs("QuliNameE").value), "", rs("QuliNameE").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
              If rs("TypeTrain").value = 1 Then
                  If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(i, .ColIndex("TypeTrain")) = "„‰ ÂÌ »«· ÊŸÌ›"
               Else
               .TextMatrix(i, .ColIndex("TypeTrain")) = "Employment"
               End If
               Else
               If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(i, .ColIndex("TypeTrain")) = "›—œÌ"
               Else
               .TextMatrix(i, .ColIndex("TypeTrain")) = "Personal"
               End If
               
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataStuCandidacy()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblStuCandidacy.ID, dbo.TblStuCandidacy.RecordDateH, dbo.TblStuCandidacy.RecordDate, dbo.TblStuCandidacy.Remarks, dbo.TblStuCandidacy.BranchID, "
    sql = sql & "                   dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStuCandidacy.CompID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    sql = sql & "                  dbo.TblCustemers.Fullcode, dbo.TblStuCandidacy.ContNoID, dbo.TblStuCandidacy.NoStudCon, dbo.TblStuCandidacy.NoStudAccept,"
    sql = sql & "                  dbo.TblStuCandidacy.NoStudRemain, dbo.TblStuCandidacy.NoStusCandid, dbo.TblStuCandidacy.DataFom, dbo.TblStuCandidacy.DataFomH,"
    sql = sql & "                  dbo.TblStuCandidacy.DateTo , dbo.TblStuCandidacy.DateToH, dbo.TblStuCandidacy.ContCode"
    sql = sql & "   FROM         dbo.TblStuCandidacy LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers ON dbo.TblStuCandidacy.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblStuCandidacy.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = True
       StrWhere = ""
    sql = sql & "  where  (dbo.TblStuCandidacy.BranchID=0 or dbo.TblStuCandidacy.BranchID is null or         dbo.TblStuCandidacy.BranchID in(" & Current_branchSql & "))"
    
    
      If Me.TxtContCode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacy.ContCode ='" & Me.TxtContCode.Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacy.ContCode  ='" & Me.TxtContCode.Text & "'"
        End If
    End If

        If Me.DcbBranch5.Text <> "" And val(DcbBranch5.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacy.BranchID = " & val(DcbBranch5.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacy.BranchID = " & val(DcbBranch5.BoundText) & ""
        End If
    End If
      If Me.DcbCompany5.Text <> "" And val(DcbCompany5.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacy.CompID = " & val(DcbCompany5.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacy.CompID = " & val(DcbCompany5.BoundText) & ""
        End If
    End If
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacy.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacy.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblStuCandidacy.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblStuCandidacy.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
    
   ''//////
    If val(Me.TxtNoStusCFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacy.NoStusCandid >=" & val(Me.TxtNoStusCFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacy.NoStusCandid >=" & val(Me.TxtNoStusCFrom.Text) & ""
        End If
    End If
    If val(Me.TxtNoStusCTo.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblStuCandidacy.NoStusCandid <=" & val(Me.TxtNoStusCTo.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblStuCandidacy.NoStusCandid <=" & val(Me.TxtNoStusCTo.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStuCandidacy.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacy.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStuCandidacy.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblStuCandidacy.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If


    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblStuCandidacy.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid5
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("NoStusCandid")) = IIf(IsNull(rs("NoStusCandid").value), "", rs("NoStusCandid").value)
                .TextMatrix(i, .ColIndex("ContCode")) = IIf(IsNull(rs("ContCode").value), "", rs("ContCode").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataCandidacyAccept()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblStuCandidacyAccept.ID, dbo.TblStuCandidacyAccept.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
    sql = sql & "                   dbo.TblStuCandidacyAccept.RecordDateH, dbo.TblStuCandidacyAccept.RecordDate, dbo.TblStuCandidacyAccept.CompID, dbo.TblCustemers.CusName,"
    sql = sql & "                  dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblStuCandidacyAccept.Remarks, dbo.TblStuCandidacyAccept.CandidacyID,"
    sql = sql & "                  dbo.TblStuCandidacyAccept.NoStudCon, dbo.TblStuCandidacyAccept.NoStudAccept, dbo.TblStuCandidacyAccept.NoStudRemain,"
    sql = sql & "                  dbo.TblStuCandidacyAccept.NoStusCandid, dbo.TblStuCandidacyAccept.AcceptDateH, dbo.TblStuCandidacyAccept.AcceptDate,"
    sql = sql & "                  dbo.TblStuCandidacyAccept.NoStudACNow , dbo.TblStuCandidacyAccept.ContNoID"
    sql = sql & "   FROM         dbo.TblStuCandidacyAccept LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers ON dbo.TblStuCandidacyAccept.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblStuCandidacyAccept.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = True
    sql = sql & "  where  (dbo.TblStuCandidacyAccept.BranchID=0 or dbo.TblStuCandidacyAccept.BranchID is null or         dbo.TblStuCandidacyAccept.BranchID in(" & Current_branchSql & "))"
       StrWhere = ""

      If val(Me.TxtCandidacyID.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacyAccept.CandidacyID =" & Me.TxtCandidacyID.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacyAccept.CandidacyID  =" & Me.TxtCandidacyID.Text & ""
        End If
    End If

        If Me.DcbBranch6.Text <> "" And val(DcbBranch6.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacyAccept.BranchID = " & val(DcbBranch6.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacyAccept.BranchID = " & val(DcbBranch6.BoundText) & ""
        End If
    End If
      If Me.DcbCompany6.Text <> "" And val(DcbCompany6.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacyAccept.CompID = " & val(DcbCompany6.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacyAccept.CompID = " & val(DcbCompany6.BoundText) & ""
        End If
    End If
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacyAccept.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacyAccept.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblStuCandidacyAccept.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblStuCandidacyAccept.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
    
   ''//////
    If val(Me.TxtNoStusCFrom1.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblStuCandidacyAccept.NoStudACNow >=" & val(Me.TxtNoStusCFrom1.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacyAccept.NoStudACNow >=" & val(Me.TxtNoStusCFrom1.Text) & ""
        End If
    End If
    If val(Me.TxtNoStusCTo1.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblStuCandidacyAccept.NoStudACNow <=" & val(Me.TxtNoStusCTo1.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblStuCandidacyAccept.NoStudACNow <=" & val(Me.TxtNoStusCTo1.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStuCandidacyAccept.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStuCandidacyAccept.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStuCandidacyAccept.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblStuCandidacyAccept.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If


    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblStuCandidacyAccept.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid6
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("NoStudACNow")) = IIf(IsNull(rs("NoStudACNow").value), "", rs("NoStudACNow").value)
                .TextMatrix(i, .ColIndex("CandidacyID")) = IIf(IsNull(rs("CandidacyID").value), "", rs("CandidacyID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataInstructor()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblInstructors.ID, dbo.TblInstructors.Name, dbo.TblInstructors.NameE, dbo.TblInstructors.FullCode, dbo.TblInstructors.UQama, dbo.TblInstructors.Addres, "
    sql = sql & "                   dbo.TblInstructors.Remarks, dbo.TblInstructors.SexID, dbo.TblInstructors.Email, dbo.TblInstructors.Phone, dbo.TblInstructors.SpecialID,"
    sql = sql & "                  dbo.TblStudentTeachers.Name AS SpecName, dbo.TblStudentTeachers.NameE AS SpecNameE"
    sql = sql & "    FROM         dbo.TblInstructors LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblStudentTeachers ON dbo.TblInstructors.SpecialID = dbo.TblStudentTeachers.ID"
       BolBegine = True
     sql = sql & "  where  (dbo.TblInstructors.BranchID=0 or dbo.TblInstructors.BranchID is null or         dbo.TblInstructors.BranchID in(" & Current_branchSql & "))"
    
    
       StrWhere = ""
  If Me.TxtName2.Text <> "" Then
        If BolBegine = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrWhere = StrWhere & "AND  dbo.TblInstructors.Name like '%" & Me.TxtName2.Text & "%'"
            Else
            StrWhere = StrWhere & "AND  dbo.TblInstructors.NameE like '%" & Me.TxtName2.Text & "%'"
            End If
        Else
            BolBegine = True
            If SystemOptions.UserInterface = ArabicInterface Then
            StrWhere = " Where dbo.TblInstructors.Name like '%" & Me.TxtName2.Text & "%'"
            Else
            StrWhere = " Where dbo.TblInstructors.NameE like '%" & Me.TxtName2.Text & "%'"
            End If
        End If
    End If
      If Me.TxtCode2.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblInstructors.FullCode like '%" & Me.TxtCode2.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstructors.FullCode like '%" & Me.TxtCode2.Text & "%'"
        End If
    End If
      If Me.TxtUQama2.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblInstructors.UQama like '%" & Me.TxtUQama2.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstructors.UQama like '%" & Me.TxtUQama2.Text & "%'"
        End If
    End If
    If Me.TxtPhone2.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblInstructors.Phone like '%" & Me.TxtPhone2.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstructors.Phone like '%" & Me.TxtPhone2.Text & "%'"
        End If
    End If
        If Me.DcbSpecial.Text <> "" And val(DcbSpecial.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblInstructors.SpecialID = " & val(DcbSpecial.BoundText) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstructors.SpecialID = " & val(DcbSpecial.BoundText) & ""
        End If
    End If
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblInstructors.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstructors.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblInstructors.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblInstructors.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If

    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblInstructors.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Sorry...There are no matching condition data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid2
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
                .TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(rs("UQama").value), "", rs("UQama").value)
                .TextMatrix(i, .ColIndex("Phone")) = IIf(IsNull(rs("Phone").value), "", rs("Phone").value)

                .TextMatrix(i, .ColIndex("Addres")) = IIf(IsNull(rs("Addres").value), "", rs("Addres").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("SpecName")) = IIf(IsNull(rs("SpecName").value), "", rs("SpecName").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                .TextMatrix(i, .ColIndex("SpecName")) = IIf(IsNull(rs("SpecNameE").value), "", rs("SpecNameE").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Sub Relod5()
      VSFlexGrid5.Visible = True
      Frame9.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ  —‘ÌÕ «·„ œ—»Ì‰ "
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Students Nomination To Companies"
      lbl(13).Caption = " Date"
   End If
   TypeTrain(2).value = True
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch5
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany5
End Sub
Sub Relod7()
Rd(2).value = True
      VSFlexGrid7.Visible = True
      Frame13.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ «·„Ã„Ê⁄«  "
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Groups"
      lbl(13).Caption = " Date"
   End If
   TypeTrain(2).value = True
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch7
  Dcombos.GetStudentDeploma Me.DcbDoplom
End Sub
Sub Relod9()
      VSFlexGrid9.Visible = True
      Frame15.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ »Ì«‰«  «·« ’«·/ ÕÃ“ «·„Ê«⁄Ìœ "
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Record Attendance"
      lbl(13).Caption = " Date"
   End If
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch9
   Dcombos.GetStudent Me.DcbStudent9, 1
   Dcombos.GetEmployees Me.DcbEmployee
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany9
   Text7.Visible = False
   DcbStudent9.Visible = False
   DcbBranch9.Visible = False
   Label1(20).Visible = False
   Label1(17).Visible = False
   Label1(18).Visible = False
   TxtUqma9.Visible = False
   lbl(43).Visible = False
   TxtStudentEmail.Visible = False
End Sub

Sub Relod10()
VSFlexGrid9.Visible = True
      VSFlexGrid10.Visible = True
      Frame15.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ »Ì«‰«  Œÿ… «·—«Õ« "
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Record Attendance"
      lbl(13).Caption = " Date"
   End If
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch9
   Dcombos.GetStudent Me.DcbStudent9, 1
   Dcombos.GetEmployees Me.DcbEmployee
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany9
   Text7.Visible = False
   DcbStudent9.Visible = False
   DcbBranch9.Visible = False
   Label1(20).Visible = False
   Label1(17).Visible = False
   Label1(18).Visible = False
   TxtUqma9.Visible = False
   lbl(43).Visible = False
    lbl(45).Visible = False
   Label1(16).Visible = False
   Text5.Visible = False
   DcbCompany9.Visible = False
   TxtMobile.Visible = False
   TxtStudentEmail.Visible = False
   TxtStudentEmail.Visible = False
   
   lbl(47).Visible = True
    lbl(48).Visible = True
   cmbRestsTypes.Visible = True
   cmbShiftType.Visible = True
   
   
   Dim sql As String
    sql = "SELECT id ,Name "
        sql = sql & " From dbo.tblRestsTypes"
        
        fill_combo cmbRestsTypes, sql
 
        sql = "SELECT SeftCode , SheftName "
        sql = sql & " From dbo.TbLSheft"
        
        fill_combo cmbShiftType, sql
End Sub




Sub Relod11()
GrdCars.Visible = True
      GrdCars.Visible = True
      Frame16.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ »Ì«‰«  «·ﬁÿ⁄ «·„ﬁœ—…"
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Record Attendance"
      lbl(13).Caption = " Date"
   End If
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch10
   
  
   Dcombos.GetCustomersSuppliers 55, Me.cmbCustomers
   Text7.Visible = False
   DcbStudent9.Visible = False
   DcbBranch9.Visible = False
   Label1(20).Visible = False
   Label1(17).Visible = False
   Label1(18).Visible = False
   TxtUqma9.Visible = False
   lbl(43).Visible = False
    lbl(45).Visible = False
   Label1(16).Visible = False
   Text5.Visible = False
   DcbCompany9.Visible = False
   TxtMobile.Visible = False
   TxtStudentEmail.Visible = False
   TxtStudentEmail.Visible = False
   
  
   
   
 
End Sub




Sub Relod8()
      VSFlexGrid8.Visible = True
      Frame14.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ  ”ÃÌ· «·Õ÷Ê— "
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Record Attendance"
      lbl(13).Caption = " Date"
   End If
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch8
   Dcombos.GetStudentCurs Me.DcbCurs
   Dcombos.GetStudentClassRooms Me.DcbHall
   Dcombos.GetStudentGroup Me.DcbGroup
   Dcombos.GeInstructor Me.DcbInstrucor
End Sub
Sub Relod6()
      VSFlexGrid6.Visible = True
      Frame11.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ «·„Ê«›ﬁ… ⁄·Ï «· —‘ÌÕ "
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Approve the Nomination"
      lbl(13).Caption = " Date"
   End If
   TypeTrain(2).value = True
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch6
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany6
End Sub

Sub Relod3()
      VSFlexGrid3.Visible = True
      Frame6.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ ÿ·»  œ—Ì»"
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Training Request"
      lbl(13).Caption = " Date"
   End If
   TypeTrain(2).value = True
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetStudentQualification Me.DcbQuali3
   Dcombos.GETNationality Me.DcbNationality
End Sub
Sub Relod4()
      VSFlexGrid4.Visible = True
      Frame7.Visible = True
   If SystemOptions.UserInterface = ArabicInterface Then
      Label1(2).Caption = "»ÕÀ «·⁄ﬁÊœ"
      lbl(13).Caption = " «—ÌŒ «·Õ—ﬂ…"
   Else
      Label1(2).Caption = "Search Contracts"
      lbl(13).Caption = " Date"
   End If
   ContType(2).value = True
   Dim Dcombos As New ClsDataCombos
   Dcombos.GetBranches Me.DcbBranch4
   Dcombos.GetStudent Me.DcbStudent
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany4
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text3.Text, EmpID
        DcbCompany6.BoundText = EmpID
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text3.Text, EmpID
        DcbCompany5.BoundText = EmpID
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetInstructorCode EmpID, Text4.Text, 1
        DcbInstrucor.BoundText = EmpID
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text5.Text, EmpID
        DcbCompany9.BoundText = EmpID
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
 Me.DcbEmployee.BoundText = GeTEmpIDByEmpCode(Text6.Text, True)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, Text7.Text, 1
        DcbStudent9.BoundText = EmpID
        End If
End Sub

Private Sub ToValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.ToValue.Text, 0)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode TxtCode.Text, EmpID
        DcbCompany4.BoundText = EmpID
    End If
End Sub

Private Sub TxtSudCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, TxtSudCode.Text, 1
        DcbStudent.BoundText = EmpID
    End If
End Sub

Private Sub VSFlexGrid1_Click()
If inde = 1 Then
FrmStudents.FindRec val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 1))
ElseIf inde = 101 Then
FrmContStudent.DcbStudent.BoundText = val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 1))
ElseIf inde = 102 Then
FrmStudentCalling.DcbStudent.BoundText = val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 1))
ElseIf inde = 103 Then
FrmReportsStudent.DcbStudent.BoundText = val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 1))
End If
End Sub

Private Sub VSFlexGrid10_Click()
If inde = 10 Then
FrmItemsClass.FindRec val(VSFlexGrid10.TextMatrix(VSFlexGrid10.Row, 1))
End If
End Sub

Private Sub VSFlexGrid2_Click()
If inde = 2 Then
FrmInstructors.FindRec val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 1))
ElseIf inde = 201 Then
FrmGroupStudents.DcbInstrucor.BoundText = val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 1))
ElseIf inde = 202 Then
FrmAttendance.DcbInstrucor.BoundText = val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 1))
ElseIf inde = 203 Then
FrmReportsStudent.instruDBox.BoundText = val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 1))
End If
End Sub

Private Sub VSFlexGrid3_Click()
If inde = 3 Then
FrmTrainingRequest.FindRec val(VSFlexGrid3.TextMatrix(VSFlexGrid3.Row, 1))
End If
End Sub

Private Sub VSFlexGrid4_Click()
If inde = 4 Then
FrmContStudent.FindRec val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
ElseIf inde = 401 Then
FrmStudentsCandidacy.TxtContCode.Text = (VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, VSFlexGrid4.ColIndex("FullCode")))
End If
End Sub

Private Sub VSFlexGrid5_Click()
If inde = 5 Then
FrmStudentsCandidacy.FindRec val(VSFlexGrid5.TextMatrix(VSFlexGrid5.Row, 1))
ElseIf inde = 501 Then
FrmStudCandidAccept.TxtCandidacyID.Text = val(VSFlexGrid5.TextMatrix(VSFlexGrid5.Row, 1))
End If
End Sub

Private Sub VSFlexGrid6_Click()
If inde = 6 Then
FrmStudCandidAccept.FindRec val(VSFlexGrid6.TextMatrix(VSFlexGrid6.Row, 1))
End If
End Sub

Private Sub VSFlexGrid7_Click()
If inde = 7 Then
FrmGroupStudents.FindRec val(VSFlexGrid7.TextMatrix(VSFlexGrid7.Row, 1))
ElseIf inde = 701 Then
FrmAttendance.DcbInstrucor.BoundText = val(VSFlexGrid7.TextMatrix(VSFlexGrid7.Row, 1))
ElseIf inde = 702 Then
FrmReportsStudent.groupDBox.BoundText = val(VSFlexGrid7.TextMatrix(VSFlexGrid7.Row, 1))
End If
End Sub

Private Sub VSFlexGrid8_Click()
If inde = 8 Then
FrmAttendance.FindRec val(VSFlexGrid8.TextMatrix(VSFlexGrid8.Row, 1))
End If
End Sub

Private Sub VSFlexGrid9_Click()
If inde = 9 Then
FrmStudentCalling.FindRec val(VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, 1))
ElseIf inde = 20 Then
    
    frmsalebill3.TxtCusID = VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, VSFlexGrid9.ColIndex("CusID"))
    frmsalebill3.DBCboClientName.BoundText = VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, VSFlexGrid9.ColIndex("CusID"))
    frmsalebill3.TxtCashCustomerName.Enabled = True
    frmsalebill3.TxtCashCustomerName.Text = VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, VSFlexGrid9.ColIndex("CusName"))
     'frmsalebill3.TxtCashCustomerName2.Text = VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, VSFlexGrid9.ColIndex("CusName"))
    
    frmsalebill3.txtCallingID = VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, VSFlexGrid9.ColIndex("ID"))
    frmsalebill3.BookingDate.value = VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, VSFlexGrid9.ColIndex("RecordDate"))
    frmsalebill3.TxtPhone = VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, VSFlexGrid9.ColIndex("Mobile"))
    frmsalebill3.TxtPhone = VSFlexGrid9.TextMatrix(VSFlexGrid9.Row, VSFlexGrid9.ColIndex("Mobile"))
    
    
    
   
End If
End Sub
