VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Begin VB.Form FrmEmployee 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13815
   HelpContextID   =   70
   Icon            =   "FrmProffisionals.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   13815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   975
      Left            =   17040
      TabIndex        =   167
      Top             =   4320
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox TxtEmp_Comm 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   50
         TabIndex        =   169
         Top             =   0
         Width           =   2295
      End
      Begin VB.TextBox TxtEmpProfitCom 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   50
         TabIndex        =   168
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·⁄„Ê·… ⁄·Ï ≈Ã„«·Ï «·„»Ì⁄« "
         Height          =   405
         Index           =   5
         Left            =   2640
         TabIndex        =   173
         Top             =   90
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·⁄„Ê·… ⁄·Ï ’«›Ï «·„»Ì⁄« "
         Height          =   405
         Index           =   8
         Left            =   2640
         TabIndex        =   172
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   255
         Index           =   52
         Left            =   0
         TabIndex        =   171
         Top             =   60
         Width           =   135
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   255
         Index           =   56
         Left            =   0
         TabIndex        =   170
         Top             =   420
         Width           =   135
      End
   End
   Begin MSDataListLib.DataCombo DcCostCenter 
      Height          =   315
      Left            =   16920
      TabIndex        =   163
      Top             =   120
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   495
      Index           =   8
      Left            =   18960
      TabIndex        =   164
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Text            =   "„—ﬂ“ «· ﬂ·›…"
      BackColor       =   16777152
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
      Left            =   16080
      TabIndex        =   91
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      Text            =   "«·„‘—Ê⁄"
      BackColor       =   16777152
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
   Begin VB.TextBox XPTxtEmpNamee 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   17400
      TabIndex        =   166
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame framx 
      BackColor       =   &H00FFFFC0&
      Caption         =   "„›—œ«  «·—« »"
      Height          =   3495
      Left            =   19920
      TabIndex        =   104
      Top             =   -3000
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command7 
         Caption         =   "Õ”«»"
         Height          =   315
         Left            =   2760
         TabIndex        =   161
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check13 
         Height          =   195
         Left            =   2640
         TabIndex        =   160
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check12 
         Height          =   195
         Left            =   2640
         TabIndex        =   159
         Top             =   2040
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check11 
         Height          =   195
         Left            =   2640
         TabIndex        =   158
         Top             =   1680
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check10 
         Height          =   195
         Left            =   2640
         TabIndex        =   157
         Top             =   1320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check9 
         Height          =   195
         Left            =   2640
         TabIndex        =   156
         Top             =   960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check8 
         Height          =   195
         Left            =   240
         TabIndex        =   155
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check7 
         Height          =   195
         Left            =   240
         TabIndex        =   154
         Top             =   2040
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check6 
         Height          =   195
         Left            =   240
         TabIndex        =   153
         Top             =   1680
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check5 
         Height          =   195
         Left            =   240
         TabIndex        =   152
         Top             =   1320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check4 
         Height          =   195
         Left            =   240
         TabIndex        =   151
         Top             =   840
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CheckBox Check3 
         Height          =   195
         Left            =   240
         TabIndex        =   150
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox TXTMANG 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   148
         Top             =   1950
         Width           =   1335
      End
      Begin VB.TextBox TXTMANGM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   147
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TXTMOB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   145
         Top             =   1590
         Width           =   1335
      End
      Begin VB.TextBox TXTMOBM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   144
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Height          =   195
         Left            =   2640
         TabIndex        =   143
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "«Œ›«¡"
         Height          =   315
         Left            =   600
         TabIndex        =   137
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   124
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   123
         Top             =   4800
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   122
         Top             =   4440
         Width           =   615
      End
      Begin VB.Frame Frame5 
         Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
         Height          =   3255
         Left            =   240
         TabIndex        =   120
         Top             =   4800
         Visible         =   0   'False
         Width           =   4935
         Begin VB.CommandButton Command6 
            Caption         =   "„Ê«›ﬁ"
            Height          =   315
            Left            =   120
            TabIndex        =   141
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            TabIndex        =   138
            Top             =   1560
            Width           =   4695
         End
         Begin VB.Frame Frame6 
            Height          =   495
            Left            =   240
            TabIndex        =   130
            Top             =   600
            Width           =   1935
            Begin VB.CommandButton Command1 
               Caption         =   "/"
               Height          =   315
               Index           =   5
               Left            =   480
               TabIndex        =   135
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "*"
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   134
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "-"
               Height          =   315
               Index           =   2
               Left            =   1200
               TabIndex        =   133
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "+"
               Height          =   315
               Index           =   3
               Left            =   1560
               TabIndex        =   132
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "="
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   131
               Top             =   120
               Width           =   375
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«Õ„«·Ì «·„ﬁ«„"
               ForeColor       =   &H000000FF&
               Height          =   15
               Index           =   55
               Left            =   0
               TabIndex        =   136
               Top             =   2520
               Width           =   1935
            End
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            TabIndex        =   129
            Top             =   1200
            Width           =   4695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "”‰ÊÌ"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   127
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "‘Â—Ì"
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   126
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "FrmProffisionals.frx":038A
            Left            =   2160
            List            =   "FrmProffisionals.frx":0391
            TabIndex        =   121
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
            TabIndex        =   140
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
            TabIndex        =   139
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
            TabIndex        =   128
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
            TabIndex        =   125
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   119
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtanotherm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   118
         Top             =   2370
         Width           =   1335
      End
      Begin VB.TextBox txtfoodm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   117
         Top             =   1170
         Width           =   1335
      End
      Begin VB.TextBox txtbusm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   116
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtsaknm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   115
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtanother 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   112
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtfood 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   110
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtbus 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   108
         Top             =   750
         Width           =   1335
      End
      Begin VB.TextBox txtsakn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   106
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
         TabIndex        =   149
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
         TabIndex        =   146
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   111
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
         TabIndex        =   109
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
         TabIndex        =   107
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
         TabIndex        =   105
         Top             =   360
         Width           =   1275
      End
   End
   Begin MSDataListLib.DataCombo dcproject 
      Height          =   315
      Left            =   14880
      TabIndex        =   162
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox TxtAccountCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   18720
      MaxLength       =   10
      TabIndex        =   142
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Index           =   7
      Left            =   19680
      TabIndex        =   103
      Top             =   2760
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "FrmProffisionals.frx":0398
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
      Left            =   14760
      TabIndex        =   102
      Top             =   2040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "FrmProffisionals.frx":03B4
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
      Left            =   18960
      TabIndex        =   101
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "FrmProffisionals.frx":03D0
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
      Left            =   14760
      TabIndex        =   100
      Top             =   1680
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "FrmProffisionals.frx":03EC
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
      Left            =   18960
      TabIndex        =   99
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "FrmProffisionals.frx":0408
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00C0C0C0&
      Caption         =   "«” ⁄·«„« "
      Height          =   975
      Index           =   25
      Left            =   13800
      TabIndex        =   93
      Top             =   4560
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton CommandÛQRY 
         Caption         =   "«” ⁄·«„  "
         Height          =   315
         Left            =   240
         TabIndex        =   97
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptExpirLinc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "«‰ Â«¡ «·—Œ’…"
         Height          =   255
         Left            =   1800
         TabIndex        =   96
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptExpirPas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "«‰ Â«¡ ÃÊ«“ «·”›—"
         Height          =   255
         Left            =   1560
         TabIndex        =   95
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptExpirEkama 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "«‰ Â«¡ «·«ﬁ«„…"
         Height          =   255
         Left            =   360
         TabIndex        =   94
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   495
      Index           =   0
      Left            =   17520
      TabIndex        =   88
      Top             =   2760
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      Text            =   "«·„ƒÂ·«  Ê«·Œ»—« "
      BackColor       =   16777152
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
      Left            =   15360
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CboWorkState 
      Enabled         =   0   'False
      Height          =   315
      Left            =   21060
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   2700
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   705
      Left            =   9900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Top             =   10140
      Visible         =   0   'False
      Width           =   2535
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   9840
      TabIndex        =   78
      ToolTipText     =   "⁄‰œ «‰‘«¡ „ÊŸ› Ì „ «‰‘«¡ 3 Õ”«»«  «·Ì… ·Â ÊÂ„  Õ”«» «·–„„ Ê Õ”«» «·«ÃÊ— «·„” Õﬁ… ÊÕ”«» «·„Œ’’« "
      Top             =   9120
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
      Left            =   9075
      TabIndex        =   79
      Top             =   9120
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
      Left            =   8355
      TabIndex        =   80
      Top             =   9120
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
      Left            =   7635
      TabIndex        =   81
      Top             =   9120
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
      Left            =   6825
      TabIndex        =   36
      Top             =   9120
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
      Left            =   6000
      TabIndex        =   82
      Top             =   9120
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
      Left            =   4350
      TabIndex        =   83
      Top             =   9120
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
      Left            =   5160
      TabIndex        =   84
      Top             =   9120
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
      Left            =   18300
      TabIndex        =   85
      Top             =   9120
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
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   615
      Index           =   1
      Left            =   17520
      TabIndex        =   89
      Top             =   3120
      Visible         =   0   'False
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
      Height          =   615
      Index           =   3
      Left            =   17280
      TabIndex        =   90
      Top             =   6480
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
      Height          =   495
      Index           =   6
      Left            =   17400
      TabIndex        =   92
      Top             =   7320
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
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   495
      Index           =   7
      Left            =   20760
      TabIndex        =   98
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
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   8535
      Left            =   0
      TabIndex        =   174
      Top             =   480
      Width           =   13785
      _cx             =   24315
      _cy             =   15055
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   128
      FrontTabColor   =   14871017
      BackTabColor    =   8454143
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   $"FrmProffisionals.frx":0424
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
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin C1SizerLibCtl.C1Elastic C1Elastic9 
         Height          =   8115
         Left            =   16830
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·—« »"
            Height          =   2175
            Index           =   23
            Left            =   2280
            TabIndex        =   386
            Top             =   120
            Width           =   4365
            Begin VB.OptionButton OptSalaryType 
               BackColor       =   &H00E2E9E9&
               Height          =   375
               Index           =   4
               Left            =   3840
               TabIndex        =   419
               Top             =   1680
               Width           =   375
            End
            Begin VB.OptionButton OptSalaryType 
               BackColor       =   &H00E2E9E9&
               Height          =   375
               Index           =   0
               Left            =   3840
               TabIndex        =   392
               Top             =   240
               Width           =   375
            End
            Begin VB.OptionButton OptSalaryType 
               BackColor       =   &H00E2E9E9&
               Height          =   375
               Index           =   1
               Left            =   3840
               TabIndex        =   391
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox TxtPercentage 
               Height          =   285
               Left            =   960
               TabIndex        =   390
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton OptSalaryType 
               BackColor       =   &H00E2E9E9&
               Height          =   375
               Index           =   2
               Left            =   3840
               TabIndex        =   389
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox TxtBYHour 
               Height          =   285
               Left            =   960
               TabIndex        =   388
               Top             =   1080
               Width           =   855
            End
            Begin VB.OptionButton OptSalaryType 
               BackColor       =   &H00E2E9E9&
               Height          =   375
               Index           =   3
               Left            =   3840
               TabIndex        =   387
               Top             =   1320
               Width           =   375
            End
            Begin VB.Label LblSalary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ··„‘—Ê⁄"
               Height          =   255
               Index           =   4
               Left            =   1920
               TabIndex        =   420
               Top             =   1800
               Width           =   1815
            End
            Begin VB.Label LblSalary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—« » ‘Â—Ì"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   396
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label LblSalary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»«·⁄„Ê·…     Õœœ «·‰”»…"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   395
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label LblSalary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»«·”«⁄Â Õœœ ﬁÌ„… «·”«⁄Â"
               Height          =   255
               Index           =   2
               Left            =   1920
               TabIndex        =   394
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label LblSalary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«‰ «Ã"
               Height          =   255
               Index           =   3
               Left            =   1920
               TabIndex        =   393
               Top             =   1440
               Width           =   1815
            End
         End
         Begin VB.Frame Fra 
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
            Height          =   7995
            Index           =   7
            Left            =   6720
            TabIndex        =   191
            Top             =   120
            Width           =   6855
            Begin VB.TextBox TxtBank 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   420
               Index           =   0
               Left            =   720
               TabIndex        =   449
               Top             =   7440
               Width           =   4455
            End
            Begin VB.TextBox txtPrefNatID 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   420
               Left            =   720
               MaxLength       =   24
               TabIndex        =   441
               Top             =   6960
               Width           =   4455
            End
            Begin VB.TextBox TxtBank 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   435
               Index           =   2
               Left            =   720
               MaxLength       =   50
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   407
               Top             =   6480
               Width           =   4455
            End
            Begin VB.TextBox TxtBank 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   435
               Index           =   1
               Left            =   720
               MaxLength       =   24
               TabIndex        =   405
               Top             =   6120
               Width           =   4455
            End
            Begin VB.ComboBox cboPayType 
               Height          =   315
               Left            =   3720
               TabIndex        =   385
               Top             =   4320
               Width           =   1455
            End
            Begin VB.TextBox TxtBankCard 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   435
               Left            =   720
               MaxLength       =   50
               TabIndex        =   373
               Top             =   5640
               Width           =   4455
            End
            Begin VB.Frame Fra 
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ  „Œ’’«    –«ﬂ—"
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
               Index           =   14
               Left            =   3720
               TabIndex        =   372
               Top             =   2880
               Width           =   3075
               Begin VB.TextBox TxtOpenBalance5 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  TabIndex        =   382
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.OptionButton OptType5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   377
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.OptionButton OptType5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   376
                  Top             =   240
                  Width           =   945
               End
               Begin VB.OptionButton OptType5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   2040
                  TabIndex        =   375
                  Top             =   240
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   345
                  Index           =   16
                  Left            =   1740
                  TabIndex        =   383
                  Top             =   630
                  Width           =   1125
               End
            End
            Begin VB.Frame Fra 
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
               Height          =   975
               Index           =   8
               Left            =   3660
               TabIndex        =   209
               Top             =   600
               Width           =   3075
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  TabIndex        =   213
                  Top             =   210
                  Width           =   915
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1110
                  TabIndex        =   212
                  Top             =   210
                  Width           =   915
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1950
                  TabIndex        =   211
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.TextBox TxtOpenBalance 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   150
                  TabIndex        =   210
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   345
                  Index           =   14
                  Left            =   1770
                  TabIndex        =   214
                  Top             =   510
                  Width           =   1125
               End
            End
            Begin VB.Frame Fra 
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
               Height          =   975
               Index           =   9
               Left            =   480
               TabIndex        =   203
               Top             =   600
               Width           =   3075
               Begin VB.TextBox TxtOpenBalance1 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   150
                  TabIndex        =   207
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1950
                  TabIndex        =   206
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1110
                  TabIndex        =   205
                  Top             =   210
                  Width           =   915
               End
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  TabIndex        =   204
                  Top             =   210
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   345
                  Index           =   15
                  Left            =   1770
                  TabIndex        =   208
                  Top             =   510
                  Width           =   1125
               End
            End
            Begin VB.Frame Fra 
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ  „Œ’’«  «·«Ã«“…"
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
               TabIndex        =   199
               Top             =   1560
               Width           =   3075
               Begin VB.TextBox TxtOpenBalance2 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  TabIndex        =   380
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  TabIndex        =   202
                  Top             =   210
                  Width           =   915
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   990
                  TabIndex        =   201
                  Top             =   210
                  Width           =   915
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1920
                  TabIndex        =   200
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   345
                  Index           =   18
                  Left            =   1740
                  TabIndex        =   381
                  Top             =   630
                  Width           =   1125
               End
            End
            Begin VB.TextBox txtopening_balance_voucher_id 
               Height          =   735
               Left            =   960
               TabIndex        =   198
               Top             =   -240
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Frame Fra 
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ  „Œ’’«  ‰Â«Ì… Œœ„…"
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
               Index           =   11
               Left            =   480
               TabIndex        =   192
               Top             =   1560
               Width           =   3075
               Begin VB.TextBox TxtOpenBalance4 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   150
                  TabIndex        =   196
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.OptionButton OptType4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1920
                  TabIndex        =   195
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.OptionButton OptType4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   960
                  TabIndex        =   194
                  Top             =   240
                  Width           =   915
               End
               Begin VB.OptionButton OptType4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  TabIndex        =   193
                  Top             =   210
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   345
                  Index           =   24
                  Left            =   1770
                  TabIndex        =   197
                  Top             =   510
                  Width           =   1125
               End
            End
            Begin MSComCtl2.DTPicker Dtp 
               Height          =   330
               Left            =   3840
               TabIndex        =   378
               Top             =   240
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               CalendarBackColor=   12648447
               CustomFormat    =   "yyyy/M/d"
               Format          =   242089987
               CurrentDate     =   38718
            End
            Begin MSDataListLib.DataCombo DcbBanck 
               Height          =   315
               Index           =   0
               Left            =   720
               TabIndex        =   431
               Top             =   4800
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBanck 
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   432
               Top             =   5160
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·„” ›Ìœ"
               Height          =   345
               Index           =   63
               Left            =   5160
               TabIndex        =   450
               Top             =   7440
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·»ÿ«ﬁ… «·„Œ ’—"
               Height          =   345
               Index           =   57
               Left            =   5310
               TabIndex        =   442
               Top             =   7020
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ⁄‰Ê«‰ «·»‰ﬂ"
               Height          =   345
               Index           =   40
               Left            =   5280
               TabIndex        =   408
               Top             =   6600
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·«Ì»«‰   "
               Height          =   345
               Index           =   31
               Left            =   5280
               TabIndex        =   406
               Top             =   6240
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   345
               Index           =   25
               Left            =   5280
               TabIndex        =   404
               Top             =   5280
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂÊœ «·»‰ﬂ"
               Height          =   345
               Index           =   11
               Left            =   5280
               TabIndex        =   401
               Top             =   4800
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·”œ«œ"
               Height          =   345
               Index           =   17
               Left            =   5280
               TabIndex        =   384
               Top             =   4320
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «· ”ÃÌ·"
               Height          =   315
               Index           =   13
               Left            =   5460
               TabIndex        =   379
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " —ﬁ„ «·Õ”«»/«·’—«›"
               Height          =   345
               Index           =   19
               Left            =   5280
               TabIndex        =   374
               Top             =   5760
               Width           =   1485
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8115
         Index           =   2
         Left            =   45
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   8115
            Left            =   0
            TabIndex        =   234
            TabStop         =   0   'False
            Top             =   -240
            Width           =   14655
            _cx             =   25850
            _cy             =   14314
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
            GridRows        =   10
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
            Begin VB.TextBox TxtEmpNotes 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H000000FF&
               Height          =   1035
               Left            =   0
               MaxLength       =   50
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   370
               Top             =   2160
               Width           =   1215
            End
            Begin VB.TextBox TxtSalary 
               Height          =   285
               Left            =   240
               TabIndex        =   367
               Text            =   "Text10"
               Top             =   240
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Frame Fra 
               Caption         =   " «—ÌŒ «Œ— „»«‘—…/«Ã«“…"
               Height          =   615
               Index           =   20
               Left            =   1200
               TabIndex        =   355
               Top             =   7560
               Width           =   6135
               Begin Dynamic_Byte.NourHijriCal lastHolidaydateH 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   356
                  Top             =   240
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin MSComCtl2.DTPicker lastHolidaydate 
                  Height          =   315
                  Left            =   2520
                  TabIndex        =   357
                  Top             =   240
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DateMoveNo 
                  Height          =   315
                  Left            =   -960
                  TabIndex        =   362
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ «Œ— „»«‘—…/«Ã«“…"
                  Height          =   405
                  Index           =   39
                  Left            =   4260
                  TabIndex        =   358
                  Top             =   240
                  Width           =   1635
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "»Ì«‰«  «·„»«‘—…"
               Height          =   615
               Left            =   7320
               TabIndex        =   348
               Top             =   7560
               Width           =   6255
               Begin Dynamic_Byte.NourHijriCal IssueDateH 
                  Height          =   315
                  Left            =   2400
                  TabIndex        =   349
                  Top             =   240
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   350
                  Top             =   240
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal LastDateH 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   351
                  Top             =   840
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin MSComCtl2.DTPicker LastDate 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   352
                  Top             =   840
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Œ— „»«‘—…"
                  Height          =   285
                  Index           =   21
                  Left            =   2700
                  TabIndex        =   354
                  Top             =   840
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ê· „»«‘—…"
                  Height          =   285
                  Index           =   9
                  Left            =   4980
                  TabIndex        =   353
                  Top             =   240
                  Width           =   795
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—Œ’… «·ﬁÌ«œ…"
               ForeColor       =   &H000000C0&
               Height          =   975
               Index           =   12
               Left            =   1320
               TabIndex        =   333
               Top             =   5520
               Width           =   5955
               Begin VB.TextBox TxtDriverLicense 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2880
                  MaxLength       =   30
                  TabIndex        =   334
                  Top             =   240
                  Width           =   1815
               End
               Begin Dynamic_Byte.NourHijriCal txtDriverLicenseStartdH 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   335
                  Top             =   240
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin Dynamic_Byte.NourHijriCal txtDriverLicenseendH 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   336
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin MSComCtl2.DTPicker DpDriverLicenseend 
                  Height          =   315
                  Left            =   2880
                  TabIndex        =   337
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTPicker5 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   338
                  Top             =   960
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«‰ Â«¡"
                  Height          =   285
                  Index           =   58
                  Left            =   4680
                  TabIndex        =   342
                  Top             =   600
                  Width           =   885
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·—Œ’…"
                  Height          =   285
                  Index           =   57
                  Left            =   4530
                  TabIndex        =   341
                  Top             =   300
                  Width           =   1065
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«’œ«—"
                  Height          =   285
                  Index           =   56
                  Left            =   1710
                  TabIndex        =   340
                  Top             =   300
                  Width           =   1005
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«‰ Â«¡"
                  Height          =   285
                  Index           =   55
                  Left            =   1860
                  TabIndex        =   339
                  Top             =   600
                  Width           =   885
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  ’«Õ» «·⁄„·"
               ForeColor       =   &H000000C0&
               Height          =   1095
               Index           =   6
               Left            =   1200
               TabIndex        =   267
               Top             =   6480
               Width           =   6075
               Begin VB.TextBox txtkafeladd 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   240
                  MaxLength       =   150
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   34
                  Top             =   600
                  Width           =   2565
               End
               Begin VB.TextBox txtkafeltel 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   3420
                  MaxLength       =   30
                  TabIndex        =   33
                  Top             =   600
                  Width           =   1845
               End
               Begin VB.TextBox txtKafelID 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   3420
                  MaxLength       =   30
                  TabIndex        =   32
                  Top             =   240
                  Width           =   1845
               End
               Begin MSDataListLib.DataCombo DcbKafelName 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   430
                  Top             =   240
                  Width           =   2565
                  _ExtentX        =   4524
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄‰Ê«‰"
                  Height          =   285
                  Index           =   33
                  Left            =   2400
                  TabIndex        =   271
                  Top             =   600
                  Width           =   1005
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ·Ì›Ê‰"
                  Height          =   285
                  Index           =   32
                  Left            =   5010
                  TabIndex        =   270
                  Top             =   660
                  Width           =   1005
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—ﬁ„"
                  Height          =   285
                  Index           =   26
                  Left            =   4890
                  TabIndex        =   269
                  Top             =   300
                  Width           =   1065
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„"
                  Height          =   285
                  Index           =   25
                  Left            =   2670
                  TabIndex        =   268
                  Top             =   300
                  Width           =   645
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÃÊ«“ «·”›—"
               ForeColor       =   &H000000C0&
               Height          =   1425
               Index           =   5
               Left            =   7320
               TabIndex        =   261
               Top             =   6180
               Width           =   6195
               Begin VB.TextBox Txt_NumPasp 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   3000
                  MaxLength       =   30
                  TabIndex        =   28
                  Top             =   240
                  Width           =   1845
               End
               Begin MSComCtl2.DTPicker Txt_DateExpPasp 
                  Height          =   315
                  Left            =   3000
                  TabIndex        =   29
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker Txt_DatePasp 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   30
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcboJobsType2 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   31
                  Top             =   1080
                  Width           =   4575
                  _ExtentX        =   8070
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbPasplace 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   433
                  Top             =   240
                  Width           =   1845
                  _ExtentX        =   3254
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Â‰… ›Ì «·ÃÊ«“"
                  Height          =   285
                  Index           =   51
                  Left            =   4680
                  TabIndex        =   266
                  Top             =   1080
                  Width           =   1365
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ﬂ«‰ «·«’œ«—"
                  Height          =   285
                  Index           =   24
                  Left            =   1920
                  TabIndex        =   265
                  Top             =   300
                  Width           =   1005
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ ÃÊ«“ «·”›— "
                  Height          =   285
                  Index           =   23
                  Left            =   4800
                  TabIndex        =   264
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«’œ«—"
                  Height          =   285
                  Index           =   22
                  Left            =   4680
                  TabIndex        =   263
                  Top             =   660
                  Width           =   1365
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«‰ Â«¡"
                  Height          =   285
                  Index           =   21
                  Left            =   1920
                  TabIndex        =   262
                  Top             =   660
                  Width           =   1005
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÂÊÌ…"
               ForeColor       =   &H000000C0&
               Height          =   975
               Index           =   4
               Left            =   1320
               TabIndex        =   252
               Top             =   4620
               Width           =   5955
               Begin VB.TextBox Tet_NumPoket 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2880
                  MaxLength       =   30
                  TabIndex        =   253
                  Top             =   240
                  Width           =   1815
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateExppoketH 
                  Height          =   255
                  Left            =   2880
                  TabIndex        =   254
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateEndpoketH 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   255
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin MSComCtl2.DTPicker Txt_DateExppoket 
                  Height          =   315
                  Left            =   2880
                  TabIndex        =   256
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker Txt_DateEndpoket 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   257
                  Top             =   960
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«‰ Â«¡"
                  Height          =   405
                  Index           =   20
                  Left            =   1860
                  TabIndex        =   260
                  Top             =   720
                  Width           =   885
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«’œ«—"
                  Height          =   285
                  Index           =   19
                  Left            =   4590
                  TabIndex        =   259
                  Top             =   660
                  Width           =   1005
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ÂÊÌ…"
                  Height          =   285
                  Index           =   18
                  Left            =   4530
                  TabIndex        =   258
                  Top             =   300
                  Width           =   1065
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—Œ’… «·⁄„·"
               ForeColor       =   &H000000C0&
               Height          =   915
               Index           =   2
               Left            =   1260
               TabIndex        =   246
               Top             =   3840
               Width           =   5955
               Begin VB.TextBox Txt_NumLicn 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   3000
                  MaxLength       =   30
                  TabIndex        =   23
                  Top             =   120
                  Width           =   1815
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateEndLincH 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   25
                  Top             =   480
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateExpLincH 
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   24
                  Top             =   480
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin MSComCtl2.DTPicker Txt_DateExpLinc 
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   247
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker Txt_DateEndLinc 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   248
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   238616577
                  CurrentDate     =   38784
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·—Œ’…"
                  Height          =   285
                  Index           =   13
                  Left            =   4650
                  TabIndex        =   251
                  Top             =   300
                  Width           =   1065
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«’œ«—"
                  Height          =   285
                  Index           =   12
                  Left            =   4710
                  TabIndex        =   250
                  Top             =   540
                  Width           =   1005
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - «·«‰ Â«¡"
                  Height          =   285
                  Index           =   11
                  Left            =   1860
                  TabIndex        =   249
                  Top             =   480
                  Width           =   1005
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  Œ«’… »«·⁄„·"
               ForeColor       =   &H000000C0&
               Height          =   1695
               Index           =   1
               Left            =   1320
               TabIndex        =   241
               Top             =   1680
               Width           =   5835
               Begin VB.TextBox TxtRegion 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   60
                  TabIndex        =   242
                  Top             =   1860
                  Width           =   2295
               End
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   18
                  Top             =   600
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   19
                  Top             =   1320
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCRegionID 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   369
                  Top             =   240
                  Width           =   4155
                  _ExtentX        =   7329
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbDepartment2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   426
                  Top             =   960
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ”„"
                  Height          =   225
                  Index           =   5
                  Left            =   4800
                  TabIndex        =   427
                  Top             =   1020
                  Width           =   885
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰ÿﬁÂ/ﬁÿ«⁄/«·”ﬂ‰"
                  Height          =   225
                  Index           =   59
                  Left            =   4320
                  TabIndex        =   368
                  Top             =   240
                  Width           =   1485
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰ÿﬁ…"
                  Height          =   225
                  Index           =   10
                  Left            =   2400
                  TabIndex        =   245
                  Top             =   1890
                  Width           =   735
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   225
                  Index           =   9
                  Left            =   4800
                  TabIndex        =   244
                  Top             =   1410
                  Width           =   885
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—…"
                  Height          =   225
                  Index           =   7
                  Left            =   4800
                  TabIndex        =   243
                  Top             =   660
                  Width           =   885
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«ﬁ«„…"
               ForeColor       =   &H000000C0&
               Height          =   1665
               Index           =   3
               Left            =   7320
               TabIndex        =   235
               Top             =   4530
               Width           =   6195
               Begin VB.TextBox Txt_placEkama 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   300
                  MaxLength       =   50
                  TabIndex        =   411
                  Top             =   240
                  Width           =   1845
               End
               Begin VB.TextBox Txt_NumEkama 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   3390
                  MaxLength       =   30
                  TabIndex        =   26
                  Top             =   240
                  Width           =   1845
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateExpEkamaH 
                  Height          =   255
                  Left            =   3390
                  TabIndex        =   27
                  Top             =   600
                  Width           =   1845
                  _ExtentX        =   3254
                  _ExtentY        =   661
               End
               Begin VB.PictureBox Picture2 
                  Height          =   0
                  Left            =   0
                  ScaleHeight     =   0
                  ScaleWidth      =   0
                  TabIndex        =   410
                  Top             =   0
                  Width           =   0
               End
               Begin Dynamic_Byte.NourHijriCal Txt_DateEndekamah 
                  Height          =   255
                  Left            =   300
                  TabIndex        =   412
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin MSDataListLib.DataCombo DcboJobsType3 
                  Height          =   315
                  Left            =   300
                  TabIndex        =   413
                  Top             =   960
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin Dynamic_Byte.NourHijriCal HowIqamaEndH 
                  Height          =   255
                  Left            =   300
                  TabIndex        =   414
                  Top             =   1320
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «‰ Â«¡ «·ÂÊÌ…"
                  Height          =   285
                  Index           =   62
                  Left            =   2040
                  TabIndex        =   415
                  Top             =   1320
                  Width           =   1725
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Â‰… "
                  Height          =   285
                  Index           =   50
                  Left            =   5400
                  TabIndex        =   240
                  Top             =   960
                  Width           =   645
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " -«·«‰ Â«¡"
                  Height          =   285
                  Index           =   17
                  Left            =   2280
                  TabIndex        =   239
                  Top             =   720
                  Width           =   885
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ﬂ«‰ «·«ﬁ«„…"
                  Height          =   285
                  Index           =   16
                  Left            =   1920
                  TabIndex        =   238
                  Top             =   300
                  Width           =   1365
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·«ﬁ«„…"
                  Height          =   285
                  Index           =   15
                  Left            =   4680
                  TabIndex        =   237
                  Top             =   180
                  Width           =   1365
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " -«·«’œ«—"
                  Height          =   285
                  Index           =   14
                  Left            =   5280
                  TabIndex        =   236
                  Top             =   600
                  Width           =   765
               End
            End
            Begin VB.Frame Fra 
               Height          =   4425
               Index           =   21
               Left            =   1320
               TabIndex        =   272
               Top             =   120
               Width           =   12135
               Begin VB.TextBox TxtMachinCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   1
                  Left            =   6000
                  TabIndex        =   428
                  Top             =   3000
                  Width           =   1815
               End
               Begin XtremeSuiteControls.CheckBox ChNoAdded 
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   423
                  Top             =   3720
                  Width           =   1935
                  _Version        =   786432
                  _ExtentX        =   3413
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "·Ì” ·Â «÷«›Ì"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox TxtMachinCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   6000
                  TabIndex        =   422
                  Top             =   2640
                  Width           =   1815
               End
               Begin VB.Frame Fra 
                  Caption         =   "«·‰Ê⁄"
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
                  Height          =   495
                  Index           =   13
                  Left            =   6000
                  TabIndex        =   416
                  Top             =   3240
                  Width           =   1875
                  Begin VB.OptionButton TypeEmp 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„ÊŸ›"
                     Height          =   255
                     Index           =   0
                     Left            =   870
                     TabIndex        =   418
                     Top             =   210
                     Width           =   915
                  End
                  Begin VB.OptionButton TypeEmp 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„‘—›"
                     Height          =   255
                     Index           =   1
                     Left            =   30
                     TabIndex        =   417
                     Top             =   210
                     Value           =   -1  'True
                     Width           =   825
                  End
               End
               Begin VB.CheckBox Chk_Stkala 
                  Caption         =   " —ﬂ «·⁄„·"
                  Height          =   255
                  Left            =   10800
                  TabIndex        =   403
                  Top             =   4380
                  Width           =   1095
               End
               Begin VB.ComboBox DcbMatrial 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   399
                  Top             =   1560
                  Width           =   1935
               End
               Begin VB.ComboBox Dcbsex 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   366
                  Top             =   1200
                  Width           =   1935
               End
               Begin VB.TextBox XPTxtEmpName 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   10980
                  MaxLength       =   50
                  TabIndex        =   274
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.TextBox XPTxtProfMail 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8750
                  MaxLength       =   50
                  TabIndex        =   15
                  Top             =   1530
                  Width           =   1695
               End
               Begin VB.TextBox XPTxtmobile 
                  Height          =   315
                  Left            =   6000
                  MaxLength       =   50
                  TabIndex        =   17
                  Top             =   1890
                  Width           =   1935
               End
               Begin VB.TextBox XPTxtPhone 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8750
                  MaxLength       =   50
                  TabIndex        =   16
                  Top             =   1890
                  Width           =   1695
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9375
                  MaxLength       =   50
                  TabIndex        =   3
                  Top             =   495
                  Width           =   1095
               End
               Begin VB.TextBox Text2 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8280
                  MaxLength       =   50
                  TabIndex        =   4
                  Top             =   495
                  Width           =   1095
               End
               Begin VB.TextBox Text3 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7200
                  MaxLength       =   50
                  TabIndex        =   5
                  Top             =   495
                  Width           =   1095
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6045
                  MaxLength       =   50
                  TabIndex        =   6
                  Top             =   495
                  Width           =   1095
               End
               Begin VB.TextBox TXT_WORK_PLACE 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6600
                  TabIndex        =   273
                  Top             =   -705
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9375
                  MaxLength       =   50
                  TabIndex        =   8
                  Top             =   855
                  Width           =   1095
               End
               Begin VB.TextBox Text6 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8280
                  MaxLength       =   50
                  TabIndex        =   9
                  Top             =   855
                  Width           =   1095
               End
               Begin VB.TextBox Text7 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7200
                  MaxLength       =   50
                  TabIndex        =   10
                  Top             =   855
                  Width           =   1095
               End
               Begin VB.TextBox Text8 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6045
                  MaxLength       =   50
                  TabIndex        =   11
                  Top             =   855
                  Width           =   1095
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9375
                  MaxLength       =   50
                  TabIndex        =   1
                  Top             =   120
                  Width           =   1095
               End
               Begin MSDataListLib.DataCombo DCNationality 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   12
                  Top             =   600
                  Width           =   1845
                  _ExtentX        =   3254
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo Dcdean 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   7
                  Top             =   600
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcjopstatus 
                  Height          =   315
                  Left            =   8745
                  TabIndex        =   20
                  Top             =   3015
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCPreFix 
                  Height          =   315
                  Left            =   8280
                  TabIndex        =   0
                  Top             =   120
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcGrade 
                  Height          =   315
                  Left            =   8750
                  TabIndex        =   13
                  Top             =   1215
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin Dynamic_Byte.NourHijriCal DOBH 
                  Height          =   315
                  Left            =   7920
                  TabIndex        =   275
                  Top             =   2295
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin MSDataListLib.DataCombo DCBranch 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   2
                  Top             =   255
                  Width           =   4785
                  _ExtentX        =   8440
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCGroupID 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   14
                  Top             =   975
                  Width           =   4845
                  _ExtentX        =   8546
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo mangerid 
                  Height          =   315
                  Left            =   8025
                  TabIndex        =   21
                  Top             =   3375
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo swapedempid 
                  Height          =   315
                  Left            =   8025
                  TabIndex        =   22
                  Top             =   3735
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DBDOB 
                  Height          =   315
                  Left            =   9110
                  TabIndex        =   363
                  Top             =   2280
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   197001217
                  CurrentDate     =   38784
                  MinDate         =   -292192
               End
               Begin MSDataListLib.DataCombo DcbSection 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   364
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   4845
                  _ExtentX        =   8546
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbContractType 
                  Height          =   315
                  Left            =   8750
                  TabIndex        =   397
                  Top             =   2640
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DtDate 
                  Height          =   315
                  Left            =   9120
                  TabIndex        =   402
                  Top             =   4380
                  Visible         =   0   'False
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   197001217
                  CurrentDate     =   38784
                  MinDate         =   -292192
               End
               Begin MSDataListLib.DataCombo DcboSpecifications 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   424
                  Top             =   1320
                  Width           =   4845
                  _ExtentX        =   8546
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker txtDateEndIndustrial 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   438
                  Top             =   3360
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   197001217
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal txtDateEndIndustrialHijri 
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   440
                  Top             =   3360
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
               End
               Begin XtremeSuiteControls.CheckBox chkShowTasks 
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   465
                  Top             =   3360
                  Width           =   1335
                  _Version        =   786432
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ÌŸÂ— ›Ï «·„Â«„"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo swapedempid2 
                  Height          =   315
                  Left            =   8040
                  TabIndex        =   466
                  Top             =   4050
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„ÊŸ› «·»œÌ· 3"
                  Height          =   285
                  Index           =   64
                  Left            =   10575
                  TabIndex        =   467
                  Top             =   4050
                  Width           =   1275
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " - ‰Â«Ì… »ÿ«ﬁ… «·«„‰ «·’‰«⁄Ì"
                  Height          =   615
                  Index           =   53
                  Left            =   4560
                  TabIndex        =   439
                  Top             =   3240
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰Ê« "
                  Height          =   285
                  Index           =   48
                  Left            =   6000
                  TabIndex        =   436
                  Top             =   2280
                  Width           =   435
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  Height          =   285
                  Index           =   47
                  Left            =   6600
                  TabIndex        =   435
                  Top             =   2280
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄„—"
                  Height          =   285
                  Index           =   44
                  Left            =   7440
                  TabIndex        =   434
                  Top             =   2280
                  Width           =   435
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ„«Ì… «·«ÃÊ—"
                  Height          =   285
                  Index           =   42
                  Left            =   7800
                  TabIndex        =   429
                  Top             =   3000
                  Width           =   915
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "›—ﬁ «·⁄„·"
                  Height          =   225
                  Index           =   8
                  Left            =   4920
                  TabIndex        =   425
                  Top             =   1350
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬂÊœ «·„ﬂ‰…"
                  Height          =   285
                  Index           =   41
                  Left            =   7440
                  TabIndex        =   421
                  Top             =   2640
                  Width           =   1275
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·Õ«·…"
                  Height          =   285
                  Index           =   61
                  Left            =   8040
                  TabIndex        =   400
                  Top             =   1575
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «· ⁄«ﬁœ"
                  Height          =   285
                  Index           =   10
                  Left            =   10560
                  TabIndex        =   398
                  Top             =   2640
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„‰ÿﬁÂ"
                  Height          =   285
                  Index           =   59
                  Left            =   2880
                  TabIndex        =   365
                  Top             =   1470
                  Visible         =   0   'False
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Index           =   46
                  Left            =   5880
                  TabIndex        =   360
                  Top             =   120
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈Ã„«·Ì «·—« »"
                  Height          =   285
                  Index           =   45
                  Left            =   6960
                  TabIndex        =   359
                  Top             =   135
                  Width           =   1275
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   285
                  Index           =   0
                  Left            =   10560
                  TabIndex        =   291
                  Top             =   480
                  Width           =   1275
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬂÊœ"
                  Height          =   285
                  Index           =   1
                  Left            =   10560
                  TabIndex        =   290
                  Top             =   165
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ÃÊ«·"
                  Height          =   285
                  Index           =   2
                  Left            =   7440
                  TabIndex        =   289
                  Top             =   1905
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·Â« ›"
                  Height          =   285
                  Index           =   3
                  Left            =   10560
                  TabIndex        =   288
                  Top             =   1920
                  Width           =   1275
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»—Ìœ «·«·ﬂ —Ê‰Ì"
                  Height          =   285
                  Index           =   3
                  Left            =   10560
                  TabIndex        =   287
                  Top             =   1560
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ«·… «·⁄„·"
                  Height          =   285
                  Index           =   7
                  Left            =   10560
                  TabIndex        =   286
                  Top             =   3075
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ  «·„Ì·«œ"
                  Height          =   285
                  Index           =   12
                  Left            =   10560
                  TabIndex        =   285
                  Top             =   2295
                  Width           =   1275
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ã‰”Ì…"
                  Height          =   225
                  Index           =   27
                  Left            =   4920
                  TabIndex        =   284
                  Top             =   645
                  Width           =   915
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·œÌ«‰…"
                  Height          =   225
                  Index           =   28
                  Left            =   1950
                  TabIndex        =   283
                  Top             =   600
                  Width           =   915
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·Ã‰”"
                  Height          =   285
                  Index           =   2
                  Left            =   8040
                  TabIndex        =   282
                  Top             =   1215
                  Width           =   675
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„Êﬁ⁄"
                  Height          =   285
                  Index           =   46
                  Left            =   4920
                  TabIndex        =   281
                  Top             =   1005
                  Width           =   915
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   285
                  Index           =   47
                  Left            =   10560
                  TabIndex        =   280
                  Top             =   855
                  Width           =   1275
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„— »…"
                  Height          =   285
                  Index           =   48
                  Left            =   10560
                  TabIndex        =   279
                  Top             =   1215
                  Width           =   1275
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›—⁄"
                  Height          =   225
                  Index           =   52
                  Left            =   4830
                  TabIndex        =   278
                  Top             =   255
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„œÌ— «·„»«‘—"
                  Height          =   285
                  Index           =   22
                  Left            =   10560
                  TabIndex        =   277
                  Top             =   3375
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„ÊŸ› «·»œÌ·"
                  Height          =   285
                  Index           =   23
                  Left            =   10560
                  TabIndex        =   276
                  Top             =   3735
                  Width           =   1275
               End
            End
            Begin DBPIXLib.DBPix20 DBPix201 
               Height          =   1335
               Left            =   0
               TabIndex        =   345
               Top             =   240
               Width           =   1335
               _Version        =   131072
               _ExtentX        =   2355
               _ExtentY        =   2355
               _StockProps     =   1
               BackColor       =   16777152
               _Image          =   "FrmProffisionals.frx":04D9
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
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   495
               Left            =   0
               TabIndex        =   346
               Top             =   1680
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   873
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
            Begin VB.PictureBox Picture1 
               Height          =   0
               Left            =   0
               ScaleHeight     =   0
               ScaleWidth      =   0
               TabIndex        =   409
               Top             =   0
               Width           =   0
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   $"FrmProffisionals.frx":04F1
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
               Height          =   4785
               Index           =   6
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   313
               Top             =   3240
               Width           =   1125
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   315
            Index           =   37
            Left            =   8400
            TabIndex        =   176
            Top             =   90
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8115
         Index           =   0
         Left            =   14430
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         Begin VB.TextBox txtCommission 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5160
            MaxLength       =   30
            TabIndex        =   463
            Top             =   2760
            Width           =   1845
         End
         Begin VB.TextBox Txt_NumPaspOld 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5190
            MaxLength       =   30
            TabIndex        =   461
            Top             =   2100
            Width           =   1845
         End
         Begin VB.ComboBox cmbToM 
            Height          =   315
            Left            =   5220
            TabIndex        =   455
            Top             =   1650
            Width           =   1935
         End
         Begin VB.ComboBox cmbInsuranceRenew 
            Height          =   315
            Left            =   5220
            TabIndex        =   453
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtCopyNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5220
            MaxLength       =   50
            TabIndex        =   451
            Top             =   780
            Width           =   1935
         End
         Begin VB.TextBox txtidxxxx 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   2
            Left            =   -3375
            TabIndex        =   178
            Top             =   13035
            Width           =   1980
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   375
            Index           =   6
            Left            =   10320
            TabIndex        =   317
            Top             =   2760
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   661
            BTYPE           =   2
            TX              =   "«·⁄ﬁÊœ"
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
            MICON           =   "FrmProffisionals.frx":057E
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
            Left            =   10320
            TabIndex        =   318
            Top             =   2280
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   661
            BTYPE           =   2
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
            MICON           =   "FrmProffisionals.frx":059A
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
            Index           =   5
            Left            =   10320
            TabIndex        =   361
            Top             =   3240
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   661
            BTYPE           =   2
            TX              =   "«·“Ì«œ«  «·”‰ÊÌ…"
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
            MICON           =   "FrmProffisionals.frx":05B6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker txtInsuranceRenewDate 
            Height          =   315
            Left            =   1770
            TabIndex        =   457
            Top             =   1200
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Format          =   197001217
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtToMDateNew 
            Height          =   315
            Left            =   1770
            TabIndex        =   459
            Top             =   1650
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Format          =   197001217
            CurrentDate     =   38784
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ê·…"
            Height          =   285
            Index           =   68
            Left            =   7050
            TabIndex        =   464
            Top             =   2820
            Width           =   1545
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ÃÊ«“ «·”›— «·ﬁœÌ„ "
            Height          =   285
            Index           =   65
            Left            =   7080
            TabIndex        =   462
            Top             =   2160
            Width           =   1545
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ÃœÌœ"
            Height          =   285
            Index           =   67
            Left            =   3660
            TabIndex        =   460
            Top             =   1665
            Width           =   1365
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ÃœÌœ"
            Height          =   285
            Index           =   66
            Left            =   3660
            TabIndex        =   458
            Top             =   1215
            Width           =   1365
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ﬁ«»· «·„«·Ï"
            Height          =   285
            Index           =   64
            Left            =   7395
            TabIndex        =   456
            Top             =   1665
            Width           =   1005
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ÃœÌœ «· √„Ì‰"
            Height          =   285
            Index           =   63
            Left            =   7275
            TabIndex        =   454
            Top             =   1215
            Width           =   1125
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·‰”Œ…"
            Height          =   285
            Index           =   54
            Left            =   7125
            TabIndex        =   452
            Top             =   810
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Index           =   53
            Left            =   17070
            TabIndex        =   179
            Top             =   2985
            Width           =   1395
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   8115
         Left            =   14730
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VB.Frame Fra 
            Height          =   7455
            Index           =   15
            Left            =   3960
            TabIndex        =   292
            Top             =   120
            Width           =   9735
            Begin VB.Frame Fra 
               Caption         =   "»Ì«‰«  ««· «»⁄Ì‰"
               Height          =   3015
               Index           =   16
               Left            =   120
               TabIndex        =   293
               Top             =   360
               Width           =   9255
               Begin VB.TextBox insuranceno1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  TabIndex        =   47
                  Top             =   960
                  Width           =   3135
               End
               Begin VB.CheckBox chkhaveinsurance 
                  Alignment       =   1  'Right Justify
                  Caption         =   "·Â  √„Ì‰"
                  Height          =   255
                  Left            =   6120
                  TabIndex        =   46
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.TextBox Txtpassportno 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  TabIndex        =   45
                  Top             =   600
                  Width           =   3135
               End
               Begin VB.TextBox Txtiqamano 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4800
                  TabIndex        =   44
                  Top             =   600
                  Width           =   3135
               End
               Begin VB.TextBox Txtname 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4800
                  TabIndex        =   42
                  Top             =   240
                  Width           =   3135
               End
               Begin VB.TextBox Txtdes1 
                  Alignment       =   1  'Right Justify
                  Height          =   915
                  Left            =   240
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   48
                  Top             =   1440
                  Width           =   7815
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   825
                  TabIndex        =   49
                  Top             =   2520
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmProffisionals.frx":05D2
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   120
                  TabIndex        =   294
                  Top             =   2520
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmProffisionals.frx":096C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcRelation 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   43
                  Top             =   240
                  Width           =   3135
                  _ExtentX        =   5530
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «· √„Ì‰"
                  Height          =   285
                  Index           =   30
                  Left            =   3600
                  TabIndex        =   300
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   285
                  Index           =   29
                  Left            =   8040
                  TabIndex        =   299
                  Top             =   1680
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ÃÊ«“"
                  Height          =   285
                  Index           =   28
                  Left            =   3600
                  TabIndex        =   298
                  Top             =   600
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·«ﬁ«„…"
                  Height          =   285
                  Index           =   27
                  Left            =   8040
                  TabIndex        =   297
                  Top             =   600
                  Width           =   1035
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«”„ «· «»⁄"
                  Height          =   255
                  Left            =   8040
                  TabIndex        =   296
                  Top             =   240
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·’·…"
                  Height          =   285
                  Index           =   26
                  Left            =   3600
                  TabIndex        =   295
                  Top             =   240
                  Width           =   1035
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid Grid 
               Height          =   3645
               Left            =   120
               TabIndex        =   301
               Top             =   3480
               Width           =   9255
               _cx             =   16325
               _cy             =   6429
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
               Rows            =   2
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProffisionals.frx":0F06
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   1
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
               WordWrap        =   -1  'True
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   8115
         Left            =   15030
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VB.Frame Fra 
            Height          =   7455
            Index           =   24
            Left            =   120
            TabIndex        =   321
            Top             =   120
            Width           =   6615
            Begin VB.Frame Frame20 
               Caption         =   "»Ì«‰«  «·Œ»—« "
               Height          =   3015
               Left            =   120
               TabIndex        =   322
               Top             =   360
               Width           =   6375
               Begin VB.TextBox Text27 
                  Alignment       =   1  'Right Justify
                  Height          =   555
                  Left            =   240
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   60
                  Top             =   1800
                  Width           =   5055
               End
               Begin VB.TextBox TxtWorkName 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   56
                  Top             =   240
                  Width           =   3255
               End
               Begin VB.TextBox TxtWorkEntity 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   57
                  Top             =   600
                  Width           =   3255
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   14
                  Left            =   825
                  TabIndex        =   61
                  Top             =   2520
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmProffisionals.frx":1078
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   15
                  Left            =   120
                  TabIndex        =   323
                  Top             =   2520
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmProffisionals.frx":1412
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSComCtl2.DTPicker workfrom 
                  Height          =   315
                  Left            =   3720
                  TabIndex        =   58
                  Top             =   960
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   237109249
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker workto 
                  Height          =   315
                  Left            =   3720
                  TabIndex        =   59
                  Top             =   1320
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   237109249
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal workfromH 
                  Height          =   375
                  Left            =   1800
                  TabIndex        =   329
                  Top             =   960
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   661
               End
               Begin Dynamic_Byte.NourHijriCal worktoH 
                  Height          =   375
                  Left            =   1800
                  TabIndex        =   330
                  Top             =   1320
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   661
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   255
                  Index           =   51
                  Left            =   5160
                  TabIndex        =   328
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«”„ «·Œ»—…"
                  Height          =   255
                  Index           =   62
                  Left            =   5160
                  TabIndex        =   327
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ì"
                  Height          =   255
                  Index           =   50
                  Left            =   5160
                  TabIndex        =   326
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   255
                  Index           =   49
                  Left            =   5160
                  TabIndex        =   325
                  Top             =   1920
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÃÂ…"
                  Height          =   255
                  Index           =   36
                  Left            =   5160
                  TabIndex        =   324
                  Top             =   600
                  Width           =   1095
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid grid02 
               Height          =   3645
               Left            =   0
               TabIndex        =   331
               Top             =   3600
               Width           =   6375
               _cx             =   11245
               _cy             =   6429
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
               Rows            =   2
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProffisionals.frx":19AC
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   1
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
               WordWrap        =   -1  'True
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
         Begin VB.Frame Fra 
            Height          =   7455
            Index           =   17
            Left            =   6960
            TabIndex        =   302
            Top             =   120
            Width           =   6615
            Begin VB.Frame Fra 
               Caption         =   "»Ì«‰«  «·„ƒÂ·«  "
               Height          =   3015
               Index           =   18
               Left            =   120
               TabIndex        =   303
               Top             =   360
               Width           =   6375
               Begin VB.TextBox TxtqualicationEntity 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   52
                  Top             =   600
                  Width           =   3375
               End
               Begin VB.TextBox txtgrade 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  TabIndex        =   51
                  Top             =   240
                  Width           =   975
               End
               Begin VB.TextBox txtyearofqualication 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  TabIndex        =   53
                  Top             =   600
                  Width           =   975
               End
               Begin VB.TextBox txtname1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   50
                  Top             =   240
                  Width           =   3375
               End
               Begin VB.TextBox Txtdes2 
                  Alignment       =   1  'Right Justify
                  Height          =   1275
                  Left            =   240
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   54
                  Top             =   1080
                  Width           =   5175
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   8
                  Left            =   825
                  TabIndex        =   55
                  Top             =   2520
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmProffisionals.frx":1AE3
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   9
                  Left            =   120
                  TabIndex        =   304
                  Top             =   2520
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmProffisionals.frx":1E7D
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÃÂ…"
                  Height          =   285
                  Index           =   34
                  Left            =   5520
                  TabIndex        =   320
                  Top             =   600
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   285
                  Index           =   35
                  Left            =   5520
                  TabIndex        =   308
                  Top             =   1200
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰… « Œ—Ã"
                  Height          =   285
                  Index           =   33
                  Left            =   1200
                  TabIndex        =   307
                  Top             =   600
                  Width           =   795
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«”„ «·„ƒÂ·"
                  Height          =   255
                  Left            =   5280
                  TabIndex        =   306
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬁœÌ—"
                  Height          =   285
                  Index           =   32
                  Left            =   1200
                  TabIndex        =   305
                  Top             =   240
                  Width           =   675
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid grid01 
               Height          =   3645
               Left            =   120
               TabIndex        =   309
               Top             =   3480
               Width           =   6375
               _cx             =   11245
               _cy             =   6429
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
               Rows            =   2
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProffisionals.frx":2417
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   1
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
               WordWrap        =   -1  'True
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   8115
         Left            =   15330
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VB.Frame Fra 
            Caption         =   "«·ﬁÌÌ„"
            Height          =   7455
            Index           =   22
            Left            =   120
            TabIndex        =   310
            Top             =   360
            Width           =   12495
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
               Height          =   6885
               Left            =   120
               TabIndex        =   311
               Top             =   480
               Width           =   12135
               _cx             =   21405
               _cy             =   12144
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
               Rows            =   2
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProffisionals.frx":2500
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   1
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
               WordWrap        =   -1  'True
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   8115
         Left            =   15630
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VB.Frame Fra 
            Caption         =   "«·„·› «·’ÕÌ"
            Height          =   7935
            Index           =   19
            Left            =   -120
            TabIndex        =   221
            Top             =   0
            Width           =   13575
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
               Height          =   7005
               Left            =   120
               TabIndex        =   222
               Top             =   600
               Width           =   13335
               _cx             =   23521
               _cy             =   12356
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
               Rows            =   2
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProffisionals.frx":2669
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   1
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
               WordWrap        =   -1  'True
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   8115
         Left            =   15930
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· √„Ì‰«  «·«Ã „«⁄ÌÂ"
            ForeColor       =   &H000000C0&
            Height          =   2055
            Index           =   0
            Left            =   8640
            TabIndex        =   314
            Top             =   240
            Width           =   4875
            Begin VB.ComboBox CboInsuranceState 
               Height          =   315
               Left            =   1140
               Style           =   2  'Dropdown List
               TabIndex        =   62
               Top             =   270
               Width           =   1695
            End
            Begin VB.TextBox TxtInsuranceNO 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1140
               TabIndex        =   63
               Top             =   600
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   1140
               TabIndex        =   64
               Top             =   960
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Format          =   237174785
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
               Height          =   315
               Left            =   1140
               TabIndex        =   65
               Top             =   1320
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ  »œ«Ì… «· √„Ì‰"
               Height          =   285
               Index           =   38
               Left            =   2850
               TabIndex        =   332
               Top             =   960
               Width           =   1875
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «· √„Ì‰"
               Height          =   285
               Index           =   4
               Left            =   2850
               TabIndex        =   316
               Top             =   240
               Width           =   1845
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·«‘ —«ﬂ"
               Height          =   285
               Index           =   6
               Left            =   2850
               TabIndex        =   315
               Top             =   600
               Width           =   1845
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   8115
         Left            =   16230
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VB.Frame Frame3 
            Caption         =   "„⁄·Ê„«  œŒÊ· «·ÕœÊœ"
            Height          =   2055
            Left            =   6840
            TabIndex        =   215
            Top             =   120
            Width           =   6255
            Begin VB.TextBox txthdodno 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   66
               Top             =   240
               Width           =   4095
            End
            Begin VB.TextBox txthdomnfaz 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   68
               Top             =   960
               Width           =   4095
            End
            Begin VB.TextBox TxtVisaNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               TabIndex        =   69
               Top             =   1320
               Width           =   4095
            End
            Begin Dynamic_Byte.NourHijriCal txthdoddate 
               Height          =   255
               Left            =   2400
               TabIndex        =   67
               Top             =   600
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
            End
            Begin MSDataListLib.DataCombo DcboJobsType1 
               Height          =   315
               Left            =   120
               TabIndex        =   70
               Top             =   1680
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„  œŒÊ· «·ÕœÊœ"
               Height          =   285
               Index           =   29
               Left            =   3960
               TabIndex        =   220
               Top             =   240
               Width           =   1785
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·œŒÊ·  "
               Height          =   285
               Index           =   30
               Left            =   3960
               TabIndex        =   219
               Top             =   600
               Width           =   1785
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "„‰›– «·œŒÊ·  "
               Height          =   285
               Index           =   31
               Left            =   3960
               TabIndex        =   218
               Top             =   960
               Width           =   1785
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Â‰… ›Ì «· √‘Ì—…"
               Height          =   285
               Index           =   49
               Left            =   3960
               TabIndex        =   217
               Top             =   1680
               Width           =   1785
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «· √‘Ì—…"
               Height          =   345
               Index           =   20
               Left            =   3960
               TabIndex        =   216
               Top             =   1320
               Width           =   1785
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   8115
         Left            =   16530
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VB.Frame Frame21 
            Caption         =   "»Ì«‰«  «·«Ã«“« "
            Height          =   7455
            Left            =   -150
            TabIndex        =   343
            Top             =   30
            Width           =   13815
            Begin VB.CheckBox chkStop 
               Alignment       =   1  'Right Justify
               Caption         =   "«Ìﬁ«› «·„Œ’’"
               Height          =   255
               Left            =   11940
               TabIndex        =   437
               Top             =   270
               Width           =   1695
            End
            Begin VSFlex8UCtl.VSFlexGrid Grid20 
               Height          =   5685
               Left            =   120
               TabIndex        =   344
               Top             =   1560
               Width           =   13575
               _cx             =   23945
               _cy             =   10028
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
               Rows            =   1
               Cols            =   16
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProffisionals.frx":27ED
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   1
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
               WordWrap        =   -1  'True
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic10 
         Height          =   8115
         Left            =   17130
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
            Height          =   7365
            Left            =   4320
            TabIndex        =   319
            Top             =   240
            Width           =   9255
            _cx             =   16325
            _cy             =   12991
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
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProffisionals.frx":2A6B
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            WordWrap        =   -1  'True
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic11 
         Height          =   8115
         Left            =   17430
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9315
            Index           =   3
            Left            =   0
            TabIndex        =   223
            TabStop         =   0   'False
            Top             =   0
            Width           =   13815
            _cx             =   24368
            _cy             =   16431
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   7875
               Index           =   4
               Left            =   -240
               TabIndex        =   224
               TabStop         =   0   'False
               Top             =   0
               Width           =   14985
               _cx             =   26432
               _cy             =   13891
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
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ»«⁄… «·’Ê—…"
                  Height          =   255
                  Left            =   10560
                  TabIndex        =   312
                  Top             =   4200
                  Width           =   1575
               End
               Begin VB.ComboBox Combo1 
                  Height          =   315
                  ItemData        =   "FrmProffisionals.frx":2BF0
                  Left            =   9120
                  List            =   "FrmProffisionals.frx":2C06
                  TabIndex        =   232
                  Top             =   3720
                  Width           =   3135
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00FFFFC0&
                  Height          =   3495
                  Left            =   9120
                  TabIndex        =   228
                  Top             =   120
                  Width           =   4575
                  Begin DBPIXLib.DBPix20 DBPix202 
                     Height          =   2055
                     Left            =   600
                     TabIndex        =   229
                     Top             =   240
                     Width           =   2895
                     _Version        =   131072
                     _ExtentX        =   5106
                     _ExtentY        =   3625
                     _StockProps     =   1
                     BackColor       =   16777152
                     _Image          =   "FrmProffisionals.frx":2C7F
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
                     Height          =   495
                     Left            =   600
                     TabIndex        =   230
                     Top             =   2400
                     Width           =   2865
                     _ExtentX        =   5054
                     _ExtentY        =   873
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
                  Begin ImpulseButton.ISButton Cmd1 
                     Height          =   375
                     Left            =   600
                     TabIndex        =   347
                     Top             =   3000
                     Width           =   2865
                     _ExtentX        =   5054
                     _ExtentY        =   661
                     ButtonPositionImage=   1
                     Caption         =   "«·‰„«–Ã Ê «·„—›ﬁ« "
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
               End
               Begin VB.TextBox txtidxxxxxx 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   3
                  Left            =   -3870
                  TabIndex        =   225
                  Top             =   12480
                  Width           =   2145
               End
               Begin ALLButtonS.ALLButton ALLButton1 
                  Height          =   375
                  Left            =   9360
                  TabIndex        =   231
                  Top             =   4200
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   661
                  BTYPE           =   2
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
                  MICON           =   "FrmProffisionals.frx":2C97
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSComCtl2.DTPicker DpCardDriverExpireDate 
                  Height          =   345
                  Left            =   4620
                  TabIndex        =   444
                  Top             =   780
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   245694465
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DpIssuingDriverCardDate 
                  Height          =   345
                  Left            =   4620
                  TabIndex        =   446
                  Top             =   330
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   245694465
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal DpIssuingDriverCardDateH 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   447
                  Top             =   330
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal DpCardDriverExpireDateH 
                  Height          =   315
                  Left            =   3480
                  TabIndex        =   448
                  Top             =   780
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «‰ Â«¡ »ÿ«ﬁ… «·”«∆ﬁ"
                  Height          =   315
                  Index           =   60
                  Left            =   6960
                  RightToLeft     =   -1  'True
                  TabIndex        =   445
                  Top             =   810
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «’œ«— »ÿ«ﬁ… «·”«∆ﬁ"
                  Height          =   375
                  Index           =   58
                  Left            =   6540
                  RightToLeft     =   -1  'True
                  TabIndex        =   443
                  Top             =   375
                  Width           =   2235
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " „«–Ã Â«„… ··ÃÊ«“« "
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   61
                  Left            =   12000
                  TabIndex        =   233
                  Top             =   3840
                  Width           =   1695
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Index           =   54
                  Left            =   13590
                  TabIndex        =   226
                  Top             =   1200
                  Width           =   825
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   43
               Left            =   8400
               TabIndex        =   227
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   8115
         Left            =   17730
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   45
         Width           =   13695
         _cx             =   24156
         _cy             =   14314
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
         GridRows        =   10
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
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   -120
      Width           =   13785
      _cx             =   24315
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰  "
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
         TabIndex        =   165
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
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   -210
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   2460
         TabIndex        =   86
         Top             =   -210
         Visible         =   0   'False
         Width           =   855
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1185
         TabIndex        =   40
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
         ButtonImage     =   "FrmProffisionals.frx":2CB3
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
         TabIndex        =   41
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
         ButtonImage     =   "FrmProffisionals.frx":304D
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
         TabIndex        =   71
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
         ButtonImage     =   "FrmProffisionals.frx":33E7
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
         TabIndex        =   72
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
         ButtonImage     =   "FrmProffisionals.frx":3781
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
         Left            =   3480
         Picture         =   "FrmProffisionals.frx":3B1B
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   60
         Left            =   3000
         TabIndex        =   371
         Top             =   240
         Width           =   7275
      End
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   120
      TabIndex        =   77
      Top             =   9120
      Width           =   855
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   2280
      TabIndex        =   76
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   0
      Left            =   3060
      TabIndex        =   75
      Top             =   9120
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   780
      TabIndex        =   74
      Top             =   9120
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   1
      Left            =   10440
      TabIndex        =   73
      Top             =   10200
      Width           =   675
   End
End
Attribute VB_Name = "FrmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim EmpReport As ClsEmployeeReport
Dim xReport As New CRAXDRT.Report
Dim NO As Double
Public DriverOnly As Integer
Private objScript As Object
Dim case_id As Integer
Dim LogTextA As String
Dim LogTexte As String
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
Dim Account_Code_dynamic3 As String
Dim Account_Code_dynamic4 As String
Dim Account_Code_dynamic5 As String
Public mIndex As Double
Public WorkShop_Job As Integer
Dim FirstPeriodDateInthisYear  As Date
Dim cCompanyInfo As New ClsCompanyInfo
  Dim Dcombos As New ClsDataCombos
  Sub LoadDept()
     Dim Dcombos As New ClsDataCombos
     Dcombos.ClearMyDataCombo DcbDepartment2
     Dcombos.GetNewDwpartMent DcbDepartment2, True, val(DcboEmpDepartments.BoundText)
  End Sub
Function reloadloadFunction()

  
                  
          Dim My_SQL As String
          
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from Nationality  "
    Else
        My_SQL = "  select  id,namee  from Nationality  "
    End If

    fill_combo DCNationality, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from dean  "
    Else
        My_SQL = "  select  id,namee  from dean  "
    End If

    fill_combo dcdean, My_SQL

    Dcombos.GetCodeing Me.DCPreFix, 6
    
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

 
    Dcombos.GetSection Me.DcbSection
    Dcombos.GetEmployees Me.mangerid
    Dcombos.GetEmployees Me.swapedempid
    Dcombos.GetEmployees Me.swapedempid2
  

    If SystemOptions.UserInterface = EnglishInterface Then
    Dcbsex.AddItem "Male"
     Dcbsex.AddItem "Female"
       DcbMatrial.AddItem "Single"
     DcbMatrial.AddItem "Married"
 
    Else
    DcbMatrial.AddItem "√⁄“»"
      DcbMatrial.AddItem "„ “ÊÃ"
     Dcbsex.AddItem "–ﬂ—"
      Dcbsex.AddItem "√‰ÀÏ"
   
    End If

 
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
     Dcombos.GetEmpGrades Me.DcGrade
   Dcombos.GetEmpRelations Me.DcRelation
     Dcombos.GetEmpLocations Me.DCGroupID
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
    Dcombos.GetEmpJobsTypes Me.DcboJobsType1
    Dcombos.GetEmpJobsTypes Me.DcboJobsType2
    Dcombos.GetEmpJobsTypes Me.DcboJobsType3
   Dcombos.GetEmpSpecifications Me.DcboSpecifications
 
    Dcombos.GetBranches Me.DCBranch

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

End Function
 
Function SHOWPIC(PICNAME As String)
    Dim xLogo As CRAXDRT.OLEObject
    StrFileName = App.path & "\" & SystemOptions.ImagesPath & "\" & PICNAME & ".JPG"

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
    Dim X As String
    On Error Resume Next

    'ALLButton1.Enabled = True
    Select Case Combo1.ListIndex + 1

        Case 1

            Dim xApp As New CRAXDRT.Application

            Dim rs As New ADODB.Recordset



If txthdodno.text = "" Then
MsgBox "Õœœ —ﬁ„ «·ÕœÊœ «Ê·«", vbInformation
Exit Sub
End If
            If SystemOptions.UserInterface = EnglishInterface Then

                X = InputBox("Specify No Of Month ")
            Else
                X = InputBox("Õœœ ⁄œœ ‘ÂÊ— «·«ﬁ«„…", " ÕœÌœ „œ… «·«ﬁ«„… »«·‘ÂÊ—")
            End If

            'If x = 0 Then MsgBox "·«»œ „‰  ÕœÌœ ⁄œœ «·‘ÂÊ— ÊÌﬂÊ‰ «—ﬁ«„ ": Exit Sub

            'Form3.Show
            'Form3.case_id = 1
            'Form3.noofmonth = x
            ' Form3.SHOWPICTURE.Caption = Me.Check1.value
            'Form3.TxtEmp_Code = Me.TxtEmp_Code.text

            sql = "SELECT * from emp_all_details WHERE Emp_ID=" & val(Me.XPTxtEmpID.text)
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT1.rpt")
            xReport.Database.SetDataSource rs
      xReport.ParameterFields(1).AddCurrentValue FrmEmployee.DcboJobsType3.text
      
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            
            FrmReport.txtPath = (system_path & "\reports\emp\REPORT1.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            xReport.reporttitle = X
            Sendkeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

        Case 2
            'Dim xApp As New CRAXDRT.Application
            'Dim Rs As New ADODB.Recordset
If Txt_NumEkama.text = "" Then
MsgBox "Õœœ —ﬁ„ «·«ﬁ«„… «Ê·«", vbInformation
Exit Sub
End If



            If SystemOptions.UserInterface = EnglishInterface Then

                X = InputBox("Specify No Of Month ")
            Else
                X = InputBox("Õœœ ⁄œœ ‘ÂÊ— «·«ﬁ«„…", " ÕœÌœ „œ… «·«ﬁ«„… »«·‘ÂÊ—")
            End If

            'If x = 0 Then MsgBox "·«»œ „‰  ÕœÌœ ⁄œœ «·‘ÂÊ— ÊÌﬂÊ‰ «—ﬁ«„ ": Exit Sub

            sql = "SELECT * from emp_all_details WHERE  Emp_ID=" & Me.XPTxtEmpID.text
     
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT2.rpt")
            xReport.Database.SetDataSource rs
       xReport.ParameterFields(1).AddCurrentValue FrmEmployee.DcboJobsType3.text
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.txtPath = (system_path & "\reports\emp\REPORT2.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            xReport.reporttitle = X
            Sendkeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

            'Form3.Show
            'Form3.case_id = 2
            'Form3.noofmonth = x
            ' Form3.SHOWPICTURE.Caption = Me.Check1.value
            'Form3.TxtEmp_Code = Me.TxtEmp_Code.text

        Case 3

            sql = "SELECT * from emp_all_details WHERE  Emp_ID=" & Me.XPTxtEmpID.text
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT3.rpt")
            xReport.Database.SetDataSource rs
 xReport.ParameterFields(1).AddCurrentValue FrmEmployee.DcboJobsType3.text
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.txtPath = (system_path & "\reports\emp\REPORT3.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            '    xReport.ReportTitle = X
            Sendkeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

            'Form3.Show
            'Form3.case_id = 3
            ' Form3.SHOWPICTURE.Caption = Me.Check1.value
            'Form3.TxtEmp_Code = Me.TxtEmp_Code.text
        Case 4
            sql = "SELECT * from emp_all_details WHERE   Emp_ID=" & Me.XPTxtEmpID.text
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT4.rpt")
            xReport.Database.SetDataSource rs
 xReport.ParameterFields(1).AddCurrentValue FrmEmployee.DcboJobsType3.text
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.txtPath = (system_path & "\reports\emp\REPORT4.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            ' xReport.ReportTitle = X
            Sendkeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

        Case 5
            outform.show
            outform.Check7.value = Me.Check1.value
            outform.txtemp_code = val(Me.XPTxtEmpID.text)

        Case 6

            If SystemOptions.UserInterface = EnglishInterface Then

                X = InputBox("Specify Date ")
            Else
                X = InputBox("Õœœ  «—ÌŒ «·Â—Ê» ", "10/02/1432")
            End If
    
            If Len(X) <> 10 Then
                If SystemOptions.UserInterface = EnglishInterface Then
        
                    MsgBox "wrong date  ex 11/02/1432 "
                Else
                    MsgBox "Õœœ  «—ÌŒ «·Â—Ê»  «·’ÕÌÕ " & "10/02/1432"
         
                End If

                Exit Sub
            End If

            sql = "SELECT * from emp_all_details WHERE  Emp_ID=" & Me.XPTxtEmpID.text
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT8.rpt")
            xReport.Database.SetDataSource rs
 
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
            FrmReport.txtPath = (system_path & "\reports\emp\REPORT8.rpt")
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            '    xReport.ReportTitle = X
            xReport.ParameterFields(1).AddCurrentValue mId(X, 1, 2)
            xReport.ParameterFields(2).AddCurrentValue mId(X, 4, 2)
            xReport.ParameterFields(3).AddCurrentValue mId(X, 9, 2)
  
            Sendkeys "{RIGHT}"

            If Check1.value = 1 Then
                SHOWPIC (FrmEmployee.XPTxtEmpID.text)
            End If

            'ALLButton1.Enabled = False
    End Select

End Sub


Private Sub RemoveGridRow()
      If Grid.rows = 1 Then Exit Sub
    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub



Private Sub RemoveGridRow2()
      If grid02.rows = 1 Then Exit Sub
    With Me.grid02

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow1()
      If Grid.rows = 1 Then Exit Sub
    With Me.grid01

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With
     IntCounter = 0
     
         With Me.VSFlexGrid3

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With
    

      IntCounter = 0
     
         With Me.grid01

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With
 IntCounter = 0


         With Me.grid02

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With
    
    
    
     IntCounter = 0


         With Me.Grid20

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("fromdate")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
       Coloring
    End With
    


End Sub
Private Sub Coloring()
    Dim i As Integer
    Dim IntCounter As Integer

    With Grid20

        For i = .FixedRows To .rows - 1
        
            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 0, i, 8) = &HFFFFC0
            Else
                .cell(flexcpBackColor, i, 0, i, 8) = vbWhite
            End If

        Next i

    End With

    'line_no1 = IntCounter

End Sub


Private Sub ALLButton2_Click(Index As Integer)
If OptSalaryType(0).value = False Then Exit Sub
    Select Case Index

        Case 0
            mO2AHELAT.show

        Case 1
         '   TABE3.show
Frame10.Visible = True
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
                frmEmpSalaryComponent.Emp_Code = txtemp_code.text
                frmEmpSalaryComponent.emp_Name(0) = Text1.text
                frmEmpSalaryComponent.emp_Name(1) = Text2.text
                frmEmpSalaryComponent.emp_Name(2) = Text3.text
                frmEmpSalaryComponent.emp_Name(3) = Text4.text
                frmEmpSalaryComponent.DEPARTEMENT.text = DcboEmpDepartments.text
                frmEmpSalaryComponent.job.text = DcboJobsType.text
                frmEmpSalaryComponent.Issue_date.value = DTPicker1.value
                frmEmpSalaryComponent.Basic_salary.text = val(TxtSalary.text)
                frmEmpSalaryComponent.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                frmEmpSalaryComponent.VSFlexGrid1.rows = 1
                frmEmpSalaryComponent.Cmd_Click (1)

            End If
 
            If val(XPTxtEmpID.text) <> 0 Then
                frmEmpSalaryComponent.show
                frmEmpSalaryComponent.Contract_ID = Me.Contract_ID
                frmEmpSalaryComponent.Retrive val(XPTxtEmpID.text)
            
                If Me.TxtModFlg.text = "E" Then
                    frmEmpSalaryComponent.Cmd_Click (1)
                    '   frmEmpSalaryComponent.Basic_salary.text = val(TxtSalary.text)
                End If
            
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
If Me.TxtModFlg.text = "N" Then
                'If SystemOptions.UserInterface = ArabicInterface Then
                'MsgBox "«Õ›Ÿ »Ì«‰«  «·„ÊŸ› «·«”«”Ì… «Ê·« "
                'Else
                'MsgBox "Save Employee Basic Information Firstly!"
                frmEmpSalaryComponentIncres.show
                frmEmpSalaryComponentIncres.Contract_ID = Me.Contract_ID
                frmEmpSalaryComponentIncres.Emp_id = val(XPTxtEmpID.text)
                frmEmpSalaryComponentIncres.Emp_Code = txtemp_code.text
                frmEmpSalaryComponentIncres.emp_Name(0) = Text1.text
                frmEmpSalaryComponentIncres.emp_Name(1) = Text2.text
                frmEmpSalaryComponentIncres.emp_Name(2) = Text3.text
                frmEmpSalaryComponentIncres.emp_Name(3) = Text4.text
                frmEmpSalaryComponentIncres.DEPARTEMENT.text = DcboEmpDepartments.text
                frmEmpSalaryComponentIncres.job.text = DcboJobsType.text
                frmEmpSalaryComponentIncres.Issue_date.value = DTPicker1.value
                frmEmpSalaryComponentIncres.Basic_salary.text = val(TxtSalary.text)
                frmEmpSalaryComponentIncres.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                frmEmpSalaryComponentIncres.VSFlexGrid1.rows = 1
                frmEmpSalaryComponentIncres.Cmd_Click (1)

            End If
 
            If val(XPTxtEmpID.text) <> 0 Then
                frmEmpSalaryComponentIncres.show
                frmEmpSalaryComponentIncres.Contract_ID = Me.Contract_ID
                frmEmpSalaryComponentIncres.Retrive val(XPTxtEmpID.text)
            
                If Me.TxtModFlg.text = "E" Then
                    frmEmpSalaryComponentIncres.Cmd_Click (1)
                    '   frmEmpSalaryComponent.Basic_salary.text = val(TxtSalary.text)
                End If
            
            Else

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "ﬂÊœ „ÊŸ› Œÿ√"
                Else
                    MsgBox "Invalid Employee Code"
                End If

            End If
        Case 6
            frmEmpContract.show
             frmEmpContract.Retrive , val(Me.XPTxtEmpID.text), True

            ' Cmd1_Click
        Case 7
            'OHDA.Show
    End Select

End Sub

Private Sub CboInsuranceState_Change()

    If Me.CboInsuranceState.ListIndex = 0 Then
       ' Me.TxtInsurValue.text = ""
       ' Me.TxtInsurValue.Enabled = False
    Else
      '  Me.TxtInsurValue.Enabled = True
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

Function addrow2()
 Dim i As Integer
 
 
      If grid02.rows = 1 Then grid02.rows = 2
         With grid02
  i = .rows
 
               .TextMatrix(i - 1, .ColIndex("name")) = (TxtWorkName.text)
      
                    
                
                .TextMatrix(i - 1, .ColIndex("qualicationEntity")) = (TxtWorkEntity.text)
                
                .TextMatrix(i - 1, .ColIndex("workfrom")) = (workfrom.value)
                .TextMatrix(i - 1, .ColIndex("workfromH")) = (workfromH.value)
                .TextMatrix(i - 1, .ColIndex("workto")) = (workto.value)
                .TextMatrix(i - 1, .ColIndex("worktoH")) = (worktoH.value)
                .TextMatrix(i - 1, .ColIndex("des")) = (Text27.text)
                  
                  '.TextMatrix(i - 1, .ColIndex("des")) = (Txtdes1.text)
                  
                
                  .rows = .rows + 1
                  TxtWorkName.text = ""
                  TxtWorkEntity.text = ""
             
                 
      '       .AutoSize 0, .Cols - 1, False
   
    End With
 
     
    ReLineGrid

End Function


Function addrow1()
 Dim i As Integer
 
 
      If grid01.rows = 1 Then grid01.rows = 2
         With grid01
  i = .rows
 
               .TextMatrix(i - 1, .ColIndex("name")) = (TxtName1.text)
      
                    
                .TextMatrix(i - 1, .ColIndex("yearofqualication")) = (txtyearofqualication.text)
                .TextMatrix(i - 1, .ColIndex("qualicationEntity")) = (TxtqualicationEntity.text)
                .TextMatrix(i - 1, .ColIndex("grade")) = txtgrade.text
                 
                  .TextMatrix(i - 1, .ColIndex("des")) = (TxtDes2.text)
                  
                
                  .rows = .rows + 1
                  TxtName1.text = ""
                  txtyearofqualication.text = ""
                  TxtqualicationEntity.text = ""
                  txtgrade.text = ""
                  TxtDes2.text = ""
                 
      '       .AutoSize 0, .Cols - 1, False
   
    End With
 
     
    ReLineGrid

End Function

Function addrow()
 Dim i As Integer
 
 

      If Grid.rows = 1 Then Grid.rows = 2
         With Grid
  i = .rows
 
               .TextMatrix(i - 1, .ColIndex("name")) = (TxtName.text)
                .TextMatrix(i - 1, .ColIndex("relationtype")) = val(Me.DcRelation.BoundText)
                    .TextMatrix(i - 1, .ColIndex("relationtypename")) = (Me.DcRelation.text)
                    
                .TextMatrix(i - 1, .ColIndex("passportno")) = (Txtpassportno.text)
                .TextMatrix(i - 1, .ColIndex("iqamano")) = (TxtIqamaNo.text)
                .TextMatrix(i - 1, .ColIndex("haveinsurance")) = IIf(chkhaveinsurance.value = vbChecked, 1, 0)
                .TextMatrix(i - 1, .ColIndex("insuranceno")) = (insuranceno1.text)
                  .TextMatrix(i - 1, .ColIndex("des")) = (Txtdes1.text)
                  
                
                  .rows = .rows + 1
                  TxtName.text = ""
                  Txtpassportno.text = ""
                  TxtIqamaNo.text = ""
                  insuranceno1.text = ""
                  Txtdes1.text = ""
                  chkhaveinsurance.value = vbUnchecked
      '       .AutoSize 0, .Cols - 1, False
   
    End With
 
     
    ReLineGrid

End Function

Private Sub Cmd_Click(Index As Integer)
      getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
       '     Me.Dtp1 = FirstPeriodDateInthisYear
       '     Me.Dtp2 = FirstPeriodDateInthisYear
       '     Me.Dtp4 = FirstPeriodDateInthisYear
     

 '  On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            XPLbl(60).Caption = ""
      '      Text1.SetFocus
            Txt_DateExpEkamaH.value = ToHijriDate(Date)
            Txt_DateEndekamah.value = ToHijriDate(Date)
            Txt_DateExpLincH.value = ToHijriDate(Date)
            VSFlexGrid3.rows = 2
            '**************************************************
            txtDriverLicenseStartdH.value = ToHijriDate(Date)
             txtDriverLicenseendH.value = ToHijriDate(Date)
            ' For i = 0 To 1
            ' DTPicker4(i).value = Date
            '  NourHijriCal2(i).value = ToHijriDate(DTPicker4(i).value)
            ' Next i
            '***************************************************
            Txt_DateEndLincH.value = ToHijriDate(Date)
            txthdoddate.value = ToHijriDate(Dcaseate)
            Txt_DateExppoketH.value = ToHijriDate(Date)
            Txt_DateEndpoketH.value = ToHijriDate(Date)
        
      
            OptType(2).value = True
            TxtSalary.text = 0
    
            OptType1(2).value = True
            OptType2(2).value = True
             OptType5(2).value = True
OptType4(2).value = True
 'OptType3(2).value = True
 ReloadCompo
Me.DCBranch.BoundText = Current_branch
OptSalaryType(0).value = True
DBPix201.ImageClear
        DBPix202.ImageClear
txtKafelID.text = cCompanyInfo.ArabComment
'Me.DcbKafelName.BoundText = cCompanyInfo.ArabCompanyName
TxtKafeltEL.text = cCompanyInfo.CompanyTel
txtkafeladd.text = cCompanyInfo.CompanyAddress
lbl(46).Caption = ""
C1Tab1.CurrTab = 0
 cboPayType.ListIndex = 0
 
 cmbInsuranceRenew.ListIndex = 1
    cmbToM.ListIndex = 1
'txtopening_balance_voucher_id.Text =0
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
    rs.Close
     '   StrSQL = "select * from  TblEmployee order by fullcode"
 
  'If WorkShop_Job = 0 Then
     '   StrSQL = "select * from  TblEmployee order by fullcode"
      '  Else
      '  StrSQL = "select * from  TblEmployee WHERE WorkShop_Job=" & WorkShop_Job & " order by fullcode"
      '  End If
     
   
    StrSQL = "select * from  TblEmployee where 1=1"
       
       If DriverOnly = 1 Then
StrSQL = "    select * from TblEmployee"
StrSQL = StrSQL & "  where   dbo.TblEmployee.JobTypeID in("
StrSQL = StrSQL & "   select dbo.TblEmpJobsTypes.JobTypeID"
StrSQL = StrSQL & "   From TblEmpJobsTypes"
StrSQL = StrSQL & "   where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1))"

    Else
    StrSQL = "select * from  TblEmployee where 1=1"
    
    End If
        
        
       If WorkShop_Job = 0 Then
       Else
        StrSQL = StrSQL & " and  WorkShop_Job=" & WorkShop_Job  '& " order by fullcode"
        End If
        
   
   
   
        If SystemOptions.usertype <> UserAdminAll Then
          '  StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
          StrSQL = StrSQL & "  AND  (BranchId=0 or BranchId is null or         BranchId in(" & Current_branchSql & "))"
          
        End If
        
        StrSQL = StrSQL & " order by fullcode "
        

   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    

    Me.Retrive val(Me.XPTxtEmpID)
            
            
     ' For i = 0 To 1
     '        DTPicker4(i).value = Date
     '         NourHijriCal2(i).value = ToHijriDate(DTPicker4(i).value)
     '        Next i
             
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
           ' Me.Dtp = FirstPeriodDateInthisYear
           ' Me.Dtp1 = FirstPeriodDateInthisYear
           ' Me.Dtp2 = FirstPeriodDateInthisYear
           VSFlexGrid3.rows = VSFlexGrid3.rows + 1
Grid.rows = Grid.rows + 1
grid01.rows = grid01.rows + 1
grid02.rows = grid02.rows + 1
      getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
       '     Me.Dtp1 = FirstPeriodDateInthisYear
       '     Me.Dtp2 = FirstPeriodDateInthisYear
       '     Me.Dtp4 = FirstPeriodDateInthisYear
     
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
      Dim EmpName As String
      Dim fullcode As String
      Dim msgstr As String
If Trim(DCNationality.text) = "”⁄ÊœÌ" Or Trim(DCNationality.text) = "”⁄ÊœÏ" Or Trim(DCNationality.text) = "saudi" Or UCase$(DCNationality.text) = "SAUDI" Then
If Me.ChekID(1, EmpName, fullcode) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
msgstr = "·«Ì„ﬂ‰  ﬂ—«—  " & CHR(13)
msgstr = msgstr & "—ﬁ„ «·ÂÊÌ… „”Ã·  ··„ÊŸ› " & " " & EmpName & CHR(13)
msgstr = msgstr & "ﬂÊœ «·„ÊŸ› " & " " & fullcode & CHR(13)
Else
msgstr = "Can not Repeat the ID number" & CHR(13)
msgstr = msgstr & "This number is registered with the Employee" & " " & EmpName & CHR(13)
msgstr = msgstr & "Code" & " " & fullcode
End If
MsgBox msgstr

Exit Sub
End If
Else
If Me.ChekID(0, EmpName, fullcode) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
msgstr = "·«Ì„ﬂ‰  ﬂ—«—  " & CHR(13)
msgstr = msgstr & "—ﬁ„ «·«ﬁ«„… „”Ã· ··„ÊŸ› " & " " & EmpName & CHR(13)
msgstr = msgstr & "ﬂÊœ «·„ÊŸ› " & " " & fullcode & CHR(13)
Else
msgstr = "Can not Repeat the ID number" & CHR(13)
msgstr = msgstr & "This number is registered with the Employee" & " " & EmpName & CHR(13)
msgstr = msgstr & "Code" & " " & fullcode
End If
MsgBox msgstr
Exit Sub

End If
End If
If Me.txtBank(1).text <> "" Then
If Len(Me.txtBank(1).text) <> 24 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» «‰ ÌﬂÊ‰ ⁄œœ «·Œ«‰«  ›Ì «·«Ì»«‰ 24 Œ«‰…"
Else
MsgBox "The No. of Iban boxes should be 24 digits"
End If
Exit Sub
End If
End If
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_ProfData

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
  Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show vbModal
 
        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            printingReport
            
            Case 8
            addrow1
            Case 9
            RemoveGridRow1
            
            Case 10
            addrow3
            
            'nnnn Case 11
           'nnnn RemoveGridRow3
            
            Case 20
            addrow
            Case 21
            RemoveGridRow
            
            
            
            Case 14
              
            addrow2
            Case 15
            RemoveGridRow2
            
            
    End Select

    Exit Sub
ErrTrap:





End Sub
'nnnn Private Sub RemoveGridRow3()

 'nnnn   With Me.Grid20
'nnnn
 'nnnn       If .Row <= 0 Then Exit Sub
 'nnnn       .RemoveItem .Row
  'nnnn  End With

  'nnnn  ReLineGrid
'nnnn End Sub
Sub addrow3()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String

    If Me.TxtModFlg.text = "R" Then
 
 Exit Sub
 
    End If
  

    
        Me.Grid20.rows = Me.Grid20.rows + 1
        LngRow = Me.Grid20.rows - 1
  
 
 

 
    On Error Resume Next
 
    With Me.Grid20
 
 
 
    ' .TextMatrix(LngRow, .ColIndex("fromdate")) = DTPicker4(0).value
        ' .TextMatrix(LngRow, .ColIndex("todate")) = DTPicker4(1).value
        '  .TextMatrix(LngRow, .ColIndex("fromdateh")) = NourHijriCal2(0).value
        ' .TextMatrix(LngRow, .ColIndex("todateh")) = NourHijriCal2(1).value
             Dim astrSplitItems() As String
            Dim Result As String
 
    Dim diff_year As Integer
    Dim Txtyear As String
    Dim TxtMonth As String
    Dim TxtDay As String
    
    'result = ExactAge(DTPicker4(0).value, DTPicker4(1).value)

    If Result = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " «—ÌŒ »œ«Ì… «·«Ã«“… ÂÊ ‰›”  «—ÌŒ ‰Â«Ì… «Ê «ﬂ»— „‰Â ﬁ„ » €ÌÌ— ﬁÌ„ «· «—ÌŒ ", vbCritical
        Else
            MsgBox "Date of start of work is the same date as the end of the work ", vbCritical
        End If
        Exit Sub
    End If

    astrSplitItems = Split(Result, "-")
    Txtyear = astrSplitItems(0)
    TxtMonth = astrSplitItems(1)
    TxtDay = astrSplitItems(2)
    
 
          .TextMatrix(LngRow, .ColIndex("day")) = TxtDay
           .TextMatrix(LngRow, .ColIndex("month")) = TxtMonth
            .TextMatrix(LngRow, .ColIndex("year")) = Txtyear
            
          
      
       ' .TextMatrix(LngRow, .ColIndex("remarks")) = (Me.txtRemarks.text)
       
          .AutoSize 0, .Cols - 1, False
    End With
    ReLineGrid
 
End Sub


Private Sub Cmd1_Click()
    On Error Resume Next
       If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
    If DCPreFix.text & txtid.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „ÊŸ› «Ê·«": Exit Sub

'    imaged.show
'    imaged.Label9.Caption = "„—›ﬁ«  «·„ÊŸ› —ﬁ„"
'    imaged.Caption = "„—›ﬁ«  «·„ÊŸ›  "
'    imaged.txtopeation_type = "„—›ﬁ«  „ÊŸ›"
''    imaged.SUBJECT_NO = DCPreFix.Text & TxtId.Text
'    imaged.Label6.Caption = "ﬂÊœ «·„ÊŸ›"
'    imaged.Adodc1.CommandType = adCmdText
'    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  „ÊŸ›' and subject_no='" & DCPreFix.Text & TxtId.Text & "'"
'    imaged.Adodc1.Refresh
'
'    If imaged.Adodc1.Recordset.RecordCount > 0 Then

'        imaged.DBPix201.Visible = True
'    Else
'        imaged.DBPix201.Visible = False
'    End If

ShowAttachments DCPreFix.text & txtid.text, "„—›ﬁ«  „ÊŸ›"
End Sub

Private Sub CmdExit_Click()
    Frame1.Visible = False
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

'Private Sub CmdEstkala_Click()
'    Frame1.Visible = True
'End Sub

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

 

Private Sub DcboEmpDepartments_Change()
LoadDept
End Sub

Private Sub DcboEmpDepartments_Click(Area As Integer)
DcboEmpDepartments_Change
End Sub

Private Sub DcboEmpDepartments_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DcboJobsType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DcboJobsType1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DcboJobsType2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DcboJobsType3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DcboSpecifications_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
reloadloadFunction
End If

End Sub

Private Sub DcbSection_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 7
    End If

End Sub

Private Sub dcdean_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DCGroupID_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
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

 

Private Sub dcjopstatus_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DCNationality_Change()
If Trim(DCNationality.text) = "”⁄ÊœÌ" Or LCase(Trim(DCNationality.text)) = "saudi" Then
Fra(3).Visible = False
Fra(2).Visible = False
Fra(5).Visible = False
Fra(6).Visible = False
Fra(4).Visible = True
Else
Fra(3).Visible = True
Fra(2).Visible = True
Fra(5).Visible = True
Fra(6).Visible = True
Fra(4).Visible = False

End If
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
      DCPreFix.text = GETNationality(val(DCNationality.BoundText))
 End If
End Sub
Function ChekID(Optional ID As Integer = 0, Optional ByRef Name As String, Optional ByRef code As String) As Boolean
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim sql As String
ChekID = False
If ID = 0 Then
If Me.Txt_NumEkama.text <> "" Then
sql = "select Fullcode,Emp_Name,Emp_Namee NumEkama from  TblEmployee where Emp_ID<>" & val(XPTxtEmpID.text) & " and NumEkama='" & Txt_NumEkama.text & " ' "
End If
Else
If Me.Tet_NumPoket.text <> "" Then
sql = "select Fullcode, Emp_Name,NumPoket,Emp_Namee from  TblEmployee where Emp_ID<>" & val(XPTxtEmpID.text) & " and NumPoket='" & Tet_NumPoket.text & " ' "
End If
End If
If sql <> "" Then
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
Name = IIf(IsNull(Rs5("Emp_Name").value), "", Rs5("Emp_Name").value)
Else
Name = IIf(IsNull(Rs5("Emp_Namee").value), "", Rs5("Emp_Namee").value)
End If
code = IIf(IsNull(Rs5("Fullcode").value), "", Rs5("Fullcode").value)
ChekID = True
Else
ChekID = False
End If
End If
End Function
Private Sub DcNationality_Click(Area As Integer)
DCNationality_Change

End Sub

Private Sub dcnationality_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub DCPreFix_Click(Area As Integer)
txtid.text = ""
End Sub

Private Sub DOBH_LostFocus()
      DBDOB.value = ToGregorianDate(DOBH.value)

End Sub

Private Sub DpCardDriverExpireDate_Change()
If Me.TxtModFlg.text <> "R" Then
         DpCardDriverExpireDateH.value = ToHijriDate(DpCardDriverExpireDate.value)
End If
End Sub

Private Sub DpCardDriverExpireDateH_LostFocus()
   DpCardDriverExpireDate.value = ToGregorianDate(DpCardDriverExpireDateH.value)
End Sub

Private Sub DpIssuingDriverCardDate_Change()
If Me.TxtModFlg.text <> "R" Then
         DpIssuingDriverCardDateH.value = ToHijriDate(DpIssuingDriverCardDate.value)
End If
End Sub

Private Sub DpIssuingDriverCardDateH_LostFocus()
   DpIssuingDriverCardDate.value = ToGregorianDate(DpIssuingDriverCardDateH.value)
End Sub

Private Sub DTPicker1_Change()
If Me.TxtModFlg.text <> "R" Then
     
         IssueDateH.value = ToHijriDate(DTPicker1.value)
       
End If

End Sub

Private Sub DBDOB_Change()
If Me.TxtModFlg.text <> "R" Then
         DOBH.value = ToHijriDate(DBDOB.value)
          lbl(47).Caption = DateDiff("yyyy", Me.DBDOB.value, Date) + 1
       
End If

End Sub


Private Sub DTPicker3_Change()
If Me.TxtModFlg.text <> "R" Then
     
         NourHijriCal1.value = ToHijriDate(DTPicker3.value)
       
End If

End Sub

Private Sub Form_Activate()
    ShowDynamicHelp Me.HelpContextID
End Sub

Private Sub Form_Load()
    system_path = App.path
    C1Tab1.CurrTab = 0
    


    
    If DriverOnly = 1 Then
 For i = 0 To 11
C1Tab1.TabVisible(i) = False
 Next i
 
 C1Tab1.TabVisible(0) = True
 C1Tab1.TabVisible(1) = True
 C1Tab1.TabVisible(9) = True
 C1Tab1.TabVisible(11) = True
 
End If

   LogTextA = "»Ì«‰«  «·„ÊŸ›Ì‰"
   LogTexte = "Employee"
    If SystemOptions.CantWorkwithComponenetinEmpScr = True Then
   
    ALLButton2(4).Enabled = False
    ALLButton2(6).Enabled = False
    ALLButton2(5).Enabled = False
    lbl(46).Visible = False
    Else
    ALLButton2(4).Enabled = True
    ALLButton2(6).Enabled = True
    ALLButton2(5).Enabled = True
    lbl(46).Visible = True
    End If



If CheckAutoCoding(6) = True Then
            'txtid.Enabled = False
            DCPreFix.Enabled = True
Else
    '  txtid.Enabled = True
       DCPreFix.Enabled = False
 End If
 
        ReloadCompo
          Dim My_SQL As String
          
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from Nationality  "
    Else
        My_SQL = "  select  id,namee  from Nationality  "
    End If

    fill_combo DCNationality, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from dean  "
    Else
        My_SQL = "  select  id,namee  from dean  "
    End If

    fill_combo dcdean, My_SQL

    Dcombos.GetCodeing Me.DCPreFix, 6
    
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
    Dcombos.GetNewDwpartMent DcbDepartment2
    Dcombos.Getemp_Contract_type DcbContractType
    Dcombos.GetSection Me.DcbSection
    Dcombos.GetEmployees Me.mangerid
    Dcombos.GetEmployees Me.swapedempid
    Dcombos.GetEmployees Me.swapedempid2
    Dcombos.GetSection Me.DCRegionID
    With cboPayType
.Clear
.AddItem "‰ﬁœ«"
.AddItem "‘Ìﬂ"
.AddItem "’—«›"
.AddItem " ÕÊÌ· »‰ﬂÌ"
.AddItem "«Œ—Ì"
End With
    
'      On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
    Dcbsex.AddItem "Male"
     Dcbsex.AddItem "Female"
        DcbMatrial.AddItem "Single"
     DcbMatrial.AddItem "Married"
      DcbMatrial.AddItem "Divorced"
       DcbMatrial.AddItem "Widowed"
       
     
        SetInterface Me
        ChangeLang
        
        Msg = "Sales Commission"
        Msg = Msg & CHR(13) & "enter the values of the sales commission"
        Msg = Msg & " and it will calculted From the Sales Total Vaules and"
        Msg = Msg & " the Sales net profit values"
    
    Else
     DcbMatrial.AddItem "√⁄“»"
      DcbMatrial.AddItem "„ “ÊÃ"
      DcbMatrial.AddItem "„ÿ·ﬁ/„ÿ·›…"
      DcbMatrial.AddItem "«—„·/√—„·…"
      
     Dcbsex.AddItem "–ﬂ—"
      Dcbsex.AddItem "√‰ÀÏ"
        Msg = "ﬁÌ„… «·⁄„Ê·… ⁄·Ï «·„»Ì⁄« "
        Msg = Msg & CHR(13) & " ≈–« ﬂ«‰ «·„ÊŸ› ÌÕ’· ⁄·Ï ⁄„Ê·… ⁄·Ï «·„»Ì⁄« "
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
     Dcombos.GetEmpGrades Me.DcGrade
     
     
    Dcombos.GetEmpRelations Me.DcRelation
    
    
     Dcombos.GetEmpLocations Me.DCGroupID
     
     
     
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
    Dcombos.GetEmpJobsTypes Me.DcboJobsType1
    Dcombos.GetEmpJobsTypes Me.DcboJobsType2
    Dcombos.GetEmpJobsTypes Me.DcboJobsType3
    

    Dcombos.GetEmpSpecifications Me.DcboSpecifications
'    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches Me.DCBranch

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

        With Me.cmbInsuranceRenew
            .Clear
            .AddItem " „ «· ÃœÌœ"
            .AddItem "·„ ÌÃœœ"
        End With


        With Me.cmbToM
            .Clear
            .AddItem " „ «· ”œÌœ"
            .AddItem "·„ Ì”œœ"
        End With
    Else

        With Me.CboInsuranceState
            .Clear
            .AddItem "Not have"
            .AddItem "  Have Insurance"
        End With


        With Me.cmbInsuranceRenew
            .Clear
            .AddItem "Updated"
            .AddItem "Not renewed"
        End With

        With Me.cmbToM
            .Clear
            .AddItem "Payment made"
            .AddItem "Not paid"
        End With
    End If
    cmbInsuranceRenew.ListIndex = 1
    cmbToM.ListIndex = 1
    
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    'rs.Open "[TblEmployee]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Dim StrSQL As String
    If DriverOnly = 1 Then
            StrSQL = "    select * from TblEmployee"
            StrSQL = StrSQL & "  where   dbo.TblEmployee.JobTypeID in("
            StrSQL = StrSQL & "   select dbo.TblEmpJobsTypes.JobTypeID"
            StrSQL = StrSQL & "   From TblEmpJobsTypes"
            StrSQL = StrSQL & "   where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1))"

    Else
            StrSQL = "select * from  TblEmployee where 1=1"
    
    End If
    
       If WorkShop_Job = 0 Then
        ' StrSQL = "select * from  TblEmployee" ' order by fullcode"
        Else
        StrSQL = StrSQL & " and  WorkShop_Job=" & WorkShop_Job  '& " order by fullcode"
        End If
        
        If SystemOptions.usertype <> UserAdminAll Then
            'StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
            
        End If
            StrSQL = StrSQL & "  AND  (BranchId=0 or BranchId is null or         BranchId in(" & Current_branchSql & "))"
            
        StrSQL = StrSQL & " Order By LEN(FullCode),FullCode "
   ' StrSQL = "select * from  TblEmployee order by CAST(Emp_Code AS int)"
    'CAST(Emp_Code AS int)

    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'WorkShop_Job = 0
    Me.TxtModFlg.text = "R"
XPBtnMove_Click 2
      If SystemOptions.SpecialVersion = True And mIndex = 0 Then
     C1Tab1.TabVisible(9) = False
End If
mIndex = 0
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

 



Sub ReloadCompo()
Dim sql As String
sql = "SELECT DISTINCT KafelName, KafelName AS KafelNames"
sql = sql & " From dbo.TblEmployee"
sql = sql & " WHERE     (NOT (KafelName IS NULL)) "
fill_combo DcbKafelName, sql
sql = "SELECT DISTINCT BankCode, BankCode AS BankCodeName"
sql = sql & " From dbo.TblEmployee"
sql = sql & " WHERE     (NOT (BankCode IS NULL)) "
fill_combo Me.DcbBanck(0), sql
sql = "SELECT DISTINCT BanckName, BanckName AS BanckNameName"
sql = sql & " From dbo.TblEmployee"
sql = sql & " WHERE     (NOT (BanckName IS NULL)) "
fill_combo Me.DcbBanck(1), sql

sql = "SELECT DISTINCT pasplace, pasplace AS pasplaceName"
sql = sql & " From dbo.TblEmployee"
sql = sql & " WHERE     (NOT (pasplace IS NULL)) "
fill_combo Me.DcbPasplace, sql
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub ISButton1_Click()

    If XPTxtEmpID.text = "" Then Exit Sub
    X = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

    If X = vbYes Then
        DBPix201.ImageLoad

        DoEvents
        MsgBox " „  Õ„Ì· «·’Ê—…"
    Else

        If X = vbNo Then
            DBPix201.TWAINAcquire
            MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"

            DoEvents
        Else

            Exit Sub
        End If
    End If

    DBPix201.ImageSaveFile (system_path & "\" & SystemOptions.ImagesPath & "\" & XPTxtEmpID.text & ".JPG")
End Sub

Private Sub ISButton2_Click()
      If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
    If XPTxtEmpID.text = "" Then Exit Sub
    X = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

    If X = vbYes Then
        DBPix202.ImageLoad

        DoEvents
        MsgBox " „  Õ„Ì· «·’Ê—…"
    Else

        If X = vbNo Then
            DBPix202.TWAINAcquire
            MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"

            DoEvents
        Else

            Exit Sub
        End If
    End If

    DBPix202.ImageSaveFile (system_path & "\" & SystemOptions.ImagesPath & "\sign" & XPTxtEmpID.text & ".JPG")
End Sub

Private Sub lastHolidaydate_Change()
If Me.TxtModFlg.text <> "R" Then
     
         lastHolidaydateH.value = ToHijriDate(lastHolidaydate.value)
       
End If

End Sub

Private Sub lastHolidaydateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
 lastHolidaydate.value = ToGregorianDate(lastHolidaydateH.value)
End If
End Sub

Private Sub mangerid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub NourHijriCal1_GotFocus()
  '  MsgBox NourHijriCal1.value
End Sub

 

Private Sub Label6_Click()

Fra(15).Visible = False
End Sub

Private Sub LastDate_Change()
If Me.TxtModFlg.text <> "R" Then
     
         LastDateH.value = ToHijriDate(LastDate.value)
       
End If

End Sub



Private Sub NourHijriCal1_LostFocus()
If Me.TxtModFlg.text <> "R" Then
 DTPicker3.value = ToGregorianDate(NourHijriCal1.value)
End If
End Sub

Private Sub OptSalaryType_Click(Index As Integer)
TxtPercentage.text = ""
TxtBYHour.text = ""
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

 Private Sub OptType4_Click(Index As Integer)
    Me.TxtOpenBalance4.Enabled = Not OptType4(2).value
    Me.TxtOpenBalance4.text = IIf(OptType4(2).value = True, 0, Me.TxtOpenBalance4.text)

End Sub
 Private Sub OptType5_Click(Index As Integer)
    Me.TxtOpenBalance5.Enabled = Not OptType5(2).value
    Me.TxtOpenBalance5.text = IIf(OptType5(2).value = True, 0, Me.TxtOpenBalance5.text)

End Sub
Private Sub swapedempid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub swapedempid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
reloadloadFunction
End If
End Sub

Private Sub Text1_Change()
If Me.TxtModFlg.text <> "R" Then
Text5.text = Translate(0, Text1.text)
End If

End Sub

Private Sub Text1_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'    If Button = vbRightButton Then
'FrmADDToDictionary.show
'FrmADDToDictionary.Txtaname.text = Text1.text
'    End If
    
End Sub

Private Sub Text2_Change()
            If Me.TxtModFlg.text <> "R" Then
            Text6.text = Translate(0, Text2.text)
            End If
End Sub

Private Sub Text2_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub Text3_Change()
If Me.TxtModFlg.text <> "R" Then
Text7.text = Translate(0, Text3.text)
End If
End Sub

Private Sub Text3_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub Text4_Change()
If Me.TxtModFlg.text <> "R" Then
Text8.text = Translate(0, Text4.text)
End If
End Sub

Private Sub Text4_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub Text5_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub Text6_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub Text7_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub Text8_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

 







Private Sub txtDateEndIndustrial_Change()

If Me.TxtModFlg.text <> "R" Then
        txtDateEndIndustrialHijri.value = ToHijriDate(txtDateEndIndustrial.value)
End If


End Sub


'Private Sub txtEmployeeInsurance_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, Me.txtEmployeeInsurance.text, 0)
'End Sub

Private Sub TxtEmp_Comm_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtEmp_Comm.text, 0)
End Sub

Private Sub TxtEmpProfitCom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtEmpProfitCom.text, 0)
End Sub

Private Sub txtid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

            rs.Find "fullcode=" & txtid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
Retrive
     


End If

End Sub

'Private Sub TxtInsurValue_KeyPress(KeyAscii As Integer)
'    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtInsurValue.text, 0)
'End Sub

Private Sub TxtModFlg_Change()
   On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
       ' dcjopstatus.Enabled = False
'Fra(13).Visible = True
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„ÊŸ›Ì‰"
            Else
                Me.Caption = "Employees Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            If SystemOptions.CantWorkwithComponenetinEmpScr = True Then
                ALLButton2(4).Enabled = False
            Else
                ALLButton2(4).Enabled = True
            End If
        dcjopstatus.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
     ' Frame8.Enabled = False
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPTxtEmpName.locked = True
            '        XPCboProfLevel.Locked = True
            XPTxtProfMail.locked = True
            xptxtphone.locked = True
            TxtSalary.locked = True
'            txtemp_code.locked = True
            XPTxtmobile.locked = True
            XPMTxtRemarks.locked = True
            Me.Txt_placEkama.locked = True
            Me.Txt_DateEndLinc.Enabled = False
           'nnnn   Me.Txt_DateEndekama.Enabled = False
            Me.Txt_DateEndpoket.Enabled = False
          'nnnn   Me.Txt_DateExpEkama.Enabled = False
            Me.Txt_DateExpLinc.Enabled = False
            Me.Txt_DateExppoket.Enabled = False
            Me.Txt_NumEkama.locked = True
            Me.Txt_NumLicn.locked = True
            Me.Txt_placEkama.locked = True
            Me.Tet_NumPoket.locked = True
            Tet_NumPoket.locked = True
            Me.Txt_NumPasp.locked = True
            Me.Txt_NumPaspOld.locked = True
            Me.Txt_DateExpPasp.Enabled = False
            Me.Txt_DatePasp.Enabled = False
            Me.Chk_EndWork.Enabled = False
            chkStop.Enabled = False
         '   Me.Chk_Stkala.Enabled = False
            '            Me.CmdEstkala.Enabled = False
            Me.DtDate.Enabled = False
'DCGroupID.Enabled = False
'DcboJobsType.Enabled = False
'DcboEmpDepartments.Enabled = False
'Txt_NumEkama.Enabled = False
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
        Frame8.Visible = False
Fra(20).Visible = False
dcjopstatus.Enabled = True
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„ÊŸ›Ì‰( ”ÃÌ· ”Ã· ÃœÌœ)"
            Else
                Me.Caption = "Employees Data(Enter New Record)"
            End If
'            Fra(13).Visible = False
          Frame8.Enabled = True
            If SystemOptions.CantWorkwithComponenetinEmpScr = True Then
                ALLButton2(4).Enabled = False
            Else
                ALLButton2(4).Enabled = True
            End If
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
       ' dcjopstatus.Enabled = True
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            XPTxtEmpName.locked = False
            '        XPCboProfLevel.Locked = False
            XPTxtProfMail.locked = False
            xptxtphone.locked = False
            XPTxtmobile.locked = False
            TxtSalary.locked = False
            XPMTxtRemarks.locked = False
'            txtemp_code.locked = False
            Me.Txt_placEkama.locked = False
            Me.Txt_DateEndLinc.Enabled = True
            'nnnn  Me.Txt_DateEndekama.Enabled = True
            Me.Txt_DateEndpoket.Enabled = True
         'nnnn    Me.Txt_DateExpEkama.Enabled = True
            Me.Txt_DateExpLinc.Enabled = True
            Me.Txt_DateExppoket.Enabled = True
            Me.Txt_NumEkama.locked = False
            Me.Txt_NumLicn.locked = False
            Me.Txt_placEkama.locked = False
            Me.Tet_NumPoket.locked = False
            Tet_NumPoket.locked = False
            Me.Txt_NumPasp.locked = False
            Me.Txt_NumPaspOld.locked = False
            Me.Txt_DateExpPasp.Enabled = True
            Me.Txt_DatePasp.Enabled = True
            Me.Chk_EndWork.Enabled = True
            chkStop.Enabled = True
           ' Me.Chk_Stkala.Enabled = True
            '            Me.CmdEstkala.Enabled = False
            Me.DtDate.Enabled = False
            DCGroupID.Enabled = True
DcboJobsType.Enabled = True
DcboEmpDepartments.Enabled = True
Txt_NumEkama.Enabled = True

        Case "E"
dcjopstatus.Enabled = False
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„ÊŸ›Ì‰(  ⁄œÌ· )"
            Else
                Me.Caption = "Employees Data(Edit Current Record)"
            End If
'Fra(3).Enabled = False

' Frame8.Enabled = False
'dcjopstatus.Enabled = False
'Fra(13).Visible = False
          '  ALLButton2(4).Enabled = True
If SystemOptions.CantWorkwithComponenetinEmpScr = True Then
                ALLButton2(4).Enabled = False
            Else
                ALLButton2(4).Enabled = True
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
            xptxtphone.locked = False
            XPTxtmobile.locked = False
            XPMTxtRemarks.locked = False
'            txtemp_code.locked = False
            Me.Txt_NumPasp.locked = False
            Me.Txt_DateExpPasp.Enabled = True
            Me.Txt_DatePasp.Enabled = True

            Me.Txt_placEkama.locked = False
            Me.Txt_DateEndLinc.Enabled = True
           'nnnn   Me.Txt_DateEndekama.Enabled = True
            Me.Txt_DateEndpoket.Enabled = True
          'nnnn   Me.Txt_DateExpEkama.Enabled = True
            Me.Txt_DateExpLinc.Enabled = True
            Me.Txt_DateExppoket.Enabled = True
            Me.Txt_NumEkama.locked = False
            Me.Txt_NumLicn.locked = False
            Me.Txt_placEkama.locked = False
            Me.Tet_NumPoket.locked = False
            Tet_NumPoket.locked = False
            Me.Chk_EndWork.Enabled = True
            chkStop.Enabled = True
          '  Me.Chk_Stkala.Enabled = True
            '            Me.CmdEstkala.Enabled = False
            Me.DtDate.Enabled = False
            
            
'            DCGroupID.Enabled = False
'DcboJobsType.Enabled = False
'DcboEmpDepartments.Enabled = False
If Txt_NumEkama.text <> "" Then
'Txt_NumEkama.Enabled = False
Else
Txt_NumEkama.Enabled = True
End If

    End Select
If SystemOptions.AllowupdateJobStatus = True Then
dcjopstatus.Enabled = True
End If
    Exit Sub
ErrTrap:
End Sub

'Private Sub TxtOtherDiscounts_KeyPress(KeyAscii As Integer)
 '   KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOtherDiscounts.text, 0)
'End Sub

Private Sub TxtSalary_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSalary.text, 0)
End Sub

Private Sub VSFlexGrid1_Click()

End Sub

 

Private Sub VSFlexGrid3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim k As Integer
Dim StrComboList As String
 Dim StrAccountCode As String
 Dim Ratval As Integer
    
    With VSFlexGrid3

        Select Case .ColKey(Col)
              Case "name"
 StrAccountCode = .ComboData
                             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SpecID1"), False, True)
                .TextMatrix(Row, .ColIndex("relationtype")) = StrAccountCode
  StrSQL = "select * from takeem where Id=" & val(StrAccountCode)
               rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("MaxValue")) = IIf(IsNull(rs("MaxDeg").value), 0, rs("MaxDeg").value)
                     .TextMatrix(Row, .ColIndex("MinValue")) = IIf(IsNull(rs("MiniDeg").value), 0, rs("MiniDeg").value)
                Else
                   .TextMatrix(Row, .ColIndex("MaxValue")) = ""
                     .TextMatrix(Row, .ColIndex("MinValue")) = ""
                End If
                
    Case "Value"
    Ratval = val(.TextMatrix(Row, .ColIndex("Value")))
    If Ratval <> 0 Then
     StrSQL = " SELECT     Id, name, namee, FromR, TOR"
StrSQL = StrSQL & " From dbo.TblRating"
StrSQL = StrSQL & " Where (TOR >= " & Ratval & ") And (FromR <= " & Ratval & ")"
rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If rs.RecordCount > 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("Rate")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                    ' .TextMatrix(Row, .ColIndex("MinValue")) = IIf(IsNull(rs("Id").value), "", rs("Id").value)
                Else
                 .TextMatrix(Row, .ColIndex("Rate")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
    
    End If
    End If
    End If
                End Select
                 If Row = .rows - 1 Then
    
            .rows = .rows + 1
        End If
                End With
             ReLineGrid
End Sub

Private Sub VSFlexGrid3_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid3
Select Case .ColKey(Col)
Case "Value"
.ComboList = ""
Case "MinValue"
Cancel = True
Case "MaxValue"
Cancel = True
Case "Rate"
.ComboList = ""
Case "des"
.ComboList = ""

End Select
End With
End Sub

Private Sub VSFlexGrid3_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim LngRow As Double
    With VSFlexGrid3

        Select Case .ColKey(Col)
Case "name"
  StrSQL = "select * from takeem"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                
                    StrComboList = VSFlexGrid3.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = VSFlexGrid3.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

                .ComboList = StrComboList
                

                
        End Select

    End With
End Sub

Private Sub workfrom_Change()
If Me.TxtModFlg.text <> "R" Then
     
         workfromH.value = ToHijriDate(workfrom.value)
       
End If
End Sub

Private Sub workto_Change()
If Me.TxtModFlg.text <> "R" Then
     
         worktoH.value = ToHijriDate(workto.value)
       
End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
'    On Error GoTo ErrTrap

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
           ' MsgBox "·„ Ì „  ÕœÌœ Õ”«»  –„„ «·„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
           ' create_accounts = False: Exit Function
         
        End If
    End If
        
    Account_Code_dynamic1 = get_account_code_branch(29, my_branch)
        
    If Account_Code_dynamic1 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic1 = "NO account" Then
           ' MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·«ÃÊ— «·„” Õﬁ… ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
         
           ' create_accounts = False: Exit Function
        End If
    End If
        
    Account_Code_dynamic2 = get_account_code_branch(30, my_branch)
        
    If Account_Code_dynamic2 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic2 = "NO account" Then
           ' MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „Œ’’«   «Ã«“…  ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
         
           ' create_accounts = False: Exit Function
        End If
    End If
        
    Account_Code_dynamic3 = get_account_code_branch(65, my_branch)
        
    If Account_Code_dynamic3 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic3 = "NO account" Then
           ' MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „œ›Ê⁄«  „ﬁœ„Â  «·„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
           ' create_accounts = False: Exit Function
         
        End If
    End If
        
    
    
    
    
    Account_Code_dynamic4 = get_account_code_branch(74, my_branch)
            
    If Account_Code_dynamic4 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic4 = "NO account" Then
         '   MsgBox "·„ Ì „  ÕœÌœ Õ”«»     „Œ’’«  ‰Â«Ì… Œœ„…  «·„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
         '   create_accounts = False: Exit Function
         
        End If
        
        
        
        
    End If
    
        
   Account_Code_dynamic5 = get_account_code_branch(93, my_branch)
            
             If Account_Code_dynamic5 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        create_accounts = False: Exit Function
    Else

        If Account_Code_dynamic5 = "NO account" Then
          '  MsgBox "·„ Ì „  ÕœÌœ Õ”«»     „Œ’’«     –«ﬂ—  «·„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
          '  create_accounts = False: Exit Function
         
        End If
    End If
            
            
    create_accounts = True
End Function
Function GETlASTiSSUEDATE(Emp_id As Integer, Optional novalue As Boolean) As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
  sql = "SELECT     max(workdate) AS MaxDate from   dbo.TblEmbarkation  WHERE     (Emp_ID = " & Emp_id & ")"
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    If Not IsNull(rs("MaxDate").value) Then
 GETlASTiSSUEDATE = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
novalue = False
Else
 GETlASTiSSUEDATE = Date
 novalue = True
 End If
 Else
 GETlASTiSSUEDATE = Date
 novalue = True
    End If

End Function
Public Sub Retrive(Optional Lngid As Long = 0)
' On Error GoTo ErrTrap
 Dim fullstr As String
ReloadCompo
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "Emp_ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

'tXTPassWord.Text = IIf(IsNull(rs("Password").value), "", (rs("Password").value))
cmbInsuranceRenew.ListIndex = 1
    cmbToM.ListIndex = 1
    
    XPTxtEmpID.text = IIf(IsNull(rs("Emp_ID").value), "", val(rs("Emp_ID").value))
      Me.TxtInsuranceNo.text = IIf(IsNull(rs("InsuranceNO").value), "", rs("InsuranceNO").value)
''////////////
Me.DcbDepartment2.BoundText = IIf(IsNull(rs("DeptID2").value), "", rs("DeptID2").value)
TxtMachinCode(0).text = IIf(IsNull(rs("MachinCode").value), "", rs("MachinCode").value)
txtBank(0).text = IIf(IsNull(rs("To_Employee_name").value), "", rs("To_Employee_name").value)

TxtMachinCode(1).text = IIf(IsNull(rs("SalaryCode").value), "", rs("SalaryCode").value)
Me.DcbBanck(1).text = IIf(IsNull(rs("BanckName").value), "", rs("BanckName").value)
Me.txtBank(1).text = IIf(IsNull(rs("BankIBan").value), "", rs("BankIBan").value)
Me.txtBank(2).text = IIf(IsNull(rs("BankIAddress").value), "", rs("BankIAddress").value)
HowIqamaEndH.value = IIf(IsNull(rs("HowIqamaEndH").value), ToHijriDate(Date), rs("HowIqamaEndH").value)
If Not IsNull(rs("NoAdded").value) Then
If rs("NoAdded").value = 1 Then
ChNoAdded.value = vbChecked
Else
ChNoAdded.value = vbUnchecked
End If
Else
ChNoAdded.value = vbUnchecked
End If
'''/////////
' Me.TxtInsurValue.text = IIf(IsNull(rs("InsuranceValue").value), "", rs("InsuranceValue").value)
      '  Me.TxtOtherDiscounts.text = IIf(IsNull(rs("OtherDiscounts").value), "", rs("OtherDiscounts").value)

 
   ' Me.txtEmployeeInsurance.text = IIf(IsNull(rs("EmployeeInsurance").value), "", rs("EmployeeInsurance").value)
    
    
   'nnnn If IsNull(rs("BlnceVocat").value) Then
  'nnnn   TxtBlncVoc.text = 0
  'nnnn   Else
 'nnnn  Me.TxtBlncVoc.text = val(rs("BlnceVocat").value)
'nnnn   End If
' val (rs("BlnceVocat").value)
    'OPENINGBALNCESDATA
    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
    txtDateEndIndustrial.value = IIf(IsNull(rs("DateEndIndustrial").value), Date, IIf(IsDate(rs("DateEndIndustrial").value), rs("DateEndIndustrial").value, Date))
    txtDateEndIndustrialHijri.value = IIf(IsNull(rs("DateEndIndustrialHijri").value), Date, rs("DateEndIndustrialHijri").value)
    If Not (IsNull(rs("OpenBalanceDate").value)) Then
        Me.Dtp.value = rs("OpenBalanceDate").value
      '  Me.Dtp1.value = rs("OpenBalanceDate").value
      '  Me.Dtp2.value = rs("OpenBalanceDate").value
        ' Me.Dtp.Enabled = True
    Else
    
        Me.Dtp.value = Date
     '   Me.Dtp1.value = Date
     '   Me.Dtp2.value = Date
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




'444444444444444444
    If Not IsNull(rs("OpenBalanceType4").value) Then
        Me.TxtOpenBalance4.text = IIf(IsNull(rs("OpenBalance4")), "", Trim(rs("OpenBalance4")))

        If rs("OpenBalanceType4").value = 0 Then
            OptType4(0).value = True
            OptType4_Click 0
        ElseIf rs("OpenBalanceType4").value = 1 Then
            OptType4(1).value = True
            OptType4_Click 1
        End If
    
    Else
        Me.TxtOpenBalance4.text = 0
        Me.OptType4(2).value = True
        OptType4_Click 2
    End If
    
'444444444444444444444444

Txtcommission = IIf(IsNull(rs("Commission")), "", Trim(rs("Commission")))
'555555555555555555
    If Not IsNull(rs("OpenBalanceType5").value) Then
        Me.TxtOpenBalance5.text = IIf(IsNull(rs("OpenBalance5")), "", Trim(rs("OpenBalance5")))

        If rs("OpenBalanceType5").value = 0 Then
            OptType5(0).value = True
            OptType5_Click 0
        ElseIf rs("OpenBalanceType5").value = 1 Then
            OptType5(1).value = True
            OptType5_Click 1
        End If
    
    Else
        Me.TxtOpenBalance5.text = 0
        Me.OptType5(2).value = True
        OptType5_Click 2
    End If
    
'555555555555555555555555
    If XPTxtEmpID.text <> "" Then
        DBPix201.ImageClear
        DBPix202.ImageClear

        If Dir(system_path & "\" & SystemOptions.ImagesPath & "\" & XPTxtEmpID.text & ".JPG") <> "" Then
            DBPix201.ImageLoadFile (system_path & "\" & SystemOptions.ImagesPath & "\" & XPTxtEmpID.text & ".JPG")
        End If

        If Dir(system_path & "\" & SystemOptions.ImagesPath & "\sign" & XPTxtEmpID.text & ".JPG") <> "" Then
            DBPix202.ImageLoadFile (system_path & "\" & SystemOptions.ImagesPath & "\sign" & XPTxtEmpID.text & ".JPG")
        End If
 
    End If

If IsNull(rs("SalaryType").value) Then
OptSalaryType(0).value = True
Else
        If (rs("SalaryType").value) = 0 Then
        OptSalaryType(0).value = True
        TxtPercentage.text = ""
         TxtBYHour.text = ""
        ElseIf (rs("SalaryType").value) = 1 Then
        OptSalaryType(1).value = True
         TxtPercentage.text = IIf(IsNull(rs("Percentage")), "", (rs("Percentage")))
         TxtBYHour.text = ""
        ElseIf (rs("SalaryType").value) = 2 Then
        OptSalaryType(2).value = True
              TxtPercentage.text = ""
         TxtBYHour.text = IIf(IsNull(rs("BYHour")), "", (rs("BYHour")))
         
        ElseIf (rs("SalaryType").value) = 3 Then
        OptSalaryType(3).value = True
               TxtPercentage.text = ""
         TxtBYHour.text = ""
       ElseIf (rs("SalaryType").value) = 4 Then
        OptSalaryType(4).value = True
               TxtPercentage.text = ""
         TxtBYHour.text = ""
        Else
        OptSalaryType(0).value = True
               TxtPercentage.text = ""
         TxtBYHour.text = ""
         
        End If
End If
If Not IsNull(rs("TypeEmp").value) Then
If rs("TypeEmp").value = 1 Then
TypeEmp(1).value = True
Else
TypeEmp(0).value = True
End If
Else
TypeEmp(0).value = True
End If
dcdean.BoundText = IIf(IsNull(rs("DeanID").value), "", rs("DeanID").value)
If dcdean.text = "" Or val(dcdean.BoundText) = 0 Then
dcdean.BoundText = Get_Employee_religon(IIf(IsNull(rs("dean").value), "", Trim(rs("dean").value)))
End If

DCNationality.BoundText = IIf(IsNull(rs("NationlID").value), "", rs("NationlID").value)
If DCNationality.text = "" Or val(DCNationality.BoundText) = 0 Then
DCNationality.BoundText = Get_Employee_Nationality(IIf(IsNull(rs("Nationality").value), "", Trim(rs("Nationality").value)))
End If
    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    
    mangerid.BoundText = IIf(IsNull(rs("mangerid").value), "", rs("mangerid").value)
    
    swapedempid.BoundText = IIf(IsNull(rs("swapedempid").value), "", rs("swapedempid").value)
    swapedempid2.BoundText = IIf(IsNull(rs("swapedempid2").value), "", rs("swapedempid2").value)
     fullstr = ""
    Me.txtid.text = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
    fullstr = DCPreFix.text & Me.txtid.text
    
    Me.TxtBankCard.text = IIf(IsNull(rs("BankCard").value), "", rs("BankCard").value)
    Me.txtPrefNatID.text = IIf(IsNull(rs("PrefNatID").value), "", rs("PrefNatID").value)
    
  
 ' TxtEmp_Code.text = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
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
    TxtEmpNotes.text = IIf(IsNull(rs("EmpNotes").value), "", Trim(rs("EmpNotes").value))
    '''////////////// 28 10 2015
     Me.DcbBanck(0).BoundText = IIf(IsNull(rs("BankCode").value), "", Trim(rs("BankCode").value))
     Me.DcbMatrial.ListIndex = IIf(IsNull(rs("MaritalStatus").value), -1, rs("MaritalStatus").value)
     Me.DcbContractType.BoundText = IIf(IsNull(rs("ContractID").value), "", rs("ContractID").value)
    
    ''''''''''''''''
        
            If SystemOptions.UserInterface = ArabicInterface Then
XPLbl(60).Caption = " ﬂÊœ «·„ÊŸ› " & fullstr & " «·«”„ : " & XPTxtEmpName.text
    Else
XPLbl(60).Caption = " Emp Code " & fullstr & " Name : " & XPTxtEmpNamee.text
    End If
    
    
        TxtAccountCode.text = IIf(IsNull(rs("Account_code").value), "", Trim(rs("Account_code").value))
   ' DcboCreditSide.BoundText = IIf(IsNull(rs("Account_code1").value), "", Trim(rs("Account_code1").value))

    txthdodno.text = IIf(IsNull(rs("hdodno").value), "", Trim(rs("hdodno").value))
    Me.DCRegionID.BoundText = IIf(IsNull(rs("RegionID").value), "", rs("RegionID").value)


    txthdomnfaz.text = IIf(IsNull(rs("hdomnfaz").value), "", Trim(rs("hdomnfaz").value))

   ' TxtSalary.text = IIf(IsNull(rs("Emp_Salary").value), "", Trim(rs("Emp_Salary").value))
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
    xptxtphone.text = IIf(IsNull(rs("Emp_Phone").value), "", Trim(rs("Emp_Phone").value))
    XPTxtmobile.text = IIf(IsNull(rs("Emp_mobile").value), "", Trim(rs("Emp_mobile").value))
    XPMTxtRemarks.text = IIf(IsNull(rs("Emp_Remark").value), "", Trim(rs("Emp_Remark").value))
    TxtEmp_Comm.text = IIf(IsNull(rs("Emp_Comm").value), "", Trim(rs("Emp_Comm").value))
    TxtEmpProfitCom.text = IIf(IsNull(rs("EmpProfitCom").value), "", Trim(rs("EmpProfitCom").value))
    Txt_placEkama.text = IIf(IsNull(rs("placeEkama").value), "", Trim(rs("placeEkama").value))
    Txt_NumEkama.text = IIf(IsNull(rs("NumEkama").value), "", Trim(rs("NumEkama").value))
    Txt_NumLicn.text = IIf(IsNull(rs("NumLicn").value), "", Trim(rs("NumLicn").value))
    Tet_NumPoket.text = IIf(IsNull(rs("NumPoket").value), "", Trim(rs("NumPoket").value))
'***************
TxtDriverLicense.text = IIf(IsNull(rs("DriverLicense").value), "", (rs("DriverLicense").value))


'**************
 

    Txt_NumPasp.text = IIf(IsNull(rs("NumPasp").value), "", Trim(rs("NumPasp").value))
    Txt_NumPaspOld.text = IIf(IsNull(rs("NumPaspOld").value), "", Trim(rs("NumPaspOld").value))
    txtKafelID.text = IIf(IsNull(rs("KafelID").value), "", Trim(rs("KafelID").value))
    Me.DcbKafelName.BoundText = IIf(IsNull(rs("KafelName").value), "", Trim(rs("KafelName").value))

    TxtKafeltEL.text = IIf(IsNull(rs("kafeltel").value), "", Trim(rs("kafeltel").value))

    txtkafeladd.text = IIf(IsNull(rs("kafeladd").value), "", Trim(rs("kafeladd").value))

    Me.DcbPasplace.BoundText = IIf(IsNull(rs("pasplace").value), "", Trim(rs("pasplace").value))
  '  DcNationality.text = IIf(IsNull(rs("Nationality").value), "", Trim(rs("Nationality").value))
  '  Dcdean.text = IIf(IsNull(rs("dean").value), "", Trim(rs("dean").value))

'Get_Employee_religon
'Get_Employee_Nationality




    'Txt_NotEndWork.text = IIf(IsNull(rs("Notsstkala").value), "", Trim(rs("Notsstkala").value))
   
   ' If rs("ChekStkala").value = True Then
      '  Chk_Stkala.value = Checked
   ' Else
      '  Chk_Stkala.value = Unchecked
'
   ' End If

    If rs("chkStop").value = True Then

        chkStop.value = Checked
    Else
        chkStop.value = Unchecked
    End If
    
    
    If rs("chkShowTasks").value = True Then

        chkShowTasks.value = Checked
    Else
        chkShowTasks.value = Unchecked
    End If
    
    


    If rs("ChekEndWork").value = True Then

        Chk_EndWork.value = Checked
    Else
        Chk_EndWork.value = Unchecked
    End If

    DtDate.value = IIf(IsNull(rs("EndWork").value), Date, rs("EndWork").value)
    If IsNull(rs("BignDateWork").value) Then
    Frame8.Visible = False
    Else
    Frame8.Visible = True
    DTPicker1.value = IIf(IsNull(rs("BignDateWork").value), Date, rs("BignDateWork").value)
    End If
    DBDOB.value = IIf(IsNull(rs("DOB").value), Date, IIf(IsDate(rs("DOB").value), rs("DOB").value, Date))
    
    DpIssuingDriverCardDate.value = IIf(IsNull(rs("IssuingDriverCardDate").value), Date, IIf(IsDate(rs("IssuingDriverCardDate").value), rs("IssuingDriverCardDate").value, Date))
    DpCardDriverExpireDate.value = IIf(IsNull(rs("CardDriverExpireDate").value), Date, IIf(IsDate(rs("CardDriverExpireDate").value), rs("CardDriverExpireDate").value, Date))

    If IsNull(rs("workstate").value) Then
        Me.CboWorkState.ListIndex = -1
    Else

        If rs("workstate").value = 1 Then
            Me.CboWorkState.ListIndex = 0
        ElseIf rs("workstate").value = 0 Then
            Me.CboWorkState.ListIndex = 1
        End If
    End If
    
    
    
 

 Me.cboPayType.ListIndex = IIf(IsNull(rs("PayType").value), 0, rs("PayType").value)
       
        
        
 
    
    
    
 If rs("sex").value = 1 Then
            Me.Dcbsex.ListIndex = 0
        ElseIf rs("sex").value = 2 Then
            Me.Dcbsex.ListIndex = 1
        End If
    
    Me.DcboEmpDepartments.BoundText = IIf(IsNull(rs("DepartmentID").value), "", rs("DepartmentID").value)
    Me.DcGrade.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)
    'Me.Dcbsex.ListIndex = val(IIf(IsNull(rs("Sex").value), -1, rs("Sex").value))
    
 '   Me.DateMoveNo.value = IIf(IsNull(rs("gradeID").value), Date, rs("gradeID").value)
   Me.DateMoveNo.value = IIf(IsNull(rs("DateMoveNo").value), Date, rs("DateMoveNo").value)
  If IsNull(rs("DateMoveNo").value) Then
  DateMoveNo.Visible = False
  End If
  '
    Me.DcCostCenter.BoundText = IIf(IsNull(rs("cost_center_id").value), "", rs("cost_center_id").value)
    Me.DCBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
Me.DCGroupID.BoundText = IIf(IsNull(rs("GroupID").value), "", rs("GroupID").value)
    Me.DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
    Me.DcbSection.BoundText = IIf(IsNull(rs("SectionID").value), "", rs("SectionID").value)
    Me.DcboJobsType1.BoundText = IIf(IsNull(rs("JobTypeID1").value), "", rs("JobTypeID1").value)
    Me.DcboJobsType2.BoundText = IIf(IsNull(rs("JobTypeID2").value), "", rs("JobTypeID2").value)
    Me.DcboJobsType3.BoundText = IIf(IsNull(rs("JobTypeID3").value), "", rs("JobTypeID3").value)
    
    
    Me.DcboSpecifications.BoundText = IIf(IsNull(rs("SpecificationID").value), "", rs("SpecificationID").value)
    Me.TxtRegion.text = IIf(IsNull(rs("Region").value), "", rs("Region").value)

    Me.dcjopstatus.BoundText = IIf(IsNull(rs("jopstatusid").value), "", rs("jopstatusid").value)
    Me.dcproject.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)

     
    If IsNull(rs("InsuranceState").value) Then
        Me.CboInsuranceState.ListIndex = 0
    Else
        Me.CboInsuranceState.ListIndex = rs("InsuranceState").value
    End If
    
    
    If IsNull(rs("InsuranceRenew").value) Then
        Me.cmbInsuranceRenew.ListIndex = 1
    Else
        Me.cmbInsuranceRenew.ListIndex = rs("InsuranceRenew").value
    End If
    
    If IsNull(rs("ToM").value) Then
        Me.cmbToM.ListIndex = 1
    Else
        Me.cmbToM.ListIndex = rs("ToM").value
    End If
    
    Me.txtCopyNo.text = IIf(IsNull(rs("CopyNo").value), "", rs("CopyNo").value)
    Me.txtInsuranceRenewDate.value = IIf(IsNull(rs("InsuranceRenewDate").value), Date, rs("InsuranceRenewDate").value)
    Me.txtToMDateNew.value = IIf(IsNull(rs("ToMDateNew").value), Date, rs("ToMDateNew").value)
    
    
''''''''/////////
    Me.DTPicker3.value = IIf(IsNull(rs("InstanceDateM").value), Date, rs("InstanceDateM").value)
    NourHijriCal1.value = IIf(IsNull(rs("InstanceDateH").value), "", rs("InstanceDateH").value)
'''''''''''///
   
    Me.TxtVisaNo.text = IIf(IsNull(rs("VisaNo").value), "", rs("VisaNo").value)
    
    
   
'txtEmployeeInsurance
    Txt_DateExpLinc.value = IIf(IsNull(rs("DateExpLinc").value), Date, rs("DateExpLinc").value)
    Txt_DateEndLinc.value = IIf(IsNull(rs("DateEndLinc").value), Date, rs("DateEndLinc").value)

    Txt_DateExppoket.value = IIf(IsNull(rs("Dateexppoket").value), Date, rs("Dateexppoket").value)
    Txt_DateEndpoket.value = IIf(IsNull(rs("dateendpoket").value), Date, rs("dateendpoket").value)

   'nnnn Txt_DateExpEkama.value = IIf(IsNull(rs("DateExpoekama").value), Date, rs("DateExpoekama").value)
   'nnnn  Txt_DateEndekama.value = IIf(IsNull(rs("DateEndekama").value), Date, rs("DateEndekama").value)

 

    Txt_DateExpEkamaH.value = IIf(IsNull(rs("DateExpoekamah").value), Date, rs("DateExpoekamah").value)
    Txt_DateEndekamah.value = IIf(IsNull(rs("DateEndekamah").value), Date, rs("DateEndekamah").value)

    Txt_DateExpLincH.value = IIf(IsNull(rs("DateExpLincH").value), Date, rs("DateExpLincH").value)
    Txt_DateEndLincH.value = IIf(IsNull(rs("DateEndLincH").value), Date, rs("DateEndLincH").value)


'**************************************************
DpDriverLicenseend.value = IIf(IsNull(rs("DriverLicenseend").value), Date, rs("DriverLicenseend").value)
    txtDriverLicenseStartdH.value = IIf(IsNull(rs("DriverLicenseStartdH").value), ToHijriDate(Date), rs("DriverLicenseStartdH").value)
    txtDriverLicenseendH.value = IIf(IsNull(rs("DriverLicenseendH").value), ToHijriDate(Date), rs("DriverLicenseendH").value)
'**************************************************



    Txt_DateExppoketH.value = IIf(IsNull(rs("Dateexppoketh").value), Date, rs("Dateexppoketh").value)
    Txt_DateEndpoketH.value = IIf(IsNull(rs("dateendpoketh").value), Date, rs("dateendpoketh").value)

    Txt_DateExpPasp.value = IIf(IsNull(rs("DateExpPasp").value), Date, rs("DateExpPasp").value)
    Txt_DatePasp.value = IIf(IsNull(rs("DateEndPasp").value), Date, rs("DateEndPasp").value)
    txthdoddate.value = IIf(IsNull(rs("hdoddate").value), Date, rs("hdoddate").value)
'


    DOBH.value = IIf(IsNull(rs("DOBH").value), Date, rs("DOBH").value)
    
    
  IssueDateH.value = IIf(IsNull(rs("IssueDateH").value), Date, rs("IssueDateH").value)
    
   LastDateH.value = IIf(IsNull(rs("LastDateH").value), Date, rs("LastDateH").value)
    
       LastDate.value = IIf(IsNull(rs("LastDate").value), Date, rs("LastDate").value)

    DpIssuingDriverCardDateH.value = IIf(IsNull(rs("IssuingDriverCardDateH").value), Date, rs("IssuingDriverCardDateH").value)
    DpCardDriverExpireDateH.value = IIf(IsNull(rs("CardDriverExpireDateH").value), Date, rs("CardDriverExpireDateH").value)

 Dim novalue As Boolean
   If IsNull(rs("lastHolidaydate").value) Then
   Fra(20).Visible = False
   Else
   Fra(20).Visible = True
     lastHolidaydateH.value = IIf(IsNull(rs("lastHolidaydateH").value), "", rs("lastHolidaydateH").value)
   lastHolidaydate.value = IIf(IsNull(rs("lastHolidaydate").value), Date, rs("lastHolidaydate").value)
   End If

'lastHolidaydate = GETlASTiSSUEDATE(val(XPTxtEmpID.text), novalue)
'If novalue = True Then
'sql = "update TblEmployee set   ChekDateIQ =0  where Emp_ID =" & val(XPTxtEmpID.text) & ""
'                                    Cn.Execute sql
'  Else
' sql = "update TblEmployee set   ChekDateIQ =1  where Emp_ID =" & val(XPTxtEmpID.text) & ""
'                                    Cn.Execute sql
' lastHolidaydate.Visible = True
' lastHolidaydateH.Visible = True
'
' End If
 
  '
    
    
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
     

          Dim RsDetails As New ADODB.Recordset
    StrSQL = " SELECT     dbo.TblEmpDetails.ID, dbo.TblEmpDetails.Emp_ID, dbo.TblEmpDetails.passportno, dbo.TblEmpDetails.iqamano, dbo.TblEmpDetails.name AS tabe3name, "
StrSQL = StrSQL & "                      dbo.TblEmpDetails.haveinsurance, dbo.TblEmpDetails.insuranceno, dbo.TblEmpDetails.des, dbo.TblEmpDetails.relationtype, dbo.TblRelations.name,"
StrSQL = StrSQL & "                      dbo.TblRelations.namee , dbo.TblEmpDetails.OprType"
StrSQL = StrSQL & " FROM         dbo.TblEmpDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRelations ON dbo.TblEmpDetails.relationtype = dbo.TblRelations.ID"
StrSQL = StrSQL & " Where (dbo.TblEmpDetails.Emp_id = " & val(Me.XPTxtEmpID.text) & ") And (dbo.TblEmpDetails.OprType = 0)"
'StrSQL = StrSQL & "  Where ( OprType =0 and dbo.TblEmpDetails.Emp_id = " & val(Me.XPTxtEmpID.text) & ")"

  

    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With Grid
     .Clear flexClearScrollable, flexClearEverything
     .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("tabe3name").value), "", RsDetails("tabe3name").value)
              .TextMatrix(i, .ColIndex("relationtype")) = IIf(IsNull(RsDetails("relationtype").value), "", RsDetails("relationtype").value)
              .TextMatrix(i, .ColIndex("passportno")) = IIf(IsNull(RsDetails("passportno").value), "", RsDetails("passportno").value)
              .TextMatrix(i, .ColIndex("iqamano")) = IIf(IsNull(RsDetails("iqamano").value), "", RsDetails("iqamano").value)
              .TextMatrix(i, .ColIndex("haveinsurance")) = IIf(IsNull(RsDetails("haveinsurance").value), 0, IIf((RsDetails("haveinsurance").value) = True, 1, 0))
              .TextMatrix(i, .ColIndex("insuranceno")) = IIf(IsNull(RsDetails("insuranceno").value), "", RsDetails("insuranceno").value)
              If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(i, .ColIndex("relationtypename")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
              Else
               .TextMatrix(i, .ColIndex("relationtypename")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
              End If
                  .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("des").value), "", RsDetails("des").value)
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
    
    
     
     
 
    
       StrSQL = " select * from TblEmpDetails "
StrSQL = StrSQL & "  Where ( OprType =1 and dbo.TblEmpDetails.Emp_id = " & val(Me.XPTxtEmpID.text) & ")"

  
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With grid01
     .Clear flexClearScrollable, flexClearEverything
     .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
             .TextMatrix(i, .ColIndex("qualicationEntity")) = IIf(IsNull(RsDetails("qualicationEntity").value), "", RsDetails("qualicationEntity").value)
             
              .TextMatrix(i, .ColIndex("yearofqualication")) = IIf(IsNull(RsDetails("yearofqualication").value), "", RsDetails("yearofqualication").value)
     
                  .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("des").value), "", RsDetails("des").value)
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
        
    
 
 
 
 
       StrSQL = " select * from TblEmpDetails "
StrSQL = StrSQL & "  Where ( OprType =2 and dbo.TblEmpDetails.Emp_id = " & val(Me.XPTxtEmpID.text) & ")"

  
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With grid02
     .Clear flexClearScrollable, flexClearEverything
     .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
             .TextMatrix(i, .ColIndex("qualicationEntity")) = IIf(IsNull(RsDetails("qualicationEntity").value), "", RsDetails("qualicationEntity").value)
             
              .TextMatrix(i, .ColIndex("workfrom")) = IIf(IsNull(RsDetails("workfrom").value), "", RsDetails("workfrom").value)
     .TextMatrix(i, .ColIndex("workfromH")) = IIf(IsNull(RsDetails("workfromH").value), "", RsDetails("workfromH").value)
     .TextMatrix(i, .ColIndex("workto")) = IIf(IsNull(RsDetails("workto").value), "", RsDetails("workto").value)
     .TextMatrix(i, .ColIndex("worktoH")) = IIf(IsNull(RsDetails("worktoH").value), "", RsDetails("worktoH").value)
     
                  .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("des").value), "", RsDetails("des").value)
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
     




      ' StrSQL = " select * from TblEmpHolidaysDetails "
'StrSQL = StrSQL & "  Where ( dbo.TblEmpHolidaysDetails.Emp_id = " & val(Me.XPTxtEmpID.text) & ")"
StrSQL = " SELECT     stratDate, stratDateH, AcuDate, AcuDateH, NoDayAct, NoDayDelay, EndDateH, EndDate, EmpID , Remark ,NoVacation "
StrSQL = StrSQL & " From dbo.TblVocationEntitlements"
StrSQL = StrSQL & " Where (empid = " & val(Me.XPTxtEmpID.text) & ")"
  
    Dim strdated As Long
    Dim strdateM As Long
   
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With Grid20
     .Clear flexClearScrollable, flexClearEverything
     .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             If Not IsNull(RsDetails("stratDate").value) Then
             .TextMatrix(i, .ColIndex("fromdate")) = IIf(IsNull(RsDetails("stratDate").value), "", RsDetails("stratDate").value)
            .TextMatrix(i, .ColIndex("fromdateh")) = IIf(IsNull(RsDetails("stratDateH").value), ToHijriDate(RsDetails("stratDate").value), RsDetails("stratDateH").value)
            End If
            If Not IsNull(RsDetails("EndDate").value) Then
             .TextMatrix(i, .ColIndex("DateExpectedH")) = IIf(IsNull(RsDetails("EndDateH").value), ToHijriDate(RsDetails("EndDate").value), RsDetails("EndDateH").value)
             .TextMatrix(i, .ColIndex("DateExpectedM")) = IIf(IsNull(RsDetails("EndDate").value), "", RsDetails("EndDate").value)
             End If
             If Not IsNull(RsDetails("AcuDate").value) Then
             .TextMatrix(i, .ColIndex("todateh")) = IIf(IsNull(RsDetails("AcuDateH").value), ToHijriDate(RsDetails("AcuDate").value), RsDetails("AcuDateH").value)
             .TextMatrix(i, .ColIndex("todate")) = IIf(IsNull(RsDetails("AcuDate").value), "", RsDetails("AcuDate").value)
             End If
             .TextMatrix(i, .ColIndex("day")) = IIf(IsNull(RsDetails("NoVacation").value), "", RsDetails("NoVacation").value)
             .TextMatrix(i, .ColIndex("a1")) = IIf(IsNull(RsDetails("NoDayAct").value), "", RsDetails("NoDayAct").value)
             .TextMatrix(i, .ColIndex("dayEx")) = IIf(IsNull(RsDetails("NoDayDelay").value), "", RsDetails("NoDayDelay").value)
             
                ' .TextMatrix(i, .ColIndex("DateExpectedM")) = IIf(IsNull(RsDetails("DateExpectedM").value), "", RsDetails("DateExpectedM").value)
                ' If Not IsNull(RsDetails("DateExpectedM").value) And Not IsNull(RsDetails("fromdate").value) Then
                ' strdated = DateDiff("d", RsDetails("fromdate").value, RsDetails("DateExpectedM").value)
                '
             '  .TextMatrix(i, .ColIndex("day")) = strdated
             '  End If
             '  If RsDetails("SpecificHolidyaType1").value = True Then
             '  .TextMatrix(i, .ColIndex("type")) = 1
             '  Else
             ' .TextMatrix(i, .ColIndex("type")) = 0
             '  End If
             '   .TextMatrix(i, .ColIndex("DateExpectedM")) = IIf(IsNull(RsDetails("DateExpectedM").value), "", RsDetails("DateExpectedM").value)
             '    If Not IsNull(RsDetails("todate").value) Then
             '    strdated = DateDiff("d", RsDetails("fromdate").value, RsDetails("todate").value)
             '
             '  .TextMatrix(i, .ColIndex("a1")) = strdated
             '  End If
             '
             '   If Not IsNull(RsDetails("DateExpectedM").value) Then
             '   If Not IsNull(RsDetails("todate").value) Then
             '    strdateM = DateDiff("d", RsDetails("DateExpectedM").value, RsDetails("todate").value)
             '    .TextMatrix(i, .ColIndex("dayEx")) = strdateM
             '    End If
             '    End If
            '.TextMatrix(i, .ColIndex("DateExpectedH")) = IIf(IsNull(RsDetails("DateExpectedH").value), "", RsDetails("DateExpectedH").value)
            '.TextMatrix(i, .ColIndex("fromdateh")) = IIf(IsNull(RsDetails("fromdateh").value), "", RsDetails("fromdateh").value)
            '   .TextMatrix(i, .ColIndex("todateh")) = IIf(IsNull(RsDetails("todateh").value), "", RsDetails("todateh").value)
                   
                   
                  .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("Remark").value), "", RsDetails("Remark").value)
            
                     Dim astrSplitItems() As String
            Dim Result As String
 
    Dim Txtyear As String
    Dim TxtMonth As String
    Dim TxtDay As String
    
 '   result = ExactAge(.TextMatrix(i, .ColIndex("fromdate")), .TextMatrix(i, .ColIndex("todate")))
'
'    astrSplitItems = Split(result, "-")
'    TxtYear = astrSplitItems(0)
'    TxtMonth = astrSplitItems(1)
'    Txtday = astrSplitItems(2)
 
'          .TextMatrix(i, .ColIndex("day")) = Txtday
'           .TextMatrix(i, .ColIndex("month")) = TxtMonth
'            .TextMatrix(i, .ColIndex("year")) = TxtYear
'
            RsDetails.MoveNext
        Next i

    End If
End With
ReLineGrid
    RsDetails.Close
    Set RsDetails = Nothing
  'astr
     '   StrSQL = " select * from TblEmpReign "
''''''''''StrSQL = StrSQL & "  Where ( dbo.TblEmpReign.Emp_ID = " & val(Me.XPTxtEmpID.text) & ")"
StrSQL = " SELECT     dbo.TblAssestes.AsName, dbo.TblAssestes.AsID, dbo.TblAssestes.AsDes, dbo.TblEmpAsestDetails.Remarks, dbo.TblEmpAsestDetails.Qunt,"
 StrSQL = StrSQL & "                     dbo.TblEmpAsest.EmpAsID , dbo.TblEmpAsestDetails.IDAseset,  dbo.TblEmpAsest.EmpAsestID, dbo.TblEmpAsest.PostedDate"
 StrSQL = StrSQL & " FROM         dbo.TblAssestes INNER JOIN"
  StrSQL = StrSQL & "                     dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
  StrSQL = StrSQL & "                     dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID"
'StrSQL = StrSQL & "  Where (dbo.TblEmpAsest.EmpAsestID = " & val(Me.XPTxtEmpID.text) & ") And (dbo.TblEmpAsestDetails.FlagAs Is Null)"
StrSQL = StrSQL & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsestDetails.EmpID = " & val(Me.XPTxtEmpID.text) & ")"
'StrSQL = StrSQL & "  Where (dbo.TblEmpAsest.EmpAsestID = " & val(Me.XPTxtEmpID.text) & ")"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With VSFlexGrid4
     .Clear flexClearScrollable, flexClearEverything
     .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("AsName").value), "", RsDetails("AsName").value)
               .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(RsDetails("Qunt").value), "", RsDetails("Qunt").value)
                
            .TextMatrix(i, .ColIndex("Date")) = IIf(IsNull(RsDetails("PostedDate").value), "", RsDetails("PostedDate").value)
               .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("Remarks").value), "", RsDetails("Remarks").value)
                   
                   
               '   .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("des").value), "", RsDetails("des").value)
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
'''////////
        StrSQL = " select * from TblHealthy "
StrSQL = StrSQL & "  Where ( dbo.TblHealthy.Emp_ID = " & val(Me.XPTxtEmpID.text) & ")"

  
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With VSFlexGrid2
     .Clear flexClearScrollable, flexClearEverything
     .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("ENTITY")) = IIf(IsNull(RsDetails("FromH").value), "", RsDetails("FromH").value)
               .TextMatrix(i, .ColIndex("clinic")) = IIf(IsNull(RsDetails("Clinic").value), "", RsDetails("Clinic").value)
                
            .TextMatrix(i, .ColIndex("Date")) = IIf(IsNull(RsDetails("HealthyDate").value), "", RsDetails("HealthyDate").value)
               .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remrks").value), "", RsDetails("Remrks").value)
                   .TextMatrix(i, .ColIndex("valuecomp")) = IIf(IsNull(RsDetails("CompanyAmount").value), "", RsDetails("CompanyAmount").value)
                 .TextMatrix(i, .ColIndex("Typename")) = IIf(IsNull(RsDetails("HeathyTreat").value), "", RsDetails("HeathyTreat").value)
                 .TextMatrix(i, .ColIndex("DES")) = IIf(IsNull(RsDetails("Des").value), "", RsDetails("Des").value)
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
  
  
  ''//////////////
 ' StrSQL = " select * from TblEmpReign "
'StrSQL = StrSQL & "  Where ( dbo.TblEmpReign.Emp_ID = " & val(Me.XPTxtEmpID.text) & ")"
'
'
'
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'  With VSFlexGrid4
 '    .Clear flexClearScrollable, flexClearEverything
 ''    .Rows = .FixedRows

  '  If Not (RsDetails.BOF Or RsDetails.EOF) Then
  '      RsDetails.MoveFirst
  '       .Rows = .FixedRows + RsDetails.RecordCount
'
'        For i = .FixedRows To .Rows - 1
'             .TextMatrix(i, .ColIndex("Ser")) = i
'             .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("ReignName").value), "", RsDetails("ReignName").value)
'               .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(RsDetails("Amount").value), "", RsDetails("Amount").value)
'
'            .TextMatrix(i, .ColIndex("Date")) = IIf(IsNull(RsDetails("ReignDate").value), "", RsDetails("ReignDate").value)
'               .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("Remrks").value), "", RsDetails("Remrks").value)
                   
                   
               '   .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("des").value), "", RsDetails("des").value)
'            RsDetails.MoveNext
'        Next i

'    End If
'End With
'    RsDetails.Close
'    Set RsDetails = Nothing
'''////////
        StrSQL = " select * from TblEvaluation "
StrSQL = StrSQL & "  Where ( dbo.TblEvaluation.Emp_ID = " & val(Me.XPTxtEmpID.text) & ")"


    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With VSFlexGrid3
     .Clear flexClearScrollable, flexClearEverything
     .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("EvalItem").value), "", RsDetails("EvalItem").value)
               .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsDetails("ActualDeg").value), "", RsDetails("ActualDeg").value)
                
            .TextMatrix(i, .ColIndex("MinValue")) = IIf(IsNull(RsDetails("MiniDeg").value), "", RsDetails("MiniDeg").value)
               .TextMatrix(i, .ColIndex("MaxValue")) = IIf(IsNull(RsDetails("MaxDeg").value), "", RsDetails("MaxDeg").value)
                   .TextMatrix(i, .ColIndex("Rate")) = IIf(IsNull(RsDetails("EvaluName").value), "", RsDetails("EvaluName").value)
                 .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDetails("Remrks").value), "", RsDetails("Remrks").value)
                ' .TextMatrix(i, .ColIndex("DES")) = IIf(IsNull(RsDetails("Des").value), "", RsDetails("Des").value)
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
  
  lbl(46).Caption = GetEmployeeSalaryAccordingToComponent(val(XPTxtEmpID.text), "", 0)
  
  
  
 ' lastHolidaydate.value = GETlASTiSSUEDATE(val(XPTxtEmpID.text))

 
     
     '    lastHolidaydateH.value = ToHijriDate(lastHolidaydate.value)
       
 

DCNationality_Change
  lbl(47).Caption = DateDiff("yyyy", Me.DBDOB.value, Date) + 1
    '''en salah
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
     Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    
 On Error GoTo ErrTrap
    XPTxtEmpName = Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text)

    If Me.TxtModFlg.text <> "R" Then

          If Text1.text = "" Then
          If SystemOptions.UserInterface = ArabicInterface Then
              Msg = "ÌÃ» «œŒ«· «”„ «·„ÊŸ› "
        Else
        Msg = "Enter Employee name"
        End If
              MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
              Text1.SetFocus
              SelectText Text1
              Exit Sub
             End If
    
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
    
        'If Not IsNumeric(TxtSalary.text) Then
        'If SystemOptions.UserInterface = ArabicInterface Then
        '         Msg = "ÌÃ» «œŒ«· «·—« » «·«”«”Ì ··„ÊŸ›  "
        '  Else
        '  Msg = " Enter Basic Salary Value  "
        '  End If
        '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '      TxtSalary.SetFocus
        '      SelectText TxtSalary
        '   Exit Sub
        'End If
        If DcboEmpDepartments.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡  ÕœÌœ «·«œ«—… «· Ì Ì »⁄Â« «·„ÊŸ›"
            Else
                Msg = " Specify Management"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '  CboWorkState.SetFocus
        
            DcboEmpDepartments.SetFocus
            Sendkeys "{F4}"
        
            Exit Sub
        End If
        
        
        
         If cboPayType.ListIndex = 2 And TxtBankCard.text = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "»—Ã«¡ ﬂ «»… —ﬁ„ «·’—«› "
                    Else
                        Msg = " Specify Card No "
                    End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
            Sendkeys "{F4}"
        
            Exit Sub
         End If
            
            
        If cboPayType.ListIndex = 3 And TxtBankCard.text = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "»—Ã«¡ ﬂ «»… —ﬁ„ «·Õ”«» «·»‰ﬂÌ "
                    Else
                        Msg = " Specify Card No "
                    End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
            Sendkeys "{F4}"
        
            Exit Sub
         End If
         
         
            
        
        
    
        If DcboJobsType.BoundText = "" Then
    
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡  ÕœÌœ ÊŸÌ›… «·„ÊŸ›"
            Else
                Msg = " Specify Job type"
            End If
        
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
            DcboJobsType.SetFocus
            Sendkeys "{F4}"
        
            Exit Sub
        End If
    
        If CboWorkState.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡  ÕœÌœ Õ«·… «·„ÊŸ› (Â· ⁄·Ï ﬁÊ… «·⁄„· √Ê  „ ›’·Â „‰ «·⁄„·)"
            Else
                Msg = " Specify Job Status"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '  CboWorkState.SetFocus
        
            dcjopstatus.SetFocus
            Sendkeys "{F4}"
        
            Exit Sub
        End If
    
        If val(Me.TxtEmp_Comm.text) > 0 Then
            If val(Me.TxtEmp_Comm.text) >= 100 Or val(Me.TxtEmp_Comm.text) < 0 Then
                Msg = "ﬁÌ„… ⁄„Ê·… «·„ÊŸ› €Ì— ’ÕÌÕ…..!!"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtEmp_Comm.SetFocus
                SelectText TxtEmp_Comm
                Exit Sub
            End If
        End If

         StrVacCode = IsRecExist("TblEmployee", "Fullcode", Me.DCPreFix.text & Trim(txtid.text), "Emp_Name", "Emp_ID<>" & val(XPTxtEmpID.text))

        If StrVacCode <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·ﬂÊœ „‰ ﬁ»·"
            Else
                Msg = " Emp Code Already Exist"
            End If
        
            MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
            '        TxtEmp_Code.SetFocus
         '   SelectText TxtEmp_Code
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
          If SystemOptions.UserInterface = ArabicInterface Then
                    
                    Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ  ··–„„  ...!!!"
        Else
           Msg = "Enter Opening Balnce   ...!!!"
           
        End If
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
                    
                   Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··«ÃÊ— «·„” Õﬁ…  ...!!!"
        Else
           Msg = "Enter  Salarie Opening Balnce    ...!!!"
           
        End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                    If TxtOpenBalance1.Enabled = True Then
                        TxtOpenBalance1.SetFocus
                    End If

                    Exit Sub
                End If
            End If
            
            If Me.OptType2(2).value = False Then
                If val(Me.TxtOpenBalance2.text) = 0 Then
                    
                                                If SystemOptions.UserInterface = ArabicInterface Then
                    
                                       Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··„Œ’’«  «·«Ã«“… ...!!!"

        Else
           Msg = "Enter  Allocation Opening Balnce    ...!!!"
           
        End If
        
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                    If TxtOpenBalance2.Enabled = True Then
                        TxtOpenBalance2.SetFocus
                    End If

                    Exit Sub
                End If
            End If
            
            
           If Me.OptType4(2).value = False Then
                   If val(Me.TxtOpenBalance4.text) = 0 Then
                                                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                                               Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··„Œ’’«  ‰Â«Ì… «·Œœ„… ...!!!"
        
                Else
                   Msg = "Enter  Allocation Opening Balnce    ...!!!"
                   
                End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                    If TxtOpenBalance4.Enabled = True Then
                        TxtOpenBalance4.SetFocus
                    End If

                    Exit Sub
                End If
            End If
            
            
                    If Me.OptType5(2).value = False Then
                   If val(Me.TxtOpenBalance5.text) = 0 Then
                                                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                                               Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··„Œ’’«  «· –«ﬂ—   ...!!!"
        
                Else
                   Msg = "Enter  Allocation Opening Balnce    ...!!!"
                   
                End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                    If TxtOpenBalance5.Enabled = True Then
                        TxtOpenBalance5.SetFocus
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
         '   Me.TxtEmp_Code.text = CStr(new_id("TblEmployee", "Emp_Code", "", True))
        
            rs.AddNew
            rs("Emp_ID").value = val(XPTxtEmpID.text)
             

            If detect_employee_work_type = 1 Then
   If Account_Code_dynamic <> "NO account" Then
                rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "- –„„ ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & "  ") '–„„
   End If
   
  If Account_Code_dynamic1 <> "NO account" Then
                rs("Account_Code1").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "- «ÃÊ— „” Õﬁ…", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & "- Salary  ") '–„„) '
                '«ÃÊ— „” Õﬁ…
 End If
        
        If Account_Code_dynamic3 <> "NO account" Then
                rs("Account_Code3").value = ModAccounts.AddNewAccount(Account_Code_dynamic3, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„œ›Ê⁄«  „ﬁœ„…  ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Adv. Payments ") '   „œ›Ê⁄«  „ﬁœ„Â
        End If
        
        If Account_Code_dynamic4 = "NO account" Then
        If Account_Code_dynamic2 <> "NO account" Then
             rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic2, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’«  ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Reserved ") '–„„) '„Œ’’« 
        End If
        
        Else
      If Account_Code_dynamic2 <> "NO account" Then
          rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic2, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’«  «Ã«“… ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Reserved Vacation ") '–„„) '„Œ’’« 
      End If
      If Account_Code_dynamic4 <> "NO account" Then
          rs("Account_Code4").value = ModAccounts.AddNewAccount(Account_Code_dynamic4, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’«   ‰Â«Ì… Œœ„…", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Reserved End Services ") '–„„) '„Œ’’« 
      End If
          End If
         
            End If
        If Account_Code_dynamic5 <> "NO account" Then
        rs("Account_Code5").value = ModAccounts.AddNewAccount(Account_Code_dynamic5, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’  –«ﬂ—    ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Adv. Payments ") '   „œ›Ê⁄«  „ﬁœ„Â
        End If
        
        Else
        
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             
             Cn.Execute "delete TblEmpDetails where Emp_ID=" & val(XPTxtEmpID.text)
             Cn.Execute "delete TblEvaluation where Emp_ID=" & val(XPTxtEmpID.text)
         '    Cn.Execute "delete TblEmpHolidaysDetails where Emp_ID=" & val(XPTxtEmpID.text)
             
            '  Rs("Account_Code").value = ModAccounts.AddNewAccount("a1a2a6", Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text), True, False)
        End If
    
  'nnnn   rs("BlnceVocat").value = val(Me.TxtBlncVoc.text)
  
  rs("To_Employee_name").value = txtBank(0).text
  
  
  rs("Commission").value = val(Txtcommission)
rs("WorkShop_Job").value = WorkShop_Job
     rs("MachinCode").value = TxtMachinCode(0).text
     rs("SalaryCode").value = TxtMachinCode(1).text
      rs("RegionID").value = IIf(DCRegionID.BoundText = "", Null, val(DCRegionID.BoundText))
     If ChNoAdded.value = vbChecked Then
     rs("NoAdded").value = 1
     Else
     rs("NoAdded").value = 0
     End If
    If OptSalaryType(0).value = True Then
            rs("SalaryType").value = 0
            rs("Percentage").value = Null
            rs("PerceTage").value = Null
            rs("BYHour").value = Null
   ElseIf OptSalaryType(1).value = True Then
          rs("SalaryType").value = 1
                   
            rs("Percentage").value = val(TxtPercentage.text)
            rs("PerceTage").value = val(TxtPercentage.text)
            
            rs("BYHour").value = Null
             
    ElseIf OptSalaryType(2).value = True Then
          rs("SalaryType").value = 2
                 rs("Percentage").value = Null
                 rs("PerceTage").value = Null
                 
            rs("BYHour").value = val(TxtBYHour.text)
            
    ElseIf OptSalaryType(3).value = True Then
           rs("SalaryType").value = 3
                       rs("Percentage").value = Null
                       rs("PerceTage").value = Null
            rs("BYHour").value = Null
         ElseIf OptSalaryType(4).value = True Then
           rs("SalaryType").value = 4
           rs("Percentage").value = Null
           rs("PerceTage").value = Null
           rs("BYHour").value = Null
    Else
         rs("SalaryType").value = Null
                     rs("Percentage").value = Null
                     rs("PerceTage").value = Null
            rs("BYHour").value = Null
    End If
    ' Me.cboPayType.ListIndex = IIf(IsNull(rs("PayType").value), 0, rs("PayType").value)
      rs("DateEndIndustrial").value = Format(txtDateEndIndustrial.value, "dd/mm/yyyy")
      rs("DateEndIndustrialHijri").value = txtDateEndIndustrialHijri.value
       If TypeEmp(1).value = True Then
       rs("TypeEmp").value = 1
       Else
       rs("TypeEmp").value = Null
       End If
        rs("DeptID2").value = val(Me.DcbDepartment2.BoundText)
        rs("HowIqamaEndH").value = HowIqamaEndH.value
        rs("PayType").value = val(Me.cboPayType.ListIndex)
        rs("Emp_Code").value = txtid.text
        rs("BankCard").value = TxtBankCard.text

            
        rs("InsuranceNO").value = TxtInsuranceNo.text
              
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtid.text) = "", Null, txtid.text)
        rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)
'''////////
rs("BankIAddress").value = txtBank(2).text
rs("BanckName").value = Me.DcbBanck(1).BoundText
rs("BankIBan").value = txtBank(1).text


    rs("PrefNatID").value = txtPrefNatID.text
''/////
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
        rs("EmpNotes").value = Trim(TxtEmpNotes.text)
    ''/// 28 10 2015
    rs("MaritalStatus").value = IIf(DcbMatrial.ListIndex = -1, Null, val(DcbMatrial.ListIndex))
    rs("BankCode").value = IIf(Me.DcbBanck(0).BoundText = "", Null, (DcbBanck(0).BoundText))
    rs("ContractID").value = IIf(DcbContractType.BoundText = "", Null, val(DcbContractType.BoundText))
    ''/////////

        If detect_employee_work_type = 1 Then
            If IsNull(rs("Account_Code").value) Or rs("Account_Code").value = "" Then
                      If Account_Code_dynamic <> "NO account" Then
                          rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "     ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " ") '   „œ›Ê⁄«  „ﬁœ„Â
                    End If
            Else
            
                ModAccounts.EditAccount rs("Account_Code").value, rs("Emp_Name").value, rs("Emp_Namee").value, , , , , , , , , , , , , , , , , True
            End If
            
            If IsNull(rs("Account_Code1").value) Or rs("Account_Code1").value = "" Then
                        If Account_Code_dynamic1 <> "NO account" Then
                            rs("Account_Code1").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "«ÃÊ— „” Õﬁ…    ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " salaries ") '   „œ›Ê⁄«  „ﬁœ„Â
                        End If
            Else
                ModAccounts.EditAccount rs("Account_Code1").value, rs("Emp_Name").value & "  «ÃÊ— „” Õﬁ… ", rs("Emp_Namee").value & " Salary ", , , , , , , , , , , , , , , , , True
            End If
            
            '555555555555555
            If Account_Code_dynamic5 <> "NO account" Then
                       If IsNull(rs("Account_Code5").value) Or rs("Account_Code5").value = "" Then
                                 If Account_Code_dynamic5 <> "NO account" Then
                                    rs("Account_Code5").value = ModAccounts.AddNewAccount(Account_Code_dynamic5, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’  –«ﬂ—      ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " salaries ") '   „œ›Ê⁄«  „ﬁœ„Â
                                End If
                      Else
                          ModAccounts.EditAccount rs("Account_Code5").value, rs("Emp_Name").value & "  „Œ’’  –«ﬂ—   ", rs("Emp_Namee").value & " Salary ", , , , , , , , , , , , , , , , , True
                      End If
            End If
            
            '55555555555555555
            If Account_Code_dynamic4 = "NO account" Then
                      If IsNull(rs("Account_Code2").value) Or rs("Account_Code2").value = "" Then
                                 If Account_Code_dynamic2 <> "NO account" Then
                                     rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic2, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’«      ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " Reserved ") '   „œ›Ê⁄«  „ﬁœ„Â
                                End If
                      Else
                      
                          ModAccounts.EditAccount rs("Account_Code2").value, rs("Emp_Name").value & "  „Œ’’« ", rs("Emp_Namee").value & "  Reserved ", , , , , , , , , , , , , , , , , True
                      End If
            
            Else
            
                     If IsNull(rs("Account_Code2").value) Or rs("Account_Code2").value = "" Then
                            If Account_Code_dynamic2 <> "NO account" Then
                                 rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic2, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’«  «Ã«“…", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & "Reserved vacation ") '   „œ›Ê⁄«  „ﬁœ„Â
                            End If
                      Else
                      
                          ModAccounts.EditAccount rs("Account_Code2").value, rs("Emp_Name").value & "  „Œ’’«  «Ã«“… ", rs("Emp_Namee").value & "  Reserved  vaction", , , , , , , , , , , , , , , , , True
                      End If
                      
            
            
                      If IsNull(rs("Account_Code4").value) Or rs("Account_Code4").value = "" Then
                                 If Account_Code_dynamic4 <> "NO account" Then
                                     rs("Account_Code4").value = ModAccounts.AddNewAccount(Account_Code_dynamic4, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„Œ’’«  ‰Â«Ì… Œœ„…", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & "Reserved End Services ") '   „œ›Ê⁄«  „ﬁœ„Â
                                End If
                      Else
                      
                          ModAccounts.EditAccount rs("Account_Code4").value, rs("Emp_Name").value & "  „Œ’’«    ‰Â«Ì… Œœ„… ", rs("Emp_Namee").value & "  Reserved  End Services", , , , , , , , , , , , , , , , , True
                      End If
                      
                      
            
            
            End If
            
            
            If IsNull(rs("Account_Code3").value) Or rs("Account_Code3").value = "" Then
                        If Account_Code_dynamic3 <> "NO account" Then
                            rs("Account_Code3").value = ModAccounts.AddNewAccount(Account_Code_dynamic3, Trim(Text1.text) & " " & Trim(Text2.text) & " " & Trim(Text3.text) & " " & Trim(Text4.text) & "„œ›Ê⁄«  „ﬁœ„…  ", True, False, Trim(Text5.text) & " " & Trim(Text6.text) & " " & Trim(Text7.text) & " " & Trim(Text8.text) & " -Adv. Payments ") '   „œ›Ê⁄«  „ﬁœ„Â
                        End If
            Else
                ModAccounts.EditAccount rs("Account_Code3").value, rs("Emp_Name").value & "  „œ›Ê⁄«  „ﬁœ„Â ", rs("Emp_Namee").value & "  Adv. Payments ", , , , , , , , , , , , , , , , , True
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
    
    
    
        If Me.OptType4(2).value = True Then
            rs("OpenBalance4").value = 0
            rs("OpenBalanceType4").value = Null
        ElseIf Me.OptType4(0).value = True Then
            rs("OpenBalance4").value = val(Me.TxtOpenBalance4.text)
            rs("OpenBalanceType4").value = 0
        ElseIf Me.OptType4(1).value = True Then
            rs("OpenBalance4").value = val(Me.TxtOpenBalance4.text)
            rs("OpenBalanceType4").value = 1
        End If
        
       If Me.OptType5(2).value = True Then
            rs("OpenBalance5").value = 0
            rs("OpenBalanceType5").value = Null
        ElseIf Me.OptType5(0).value = True Then
            rs("OpenBalance5").value = val(Me.TxtOpenBalance5.text)
            rs("OpenBalanceType5").value = 0
        ElseIf Me.OptType5(1).value = True Then
            rs("OpenBalance5").value = val(Me.TxtOpenBalance5.text)
            rs("OpenBalanceType5").value = 1
        End If
        
        
        rs("OpenBalanceDate").value = Me.Dtp.value
            
 

        '    Rs("Account_Code1").value = DcboCreditSide.BoundText
     
  '       rs("Password").value = Trim(tXTPassWord.Text)
         
        rs("hdodno").value = Trim(txthdodno.text)
    
        rs("hdoddate").value = txthdoddate.value
    
        rs("hdomnfaz").value = Trim(txthdomnfaz.text)
    
       ' rs("Emp_Salary").value = IIf(TxtSalary.text = "", Null, Trim(TxtSalary.text))
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
        rs("Emp_Phone").value = IIf(xptxtphone.text = "", "", Trim(xptxtphone.text))
        rs("Emp_mobile").value = IIf(XPTxtmobile.text = "", "", Trim(XPTxtmobile.text))
        rs("Emp_Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("Emp_Comm").value = IIf(TxtEmp_Comm.text = "", 0, val(TxtEmp_Comm.text))
        rs("EmpProfitCom").value = IIf(TxtEmpProfitCom.text = "", 0, val(TxtEmpProfitCom.text))
        rs("placeEkama").value = IIf(Txt_placEkama.text = "", Null, Trim(Txt_placEkama.text))
        rs("NumEkama").value = IIf(Txt_NumEkama.text = "", IIf(Tet_NumPoket.text = "", Null, Trim(Tet_NumPoket.text)), Trim(Txt_NumEkama.text))
         rs("DateExpoekamah").value = Txt_DateExpEkamaH.value
    
       rs("DateEndekamah").value = Txt_DateEndekamah.value
    
        rs("DateExpoekama").value = ToGregorianDate(Txt_DateExpEkamaH.value)
        rs("DateEndekama").value = ToGregorianDate(Txt_DateEndekamah.value)
    
        rs("DateExpLincH").value = Txt_DateExpLincH.value
        rs("DateEndLincH").value = Txt_DateEndLincH.value
    
    '*********************************
          rs("DriverLicenseStartdH").value = txtDriverLicenseStartdH.value
        rs("DriverLicenseendH").value = txtDriverLicenseendH.value
         rs("DriverLicenseend").value = ToGregorianDate(txtDriverLicenseendH.value)
    
    '**********************************
    
       rs("DOBH").value = DOBH.value
      '    rs("IssueDateH").value = IssueDateH.value
             rs("LastDateH").value = LastDateH.value
                 rs("LastDate").value = LastDate.value
                 
                rs("IssuingDriverCardDateH").value = DpIssuingDriverCardDateH.value
                rs("CardDriverExpireDateH").value = DpCardDriverExpireDateH.value
          
           '      rs("lastHolidaydateH").value = lastHolidaydateH.value
       '          rs("lastHolidaydate").value = lastHolidaydate.value
                 
                 
 
    'LastDate
    
 
       
       
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


'*********************
rs("DriverLicense").value = IIf(TxtDriverLicense.text = "", Null, Trim(TxtDriverLicense.text))

'************************
        rs("NumPasp").value = IIf(Txt_NumPasp.text = "", Null, Trim(Txt_NumPasp.text))
         rs("NumPaspOld").value = IIf(Txt_NumPaspOld.text = "", Null, Trim(Txt_NumPaspOld.text))
        
        rs("KafelID").value = IIf(txtKafelID.text = "", Null, Trim(txtKafelID.text))
        rs("KafelName").value = IIf(Me.DcbKafelName.text = "", Null, Trim(DcbKafelName.text))
     
        rs("kafeltel").value = IIf(TxtKafeltEL.text = "", Null, Trim(TxtKafeltEL.text))
        rs("kafeladd").value = IIf(txtkafeladd.text = "", Null, Trim(txtkafeladd.text))
       
        rs("pasplace").value = IIf(Me.DcbPasplace.BoundText = "", Null, Trim(Me.DcbPasplace.BoundText))
        rs("NationlID").value = IIf(DCNationality.BoundText = "", Null, val(DCNationality.BoundText))
        rs("Nationality").value = IIf(DCNationality.text = "", Null, Trim(DCNationality.text))
        rs("dean").value = IIf(dcdean.text = "", Null, Trim(dcdean.text))
        rs("DeanID").value = IIf(dcdean.BoundText = "", Null, val(dcdean.BoundText))
        rs("project_id").value = IIf(dcproject.text = "", Null, dcproject.BoundText)
        rs("BranchId").value = IIf(Me.DCBranch.text = "", Null, Me.DCBranch.BoundText)
   
           rs("VisaNo").value = IIf(TxtVisaNo.text = "", Null, Trim(TxtVisaNo.text))


        rs("DateExpPasp").value = Txt_DateExpPasp.value
        rs("DateEndPasp").value = Txt_DatePasp.value
        '    Rs("Notsstkala").Value = IIf(Txt_NotEndWork.text = "", "", Trim(Txt_NotEndWork.text))
      '  rs("Notsstkala").value = IIf(Txt_NotEndWork.text = "", "", Trim(Txt_NotEndWork.text))

        'If Chk_Stkala.value = Checked Then
        '    rs("ChekStkala").value = 1
       '' Else
       '     rs("ChekStkala").value = 0
       ' End If
    
    
        
        If chkShowTasks.value = Checked Then
            rs("chkShowTasks").value = 1
        Else
            rs("chkShowTasks").value = 0
        End If

    
        If chkStop.value = Checked Then
            rs("chkStop").value = 1
        Else
            rs("chkStop").value = 0
        End If
    
        If Chk_EndWork.value = Checked Then
            rs("ChekEndWork").value = 1
        Else
            rs("ChekEndWork").value = 0
        End If
    
       ' If Chk_Stkala.value = Checked Or Chk_EndWork.value = Checked Then
       '     rs("EndWork").value = DtDate.value
       ' Else
            rs("EndWork").value = Null
       ' End If
 
 '       rs("BignDateWork").value = DTPicker1.value
        rs("DOB").value = Format(DBDOB.value, "dd/mm/yyyy")

        rs("IssuingDriverCardDate").value = Format(DpIssuingDriverCardDate.value, "dd/mm/yyyy")
        rs("CardDriverExpireDate").value = Format(DpCardDriverExpireDate.value, "dd/mm/yyyy")
    

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
    
        If val(Me.DcGrade.BoundText) = 0 Then
            rs("gradeID").value = Null
        Else
            rs("gradeID").value = val(Me.DcGrade.BoundText)
        End If
    '
     If val(Me.Dcbsex.ListIndex) >= 0 Then
            rs("Sex").value = val(Me.Dcbsex.ListIndex) + 1
        
        End If
        If val(Me.DcbSection.BoundText) = 0 Then
            rs("SectionID").value = Null
        Else
            rs("SectionID").value = val(Me.DcbSection.BoundText)
        End If
        If val(Me.DCGroupID.BoundText) = 0 Then
            rs("GroupID").value = Null
        Else
            rs("GroupID").value = val(Me.DCGroupID.BoundText)
        End If
     '   rs("DateMoveNo").value = Me.DateMoveNo.value
        
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



        If val(Me.DcboJobsType1.BoundText) = 0 Then
            rs("JobTypeID1").value = Null
        Else
            rs("JobTypeID1").value = val(Me.DcboJobsType1.BoundText)
        End If
        
        
       If val(Me.mangerid.BoundText) = 0 Then
            rs("mangerid").value = Null
        Else
            rs("mangerid").value = val(Me.mangerid.BoundText)
        End If
        
        
               If val(Me.swapedempid.BoundText) = 0 Then
            rs("swapedempid").value = Null
        Else
            rs("swapedempid").value = val(Me.swapedempid.BoundText)
        End If
        
        
        
        
                
        
               If val(Me.swapedempid2.BoundText) = 0 Then
            rs("swapedempid2").value = Null
        Else
            rs("swapedempid2").value = val(Me.swapedempid2.BoundText)
        End If
        
        
        
        
        
        
        
        
                If val(Me.DcboJobsType2.BoundText) = 0 Then
            rs("JobTypeID2").value = Null
        Else
            rs("JobTypeID2").value = val(Me.DcboJobsType2.BoundText)
        End If
        
        
        
                If val(Me.DcboJobsType3.BoundText) = 0 Then
            rs("JobTypeID3").value = Null
        Else
            rs("JobTypeID3").value = val(Me.DcboJobsType3.BoundText)
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

        'rs("InsuranceValue").value = val(Me.TxtInsurValue.text)
       ' rs("OtherDiscounts").value = val(Me.TxtOtherDiscounts.text)
      '  rs("EmployeeInsurance").value = val(Me.txtEmployeeInsurance.text)
        
        'txtEmployeeInsurance
         rs("InstanceDateM").value = Me.DTPicker3.value
        rs("InstanceDateH").value = NourHijriCal1.value
        

        If Me.cmbInsuranceRenew.ListIndex = 0 Or Me.cmbInsuranceRenew.ListIndex = -1 Then
            rs("InsuranceRenew").value = 0
        ElseIf Me.cmbInsuranceRenew.ListIndex = 1 Then
            rs("InsuranceRenew").value = 1
        End If
        
        If Me.cmbToM.ListIndex = 0 Or Me.cmbToM.ListIndex = -1 Then
            rs("ToM").value = 0
        ElseIf Me.cmbToM.ListIndex = 1 Then
            rs("ToM").value = 1
        End If
        rs("InsuranceRenewDate").value = Me.txtInsuranceRenewDate.value
        rs("ToMDateNew").value = Me.txtToMDateNew.value
        rs("CopyNo").value = Trim$(Me.txtCopyNo.text)
             
        
         
        
        
    
        '    If Dir(system_path & "\"& SystemOptions.ImagesPath &"\" & XPTxtEmpID.text & ".JPG") <> "" Then
        '     Rs("ItemPhoto").value = DBPix201.ImageLoadFile(system_path & "\"& SystemOptions.ImagesPath &"\" & XPTxtEmpID.text & ".JPG")
 
        '    End If

        'OPENING Balance Voucher
       If detect_employee_work_type = 1 Then
    
            If val(TxtOpenBalance.text) <> 0 Or val(TxtOpenBalance1.text) <> 0 Or val(TxtOpenBalance2.text) <> 0 Or val(TxtOpenBalance4.text) <> 0 Or val(TxtOpenBalance5.text) <> 0 Then
                If val(rs("opening_balance_voucher_id").value & "") = 0 Then
                    txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
                Else
                    txtopening_balance_voucher_id.text = val(rs("opening_balance_voucher_id").value & "")
                End If
                rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
            Else
                rs("opening_balance_voucher_id").value = Null
            End If
    
        End If
        
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
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
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
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
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
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code1").value, val(Me.TxtOpenBalance1.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance1.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
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
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance1.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code1").value, val(Me.TxtOpenBalance1.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
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
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code2").value, val(Me.TxtOpenBalance2.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance2.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
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
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance2.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code2").value, val(Me.TxtOpenBalance2.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
                 
            End If
        End If
   '''''''''444444444444444444444444
           '33333333333333333333
           If Account_Code_dynamic4 <> "NO account" Then
        If Me.OptType4(0).value = True Or Me.OptType4(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
                If Me.OptType4(0).value = True Then
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
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code4").value, val(Me.TxtOpenBalance4.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance4.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                ElseIf Me.OptType4(1).value = True Then
            
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
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance4.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code4").value, val(Me.TxtOpenBalance4.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
                 
            End If
        End If
        End If
   
        
  '555555555555555555555555555555555555555555555555555555
  '''''''''444444444444444444444444
           '33333333333333333333
           If Account_Code_dynamic5 <> "NO account" Then
        If Me.OptType5(0).value = True Or Me.OptType5(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
                If Me.OptType5(0).value = True Then
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
        
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code5").value, val(Me.TxtOpenBalance5.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance5.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                ElseIf Me.OptType5(1).value = True Then
            
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
                 
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance5.text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code5").value, val(Me.TxtOpenBalance5.text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(DCBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
                 
            End If
        End If
        End If
  
  '5555555555555555555555555555555555555555555
        rs.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

    End If



'»Ì‰«  ÃœÌœ…

 Dim RsDetails As ADODB.Recordset
 
 

        Set RsDetails = New ADODB.Recordset
      '  RsDetails.Open "TblEmpDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
           StrSQL = "SELECT     * from dbo.TblEmpDetails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
With Grid
                    For i = .FixedRows To .rows - 1
                    If .TextMatrix(i, .ColIndex("name")) <> "" Then
                        RsDetails.AddNew
                             RsDetails("Emp_ID").value = val(XPTxtEmpID.text)
                     RsDetails("name").value = .TextMatrix(i, .ColIndex("name"))
                     
                       RsDetails("passportno").value = .TextMatrix(i, .ColIndex("passportno"))
                          RsDetails("relationtype").value = val(.TextMatrix(i, .ColIndex("relationtype")))
                          
                        RsDetails("iqamano").value = .TextMatrix(i, .ColIndex("iqamano"))
                        RsDetails("haveinsurance").value = .TextMatrix(i, .ColIndex("haveinsurance"))
                             RsDetails("insuranceno").value = .TextMatrix(i, .ColIndex("insuranceno"))
                             RsDetails("des").value = .TextMatrix(i, .ColIndex("des"))
                               RsDetails("OprType").value = 0
                        RsDetails.update
                        End If
                    Next i
 End With
RsDetails.Close
Set RsDetails = Nothing



        Set RsDetails = New ADODB.Recordset
     '   RsDetails.Open "TblEmpDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
           StrSQL = "SELECT     * from dbo.TblEmpDetails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
With grid01
                    For i = .FixedRows To .rows - 1
                    If .TextMatrix(i, .ColIndex("name")) <> "" Then
                        RsDetails.AddNew
                             RsDetails("Emp_ID").value = val(XPTxtEmpID.text)
                     RsDetails("name").value = .TextMatrix(i, .ColIndex("name"))
                     
                       RsDetails("yearofqualication").value = .TextMatrix(i, .ColIndex("yearofqualication"))
                          RsDetails("qualicationEntity").value = (.TextMatrix(i, .ColIndex("qualicationEntity")))
                          
                        RsDetails("grade").value = .TextMatrix(i, .ColIndex("grade"))
                       
                             RsDetails("des").value = .TextMatrix(i, .ColIndex("des"))
                               RsDetails("OprType").value = 1
                        RsDetails.update
                        End If
                    Next i
 End With
RsDetails.Close
Set RsDetails = Nothing



        Set RsDetails = New ADODB.Recordset
        'RsDetails.Open "TblEmpDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            StrSQL = "SELECT     * from dbo.TblEmpDetails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    

With grid02
                    For i = .FixedRows To .rows - 1
                    If .TextMatrix(i, .ColIndex("name")) <> "" Then
                        RsDetails.AddNew
                             RsDetails("Emp_ID").value = val(XPTxtEmpID.text)
                     RsDetails("name").value = .TextMatrix(i, .ColIndex("name"))
                     
                       
                          RsDetails("qualicationEntity").value = (.TextMatrix(i, .ColIndex("qualicationEntity")))
                          
                      RsDetails("workfrom").value = .TextMatrix(i, .ColIndex("workfrom"))
                      RsDetails("workfromH").value = .TextMatrix(i, .ColIndex("workfromH"))
                      
                      RsDetails("workto").value = .TextMatrix(i, .ColIndex("workto"))
                      RsDetails("worktoH").value = .TextMatrix(i, .ColIndex("worktoH"))
                      
                       
                       
                             RsDetails("des").value = .TextMatrix(i, .ColIndex("des"))
                               RsDetails("OprType").value = 2
                        RsDetails.update
                        End If
                    Next i
 End With
RsDetails.Close
Set RsDetails = Nothing






'»Ì‰«  ÃœÌœ…

  
 
 

    '    Set RsDetails = New ADODB.Recordset
        'RsDetails.Open "TblEmpHolidaysDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    '        StrSQL = "SELECT     * from dbo.TblEmpHolidaysDetails Where (1 = -1)"
   'RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
'With Grid20

'                    For i = .FixedRows To .Rows - 1
'                    If .TextMatrix(i, .ColIndex("fromdate")) <> "" Then
'                        RsDetails.AddNew
'                             RsDetails("Emp_ID").value = val(XPTxtEmpID.text)
'                     RsDetails("fromdate").value = .TextMatrix(i, .ColIndex("fromdate"))
'
'                       RsDetails("todate").value = .TextMatrix(i, .ColIndex("todate"))
'
'                       RsDetails("fromdateh").value = .TextMatrix(i, .ColIndex("fromdateh"))
'
'                       RsDetails("todateh").value = .TextMatrix(i, .ColIndex("todateh"))
'
'                             RsDetails("des").value = .TextMatrix(i, .ColIndex("des"))
'                       RsDetails("cheke").value = 0
'                       RsDetails("day").value = .TextMatrix(i, .ColIndex("day"))
'                       RsDetails("month").value = .TextMatrix(i, .ColIndex("month"))
'                       RsDetails("year").value = .TextMatrix(i, .ColIndex("year"))
'
'                        RsDetails.update
'                        End If
'                    Next i
' End With
'RsDetails.Close
'Set RsDetails = Nothing
''''/////////salahstr
' Dim RsDetails As ADODB.Recordset
 
 

        Set RsDetails = New ADODB.Recordset
   '     RsDetails.Open "TblHealthy", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
             StrSQL = "SELECT     * from dbo.TblHealthy Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
With Me.VSFlexGrid2
                    For i = .FixedRows To .rows - 1
                    If .TextMatrix(i, .ColIndex("ENTITY")) <> "" Then
                        RsDetails.AddNew
                             RsDetails("Emp_ID").value = val(XPTxtEmpID.text)
                           '  RsDetails("HealthyDate").value = .TextMatrix(i, .ColIndex("Date"))
                             RsDetails("HealthyDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("Date"))), .TextMatrix(i, .ColIndex("Date")), Null)
                     RsDetails("FromH").value = .TextMatrix(i, .ColIndex("ENTITY"))
                     
                       RsDetails("Clinic").value = .TextMatrix(i, .ColIndex("clinic"))
                          RsDetails("Remrks").value = .TextMatrix(i, .ColIndex("Remarks"))
                          
                        RsDetails("CompanyAmount").value = .TextMatrix(i, .ColIndex("valuecomp"))
                        RsDetails("HeathyTreat").value = .TextMatrix(i, .ColIndex("Typename"))
                             RsDetails("Des").value = .TextMatrix(i, .ColIndex("DES"))
                         
                        RsDetails.update
                        End If
                    Next i
 End With
RsDetails.Close
Set RsDetails = Nothing


 'Dim RsDetails As ADODB.Recordset
 
 

        Set RsDetails = New ADODB.Recordset
      '  RsDetails.Open "TblEvaluation", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
             StrSQL = "SELECT     * from dbo.TblEvaluation Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
With VSFlexGrid3
                    For i = .FixedRows To .rows - 1
                    If .TextMatrix(i, .ColIndex("name")) <> "" Then
                        RsDetails.AddNew
                             RsDetails("Emp_ID").value = val(XPTxtEmpID.text)
                             RsDetails("EvalItem").value = .TextMatrix(i, .ColIndex("name"))
                     RsDetails("ActualDeg").value = .TextMatrix(i, .ColIndex("Value"))
                     
                       RsDetails("MaxDeg").value = .TextMatrix(i, .ColIndex("MaxValue"))
                          RsDetails("MiniDeg").value = .TextMatrix(i, .ColIndex("MinValue"))
                          
                        RsDetails("EvaluName").value = .TextMatrix(i, .ColIndex("Rate"))
                        RsDetails("Remrks").value = .TextMatrix(i, .ColIndex("des"))
                       
                        RsDetails.update
                        End If
                    Next i
 End With
RsDetails.Close
Set RsDetails = Nothing

 ' Set RsDetails = New ADODB.Recordset
 '       RsDetails.Open "TblEmpReign", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
 
'With VSFlexGrid4
'                    For i = .FixedRows To .Rows - 1
'                    If .TextMatrix(i, .ColIndex("name")) <> "" Then
''                        RsDetails.AddNew
 '                            RsDetails("Emp_ID").value = val(XPTxtEmpID.text)
 ''                            RsDetails("ReignName").value = .TextMatrix(i, .ColIndex("name"))
  '                   RsDetails("Amount").value = .TextMatrix(i, .ColIndex("Qty"))
  '                  RsDetails("ReignDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("Date"))), .TextMatrix(i, .ColIndex("Date")), Null)
                       'RsDetails("ReignDate").value = .TextMatrix(i, .ColIndex("Date"))
  '                        RsDetails("Remrks").value = .TextMatrix(i, .ColIndex("des"))
                          
                       ' RsDetails("EvaluName").value = .TextMatrix(i, .ColIndex("Rate"))
                       ' RsDetails("Remrks").value = .TextMatrix(i, .ColIndex("des"))
                       
  '                      RsDetails.update
  '                      End If
  '                  Next i
 'End With
'RsDetails.Close
'Set RsDetails = Nothing
'''//salahend

    ' ⁄œÌ· „—ﬂ“ «· ﬂ·›…

    If Me.DcCostCenter.BoundText = "" Then
        Dim X As Boolean
        X = UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code"), False, 0, "")
        X = UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code1"), False, 0, "")
        X = UPDATE_ACCOUNT_COST_CENTER(get_EMPLOYEE_Account(val(XPTxtEmpID.text), "Account_Code2"), False, 0, "")
   
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
   ' updateEmployeeSalaryComponent val(Me.XPTxtEmpID.text), Me.TxtSalary.text
    ' ÕœÌÀ «·„›—œ«  «·«·Ì…
    'addSalaryComponentToEmployee val(Me.XPTxtEmpID.text)
 
    Select Case Me.TxtModFlg.text

        Case "N"

            'updateResults
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„ÊŸ› " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
      
            Else
                Msg = " This Employee Data Was Saved" & CHR(13)
                Msg = Msg + "Do you want To enter Another Employee"
            End If
  
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub

            End If
        
        Case "E"

                MsgBox "Amendments have been saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title


            '  updateResults
 
    End Select
  '  Fra(15).Visible = False
    
 
    
    'rs.Close
    
'    StrSQL = "select * from  TblEmployee order by fullcode"
 
 
 
    StrSQL = "select * from  TblEmployee where 1=1"
       If WorkShop_Job = 0 Then
       
        Else
        StrSQL = StrSQL & " and  WorkShop_Job=" & WorkShop_Job  '& " order by fullcode"
        End If
      
   
        If SystemOptions.usertype <> UserAdminAll Then
     '       StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
            
        End If
        
     '   StrSQL = StrSQL & " order by fullcode "
        
        StrSQL = StrSQL & "  AND  (BranchId=0 or BranchId is null or         BranchId in(" & Current_branchSql & "))"
            
        StrSQL = StrSQL & " order by fullcode "
        
'  If WorkShop_Job = 0 Then
'        StrSQL = "select * from  TblEmployee order by fullcode"
'        Else
'        StrSQL = "select * from  TblEmployee WHERE WorkShop_Job=" & WorkShop_Job & " order by fullcode"
'        End If
     
     

  '  rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
'    rs.Open "[TblEmployee]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    TxtModFlg.text = "R"
    Me.Retrive val(Me.XPTxtEmpID)
       
   
    
    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        Cn.RollbackTrans
        BeginTrans = False
    End If

    'Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
     
    
        Msg = "Sorry........ Error During Saving " & CHR(13)
   
 
 
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
            rs.Find "Emp_ID='" & val(XPTxtEmpID.text) & "'", , adSearchForward, adBookmarkFirst

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

        If mId(src, i, 1) = "+" Or mId(src, i, 1) = "-" Or mId(src, i, 1) = "*" Or mId(src, i, 1) = "/" Or mId(src, i, 1) = "=" Then
            new_pos = i
            cuttent_operand = mId(src, last_pos, new_pos - last_pos)

            If InStr(cuttent_operand, "A") > 0 Then
                cuttent_operand = get_value(cuttent_operand)
            End If

            new_str = new_str & cuttent_operand & mId(src, i, 1)

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
'    Cmd_Click (1)
'    OptType(2).value = True
'    TxtOpenBalance.text = 0

    'OptType1(2).value = True
    'TxtOpenBalance1.text = 0

    'OptType2(2).value = True
    'TxtOpenBalance2.text = 0

'    Cmd_Click (2)

End Function

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ «·„ÊŸ›    " & txtid.text & CHR(13) & "   «·«”„ " & XPLbl(60).Caption & CHR(13)
                     
    LogTexte = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " Code      " & txtid.text & CHR(13) & "   Name " & XPLbl(60).Caption & CHR(13)
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub Del_ProfData()

    Dim Msg As String
    Dim StrSQL As String

    'On Error GoTo ErrTrap
       If check_employee_transations(val(XPTxtEmpID)) = False Then

        Exit Sub

    End If
    
    DeleteOpeningBalance
    StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
             
 

    If XPTxtEmpID.text <> "" Then

        Msg = "”Ì „ Õ–› »Ì«‰«  «·„ÊŸ› —ﬁ„ " & CHR(13)
'        Msg = Msg + Me.DCPreFix.text & Trim(TxtEmp_Code.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
             CuurentLogdata ("D")
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„ÊŸ› "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Function check_employee_transations(Emp_id As String) As Boolean
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    check_employee_transations = True
 
    StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code") & "' or  Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code1") & "' or Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code2") & "' or  Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code3") & "'"
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & CHR(13)
        Msg = Msg + "·«‰… „”Ã· ›Ì »⁄÷ «·ﬁÌÊœ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        check_employee_transations = False
        Exit Function
    End If
    
    RsTemp.Close
    StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS1 where Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code") & "' or  Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code1") & "' or Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code2") & "' or  Account_Code='" & get_EMPLOYEE_Account(Emp_id, "Account_Code3") & "'"
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & CHR(13)
        Msg = Msg + "·«‰… „”Ã· ›Ì   «·«—’œ… «·«›  «ÕÌ… "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

    End If
    
       If ModAccounts.DeleteAccount(get_EMPLOYEE_Account(Emp_id, "Account_Code5")) = True Then
        check_employee_transations = True

    End If
    
    
    
    
    
    
    
    
    
    RsTemp.Close
    
    
    StrSQL = " select Emp_id  FROM         dbo.Transactions Where (Emp_id = " & Emp_id & ")"

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & CHR(13)
        Msg = Msg + "·«‰… „”Ã· ›Ì »⁄÷ «·Õ—ﬂ«  «· Ã«—Ì… "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        check_employee_transations = False
        Exit Function
    End If
    
    
    RsTemp.Close
    
    
    StrSQL = " SELECT     EmpId FROM         dbo.Notes WHERE     (EmpId = " & Emp_id & ") "

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & CHR(13)
        Msg = Msg + "·«‰… „”Ã· ›Ì »⁄÷ «·Õ—ﬂ«  «·„«·Ì… "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        check_employee_transations = False
        Exit Function
    End If
    
    
    
   RsTemp.Close
    
    
    StrSQL = " SELECT     EmpID FROM         dbo.TBLSalesRepData    WHERE     (EmpId = " & Emp_id & ") "

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & CHR(13)
        Msg = Msg + "·«‰… „— »ÿ »ÃœÊ· «·„‰«œÌ» "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        check_employee_transations = False
        Exit Function
    End If
        
        
    RsTemp.Close
    
    StrSQL = " SELECT     Emp_id  FROM         dbo.Contract    WHERE     (Emp_id = " & Emp_id & ") "

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & CHR(13)
        Msg = Msg + "·«‰… „— »ÿ »ÃœÊ· ⁄ﬁÊœ «·„ÊŸ›Ì‰ "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        check_employee_transations = False
        Exit Function
    End If
    
    
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
'            Cmd_Click (0)
        Else
'            KeyCode = 0
'            SendKeys "{TAB}"
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
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „ÊŸ› ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·„ÊŸ›" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„ÊŸ› «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  „ÊŸ›" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „ÊŸ›" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap, True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New Record ..." & Wrap & "Click here to add a new employee" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the current employee data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the new record or " & Wrap & "save the edit in the " & Wrap & "current record", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo" & Wrap & "Undo in the adding new record" & Wrap & "Or undo in the current editing" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current employee data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Search for an employee" & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist Record" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next" & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last" & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Employees Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help" & Wrap & "Show the Help File" & Wrap & "" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub
Function printingReport(Optional NoteSerial As String)
On Error Resume Next
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  If C1Tab1.CurrTab = 0 Or C1Tab1.CurrTab = 1 Or C1Tab1.CurrTab = 6 Or C1Tab1.CurrTab = 7 Or C1Tab1.CurrTab = 9 Then
 MySQL = "SELECT     dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, "
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.BlnceVocat, dbo.TblEmployee.InstanceDateH, dbo.TblEmployee.InstanceDateM,"
MySQL = MySQL & "                       dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job, dbo.TblEmployee.BYHour, dbo.TblEmployee.Percentage, dbo.TblEmployee.SalaryType,"
MySQL = MySQL & "                       dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseend, dbo.TblEmployee.DriverLicense,"
MySQL = MySQL & "                       dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.OpenBalance4, dbo.TblEmployee.OpenBalanceType4,"
MySQL = MySQL & "                       dbo.TblEmployee.swapedempid, dbo.TblEmployee.mangerid, dbo.TblEmployee.VisaNo, dbo.TblEmployee.LastDateH, dbo.TblEmployee.JobTypeID3,"
MySQL = MySQL & "                       TblEmpJobsTypes_3.JobTypeName AS JobTypeName3, TblEmpJobsTypes_3.JobTypeNamee AS JobTypeNamee3, dbo.TblEmployee.JobTypeID2,"
MySQL = MySQL & "                       TblEmpJobsTypes_2.JobTypeName AS JobTypeName2, TblEmpJobsTypes_2.JobTypeNamee AS JobTypeNamee2, dbo.TblEmployee.JobTypeID1,"
MySQL = MySQL & "                       TblEmpJobsTypes_1.JobTypeName AS JobTypeName1, TblEmpJobsTypes_1.JobTypeNamee AS JobTypeNamee1, dbo.TblEmployee.LastDate,"
MySQL = MySQL & "                      dbo.TblEmployee.IssueDateH, dbo.TblEmployee.DOBH, dbo.TblEmployee.gradeID, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.BankCard,"
MySQL = MySQL & "                       dbo.TblEmployee.DriverId, dbo.TblEmployee.Account_Code5, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.Account_Code3, dbo.TblEmployee.OpenBalance2,"
MySQL = MySQL & "                       dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance1, dbo.TblEmployee.OpenBalanceType1, dbo.TblEmployee.OpenBalance,"
MySQL = MySQL & "                       dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id, dbo.TblEmployee.prifix,"
MySQL = MySQL & "                       dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.term_fullcode, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                       dbo.TblBranchesData.branch_namee, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_id, dbo.TblEmployee.opr_fullcode, dbo.TblEmployee.dateendpoketh,"
MySQL = MySQL & "                       dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.project_id, dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.Account_code1,"
MySQL = MySQL & "                       dbo.TblEmployee.Account_code, dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_others1,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_mang,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_sakn,"
MySQL = MySQL & "                       dbo.TblEmployee.kafeladd, dbo.TblEmployee.kafeltel, dbo.TblEmployee.DOB, dbo.TblEmployee.Notsstkala, dbo.TblEmployee.EndWork,"
MySQL = MySQL & "                       dbo.TblEmployee.BignDateWork, dbo.TblEmployee.CustNum, dbo.TblEmployee.EmpNum, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.Dateexppoket,"
MySQL = MySQL & "                       dbo.TblEmployee.NumPoket, dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLinc,"
MySQL = MySQL & "                       dbo.TblEmployee.NumLicn, dbo.TblEmployee.placeEkama, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.InsuranceState,"
MySQL = MySQL & "                       dbo.TblEmployee.Region, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.pasplace, dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.DateEndekama,"
MySQL = MySQL & "                       dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_Mail,"
MySQL = MySQL & "                       dbo.TblEmployee.workstate, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.KafelName, dbo.TblEmployee.hdoddate,"
MySQL = MySQL & "                       dbo.TblEmployee.hdodno, dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.NumPasp, dbo.TblEmployee.KafelID,"
MySQL = MySQL & "                       dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.NumEkama, dbo.TblEmployee.Emp_Salary_others,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.placeWORK, dbo.TblEmployee.JobTypeID, TblEmpJobsTypes_1.JobTypeName,"
MySQL = MySQL & "                       TblEmpJobsTypes_1.JobTypeNamee, dbo.TblEmployee.dean, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_ID, dbo.EmpSalaryComponent.AccountCode,"
MySQL = MySQL & "                       dbo.EmpSalaryComponent.AccountName, dbo.EmpSalaryComponent.[Value], dbo.EmpSalaryComponent.des, dbo.EmpSalaryComponent.eq_text,"
MySQL = MySQL & "                       dbo.EmpSalaryComponent.specific_value, dbo.EmpSalaryComponent.percentage AS percentageComp, dbo.EmpSalaryComponent.min_val,"
MySQL = MySQL & "                       dbo.EmpSalaryComponent.max_val, dbo.EmpSalaryComponent.is_fixed, dbo.EmpSalaryComponent.mofrad_type, dbo.EmpSalaryComponent.ModDate,"
MySQL = MySQL & "                       dbo.EmpSalaryComponent.Flagx, dbo.EmpSalaryComponent.EntIncresDataM, dbo.EmpSalaryComponent.EntIncresDataH, dbo.mofrdat.mofrad_name,"
MySQL = MySQL & "                       dbo.mofrdat.specific_value AS specific_valueM, dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName,"
MySQL = MySQL & "                       dbo.TblEmpDepartments.DepartmentNamee , dbo.TblEmployee.GroupID, dbo.EmpGroupDep.GroupName, dbo.EmpGroupDep.Ename , dbo.TblEmployee.Emp_Remark, "
MySQL = MySQL & "                      dbo.TblEmployee.EmpNotes "
MySQL = MySQL & "  FROM         dbo.TblEmpJobsTypes TblEmpJobsTypes_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmpJobsTypes TblEmpJobsTypes_4 ON dbo.TblEmployee.JobTypeID = TblEmpJobsTypes_4.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id ON"
MySQL = MySQL & "                       TblEmpJobsTypes_1.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmpJobsTypes TblEmpJobsTypes_2 ON dbo.TblEmployee.JobTypeID2 = TblEmpJobsTypes_2.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmpJobsTypes TblEmpJobsTypes_3 ON dbo.TblEmployee.JobTypeID3 = TblEmpJobsTypes_3.JobTypeID FULL OUTER JOIN"
MySQL = MySQL & "                       dbo.mofrdat RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode ON"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_id = dbo.EmpSalaryComponent.Emp_id"
MySQL = MySQL & " Where (dbo.TblEmployee.Emp_id = " & val(XPTxtEmpID.text) & ")"

' MySQL = MySQL & "    Where (dbo.TblQuesEmp.id = " & val(XPTxtID.text) & ")"
 If C1Tab1.CurrTab = 0 Or C1Tab1.CurrTab = 1 Then
 If val(DCNationality.BoundText) = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmpSaudi.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmpSaudi.rpt"
        End If
        Else
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp.rpt"
        End If
        End If
        End If
     If C1Tab1.CurrTab = 6 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp Instance.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp Instance.rpt"
        End If
        End If
          If C1Tab1.CurrTab = 9 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp Account.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp Account.rpt"
        End If
        End If
       If C1Tab1.CurrTab = 7 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmpvisa.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmpvisa.rpt"
        End If
        End If
        
Else

 '''////////////////////////
MySQL = " SELECT     dbo.TblEmpDetails.passportno, dbo.TblEmpDetails.iqamano, dbo.TblEmpDetails.name, dbo.TblEmpDetails.haveinsurance, dbo.TblEmpDetails.insuranceno, "
    MySQL = MySQL & "                   dbo.TblEmpDetails.yearofqualication, dbo.TblEmpDetails.qualicationEntity, dbo.TblEmpDetails.grade, dbo.TblEmpDetails.workfrom, dbo.TblEmpDetails.workto,"
     MySQL = MySQL & "                  dbo.TblEmpDetails.workfromH, dbo.TblEmpDetails.worktoH, dbo.TblEmpDetails.des, dbo.TblEmpDetails.illdate, dbo.TblEmpDetails.illdateH,"
    MySQL = MySQL & "                   dbo.TblEmpDetails.OprType, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
   MySQL = MySQL & "                    dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.JobTypeID,"
   MySQL = MySQL & "                    dbo.TblEmployee.placeWORK, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.NumEkama,"
   MySQL = MySQL & "                    dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.KafelID, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
    MySQL = MySQL & "                   dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate, dbo.TblEmployee.KafelName, dbo.TblEmployee.hdomnfaz,"
     MySQL = MySQL & "                  dbo.TblEmployee.jopstatusid, dbo.TblEmployee.workstate, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.DateEndekama,"
    MySQL = MySQL & "                   dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.pasplace, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region, dbo.TblEmployee.InsuranceValue,"
   MySQL = MySQL & "                    dbo.TblEmployee.InsuranceState, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc,"
   MySQL = MySQL & "                    dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
   MySQL = MySQL & "                    dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmpDetails.relationtype,"
  MySQL = MySQL & "                     dbo.TblEmpDetails.Emp_ID, dbo.TblRelations.name AS namerelation, dbo.TblRelations.namee"
 MySQL = MySQL & " FROM         dbo.TblEmpDetails LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblRelations ON dbo.TblEmpDetails.relationtype = dbo.TblRelations.ID RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblEmpDetails.Emp_ID = dbo.TblEmployee.Emp_ID"
 MySQL = MySQL & ""
 If C1Tab1.CurrTab = 2 Then
 MySQL = MySQL & "  Where (dbo.TblEmpDetails.OprType = 0) And (dbo.TblEmployee.Emp_id = " & val(XPTxtEmpID.text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpFamily.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpFamily.rpt"
        End If
End If
 If C1Tab1.CurrTab = 3 Then
 MySQL = MySQL & "  Where (dbo.TblEmpDetails.OprType = 1 or dbo.TblEmpDetails.OprType=2) And (dbo.TblEmployee.Emp_id = " & val(XPTxtEmpID.text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpDecExp.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpDecExp.rpt"
        End If
End If

End If
If C1Tab1.CurrTab = 8 Then
MySQL = "SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, "
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.WorkShop_Job,"
MySQL = MySQL & "                      dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Account_Code3, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.SalaryType, dbo.TblEmployee.Percentage,"
MySQL = MySQL & "                      dbo.TblEmployee.BYHour, dbo.TblEmployee.BankCard, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType1, dbo.TblEmployee.OpenBalance1,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.OpenBalance4,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id, dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate,"
MySQL = MySQL & "                      dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.placeWORK, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_sakn1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama, dbo.TblEmployee.Nationality,"
MySQL = MySQL & "                      dbo.TblEmployee.dean, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Salary_others,"
MySQL = MySQL & "                      dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.KafelID, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.KafelName, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.workstate, dbo.TblEmployee.DateEndekama,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.pasplace, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region, dbo.TblEmployee.InsuranceValue,"
MySQL = MySQL & "                      dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH,"
MySQL = MySQL & "                      dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.EmpNum,"
MySQL = MySQL & "                      dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork,"
MySQL = MySQL & "                      dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB, dbo.TblEmployee.kafeltel, dbo.TblEmployee.kafeladd,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc, dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc, dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1, dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1,"
MySQL = MySQL & "                      dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.project_id, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
MySQL = MySQL & "                      dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.BranchId,"
MySQL = MySQL & "                      dbo.TblEmployee.Account_Code5, dbo.TblEmployee.DriverId, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.gradeID, dbo.TblEmployee.DOBH,"
MySQL = MySQL & "                      dbo.TblEmployee.IssueDateH, dbo.TblEmployee.LastDate, dbo.TblEmployee.LastDateH, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.JobTypeID2,"
MySQL = MySQL & "                      dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.VisaNo, dbo.TblEmployee.GroupID, dbo.TblEmployee.mangerid, dbo.TblEmployee.swapedempid,"
MySQL = MySQL & "                      dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.DriverLicense, dbo.TblEmployee.DriverLicenseend,"
MySQL = MySQL & "                      dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.PerceTage, dbo.TblEmpJobsTypes.JobTypeName,"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes.JobTypeNamee, TblEmpJobsTypes_1.JobTypeName AS jobnmae3, TblEmpJobsTypes_1.JobTypeNamee AS jobnamee3,"
MySQL = MySQL & "                      TblEmpJobsTypes_2.JobTypeName AS jobnmae1, TblEmpJobsTypes_2.JobTypeNamee AS jobnmaee1, TblEmpJobsTypes_3.JobTypeName AS jobnmae2,"
MySQL = MySQL & "                      TblEmpJobsTypes_3.JobTypeNamee AS jobnmaee2, dbo.TblVocationEntitlements.stratDate, dbo.TblVocationEntitlements.stratDateH,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.AcuDate, dbo.TblVocationEntitlements.AcuDateH, dbo.TblVocationEntitlements.EndDateH, dbo.TblVocationEntitlements.EndDate,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.NoDayAct, dbo.TblVocationEntitlements.NoDayDelay, dbo.TblVocationEntitlements.NoVacation,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.remark"
MySQL = MySQL & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblVocationEntitlements ON dbo.TblEmployee.Emp_ID = dbo.TblVocationEntitlements.EmpID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_3 ON dbo.TblEmployee.JobTypeID2 = TblEmpJobsTypes_3.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_2 ON dbo.TblEmployee.JobTypeID1 = TblEmpJobsTypes_2.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_1 ON dbo.TblEmployee.JobTypeID3 = TblEmpJobsTypes_1.JobTypeID"
MySQL = MySQL & " Where (dbo.TblEmployee.Emp_id = " & val(XPTxtEmpID.text) & ")"
' MySQL = MySQL & "  Where (dbo.TblEmpDetails.OprType = 1 or dbo.TblEmpDetails.OprType=2) And (dbo.TblEmployee.Emp_id = " & val(XPTxtEmpID.text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpHolday.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpHolday.rpt"
        End If
End If
''
If C1Tab1.CurrTab = 4 Then
MySQL = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.WorkShop_Job,"
MySQL = MySQL & "                      dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Account_Code3, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.SalaryType, dbo.TblEmployee.Percentage,"
MySQL = MySQL & "                      dbo.TblEmployee.BYHour, dbo.TblEmployee.BankCard, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType1, dbo.TblEmployee.OpenBalance1,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.OpenBalance4,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id, dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate,"
MySQL = MySQL & "                      dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.placeWORK, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_sakn1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama, dbo.TblEmployee.Nationality,"
MySQL = MySQL & "                      dbo.TblEmployee.dean, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Salary_others,"
MySQL = MySQL & "                      dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.KafelID, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.KafelName, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.workstate, dbo.TblEmployee.DateEndekama,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.pasplace, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region, dbo.TblEmployee.InsuranceValue,"
MySQL = MySQL & "                      dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH,"
MySQL = MySQL & "                      dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.EmpNum,"
MySQL = MySQL & "                      dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork,"
MySQL = MySQL & "                      dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB, dbo.TblEmployee.kafeltel, dbo.TblEmployee.kafeladd,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc, dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc, dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1, dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1,"
MySQL = MySQL & "                      dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.project_id, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
MySQL = MySQL & "                      dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.BranchId,"
MySQL = MySQL & "                      dbo.TblEmployee.Account_Code5, dbo.TblEmployee.DriverId, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.gradeID, dbo.TblEmployee.DOBH,"
MySQL = MySQL & "                      dbo.TblEmployee.IssueDateH, dbo.TblEmployee.LastDate, dbo.TblEmployee.LastDateH, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.JobTypeID2,"
MySQL = MySQL & "                      dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.VisaNo, dbo.TblEmployee.GroupID, dbo.TblEmployee.mangerid, dbo.TblEmployee.swapedempid,"
MySQL = MySQL & "                      dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.DriverLicense, dbo.TblEmployee.DriverLicenseend,"
MySQL = MySQL & "                      dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.PerceTage,  dbo.TblEmployee.InstanceDateM,"
MySQL = MySQL & "                      dbo.TblEmployee.InstanceDateH, dbo.TblEvaluation.EvalItem, dbo.TblEvaluation.ActualDeg, dbo.TblEvaluation.MaxDeg, dbo.TblEvaluation.EvaluName,"
MySQL = MySQL & "                      dbo.TblEvaluation.MiniDeg , dbo.TblEvaluation.Remrks, dbo.TblEvaluation.EvaluDate"
MySQL = MySQL & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEvaluation ON dbo.TblEmployee.Emp_ID = dbo.TblEvaluation.Emp_ID"
MySQL = MySQL & " Where (dbo.TblEmployee.Emp_id = " & val(XPTxtEmpID.text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpEvaluation.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpEvaluation.rpt"
        End If
End If
'''
If C1Tab1.CurrTab = 5 Then
MySQL = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.WorkShop_Job,"
MySQL = MySQL & "                      dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Account_Code3, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.SalaryType, dbo.TblEmployee.Percentage,"
MySQL = MySQL & "                      dbo.TblEmployee.BYHour, dbo.TblEmployee.BankCard, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType1, dbo.TblEmployee.OpenBalance1,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.OpenBalance4,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id, dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate,"
MySQL = MySQL & "                      dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.placeWORK, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_sakn1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama, dbo.TblEmployee.Nationality,"
 MySQL = MySQL & "                     dbo.TblEmployee.dean, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Salary_others,"
MySQL = MySQL & "                      dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.KafelID, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.KafelName, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.workstate, dbo.TblEmployee.DateEndekama,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.pasplace, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region, dbo.TblEmployee.InsuranceValue,"
MySQL = MySQL & "                      dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH,"
MySQL = MySQL & "                      dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.EmpNum,"
MySQL = MySQL & "                      dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork,"
MySQL = MySQL & "                      dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB, dbo.TblEmployee.kafeltel, dbo.TblEmployee.kafeladd,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc, dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc, dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1, dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1,"
MySQL = MySQL & "                      dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.project_id, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
MySQL = MySQL & "                      dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.BranchId,"
MySQL = MySQL & "                      dbo.TblEmployee.Account_Code5, dbo.TblEmployee.DriverId, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.gradeID, dbo.TblEmployee.DOBH,"
MySQL = MySQL & "                      dbo.TblEmployee.IssueDateH, dbo.TblEmployee.LastDate, dbo.TblEmployee.LastDateH, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.JobTypeID2,"
MySQL = MySQL & "                      dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.VisaNo, dbo.TblEmployee.GroupID, dbo.TblEmployee.mangerid, dbo.TblEmployee.swapedempid,"
MySQL = MySQL & "                      dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.DriverLicense, dbo.TblEmployee.DriverLicenseend,"
MySQL = MySQL & "                      dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.PerceTage , dbo.TblEmployee.InstanceDateM,"
MySQL = MySQL & "                      dbo.TblEmployee.InstanceDateH, dbo.TblHealthy.FromH, dbo.TblHealthy.Clinic, dbo.TblHealthy.HealthyDate, dbo.TblHealthy.Remrks, dbo.TblHealthy.Des,"
MySQL = MySQL & "                      dbo.TblHealthy.HeathyTreat , dbo.TblHealthy.CompanyAmount"
MySQL = MySQL & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblHealthy ON dbo.TblEmployee.Emp_ID = dbo.TblHealthy.Emp_ID"
MySQL = MySQL & " Where (dbo.TblEmployee.Emp_id = " & val(XPTxtEmpID.text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpHelthey.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpHelthey.rpt"
        End If
End If
If C1Tab1.CurrTab = 10 Then
'MySQL = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.WorkShop_Job,"
'MySQL = MySQL & "                      dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Account_Code3, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.SalaryType, dbo.TblEmployee.Percentage,"
'MySQL = MySQL & "                      dbo.TblEmployee.BYHour, dbo.TblEmployee.BankCard, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
'MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType1, dbo.TblEmployee.OpenBalance1,"
'MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.OpenBalance4,"
'MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id, dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate,"
'MySQL = MySQL & "                      dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.placeWORK, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_sakn1,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama, dbo.TblEmployee.Nationality,"
'MySQL = MySQL & "                      dbo.TblEmployee.dean, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Salary_others,"
'MySQL = MySQL & "                      dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.KafelID, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
'MySQL = MySQL & "                      dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.KafelName, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.workstate, dbo.TblEmployee.DateEndekama,"
'MySQL = MySQL & "                      dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.pasplace, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region, dbo.TblEmployee.InsuranceValue,"
'MySQL = MySQL & "                      dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH,"
'MySQL = MySQL & "                      dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.EmpNum,"
'MySQL = MySQL & "                      dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork,"
'MySQL = MySQL & "                      dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB, dbo.TblEmployee.kafeltel, dbo.TblEmployee.kafeladd,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc, dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc, dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1, dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1,"
'MySQL = MySQL & "                      dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.project_id, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
'MySQL = MySQL & "                      dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.BranchId,"
'MySQL = MySQL & "                      dbo.TblEmployee.Account_Code5, dbo.TblEmployee.DriverId, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.gradeID, dbo.TblEmployee.DOBH,"
'MySQL = MySQL & "                      dbo.TblEmployee.IssueDateH, dbo.TblEmployee.LastDate, dbo.TblEmployee.LastDateH, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.JobTypeID2,"
'MySQL = MySQL & "                      dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.VisaNo, dbo.TblEmployee.GroupID, dbo.TblEmployee.mangerid, dbo.TblEmployee.swapedempid,"
'MySQL = MySQL & "                      dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.DriverLicense, dbo.TblEmployee.DriverLicenseend,"
'MySQL = MySQL & "                      dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.PerceTage,   dbo.TblEmployee.InstanceDateM,"
'MySQL = MySQL & "                      dbo.TblEmployee.InstanceDateH , dbo.TblEmpReign.ReignName, dbo.TblEmpReign.Amount, dbo.TblEmpReign.ReignDate, dbo.TblEmpReign.Remrks"
'MySQL = MySQL & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
' MySQL = MySQL & "                     dbo.TblEmpReign ON dbo.TblEmployee.Emp_ID = dbo.TblEmpReign.Emp_ID"
'MySQL = MySQL & " Where (dbo.TblEmployee.Emp_id = " & val(XPTxtEmpID.text) & ")"
MySQL = " SELECT     dbo.TblAssestes.AsName, dbo.TblAssestes.AsID, dbo.TblAssestes.AsDes, dbo.TblEmpAsestDetails.Remarks, dbo.TblEmpAsestDetails.Qunt, "
 MySQL = MySQL & "                     dbo.TblEmpAsestDetails.FlagAs, dbo.TblEmpAsest.EmpAsID, dbo.TblEmpAsestDetails.IDAseset, dbo.TblEmpAsest.EmpAsestID, dbo.TblEmpAsest.PostedDate,"
 MySQL = MySQL & "                     dbo.TblEmpAsest.RecordDate, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.Emp_Namee,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode,"
 MySQL = MySQL & "                     dbo.TblEmpAsestDetails.Remark2 , dbo.TblEmpAsestDetails.EmpID, dbo.TblEmployee.Emp_id"
MySQL = MySQL & " FROM         dbo.TblAssestes INNER JOIN"
MySQL = MySQL & "                      dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
MySQL = MySQL & "                      dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblEmpAsestDetails.EmpID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & " Where (dbo.TblEmployee.Emp_id =" & val(XPTxtEmpID.text) & ") And (dbo.TblEmpAsestDetails.FlagAs Is Null)"
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpReign.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpReign.rpt"
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

    Dim cCompanyInfo As New ClsCompanyInfo
Dim valmofrd As String

valmofrd = GetEmployeeSalaryAccordingToComponent(val(Me.XPTxtEmpID.text), "", 0)

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
       xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(valmofrd), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  xReport.ParameterFields(12).AddCurrentValue valmofrd
If C1Tab1.CurrTab = 0 Then
    xReport.ParameterFields(13).AddCurrentValue CStr(TxtEmpNotes)
    xReport.ParameterFields(14).AddCurrentValue CStr(DCRegionID.text)
    
End If

'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
        
          Dim xLogo As CRAXDRT.OLEObject
   ' StrFileName = App.path & "\"& SystemOptions.ImagesPath &"\" & PICNAME & ".JPG"
  If C1Tab1.CurrTab = 0 Or C1Tab1.CurrTab = 1 Then
  
         If Dir(App.path & "\" & SystemOptions.ImagesPath & "\" & val(XPTxtEmpID.text) & ".JPG") <> "" Then
          
    
        
            Set xLogo = xReport.Areas(1).Sections(2).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\" & val(FrmEmployee.XPTxtEmpID.text) & ".JPG", 500, 300)
            xLogo.Width = 1700
            xLogo.Height = 2000
            xLogo.backcolor = vbWhite
            xLogo.BorderColor = 255
            xLogo.CloseAtPageBreak = True
            
          End If
            
End If
  
    Set CViewer = New ClsReportViewer
    
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
       

        
            
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function

'Private Sub printingReport8()
'    Dim sql As String
'
'    'Dim Rs As ADODB.Recordset
''    Dim xReport As New CRAXDRT.Report
  '  Dim xApp As New CRAXDRT.Application
 ''   Dim rs As New ADODB.Recordset
  '  Dim reportpatath As String

  '  sql = "select * From emp_all_details ORDER BY CAST(Emp_Code AS integer) ASC "
    'sql = "select * From emp_all_details  where emp_id=" & val(Me.XPTxtEmpID.text)
    
  ' sql = "SELECT     dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, "
  ' sq = sql & "                    dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
  'sql = sql & "           dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary,"
'sql = sql & "          dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.workstate, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.JobTypeID,"
'sql = sql & "          dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.OtherDiscounts,"
'sql = sql & "          dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama, dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.DateEndekama,"
'sql = sql & "          dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc,"
'sql = sql & "          dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket,"
'sql = sql & "          dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.EmpNum, dbo.TblEmployee.CustNum,"
'sql = sql & "          dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork, dbo.TblEmployee.Notsstkala,"
'sql = sql & "          dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB, dbo.TblEmployee.KafelID, dbo.TblEmployee.KafelName, dbo.TblEmployee.pasplace,"
'sql = sql & "           dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.kafeltel,"
'sql = sql & "          dbo.TblEmployee.jopstatusid, dbo.TblEmployee.kafeladd, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
'sql = sql & "          dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_others,"
'sql = sql & "          dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1,"
'sql = sql & "          dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1,"
'sql = sql & "          dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc, dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc,"
'sql = sql & "          dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc, dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1,"
'sql = sql & "          dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1, dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1,"
'sql = sql & "          dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.placeWORK, dbo.TblEmployee.project_id, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Dateexppoketh,"
'sql = sql & "          dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode, dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode,"
'sql = sql & "          dbo.TblEmployee.BYHour, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3,"
'sql = sql & "          dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.BranchId, dbo.TblEmployee.Fullcode, dbo.TblEmployee.IssueDateH, dbo.TblEmployee.LastDate,"
'sql = sql & "          dbo.TblEmployee.LastDateH, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.JobTypeID2, dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.GroupID,"
'sql = sql & "          dbo.TblEmployee.VisaNo, dbo.TblEmployee.mangerid, dbo.TblEmployee.swapedempid, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.lastHolidaydate,"
'sql = sql & "          dbo.TblEmployee.DriverLicense, dbo.TblEmployee.DriverLicenseend, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.SalaryType,"
'sql = sql & "          dbo.TblEmployee.Percentage, dbo.TblEmployee.DriverLicenseendH, dbo.jopstatus.color, dbo.jopstatus.namee, dbo.TblEmployee.BankCard,"
'sql = sql & "          dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.DOBH, dbo.TblEmployee.gradeID, dbo.TblEmpDetails.relationtype, dbo.TblEmpDetails.iqamano,"
'sql = sql & "          dbo.TblEmpDetails.passportno, dbo.TblEmpDetails.OprType, dbo.TblEmpHolidaysDetails.fromdate, dbo.TblEmpHolidaysDetails.todate,"
'sql = sql & "          dbo.TblEmpHolidaysDetails.fromdateH , dbo.TblEmpHolidaysDetails.todateH, dbo.TblEmpHolidaysDetails.des"
'sql = sql & "          FROM         dbo.TblEmployee LEFT OUTER JOIN"
'sql = sql & "           dbo.TblEmpHolidaysDetails ON dbo.TblEmployee.Emp_ID = dbo.TblEmpHolidaysDetails.Emp_ID LEFT OUTER JOIN"
'sql = sql & "          dbo.TblEmpDetails ON dbo.TblEmployee.Emp_ID = dbo.TblEmpDetails.Emp_ID LEFT OUTER JOIN"
'sql = sql & "          dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
'sql = sql & "          dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
'sql = sql & "          dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
'sql = sql & "          WHERE     (dbo.TblEmpDetails.OprType = 0)"
'sql = sql & "  and     TblEmployee.Emp_id = " & val(Me.XPTxtEmpID.text)
     
'    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
'If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp.rpt"
'        Else
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp.rpt"
'        End If
'
'    Set xReport = xApp.OpenReport(reportpatath)
'    xReport.Database.SetDataSource rs
'
'    Set FrmReport = New FrmReportViewer
'    FrmReport.CRViewer.ReportSource = xReport
'
'    FrmReport.CRViewer.ViewReport
'    FrmReport.txtpath = reportpatath
'    FrmReport.show
'    Screen.MousePointer = vbDefault
'    '      xReport.ReportTitle = X
'    SendKeys "{RIGHT}"
'
    'Dim Msg As String
'    'On Error GoTo ErrTrap
'    'If XPTxtEmpID.text <> "" Then
'    '    Set EmpReport = New ClsEmployeeReport
'    '    EmpReport.EmpData XPTxtEmpID.text
'    'Else
'    '    Msg = "⁄„·Ì… «·ÿ»«⁄… €Ì— „ «Õ… Õ«·Ì«"
'    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    'End If
'    'Exit Sub
'    'ErrTrap:
'End Sub

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
lbl(59).Caption = "Section"
lbl(45).Caption = "Salary"
LblSalary(4).Caption = "Projects"
lbl(42).Caption = "W.P."
lbl(44).Caption = "Age"
lbl(48).Caption = "Years"
'****************************
lbl(58).Caption = "Driver ID Start Date"
lbl(60).Caption = "Driver ID End Date"


ChNoAdded.RightToLeft = False
ChNoAdded.Caption = "No Additional"
Fra(23).Caption = "Salary Type"
LblSalary(0).Caption = "Salary"
LblSalary(1).Caption = "Comission"
LblSalary(2).Caption = "By Hour"
LblSalary(3).Caption = "By Production"
Chk_Stkala.Caption = "Leave"
Fra(16).Caption = "Follower Data"
Label5.Caption = "Name"
lbl(27).Caption = "Iqam No"
lbl(29).Caption = "Remarks"
lbl(17).Caption = "payment type"
lbl(26).Caption = "Relation"
lbl(28).Caption = "Passport No."
lbl(25).Caption = "Banck Name"
lbl(40).Caption = "Banck Address"
lbl(31).Caption = "IBAN"
lbl(41).Caption = "Machin Code"
Fra(13).Caption = "Type"
TypeEmp(0).RightToLeft = False
TypeEmp(1).RightToLeft = False
TypeEmp(0).Caption = "Emp."
TypeEmp(1).Caption = "Mang."
With cboPayType
.Clear
.AddItem "Cash"
.AddItem "Cheque"
.AddItem "ATM"
.AddItem "Transfer"
.AddItem "Others"

End With
''''''''''''''''sa
XPLbl(61).Caption = "Status"
lbl(11).Caption = "Bank Code"
lbl(10).Caption = "Type contract"

lbl(30).Caption = "Insurance No."

lbl(29).Caption = "Remarks"
chkhaveinsurance.Caption = "have Insurance"

Cmd(20).Caption = "Add"
Cmd(21).Caption = "Delete"

    With Me.Grid
       .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("name")) = "name"
        .TextMatrix(0, .ColIndex("relationtypename")) = "Relation"
        .TextMatrix(0, .ColIndex("iqamano")) = "Iqama no"
        .TextMatrix(0, .ColIndex("passportno")) = "Passport No"
                .TextMatrix(0, .ColIndex("insuranceno")) = "insurance  No"
                        .TextMatrix(0, .ColIndex("des")) = "Remarks"
                        
    End With
    '''vaction
       With Me.Grid20
       .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("fromdateh")) = "From Date Hegira"
        .TextMatrix(0, .ColIndex("DateExpectedH")) = "To Date"
        .TextMatrix(0, .ColIndex("todateh")) = "Return Date Hegira"
        .TextMatrix(0, .ColIndex("fromdate")) = "From Date "
                .TextMatrix(0, .ColIndex("DateExpectedM")) = "TO Date "
                        .TextMatrix(0, .ColIndex("todate")) = "To Return Date"
                 .TextMatrix(0, .ColIndex("day")) = "Allowed"
                .TextMatrix(0, .ColIndex("a1")) = "Actual"
                        .TextMatrix(0, .ColIndex("dayEx")) = "Delay"
    End With
   'nnnn lbl(42).Caption = "Opening balance vacations"
  'nnnn  lbl(44).Caption = "Day"
 

    With Me.VSFlexGrid4
       .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("name")) = "name"
        .TextMatrix(0, .ColIndex("Date")) = "Date"
        .TextMatrix(0, .ColIndex("Qty")) = "Qty"
  
                        .TextMatrix(0, .ColIndex("des")) = "Remarks"
                        
    End With
    
    

Frame21.Caption = "  Data Holiday"
ALLButton2(5).Caption = "Annual Increases"
Fra(12).Caption = "Driver's license "
XPLbl(57).Caption = "No license"
XPLbl(58).Caption = "End"
XPLbl(55).Caption = "End"
XPLbl(56).Caption = "Start"
    With Me.VSFlexGrid2
       .TextMatrix(0, .ColIndex("Date")) = "Date"
        .TextMatrix(0, .ColIndex("ENTITY")) = "ENTITY"
        .TextMatrix(0, .ColIndex("clinic")) = "clinic"
        .TextMatrix(0, .ColIndex("des")) = "des"
  
                        .TextMatrix(0, .ColIndex("TypeName")) = "Vac type"
                         .TextMatrix(0, .ColIndex("valuecomp")) = "Company Payed"
                          .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
                          
                        
    End With
    
Fra(22).Caption = "Rating"


 


    With Me.VSFlexGrid3
       .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("name")) = "name"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("MinValue")) = "Min Value"
  
                        .TextMatrix(0, .ColIndex("MaxValue")) = "Max Value"
                         .TextMatrix(0, .ColIndex("rate")) = "rate"
                          .TextMatrix(0, .ColIndex("des")) = "Remarks"
                          
                        
    End With
    
    

Fra(18).Caption = "Qualifications"

Label7.Caption = "Name"
lbl(34).Caption = "Entity"
 lbl(35).Caption = "Remarks"
 lbl(32).Caption = "Grade"
lbl(33).Caption = "Year"

lbl(32).Caption = "Grade"

Cmd(8).Caption = "add"
Cmd(9).Caption = "Remove"
  
    With Me.grid01
       .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("name")) = "name"
        .TextMatrix(0, .ColIndex("qualicationEntity")) = "Qualication Entity"
        .TextMatrix(0, .ColIndex("yearofqualication")) = "Year"
  
                        .TextMatrix(0, .ColIndex("grade")) = "Grade"
                      
                          .TextMatrix(0, .ColIndex("des")) = "Remarks"
                         
    End With
    
    
    
    
    
 Frame20.Caption = "Experience"
lbl(62).Caption = "Name"
lbl(36).Caption = "Entity"
 lbl(51).Caption = "From"
 lbl(50).Caption = "To"
 

lbl(49).Caption = "Remarks"

Cmd(14).Caption = "add"
Cmd(15).Caption = "Remove"
  
  
 
    With Me.grid02
       .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("name")) = "name"
       .TextMatrix(0, .ColIndex("qualicationEntity")) = "Qualication Entity"
       
  
                        .TextMatrix(0, .ColIndex("workfrom")) = "work from"
                      .TextMatrix(0, .ColIndex("workfromH")) = "work from H"
                      .TextMatrix(0, .ColIndex("workto")) = "work to"
                      .TextMatrix(0, .ColIndex("worktoH")) = "work to H"
                      
                          .TextMatrix(0, .ColIndex("des")) = "Remarks"
                         
    End With
    
    
    
    lbl(21).Caption = "Last"
Fra(20).Caption = "Last Holiday Dates"
lbl(39).Caption = " Date"
 

XPLbl(49).Caption = "Job"

'XPLbl(53).Caption = "Comp. %"
'XPLbl(54).Caption = "Employee. %"




'**********************************
 


    XPLbl(46).Caption = "Work Place"
    SuperLabel1(5).text = "Project"
    'lblB(9).text = "Branch"
'    lbl(31).Caption = "Credit"
    lbl(6).Visible = False
XPLbl(48).Caption = "Grade"
    Dim XPic As IPictureDisp
    Fra(8).Caption = "OB  Accounts debitors"
    Fra(9).Caption = "OB Due Salary"
    Fra(10).Caption = "OB  Holiday Allocations "
    
    '******************************
    Fra(11).Caption = "OB  End Service Allocations "
    
        OptType4(0).Caption = "Depit"
    OptType4(1).Caption = "Credit"
    OptType4(2).Caption = "NA"
    lbl(24).Caption = "Balance"
'    lbl(25).Caption = "Date"
    
'    Frame4.Caption = "Bank Data"
    lbl(19).Caption = "Acc. No"
    XPLbl(50).Caption = "Job"
    
    '******************************
    
    
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
    XPLbl(38).Caption = "Start Date"
    Command5.Caption = "Hide"

    XPLbl(27).Caption = "Nationality"
    XPLbl(28).Caption = "Religion"
    lbl(12).Caption = "DOB"
   ' lbl(11).Caption = "Leaving"
    'Chk_Stkala.Caption = "Resignation"
    Chk_EndWork.Caption = "Separation"
    chkStop.Caption = "Stop monthly subscription"
   ' lbl(10).Caption = "Date"
 '   CmdEstkala.Caption = "Reason"
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
    XPLbl(7).Caption = "Department"
    XPLbl(5).Caption = "Sections"
    XPLbl(59).Caption = "Zone\Sector"
    XPLbl(9).Caption = "Job"
    XPLbl(62).Caption = "End Date"
    XPLbl(8).Caption = "Work Teams"
    lbl(9).Caption = "Start Date"
    Fra(0).Caption = "Insurance"
    XPLbl(4).Caption = "Status"
  '  XPLbl(5).Caption = "value"
    XPLbl(6).Caption = "Insu. No"
    Fra(2).Caption = "Licence"
    XPLbl(13).Caption = "No"
    XPLbl(12).Caption = "Start"
    XPLbl(11).Caption = "End"
    Fra(4).Caption = "Saudi ID"
    Fra(3).Caption = "IQama No"

  '  Fra(8).Caption = "Opening Balance Depitors"
    OptType(0).Caption = "Depit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "NA"
    lbl(14).Caption = "Balance"
    lbl(13).Caption = "Date"

  '  Fra(9).Caption = "Opening Balance  "
    OptType1(0).Caption = "Depit"
    OptType1(1).Caption = "Credit"
    OptType1(2).Caption = "NA"
    lbl(15).Caption = "Balance"
    lbl(16).Caption = "Date"

  '  Fra(10).Caption = "Opening Balance "
    OptType2(0).Caption = "Depit"
    OptType2(1).Caption = "Credit"
    OptType2(2).Caption = "NA"
    lbl(18).Caption = "Balance"
    
       Fra(14).Caption = "O B Ticket"
    OptType5(0).Caption = "Depit"
    OptType5(1).Caption = "Credit"
    OptType5(2).Caption = "NA"
    lbl(16).Caption = "Balance"
    
    
    
    
    
    
'    lbl(17).Caption = "Date"

    XPLbl(18).Caption = "No"
    XPLbl(19).Caption = "Start"
    XPLbl(20).Caption = "End"
  '  CmdExit.Caption = "Exit"
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
    Fra(25).Caption = "Query"
    OptExpirLinc.Caption = "License"
    OptExpirEkama.Caption = "Residence"
    OptExpirPas.Caption = "Passport"
    CommandÛQRY.Caption = "Query"

    Frame3.Caption = "Entry Data"

    XPLbl(29).Caption = "NO"
    XPLbl(30).Caption = "Date"
    XPLbl(31).Caption = "Port"
    lbl(61).Caption = "Forms"
    Check1.Caption = "Print Image"
    ALLButton1.Caption = "Print"
    Combo1.Clear
    Combo1.AddItem "new Residence"
    Combo1.AddItem "Renew Residence"
    Combo1.AddItem "Residence Replacement"
    Combo1.AddItem "Residence Damaged"
    Combo1.AddItem "Visa"
    Combo1.AddItem "Absence Form"

    Me.Caption = "Employees Data"
    EleHeader.Caption = Me.Caption
    XPLbl(1).Caption = "Employee Code"
    XPLbl(0).Caption = " Name AR"
    XPLbl(47).Caption = " Name ENG"
    
    XPLbl(2).Caption = "Gender"
    
    
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

 Me.C1Tab1.TabCaption(0) = "Basic Data"
    Me.C1Tab1.TabCaption(1) = "Components & Contracts"
    Me.C1Tab1.TabCaption(2) = "Follower"
    Me.C1Tab1.TabCaption(3) = "Qualifications and experience"
    
   Me.C1Tab1.TabCaption(4) = "Rating"
     Me.C1Tab1.TabCaption(5) = "Health file"
  Me.C1Tab1.TabCaption(6) = "Insurance Data"
  Me.C1Tab1.TabCaption(7) = "Visas"
  
  Me.C1Tab1.TabCaption(8) = "Vacations and direct-entry"
  Me.C1Tab1.TabCaption(9) = "Accounting Data "
Me.C1Tab1.TabCaption(10) = "Testament kind"
Me.C1Tab1.TabCaption(11) = "Attachments"



XPLbl(52).Caption = "Branch"

lbl(22).Caption = "Manger"

lbl(23).Caption = "Alternative"

XPLbl(51).Caption = "Passport Job"

XPLbl(51).Caption = "Passport Job"

Frame8.Caption = "Data Direct"
lbl(20).Caption = "Visa No."
lbl(38).Caption = "Start Date"
Fra(19).Caption = "Health File"

End Sub


Private Sub txtDateEndIndustrialHijri_LostFocus()
      txtDateEndIndustrial.value = ToGregorianDate(txtDateEndIndustrialHijri.value)

End Sub

