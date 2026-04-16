VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCarsPlan 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ŒÿÂ ’Ì«‰Â «·„—ﬂ»« "
   ClientHeight    =   7755
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14775
   Icon            =   "FrmCarsPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   14775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtNoHour 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   152
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox TXTCurrentKM 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   131
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Frame Frame8 
      Caption         =   "»Ì«‰«  «·’Ì«‰Â"
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   116
      Top             =   1920
      Width           =   14745
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7770
         RightToLeft     =   -1  'True
         TabIndex        =   165
         Top             =   210
         Width           =   1005
      End
      Begin VB.Frame Frame4 
         Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
         Height          =   615
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   160
         Top             =   600
         Width           =   2655
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            Caption         =   " Ê«—ÌŒ"
            Height          =   195
            Index           =   1
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂ„"
            Height          =   195
            Index           =   0
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            Caption         =   "”«⁄« "
            Height          =   195
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "›Ì Õ«·Â «·”«⁄« "
         Height          =   615
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   157
         Top             =   600
         Width           =   2415
         Begin VB.TextBox TxtHour 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂ· ⁄œœ ”«⁄« "
            Height          =   315
            Index           =   18
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   240
            Width           =   1245
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "›Ì Õ«·Â «·Êﬁ "
         Height          =   855
         Left            =   12570
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   1530
         Visible         =   0   'False
         Width           =   4935
         Begin VB.Frame Frame10 
            Caption         =   "Õœœ"
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   240
            Width           =   2655
            Begin VB.OptionButton optintervals 
               Alignment       =   1  'Right Justify
               Caption         =   "œﬁÌﬁÂ"
               Height          =   195
               Index           =   5
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   146
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optintervals 
               Alignment       =   1  'Right Justify
               Caption         =   "À«‰ÌÂ"
               Height          =   195
               Index           =   4
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optintervals 
               Alignment       =   1  'Right Justify
               Caption         =   "”«⁄…"
               Height          =   195
               Index           =   3
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.TextBox Txtpriod2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Text            =   "1"
            Top             =   480
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Left            =   3480
            TabIndex        =   147
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Format          =   109445122
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Êﬁ  «·√”«”"
            Height          =   375
            Index           =   16
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —… »Ì‰ ﬂ· „—…"
            Height          =   315
            Index           =   15
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   120
            Width           =   1245
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "›Ì Õ«·Â ﬂÌ·Ê„ —« "
         Height          =   615
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   138
         Top             =   600
         Width           =   1815
         Begin VB.TextBox txtKMCount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂ· ⁄œœ ﬂ„"
            Height          =   315
            Index           =   11
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "›Ì Õ«·Â  Ê«—ÌŒ"
         Height          =   855
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   121
         Top             =   540
         Width           =   4815
         Begin VB.TextBox TxXPeriod 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Text            =   "1"
            Top             =   480
            Width           =   615
         End
         Begin VB.Frame Frame5 
            Caption         =   "Õœœ"
            Height          =   495
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   240
            Width           =   2295
            Begin VB.OptionButton optintervals 
               Alignment       =   1  'Right Justify
               Caption         =   "”‰Â"
               Height          =   195
               Index           =   2
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optintervals 
               Alignment       =   1  'Right Justify
               Caption         =   "ÌÊ„"
               Height          =   195
               Index           =   0
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optintervals 
               Alignment       =   1  'Right Justify
               Caption         =   "‘Â—"
               Height          =   195
               Index           =   1
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin MSComCtl2.DTPicker DPStartDate 
            Height          =   345
            Left            =   3240
            TabIndex        =   127
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Format          =   180551681
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —… »Ì‰ ﬂ· „—…"
            Height          =   315
            Index           =   8
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ  «·√”«”"
            Height          =   375
            Index           =   6
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.TextBox TxtNoOfIteration 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   117
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DCMaintenanceTypes 
         Height          =   315
         Left            =   2160
         TabIndex        =   118
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton CmdAdd 
         Height          =   390
         Left            =   0
         TabIndex        =   137
         Top             =   1200
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   688
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "≈÷«›…"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCarsPlan.frx":000C
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DcbGroup 
         Height          =   315
         Left            =   10320
         TabIndex        =   155
         Top             =   240
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»ÕÀ «·”—Ì⁄"
         Height          =   315
         Index           =   20
         Left            =   8850
         RightToLeft     =   -1  'True
         TabIndex        =   164
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›∆…"
         Height          =   315
         Index           =   19
         Left            =   13560
         RightToLeft     =   -1  'True
         TabIndex        =   156
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·’Ì«‰Â"
         Height          =   315
         Index           =   4
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   120
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·„—« "
         Height          =   315
         Index           =   7
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   119
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   8970
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   114
      Top             =   720
      Width           =   1485
   End
   Begin VB.TextBox txtLastKMCounter 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   16800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   105
      Top             =   2880
      Width           =   1365
   End
   Begin VB.TextBox txtModel 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   19800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   102
      Top             =   2880
      Width           =   1605
   End
   Begin VB.TextBox TxtLicenseNO 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   19800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   101
      Top             =   2520
      Width           =   1605
   End
   Begin VB.TextBox txtBoardNO 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   18360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   100
      Top             =   3240
      Width           =   1605
   End
   Begin VB.TextBox txtopening_balance_voucher_id 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   16440
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   95
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtNoteID1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   94
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   16320
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   93
      Top             =   5880
      Width           =   2325
   End
   Begin VB.TextBox TxtSalePrice 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   15600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   92
      Top             =   5400
      Width           =   2325
   End
   Begin VB.TextBox txtNoteID 
      Height          =   285
      Left            =   16200
      RightToLeft     =   -1  'True
      TabIndex        =   87
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtNoteSerial 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   17160
      RightToLeft     =   -1  'True
      TabIndex        =   86
      Top             =   6120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Left            =   15360
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Top             =   75
      Width           =   2175
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÃœÌœ"
         Height          =   195
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   120
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "«›  «ÕÌ"
         Height          =   195
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtPurchaseBillId 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   15480
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   360
      Width           =   1245
   End
   Begin VB.TextBox TxtKhordaPrice 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   15360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   720
      Width           =   2325
   End
   Begin VB.TextBox TxtCurrentValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   15480
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   600
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Caption         =   "»Ì«‰«  «·«Â·«ﬂ"
      Height          =   2415
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   9360
      Width           =   8535
   End
   Begin VB.TextBox txtinstallDo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   15960
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   3360
      Width           =   1605
   End
   Begin VB.TextBox txtinstallmentresult 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   16680
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   720
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  „Ã„Ê⁄Â «·«’·"
      Enabled         =   0   'False
      Height          =   2895
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   3840
      Width           =   6375
      Begin VB.TextBox TxtAge 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   360
         Width           =   525
      End
      Begin VB.OptionButton Optxxx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·Ì” ·Â «Â·«ﬂ"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   120
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Optx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·Â «Â·«ﬂ"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   2370
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox TXT24 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1080
         Width           =   3885
      End
      Begin VB.TextBox TXT26 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1440
         Width           =   3885
      End
      Begin VB.TextBox TXT25 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1800
         Width           =   3885
      End
      Begin VB.TextBox TXT31 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   2160
         Width           =   3885
      End
      Begin VB.TextBox TXT40 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2520
         Width           =   3885
      End
      Begin VB.TextBox txtPercentage2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2760
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   720
         Width           =   1245
      End
      Begin VB.TextBox TXtPercentage1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2760
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„— «·«› —«÷Ì ··«’· »«·‘Â—"
         Height          =   255
         Index           =   9
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«» «·«’·  »«·„Ì“«‰Ì…"
         Height          =   255
         Index           =   111
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«» „Ã„⁄ «·«Â·«ﬂ"
         Height          =   255
         Index           =   112
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»    „’—Ê›«  «·«Â·«ﬂ"
         Height          =   255
         Index           =   113
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»   «—»«Õ »Ì⁄"
         Height          =   255
         Index           =   114
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   2160
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»   Œ”«∆— »Ì⁄"
         Height          =   255
         Index           =   115
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·«Â·«ﬂ"
         Height          =   255
         Index           =   109
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·«Â·«ﬂ ⁄‰œ «·«Ìﬁ«›"
         Height          =   255
         Index           =   110
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   720
         Width           =   1995
      End
   End
   Begin VB.ComboBox cStatus 
      Height          =   315
      ItemData        =   "FrmCarsPlan.frx":03A6
      Left            =   15240
      List            =   "FrmCarsPlan.frx":03B6
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox CBoDepreciation_Type_id 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmCarsPlan.frx":03F6
      Left            =   14880
      List            =   "FrmCarsPlan.frx":0400
      TabIndex        =   42
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox TxtnoOfInst 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   14760
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3360
      Width           =   1605
   End
   Begin VB.TextBox txtinstallValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   15960
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   600
      Width           =   2325
   End
   Begin VB.TextBox TxtAccDepreciation 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   15960
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   480
      Width           =   1605
   End
   Begin VB.TextBox XPTxtID0 
      Height          =   285
      Left            =   6960
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtPurchasePrice 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   15360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   720
      Width           =   1605
   End
   Begin VB.TextBox TxtNotes 
      Alignment       =   1  'Right Justify
      Height          =   555
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   4245
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   675
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -120
      Width           =   14835
      _cx             =   26167
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   20.25
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
      Caption         =   "ŒÿÂ ’Ì«‰Â «·„—ﬂ»« /«·„⁄œ« "
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
         Left            =   1155
         TabIndex        =   3
         Top             =   120
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
         ButtonImage     =   "FrmCarsPlan.frx":0423
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
         Left            =   90
         TabIndex        =   4
         Top             =   120
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
         ButtonImage     =   "FrmCarsPlan.frx":07BD
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
         Left            =   1680
         TabIndex        =   5
         Top             =   120
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
         ButtonImage     =   "FrmCarsPlan.frx":0B57
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
         Left            =   615
         TabIndex        =   6
         Top             =   120
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
         ButtonImage     =   "FrmCarsPlan.frx":0EF1
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   7
      Top             =   7275
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   7680
      TabIndex        =   8
      Top             =   7275
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   6795
      TabIndex        =   9
      Top             =   7275
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   5925
      TabIndex        =   10
      Top             =   7275
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   3960
      TabIndex        =   11
      Top             =   7275
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
      Index           =   6
      Left            =   1290
      TabIndex        =   12
      Top             =   7275
      Width           =   795
      _ExtentX        =   1402
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
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   345
      Left            =   9960
      TabIndex        =   23
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   155779073
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   6840
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DpLicenseExpireDate 
      Height          =   345
      Left            =   15600
      TabIndex        =   27
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      Format          =   155779073
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   16800
      TabIndex        =   37
      Top             =   1920
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DPReceiveDate 
      Height          =   345
      Left            =   15000
      TabIndex        =   40
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   609
      _Version        =   393216
      Format          =   155779073
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   44
      Top             =   7275
      Width           =   1035
      _ExtentX        =   1826
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   15000
      TabIndex        =   45
      Top             =   5640
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«Ìﬁ«› «·«Â·«ﬂ"
      BackColor       =   14871017
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
      Index           =   8
      Left            =   18000
      TabIndex        =   46
      Top             =   5640
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "≈⁄«œ…  ‘€Ì· «·«Â·«ﬂ"
      BackColor       =   14871017
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
      Index           =   9
      Left            =   15840
      TabIndex        =   47
      Top             =   5640
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«· Œ·’ „‰ «·«’·"
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
      Height          =   315
      Index           =   10
      Left            =   2400
      TabIndex        =   76
      Top             =   360
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «·›« Ê—…"
      BackColor       =   14871017
      ForeColor       =   16711680
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
      ColorToggledText=   16711680
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSComCtl2.DTPicker DpPurchaseDate 
      Height          =   345
      Left            =   15360
      TabIndex        =   77
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      Format          =   155779073
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DpTestExpireDate 
      Height          =   345
      Left            =   15360
      TabIndex        =   81
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   609
      _Version        =   393216
      Format          =   155844609
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   11
      Left            =   15360
      TabIndex        =   88
      Top             =   5235
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
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
   Begin MSDataListLib.DataCombo DCGroup 
      Height          =   315
      Left            =   15360
      TabIndex        =   97
      Top             =   2520
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcEmployee 
      Height          =   315
      Left            =   15360
      TabIndex        =   98
      Top             =   2880
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   12
      Left            =   2160
      TabIndex        =   99
      Top             =   7275
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
   Begin MSComCtl2.DTPicker DpInsuranceExpireDate 
      Height          =   345
      Left            =   15120
      TabIndex        =   104
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   609
      _Version        =   393216
      Format          =   155844609
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DCInsuranceCompanyId 
      Height          =   315
      Left            =   15600
      TabIndex        =   107
      Top             =   2880
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Dynamic_Byte.NourHijriCal DpLicenseExpireDateH 
      Height          =   255
      Left            =   16800
      TabIndex        =   109
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
   End
   Begin Dynamic_Byte.NourHijriCal DpTestExpireDateH 
      Height          =   255
      Left            =   15600
      TabIndex        =   110
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
   End
   Begin Dynamic_Byte.NourHijriCal DpInsuranceExpireDateH 
      Height          =   255
      Left            =   17640
      TabIndex        =   111
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2940
      Left            =   -1800
      TabIndex        =   112
      Top             =   3720
      Width           =   16740
      _cx             =   29527
      _cy             =   5186
      Appearance      =   2
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   29
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCarsPlan.frx":128B
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
   End
   Begin MSComCtl2.DTPicker DPRecordDate 
      Height          =   345
      Left            =   6120
      TabIndex        =   132
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   155844609
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcyearsData 
      Height          =   315
      Left            =   120
      TabIndex        =   135
      Top             =   1200
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo Dccar 
      Height          =   315
      Left            =   6120
      TabIndex        =   136
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdDelete 
      Height          =   390
      Left            =   9960
      TabIndex        =   150
      Top             =   6600
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–›"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmCarsPlan.frx":1713
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton cmdClear 
      Height          =   390
      Left            =   8760
      TabIndex        =   151
      Top             =   6600
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–› «·ﬂ·"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmCarsPlan.frx":1CAD
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   13
      Left            =   3000
      TabIndex        =   154
      Top             =   7275
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·”«⁄« "
      Height          =   315
      Index           =   17
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   153
      Top             =   720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”‰Â «·„«·Ì…"
      Height          =   315
      Index           =   14
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   134
      Top             =   1200
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·ŒÿÂ"
      Height          =   315
      Index           =   13
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   133
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁ—«¡… «·Õ«·Ì… ··⁄œ«œ"
      Height          =   315
      Index           =   12
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   130
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ê’› «·ŒÿÂ"
      Height          =   195
      Index           =   124
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   115
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·ŒÿÂ"
      Height          =   315
      Index           =   10
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   113
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‘—ﬂ… «· √„Ì‰"
      Height          =   375
      Index           =   3
      Left            =   18240
      RightToLeft     =   -1  'True
      TabIndex        =   108
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Œ— ﬁ—«¡… ··⁄œ«œ"
      Height          =   255
      Index           =   2
      Left            =   18240
      RightToLeft     =   -1  'True
      TabIndex        =   106
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ  «·‘—«¡"
      Height          =   375
      Index           =   1
      Left            =   16800
      RightToLeft     =   -1  'True
      TabIndex        =   103
      Top             =   3240
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„ﬂ”» «Ê «·Œ”«—…"
      Height          =   375
      Left            =   17280
      RightToLeft     =   -1  'True
      TabIndex        =   91
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”⁄— «·»Ì⁄"
      Height          =   255
      Left            =   18720
      RightToLeft     =   -1  'True
      TabIndex        =   90
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   195
      Index           =   0
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   89
      Top             =   6120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ‰Â«Ì… «·›Õ’"
      Height          =   375
      Index           =   120
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   82
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ‰Â«Ì… «·«” „«—…"
      Height          =   255
      Index           =   128
      Left            =   18240
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   2520
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ›« Ê—… «·‘—«¡"
      Height          =   255
      Index           =   116
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·«’· ﬂŒ—œ…"
      Height          =   375
      Index           =   121
      Left            =   15360
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·«” „«—…"
      Height          =   315
      Index           =   106
      Left            =   21480
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰›–"
      Height          =   255
      Index           =   130
      Left            =   15360
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ »ﬁÏ"
      Height          =   255
      Index           =   123
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «··ÊÕ…"
      Height          =   255
      Index           =   105
      Left            =   20040
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«” ·«„"
      Height          =   375
      Index           =   119
      Left            =   15720
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»⁄ÂœÂ"
      Height          =   315
      Index           =   104
      Left            =   20040
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   315
      Index           =   117
      Left            =   17040
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
      Height          =   315
      Index           =   103
      Left            =   20040
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Œ— ﬁ—«¡… ··⁄œ«œ"
      Height          =   255
      Index           =   108
      Left            =   19200
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3360
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… ﬁ”ÿ «·«Â·«ﬂ"
      Height          =   255
      Index           =   122
      Left            =   16920
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„⁄ «·«Â·«ﬂ"
      Height          =   255
      Index           =   129
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ  ‰Â«Ì… «· √„Ì‰"
      Height          =   375
      Index           =   127
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2520
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·«’·"
      Height          =   255
      Index           =   118
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   -120
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   315
      Index           =   5
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6840
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LngDevID 
      Height          =   375
      Left            =   6960
      TabIndex        =   20
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ÊœÌ·"
      Height          =   315
      Index           =   107
      Left            =   21480
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   5010
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   6750
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   6750
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   126
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6720
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   125
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   6720
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„⁄œ…"
      Height          =   315
      Index           =   101
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   1395
   End
End
Attribute VB_Name = "FrmCarsPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RSAss As New ADODB.Recordset
Dim FirstPeriodDateInthisYear  As Date
Dim TTP As clstooltip

Private Sub Removeˆall()
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Function addrow()

    Dim j As Integer
    Dim lastrow As Integer
    Dim sql As String
    Dim i As Integer
    Dim Msg  As String
    Dim startkm As Double
    Dim sumKm As Double
    Dim SumHour As Double
    Dim StartDate As Date
    Dim sumdate As Date
    startkm = val(TXTCurrentKM.Text)
    sumKm = startkm
    StartDate = DPStartDate.value
    sumdate = StartDate
SumHour = 0
    If DCMaintenanceTypes.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»       «Œ Ì«— ‰Ê⁄ «·’Ì«‰…  ...!!!"
        Else
            Msg = "must Specify  Maintenance type Name ...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
        DCMaintenanceTypes.SetFocus
      SendKeys "{F4}"
        Exit Function
    End If
 
    If val(TxtNoOfIteration.Text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õœœ ⁄œœ „—«  «· ﬂ—«—  ...!!!"
        Else
            Msg = "must Specify  No Of Iterations ...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtNoOfIteration.SetFocus
        Exit Function
    End If

    If Opt(0).value = True Then ' km
        If val(txtKMCount.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ ⁄œœ   «·ﬂÌ·Ê„ —«   ...!!!"
            Else
                Msg = "must Specify  KM's  ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            txtKMCount.SetFocus
            Exit Function
        End If
ElseIf Opt(2).value = True Then
If val(TxtHour.Text) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ ⁄œœ  «·”«⁄«    ...!!!"
            Else
                Msg = "must Specify  Hours...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 
            Exit Function
End If
    Else

        If val(TxXPeriod.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ     «·› —«  »Ì‰ «· Ê«—ÌŒ   ...!!!"
            Else
                Msg = "must Specify  Intervals...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxXPeriod.SetFocus
            Exit Function
        End If

    End If
 
    With Grid
        lastrow = .Rows
    
        If val(TxtNoOfIteration.Text) > 0 Then
            .Rows = val(TxtNoOfIteration.Text) + lastrow
      
            For i = lastrow To val(TxtNoOfIteration.Text) + lastrow - 1
                .TextMatrix(i, .ColIndex("MaintenanceID")) = val(DCMaintenanceTypes.BoundText)
                .TextMatrix(i, .ColIndex("MaintenanceType")) = DCMaintenanceTypes.Text
                .TextMatrix(i, .ColIndex("GroupID")) = val(DcbGroup.BoundText)
                .TextMatrix(i, .ColIndex("Group")) = Me.DcbGroup.Text

                If Opt(0).value = True Then
                    .TextMatrix(i, .ColIndex("AlarmType")) = 0
                    sumKm = sumKm + txtKMCount.Text
                    .TextMatrix(i, .ColIndex("AlarmInKM")) = sumKm
                    .TextMatrix(i, .ColIndex("AlarmINDate")) = ""
                     .TextMatrix(i, .ColIndex("AlarmINTime")) = ""
                     ElseIf Opt(2).value = True Then
                     SumHour = SumHour + val(TxtHour.Text)
                     .TextMatrix(i, .ColIndex("AlarmType")) = 2
                      .TextMatrix(i, .ColIndex("AlarmINDate")) = ""
                    .TextMatrix(i, .ColIndex("AlarmINTime")) = ""
                     .TextMatrix(i, .ColIndex("hour")) = SumHour
                     
                      '                     If optintervals(4).value = True Then
                      '  sumdate = DateAdd("s", val(Txtpriod2.text), sumdate)
                  '  ElseIf optintervals(5).value = True Then
                  '      sumdate = DateAdd("n", val(Txtpriod2.text), sumdate)
                  '  ElseIf optintervals(3).value = True Then
                  '      sumdate = DateAdd("h", val(Txtpriod2.text), sumdate)
                  '
                  '  End If
                
                    .TextMatrix(i, .ColIndex("AlarmInKM")) = ""
                    .TextMatrix(i, .ColIndex("AlarmINDate")) = ""
                    .TextMatrix(i, .ColIndex("AlarmINTime")) = FormatDateTime(sumdate, vbShortTime)
                    
                Else
                    .TextMatrix(i, .ColIndex("AlarmType")) = 1

                    If optintervals(0).value = True Then
                        sumdate = DateAdd("D", val(TxXPeriod.Text), sumdate)
                    ElseIf optintervals(1).value = True Then
                        sumdate = DateAdd("M", val(TxXPeriod.Text), sumdate)
                    ElseIf optintervals(2).value = True Then
                        sumdate = DateAdd("YYYY", val(TxXPeriod.Text), sumdate)
                                 
                    End If
                
                    .TextMatrix(i, .ColIndex("AlarmInKM")) = ""
                    .TextMatrix(i, .ColIndex("AlarmINDate")) = sumdate
                    .TextMatrix(i, .ColIndex("AlarmINTime")) = ""
                End If
 
            Next i
 
            '    .AutoSize 0, .Cols - 1, False
        End If

    End With

    ReLineGrid

End Function

Private Sub cmdAdd_Click()
    addrow
End Sub
 
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("MaintenanceID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Private Sub cmdClear_Click()
    Removeˆall
End Sub

Private Sub CmdDelete_Click()
    RemoveGridRow
End Sub

Private Sub DcbGroup_Change()
DcbGroup_Click (0)
End Sub

Private Sub DcbGroup_Click(Area As Integer)
loadDcbMaint val(DcbGroup.BoundText)
End Sub

Private Sub Dccar_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
          Load FrmCasrShearches
         FrmCasrShearches.SendForm = "FrmCarsPlan"
         FrmCasrShearches.Show vbModal
    End If
End Sub

Private Sub DCMaintenanceTypes_Click(Area As Integer)
    Dim km As String

    If Opt(0).value = True Then
'        getMaintenancetypeInformations val(DCMaintenanceTypes.BoundText), , km
'        txtKMCount.text = km
    End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
                            
          If Grid.ColIndex("AlarmINDate") = Col Then
          Else
    Cancel = True
    End If
End Sub

Public Sub Cmd_Click(Index As Integer)
    Dim msgstr  As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            TxtModFlg.Text = "N"
            clear_all Me
 
            Me.DCboUserName.BoundText = user_id
            Me.dcBranch.BoundText = branch_id
            DPRecordDate.value = Date
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            Opt(0).value = True
            optintervals(0).value = True

        Case 1
            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id
          
            '              Me.dcBranch.BoundText = my_branch
            CuurentLogdata
            Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
        
        Case 2

            SaveData

        Case 3
            Call Undo

        Case 4
            Del_AssetType

        Case 5
            VIEW_ATTACH
    
        Case 6
            Unload Me

        Case 7 ' «Ìﬁ«› «·«Â·«ﬂ
            cStatus.ListIndex = 1
            Cmd(7).Enabled = False

        Case 8 ' «⁄«œ…  ‘€Ì· «·«Â·«ﬂ
            cStatus.ListIndex = 0
            Cmd(8).Enabled = False

        Case 9 ' «· Œ·’ „‰ «·«’·
    
            cStatus.ListIndex = 3
            Cmd(9).Enabled = False
    
        Case 10

'wael
'             FrmExpenses4.Show
'
'            If val(Me.txtPurchaseBillId.Text) = 0 Then
'                 FrmExpenses4.Retrive -1
'            Else
'                 FrmExpenses4.Retrive val(Me.txtPurchaseBillId.Text)
'            End If
'
        Case 11
            ShowGL_cc Me.txtNoteSerial.Text, , 200
           Case 12
                 If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report val(Me.XPTxtID.Text)
               
            End If

        Case 13
   
              FrmSearchCarsPlan.Indx = 0
              Load FrmSearchCarsPlan
              FrmSearchCarsPlan.Indx = 0
              FrmSearchCarsPlan.Show vbModal
       
    End Select

    Exit Sub
ErrTrap:
End Sub

Function VIEW_ATTACH()
    'On Error Resume Next''
 
    'If TxtEmp_Code.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „ÊŸ› «Ê·«": Exit Sub

    imaged.Show
    imaged.Label9.Caption = "„—›ﬁ«  «·„⁄œÂ/«·”Ì«—… —ﬁ„"
    imaged.Caption = "„—›ﬁ«  «·„⁄œÂ/«·”Ì«—…  "
    imaged.txtopeation_type = "„—›ﬁ«  «·„⁄œÂ/«·”Ì«—…"
    imaged.SUBJECT_NO = XPTxtID 'TxtEmp_Code.text
    imaged.Label6.Caption = "ﬂÊœ «·„⁄œÂ/«·”Ì«—…"
    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  «·„⁄œÂ/«·”Ì«—…' and subject_no='" & XPTxtID & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Function
 
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ «·ŒÿÂ " & XPTxtID & CHR(13) & " Ê’› «·ŒÿÂ       " & TxtNotes & CHR(13) & "   «·„⁄œÂ/«·”Ì«—…   " & Dccar & CHR(13) & "      ﬁ—«¡… «·⁄œ«œ «·Õ«·Ì…   " & TXTCurrentKM
    LogTextA = "    ‘«‘… " & ScreenNameEnglish & CHR(13) & "plan Code   " & XPTxtID & CHR(13) & "Plan Desc" & TxtNotes & CHR(13) & "   Car   " & Dccar & CHR(13) & "   Current KM  " & TXTCurrentKM
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , val(XPTxtID)
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , val(XPTxtID)
    End If
    
End Function
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.TblCarMaintenancePlan.Planid, dbo.TblCarMaintenancePlan.RecordDate, dbo.TblCarMaintenancePlan.CurrentKM, dbo.TblCarMaintenancePlan.CarId, "
MySQL = MySQL & "                      dbo.TblCarsData.BoardNO, dbo.TblCarMaintenancePlanDetails.MaintenanceID, dbo.TblMaintenanceType.name, dbo.TblCarMaintenancePlan.PlanYear,"
MySQL = MySQL & "                      dbo.TblMaintenanceType.namee, dbo.TblCarMaintenancePlanDetails.alarmType, dbo.TblCarMaintenancePlanDetails.Done,"
MySQL = MySQL & "                      dbo.TblCarMaintenancePlanDetails.DoneDate, dbo.TblCarMaintenancePlanDetails.CurrentKM AS CurrentKMDet, dbo.TblCarMaintenancePlanDetails.AlarmInKM,"
MySQL = MySQL & "                      dbo.TblCarMaintenancePlanDetails.AlarmINDate, dbo.TblCarMaintenancePlanDetails.AlarmINTime, dbo.TblCarMaintenancePlanDetails.NoHour,"
MySQL = MySQL & "                      dbo.TblCarMaintenancePlan.Remarks, dbo.TblCarMaintenancePlanDetails.GroupID, TblMaintenanceType_1.name AS GroupName,"
MySQL = MySQL & "                      TblMaintenanceType_1.namee AS GroupNameE"
MySQL = MySQL & " FROM         dbo.TblMaintenanceType TblMaintenanceType_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarMaintenancePlanDetails ON TblMaintenanceType_1.id = dbo.TblCarMaintenancePlanDetails.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblMaintenanceType ON dbo.TblCarMaintenancePlanDetails.MaintenanceID = dbo.TblMaintenanceType.id RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarMaintenancePlan ON dbo.TblCarMaintenancePlanDetails.Planid = dbo.TblCarMaintenancePlan.Planid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarsData ON dbo.TblCarMaintenancePlan.CarId = dbo.TblCarsData.id"
MySQL = MySQL & " Where (dbo.TblCarMaintenancePlan.Planid =" & val(XPTxtID.Text) & ")"
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarsPlan.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarsPlanE.rpt"
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
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
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
Sub loadDcbMaint(Optional ID As Double)
Dim My_SQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "SELECT     id, name, namee"
    Else
    My_SQL = "SELECT     id, namee"
    End If
My_SQL = My_SQL & " From dbo.TblMaintenanceType"
My_SQL = My_SQL & " where MainType<>1"
If ID <> 0 Then
My_SQL = My_SQL & " and FollowID=" & ID & " "
End If
 
    fill_combo DCMaintenanceTypes, My_SQL
End Sub
Sub loadDcb()
Dim My_SQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "SELECT     id, name, namee"
    Else
    My_SQL = "SELECT     id, namee"
    End If
My_SQL = My_SQL & " From dbo.TblMaintenanceType"
My_SQL = My_SQL & " where MainType=1 "
    fill_combo DcbGroup, My_SQL
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
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

Private Sub Form_Load()
    'On Error GoTo ErrTrap
    Dim Dcombos As New ClsDataCombos

    'Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetAccountingCodes Me.DCboUserName
 
    Dcombos.GetTblyearsData DcyearsData
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCars Me.Dccar
    'Dcombos.GetCarsMaintenanceTypes Me.DCMaintenanceTypes, , 0
     loadDcb
    Dim My_SQL As String

    ScreenNameArabic = " ŒÿÂ ’Ì«‰Â   «·„⁄œ« /«·”Ì«—«   "
    ScreenNameEnglish = "Cars Maintenance Plan Data"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Dcombos.GetBranches dcBranch

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Dcombos.GetAccountingCodes Me.DcboCreditSide

    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With

    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With

    AddTip
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    If publicCarId <> 0 Then
      StrSQL = "select * From TblCarMaintenancePlan where CarId=" & publicCarId
     Else
     StrSQL = "select * From TblCarMaintenancePlan "
    End If
    StrSQL = StrSQL & "  Order By Planid"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'rs.Open "TblCarMaintenancePlan", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Me.TxtModFlg.Text = "R"

    Retrive

    If rs.RecordCount = 0 Then
        Cmd_Click (0)
        Dccar.BoundText = publicCarId
        Exit Sub
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

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

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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
    Exit Sub
ErrTrap:
End Sub
 
Private Sub Opt_Click(Index As Integer)
Select Case Index
Case 0
Frame6.Enabled = True
Frame7.Enabled = False
Frame11.Enabled = False
Case 1
Frame7.Enabled = True
Frame6.Enabled = False
Frame11.Enabled = False
Case 2
Frame7.Enabled = False
Frame6.Enabled = False
Frame11.Enabled = True

End Select

End Sub

Private Sub Text3_Change()

                 
                   Dim Dcombos As New ClsDataCombos
    
    Dcombos.GetQuicSearch DCMaintenanceTypes, Text3, "TblMaintenanceType"
End Sub

Private Sub TXTCurrentKM_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TXTCurrentKM.Text, 0)

End Sub

Private Sub TxtKhordaPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtKhordaPrice.Text, 0)
End Sub

Private Sub txtKMCount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtKMCount.Text, 0)

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            '  txtLastKMCounter.locked = True
            '   Me.Caption = "«·«’Ê· «·À«» …"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            cStatus.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            Frame3.Enabled = False
        
            If rs.RecordCount < 1 Then
                '      Me.XPBtnMove(0).Enabled = False
                '      Me.XPBtnMove(1).Enabled = False
                '      Me.XPBtnMove(2).Enabled = False
                '      Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
            '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« ( ÃœÌœ )"
            txtLastKMCounter.locked = False
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            cStatus.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
         
        Case "E"
            ' txtLastKMCounter.locked = True
            Frame3.Enabled = False
            '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            cStatus.Enabled = False
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
         
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Not (rs.EOF Or rs.BOF) Then
        If Lngid <> 0 Then
            rs.Find "Planid=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If

    End If
                        
    Me.XPTxtID.Text = IIf(IsNull(rs("Planid").value), 0, rs("Planid").value)
    TxtNotes.Text = IIf(IsNull(rs("REMARKS").value), "", rs("REMARKS").value)
    Me.DcbGroup.BoundText = IIf(IsNull(rs("GroupID").value), 0, rs("GroupID").value)
    Dccar.BoundText = IIf(IsNull(rs("CarId").value), 0, rs("CarId").value)
    DPRecordDate.value = rs("RecordDate").value
    DcyearsData.BoundText = IIf(IsNull(rs("PlanYear").value), 0, rs("PlanYear").value) 'IIf(val(rs("PlanYear").value) = 0, 0, rs("PlanYear").value)
 
    Me.TXTCurrentKM.Text = IIf(IsNull(rs("CurrentKM").value), "", rs("CurrentKM").value)
   ' Me.TxtNoHour.text = IIf(IsNull(rs("NoHour").value), "", rs("NoHour").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), 0, rs("UserID").value)

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim i As Integer
 
    StrSQL = "SELECT  TblCarMaintenancePlanDetails.Done , TblCarMaintenancePlanDetails.cancelreason,   dbo.TblCarMaintenancePlan.Planid,dbo.TblCarMaintenancePlanDetails.OrderMaintinID, dbo.TblCarMaintenancePlanDetails.alarmType, dbo.TblCarMaintenancePlanDetails.MaintenanceID, "
    StrSQL = StrSQL & "                  dbo.TblMaintenanceType.name, dbo.TblCarMaintenancePlanDetails.AlarmInKM, dbo.TblCarMaintenancePlanDetails.AlarmINDate,"
    StrSQL = StrSQL & "                  dbo.TblCarMaintenancePlanDetails.AlarmINTime, dbo.TblCarMaintenancePlanDetails.NoHour, dbo.TblMaintenanceType.namee,"
    StrSQL = StrSQL & "                  dbo.TblCarMaintenancePlanDetails.GroupID, TblMaintenanceType_1.name AS GroupName, TblMaintenanceType_1.namee AS GroupNameE"
    StrSQL = StrSQL & "   FROM         dbo.TblMaintenanceType TblMaintenanceType_1 RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCarMaintenancePlanDetails ON TblMaintenanceType_1.id = dbo.TblCarMaintenancePlanDetails.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblMaintenanceType ON dbo.TblCarMaintenancePlanDetails.MaintenanceID = dbo.TblMaintenanceType.id RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCarMaintenancePlan ON dbo.TblCarMaintenancePlanDetails.Planid = dbo.TblCarMaintenancePlan.Planid"
    StrSQL = StrSQL & " WHERE     (dbo.TblCarMaintenancePlan.Planid = " & val(Me.XPTxtID.Text) & ")"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(RsDev("GroupID").value), 0, RsDev("GroupID").value)
                .TextMatrix(i, .ColIndex("hour")) = IIf(IsNull(RsDev("NoHour").value), 0, RsDev("NoHour").value)
                .TextMatrix(i, .ColIndex("alarmType")) = IIf(IsNull(RsDev("alarmType").value), 0, RsDev("alarmType").value)
                .TextMatrix(i, .ColIndex("MaintenanceID")) = IIf(IsNull(RsDev("MaintenanceID").value), "", RsDev("MaintenanceID").value)
                .TextMatrix(i, .ColIndex("OrderMaintinID")) = IIf(IsNull(RsDev("OrderMaintinID").value), "", RsDev("OrderMaintinID").value)
                .TextMatrix(i, .ColIndex("done")) = IIf(IsNull(RsDev("done").value), "", RsDev("done").value)
                
                
            If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("Group")) = IIf(IsNull(RsDev("GroupName").value), "", RsDev("GroupName").value)
                .TextMatrix(i, .ColIndex("MaintenanceType")) = IIf(IsNull(RsDev("NAME").value), "", RsDev("NAME").value)
                Else
                .TextMatrix(i, .ColIndex("Group")) = IIf(IsNull(RsDev("GroupNameE").value), "", RsDev("GroupNameE").value)
                .TextMatrix(i, .ColIndex("MaintenanceType")) = IIf(IsNull(RsDev("NameE").value), "", RsDev("NameE").value)
                End If
            .TextMatrix(i, .ColIndex("AlarmInKM")) = IIf(IsNull(RsDev("AlarmInKM").value), "", RsDev("AlarmInKM").value)
                .TextMatrix(i, .ColIndex("cancelreason")) = IIf(IsNull(RsDev("cancelreason").value), "", RsDev("cancelreason").value)
                .TextMatrix(i, .ColIndex("AlarmINDate")) = IIf(IsNull(RsDev("AlarmINDate").value), "", RsDev("AlarmINDate").value)
                   If Not IsNull(RsDev("AlarmINTime").value) Then
                   Dim ArrivalTime1 As Date
   If IsDate(RsDev("AlarmINTime").value) Then
         ArrivalTime1 = FormatDateTime(RsDev("AlarmINTime").value, vbShortTime)
         .TextMatrix(i, .ColIndex("AlarmINTime")) = ArrivalTime1
   End If
    End If
                         
                RsDev.MoveNext
            Next i
 
        End With

    End If

    RsDev.Close
    ReLineGrid

    Exit Sub
ErrTrap:
End Sub
 
Private Sub TxtNoOfIteration_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtNoOfIteration.Text, 0)

End Sub

Private Sub TxXPeriod_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxXPeriod.Text, 0)

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    'On Error GoTo ErrTrap
 
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

Private Sub SaveData()
    Dim sql As String
    Dim TblCarKMFOLLOWid As Double
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim RsDev As New ADODB.Recordset
    Dim RsNot As New ADODB.Recordset

    Dim BeginTrans As Boolean
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        
        Dim i As Long, mOrderMaintinID As Long
        For i = 1 To Grid.Rows - 1
            mOrderMaintinID = val(Grid.TextMatrix(i, Grid.ColIndex("OrderMaintinID")))
            If mOrderMaintinID <> 0 Then
                 MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ· ·ÊÃÊœ «·»‰œ —ﬁ„ " & i & " ›Ï Õ—ﬂ… «„— ‘€· —ﬁ„ " & mOrderMaintinID, vbCritical
                SendKeys "{F4}"
                Exit Sub
            End If
        Next
        
        If Dccar.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «”„ «·„⁄œÂ/«·”Ì«—…  «Ê·«", vbCritical
            Else
                MsgBox "Select Car Name Firstly    ", vbCritical
            End If
    
            Dccar.SetFocus
         
            SendKeys "{F4}"
            Exit Sub
        End If
 
       ' If val(Me.DcyearsData.BoundText) = 0 Then
       '     If SystemOptions.UserInterface = ArabicInterface Then
       '         MsgBox "Õœœ  «·”‰Â «·„«·ÌÂ", vbCritical
       '     Else
       '         MsgBox "Select year   ", vbCritical
       '     End If
'
'            DcyearsData.SetFocus
'            SendKeys "{F4}"
'            Exit Sub
'        End If
 
        Select Case Me.TxtModFlg.Text

            Case "N"
 
            Case "E"
        
        End Select

        If Me.TxtModFlg.Text = "N" Then
   
        End If

        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.Text

            Case "N"
                XPTxtID.Text = CStr(new_id("TblCarMaintenancePlan", "Planid", "", True))
            
                rs.AddNew
            
            Case "E"
            
            
            
                Cn.Execute "delete TblCarMaintenancePlanDetails where Planid=" & val(Me.XPTxtID.Text)
        
        End Select

        rs("Planid").value = val(Me.XPTxtID.Text)
        rs("REMARKS").value = IIf(Trim(TxtNotes.Text) = "", Null, TxtNotes.Text)
        rs("CarId").value = IIf(val(Dccar.BoundText) = 0, Null, val(Dccar.BoundText))
        rs("RecordDate").value = DPRecordDate.value
        rs("PlanYear").value = IIf(val(DcyearsData.BoundText) = 0, Null, DcyearsData.BoundText)
        rs("CurrentKM").value = IIf(Trim(TXTCurrentKM.Text) = "", Null, TXTCurrentKM.Text)
       ' rs("NoHour").value = IIf(Trim(TxtNoHour.text) = "", Null, TxtNoHour.text)
        rs("UserID").value = IIf(val(Me.DCboUserName.BoundText) = 0, Null, DCboUserName.BoundText)
        rs("GroupID").value = IIf(val(Me.DcbGroup.BoundText) = 0, Null, val(DcbGroup.BoundText))
                         
        rs.update
    End If
 
    '**************************************************************************
 
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TblCarMaintenancePlanDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    

    With Me.Grid

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("MaintenanceID")) <> "" Then
 
                RsDev.AddNew
                RsDev("Planid").value = val(Me.XPTxtID.Text)
                RsDev("GroupID").value = val(.TextMatrix(i, .ColIndex("GroupID")))
                RsDev("MaintenanceID").value = val(.TextMatrix(i, .ColIndex("MaintenanceID")))
                RsDev("alarmType").value = val(.TextMatrix(i, .ColIndex("AlarmType")))
                RsDev("NoHour").value = val(.TextMatrix(i, .ColIndex("hour")))
                RsDev("AlarmInKM").value = val(.TextMatrix(i, .ColIndex("AlarmInKM")))
                RsDev("AlarmINDate").value = IIf((.TextMatrix(i, .ColIndex("AlarmINDate"))) = "", Null, (.TextMatrix(i, .ColIndex("AlarmINDate"))))
                RsDev("done").value = IIf((.TextMatrix(i, .ColIndex("done"))) = "", Null, (.TextMatrix(i, .ColIndex("done"))))
                
                
              'sa   RsDev("AlarmINTime").value = IIf((.TextMatrix(i, .ColIndex("AlarmINTime"))) = "", Null, FormatDateTime((.TextMatrix(i, .ColIndex("AlarmINTime"))), vbShortTime))
                RsDev.update
            End If

        Next i

    End With
    
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"
              If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & CHR(13)
                    Msg = Msg + "do you new Operation?"
        
                End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If
           
        Case "E"
             If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
    End Select
    
    TxtModFlg.Text = "R"
 
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
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
            rs.Find "Planid=" & val(XPTxtID.Text) & "", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_AssetType()
    Dim msgstr  As String

    Dim sql As String

    Dim Msg As String
    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (Me.XPTxtID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                '       sql = "Delete   from notes where NoteID=" & Val(txtNoteID.text)
                '        Cn.Execute sql
        
                ' sql = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & Val(txtopening_balance_voucher_id.text)
                '    Cn.Execute sql, , adExecuteNoRecords
                '   sql = "delete  FixedAssetInstallmentsDetails where FixedAssetID=" & Val(Me.XPTxtID.text)
                '   Cn.Execute sql, , adExecuteNoRecords
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Grid.Clear flexClearScrollable, flexClearEverything
                    Grid.Rows = 1
                    TxtModFlg_Change
             
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate

End Sub
 
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ‰Ê⁄ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·‰Ê⁄ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

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
    lbl(19).Caption = "Group"
    Me.Caption = "Vehicles Maintenance Plan"
    Me.Ele.Caption = Me.Caption
    Me.lbl(10).Caption = "Code"
    Me.lbl(13).Caption = "Date"
    Me.lbl(14).Caption = "Year"
    Me.lbl(101).Caption = "Car"
    Me.lbl(124).Caption = "Remark"
    Me.lbl(13).Caption = "Date"
    Frame11.Caption = "Case Hours"
    Frame8.Caption = "Data"
    Me.lbl(13).Caption = "Date"
    Me.lbl(4).Caption = "Select Maint."
    Me.lbl(7).Caption = "Date"
 
    Me.lbl(6).Caption = "No Of Check up"
    Frame4.Caption = "Calculation Method"
    Opt(0).Caption = "KM"
    Opt(1).Caption = "Dates"
     Opt(2).Caption = "Time"
    Frame6.Caption = "KM Case"
 lbl(16).Caption = "Time"
 Me.lbl(18).Caption = "Hours Count"
    Me.lbl(11).Caption = "Km Count"
    Frame7.Caption = "Dates Case"
    Frame9.Caption = "Time Case"
    Me.lbl(6).Caption = "Start Date"
    Me.lbl(8).Caption = "Period"
    Me.lbl(15).Caption = "Period"

    Me.lbl(11).Caption = "Km Count"

    Frame5.Caption = "Select"
    optintervals(4).Caption = "Second"
    optintervals(5).Caption = "Minute"
    optintervals(3).Caption = "Hours"
Frame10.Caption = "Select"
    optintervals(0).Caption = "Day"
    optintervals(1).Caption = "Month"
    optintervals(2).Caption = "Year"
    CmdAdd.Caption = "Add"
    CmdDelete.Caption = "Delete"
    cmdClear.Caption = "Clear All"
    Cmd(12).Caption = "Search"

    lbl(12).Caption = "Current KM "
 
    Me.lbl(124).Caption = "Remark"

    Cmd(8).Caption = "Depreciation Restart"

    Me.lbl(125).Caption = "Current Record:"
    Me.lbl(126).Caption = "Records NO:"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(12).Caption = "Print"
    Me.Cmd(13).Caption = "Search"
    Cmd(7).Caption = "Stop Dep"

    With CBoDepreciation_Type_id
        .Clear
        .AddItem "fixed "
        .AddItem "Decreasing"
    End With

    lbl(5).Caption = "By"
    Cmd(5).Caption = "Attachments"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("MaintenanceType")) = "Maintenance Type"
        .TextMatrix(0, .ColIndex("AlarmInKM")) = "Alarm In KM"
        .TextMatrix(0, .ColIndex("AlarmINDate")) = "Alarm In Date"
        .TextMatrix(0, .ColIndex("AlarmINTime")) = "Alarm In Time"
        .TextMatrix(0, .ColIndex("hour")) = "Num Hours"
        .TextMatrix(0, .ColIndex("Group")) = "Group"


    End With

End Sub

